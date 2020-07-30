# -*- coding: utf-8 -*-
'''
Interactive Python Console with access to Office objects

Created on 23.01.2014
@author: cschmitt
'''

from __future__ import absolute_import

# import code
from cStringIO import StringIO
import sys
import traceback
import re
import math

from bkt.helpers import settings, endings_to_windows


class Mock(object):
    def __init__(self, *args, **kwargs):
        pass
    
    def __getitem__(self, key):
        return self
    
    def __setitem__(self, key, value):
        pass
    
    def __call__(self, *args, **kwargs):
        return self
    
    def __getattr__(self, attr):
        return self
    
    def __iter__(self):
        return iter([])

    def __iadd__(self, other):
        return self


load_ui = True

if load_ui:
    from . import dotnet
    F = dotnet.import_forms()
    D = dotnet.import_drawing()
    
    F.Application.EnableVisualStyles()  
else:
    F = Mock()
    D = Mock()

class FileWrapper(object):
    def __init__(self, fd):
        self._fd = fd
        self._callbacks = set()
        self._buffer = StringIO()
        
    def __getattr__(self, attr):
        return getattr(self._fd, attr)
    
    def _exec_callbacks(self, data):
        for callback in self._callbacks:
            callback(data)
    
    def write(self, data, *args, **kwargs):
        res = self._fd.write(data, *args, **kwargs)
        self._buffer.write(data)
        self._exec_callbacks(data)
        return res
    
    def flush(self):
        self._fd.flush()
        self._exec_callbacks("")


_stderr = sys.stderr    
_stdout = sys.stdout

def redirect_output():
    sys.stdout = FileWrapper(sys.stdout)
    sys.stderr = sys.stdout

def restore_output():
    sys.stdout = _stdout
    sys.stderr = _stderr

redirect_output()


def create_bottom_buttons(*buttons):
    button_labels = buttons
    buttons = []
    
    for label in button_labels:
        btn = F.Button()
        btn.Margin = F.Padding(5)
        #btn.Dock = F.DockStyle.Bottom
        btn.Width = 100
        btn.Height = 30
        btn.Text = label
        buttons.append(btn)

    bottom = F.FlowLayoutPanel()
    bottom.AutoSize = True
    bottom.WrapContents = False
    bottom.Dock = F.DockStyle.Bottom
    bottom.FlowDirection = F.FlowDirection.RightToLeft
    
    for btn in reversed(buttons):
        bottom.Controls.Add(btn)
    
    #F.MessageBox.Show("bottom.Controls.Count = %d\r\nbuttons = %s" % (bottom.Controls.Count,buttons))
    
    #return bottom
    return [bottom] + buttons

def create_console_textbox():
    box = F.TextBox()
    box.Font = D.Font(D.FontFamily.GenericMonospace, D.SystemFonts.MessageBoxFont.Size)
    box.Multiline = True
    box.ReadOnly = True
    box.WordWrap = False
    box.ScrollBars = F.ScrollBars.Both
    return box

def set_default_size(form):
    form.Width = 1440
    form.Height = 900
    form.StartPosition = F.FormStartPosition.CenterScreen

def create_input_textbox(cls=F.TextBox, *cargs, **ckwargs):
    box = cls(*cargs, **ckwargs)
    box.Font = D.Font(D.FontFamily.GenericMonospace, D.SystemFonts.MessageBoxFont.Size)
    box.Multiline = True
    box.WordWrap = False
    box.ScrollBars = F.ScrollBars.Both
    return box

def set_text(box, text, scroll_down=False):
    box.Text = text
    box.SelectionLength = 0
    if scroll_down:
        box.SelectionStart = len(box.Text)
        box.ScrollToCaret()

class ConsoleInput(F.TextBox):
    def __new__(cls, intercept):
        ins = F.TextBox.__new__(cls)
        ins.intercept = intercept
        return ins
    
    def ProcessCmdKey(self, msg, keys):
        if self.intercept(msg, keys):
            return True
        return F.TextBox.ProcessCmdKey(self, msg, keys)

CONSOLE_HELP_TEXT = '''
---------------------
  BKT Console Help
---------------------

USAGE:
    1. Enter code into the input text field
    2. Use CTRL+RETURN to execute
    3. Use UP and DOWN to browse command history
    4. Use TAB to complete simple statements

LIMITATIONS:
    * Other input than a single interactive statement will raise an exception, e.g.
      >>> print 'foo'
      ... print 'bar'
      You may use instead:
      >>> print 'foo'; print 'bar'
    * Code execution can not be interrupted via CTRL+C or otherwise
      since it is executed on the main thread.
    * Currently, _all_ exceptions during execution are caught and printed.
    
Available variables:
    context: wrapped .NET context object, access application via context.app
'''
        
class InteractiveConsole(F.Form):
    def __init__(self):
        self.Text = "BKT IronPython Console"
        self.Font = D.SystemFonts.MessageBoxFont
        self.history = settings.get("bkt.console.history", [])
        self.history_cursor = None
        self.history_uncommitted = None
        self.last_input = None
        
        set_default_size(self)
        sys.stdout._callbacks.add(self.stdout_callback)
        #self.interpreter = code.InteractiveInterpreter()
        self._globals = {}
        exec "import clr" in self._globals
        
        self.output = create_console_textbox()
        self.output.Dock = F.DockStyle.Fill 

        self.input = create_input_textbox(ConsoleInput, intercept=self.intercept_input)
        self.input.Dock = F.DockStyle.Fill
        set_text(self.output,
                 endings_to_windows(sys.stdout._buffer.getvalue()),
                 True)
        
        hsplit = F.SplitContainer()
        hsplit.Orientation = F.Orientation.Horizontal
        hsplit.SplitterDistance = 200
        hsplit.Dock = F.DockStyle.Fill

        hsplit.Panel1.Controls.Add(self.output)
        hsplit.Panel2.Controls.Add(self.input)
        
        self.Controls.Add(hsplit)
        self.FormClosing += self.cancel_close
        
        bottom, help, execute = create_bottom_buttons('Help', 'Execute')
        execute.Click += self.execute
        
        def print_help(*args):
            print CONSOLE_HELP_TEXT.strip()
        
        help.Click += print_help
        
        #refresh.Click += self.refresh
        
        self.Controls.Add(bottom)
        
    def intercept_input(self, msg, keys):
        if keys == F.Keys.Control | F.Keys.Return:
            #print 'intercepted CTRL+RETURN'
            try:
                self.execute()
            finally:
                return True
        elif keys == F.Keys.Up:
            return self.check_history_up()
        elif keys == F.Keys.Down:
            return self.check_history_down()
        elif keys == F.Keys.Control | F.Keys.A:
            self.input.SelectionStart = 0
            self.input.SelectionLength = self.input.Text.Length
        elif keys == F.Keys.Control | F.Keys.C and self.input.SelectionLength == 0:
            print endings_to_windows(self.input.Text.strip(), prepend='... ', prepend_first='>>> ')
            print "KeyboardInterrupt"
            self.input.Text = ""
        elif keys == F.Keys.Tab:
            self.tab_complete()
            return True
    
    def check_history_up(self):
        box = self.input
        if box.SelectionLength > 0:
            return False
        line_break = box.Text.find('\n')
        if not (line_break < 0 or box.SelectionStart < line_break):
            return False
        if not self.history:
            return False
        if self.history_cursor is None:
            self.history_cursor = len(self.history)-1
        elif self.history_cursor == 0:
            return False
        else:
            self.history_cursor -= 1
        
        box.Text = self.history[self.history_cursor]
        box.SelectionStart = 0
        box.SelectionLength = 0
        
        return True

    def check_history_down(self):
        box = self.input
        if box.SelectionLength > 0:
            return False
        line_break = box.Text.rfind('\n')
        if not (line_break < 0 or box.SelectionStart > line_break):
            return False
        if self.history_cursor is None or self.history_cursor >= len(self.history)-1:
            return False
        
        self.history_cursor += 1
        
        box.Text = self.history[self.history_cursor]
        box.SelectionStart = 0
        box.SelectionLength = 0
        
        return True
        
    def cancel_close(self, sender, e):
        e.Cancel = True
        self.Visible = False
        
    def stdout_callback(self, data):
        if data:
            self.output.AppendText(endings_to_windows(data))
            
    def scroll_down(self):
        box = self.input
        box.SelectionStart = len(str(box.Text))
        box.ScrollToCaret()
    

    SIMPLE_STRING1 = r'\"[^\"]*\"'
    SIMPLE_STRING2 = r"\'[^\']*\'"
    SIMPLE_VAR_OR_INT = r"[a-zA-Z0-9_]+"
    
    #GETITEM = r'\[(' + SIMPLE_STRING1 + r')|(' + SIMPLE_STRING2 + r')|(' + SIMPLE_VAR_OR_INT + r')\]'
    GETITEM = r'\[[0-9]+\]'
    SIMPLEFUNCCALL = r'\( *[a-zA-Z0-9_]+( *, *[a-zA-Z0-9_]+)* *\)'
    INTER_DOT_ELEM = r'[a-zA-Z0-9_]+(' + GETITEM + r'|' + SIMPLEFUNCCALL + r')?'
    PRECEDING_STUFF = r'(.*[, \(\[])?'
    ALLOWED_COMPLETION_INPUT = r'^' + PRECEDING_STUFF + r'((' + INTER_DOT_ELEM + r'\.)*(' + INTER_DOT_ELEM + r')?)$'
    
    allowed = re.compile(ALLOWED_COMPLETION_INPUT)
    
    def print_options(self, source, options, max_width=80):
        col_width = max(len(o) for o in options) + 1
        num_cols = max_width/col_width
        num_rows = int(math.ceil(float(len(options))/float(num_cols)))
        rows = ['' for row in xrange(num_rows)]
        
        def pad(s):
            return s + ((col_width-len(s))*' ')
        
        for i, option in enumerate(options):
            col, row = divmod(i, num_rows)
            rows[row] += pad(option)
        
        print('--- completion options for "' + source + '" ---')
        for row in rows:
            print(row)
            
    def longest_common_prefix(self, items):
        items = list(items)
        min_len = min(len(item) for item in items)
        if min_len == 0:
            return ''
        
        for i in xrange(min_len):
            # check whether character at position i is equal
            if len(set([item[i] for item in items])) > 1:
                # if not, return prefix (without i'th character)
                return items[0][:i]
        
        # never breaked, first min_len characters are all equal
        return items[0][:min_len+1]
    
    def replace_input(self, left, right):
        self.input.Text = left + right
        self.input.SelectionStart = len(left)
        self.input.SelectionLength = 0
    
    def tab_complete(self):
        source = self.input.Text
        if '\n' in source.strip() or '\r' in source.strip():
            return
        
        cursor_pos = self.input.SelectionStart + self.input.SelectionLength
        left_of_cursor, right_of_cursor = source[:cursor_pos], source[cursor_pos:]
        source_match = self.allowed.match(left_of_cursor)
        if source.strip() and not source_match:
            print("completion not allowed for %r" % source)
            return
        
        source_for_completion = source_match.group(2)
        if not '.' in source_for_completion:
            head = ''
            tail = source
        else:
            head, tail = source_for_completion.rsplit('.', 1)
        
        
        try:
            if head:
                #print("calling eval() with %r" % head)
                head_value = eval(head, globals=self._globals)
                if head_value is None:
                    return
                raw_options = dir(head_value)
            else:
                raw_options = sorted(self._globals)
            
            tail_match = tail.lower()
            if tail_match:
                options = [attr for attr in raw_options if attr.lower().startswith(tail_match)]
            else:
                options = raw_options
                
            if not tail.startswith('_'):
                options = [o for o in options if not o.startswith('_')]
                
            if not options:
                return
            
            
            if len(options) == 1:
                if head:
                    new_input = head + '.' + options[0]
                else:
                    new_input = options[0]
                
            else:
                prefix = self.longest_common_prefix(options)
                if prefix and len(prefix) > len(tail):
                    if head:
                        new_input = head + '.' + prefix
                    else:
                        new_input = prefix
                
                if source != self.last_input:
                    self.print_options(source_for_completion, options)
            
            if source_match.group(1):
                new_input = source_match.group(1) + new_input 
            
            self.last_input = source
            self.replace_input(new_input, right_of_cursor)

        except (SyntaxError, NameError, AttributeError):
            pass
        except  Exception, e:
            print(e)
    
    def execute(self, *args):
        source = self.input.Text
        
        if not source.strip():
            return
        
        source_orig = source
        
        def restore_input():
            self.input.Text = source_orig
            self.input.SelectionStart = 0
            self.input.SelectionLength = self.input.Text.Length
        
        if not (source.endswith('\n') or source.endswith('\r\n')):
            source += '\r\n'
        self.input.Text = ""
        self.history_cursor = None

        # print input
        print endings_to_windows(source.strip(), prepend='... ', prepend_first='>>> ')
        
        # save input
        if not self.history or source_orig != self.history[-1]:
            self.history.append(source_orig)
            settings["bkt.console.history"] = self.history[-5:] #store last 5 items
        
        # compile code
        try:
            code = compile(source, '<input>', 'single')
        except BaseException, ex:
            print ex
            restore_input()
            return
        
        # execute code
        try:
            exec code in self._globals
            self.last_input = None
            #if not self.history or source_orig != self.history[-1]:
            #    self.history.append(source_orig)
        except:
            traceback.print_exc()
            restore_input()
         
class ConsoleTextMessage(F.Form):
    def __init__(self):
        self.Text = "BKT Message"
        self.Font = D.SystemFonts.MessageBoxFont
        self.console = create_console_textbox()
        self.create_textbox()
        self.create_bottom()
        set_default_size(self)
        
    def set_text(self, text):
        box = self.box
        box.Text = text
        box.SelectionLength = 0
        box.SelectionStart = 0
        
    def create_textbox(self):
        box = create_console_textbox()
        box.Dock = F.DockStyle.Fill

        self.Controls.Add(box)
        self.box = box
        
    def create_bottom(self):
        bottom, ok = create_bottom_buttons('OK')
        ok.DialogResult = F.DialogResult.OK
        self.Controls.Add(bottom)
        self.AcceptButton = ok

# _parent = None
# 
# class ParentHackDialog(F.CommonDialog):
#     def RunDialog(self, parent):
#         global _parent
#         _parent = parent
#         return
# 
# ParentHackDialog().ShowDialog()

class ConsoleMessageDialog(F.CommonDialog):
    def __init__(self):
        self.form = ConsoleTextMessage()
        self.text = None
    
    def RunDialog(self, parent):
        self.form.CenterToScreen()
        self.form.set_text(self.text)
        res = self.form.ShowDialog()
        self.form.set_text("")
        return res == F.DialogResult.OK

_msg_dialog = None

def show_message(text):
    global _msg_dialog
    if _msg_dialog is None:
        _msg_dialog = ConsoleMessageDialog()
    
    _msg_dialog.text = text
    _msg_dialog.ShowDialog()
    _msg_dialog.text = None
    
def open_form(text):
    form = ConsoleTextMessage()
    form.set_text(text)
    form.Visible = True

console = InteractiveConsole()


class InputMessageBox(F.Form):
    def __init__(self, ok_method):
        self.Text = "input message box"
        #self.Font = D.SystemFonts.MessageBoxFont
        
        self.Width = 800
        self.Height = 600
        self.StartPosition = F.FormStartPosition.CenterScreen
        
        
        self.input = create_input_textbox()
        self.input.Dock = F.DockStyle.Fill
        
        self.Controls.Add(self.input)
        self.FormClosing += self.cancel_close
        
        bottom, cancel, ok = create_bottom_buttons('Cancel', 'OK')
        self.Controls.Add(bottom)
        cancel.Click += self.cancel
        ok.Click += self.ok
        
        self.ok_method = ok_method
     
    def set_text(self, text):
        self.input.Text = text
        
    def intercept_input(self, msg, keys):
        pass
        # try:
        #     source = self.input.Text
        # finally:
        #     return True
        
    def cancel_close(self, sender, e):
        e.Cancel = True
        self.Visible = False
    
    def cancel(self, sender, e):
        self.Visible = False
    
    def ok(self, *args):
        source = self.input.Text
        self.ok_method(source)


def show_input(text, ok_method=None):
    form = InputMessageBox(ok_method)
    form.set_text(text)
    form.Visible = True


def main():
    r = InteractiveConsole.allowed
    print(r.match('a[test]').group())
    
    return
    show_message("this is a test")
    open_form("this is another test")
    import time
    time.sleep(10)

if __name__ == '__main__':
    main()
