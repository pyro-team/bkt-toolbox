# -*- coding: utf-8 -*-
'''
Abstract class for WPF windows and popups, standard input and progress bar windows

Created on 23.01.2014
@author: cschmitt, fstallmann
'''

from __future__ import absolute_import

import logging
# import traceback
import os.path

import System
Window = System.Windows.Window
Popup = System.Windows.Controls.Primitives.Popup
BitmapImage = System.Windows.Media.Imaging.BitmapImage
BackgroundWorker = System.ComponentModel.BackgroundWorker

from bkt import dotnet
wpf = dotnet.import_wpf()
bkt_addin = dotnet.import_bkt()

from bkt.library.wpf.notify import NotifyPropertyChangedBase, notify_property
from bkt.apps import Resources



# =======================
# = UI MODEL AND WINDOW =
# =======================


class Singleton(type):
    _instances = {}
    def __call__(cls, *args, **kwargs):
        if cls not in cls._instances:
            cls._instances[cls] = super(Singleton, cls).__call__(*args, **kwargs)
        return cls._instances[cls]

class ViewModelSingleton(NotifyPropertyChangedBase):
    __metaclass__ = Singleton

    def __init__(self):
        super(ViewModelSingleton, self).__init__()

class ViewModelAsbtract(NotifyPropertyChangedBase):
    def __init__(self):
        super(ViewModelAsbtract, self).__init__()


class WpfWindowAbstract(bkt_addin.BktWindow):
    _filename = None
    _vm_class = None
    _vm       = None
    _context  = None

    # @staticmethod
    # def get_main_window_handle():
    #     ''' returns main window hwnd handle of current process '''
    #     try:
    #         return System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle
    #     except:
    #         #logging.error(traceback.format_exc())
    #         return 0
    
    @classmethod
    def create_and_show_dialog(cls, *args,**kwargs):
        wnd = cls(*args,**kwargs)
        wnd.show_dialog()
        return wnd
    
    def show_dialog(self, modal=True):
        if self._context is not None:
            self.SetOwner(self._context.addin.GetWindowHandle())
        # System.Windows.Interop.WindowInteropHelper(self).Owner = self.get_main_window_handle()
        if modal:
            return self.ShowDialog()
        else:
            return self.Show()

    def __init__(self, context=None):
        wpf.LoadComponent(self, self._filename)

        if context is not None:
            self._context = context
        if self._vm_class is not None:
            self._vm = self._vm_class()
        if self._vm is not None:
            self.DataContext = self._vm

    def __getattr__(self, name):
        # provides easy access to XAML elements (e.g. self.Button)
        return self.root.FindName(name)

    def cancel(self, sender, event):
        self.Close()


class WpfPopupAbstract(Popup):
    _filename = None
    _vm_class = None
    _vm       = None

    def __init__(self, context=None):
        wpf.LoadComponent(self, self._filename)

        self._context = context
        if self._vm_class is not None:
            self._vm = self._vm_class()
        if self._vm is not None:
            self.DataContext.DataContext = self._vm

    def __getattr__(self, name):
        # provides easy access to XAML elements (e.g. self.Button)
        return self.root.FindName(name)
    
    def cancel(self, sender, event):
        self.IsOpen = False



def convert_bitmap_to_bitmapsource(bitmap):
    return System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
        bitmap.GetHbitmap(),
        System.IntPtr.Zero,
        System.Windows.Int32Rect.Empty,
        System.Windows.Media.Imaging.BitmapSizeOptions.FromEmptyOptions()
    )

def load_bitmapimage(image_name):
        path = Resources.images.locate(image_name)  #@UndefinedVariable
        if path is None:
            raise IndexError("Image file not found")
        return BitmapImage(System.Uri(path))


def endings_to_windows(text, prepend="", prepend_first=""):
    def _iter():
        first = True
        for line in text.split('\n'):
            if not line.endswith('\r'):
                line = line + '\r'
            if first:
                line = prepend_first + line
                first = False
            else:
                line = prepend + line
            yield line
    
    return '\n'.join(_iter())

def endings_to_unix(text):
    def _iter():
        for line in text.split('\n'):
            if line.endswith('\r'):
                yield line[:-1]
            else:
                yield line
    
    return '\n'.join(_iter())



F = None
D = None

class UserInputBox(object):
    def __init__(self, text, title="BKT"):
        global F, D
        if F is None:
            F = dotnet.import_forms()
            D = dotnet.import_drawing()
            F.Application.EnableVisualStyles()  

        prompt = F.Form()
        prompt.Width = 500
        prompt.Height = 10
        prompt.Text = title
        prompt.StartPosition = F.FormStartPosition.CenterScreen
        prompt.AutoSize = True
        #prompt.TopMost = True
        prompt.MinimizeBox = False
        prompt.MaximizeBox = False
        prompt.ShowInTaskbar = False
        prompt.SizeGripStyle = F.SizeGripStyle.Hide
        prompt.FormBorderStyle = F.FormBorderStyle.FixedDialog
        #TODO: allow adjustment of width

        superPanel = F.FlowLayoutPanel()
        superPanel.AutoSize = True
        superPanel.WrapContents = False
        superPanel.FlowDirection = F.FlowDirection.TopDown
        superPanel.Dock = F.DockStyle.Top
        superPanel.DockPadding.All = 10
        #superPanel.Anchor = (F.AnchorStyles.Top | F.AnchorStyles.Left | F.AnchorStyles.Right)
        #superPanel.BorderStyle = F.BorderStyle.Fixed3D

        confirmation = F.Button()
        confirmation.Text = "OK"
        confirmation.Width = 228
        confirmation.Height = 30
        confirmation.Click += self.confirmation

        cancel = F.Button()
        cancel.Text = "Abbrechen"
        cancel.Width = confirmation.Width
        cancel.Height = confirmation.Height
        cancel.Click += self.cancel

        buttonsPanel = F.FlowLayoutPanel()
        buttonsPanel.AutoSize = True
        buttonsPanel.Height = confirmation.Height
        buttonsPanel.WrapContents = False
        buttonsPanel.FlowDirection = F.FlowDirection.LeftToRight
        buttonsPanel.Controls.Add(confirmation)
        buttonsPanel.Controls.Add(cancel)
        buttonsPanel.Dock = F.DockStyle.Bottom
        buttonsPanel.DockPadding.All = 10
        #buttonsPanel.BorderStyle = F.BorderStyle.Fixed3D

        prompt.AcceptButton = confirmation
        prompt.CancelButton = cancel

        self.superPanel = superPanel
        self.buttonsPanel = buttonsPanel
        self.prompt = prompt
        self.input = []
        self.values = {}

        self._add_label(text)
        # self._add_textbox(default, multiline)

    def _add_label(self, text=None):
        textLabel = F.Label()
        textLabel.Width = self.prompt.Width - 40
        textLabel.MaximumSize = D.Size(self.prompt.Width - 40, 200)
        textLabel.Text = text
        textLabel.Padding = F.Padding(0, 10, 0, 0)
        textLabel.AutoSize = True
        #self.prompt.Height += textLabel.Height + 5
        self.superPanel.Controls.Add(textLabel)
        return textLabel

    def _add_textbox(self, input_id, text=None, multiline=False):
        textBox = F.TextBox()
        textBox.Width = self.prompt.Width - 40
        textBox.Text = text
        #textBox.Anchor = (F.AnchorStyles.Left | F.AnchorStyles.Right)
        if multiline:
            textBox.Multiline = True
            textBox.ScrollBars = F.ScrollBars.Vertical
            textBox.AcceptsReturn = True
            #textBox.AcceptsTab = True
            textBox.WordWrap = True
            textBox.Height = 50
        #self.prompt.Height += textBox.Height + 5
        self.superPanel.Controls.Add(textBox)
        self.input.append((input_id, textBox, "Text"))
        return textBox

    #TODO: Automatically create history of entries
    #TODO: support dict as dropdown and return key-name instead of value
    def _add_combobox(self, input_id, text=None, dropdown=[], editable=True, selected_index=None, return_value="Text"):
        comboBox = F.ComboBox()
        comboBox.Width = self.prompt.Width - 40
        comboBox.DropDownStyle = F.ComboBoxStyle.DropDown if editable else F.ComboBoxStyle.DropDownList
        
        for item in dropdown:
            comboBox.Items.Add(item)
        
        if text is not None:
            comboBox.Text = text #if text is found in items, automatically selects it
        if selected_index is not None:
            comboBox.SelectedIndex = selected_index
        
        self.superPanel.Controls.Add(comboBox)
        self.input.append((input_id, comboBox, return_value))
        return comboBox

    def _add_spinner(self, input_id, value=0, min_value=0, max_value=100, return_value="Value"):
        spinnerBox = F.NumericUpDown()
        spinnerBox.Width = self.prompt.Width - 40
        if value is not None:
            spinnerBox.Value = value
        else:
            spinnerBox.Value = 0
            spinnerBox.Text = ""
        spinnerBox.Minimum = min_value
        spinnerBox.Maximum = max_value
        self.superPanel.Controls.Add(spinnerBox)
        self.input.append((input_id, spinnerBox, return_value))
        return spinnerBox

    def _add_checkbox(self, input_id, text=None, checked=False):
        checkBox = F.CheckBox()
        checkBox.Width = self.prompt.Width - 40
        checkBox.Text = text
        #checkBox.Padding = F.Padding(0, 10, 0, 0)
        checkBox.Checked = checked
        #self.prompt.Height += checkBox.Height + 5
        self.superPanel.Controls.Add(checkBox)
        self.input.append((input_id, checkBox, "Checked"))
        return checkBox

    def _add_listbox(self, input_id, lb_list, lb_return="Items"):
        listBox = F.ListBox()
        listBox.Width = self.prompt.Width - 40
        listBox.HorizontalScrollbar = True
        for item_name in lb_list:
            listBox.Items.Add(item_name)
        self.superPanel.Controls.Add(listBox)
        self.input.append( (input_id, listBox, lambda x: list(getattr(x, lb_return, []))) )
        return listBox

    def _add_checked_listbox(self, input_id, clb_list, clb_return="CheckedItems"):
        clb = F.CheckedListBox()
        clb.Width = self.prompt.Width - 40
        clb.HorizontalScrollbar = True
        clb.CheckOnClick = True
        for item_name, item_checked in clb_list:
            clb.Items.Add(item_name, item_checked)
        self.superPanel.Controls.Add(clb)
        self.input.append( (input_id, clb, lambda x: list(getattr(x, clb_return, []))) )
        return clb

    def _add_radio_buttons(self, input_id, text=None, radio_list=[], checked_index = 0):
        def _get_checked(radiobuttons):
            for radio in radiobuttons:
                if radio.Checked:
                    return radio.Text
            return None

        groupBox = F.GroupBox()
        groupBox.Width = self.prompt.Width - 40
        groupBox.Padding = F.Padding(0, 10, 0, 0)
        groupBox.Text = text
        top = 20
        radioButtons = []
        for i, radio_el in enumerate(radio_list):
            radioButton = F.RadioButton()
            radioButton.Width = self.prompt.Width - 60
            radioButton.Text = radio_el
            radioButton.Top = top
            radioButton.Left = 10
            #Check element
            if i == checked_index:
                radioButton.Checked = True
            top += radioButton.Height
            radioButtons.append(radioButton)
            groupBox.Controls.Add(radioButton)
        groupBox.Height = top + 10
        self.superPanel.Controls.Add(groupBox)
        self.input.append( (input_id, radioButtons, _get_checked) )
        return groupBox

    def _add_custom(self, input_id, input_element):
        input_element.Width = self.prompt.Width - 40
        self.superPanel.Controls.Add(input_element)
        self.input.append((input_id, input_element))
        return input_element

    def confirmation(self, sender, e):
        for con_input in self.input:
            if len(con_input) > 2 and callable(con_input[2]):
                self.values[con_input[0]] = con_input[2](con_input[1])
            elif len(con_input) > 2 and type(con_input[2]) == str:
                self.values[con_input[0]] = getattr(con_input[1], con_input[2], None)
            else:
                self.values[con_input[0]] = con_input[1]
        self.prompt.Close()

    def cancel(self, sender, e):
        self.prompt.Close()

    def show(self, dialog=True):
        self.prompt.Controls.Add(self.superPanel)
        self.prompt.Controls.Add(self.buttonsPanel)
        self.buttonsPanel.BringToFront()
        if dialog:
            self.prompt.ShowDialog()
            return self.values
        else:
            self.prompt.TopMost = True
            self.prompt.Show()



class WpfUserInput(WpfWindowAbstract):
    _filename = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'ui_inputbox.xaml')

    def __init__(self, text, title, default, multiline):
        super(WpfUserInput, self).__init__()

        self.Title = title

        if multiline:
            self.input_box.AcceptsReturn = True
            self.input_box.TextWrapping = System.Windows.TextWrapping.Wrap
            self.input_box.Height = 100

        self.text_label.Text = text
        self.input_box.Text = default or ""
        self.input_box.SelectAll()

    def cancel(self, sender, event):
        self.DialogResult = False
        self.Close()
    
    def confirm(self, sender, event):
        self.DialogResult = True
        self.Close()

def show_user_input(text, title, default=None, multiline=False):
    wnd = WpfUserInput(text, title, default, multiline)
    res = wnd.show_dialog()
    return wnd.input_box.Text if res is True else None

    # form = UserInputBox(text, title)
    # form._add_textbox("input", default, multiline)
    # value = form.show()
    # return None if "input" not in value else value["input"]


class WpfProgressBar(WpfWindowAbstract):
    _filename = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'ui_progressbar.xaml')

    def __init__(self, work_func, context=None, indeterminate=False):
        super(WpfProgressBar, self).__init__(context=context)

        self.progress_bar.IsIndeterminate = indeterminate

        self.bw = BackgroundWorker()
        self.bw.WorkerReportsProgress = True
        self.bw.WorkerSupportsCancellation = True
        
        self.bw.DoWork += self.bw_DoWork
        self.bw.ProgressChanged += self.bw_ProgressChanged
        self.bw.RunWorkerCompleted += self.bw_RunWorkerCompleted

        self.ContentRendered += self.startProgress
        self.Closing += self.stopProgress

        self._work_func = work_func
        # self.Title = title

    def bw_DoWork(self, sender, e):
        # sender.ReportProgress(2)
        e.Result = self._work_func(worker=sender)
        if sender.CancellationPending:
            e.Cancel = True

    def bw_ProgressChanged(self, sender, e):
        try:
            self.progress_bar.Value = e.ProgressPercentage
            if e.UserState and len(e.UserState.ToString()) > 0:
                self.progress_text.Text = e.UserState.ToString()
        except:
            logging.error("Error updating progress bar")
            # logging.debug(traceback.format_exc())

    def bw_RunWorkerCompleted(self, sender, e):
        self.Close()
    
    def startProgress(self, sender, e):
        self.bw.RunWorkerAsync()

    def stopProgress(self, sender, e):
        self.bw.CancelAsync() #send cancel request
        if self.bw.IsBusy:
            e.Cancel = True #cancel closing

def execute_with_progress_bar(work_func, context=None, modal=True, indeterminate=False):
    wnd = WpfProgressBar(work_func, context=context, indeterminate=indeterminate)
    wnd.show_dialog(modal=modal)
    return wnd