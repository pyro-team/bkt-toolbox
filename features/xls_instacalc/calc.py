# -*- coding: utf-8 -*-
'''
Created on 2017-07-18
@author: Florian Stallmann
'''

import bkt
import bkt.library.excel.helpers as xllib
import bkt.library.excel.constants as xlcon

import bkt.dotnet as dotnet
F = dotnet.import_forms() #required to copy text to clipboard
D = dotnet.import_drawing() #required for sizes

calc_history = []

class CalcHistoryForm(object):
    def __init__(self, calc_form):
        self.calc_form = calc_form

        prompt = F.Form();
        # prompt.Width = 400;
        # prompt.Height = 230;
        prompt.ClientSize = D.Size(400, 230)
        prompt.Text = "Sofort-Mini-Rechner Verlauf";
        prompt.StartPosition = F.FormStartPosition.CenterScreen;
        prompt.TopMost = True
        prompt.MinimizeBox = False
        prompt.MaximizeBox = False
        prompt.ShowInTaskbar = False
        prompt.SizeGripStyle = F.SizeGripStyle.Hide
        prompt.FormBorderStyle = F.FormBorderStyle.FixedSingle
        prompt.ShowIcon = False
        prompt.Padding = F.Padding(10)
        self.prompt = prompt

        listView = F.ListView()
        listView.Dock = F.DockStyle.Fill
        listView.View = F.View.Details
        listView.FullRowSelect = True
        listView.GridLines = True

        listView.DoubleClick += self.item_selected

        listView.Columns.Add("Formel", -2, F.HorizontalAlignment.Left)
        listView.Columns.Add("Ergebnis", -2, F.HorizontalAlignment.Right)

        for calc, result in reversed(calc_history):
            item1 = F.ListViewItem(calc)
            item1.SubItems.Add(result)
            listView.Items.Add(item1)
        self.listview = listView

        prompt.Controls.Add(listView)

    def item_selected(self, sender, e):
        self.calc_form.inputbox.Text = self.listview.SelectedItems[0].SubItems[0].Text
        self.prompt.Close()

    def show(self):
        self.prompt.ShowDialog(self.calc_form.prompt)


class CalcForm(object):
    def __init__(self, context, default_text="3,142+MITTELWERT(42;73)"):
        self.context = context
        self.application = context.app
        self.dec_sep = self.application.International(xlcon.XlApplicationInternational["xlDecimalSeparator"])
        self.numberformat = "#.##0,00##########"
        #self.type_range = type(self.application.Range("A1"))

        width = 200

        prompt = F.Form();
        # prompt.Width = width+38;
        # prompt.Height = 198;
        prompt.ClientSize = D.Size(width+20, 170)
        prompt.MinimumSize = prompt.Size
        prompt.MaximumSize = D.Size(width+400, prompt.Height)

        prompt.Text = "Sofort-Mini-Rechner";
        prompt.StartPosition = F.FormStartPosition.CenterScreen;
        #prompt.TopMost = True
        prompt.MinimizeBox = False
        prompt.MaximizeBox = False
        prompt.ShowInTaskbar = False
        # prompt.SizeGripStyle = F.SizeGripStyle.Hide
        # prompt.FormBorderStyle = F.FormBorderStyle.FixedSingle
        prompt.FormClosed += self.form_closed
        #prompt.Icon = D.SystemIcons.Application
        prompt.ShowIcon = False
        self.prompt = prompt

        toolTip = F.ToolTip()
        toolTip.ShowAlways = True #show also when form is inactive

        timer = F.Timer()
        timer.Interval = 250
        timer.Tick += self.timer_tick
        self.timer = timer

        tablePanel = F.TableLayoutPanel()
        tablePanel.Top = 10
        tablePanel.Left = 10
        tablePanel.Size = D.Size(width, 150)
        tablePanel.Anchor = F.AnchorStyles.Top | F.AnchorStyles.Bottom | F.AnchorStyles.Left | F.AnchorStyles.Right
        tablePanel.AutoSize = True
        tablePanel.ColumnCount = 2
        tablePanel.ColumnStyles.Add(F.ColumnStyle(F.SizeType.Percent, 50)) #50% width
        tablePanel.ColumnStyles.Add(F.ColumnStyle(F.SizeType.Percent, 50)) #50% width
        tablePanel.RowCount = 5
        tablePanel.RowStyles.Add(F.RowStyle()) #auto height
        tablePanel.RowStyles.Add(F.RowStyle()) #auto height
        tablePanel.RowStyles.Add(F.RowStyle()) #auto height
        tablePanel.RowStyles.Add(F.RowStyle()) #auto height
        tablePanel.RowStyles.Add(F.RowStyle()) #auto height

        ### ROW 1 ###

        btn_take_address = F.Button()
        btn_take_address.Anchor = F.AnchorStyles.Left | F.AnchorStyles.Right
        btn_take_address.Text = "→□ Adresse"
        btn_take_address.Click += self.take_address
        toolTip.SetToolTip(btn_take_address, "Adresse der aktuellen Selektion in Eingabe einfügen.\nMit Umschalt-Taste wird der Tabellenname vorangestellt.")

        btn_take_value = F.Button()
        btn_take_value.Anchor = F.AnchorStyles.Left | F.AnchorStyles.Right
        btn_take_value.Text = "→□ Wert"
        btn_take_value.Click += self.take_value
        toolTip.SetToolTip(btn_take_value, "Wert der aktiven Zelle in Eingabe einfügen.")

        tablePanel.Controls.Add(btn_take_address,0,0) #col, row
        tablePanel.Controls.Add(btn_take_value,1,0) #col, row

        ### ROW 2 ###

        inputPanel = F.Panel()
        inputPanel.AutoSize = True
        inputPanel.Size = D.Size(width,22)
        inputPanel.Dock = F.DockStyle.Fill

        #INFO: ComboBox does not have Paste-Function, so using TextBox
        #inputBox = F.ComboBox()
        inputBox = F.TextBox()
        inputBox.AutoSize = True
        inputBox.Size = D.Size(width-20,22)
        inputBox.Anchor = F.AnchorStyles.Left | F.AnchorStyles.Right
        inputBox.Text = default_text
        inputBox.TextChanged += self.form_changed
        inputPanel.Controls.Add(inputBox)
        self.inputbox = inputBox

        btn_history = F.Button()
        btn_history.Left = width-20
        btn_history.Text = "V"
        btn_history.Size = D.Size(20,inputBox.Height)
        btn_history.Anchor = F.AnchorStyles.Right
        inputPanel.Controls.Add(btn_history)
        btn_history.Click += self.show_history
        toolTip.SetToolTip(btn_history, "Verlauf anzeigen")

        tablePanel.SetColumnSpan(inputPanel, 2) #col, row
        tablePanel.Controls.Add(inputPanel,0,1) #col, row

        ### ROW 3 ###

        resultPanel = F.Panel()
        resultPanel.AutoSize = True
        resultPanel.Size = D.Size(width,22)
        resultPanel.Dock = F.DockStyle.Fill

        resultBox = F.TextBox()
        resultBox.Size = D.Size(width-40,22)
        resultBox.Anchor = F.AnchorStyles.Left | F.AnchorStyles.Right
        resultBox.ReadOnly = True
        #resultBox.GotFocus += self.result_focus
        resultPanel.Controls.Add(resultBox)
        self.resultbox = resultBox

        confirmation = F.Button()
        confirmation.Left = width-40
        confirmation.Size = D.Size(20,resultBox.Height)
        confirmation.Anchor = F.AnchorStyles.Right
        confirmation.Text = "A"
        resultPanel.Controls.Add(confirmation)
        confirmation.Click += self.okay_clicked
        toolTip.SetToolTip(confirmation, "Ergebnis neu berechnen")

        btn_copy = F.Button()
        btn_copy.Left = width-20
        btn_copy.Text = "K"
        btn_copy.Size = D.Size(20,resultBox.Height)
        btn_copy.Anchor = F.AnchorStyles.Right
        resultPanel.Controls.Add(btn_copy)
        btn_copy.Click += self.copy_result
        toolTip.SetToolTip(btn_copy, "Ergebnis in Zwischenablage kopieren")

        tablePanel.SetColumnSpan(resultPanel, 2) #col, row
        tablePanel.Controls.Add(resultPanel,0,2) #col, row

        ### ROW 4 ###

        checkBox = F.CheckBox()
        checkBox.AutoSize = True
        checkBox.Text = "Live-Auswertung"
        checkBox.Checked = True
        checkBox.CheckedChanged += self.form_changed
        self.livecalc = checkBox

        tablePanel.SetColumnSpan(checkBox, 2) #col, row
        tablePanel.Controls.Add(checkBox,0,3) #col, row

        ### ROW 5 ###

        btn_formula_to_cell = F.Button()
        btn_formula_to_cell.Anchor = F.AnchorStyles.Left | F.AnchorStyles.Right
        btn_formula_to_cell.Text = "Formel →□"
        btn_formula_to_cell.Click += self.formula_to_cell
        toolTip.SetToolTip(btn_formula_to_cell, "Eingegebene Formel in aktive Zelle einfügen.")

        btn_value_to_cell = F.Button()
        btn_value_to_cell.Anchor = F.AnchorStyles.Left | F.AnchorStyles.Right
        btn_value_to_cell.Text = "Wert →□"
        btn_value_to_cell.Click += self.value_to_cell
        toolTip.SetToolTip(btn_value_to_cell, "Aktuelles Ergebnis in aktive Zelle einfügen.")

        tablePanel.Controls.Add(btn_formula_to_cell,0,4) #col, row
        tablePanel.Controls.Add(btn_value_to_cell,1,4) #col, row

        prompt.Controls.Add(tablePanel)

        ### INVISIBLE ###

        cancel = F.Button()
        cancel.Top = prompt.Height + 50 #outside form
        cancel.Left = 10
        # cancel.Visible = False #does not work if invisible
        cancel.Text = "Schließen"
        prompt.Controls.Add(cancel)
        cancel.Click += self.cancel_clicked

        prompt.AcceptButton = confirmation;
        prompt.CancelButton = cancel;

        self._recalc(True)

    # def result_focus(self, sender, e):
    #     self.resultbox.SelectAll()

    def show_history(self, sender, e):
        form = CalcHistoryForm(self)
        form.show()

    def copy_result(self, sender, e):
        self._add_to_history()
        F.Clipboard.SetText(self.resultbox.Text)
        self.inputbox.Focus()

    def cancel_clicked(self, sender, e):
        self.prompt.Close()

    def okay_clicked(self, sender, e):
        # #Insert item at top
        # self.inputbox.Items.Insert(0, self.inputbox.Text)
        # if self.inputbox.Items.Count > 20:
        #     self.inputbox.Items.RemoveAt(20)
        # #self.inputbox.Items.Add(self.inputbox.Text)
        self._recalc(True)
        self.inputbox.SelectAll()
        self.inputbox.Focus()

    def timer_tick(self, sender, e):
        self.timer.Stop()
        self._recalc()

    def form_changed(self, sender, e):
        #self._recalc()
        self.timer.Stop()
        self.timer.Start()

    def sheet_changed(self, sheet, target):
        self._recalc()

    def _recalc(self, add_to_history=False):
        if not self.livecalc.Checked:
            return
        try:
            self.resultbox.Text = xllib.xls_evaluate(self.inputbox.Text, self.dec_sep, self.numberformat)
        except:
            self.resultbox.Text = "UNGÜLTIGE EINGABE"

        if add_to_history:
            self._add_to_history()

    def _add_to_history(self):
        el = (self.inputbox.Text, self.resultbox.Text)
        #delete duplicate and re-add it to new position
        try:
            calc_history.remove(el)
        except:
            pass
        calc_history.append(el)

    def take_address(self, sender, e):
        if self.prompt.ModifierKeys == F.Keys.Shift:
            address = xllib.get_address_external(self.application.ActiveWindow.RangeSelection)
        else:
            address = self.application.ActiveWindow.RangeSelection.AddressLocal(False, False)
        try:
            self.inputbox.Paste( address )
            self.inputbox.Focus()
            self._add_to_history()
        except:
            pass

    def take_value(self, sender, e):
        try:
            cell = self.application.ActiveCell
            if cell.Value2 is not None:
                self.inputbox.Paste( unicode(self.application.ActiveCell.Value2).replace('.', self.dec_sep) )
            self.inputbox.Focus()
            self._add_to_history()
        except:
            pass

    def formula_to_cell(self, sender, e):
        try:
            self._add_to_history()
            res = self.inputbox.Text
            self.application.ActiveCell.FormulaLocal = '=' + res.lstrip('=')
        except:
            pass

    def value_to_cell(self, sender, e):
        try:
            self._add_to_history()
            res = self.application.Evaluate(self.inputbox.Text)
            self.application.ActiveCell.Value = res
        except:
            pass

    def show(self):
        self.application.SheetSelectionChange += self.sheet_changed
        try:
            handle = self.context.addin.GetW32WindowHandle()
            if not handle:
                raise ValueError("no window handle")
            self.prompt.Show(handle)
        except:
            self.prompt.TopMost = True
            self.prompt.Show()

    def form_closed(self, sender, e):
        self.application.SheetSelectionChange -= self.sheet_changed
        pass


class Calc(object):
    # @staticmethod
    # def run_win_calc(application):
    #     import os
    #     os.system('start calc.exe')

    @classmethod
    def show_calc(cls, context):
        calcform = CalcForm(context)
        calcform.show()

calculator = Calc()


rechner_gruppe = bkt.ribbon.Group(
    label="Mini-Rechner",
    image_mso="Calculator",
    children=[
        bkt.ribbon.Button(
            id = 'run_eval',
            label="Neuer Mini-Rechner",
            show_label=True,
            size='large',
            image_mso='Calculator',
            screentip="Zeigt einen kleines Tool zum Auswerten von Formeln und Berechnungen an an",
            on_action=bkt.Callback(calculator.show_calc, context=True),
            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
        ),
    ]
)

bkt.excel.add_tab(bkt.ribbon.Tab(
    id='bkt_excel_toolbox_advanced',
    #id_q='nsBKT:excel_toolbox_advanced',
    label=u'Toolbox 3/3 BETA',
    insert_before_mso="TabHome",
    get_visible=bkt.Callback(lambda: True),
    children = [
        rechner_gruppe,
    ]
), True)