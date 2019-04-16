# -*- coding: utf-8 -*-
'''
Created on 2017-07-18
@author: Florian Stallmann
'''

import bkt
import bkt.library.excel.helpers as xllib
import bkt.library.excel.constants as xlcon

# import System.Convert as Converter #required for convert to double
import System.DateTime as DateTime #required for parse as date and time

import bkt.dotnet as dotnet
Forms = dotnet.import_forms() #required to copy text to clipboard

class CellsOps(object):
    # hidden_columns = {}
    # hidden_rows = {}

    last_formula = "*100"
    last_prepend = "ID-"
    last_append = "...!?"
    last_slice_pos = "2:"
    last_slice_text = "/"

    @staticmethod
    def _set_hidden_name(key, sheet, rng):
        try:
            #sheet.Names(key).RefersToLocal = "=" + address
            sheet.Names(key).RefersTo = rng
        except:
            #sheet.Names.Add(Name=key, RefersToLocal="=" + address, Visible=False)
            sheet.Names.Add(Name=key, RefersTo=rng, Visible=False)

    @staticmethod
    def _get_hidden_name(key, sheet, delete=True):
        try:
            # rng = sheet.Names(key).RefersToRange -> Cuts off too long range strings, below method is better
            addr = sheet.Names(key).RefersToLocal[1:]
            pos = addr.find("!")+1
            addr = addr.replace(addr[0:pos], "") #remove sheet names to support much longer range strings
            if delete:
                sheet.Names(key).Delete()
            return sheet.Range(addr)
        except:
            return None

    # @staticmethod
    # def _del_hidden_name(key, sheet):
    #     try:
    #         name = sheet.Names(key).Delete()
    #     except:
    #         pass

    @classmethod
    def prepend_text(cls, cells, application):
        input_text = bkt.ui.show_user_input("Text eingeben, der vor alle Zellen geschrieben werden soll. Existierende Formeln werden überschrieben und durch Werte ersetzt.\n\nMögliche Platzhalter: [counter], [row], [column].", "Text voranstellen", cls.last_prepend)
        if not input_text:
            return

        if not xllib.confirm_no_undo(): return

        cls.last_prepend = input_text
        number_format = application.International(xlcon.XlApplicationInternational["xlGeneralFormatName"])

        counter = 1
        for cell in cells:
            input_text_local = input_text.replace("[counter]", str(counter)).replace("[row]", str(cell.Row)).replace("[column]", str(cell.Column))
            cell.Value = input_text_local + cell.Text
            cell.NumberFormatLocal = number_format
            counter += 1

    @classmethod
    def append_text(cls, cells, application):
        input_text = bkt.ui.show_user_input("Text eingeben, der hinter alle Zellen geschrieben werden soll. Existierende Formeln werden überschrieben und durch Werte ersetzt.\n\nMögliche Platzhalter: [counter], [row], [column].", "Text anhängen", cls.last_append)
        if not input_text:
            return

        if not xllib.confirm_no_undo(): return

        cls.last_append = input_text
        number_format = application.International(xlcon.XlApplicationInternational["xlGeneralFormatName"])

        counter = 1
        for cell in cells:
            input_text_local = input_text.replace("[counter]", str(counter)).replace("[row]", str(cell.Row)).replace("[column]", str(cell.Column))
            cell.Value = cell.Text + input_text_local
            cell.NumberFormatLocal = number_format
            counter += 1

    @classmethod
    def slice_text(cls, cells, application):
        def _get_slicer(pos_text):
            input_params = pos_text.strip(' \t\n\r[]').split(":")
            
            #Extract single character
            if len(input_params) == 1:
                start = int(input_params[0])
                stop = None if start < 0 else start + 1
            
            #Extract string with start and stop
            elif len(input_params) == 2:
                start = 0 if not input_params[0] else int(input_params[0])
                stop = None if not input_params[1] else int(input_params[1])

            else:
                raise ValueError('invalid number of parameters')
            
            if stop is not None and ((start > 0 and stop > 0) or (start < 0 and stop < 0)) and start >= stop:
                raise ValueError('no text remains as start is after stop')

            return slice(start, stop)

        preview_cell = application.ActiveCell
        def _preview(sender, e):
            try:
                if text.Text == '':
                    txt_preview.Text = ''
                else:
                    s = _get_slicer(text.Text)
                    txt_preview.Text = preview_cell.Text[s]
            except:
                txt_preview.Text = "FEHLER"

        explanation = '''Start- und Stopp-Position zum Schneiden getrennt mit ":" eingeben. Ist keine Start-Position gegeben, wird diese auf 0 gesetzt. Ist keine Stopp-Position gegeben, wird diese bis Textende gesetzt. Eine negative Position bedeutet, dass diese vom Textende berechnet wird.

  Beispiel für "ABCDEF":
  [2:]   = CDEF  Entferne die beiden ersten Zeichen
  [:2]   = AB    Entferne alles nach dem zweiten Zeichen
  [-2:]  = EF    Entferne alles bis zum vorletzten Zeichen
  [2:-2] = CD    Entferne 2 Zeichen an Anfang und Ende'''

        user_form = bkt.ui.UserInputBox(explanation, "Text anhand Position schneiden")
        text = user_form._add_textbox("text", cls.last_slice_pos)
        text.TextChanged += _preview
        
        user_form._add_label("Vorschau für aktive Zelle:")
        txt_preview = user_form._add_textbox("preview")
        txt_preview.ReadOnly = True
        _preview(None, None)

        form_return = user_form.show()
        if len(form_return) == 0 or form_return["text"] == '':
            return
        
        if not xllib.confirm_no_undo(): return

        number_format = application.International(xlcon.XlApplicationInternational["xlGeneralFormatName"])
        cls.last_slice_pos = form_return["text"]

        try:
            s = _get_slicer(form_return["text"])
        except:
            bkt.helpers.message("Ungültige Eingabe!")
            return

        for cell in cells:
            cell.Value = cell.Text[s]
            cell.NumberFormatLocal = number_format

    @classmethod
    def find_and_slice_text(cls, cells, application):
        def _slice_text(initial_text, search_text, find_method, rslice):
            pos = find_method(initial_text, search_text)
            if pos == -1:
                return initial_text
            start = pos+len(search_text) if rslice else 0
            stop = None if rslice else pos
            s = slice(start, stop)
            return initial_text[s]

        preview_cell = application.ActiveCell
        def _preview(sender, e):
            try:
                find_method = str.rfind if cb_rfind.Checked else str.find
                txt_preview.Text = _slice_text(preview_cell.Text, text.Text, find_method, cb_rslice.Checked)
            except:
                txt_preview.Text = "FEHLER"

        user_form = bkt.ui.UserInputBox("Gibt Zelleninhalt von Beginn bis zum eingegebenen Text zurück. Wird der Text nicht gefunden, bleibt der Zelleninhalt unverändert.", "Text anhand Zeichen schneiden")
        text = user_form._add_textbox("text", cls.last_slice_text)
        text.TextChanged += _preview
        cb_rslice = user_form._add_checkbox("rslice", "Rechten Teil zurückgeben (ab eingegebenem Text bis Ende)")
        cb_rslice.CheckedChanged += _preview
        cb_rfind = user_form._add_checkbox("rfind", "Von rechts anfangen zu suchen")
        cb_rfind.CheckedChanged += _preview
        
        user_form._add_label("Vorschau für aktive Zelle:")
        txt_preview = user_form._add_textbox("preview")
        txt_preview.ReadOnly = True
        _preview(None, None)

        form_return = user_form.show()
        if len(form_return) == 0 or form_return["text"] == '':
            return

        if not xllib.confirm_no_undo(): return

        find_method = str.rfind if form_return["rfind"] else str.find
        
        number_format = application.International(xlcon.XlApplicationInternational["xlGeneralFormatName"])
        cls.last_slice_text = form_return["text"]

        for cell in cells:
            try:
                cell.Value = _slice_text(cell.Text, form_return["text"], find_method, form_return["rslice"])
                cell.NumberFormatLocal = number_format
            except:
                pass


    @classmethod
    def apply_formula(cls, cells, application):
        dec_sep = application.International(xlcon.XlApplicationInternational["xlDecimalSeparator"])
        preview_cell_format = application.ActiveCell.NumberFormatLocal
        preview_cell_formula = application.ActiveCell.FormulaLocal.lstrip("=")
        try:
            preview_cell_value = str(application.ActiveCell.Value2).replace('.', dec_sep)
        except:
            preview_cell_value = -2146826273 #"#Value!"

        def _preview(sender, e):
            try:
                create_formulas = text.Text[0] == "="
                formula = text.Text if create_formulas or "[cell]" in text.Text else "([cell])" + text.Text
                if create_formulas:
                    formula = formula.replace("[cell]", preview_cell_formula)
                    txt_preview.Text = formula
                else:
                    formula = formula.replace("[cell]", preview_cell_value)
                    txt_preview.Text = xllib.xls_evaluate(formula, dec_sep, preview_cell_format)

            except:
                txt_preview.Text = "FEHLER"

        user_form = bkt.ui.UserInputBox("Hier kann eine Formel auf alle markierten Zellen angewendet werden. Soll der Zelleninhalt nicht am Anfang stehen, können Sie mit dem Platzhalter [cell] arbeiten. Standardmäßig wird der resultierende Wert eingefügt (sofern die Formel nicht fehlerhaft ist). Wenn Ihre Eingabe mit '=' beginnt, wird eine Formel erstellt. In der Auswahlbox finden Sie Beispiele für mögliche Eingaben.", "Formel anwenden")
        text = user_form._add_combobox("formula", cls.last_formula, ["*100", "/100", "*(-1)", "+A1", "/SUMME(A1:A3)", "ABS([cell])", "RUNDEN([cell];2)", "ABRUNDEN([cell];2)", "AUFRUNDEN([cell];2)", "KÜRZEN([cell])", "1/[cell]", "=([cell])*100"])
        text.TextChanged += _preview
        user_form._add_checkbox("skip_existing_formulas", "Bestehende Formeln überspringen und nicht verändern")
        
        user_form._add_label("Vorschau für aktive Zelle:")
        txt_preview = user_form._add_textbox("preview")
        txt_preview.ReadOnly = True
        _preview(None, None)

        form_return = user_form.show()
        if len(form_return) == 0:
            return

        if not xllib.confirm_no_undo(): return

        err_counter = 0
        formula = form_return["formula"]
        cls.last_formula = formula

        create_formulas = formula[0] == "="
        formula = formula if create_formulas or "[cell]" in formula else "([cell])" + formula

        for cell in cells:
            if cell.Value2 is None or (cell.HasFormula and form_return["skip_existing_formulas"]):
                continue

            if cell.HasFormula:
                new_formula = formula.replace("[cell]", cell.FormulaLocal[1:])
            else:
                new_formula = formula.replace("[cell]", cell.FormulaLocal)

            try:
                if create_formulas:
                    cell.FormulaLocal = new_formula
                else:
                    cell.FormulaLocal = "=" + new_formula
                    #On error do not replace with int value of error
                    if not application.WorksheetFunction.IsError(cell):
                        cell.Value = cell.Value()
            except:
                err_counter += 1
                #bkt.helpers.exception_as_message(str(cell.AddressLocal()))

        if err_counter > 0:
            bkt.helpers.message("Fehler! Formel war auf " + str(err_counter) + " Zelle(n) nicht anwendbar.")

    @staticmethod
    def merge_cells(cells, join="\r\n"):
        if not xllib.confirm_no_undo(): return
        target_cell = next(cells)
        for cell in cells:
            target_cell.Value = target_cell.Value() + join + cell.Value()
            cell.Value = None


    @staticmethod
    def formula_to_values(areas):
        if not xllib.confirm_no_undo(): return
        for area in areas:
            area.Value = area.Value()

    @staticmethod
    def values_to_showntext(areas):
        if not xllib.confirm_no_undo(): return
        for area in areas:
            for cell in iter(area.Cells):
                if cell.Value2 is None:
                    continue
                cell.Value = "'" + cell.Text
            area.NumberFormat = "@" #Text

    @staticmethod
    def text_to_numbers(areas, application):
        if not xllib.confirm_no_undo(): return
        general_format = application.International(xlcon.XlApplicationInternational["xlGeneralFormatName"])
        for area in areas:
            #area.NumberFormatLocal = general_format
            #area.Value = application.WorksheetFunction.NumberValue( area )
            for cell in iter(area.Cells):
                if cell.HasFormula or cell.Value2 is None:
                    continue
                if cell.NumberFormat == "@": #Text
                    cell.NumberFormatLocal = general_format
                try:
                    # cell.Value = Converter.ToDouble(cell.Value())
                    cell.Value = application.WorksheetFunction.NumberValue( cell )
                except:
                    cell.Value = cell.Value()

    @staticmethod
    def numbers_to_text(areas):
        if not xllib.confirm_no_undo(): return
        for area in areas:
            area.NumberFormat = "@" #Text
            for cell in iter(area.Cells):
                if cell.Value2 is None:
                    continue
                cell.Value = "'" + cell.Text

    @staticmethod
    def text_to_datetime(areas, application):
        if not xllib.confirm_no_undo(): return
        general_format = application.International(xlcon.XlApplicationInternational["xlGeneralFormatName"])
        for area in areas:
            for cell in iter(area.Cells):
                if cell.HasFormula or cell.Value2 is None or isinstance(cell.Value(), DateTime):
                    continue
                if cell.NumberFormat == "@": #Text
                    cell.NumberFormatLocal = general_format
                try:
                    cell.Value = DateTime.Parse( cell.Text )
                except:
                    pass

    @staticmethod
    def text_to_formula(areas, application):
        if not xllib.confirm_no_undo(): return
        general_format = application.International(xlcon.XlApplicationInternational["xlGeneralFormatName"])
        for area in areas:
            #area.NumberFormatLocal = general_format
            #area.FormulaLocal = area.Value()
            for cell in iter(area.Cells):
                if cell.Text[0] != "=":
                    continue
                cell.NumberFormatLocal = general_format
                cell.FormulaLocal = cell.Value()

    @staticmethod
    def formula_to_text(areas):
        if not xllib.confirm_no_undo(): return
        for area in areas:
            #area.NumberFormat = "@" #Text
            #area.Value = area.FormulaLocal
            for cell in iter(area.Cells):
                if not cell.HasFormula:
                    continue
                cell.NumberFormat = "@" #Text
                cell.Value = "'" + cell.FormulaLocal

    @staticmethod
    def formula_to_absolute(cells, application):
        if not xllib.confirm_no_undo(): return
        for cell in cells:
            if cell.HasFormula:
                cell.Formula = Application.ConvertFormula(cell.Formula, 1, 1, 1) #xlA1, xlA1, xlAbsolute

    @staticmethod
    def formula_to_relative(cells, application):
        if not xllib.confirm_no_undo(): return
        for cell in cells:
            if cell.HasFormula:
                cell.Formula = Application.ConvertFormula(cell.Formula, 1, 1, 4) #xlA1, xlA1, xlRelative

    @staticmethod
    def prohibit_duplicates(areas, application):
        if not xllib.confirm_no_undo("Dies überschreibt bestehende Datenüberprüfungen und kann nicht rückgängig gemacht werden. Ausführen?"): return
        for area in areas:
            vali_form = "=COUNTIF(" + area.Address(True, True, 1) + "," + area.Cells(1).Address(False, False, 1) + ")=1" #xlA1
            vali_form = xllib.formula_int2local(vali_form)
            area.Validation.Delete()
            area.Validation.Add(7, 1, 1, vali_form) #xlValidateCustom, xlValidAlertStop, xlBetween
            #area.Validation.ShowError = True
            #area.Validation.ErrorTitle = "Duplicate Value"
            #area.Validation.ErrorMessage = "This value was already entered. All values must be unique. Please try again."

    @staticmethod
    def subtotal(application, selection, func="Sum"):
        try:
            selection = selection.SpecialCells(xlcon.XlCellType["xlCellTypeVisible"])
            value = str( getattr(application.WorksheetFunction, func)(selection) )
            #value = str(application.WorksheetFunction.Subtotal(xlcon.subtotalFunction[func], selection))

            value = value.replace('.', application.International(xlcon.XlApplicationInternational["xlDecimalSeparator"]))
            Forms.Clipboard.SetText(value)
        except:
            bkt.helpers.message('Fehler beim Kopieren!')
        #bkt.helpers.message('Kopiert: ' + value)

    @staticmethod
    def enabled_subtotal(application, selection):
        try:
            #count number of cells that contain numbers
            return application.WorksheetFunction.Count(selection) > 0
            #application.WorksheetFunction.Subtotal(xlcon.subtotalFunction["AVG"], selection)
            #return True
        except:
            return False

    @staticmethod
    def trim(application, areas):
        if not xllib.confirm_no_undo(): return
        for area in areas:
            area.Value = application.WorksheetFunction.Trim(area)

    @staticmethod
    def clean(application, areas):
        if not xllib.confirm_no_undo(): return
        for area in areas:
            area.Value = application.WorksheetFunction.Clean(area)

    @staticmethod
    def trim_python(application, cells):
        if not xllib.confirm_no_undo(): return
        for cell in cells:
            cell.Value = cell.Text.strip()

    @staticmethod
    def fill_down(cells, application):
        if not xllib.confirm_no_undo(): return

        # to_be_filled = None
        for cell in cells:
            if cell.Row == 1:
                continue
            if cell.Value2 is None:
                try:
                    cell.Value = cell.Offset(-1,0).MergeArea(1).Value()
                    # to_be_filled = xllib.range_union(to_be_filled, cell, application)
                except:
                    pass

        # if to_be_filled is not None:
        #     to_be_filled.FormulaR1C1 = "=R[-1]C"
        #     to_be_filled.Value = to_be_filled.Value()
       
        # for area in areas:
        #     empty_cells = area.SpecialCells(4) #xlCellTypeBlanks
        #     empty_cells.FormulaR1C1 = "=R[-1]C"
        #     area.Value = area.Value()

    @staticmethod
    def undo_fill_down(cells, application):
        if not xllib.confirm_no_undo(): return

        to_be_deleted = None
        for cell in cells:
            if cell.Row == 1:
                continue
            try:
                if cell.Value() == cell.Offset(-1,0).MergeArea(1).Value():
                    to_be_deleted = xllib.range_union(to_be_deleted, cell)
            except:
                pass

        if to_be_deleted is not None:
            to_be_deleted.Value = None

    @classmethod
    def toggle_hidden_columns(cls, sheet, application, selection):
        area = sheet.UsedRange

        #Restore hidden columns if sheet is the same
        hidden_cols = cls._get_hidden_name("BKT_HIDDEN_COLS", sheet)
        #if sheet.Name in cls.hidden_columns:
        if hidden_cols is not None:
            #sheet.Range(cls.hidden_columns[sheet.Name]).EntireColumn.Hidden = True
            #del cls.hidden_columns[sheet.Name]
            #sheet.Range(hidden_cols).EntireColumn.Hidden = True
            hidden_cols.EntireColumn.Hidden = True
            #cls._del_hidden_name("BKT_HIDDEN_COLS", sheet)

        #Show hidden columns and store them
        else:
            #hidden_cols = None
            for i in xrange(1,area.Columns.Count+1):
                if area.Columns(i).EntireColumn.Hidden:
                    hidden_cols = xllib.range_union(hidden_cols, area.Columns(i).EntireColumn)

            if hidden_cols is not None:
                hidden_cols.EntireColumn.Hidden = False
                #cls.hidden_columns[sheet.Name] = hidden_cols.AddressLocal(False, False)
                #cls._set_hidden_name("BKT_HIDDEN_COLS", sheet, hidden_cols.AddressLocal(True, True))
                cls._set_hidden_name("BKT_HIDDEN_COLS", sheet, hidden_cols)
            
            #If entire rows are selected hide them
            elif selection.Address() == selection.EntireColumn.Address():
                selection.EntireColumn.Hidden = True
            
            else:
                bkt.helpers.message("Keine ausgeblendeten Spalten im genutzten Bereich gefunden.")


    @classmethod
    def toggle_hidden_rows(cls, sheet, application, selection):
        area = sheet.UsedRange

        #Restore hidden rows if sheet is the same
        hidden_rows = cls._get_hidden_name("BKT_HIDDEN_ROWS", sheet)
        #if sheet.Name in cls.hidden_rows:
        if hidden_rows is not None:
            #sheet.Range(cls.hidden_rows[sheet.Name]).EntireRow.Hidden = True
            #del cls.hidden_rows[sheet.Name]
            #sheet.Range(hidden_rows).EntireRow.Hidden = True
            hidden_rows.EntireRow.Hidden = True
            #cls._del_hidden_name("BKT_HIDDEN_ROWS", sheet)

        #Show hidden rows and store them
        else:
            #hidden_rows = None
            for i in xrange(1,area.Rows.Count+1):
                if area.Rows(i).EntireRow.Hidden:
                    hidden_rows = xllib.range_union(hidden_rows, area.Rows(i).EntireRow)

            if hidden_rows is not None:
                hidden_rows.EntireRow.Hidden = False
                #cls.hidden_rows[sheet.Name] = hidden_rows.AddressLocal(False, False)
                #cls._set_hidden_name("BKT_HIDDEN_ROWS", sheet, hidden_rows.AddressLocal(True, True))
                cls._set_hidden_name("BKT_HIDDEN_ROWS", sheet, hidden_rows)
            
            #If entire rows are selected hide them
            elif selection.Address() == selection.EntireRow.Address():
                selection.EntireRow.Hidden = True
            
            else:
                bkt.helpers.message("Keine ausgeblendeten Zeilen im genutzten Bereich gefunden.")

    @staticmethod
    def show_all_cells(sheet):
        sheet.Columns.EntireColumn.Hidden = False
        sheet.Rows.EntireRow.Hidden = False

    @staticmethod
    def hide_unused_areas(sheet):
        selection = xllib.get_unused_ranges(sheet)

        for rng in selection:
            rng.Hidden = True


    @staticmethod
    def paste_on_visible(application, sheet, cell, pasteType=xlcon.XlPasteType["xlPasteAll"]):
        if not xllib.confirm_no_undo(): return

        xllib.freeze_app(disable_display_alerts=True)
        temporary_sheet = xllib.create_temp_sheet()

        try:
            temporary_sheet.Cells(cell.Row, cell.Column).PasteSpecial(pasteType)
            rows = temporary_sheet.UsedRange.Rows.Count
            cols = temporary_sheet.UsedRange.Columns.Count
            
            ### METHOD 1: COPY CELL BY CELL ###
            #FIXME: cache area of visible columns once determined in first loop
            # cur_cell = cell
            # for i in range(1,rows+1):
            #     for j in range(1,cols+1):
            #         temporary_sheet.UsedRange.Cells(i, j).Copy()
            #         cur_cell.PasteSpecial(pasteType)
            #         if j < cols:
            #             cur_cell = xllib.get_next_visible_cell(cur_cell, 'right')
            #     if i < rows:
            #         cur_cell = sheet.Cells(cur_cell.Row, cell.Column)
            #         cur_cell = xllib.get_next_visible_cell(cur_cell, 'bottom')
            # sheet.Range(cell, cur_cell).Select()
        
            ### METHOD 2: INSERT BLANKS AND PASTE USING SKIP BLANKS ###
            i = cell.Row
            rows_to_check = i+rows
            while i <= rows_to_check:
                if sheet.Cells(i,1).EntireRow.Hidden:
                    temporary_sheet.Cells(i,1).EntireRow.Insert()
                    rows_to_check += 1
                i += 1

            i = cell.Column
            cols_to_check = i+cols
            while i <= cols_to_check:
                if sheet.Cells(1,i).EntireColumn.Hidden:
                    temporary_sheet.Cells(1,i).EntireColumn.Insert()
                    cols_to_check += 1
                i += 1

            temporary_sheet.UsedRange.Copy()
            cell.PasteSpecial(pasteType, SkipBlanks=True)
            
        except:
            bkt.helpers.message("Sorry, etwas ist schiefgelaufen!?")

        temporary_sheet.Delete()
        xllib.unfreeze_app()


class Format(object):
    @staticmethod
    def hide_zero(cells, application, pressed):
        if not xllib.confirm_no_undo(): return
        for cell in cells:
            if pressed:
                formats = cell.NumberFormat.split(";")
                formats = formats + ['']*(3-len(formats))
                formats[2] = ''
                cell.NumberFormat = ";".join(formats)
                #cell.NumberFormat = '0;;;@'
            else:
                if cell.NumberFormat == '0;;;@':
                    cell.NumberFormatLocal = application.International(xlcon.XlApplicationInternational["xlGeneralFormatName"])
                    return

                formats = cell.NumberFormat.split(";")
                if len(formats) == 3:
                    del formats[2]
                elif len(formats) >= 4:
                    #(.*) (.*),(0*).* (.*)
                    formats[2] = "0"
                cell.NumberFormat = ";".join(formats)

    @staticmethod
    def hide_zero_pressed(cell):
        formats = cell.NumberFormat.split(";")
        return len(formats) >= 3 and formats[2] == ''
        #return cell.NumberFormat == '0;;;@'

    @staticmethod
    def hide_zero_simple(cells, application):
        if not xllib.confirm_no_undo(): return
        for cell in cells:
            cell.NumberFormat = '0;;;@'

    @staticmethod
    def number_in_thousand(cells):
        if not xllib.confirm_no_undo(): return
        #TODO: Make buttons smart: recognize number format and adjust it instead of replacing it
        for cell in cells:
            cell.NumberFormat = '_-* #.##0,0. "k"_-;-* #.##0,0. "k"_-;_-* "-"? "k"_-;_-@_-'

    @staticmethod
    def number_in_million(cells):
        if not xllib.confirm_no_undo(): return
        #TODO: Make buttons smart: recognize number format and adjust it instead of replacing it
        for cell in cells:
            cell.NumberFormat = '_-* #.##0,0.. "Mio."_-;-* #.##0,0.. "Mio."_-;_-* "-"? "Mio."_-;_-@_-'

    @staticmethod
    def merged_cells_to_center_across(cells):
        if not xllib.confirm_no_undo(): return

        for cell in cells:
            if cell.MergeCells and cell.MergeArea.Rows.Count == 1 and cell.MergeArea.HorizontalAlignment == -4108: #xlCenter
                area = cell.MergeArea
                cell.MergeCells = False
                area.HorizontalAlignment = 7 #xlCenterAcrossSelection

    @staticmethod
    def merged_cells_to_unmerged_filled(cells):
        if not xllib.confirm_no_undo(): return

        for cell in cells:
            if cell.MergeCells:
                area = cell.MergeArea
                cell.MergeCells = False
                if cell.HasFormula:
                    area.Formula = cell.Formula
                else:
                    area.Value = cell.Value()

    @staticmethod
    def horiz_align(selection, alignment, pressed):
        if not xllib.confirm_no_undo(): return
        
        if not pressed:
            selection.HorizontalAlignment = 1 #xlGeneral
        else:
            selection.HorizontalAlignment = alignment

    @staticmethod
    def horiz_align_pressed(selection, alignment):
        return selection.HorizontalAlignment == alignment


zellen_inhalt_gruppe = bkt.ribbon.Group(
    label="Zellen-Inhalte",
    image_mso="Formula",
    children=[
        bkt.ribbon.Button(
            id = 'apply_formula',
            label="Formel anwenden…",
            show_label=True,
            size='large',
            image_mso='Formula',
            supertip="Eine Formel auf alle ausgwählten Zellen anwenden.",
            on_action=bkt.Callback(CellsOps.apply_formula, cells=True, application=True),
            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
        ),
        bkt.ribbon.Menu(
            label="Textwerkzeuge",
            show_label=True,
            image_mso='FormControlEditBox',
            screentip="Verschiedene Text-Manipulationen",
            supertip="Text hinzufügen oder schneiden.",
            children=[
                bkt.ribbon.Button(
                    id = 'prepend_text',
                    label="Text voranstellen…",
                    show_label=True,
                    #image_mso='FormControlEditBox',
                    supertip="Einen Text allen ausgewählten Zellen voranstellen.",
                    on_action=bkt.Callback(CellsOps.prepend_text, cells=True, application=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.Button(
                    id = 'append_text',
                    label="Text anhängen…",
                    show_label=True,
                    #image_mso='FormControlEditBox',
                    supertip="Einen Text allen ausgewählten Zellen anhängen.",
                    on_action=bkt.Callback(CellsOps.append_text, cells=True, application=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.MenuSeparator(),
                bkt.ribbon.Button(
                    id = 'slice_text',
                    label="Text anhand Position schneiden…",
                    show_label=True,
                    #image_mso='FormControlEditBox',
                    supertip="Einen Text vorne oder hinten nach gegebener Position abschneiden.",
                    on_action=bkt.Callback(CellsOps.slice_text, cells=True, application=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.Button(
                    id = 'find_and_slice_text',
                    label="Text anhand Zeichen schneiden…",
                    show_label=True,
                    #image_mso='FormControlEditBox',
                    supertip="Einen Text vorne oder hinten nach gegebenem Zeichen abschneiden.",
                    on_action=bkt.Callback(CellsOps.find_and_slice_text, cells=True, application=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.MenuSeparator(),
                bkt.mso.control.ReplaceDialog(),
            ]
        ),
        bkt.ribbon.SplitButton(
            children=[
                bkt.ribbon.Button(
                    id = 'formula_to_values',
                    label="Formeln zu Werten",
                    show_label=True,
                    image_mso='ShowFormulas',
                    supertip="Formeln in allen ausgewählten Zellen durch jeweilige Werte ersetzen. Zellen ohne Formeln bleiben unverändert.",
                    on_action=bkt.Callback(CellsOps.formula_to_values, areas=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.Menu(children=[
                    bkt.ribbon.MenuSeparator(title="Werte/Text"),
                    bkt.ribbon.Button(
                        id = 'formula_to_values2',
                        label="Formeln zu Werten",
                        show_label=True,
                        image_mso='ShowFormulas',
                        supertip="Formeln in allen ausgewählten Zellen durch jeweilige Werte ersetzen. Zellen ohne Formeln bleiben unverändert.",
                        on_action=bkt.Callback(CellsOps.formula_to_values, areas=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'values_to_showntext',
                        label="Zu angezeigtem Text",
                        show_label=True,
                        #image_mso='PasteTextOnly',
                        supertip="Werte in allen ausgewählten Zellen durch den tatsächlich angezeigten Text ersetzen. Dabei wird das Zellenformat auf 'Text' geändert.",
                        on_action=bkt.Callback(CellsOps.values_to_showntext, areas=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.MenuSeparator(title="Zahlen/Daten"),
                    bkt.ribbon.Button(
                        id = 'numbers_to_text',
                        label="Zahlenwerte zu Text",
                        show_label=True,
                        #image_mso='PasteTextOnly',
                        supertip="Konvertiert numerische Werte (Zahlen, Datum, Zeit) in als Text gespeicherte Zahlen. Dabei wird das Zellenformat auf 'Text' geändert.",
                        on_action=bkt.Callback(CellsOps.numbers_to_text, areas=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'text_to_numbers',
                        label="Text zu Zahlen",
                        show_label=True,
                        #image_mso='PasteValues',
                        supertip="Konvertiert als Text gespeicherte Zahlen in echte Zahlen. Dabei wird das Zellenformat auf 'Standard' geändert.",
                        on_action=bkt.Callback(CellsOps.text_to_numbers, areas=True, application=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'text_to_datetime',
                        label="Text zu Datum/Zeit",
                        show_label=True,
                        #image_mso='PasteTextOnly',
                        supertip="Konvertiert als Text gespeicherte Datum- und Zeitwerte in ein echtes Datum ggf. mit Uhrzeit. Dabei wird das Zellenformat auf 'Standard' geändert.",
                        on_action=bkt.Callback(CellsOps.text_to_datetime, areas=True, application=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.MenuSeparator(title="Formeln"),
                    bkt.ribbon.Button(
                        id = 'text_to_formula',
                        label="Text zu Formeln",
                        show_label=True,
                        #image_mso='PasteFormulas',
                        supertip="Konvertiert als Text gespeicherte Formeln in echte Formeln. Dabei wird das Zellenformat auf 'Standard' geändert.",
                        on_action=bkt.Callback(CellsOps.text_to_formula, areas=True, application=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'formula_to_text',
                        label="Formeln zu Text",
                        show_label=True,
                        #image_mso='PasteTextOnly',
                        supertip="Konvertiert Formeln in als Text gespeicherte Formeln. Dabei wird das Zellenformat auf 'Text' geändert. Zellen ohne Formeln bleiben unverändert.",
                        on_action=bkt.Callback(CellsOps.formula_to_text, areas=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'formula_to_absolute',
                        label="Formeln A1 zu $A$1",
                        show_label=True,
                        #image_mso='PasteFormulas',
                        supertip="Konvertiert Referenzen in Formeln zu absoluten Referenzen.",
                        on_action=bkt.Callback(CellsOps.formula_to_absolute, cells=True, application=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'formula_to_relative',
                        label="Formeln $A$1 zu A1",
                        show_label=True,
                        #image_mso='PasteFormulas',
                        supertip="Konvertiert Referenzen in Formeln zu relativen Referenzen.",
                        on_action=bkt.Callback(CellsOps.formula_to_relative, cells=True, application=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                ])
            ]
        ),
        bkt.ribbon.SplitButton(
            children=[
                bkt.ribbon.Button(
                    id = 'cells_trim',
                    label="Glätten (Trim)",
                    show_label=True,
                    image_mso='TextDirectionContext',
                    supertip="Entferne überflüssige Leerzeichen am Anfang oder Ende aller selektierten Zellen (wie Excel-Funktion GLÄTTEN).",
                    on_action=bkt.Callback(CellsOps.trim, application=True, areas=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.Menu(children=[
                    bkt.ribbon.Button(
                        id = 'cells_trim2',
                        label="Glätten/Kürzen (Trim)",
                        show_label=True,
                        image_mso='TextDirectionContext',
                        supertip="Entferne überflüssige Leerzeichen am Anfang oder Ende aller selektierten Zellen (wie Excel-Funktion GLÄTTEN).",
                        on_action=bkt.Callback(CellsOps.trim, application=True, areas=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'cells_trim_python',
                        label="Erweitertes Glätten/Kürzen (Trim)",
                        show_label=True,
                        # image_mso='TextDirectionContext',
                        supertip="Entferne überflüssige Leerzeichen am Anfang oder Ende aller selektierten Zellen mit Pythons Strip-Funktion.",
                        on_action=bkt.Callback(CellsOps.trim_python, application=True, cells=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'cells_clean',
                        label="Säubern/Bereinigen (Clean)",
                        show_label=True,
                        #image_mso='TextDirectionContext',
                        supertip="Entferne nicht-druckbare Zeichen in allen selektierten Zellen (wie Excel-Funktion SÄUBERN).",
                        on_action=bkt.Callback(CellsOps.clean, application=True, areas=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                ])
            ]
        ),
        bkt.ribbon.SplitButton(
            children=[
                bkt.ribbon.Button(
                    id = 'cells_fill_down',
                    label="Leere Zellen nach unten füllen",
                    show_label=True,
                    image_mso='FillDown',
                    supertip="Leere Zellen im selektierten Bereich mit jeweils gefüllter Zelle darüber füllen.",
                    on_action=bkt.Callback(CellsOps.fill_down, cells=True, application=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.Menu(children=[
                    bkt.ribbon.Button(
                        id = 'cells_fill_down2',
                        label="Leere Zellen nach unten füllen",
                        show_label=True,
                        image_mso='FillDown',
                        supertip="Leere Zellen im selektierten Bereich mit jeweils gefüllter Zelle darüber füllen.",
                        on_action=bkt.Callback(CellsOps.fill_down, cells=True, application=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'cells_undo_fill_down',
                        label="Nach unten gefüllte Zellen wieder leeren",
                        show_label=True,
                        image_mso='FillUp',
                        supertip="Sich wiederholende Zellenwerte löschen, sodass nur jeweils oberste Zelle gefüllt bleibt.",
                        on_action=bkt.Callback(CellsOps.undo_fill_down, cells=True, application=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.MenuSeparator(),
                    bkt.ribbon.Button(
                        id = 'cells_merge',
                        label="Zell-Inhalte zusammenführen",
                        show_label=True,
                        # image_mso='FillUp',
                        supertip="Fügt alle Zellen in aktive Zelle getrennt mit Zeilenumbruch ein",
                        on_action=bkt.Callback(CellsOps.merge_cells, cells=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                ])
            ]
        ),
        #TODO: Zellen mit gleichen Werten verbinden
        #TODO: Zellen nicht mehr verbinden und Werte in einzelne Zellen füllen
        bkt.ribbon.SplitButton(
            get_enabled = bkt.Callback(CellsOps.enabled_subtotal, application=True, selection=True),
            children=[
                bkt.ribbon.Button(
                    id = 'selection_subtotal_sum',
                    label="Kopiere Summe markierter Zellen",
                    show_label=True,
                    image_mso='Copy',
                    supertip="Kopiere die Summe über die selektierten sichtbaren Zellen in die Zwischenablage.",
                    on_action=bkt.Callback(lambda application, selection: CellsOps.subtotal(application, selection), application=True, selection=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.Menu(children=[
                    bkt.ribbon.Button(
                        id = 'selection_subtotal_sum2',
                        label="Kopiere Summe markierter Zellen",
                        show_label=True,
                        image_mso='Copy',
                        supertip="Kopiere die Summe über die selektierten sichtbaren Zellen in die Zwischenablage.",
                        on_action=bkt.Callback(lambda application, selection: CellsOps.subtotal(application, selection), application=True, selection=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'selection_subtotal_avg',
                        label="Kopiere Mittelwert markierter Zellen",
                        show_label=True,
                        #image_mso='Copy',
                        supertip="Kopiere den Mittelwert über die selektierten sichtbaren Zellen in die Zwischenablage.",
                        on_action=bkt.Callback(lambda application, selection: CellsOps.subtotal(application, selection, "Average"), application=True, selection=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'selection_subtotal_min',
                        label="Kopiere Minimum markierter Zellen",
                        show_label=True,
                        #image_mso='Copy',
                        supertip="Kopiere das Minimum über die selektierten sichtbaren Zellen in die Zwischenablage.",
                        on_action=bkt.Callback(lambda application, selection: CellsOps.subtotal(application, selection, "Min"), application=True, selection=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'selection_subtotal_max',
                        label="Kopiere Maximum markierter Zellen",
                        show_label=True,
                        #image_mso='Copy',
                        supertip="Kopiere das Maximum über die selektierten sichtbaren Zellen in die Zwischenablage.",
                        on_action=bkt.Callback(lambda application, selection: CellsOps.subtotal(application, selection, "Max"), application=True, selection=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                ])
            ]
        ),
        bkt.ribbon.SplitButton(
            get_enabled = bkt.Callback(lambda: Forms.Clipboard.ContainsText()),
            children=[
                bkt.ribbon.Button(
                    id = 'paste_on_visible_all',
                    label="Einfügen auf sichtbare Zellen",
                    show_label=True,
                    image_mso='PasteTableByOverwritingCells',
                    supertip="Fügt den Inhalt der Zwischenablage nur auf sichtbare Zellen ein. Ausgeblendete bzw. herausgefilterte Zellen werden übersprungen.",
                    on_action=bkt.Callback(CellsOps.paste_on_visible, application=True, sheet=True, cell=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.Menu(children=[
                    bkt.ribbon.Button(
                        id = 'paste_on_visible_all2',
                        label="Einfügen auf sichtbare Zellen",
                        show_label=True,
                        image_mso='PasteTableByOverwritingCells',
                        supertip="Fügt den Inhalt der Zwischenablage nur auf sichtbare Zellen ein. Ausgeblendete bzw. herausgefilterte Zellen werden übersprungen.",
                        on_action=bkt.Callback(CellsOps.paste_on_visible, application=True, sheet=True, cell=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'paste_on_visible_values',
                        label="Werte einfügen auf sichtbare Zellen",
                        show_label=True,
                        image_mso='PasteValues',
                        supertip="Fügt den Inhalt der Zwischenablage als Werte nur auf sichtbare Zellen ein. Ausgeblendete bzw. herausgefilterte Zellen werden übersprungen.",
                        on_action=bkt.Callback(lambda application, sheet, cell: CellsOps.paste_on_visible(application, sheet, cell, xlcon.XlPasteType["xlPasteValues"]), application=True, sheet=True, cell=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'paste_on_visible_formulas',
                        label="Formeln einfügen auf sichtbare Zellen",
                        show_label=True,
                        image_mso='PasteFormulas',
                        supertip="Fügt den Inhalt der Zwischenablage als Formeln nur auf sichtbare Zellen ein. Ausgeblendete bzw. herausgefilterte Zellen werden übersprungen.",
                        on_action=bkt.Callback(lambda application, sheet, cell: CellsOps.paste_on_visible(application, sheet, cell, xlcon.XlPasteType["xlPasteFormulas"]), application=True, sheet=True, cell=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                ])
            ]
        ),
        #TODO: Upper/Lower/Proper-Case
        #TODO: Formatierung gezielt übertragen (Auswahl Zellenformat, Benutzerdefinierte Format., Datenvalidierung)
        #TODO: Benutzerdefinierte Formatierung konsolidieren (wenn farbe und typ identisch, range_union)
        #TODO: Unit/Currency Conversion
        #TODO: Alle Kommentare löschen (?)
    ]
)


zellen_format_gruppe = bkt.ribbon.Group(
    label="Zellen-Formate",
    image_mso="TableColumnsInsertLeftExcel",
    children=[
        bkt.ribbon.SplitButton(
            size="large",
            children=[
                bkt.ribbon.Button(
                    id = 'toggle_hidden_columns',
                    label="Spalten ein/ausblenden",
                    show_label=True,
                    image_mso='TableColumnsInsertLeftExcel',
                    supertip="Alle ausgeblendeten Spalten zwischen aus- und einblenden umschalten.\n\nWenn keine ausgeblendeten Spalten zwischengespeichert bzw. im Blatt vorhanden sind, und Spalten markiert sind, werden diese ausgeblendet.",
                    on_action=bkt.Callback(CellsOps.toggle_hidden_columns, sheet=True, application=True, selection=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.Menu(children=[
                    bkt.ribbon.Button(
                        id = 'toggle_hidden_columns2',
                        label="Spalten ein/ausblenden",
                        show_label=True,
                        image_mso='TableColumnsInsertLeftExcel',
                        supertip="Alle ausgeblendeten Spalten zwischen aus- und einblenden umschalten.\n\nWenn keine ausgeblendeten Spalten zwischengespeichert bzw. im Blatt vorhanden sind, und Spalten markiert sind, werden diese ausgeblendet.",
                        on_action=bkt.Callback(CellsOps.toggle_hidden_columns, sheet=True, application=True, selection=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'toggle_hidden_rows',
                        label="Zeilen ein/ausblenden",
                        show_label=True,
                        image_mso='TableRowsInsertBelowExcel',
                        supertip="Alle ausgeblendeten Zeilen zwischen aus- und einblenden umschalten.\n\nWenn keine ausgeblendeten Zeilen zwischengespeichert bzw. im Blatt vorhanden sind, und Zeilen markiert sind, werden diese ausgeblendet.",
                        on_action=bkt.Callback(CellsOps.toggle_hidden_rows, sheet=True, application=True, selection=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.MenuSeparator(),
                    bkt.ribbon.Button(
                        id = 'show_all_cells',
                        label="Alle Spalten und Zeilen einblenden",
                        show_label=True,
                        #image_mso='TableInsertMultidiagonalCell',
                        supertip="Alle ausgeblendeten Spalten und Zeilen wieder einblenden.",
                        on_action=bkt.Callback(CellsOps.show_all_cells, sheet=True, require_worksheet=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'hide_unused_areas',
                        label="Ungenutzten Bereich ausblenden",
                        show_label=True,
                        #image_mso='ViewGridlinesToggleExcel',
                        supertip="Alle Spalten und Zeilen des nicht genutzten Bereichs ausblenden.",
                        on_action=bkt.Callback(CellsOps.hide_unused_areas, sheet=True, require_worksheet=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                ])
            ]
        ),
        bkt.ribbon.Button(
            id = 'prohibit_duplicates',
            label="Duplikate verbieten",
            show_label=True,
            image_mso='DataValidation',
            supertip="Verbietet Duplikate innerhalb der jeweils selektierten Bereiche über eine Datenüberprüfung. Bestehende Datenüberprüfungen werden dabei überschrieben.",
            on_action=bkt.Callback(CellsOps.prohibit_duplicates, areas=True, application=True),
            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
        ),
        bkt.ribbon.SplitButton(
            children=[
                bkt.ribbon.ToggleButton(
                    id = 'hide_zero',
                    label="0 ein-/ausblenden",
                    show_label=True,
                    image='hide_zero',
                    screentip="Nullwerte verstecken",
                    supertip="Per Zellen-Format 0-Werte ausblenden und wieder einblenden. Dabei wird versucht, dass bestehende Zellen-Format zu erkennen und entsprechend anzupassen.",
                    on_toggle_action=bkt.Callback(Format.hide_zero, cells=True, application=True),
                    get_pressed=bkt.Callback(Format.hide_zero_pressed, cell=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.Menu(children=[
                    bkt.ribbon.Button(
                        id = 'hide_zero_simple',
                        label="0-Werte verstecken",
                        show_label=True,
                        image='hide_zero',
                        screentip="Nullwerte verstecken",
                        supertip="Per Zellen-Format ('0;;;@') 0-Werte ausblenden. Bestehendes Zellen-Format wird überschrieben.",
                        on_action=bkt.Callback(lambda cells, application: Format.hide_zero_simple(cells, application), cells=True, application=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'number_in_thousand',
                        label="Tausender zu 0,0 k",
                        show_label=True,
                        image='number_in_thousand',
                        screentip="Tausenderbeträge übersichtlich darstellen",
                        supertip="Per Zellen-Format Tausenderbeträge als x,x k. anzeigen. Bestehendes Zellen-Format wird überschrieben.",
                        on_action=bkt.Callback(Format.number_in_thousand, cells=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'number_in_million',
                        label="Mio.-Werte zu 0,0 M",
                        show_label=True,
                        image='number_in_million',
                        screentip="Millionenbeträge übersichtlich darstellen",
                        supertip="Per Zellen-Format Millionenbeträge als x,x Mio. anzeigen. Bestehendes Zellen-Format wird überschrieben.",
                        on_action=bkt.Callback(Format.number_in_million, cells=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    )
                ]),
            ]
        ),
        bkt.ribbon.Menu(
            label="Zellen u. Ausrichtung",
            show_label=True,
            image_mso='AlignJustify',
            screentip="Verbundene Zellen ersetzen und ungewöhnliche Textausrichtungen nutzen",
            #supertip="Text hinzufügen oder schneiden.",
            children=[
                bkt.ribbon.Button(
                    id = 'merged_cells_to_center_across',
                    label="Verbundene Zellen ersetzen durch Über Auswahl zentrieren",
                    show_label=True,
                    #image_mso='FormControlEditBox',
                    supertip="Ersetzt verbundene Zellen innerhalb der aktuellen Auswahl durch die horizontale Ausrichtung 'Über Auswahl zentrieren', sofern die verbundenen Zellen aus einer Zeile bestehen und bisher zentriert formatiert waren.",
                    on_action=bkt.Callback(Format.merged_cells_to_center_across, cells=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.Button(
                    id = 'merged_cells_to_unmerged_filled',
                    label="Verbundene Zellen aufheben und Inhalte verteilen",
                    show_label=True,
                    #image_mso='FormControlEditBox',
                    supertip="Hebt verbundene Zellen innerhalb der aktuellen Auswahl auf und fügt den ursprünglichen Zelleninhalt in alle Zellen ein.",
                    on_action=bkt.Callback(Format.merged_cells_to_unmerged_filled, cells=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.MenuSeparator(),
                # bkt.mso.control.AlignJustify,
                # bkt.mso.control.ParagraphDistributed,
                bkt.ribbon.ToggleButton(
                    id = 'halign_justify',
                    label="Blocksatz",
                    show_label=True,
                    image_mso='AlignJustify',
                    supertip="Ausgewählte zellen als Blocksatz ausrichten.",
                    on_toggle_action=bkt.Callback(lambda selection, pressed: Format.horiz_align(selection, -4130, pressed), selection=True), #xlHAlignJustify
                    get_pressed=bkt.Callback(lambda selection: Format.horiz_align_pressed(selection, -4130), selection=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.ToggleButton(
                    id = 'halign_distributed',
                    label="Gleichmäßig verteilt",
                    show_label=True,
                    image_mso='ParagraphDistributed',
                    supertip="Ausgewählte zellen horizontal verteilt ausrichten (extremer Blocksatz).",
                    on_toggle_action=bkt.Callback(lambda selection, pressed: Format.horiz_align(selection, -4117, pressed), selection=True), #xlHAlignDistributed
                    get_pressed=bkt.Callback(lambda selection: Format.horiz_align_pressed(selection, -4117), selection=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.ToggleButton(
                    id = 'halign_centeracross',
                    label="Über Auswahl zentrieren",
                    show_label=True,
                    #image_mso='FormControlEditBox',
                    supertip="Ausgewählte zellen 'Über Auswahl zentriert' ausrichten, d.h es werden verbundene Zellen simuliert.",
                    on_toggle_action=bkt.Callback(lambda selection, pressed: Format.horiz_align(selection, 7, pressed), selection=True), #xlHAlignCenterAcrossSelection 
                    get_pressed=bkt.Callback(lambda selection: Format.horiz_align_pressed(selection, 7), selection=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.ToggleButton(
                    id = 'halign_fill',
                    label="Ausfüllen (Text bis Ende wiederholen)",
                    show_label=True,
                    #image_mso='FormControlEditBox',
                    supertip="Ausgewählte zellen 'Ausfüllen', d.h. Zelleninhalt wird optisch wiederholt über die gesamte Zellenbreite.",
                    on_toggle_action=bkt.Callback(lambda selection, pressed: Format.horiz_align(selection, 5, pressed), selection=True), #xlHAlignFill
                    get_pressed=bkt.Callback(lambda selection: Format.horiz_align_pressed(selection, 5), selection=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
            ]
        ),
        bkt.ribbon.DialogBoxLauncher(idMso='CellAlignmentOptions')
    ]
)
