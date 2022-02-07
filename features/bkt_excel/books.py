# -*- coding: utf-8 -*-
'''
Created on 2017-07-18
@author: Florian Stallmann
'''

from __future__ import absolute_import

import os.path #required to split filenames
import tempfile #required to copy color scheme
import logging

import System.Array #required to copy sheets

import bkt
import bkt.library.excel.helpers as xllib
import bkt.library.excel.constants as xlcon

import bkt.dotnet as dotnet
Forms = dotnet.import_forms() #required for save as dialog

class BooksOps(object):
    @staticmethod
    def reset_workbook(workbook, application):
        # Show warning regardless of setting "ignore warnings"
        if not bkt.message.confirmation("Dies löscht alles Änderungen seit dem letzten Speichern und kann nicht rückgängig gemacht werden. Ausführen?"): return
        wb_path = workbook.FullName
        wb_updatelinks = workbook.UpdateLinks
        wb_readonly = workbook.ReadOnly
        active_sheet = workbook.ActiveSheet.Name
        workbook.Close(False)
        application.Workbooks.Open(wb_path, wb_updatelinks, wb_readonly, IgnoreReadOnlyRecommended=True) #Open( FileName , UpdateLinks , ReadOnly , Format , Password , WriteResPassword , IgnoreReadOnlyRecommended , Origin , Delimiter , Editable , Notify , Converter , AddToMru , Local , CorruptLoad )
        try:
            application.ActiveWorkbook.Sheets[active_sheet].Activate()
        except:
            pass

    @staticmethod
    def theme_apply(workbook):
        fileDialog = Forms.OpenFileDialog()
        fileDialog.Filter = "Excel (*.xlsx;*.xls;*.xltx;*.xlt)|*.xlsx;*.xls;*.xltx;*.xlt|Alle Dateien (*.*)|*.*"
        fileDialog.Title = "Excel-Datei auswählen"

        if workbook.Path:
            fileDialog.InitialDirectory = workbook.Path + '\\'

        if not fileDialog.ShowDialog() == Forms.DialogResult.OK:
            return
        schemePath = fileDialog.FileName

        try:
            workbook.ApplyTheme(schemePath)
            bkt.message("Theme erfolgreich angewendet!")
        except:
            bkt.message("Fehler beim Anwenden des Themes! Eventuell ist die Arbeitsmappe oder einzelne Blätter geschützt.")

    @classmethod
    def theme_color_export(cls, workbook):
        cls._theme_export(workbook, "color")
    @classmethod
    def theme_color_import(cls, workbook):
        cls._theme_import(workbook, "color")

    @classmethod
    def theme_font_export(cls, workbook):
        cls._theme_export(workbook, "font")
    @classmethod
    def theme_font_import(cls, workbook):
        cls._theme_import(workbook, "font")

    @staticmethod
    def _theme_export(workbook, theme_type):
        #Using SaveFileDialog because application.FileDialog does not support xml-File-Filter
        fileDialog = Forms.SaveFileDialog()
        fileDialog.Filter = "XML (*.xml)|*.xml|Alle Dateien (*.*)|*.*"
        if workbook.Path:
            fileDialog.InitialDirectory = workbook.Path + '\\'
        if theme_type == "color":
            fileDialog.FileName = 'colorscheme.xml'
        else:
            fileDialog.FileName = 'fontscheme.xml'
        fileDialog.Title = "Speicherort auswählen"
        fileDialog.RestoreDirectory = True

        if not fileDialog.ShowDialog() == Forms.DialogResult.OK:
            return
        schemePath = fileDialog.FileName

        if theme_type == "color":
            workbook.Theme.ThemeColorScheme.Save(schemePath)
        else:
            workbook.Theme.ThemeFontScheme.Save(schemePath)

        bkt.message("Scheme-Datei erfolgreich exportiert!")

    @staticmethod
    def _theme_import(workbook, theme_type):
        fileDialog = Forms.OpenFileDialog()
        fileDialog.Filter = "XML (*.xml)|*.xml|Alle Dateien (*.*)|*.*"
        fileDialog.Title = "XML Scheme-Datei auswählen"

        if workbook.Path:
            fileDialog.InitialDirectory = workbook.Path + '\\'

        if not fileDialog.ShowDialog() == Forms.DialogResult.OK:
            return
        schemePath = fileDialog.FileName

        try:
            if theme_type == "color":
                workbook.Theme.ThemeColorScheme.Load(schemePath)
            else:
                workbook.Theme.ThemeFontScheme.Load(schemePath)
            bkt.message("Scheme-Datei erfolgreich importiert!")
        except:
            bkt.message("Fehler beim Import!")

    @staticmethod
    def copy_selected_sheets(workbook, application):
        #Workaround to copy multiple sheets with tables at once: https://blogs.office.com/en-us/2009/08/31/copying-worksheets-with-a-list-or-table/
        xllib.freeze_app(disable_calculation=True)
        try:
            tmp_active_window = application.ActiveWindow
            tmp_window = workbook.NewWindow()
            tmp_active_window.SelectedSheets.Copy()
            tmp_window.Close()
            #Copy color scheme
            fullFileName = os.path.join(tempfile.gettempdir(), "bkt_colorscheme.xml")
            workbook.Theme.ThemeColorScheme.Save(fullFileName)
            application.ActiveWorkbook.Theme.ThemeColorScheme.Load(fullFileName)
        except:
            bkt.helpers.exception_as_message()
        xllib.unfreeze_app()

    @staticmethod
    def close_workbooks(workbook, application):
        workbooks = list(iter(application.Workbooks))
        for cur_wb in workbooks:
            if workbook.Name == cur_wb.Name:
                continue
            cur_wb.Close()

    @staticmethod
    def _get_worksheet_list(sheets, selected_sheets, include_hidden=False):
        selected_sheets = [sheet.Name for sheet in selected_sheets]
        sel_worksheets = []
        for sheet in sheets:
            #exclude strange worksheet types and very hidden sheets
            if getattr(sheet, "Type", None) != xlcon.XlSheetType["xlWorksheet"] or sheet.Visible == xlcon.XlSheetVisibility["xlSheetVeryHidden"]:
                continue
            if len(selected_sheets) == 1:
                if sheet.Visible == xlcon.XlSheetVisibility["xlSheetVisible"]:
                    sel_worksheets.append((sheet.Name, True))
                elif include_hidden:
                    sel_worksheets.append((sheet.Name, False))
            else:
                sel_worksheets.append((sheet.Name, sheet.Name in selected_sheets))
        return sel_worksheets

    @classmethod
    def copy_sheets_and_save(cls, sheets, selected_sheets, workbook, application):
        if not workbook.Path:
            bkt.message("Bitte erst die Arbeitsmappe speichern!")
            return

        #generate list for checked listbox, if multiple sheets are selected mark them as checked, otherwise all are checked
        sel_worksheets = cls._get_worksheet_list(sheets, selected_sheets)

        user_form = bkt.ui.UserInputBox("Diese Funktion kopiert jedes Blatt in eine einzelne Datei. Bitte die Arbeitsblätter zum Speichern auswählen:", "Arbeitsblätter getrennt speichern")
        user_form._add_checked_listbox("sel_worksheets", sel_worksheets)
        user_form._add_label("Bitte den Dateinamen eingeben. Erlaube Platzhalter: [counter], [workbook], [sheet]")
        user_form._add_combobox("filename", "[counter]_[sheet].xlsx", ["[counter]_[sheet].xlsx", "[workbook]_[sheet].xlsx", "[sheet].xlsx"])
        user_form._add_checkbox("do_not_close", "Arbeitsmappen geöffnet lassen")
        form_return = user_form.show()
        if len(form_return) == 0:
            return

        #worksheets to be consolidated
        sel_worksheets = form_return["sel_worksheets"]
        if len(sel_worksheets) == 0:
            bkt.message("Keine Blätter ausgewählt.")
            return

        err_counter = 0
        input_filename = form_return["filename"].replace("[workbook]", workbook.Name)

        counter = 0
        for sheet in sheets:
            if sheet.Name not in sel_worksheets:
                continue
            try:
                sheet.Copy()
                counter += 1
                new_filename = input_filename.replace("[counter]", str(counter)).replace("[sheet]", sheet.Name)
                application.ActiveWorkbook.SaveAs(workbook.Path + '\\' + new_filename)
                if not form_return["do_not_close"]:
                    application.ActiveWorkbook.Close(True)
            except:
                err_counter += 1

        if err_counter > 0:
            bkt.message("Fehler! " + str(err_counter) + " Blatt/Blätter konnte(n) nicht kopiert werden.")

    @classmethod
    def consolidate_file_workbooks(cls, workbook, sheets, application):
        fileDialog = application.FileDialog(3) #msoFileDialogFilePicker
        fileDialog.Filters.Add("Excel", "*.xls; *.xlsx; *.xlsm; *.xlsb", 1)
        fileDialog.Filters.Add("Alle Dateien", "*.*", 2)
        if workbook.Path:
            fileDialog.InitialFileName = workbook.Path + '\\'
        fileDialog.title = "Excel-Dateien auswählen"
        fileDialog.AllowMultiSelect = True

        if fileDialog.Show() == 0: #msoFalse
            return

        application.StatusBar = "Einstellungen für Konsolidierung"
        workbooks = [(wb, True) for wb in list(iter(fileDialog.SelectedItems))]
        cls._consolidate_workbooks(workbooks, [sheet.Name for sheet in sheets], application)
        application.StatusBar = False

    @classmethod
    def consolidate_open_workbooks(cls, workbook, sheets, application):
        application.StatusBar = "Einstellungen für Konsolidierung"
        workbooks = [(wb.Name, True) for wb in list(iter(application.Workbooks))]
        cls._consolidate_workbooks(workbooks, [sheet.Name for sheet in sheets], application)
        application.StatusBar = False


    @staticmethod
    def _consolidate_workbooks(workbooks, sheets, application):
        user_form = bkt.ui.UserInputBox("Diese Funktion kopiert die Blätter mehrerer Arbeitsmappen in eine Mappe. Bitte die Arbeitsmappen zur Konsolidierung auswählen:", "Arbeitsmappen konsolidieren")
        user_form._add_checked_listbox("sel_workbooks", workbooks)
        user_form._add_label("Komma-getrennte Liste von Blattnamen, die ausschließlich konsolidiert werden:")
        user_form._add_combobox("include_sheets", dropdown=sheets)
        user_form._add_label("Komma-getrennte Liste von Blattnamen, die nicht konsolidiert werden:")
        user_form._add_combobox("exclude_sheets", dropdown=sheets)
        user_form._add_checkbox("deduplicate", "Doppelte Blätter bzw. Blattnamen nur einmal kopieren")
        user_form._add_checkbox("include_hidden", "Versteckte Blätter kopieren", True)
        user_form._add_checkbox("add_wb_name", "Name der Arbeitsmappe vor Blattnamen schreiben")
        user_form._add_checkbox("add_report", "Neues Blatt mit Zusammenfassung der Konsolidierung einfügen", True)
        form_return = user_form.show()
        if len(form_return) == 0:
            return

        #sel_workbooks = list(form_return["sel_workbooks"].Item)
        sel_workbooks = form_return["sel_workbooks"]
        if len(sel_workbooks) == 0:
            bkt.message("Keine Arbeitsmappen ausgewählt.")
            return

        if form_return["exclude_sheets"] == '':
            exclude_sheets = []
        else:
            exclude_sheets = form_return["exclude_sheets"].split(',')
            exclude_sheets = map(str.strip, exclude_sheets)

        if form_return["include_sheets"] == '':
            include_sheets = []
        else:
            include_sheets = form_return["include_sheets"].split(',')
            include_sheets = map(str.strip, include_sheets)

        xllib.freeze_app(disable_display_alerts=True)
        application.StatusBar = "Konsolidiere Mappen"

        #Create new workbook and store created default sheets
        new_wb = application.Workbooks.Add()
        new_wb_sheets = list(iter(new_wb.Sheets))

        #Rename created default sheets
        for i, sheet in enumerate(new_wb_sheets):
            sheet.Name = "BKT_TEMP_"  + str(i)

        err_counter = 0
        report = []
        all_sheets = set()
        for wb_name in sel_workbooks:
            application.StatusBar = "Konsolidiere Mappe " + wb_name
            #Test if workbook is open, otherwise open it in read-only mode
            close = False
            try:
                cur_wb = application.Workbooks[os.path.basename(wb_name)]
            except:
                try:
                    cur_wb = application.Workbooks.Open(wb_name, 0, True, IgnoreReadOnlyRecommended=True) #Open( FileName , UpdateLinks , ReadOnly , Format , Password , WriteResPassword , IgnoreReadOnlyRecommended , Origin , Delimiter , Editable , Notify , Converter , AddToMru , Local , CorruptLoad )
                    close = True
                except:
                    err_counter +=1
                    report.append((wb_name, "", "", "FEHLER BEIM ÖFFNEN"))
                    #bkt.helpers.exception_as_message()
                    logging.exception("error opening workbook")
                    continue

            err_counter_sheets = 0
            sheets_to_copy = []
            orig_sheet_names = []
            #Iterate sheets and determine which one to copy and save original name
            for cur_sh in list(iter(cur_wb.Sheets)):
                if cur_sh.Name in exclude_sheets or \
                (len(include_sheets) > 0 and cur_sh.Name not in include_sheets) or \
                (cur_sh.Visible != xlcon.XlSheetVisibility["xlSheetVisible"] and not form_return["include_hidden"]):
                    continue

                if form_return["deduplicate"] and cur_sh.Name in all_sheets:
                    report.append((cur_wb.Name, cur_sh.Name, "", "DUPLIKAT ÜBERSPRUNGEN"))
                    continue
                
                all_sheets.add(cur_sh.Name)
                sheets_to_copy.append(cur_sh.Index)
                orig_sheet_names.append(cur_sh.Name)

                ### OLD METHOD (copy sheets individually):
                # try:
                #     #Copy sheet, store original name, add workbook name to original name if required
                #     orig_sheet_name = cur_sh.Name
                #     cur_sh.Copy(After=new_wb.Sheets(new_wb.Sheets.Count))
                #     new_sh = new_wb.Sheets(new_wb.Sheets.Count)
                #     if(form_return["add_wb_name"]):
                #         new_name = cur_wb.Name.rsplit('.', 1)[0] + " " + orig_sheet_name
                #         xllib.rename_sheet(new_sh, new_name)
                #         #new_sh.Name = new_name[:31] #max sheet name length is 31 characters
                #     report.append((cur_wb.Name, orig_sheet_name, new_sh.Name, "OK"))
                # except:
                #     err_counter_sheets += 1
                #     report.append((cur_wb.Name, cur_sh.Name, "", "FEHLER"))
                #     #bkt.helpers.exception_as_message()

            #Copy and rename sheets
            if len(sheets_to_copy) > 0:
                cur_index = new_wb.Sheets.Count
                cur_wb_name = cur_wb.Name.rsplit('.', 1)[0] #filename without ending
                try:
                    #New window as workaround to copy multiple sheets with tables
                    tmp_window = cur_wb.NewWindow()
                    cur_wb.Sheets(System.Array[int](sheets_to_copy)).Copy(After=new_wb.Sheets(cur_index))
                    tmp_window.Close()
                    #Rename sheets
                    for i in range(cur_index+1, cur_index+len(sheets_to_copy)+1):
                        orig_sheet_name = orig_sheet_names[i-1-cur_index]
                        new_sh = new_wb.Sheets(i)
                        if(form_return["add_wb_name"]):
                            new_name = cur_wb_name + " " + orig_sheet_name
                        else:
                            new_name = orig_sheet_name
                        xllib.rename_sheet(new_sh, new_name)
                        report.append((cur_wb.Name, orig_sheet_name, new_sh.Name, "OK"))
                except:
                    err_counter_sheets += 1
                    report.append((cur_wb.Name, "", "", "FEHLER BEIM KOPIEREN"))
                    #bkt.helpers.exception_as_message()
                    logging.exception("error copying workbook")

            if err_counter_sheets > 0:
                err_counter +=1

            if close:
                cur_wb.Close(False)

        #Delete created default sheets
        for sheet in new_wb_sheets:
            sheet.Delete()

        #Generate report sheet
        if(form_return["add_report"]):
            list_sheet = new_wb.Worksheets.Add(Before=new_wb.Worksheets(1))
            #list_sheet.Name = "BKT KONSOLIDIERUNG"
            xllib.rename_sheet(list_sheet, "BKT KONSOLIDIERUNG")
            list_sheet.Cells(1,1).Value = "Arbeitsmappe"
            list_sheet.Cells(1,2).Value = "Blattname (alt)"
            list_sheet.Cells(1,3).Value = "Blattname (neu)"
            list_sheet.Cells(1,4).Value = "Status"
            list_sheet.Range("A1:D1").Font.Bold = True
            cur_row = 2
            for wb, sh_old, sh_new, status in report:
                new_wb.Sheets(1).Cells(cur_row, 1).Value = wb
                new_wb.Sheets(1).Cells(cur_row, 2).Value = sh_old
                new_wb.Sheets(1).Cells(cur_row, 3).Value = sh_new
                new_wb.Sheets(1).Cells(cur_row, 4).Value = status
                cur_row += 1
            list_sheet.UsedRange.Columns.AutoFit()

        application.StatusBar = False
        xllib.unfreeze_app()

        if err_counter > 0:
            bkt.message("Fehler! " + str(err_counter) + " Arbeitemappe(n) konnte(n) nicht oder nur teilweise konsolidiert werden.")

    @classmethod
    def consolidate_worksheets(cls, workbook, sheet, sheets, selected_sheets, application):
        dropdown = ["[UsedRange]", "[TableRange]", "[Selection]", sheet.UsedRange.AddressLocal(False, False)]
        #TODO: [TableRange] einfügen mit automatischer Erkennung der Tabellen in einem Sheet inkl. Kopfzeile und Ergebniszeile

        #if area selected, take address address as default
        selection = application.ActiveWindow.RangeSelection
        if selection and selection.Cells.Count > 1:
            default_range = selection.AddressLocal(False, False)
            default_skip = 0
            dropdown.append(default_range)
        else:
            default_range = "[UsedRange]"
            default_skip = 1
        
        #Add ranges of defined names to dropdown
        for name in list(iter(workbook.Names)):
            try:
                if name.Visible:
                    # dropdown.append(name.RefersToRange.AddressLocal(False, False))
                    _ = name.RefersToRange #test if name is refering to a range (and not a function)
                    dropdown.append(name.NameLocal)
            except:
                pass

        consol_mode_list = ["Untereinander (Zeilen)", "Nebeneinander (Spalten)"]
        pastemode_list = ["Alles einfügen", "Werte", "Werte und Zahlenformate", "Werte und Quellformatierung", "Formeln", "Formeln und Zahlenformate", "Formeln und Quellformatierung", "Referenzen", "Referenzen und Quellformatierung"]
        pastemode_values = [
            [xlcon.XlPasteType["xlPasteAll"]], 
            [xlcon.XlPasteType["xlPasteValues"]], 
            [xlcon.XlPasteType["xlPasteValuesAndNumberFormats"]], 
            [xlcon.XlPasteType["xlPasteValues"], xlcon.XlPasteType["xlPasteFormats"]], 
            [xlcon.XlPasteType["xlPasteFormulas"]], 
            [xlcon.XlPasteType["xlPasteFormulasAndNumberFormats"]], 
            [xlcon.XlPasteType["xlPasteFormulas"], xlcon.XlPasteType["xlPasteFormats"]], 
            ["PASTE_LINK"],
            ["PASTE_LINK", xlcon.XlPasteType["xlPasteFormats"]]
        ]

        #generate list for checked listbox, if multiple sheets are selected mark them as checked, otherwise all are checked
        sel_worksheets = cls._get_worksheet_list(sheets, selected_sheets)
        #TODO: allow re-ordering ot sheets

        user_form = bkt.ui.UserInputBox("Diese Funktion kopiert die Zellen mehrerer Arbeitsblätter in ein Blatt. Bitte die Arbeitsblätter zur Konsolidierung auswählen:", "Arbeitsblätter konsolidieren")
        user_form._add_checked_listbox("sel_worksheets", sel_worksheets)
        user_form._add_radio_buttons("consolidate_columns", "Konsolidierungs-Modus", consol_mode_list)
        user_form._add_label("Bereich zum Konsolidieren eingeben, d.h. eine benannter Bereich oder eine Adresse wie A1:D5. [UsedRange] ermittelt automatisch den genutzten Bereich je Arbeitsblatt. [Selection] nimmt den jeweils im Sheet ausgewählten Bereich.")
        user_form._add_combobox("range", default_range, sorted(set(dropdown)))
        user_form._add_label("Zeilen/Spalten überspringen, z.B. für Titelzeilen:")
        user_form._add_spinner("skip_rows", default_skip, max_value=sheet.Cells.Rows.Count-1)
        user_form._add_checkbox("insert_skip_rows", "Übersprungene Zeilen/Spalten aus erstem Blatt einfügen (bspw. Überschriften)", True)
        user_form._add_label("Zeilen/Spalten abtrennen, z.B. für Ergebnis-/Summenzeilen:")
        user_form._add_spinner("cut_rows", 0, max_value=sheet.Cells.Rows.Count-1)
        user_form._add_checkbox("insert_sheet_names", "Jeweiligen Blattnamen als erste Spalte/Zeile einfügen")
        user_form._add_label("Einfügemodus:")
        user_form._add_combobox("pastemode", dropdown=pastemode_list, selected_index=0, editable=False, return_value="SelectedIndex")
        form_return = user_form.show()
        if len(form_return) == 0:
            return

        #worksheets to be consolidated
        sel_worksheets = form_return["sel_worksheets"]
        if len(sel_worksheets) == 0:
            bkt.message("Keine Blätter ausgewählt.")
            return

        #Number of skipped rows
        try:
            skip_rows = form_return["skip_rows"]
            skip_rows = 0 if skip_rows == '' else int(skip_rows)
            cut_rows = form_return["cut_rows"]
            cut_rows = 0 if cut_rows == '' else int(cut_rows)
        except:
            bkt.message("Fehler, Eingabe ist keine Zahl!")
            return
        err_counter = 0

        insert_skip_rows = form_return["insert_skip_rows"]
        insert_column = 1 if not form_return["insert_sheet_names"] else 2
        insert_row = 1 if insert_skip_rows else skip_rows+1

        xllib.freeze_app(disable_display_alerts=True, disable_events=True)
        application.StatusBar = "Konsolidiere Blätter"

        consolidate_columns = form_return["consolidate_columns"] == consol_mode_list[1]
        paste_types =  pastemode_values[form_return["pastemode"]]
        new_sheet = workbook.Worksheets.Add()
        #new_sheet.Name = "BKT KONSOLIDIERUNG"
        xllib.rename_sheet(new_sheet, "BKT KONSOLIDIERUNG")
        if consolidate_columns:
            cur_cell = new_sheet.Cells(insert_column, insert_row)
        else:
            cur_cell = new_sheet.Cells(insert_row, insert_column)
        # for sheet in sheets:
        for sheet_name in sel_worksheets:
            try:
                sheet = workbook.Worksheets[sheet_name]
                application.StatusBar = "Konsolidiere Blatt " + sheet.Name
                # if sheet.Name not in sel_worksheets:
                #     continue
                #Determine range to copy
                form_range = form_return["range"]
                if form_range == "[UsedRange]":
                    rng_to_copy = sheet.UsedRange
                elif form_range == "[TableRange]":
                    rng_to_copy = sheet.ListObjects[1].Range
                elif form_range == "[Selection]":
                    sheet.Activate()
                    rng_to_copy = application.ActiveWindow.RangeSelection
                    new_sheet.Activate()
                else:
                    rng_to_copy = sheet.Range(form_range)
                
                #FIXME: Rows.Count does not return correct value for multiple areas (max of all area rows)
                if consolidate_columns:
                    rows_to_insert = rng_to_copy.Columns.Count
                else:
                    rows_to_insert = rng_to_copy.Rows.Count
                rows_to_skip = skip_rows

                #Reduce rows if rows should be cut
                if cut_rows > 0:
                    rows_to_insert -= cut_rows

                #Reduce rows if rows should be skipped
                if skip_rows > 0:
                    if insert_skip_rows:
                        #Insert skipped rows in first iteration => no shift in first iteration
                        insert_skip_rows = False
                        rows_to_skip = 0
                    else:
                        rows_to_insert -= skip_rows
                
                #If no rows to insert continue
                if rows_to_insert <= 0:
                    continue

                #FIXME: Resize does not work for multiple areas! This is a workaround until method can handle multiple areas.
                if rng_to_copy.Areas.Count == 1:
                    if consolidate_columns:
                        rng_to_copy = rng_to_copy.Offset(0,rows_to_skip).Resize(ColumnSize=rows_to_insert)
                    else:
                        rng_to_copy = rng_to_copy.Offset(rows_to_skip,0).Resize(RowSize=rows_to_insert)
                
                #Copy action
                rng_to_copy.Copy()

                #Paste values/formats on current cell
                for ptype in paste_types:
                    if ptype == "PASTE_LINK":
                        cur_cell.Select()
                        cur_cell.Parent.Paste(Link=True)
                    else:
                        cur_cell.PasteSpecial(ptype)
                
                if consolidate_columns:
                    rows_pasted = new_sheet.UsedRange.Column + new_sheet.UsedRange.Columns.Count - cur_cell.Column
                else:
                    rows_pasted = new_sheet.UsedRange.Row + new_sheet.UsedRange.Rows.Count - cur_cell.Row

                #Insert sheet name as first column
                if form_return["insert_sheet_names"]:
                    if consolidate_columns:
                        cur_cell.Offset(-1,0).Resize(ColumnSize=rows_pasted).Value = sheet.Name
                    else:
                        cur_cell.Offset(0,-1).Resize(RowSize=rows_pasted).Value = sheet.Name
                    #cur_cell.Offset(0,-1).Resize(rng_to_copy.Rows.Count).Value = sheet.Name

                if consolidate_columns:
                    cur_cell = new_sheet.Cells(insert_column, cur_cell.Column + rows_pasted)
                else:
                    cur_cell = new_sheet.Cells(cur_cell.Row + rows_pasted, insert_column)
            except:
                #bkt.helpers.exception_as_message()
                logging.exception("error consolidating sheet %s", sheet_name)
                err_counter += 1
        
        new_sheet.UsedRange.Columns.AutoFit()
        new_sheet.Range("A1").Select()
        new_sheet.Activate()

        application.StatusBar = False
        xllib.unfreeze_app()

        if err_counter > 0:
            bkt.message("Fehler! " + str(err_counter) + " Blatt/Blätter konnte(n) nicht konsolidiert werden.")


mappen_gruppe = bkt.ribbon.Group(
    id="group_workbook",
    label="Arbeitsmappe",
    image_mso="ExportExcel",
    auto_scale=True,
    children=[
        bkt.ribbon.Button(
            id = 'reset_workbook',
            label="Datei zurücksetzen",
            show_label=True,
            size="large",
            image_mso='ResetCurrentView',
            supertip="Datei auf den zuletzt gespeicherten Zustand zurücksetzen. Die Datei wird dazu geschlossen ohne zu speichern und neu geöffnet.",
            on_action=bkt.Callback(BooksOps.reset_workbook, workbook=True, application=True),
            #get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            get_enabled = bkt.Callback(lambda workbook: workbook.Path != '', workbook=True),
        ),
        bkt.ribbon.SplitButton(
            size='large',
            children=[
                bkt.ribbon.Button(
                    id = 'consolidate_open_workbooks',
                    label="Arbeitsmappen konsolidieren…",
                    show_label=True,
                    image_mso='ReviewShareWorkbook',
                    screentip="Geöffnete Arbeitsmappen konsolidieren",
                    supertip="Konsolidiert die ausgewählten geöffneten Arbeitsmappen in einer neuen Arbeitsmappe, d.h. kopiert alle Blätter in eine Arbeitsmappe. Die aktuelle Arbeitsmappe wird nicht verändert.",
                    on_action=bkt.Callback(BooksOps.consolidate_open_workbooks, workbook=True, sheets=True, application=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.Menu(children=[
                    bkt.ribbon.Button(
                        id = 'consolidate_open_workbooks2',
                        label="Geöffnete Arbeitsmappen konsolidieren…",
                        show_label=True,
                        image_mso='ReviewShareWorkbook',
                        screentip="Geöffnete Arbeitsmappen konsolidieren",
                        supertip="Konsolidiert die ausgewählten geöffneten Arbeitsmappen in einer neuen Arbeitsmappe, d.h. kopiert alle Blätter in eine Arbeitsmappe. Die aktuelle Arbeitsmappe wird nicht verändert.",
                        on_action=bkt.Callback(BooksOps.consolidate_open_workbooks, workbook=True, sheets=True, application=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'consolidate_file_workbooks',
                        label="Dateien zum Konsolidieren auswählen…",
                        show_label=True,
                        #image_mso='ReviewShareWorkbook',
                        screentip="Mehrere Dateien konsolidieren",
                        supertip="Konsolidiert die Blätter der ausgewählten Dateien in einer neuen Arbeitsmappe. Die aktuelle Arbeitsmappe wird nicht verändert.",
                        on_action=bkt.Callback(BooksOps.consolidate_file_workbooks, workbook=True, sheets=True, application=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                ])
            ]
        ),
        bkt.ribbon.Button(
            id = 'consolidate_worksheets',
            label="Blätter in ein Blatt konsolidieren…",
            show_label=True,
            size="large",
            image_mso='ReviewCombineRevisions',
            screentip="Blätter dieser Arbeitsmappe in ein Blatt konsolidieren",
            supertip="Konsolidiert alle ausgewählten Blätter dieser Arbeitsmappe in einem neuen Blatt.",
            on_action=bkt.Callback(BooksOps.consolidate_worksheets, workbook=True, sheet=True, sheets=True, selected_sheets=True, application=True),
            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
        ),
        bkt.ribbon.Button(
            id = 'close_workbooks',
            label="Alle Anderen schließen",
            show_label=True,
            image_mso='CloseAllItems',
            supertip="Alle Arbeitsmappen außer der aktuellen Mappe schließen. Gibt es ungespeicherte Änderungen, kommt vorab eine Meldung.",
            on_action=bkt.Callback(BooksOps.close_workbooks, workbook=True, application=True),
            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
        ),
        bkt.ribbon.Menu(
            label="Blätter kopieren",
            show_label=True,
            image_mso='ExportExcel',
            screentip="Blätter der Arbeitsmappe kopieren",
            supertip="Blätter aus der aktuellen Arbeitsmappe getrennt kopieren und einzeln speichern",
            children=[
                #bkt.ribbon.MenuSeparator(title="Blätter"),
                bkt.ribbon.Button(
                    id = 'copy_selected_sheets',
                    label="Markierte Blätter in neue Arbeitsmappe kopieren",
                    show_label=True,
                    #image_mso='ExportExcel',
                    supertip="Kopiert die markierten Arbeitsblätter in eine neue Arbeitsmappe. Dies funktioniert auch, wenn mehrere Blätter Listen und Tabellen enthalten. Die aktuelle Arbeitsmappe wird nicht verändert.",
                    on_action=bkt.Callback(BooksOps.copy_selected_sheets, workbook=True, application=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.Button(
                    id = 'copy_sheets_and_save',
                    label="Blätter in jeweils eigene Datei speichern…",
                    show_label=True,
                    #image_mso='ExportExcel',
                    supertip="Kopiert alle sichtbaren Blätter jeweils in eine neue Arbeitsmappe und speichert diese im gleichen Verzeichnis wie die aktuelle Arbeitsmappe. Die aktuelle Arbeitsmappe wird nicht verändert.",
                    on_action=bkt.Callback(BooksOps.copy_sheets_and_save, sheets=True, selected_sheets=True, workbook=True, application=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
            ]
        ),
        bkt.ribbon.Menu(
            label="Theme & Color Scheme",
            show_label=True,
            image_mso='SchemeColorsGallery',
            children = [
                bkt.ribbon.Button(
                    id = 'theme_apply',
                    label="Theme aus Datei anwenden",
                    supertip="Beliebige Excel-Datei auswählen und dessen Theme (Farb- und Fontschema) auf aktuelle Datei anwenden.",
                    image_mso='SchemeColorsGallery',
                    on_action=bkt.Callback(BooksOps.theme_apply, workbook=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.MenuSeparator(),
                bkt.ribbon.Menu(
                    label="Color Scheme",
                    show_label=True,
                    # image_mso='SchemeColorsGallery',
                    children = [
                        bkt.ribbon.Button(
                            id = 'theme_color_export',
                            label="Export",
                            screentip="Farbschema Export",
                            supertip="Exportiere Farbschema der Arbeitsmappe als XML-Datei.",
                            on_action=bkt.Callback(BooksOps.theme_color_export, workbook=True),
                            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                        ),
                        bkt.ribbon.Button(
                            id = 'theme_color_import',
                            label="Import",
                            screentip="Farbschema Import",
                            supertip="Importiere Farbschema aus einer XML-Datei in die Arbeitsmappe.",
                            on_action=bkt.Callback(BooksOps.theme_color_import, workbook=True),
                            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                        ),
                    ]
                ),
                bkt.ribbon.Menu(
                    label="Font Scheme",
                    show_label=True,
                    # image_mso='SchemeColorsGallery',
                    children = [
                        bkt.ribbon.Button(
                            id = 'theme_font_export',
                            label="Export",
                            screentip="Fontschema Export",
                            supertip="Exportiere Fontschema der Arbeitsmappe als XML-Datei.",
                            on_action=bkt.Callback(BooksOps.theme_font_export, workbook=True),
                            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                        ),
                        bkt.ribbon.Button(
                            id = 'theme_font_import',
                            label="Import",
                            screentip="Fontschema Import",
                            supertip="Importiere Fontschema aus einer XML-Datei in die Arbeitsmappe.",
                            on_action=bkt.Callback(BooksOps.theme_font_import, workbook=True),
                            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                        ),
                    ]
                ),
            ]
        ),
    ]
)