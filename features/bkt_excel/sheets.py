# -*- coding: utf-8 -*-
'''
Created on 2017-07-18
@author: Florian Stallmann
'''

import bkt
import bkt.library.excel.helpers as xllib
import bkt.library.excel.constants as xlcon

import os #for filelist
from datetime import datetime #for filelist

from System import DBNull #for list of cond format

class SheetsOps(object):
    very_hidden_sheets = set()
    hidden_sheets = set()

    @staticmethod
    def hide_sheets(selected_sheets, visibility=xlcon.XlSheetVisibility["xlSheetHidden"]):
        try:
            for sheet in selected_sheets:
                sheet.Visible = visibility
        except:
            bkt.helpers.message("Fehler beim Ausblenden. Es muss mind. ein sichtbares Blatt geben.")

    @classmethod
    def hide_sheets_veryhidden(cls, selected_sheets):
        cls.hide_sheets(selected_sheets, xlcon.XlSheetVisibility["xlSheetVeryHidden"])

    @staticmethod
    def show_hidden_sheets(sheets):
        counter = 0
        for sheet in sheets:
            if sheet.Visible == xlcon.XlSheetVisibility["xlSheetHidden"]:
                sheet.Visible = xlcon.XlSheetVisibility["xlSheetVisible"]
                counter += 1
        bkt.helpers.message("Es wurden " + str(counter) + " Blätter eingeblendet.")

    @staticmethod
    def show_veryhidden_sheets(sheets):
        counter = 0
        for sheet in sheets:
            if sheet.Visible == xlcon.XlSheetVisibility["xlSheetVeryHidden"]:
                sheet.Visible = xlcon.XlSheetVisibility["xlSheetVisible"]
                counter += 1
        bkt.helpers.message("Es wurden " + str(counter) + " Blätter eingeblendet.")

    @classmethod
    def toggle_hidden_sheets(cls, sheets, selected_sheets):
        try:
            cls.toggl_sheet_visibility(sheets, cls.hidden_sheets, xlcon.XlSheetVisibility["xlSheetHidden"])
        except:
            cls.hide_sheets(selected_sheets, xlcon.XlSheetVisibility["xlSheetHidden"])

    @classmethod
    def toggle_veryhidden_sheets(cls, sheets, selected_sheets):
        try:
            cls.toggl_sheet_visibility(sheets, cls.very_hidden_sheets, xlcon.XlSheetVisibility["xlSheetVeryHidden"])
        except:
            cls.hide_sheets(selected_sheets, xlcon.XlSheetVisibility["xlSheetVeryHidden"])

    @staticmethod
    def toggl_sheet_visibility(sheets, hidden_sheet_set, visibility=xlcon.XlSheetVisibility["xlSheetHidden"]):
        if len(hidden_sheet_set) == 0:
            for sheet in sheets:
                if sheet.Visible == visibility:
                    sheet.Visible = xlcon.XlSheetVisibility["xlSheetVisible"]
                    hidden_sheet_set.add(sheet.Name)
            if len(hidden_sheet_set) == 0:
                raise AssertionError("No hidden sheets found")
                #bkt.helpers.message("Keine versteckten Blätter gefunden.")
        else:
            for sheet in sheets:
                if sheet.Name in hidden_sheet_set:
                    sheet.Visible = visibility
            hidden_sheet_set.clear()

    @staticmethod
    def show_all_sheets(sheets):
        for sheet in sheets:
            sheet.Visible = -1 #xlSheetVisible

    @classmethod
    def sheets_base_list(cls, workbook, sheets):
        list_sheet = workbook.Worksheets.Add()
        # explanation = list_sheet.Range("A1:C1")
        # explanation.MergeCells = True
        # explanation.WrapText = True
        # explanation.Value = "Umbenennen: XXX\nSortieren: XXX\nErstellen: XXX"

        xllib.rename_sheet(list_sheet, "BKT MULTIEDIT")
        cls._create_list_header(list_sheet, ["#", "Alter Name", "Neuer Name"], row=1)
        cur_row = 2
        for i, sheet in enumerate(sheets, start=1):
            if sheet.Visible != xlcon.XlSheetVisibility["xlSheetVisible"] or sheet.Type != xlcon.XlSheetType["xlWorksheet"]:
                continue
            list_sheet.Cells(cur_row,1).Value = i
            list_sheet.Cells(cur_row,2).Value = sheet.Name
            cur_row += 1
        list_sheet.UsedRange.Columns.AutoFit()

    @staticmethod
    def rename_all_sheets(workbook, areas, application):
        if areas[0].Columns.Count != 2:
            bkt.helpers.message("Es müssen genau 2 Spalten (Alter Name, Neuer Name) ausgewählt werden")
            return
        
        if not xllib.confirm_no_undo(): return
        
        err_counter = 0
        for row in areas[0].Rows:
            old_name = row.Cells(1).Text
            new_name = row.Cells(2).Text
            if not old_name or not new_name:
                continue
            try:
                workbook.Sheets[old_name].Name = new_name[:31] #max sheet name length is 31
            except:
                err_counter += 1
        
        if err_counter > 0:
            bkt.helpers.message("Fehler! " + str(err_counter) + " Blatt/Blätter konnte(n) nicht umbenannt werden.")

    @classmethod
    def sort_all_sheets(cls, workbook, areas, application, sheet):
        if areas[0].Columns.Count != 1 or areas[0].Cells.Count == 1:
            bkt.helpers.message("Es muss genau 1 Spalte (mit Blattnamen in gewünschter Reihenfolge) ausgewählt werden")
            return

        if not xllib.confirm_no_undo(): return

        #Make all sheets visible
        hidden_sheets = set()
        all_sheets = list(iter(workbook.Sheets))
        try:
            cls.toggl_sheet_visibility(all_sheets, hidden_sheets)
        except:
            pass
            #bkt.helpers.exception_as_message()
        
        #Sort sheets
        err_counter = 0
        for row in areas[0].Rows:
            name = row.Cells(1).Text
            if not name:
                continue
            try:
                workbook.Sheets[name].Move(After=workbook.Sheets[workbook.Sheets.Count])
            except:
                err_counter += 1
                #bkt.helpers.exception_as_message(name)
        
        #Restore sheet visibility
        if len(hidden_sheets) > 0:
            try:
                cls.toggl_sheet_visibility(all_sheets, hidden_sheets)
            except:
                pass
                #bkt.helpers.exception_as_message()

        sheet.Activate()

        if err_counter > 0:
            bkt.helpers.message("Fehler! " + str(err_counter) + " Blatt/Blätter konnte(n) nicht umsortiert werden.")
    
    @staticmethod
    def create_sheets(workbook, areas, application):
        if areas[0].Columns.Count != 1 or areas[0].Cells.Count == 1:
            bkt.helpers.message("Es muss genau 1 Spalte (mit anzulegenden Blattnamen) ausgewählt werden")
            return

        if not xllib.confirm_no_undo(): return
        
        err_counter = 0
        for row in areas[0].Rows:
            name = row.Cells(1).Text
            if not name:
                continue
            try:
                new_sheet = workbook.Worksheets.Add()
                new_sheet.Name = name[:31] #max sheet name length is 31
            except:
                err_counter += 1
        
        if err_counter > 0:
            bkt.helpers.message("Fehler! " + str(err_counter) + " Blatt/Blätter konnte(n) nicht angelegt werden.")

    @staticmethod
    def _create_list_header(list_sheet, header, row=1):
        for i,h in enumerate(header, start=1):
            list_sheet.Cells(row,i).Value = h
        list_sheet.Range(list_sheet.Cells(row,1),list_sheet.Cells(row,len(header))).Font.Bold = True

    @classmethod
    def list_properties(cls, workbook):
        doctypes = {
            1: "Zahl",
            2: "Ja/Nein",
            3: "Datum",
            4: "Text",
            5: "Zahl",
        }

        list_sheet = workbook.Worksheets.Add()
        xllib.rename_sheet(list_sheet, "BKT LISTE DOKU. EIG.")
        cls._create_list_header(list_sheet, ["Typ", "Name", "Wert", "Datentyp"])
        cur_row = 2
        for prop in iter(workbook.BuiltinDocumentProperties):
            list_sheet.Cells(cur_row,1).Value = "Standard"
            try:
                list_sheet.Cells(cur_row,3).Value = prop.Value()
                list_sheet.Cells(cur_row,2).Value = prop.Name()
                list_sheet.Cells(cur_row,4).Value = doctypes[prop.Type()]
                cur_row += 1
            except:
                pass

        for prop in iter(workbook.CustomDocumentProperties):
            list_sheet.Cells(cur_row,1).Value = "Benutzerdefiniert"
            try:
                list_sheet.Cells(cur_row,3).Value = prop.Value()
                list_sheet.Cells(cur_row,2).Value = prop.Name()
                list_sheet.Cells(cur_row,4).Value = doctypes[prop.Type()]
                cur_row += 1
            except:
                pass

        list_sheet.UsedRange.Columns.AutoFit()

    @classmethod
    def list_names(cls, workbook, sheets):
        list_sheet = workbook.Worksheets.Add()
        #list_sheet.Name = "BKT LISTE NAMEN"
        xllib.rename_sheet(list_sheet, "BKT LISTE NAMEN")
        cls._create_list_header(list_sheet, ["Typ", "Name", "Bezug", "Bereich"])
        #list_sheet.Range("A2").ListNames()
        cur_row = 2
        for name in iter(workbook.Names):
            if not name.Visible:
                continue
            ident = name.NameLocal.split("!",1)
            list_sheet.Cells(cur_row,1).Value = "Name"
            list_sheet.Cells(cur_row,2).Value = "'" + ident[-1] #last element
            list_sheet.Cells(cur_row,3).Value = "'" + name.RefersToLocal
            list_sheet.Cells(cur_row,4).Value = "Arbeitsmappe" if len(ident) == 1 else ident[0].strip("'")
            cur_row += 1

        for sheet in sheets:
            for obj in iter(sheet.ListObjects):
                list_sheet.Cells(cur_row,1).Value = "Tabelle"
                list_sheet.Cells(cur_row,2).Value = "'" + obj.Name #FIXME: use DisplayName instead???
                list_sheet.Cells(cur_row,3).Value = "'=" + xllib.get_address_external(obj.Range, True, True)
                list_sheet.Cells(cur_row,4).Value = "Arbeitsmappe"
                cur_row += 1

        list_sheet.UsedRange.Columns.AutoFit()

    @classmethod
    def list_comments(cls, sheet, workbook):
        list_sheet = workbook.Worksheets.Add()
        xllib.rename_sheet(list_sheet, "BKT LISTE KOMMENTARE")
        cls._create_list_header(list_sheet, ["Zelle", "Autor", "Text"])
        cur_row = 2

        list_sheet.Range("C:C").ColumnWidth = 30
        for comment in iter(sheet.Comments):
            list_sheet.Cells(cur_row,1).Value = comment.Parent.AddressLocal(False, False)
            list_sheet.Cells(cur_row,2).Value = comment.Author
            list_sheet.Cells(cur_row,3).Value = comment.Text()
            cur_row += 1

        list_sheet.UsedRange.Columns.AutoFit()


    @classmethod
    def list_cond_formats(cls, sheet, workbook):
        def _dict_by_value(input_dict, search_value):
            for key, value in input_dict.iteritems():
                if value == search_value:
                    return key

        def _getattr(obj, name, default=None):
            try:
                return getattr(obj, name, default)
            except:
                return default

        def _copy_values(from_obj, to_obj, attribute_list):
            for attr in attribute_list:
                val = _getattr(from_obj, attr, None)
                if val is not DBNull and val is not None:
                    setattr(to_obj, attr, val)

        list_sheet = workbook.Worksheets.Add()
        xllib.rename_sheet(list_sheet, "BKT LISTE BEN. FORMAT.")
        cls._create_list_header(list_sheet, ["Priorität", "Typ", "Formel 1", "Formel 2", "Text", "Operator", "Format", "Bereich", "Anhalten"])
        cur_row = 2

        # IMPORTANT LINE! For some reason excel crashs when accessing border/font color if sheet is not active
        sheet.Activate()

        for fcond in iter(sheet.Cells.FormatConditions):
            list_sheet.Cells(cur_row,1).Value = fcond.Priority
            list_sheet.Cells(cur_row,2).Value = _dict_by_value(xlcon.XlFormatConditionType, fcond.Type)
            
            list_sheet.Cells(cur_row,3).Value = "'" + _getattr(fcond, "Formula1", '')
            list_sheet.Cells(cur_row,4).Value = "'" + _getattr(fcond, "Formula2", '')
            list_sheet.Cells(cur_row,5).Value = "'" + _getattr(fcond, "Text", '')

            operator = _getattr(fcond, "Operator", None)
            list_sheet.Cells(cur_row,6).Value = None if operator is None else _dict_by_value(xlcon.XlFormatConditionOperator, operator)
            
            #Format
            list_sheet.Cells(cur_row,7).Value = "AaBbCcYyZz"
            _copy_values(fcond.Interior, list_sheet.Cells(cur_row,7).Interior, ["Color", "Pattern", "PatternColor"])
            _copy_values(fcond.Borders, list_sheet.Cells(cur_row,7).Borders, ["Color", "LineStyle", "Weight"])
            _copy_values(fcond.Font, list_sheet.Cells(cur_row,7).Font, ["Color", "FontStyle"])
            
            list_sheet.Cells(cur_row,8).Value = "'=" + xllib.get_address_external(fcond.AppliesTo, True, True)
            list_sheet.Cells(cur_row,9).Value = "X" if fcond.StopIfTrue else None
            cur_row += 1
        
        list_sheet.Activate()
        list_sheet.UsedRange.Columns.AutoFit()

    @classmethod
    def list_sheets(cls, workbook, sheets):
        list_sheet = workbook.Worksheets.Add()
        #list_sheet.Name = "BKT LISTE BLÄTTER"
        xllib.rename_sheet(list_sheet, "BKT LISTE BLÄTTER")
        cls._create_list_header(list_sheet, ["Name", "Genutzter Bereich", "Zeilen", "Spalten", "Tab-Farbe", "Sichtbar", "Geschützt"])
        cur_row = 2
        for sheet in sheets:
            if sheet.Visible == xlcon.XlSheetVisibility["xlSheetVeryHidden"]:
                continue
            if sheet.Type == xlcon.XlSheetType["xlWorksheet"]:
                list_sheet.Hyperlinks.Add(list_sheet.Cells(cur_row,1), "", "'" + sheet.Name + "'!A1", "", sheet.Name) #anchor, address, subaddress, screentip, texttodisplay
                list_sheet.Cells(cur_row,2).Value = "'=" + xllib.get_address_external(sheet.UsedRange, True, True)
                list_sheet.Cells(cur_row,3).Value = sheet.UsedRange.Rows.Count
                list_sheet.Cells(cur_row,4).Value = sheet.UsedRange.Columns.Count
                if sheet.Tab.Color:
                    list_sheet.Cells(cur_row,5).Interior.Color = sheet.Tab.Color
                list_sheet.Cells(cur_row,6).Value = "X" if sheet.Visible == xlcon.XlSheetVisibility["xlSheetVisible"] else None
                list_sheet.Cells(cur_row,7).Value = "X" if sheet.ProtectContents else None
            else:
                list_sheet.Cells(cur_row,1).Value = sheet.Name
            cur_row += 1
        list_sheet.UsedRange.Columns.AutoFit()

    @classmethod
    def list_workbooks(cls, workbook, application):
        list_sheet = workbook.Worksheets.Add()
        #list_sheet.Name = "BKT LISTE ARBEITSMAPPEN"
        xllib.rename_sheet(list_sheet, "BKT LISTE ARBEITSMAPPEN")
        cls._create_list_header(list_sheet, ["Name", "Ordner", "Pfad", "Anz. Blätter", "Liste Blätter"])
        cur_row = 2
        for wb in list(iter(application.Workbooks)):
            list_sheet.Cells(cur_row,1).Value = wb.Name
            list_sheet.Cells(cur_row,2).Value = wb.Path
            if wb.FullName:
                # list_sheet.Cells(cur_row,3).Value = wb.FullName
                list_sheet.Hyperlinks.Add(list_sheet.Cells(cur_row,3), wb.FullName, "", "", wb.FullName) #anchor, address, subaddress, screentip, texttodisplay
            list_sheet.Cells(cur_row,4).Value = wb.Worksheets.Count
            list_sheet.Cells(cur_row,5).Value = ", ".join([sh.Name for sh in wb.Worksheets if sh.Visible != xlcon.XlSheetVisibility["xlSheetVeryHidden"]])
            cur_row += 1
        list_sheet.UsedRange.Columns.AutoFit()

    @classmethod
    def file_list(cls, application, workbook, recursive=False):
        fileDialog = application.FileDialog(4) #msoFileDialogFolderPicker
        if workbook.Path:
            fileDialog.InitialFileName = workbook.Path + '\\'
        fileDialog.title = "Ordner für Dateiliste auswählen"

        #TODO: Filter auf Dateityp einbauen

        if fileDialog.Show() == 0: #msoFalse
            return
        folder = fileDialog.SelectedItems(1)
        if not os.path.isdir(folder):
            return

        xllib.freeze_app()
        application.StatusBar = "Erstelle Dateiliste"

        sheet = workbook.Worksheets.Add()
        #sheet.Name = "BKT DATEILISTE"
        xllib.rename_sheet(sheet, "BKT DATEILISTE")
        cls._create_list_header(sheet, ["Name", "Typ", "Größe", "Erstellt", "Geändert", "Ordner", "Pfad"], 2)

        total = cls.create_list_of_folder(application, folder, sheet, 3, recursive)
        total -= 3

        application.StatusBar = False
        xllib.unfreeze_app()

        sheet.UsedRange.Columns.AutoFit()
        sheet.Cells(1,1).Value = "Dateiliste mit " + str(total) + " Dateien für Ordner: " + os.path.normpath(folder)


    @classmethod
    def create_list_of_folder(cls, application, folder, sheet, cur_row, recursive=False):
        if not os.path.isdir(folder):
            return 0
        subfolders = []

        application.StatusBar = "Erstelle Dateiliste für Ordner " + folder
        # bkt.helpers.message("Liste für Ordner: " + folder)

        for file in os.listdir(folder):
            #full_path = folder + "\\" + file
            full_path = os.path.join(folder, file)

            if os.path.isdir(full_path) == True:
                subfolders.append(full_path)

            if os.path.isfile(full_path) == True:
                try:
                    full_path = os.path.normpath(full_path)
                    root,ext = os.path.splitext(file)
                    sheet.Cells(cur_row,1).Value = root
                    sheet.Cells(cur_row,2).Value = ext
                    sheet.Cells(cur_row,3).Value = str(os.path.getsize(full_path))
                    sheet.Cells(cur_row,4).Value = datetime.fromtimestamp(os.path.getctime(full_path)).strftime('%Y-%m-%d %H:%M:%S')
                    sheet.Cells(cur_row,5).Value = datetime.fromtimestamp(os.path.getmtime(full_path)).strftime('%Y-%m-%d %H:%M:%S')
                    sheet.Cells(cur_row,6).Value = os.path.normpath(folder)
                    #sheet.Cells(cur_row,7).Value = full_path
                    sheet.Hyperlinks.Add(sheet.Cells(cur_row,7), full_path, "", "", full_path) #anchor, address, subaddress, screentip, texttodisplay
                except:
                    #Fallback: Simple info
                    sheet.Cells(cur_row,1).Value = file
                    sheet.Cells(cur_row,6).Value = folder
                    sheet.Cells(cur_row,7).Value = full_path
                cur_row += 1

        # bkt.helpers.message("Unterordner: " + str(len(subfolders)))

        if recursive:
            for subs in subfolders:
                cur_row = cls.create_list_of_folder(application, subs, sheet, cur_row, recursive)

        return cur_row


blatt_gruppe = bkt.ribbon.Group(
    label="Blätter",
    image_mso="SheetInsert",
    auto_scale=True,
    children=[
        bkt.ribbon.SplitButton(
            size="large",
            children=[
                bkt.ribbon.Button(
                    id = 'toggle_hidden_sheets',
                    label="Blätter ein/ausblenden",
                    show_label=True,
                    image_mso='SheetInsert',
                    #image="toggle_hidden_sheets",
                    supertip="Alle ausgeblendeten Blätter zwischen ein- und ausblenden umschalten.\n\nSind keine Blätter ausgeblendet, werden die ausgewählten Blätter bzw. das aktuelle Blatt ausgeblendet.",
                    on_action=bkt.Callback(SheetsOps.toggle_hidden_sheets, sheets=True, selected_sheets=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.Menu(children=[
                    bkt.ribbon.MenuSeparator(title="Ein-/Ausblenden"),
                    bkt.ribbon.Button(
                        id = 'toggle_hidden_sheets2',
                        label="Blätter ein/ausblenden",
                        show_label=True,
                        image_mso='SheetInsert',
                        #image="toggle_hidden_sheets",
                        supertip="Alle ausgeblendeten Blätter zwischen ein- und ausblenden umschalten.\n\nSind keine Blätter ausgeblendet, werden die ausgewählten Blätter bzw. das aktuelle Blatt ausgeblendet.",
                        on_action=bkt.Callback(SheetsOps.toggle_hidden_sheets, sheets=True, selected_sheets=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'toggle_veryhidden_sheets',
                        label="Blätter anzeigen/verstecken (xlVeryHidden)",
                        show_label=True,
                        #image_mso='CreateQueryFromWizard',
                        supertip="Alle versteckten (xlVeryHidden) Blätter zwischen anzeigen und verstecken umschalten. \n\nind keine Blätter versteckt, werden die ausgewählten Blätter bzw. das aktuelle Blatt versteckt.",
                        on_action=bkt.Callback(SheetsOps.toggle_veryhidden_sheets, sheets=True, selected_sheets=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.MenuSeparator(title="Ausblenden"),
                        bkt.ribbon.Button(
                            id = 'hide_sheets',
                            label="Blatt ausblenden",
                            show_label=True,
                            #image_mso='CreateQueryFromWizard',
                            supertip="Aktuelles Blatt bzw. ausgewählte Blätter ausblenden.",
                            on_action=bkt.Callback(SheetsOps.hide_sheets, selected_sheets=True),
                            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                        ),
                        bkt.ribbon.Button(
                            id = 'hide_sheets_veryhidden',
                            label="Blatt verstecken (xlVeryHidden)",
                            show_label=True,
                            #image_mso='CreateQueryFromWizard',
                            supertip="Aktuelles Blatt bzw. ausgewählte Blätter verstecken (xlVeryHidden), sodass diese nur über die Toolbox oder ein Makro wieder sichtbar gemacht werden können.",
                            on_action=bkt.Callback(SheetsOps.hide_sheets_veryhidden, selected_sheets=True),
                            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                        ),
                    bkt.ribbon.MenuSeparator(title="Einblenden"),
                    bkt.ribbon.Button(
                        id = 'show_hidden_sheets',
                        label="Alle ausgeblendeten Blätter einblenden",
                        show_label=True,
                        #image_mso='QueryMakeTable',
                        supertip="Alle ausgeblendeten Blätter wieder einblenden.",
                        on_action=bkt.Callback(SheetsOps.show_hidden_sheets, sheets=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'show_veryhidden_sheets',
                        label="Alle versteckten Blätter (nur xlVeryHidden) einblenden",
                        show_label=True,
                        #image_mso='QueryMakeTable',
                        supertip="Alle ausgeblendeten oder versteckten (xlVeryHidden) Blätter wieder einblenden.",
                        on_action=bkt.Callback(SheetsOps.show_veryhidden_sheets, sheets=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'show_all_sheets',
                        label="Alle Blätter (inkl. xlVeryHidden) einblenden",
                        show_label=True,
                        #image_mso='QueryMakeTable',
                        supertip="Alle ausgeblendeten oder versteckten (xlVeryHidden) Blätter wieder einblenden.",
                        on_action=bkt.Callback(SheetsOps.show_all_sheets, sheets=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                ])
            ]
        ),
        bkt.ribbon.Menu(
            label="Listen erstellen",
            show_label=True,
            size="large",
            image_mso='TableOfContentsDialog',
            #image="list_sheets",
            screentip="Verschiedene Listen erstellen",
            supertip="Liste aller Blätter, Arbeitsmappen, Dateliste, ...",
            children=[
                bkt.ribbon.MenuSeparator(title="Listen zum Blatt"),
                bkt.ribbon.Button(
                    id = 'list_comments',
                    label="Liste aller Kommentare",
                    show_label=True,
                    #image_mso='SheetInsert',
                    supertip="Erstellt Liste aller Kommentare des aktuellen Blatts in neuem Blatt.",
                    on_action=bkt.Callback(SheetsOps.list_comments, sheet=True, workbook=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.Button(
                    id = 'list_cond_formats',
                    label="Liste aller bedingten Formatierungen",
                    show_label=True,
                    #image_mso='SheetInsert',
                    supertip="Erstellt Liste aller bedingten Formatierungen des aktuellen Blatts in neuem Blatt.",
                    on_action=bkt.Callback(SheetsOps.list_cond_formats, sheet=True, workbook=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.MenuSeparator(title="Listen zur Mappe"),
                bkt.ribbon.Button(
                    id = 'list_names',
                    label="Liste aller Namen",
                    show_label=True,
                    #image_mso='SheetInsert',
                    supertip="Erstellt Liste aller Namen dieser Arbeitsmappe in neuem Blatt.",
                    on_action=bkt.Callback(SheetsOps.list_names, workbook=True, sheets=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.Button(
                    id = 'list_sheets',
                    label="Liste aller Arbeitsblätter",
                    show_label=True,
                    #image_mso='SheetInsert',
                    supertip="Erstellt Liste aller Blätter dieser Arbeitsmappe in einem neuen Blatt.",
                    on_action=bkt.Callback(SheetsOps.list_sheets, workbook=True, sheets=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.Button(
                    id = 'list_properties',
                    label="Liste aller Dokumenteneigenschaften",
                    show_label=True,
                    #image_mso='SheetInsert',
                    supertip="Erstellt Liste aller Dokumenteneigenschaften dieser Arbeitsmappe in neuem Blatt.",
                    on_action=bkt.Callback(SheetsOps.list_properties, workbook=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.Button(
                    id = 'list_workbooks',
                    label="Liste aller geöffneten Arbeitsmappen",
                    show_label=True,
                    #image_mso='SheetInsert',
                    supertip="Erstellt Liste aller geöffneten Arbeitsmappen inkl. Pfad in einem neuen Blatt.",
                    on_action=bkt.Callback(SheetsOps.list_workbooks, workbook=True, application=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.MenuSeparator(title="Dateilisten"),
                bkt.ribbon.Button(
                    id = 'file_list',
                    label="Dateiliste erstellen…",
                    show_label=True,
                    image_mso='FileVersionHistory',
                    supertip="Wähle Ordner und erstelle Liste aller Dateien in diesem Ordner in neuem Blatt.",
                    on_action=bkt.Callback(lambda application, workbook: SheetsOps.file_list(application, workbook, False), application=True, workbook=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.Button(
                    id = 'file_list_recursive',
                    label="Dateiliste erstellen (rekursiv)…",
                    show_label=True,
                    #image_mso='FileVersionHistory',
                    supertip="Wähle Ordner und erstelle Liste aller Dateien in diesem Ordner und allen Unterordnern in neuem Blatt.",
                    on_action=bkt.Callback(lambda application, workbook: SheetsOps.file_list(application, workbook, True), application=True, workbook=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
            ]
        ),
        bkt.ribbon.Menu(
            label="Umbenennen/ Sortieren",
            show_label=True,
            size="large",
            image_mso='Rename',
            screentip="Mehrere Blätter umbennen, sortieren oder neu erstellen",
            supertip="Mehrere Arbeitsblätter gemäß einer Auswahl umbenennen oder erstellen aus Vorlage",
            children=[
                bkt.ribbon.Button(
                    id = 'sheets_base_list',
                    label="Basisliste erstellen",
                    show_label=True,
                    #image_mso='SheetInsert',
                    supertip="Erstellt Liste aller Blätter dieser Arbeitsmappe in einem neuen Blatt.",
                    on_action=bkt.Callback(SheetsOps.sheets_base_list, workbook=True, sheets=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.MenuSeparator(),
                bkt.ribbon.Button(
                    id = 'rename_all_sheets',
                    label="Blätter gemäß aktueller Auswahl umbenennen",
                    show_label=True,
                    #image_mso='Rename',
                    supertip="Alle Blätter gemäß der Auswahl umbenennen. Die Auswahl muss aus genau 2 Spalten bestehen, wobei die erste Spalten den alten Namen und die zweite Spalte den neuen Namen enthält. Leere Namen werden übersprungen.",
                    on_action=bkt.Callback(SheetsOps.rename_all_sheets, workbook=True, areas=True, areas_max=1, application=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.Button(
                    id = 'sort_all_sheets',
                    label="Blätter gemäß aktueller Auswahl sortieren",
                    show_label=True,
                    #image_mso='Rename',
                    supertip="Alle Blätter gemäß der Auswahl sortieren. Die Auswahl muss aus genau einer Spalten mit den Blattnamen in der gewünschten Reihenfolge bestehen.",
                    on_action=bkt.Callback(SheetsOps.sort_all_sheets, sheet=True, workbook=True, areas=True, areas_max=1, application=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.Button(
                    id = 'create_sheets',
                    label="Blätter gemäß aktueller Auswahl erstellen",
                    show_label=True,
                    #image_mso='Rename',
                    supertip="Neue Blätter gemäß der aktuellen Auswahl erstellen. Die Auswahl muss aus genauer einer Spalte mit den anzulegenden Blattnamen bestehen.",
                    on_action=bkt.Callback(SheetsOps.create_sheets, workbook=True, areas=True, areas_max=1, application=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
            ]
        )
    ]
)