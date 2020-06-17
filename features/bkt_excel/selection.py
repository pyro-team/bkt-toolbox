# -*- coding: utf-8 -*-
'''
Created on 2017-07-18
@author: Florian Stallmann
'''

from __future__ import absolute_import

import bkt
import bkt.library.excel.helpers as xllib

class SelectionOps(object):
    move_resize = False
    saved_selection = None

    @classmethod
    def toggle_move_resize(cls, pressed):
        cls.move_resize = pressed

    @staticmethod
    def get_selection_height(areas):
        return areas[0].Rows.Count

    @staticmethod
    def set_selection_height(areas, value):
        areas[0].Resize(max(1,value), areas[0].Columns.Count).Select()

    @staticmethod
    def get_selection_width(areas):
        return areas[0].Columns.Count

    @staticmethod
    def set_selection_width(areas, value):
        areas[0].Resize(areas[0].Rows.Count, max(1,value)).Select()

    @classmethod
    def _inc_dec_selection_size(cls, areas, direction):
        bottom_right_cell = areas[0].Cells(areas[0].Rows.Count,areas[0].Columns.Count)
        if bkt.get_key_state(bkt.KeyCodes.CTRL):
            cell = xllib.get_next_cell(bottom_right_cell, direction)
        else:
            cell = xllib.get_next_visible_cell(bottom_right_cell, direction)

        if direction in ['top', 'bottom']:
            cls.set_selection_height(areas, cell.Row - areas[0].Cells(1,1).Row + 1)
        else:
            cls.set_selection_width(areas, cell.Column - areas[0].Cells(1,1).Column + 1)

    @staticmethod
    def get_selection_left(selection, areas):
        return selection.Column -1
        #return areas[0].Cells(1,1).Column -1

    @classmethod
    def set_selection_left(cls, selection, areas, value):
        try:
            delta = value-selection.Column+1
            #delta = value-areas[0].Cells(1,1).Column+1
            if cls.move_resize and len(areas) == 1:
                areas[0].Offset(0, delta).Resize(areas[0].Rows.Count, max(1,areas[0].Columns.Count-delta)).Select()
            else:
                selection.Offset(0, delta).Select()
                #areas[0].Offset(0, delta).Select()
        except:
            pass

    @staticmethod
    def get_selection_top(selection, areas):
        return selection.Row -1
        #return areas[0].Cells(1,1).Row -1

    @classmethod
    def set_selection_top(cls, selection, areas, value):
        try:
            delta = value-selection.Row+1
            #delta = value-areas[0].Cells(1,1).Row+1
            if cls.move_resize and len(areas) == 1:
                areas[0].Offset(delta, 0).Resize(max(1,areas[0].Rows.Count-delta), areas[0].Columns.Count).Select()
            else:
                selection.Offset(delta, 0).Select()
                #areas[0].Offset(delta, 0).Select()
        except:
            pass

    @classmethod
    def _inc_dec_selection_pos(cls, selection, areas, direction):
        if bkt.get_key_state(bkt.KeyCodes.CTRL):
            cell = xllib.get_next_cell(areas[0].Cells(1,1), direction)
        else:
            cell = xllib.get_next_visible_cell(areas[0].Cells(1,1), direction)

        if direction in ['top', 'bottom']:
            cls.set_selection_top(selection, areas, cell.Row-1)
        else:
            cls.set_selection_left(selection, areas, cell.Column-1)

    @staticmethod
    def select_empty_rows(application, sheet):
        cells_selected = None
        area = sheet.UsedRange
        #for area in areas:
        for i in xrange(1,area.Rows.Count+1):
            if application.WorksheetFunction.CountA(area.Rows(i).EntireRow) == 0:
                cells_selected = xllib.range_union(cells_selected, area.Rows(i).EntireRow)

        if not cells_selected:
            bkt.message("Keine leeren Zeilen im genutzten Bereich!")
        else:
            cells_selected.Select()

    @staticmethod
    def select_entire_rows(application, selection):
        selection.EntireRow.Select()

    @staticmethod
    def select_entire_columns(application, selection):
        selection.EntireColumn.Select()

    @staticmethod
    def select_empty_columns(application, sheet):
        cells_selected = None
        area = sheet.UsedRange
        #for area in areas:
        for i in xrange(1,area.Columns.Count+1):
            if application.WorksheetFunction.CountA(area.Columns(i).EntireColumn) == 0:
                cells_selected = xllib.range_union(cells_selected, area.Columns(i).EntireColumn)

        if not cells_selected:
            bkt.message("Keine leeren Spalten im genutzten Bereich!")
        else:
            cells_selected.Select()

    @staticmethod
    def select_used_range(sheet=True):
        sheet.UsedRange.Select()

    @staticmethod
    def select_samecolor(application, sheet, cells):
        cells_selected = None
        cells_colors = set()
        for mastercell in cells:
            cells_colors.add((mastercell.Interior.ThemeColor,mastercell.Interior.Color))
        
        area = sheet.UsedRange
        for cell in iter(area.Cells):
            if (cell.Interior.ThemeColor,cell.Interior.Color) in cells_colors:
                cells_selected = xllib.range_union(cells_selected, cell)
        cells_selected.Select()

    @staticmethod
    def select_unused_areas(sheet, application):
        selection = xllib.get_unused_ranges(sheet)

        if len(selection) == 0:
            bkt.message("Kein ungenutzter Bereich!")
        elif len(selection) == 1:
            selection[0].Select()
        else:
            application.Union(*selection).Select()

        # Alternative method using range subtract:
        # selection = xllib.range_substract(sheet.Cells, sheet.UsedRange, application)
        # if selection:
        #     selection.Select()

    @staticmethod
    def invert_selection(application, sheet, selection):
        rng_input = application.InputBox("Gesamtbereich auswählen:", "Bereich auswählen", sheet.UsedRange.AddressLocal(), type=8) #text, title, default, type=8 (cell range)
        if not rng_input:
            return
        cells_selected = xllib.range_substract(rng_input, selection)
        if not cells_selected:
            bkt.message("Leerer Bereich, keine Zellen zum Markieren!")
        else:
            cells_selected.Select()

    @staticmethod
    def deselect(application, selection):
        rng_input = application.InputBox("Bereich zum Deselektieren auswählen:", "Bereich auswählen", type=8) #text, title, default, type=8 (cell range)
        if not rng_input:
            return
        cells_selected = xllib.range_substract(selection, rng_input)
        if not cells_selected:
            bkt.message("Leerer Bereich, keine Zellen zum Markieren!")
        else:
            cells_selected.Select()

    @staticmethod
    def select_intersection(application, selection):
        rng_input = application.InputBox("Bereich auswählen:", "Bereich auswählen", type=8) #text, title, default, type=8 (cell range)
        if not rng_input:
            return
        cells_selected = application.Intersect(selection, rng_input)
        if not cells_selected:
            bkt.message("Leerer Bereich, keine Zellen zum Markieren!")
        else:
            cells_selected.Select()

    @staticmethod
    def select_union(application, selection):
        rng_input = application.InputBox("Bereich auswählen:", "Bereich auswählen", type=8) #text, title, default, type=8 (cell range)
        if not rng_input:
            return
        cells_selected = application.Union(selection, rng_input)
        if not cells_selected:
            bkt.message("Leerer Bereich, keine Zellen zum Markieren!")
        else:
            cells_selected.Select()

    @staticmethod
    def select_symdiff(application, selection):
        rng_input = application.InputBox("Bereich auswählen:", "Bereich auswählen", type=8) #text, title, default, type=8 (cell range)
        if not rng_input:
            return
        cells_selected = application.Union(selection, rng_input)
        intersection = application.Intersect(selection, rng_input)
        if intersection:
            cells_selected = xllib.range_substract(cells_selected, intersection)
        if not cells_selected:
            bkt.message("Leerer Bereich, keine Zellen zum Markieren!")
        else:
            cells_selected.Select()

    # @staticmethod
    # def select_by_value(cell, selection):
    #     default = cell.Value() if cell.HasFormula else cell.Formula
    #     value_to_select = bkt.ui.show_user_input("Wert eingeben:", "Selektieren nach Wert", default)
    #     if not value_to_select:
    #         return

    @classmethod
    def save_selection(cls, selection):
        cls.saved_selection = selection.AddressLocal(False, False)

    @classmethod
    def restore_selection(cls, application):
        application.Range(cls.saved_selection).Select()

    @classmethod
    def enabled_restore_selection(cls):
        return cls.saved_selection != None

    @staticmethod
    def selection_address(application, selection):
        rng_input = application.InputBox("Bereich zum Selektieren auswählen:", "Selektierte Adresse", default=selection.AddressLocal(False, False), type=8) #text, title, default, type=8 (cell range)
        if not rng_input:
            return
        rng_input.Select()


selektion_gruppe = bkt.ribbon.Group(
    label="Selektion",
    image_mso="SelectCurrentRegion",
    #auto_scale=True,
    children=[
        bkt.ribbon.SplitButton(
            size='large',
            children=[
                bkt.ribbon.Button(
                    id = 'deselect',
                    label="Deselektieren…",
                    show_label=True,
                    image_mso='SelectCell',
                    supertip="Aktuelle Selektion um gewählten Bereich reduzieren (Komplement).",
                    on_action=bkt.Callback(SelectionOps.deselect, application=True, selection=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.Menu(children=[
                    bkt.ribbon.Button(
                        id = 'deselect2',
                        label="Deselektieren…",
                        show_label=True,
                        image_mso='SelectCell',
                        supertip="Aktuelle Selektion um gewählten Bereich reduzieren (Komplement).",
                        on_action=bkt.Callback(SelectionOps.deselect, application=True, selection=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'select_intersection',
                        label="Schnittmenge bilden…",
                        show_label=True,
                        image_mso='ShapesObjectIntersect',
                        supertip="Schnittmenge aus aktueller Selektion und gewähltem Bereich.",
                        on_action=bkt.Callback(SelectionOps.select_intersection, application=True, selection=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'select_union',
                        label="Vereinigung bilden…",
                        show_label=True,
                        image_mso='ShapesObjectUnion',
                        supertip="Vereinigungsmenge aus aktueller Selektion und gewähltem Bereich.",
                        on_action=bkt.Callback(SelectionOps.select_union, application=True, selection=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'select_symdiff',
                        label="Symmetrische Differenz bilden…",
                        show_label=True,
                        image_mso='ShapesObjectCombine',
                        supertip="Symmetrische Differenz (Vereinigungsmenge abzüglich Schnittmenge) aus aktueller Selektion und gewähltem Bereich.",
                        on_action=bkt.Callback(SelectionOps.select_symdiff, application=True, selection=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                ])
            ]
        ),
        bkt.ribbon.Box(box_style="vertical",
            children = [
                bkt.ribbon.SplitButton(
                    children=[
                        bkt.ribbon.Button(
                            id = 'select_used_range',
                            label="Genutzten Bereich markieren",
                            show_label=True,
                            image_mso='SelectCurrentRegion',
                            supertip="Markiert den genutzten Bereich (UsedRange).",
                            on_action=bkt.Callback(SelectionOps.select_used_range, sheet=True, require_worksheet=True),
                            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                        ),
                        bkt.ribbon.Menu(children=[
                            bkt.ribbon.Button(
                                id = 'select_used_range2',
                                label="Genutzten Bereich markieren",
                                show_label=True,
                                image_mso='SelectCurrentRegion',
                                supertip="Markiert den genutzten Bereich (UsedRange).",
                                on_action=bkt.Callback(SelectionOps.select_used_range, sheet=True, require_worksheet=True),
                                get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                            ),
                            bkt.mso.control.SelectCurrentRegion(show_label=True, show_image=False),
                            bkt.ribbon.Button(
                                id = 'select_unused_areas',
                                label="Ungenutzten Bereich markieren",
                                show_label=True,
                                #image_mso='SelectCurrentRegion',
                                supertip="Markiert den ungenutzten Bereich.",
                                on_action=bkt.Callback(SelectionOps.select_unused_areas, sheet=True, application=True, require_worksheet=True),
                                get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                            ),
                        ])
                    ]
                ),
                bkt.mso.control.TableSelectVisibleCells(show_label=True),
                bkt.ribbon.Button(
                    id = 'invert_selection',
                    label="Markierung umkehren…",
                    show_label=True,
                    image_mso='PictureContrastGallery',
                    supertip="Markiert den nicht markierten Bereich innerhalb des gewählten Bereichs.",
                    on_action=bkt.Callback(SelectionOps.invert_selection, selection=True, sheet=True, application=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.SplitButton(
                    children=[
                        bkt.ribbon.Button(
                            id = 'select_empty_rows',
                            label="Leere Zeilen",
                            show_label=True,
                            image_mso='TableRowSelect',
                            supertip="Alle leeren Zeilen innerhalb der benutzen Zellen (UsedRange) markieren.",
                            on_action=bkt.Callback(SelectionOps.select_empty_rows, application=True, sheet=True, require_worksheet=True),
                            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                        ),
                        bkt.ribbon.Menu(children=[
                            bkt.ribbon.Button(
                                id = 'select_empty_rows2',
                                label="Leere Zeilen markieren",
                                show_label=True,
                                image_mso='TableRowSelect',
                                supertip="Alle leeren Zeilen innerhalb der benutzen Zellen (UsedRange) markieren.",
                                on_action=bkt.Callback(SelectionOps.select_empty_rows, application=True, sheet=True, require_worksheet=True),
                                get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                            ),
                            bkt.ribbon.Button(
                                id = 'select_entire_rows',
                                label="Ganze Zeilen markieren",
                                show_label=True,
                                #image_mso='TableRowSelect',
                                supertip="Alle Zeilen der aktuell markierten Zellen markieren.",
                                on_action=bkt.Callback(SelectionOps.select_entire_rows, application=True, selection=True),
                                get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                            ),
                        ])
                    ]
                ),
                bkt.ribbon.SplitButton(
                    children=[
                        bkt.ribbon.Button(
                            id = 'select_empty_columns',
                            label="Leere Spalten",
                            show_label=True,
                            image_mso='TableColumnSelect',
                            supertip="Alle leeren Spalten innerhalb der benutzen Zellen (UsedRange) markieren.",
                            on_action=bkt.Callback(SelectionOps.select_empty_columns, application=True, sheet=True, require_worksheet=True),
                            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                        ),
                        bkt.ribbon.Menu(children=[
                            bkt.ribbon.Button(
                                id = 'select_empty_columns2',
                                label="Leere Spalten markieren",
                                show_label=True,
                                image_mso='TableColumnSelect',
                                supertip="Alle leeren Spalten innerhalb der benutzen Zellen (UsedRange) markieren.",
                                on_action=bkt.Callback(SelectionOps.select_empty_columns, application=True, sheet=True, require_worksheet=True),
                                get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                            ),
                            bkt.ribbon.Button(
                                id = 'select_entire_columns',
                                label="Ganze Spalten markieren",
                                show_label=True,
                                #image_mso='TableColumnSelect',
                                supertip="Alle Spalten der aktuell markierten Zellen markieren.",
                                on_action=bkt.Callback(SelectionOps.select_entire_columns, application=True, selection=True),
                                get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                            ),
                        ])
                    ]
                ),
                bkt.ribbon.Button(
                    id = 'select_samecolor',
                    label="Gleiche Farbe",
                    show_label=True,
                    image_mso='SelectedTaskGoTo',
                    supertip="Alle Zellen mit der gleichen Hintergrundfarbe wie die ausgewählten Zellen markieren.",
                    on_action=bkt.Callback(SelectionOps.select_samecolor, application=True, sheet=True, cells=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                # bkt.ribbon.Button(
                #     id = 'select_by_value',
                #     label="Werte selektieren",
                #     show_label=True,
                #     image_mso='HappyFace',
                #     supertip="XXX",
                #     on_action=bkt.Callback(SelectionOps.select_by_value, cell=True, selection=True),
                #     get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                # ),
            ]
        ),
        bkt.ribbon.Separator(),
        bkt.ribbon.Box(box_style="horizontal",
            children = [
                #TODO: areas_max entfernen und Auswahl für jede Area einzeln verschieben
                bkt.ribbon.SpinnerBox(
                    id = 'selection_top',
                    label=u"Oben",
                    show_label=False,
                    image_mso='ObjectNudgeDown',
                    screentip="Reihen oberhalb der Selektion",
                    on_change = bkt.Callback(SelectionOps.set_selection_top, selection=True, areas=True, areas_min=1),
                    get_text  = bkt.Callback(SelectionOps.get_selection_top, selection=True, areas=True, areas_min=1),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    increment = bkt.Callback(lambda selection, areas: SelectionOps._inc_dec_selection_pos(selection, areas, 'bottom'), selection=True, areas=True, areas_min=1),
                    decrement = bkt.Callback(lambda selection, areas: SelectionOps._inc_dec_selection_pos(selection, areas, 'top'), selection=True, areas=True, areas_min=1),
                    #big_step = 1,
                    #small_step = 1
                ),
                bkt.ribbon.SpinnerBox(
                    id = 'selection_height',
                    label=u"Reihen",
                    show_label=False,
                    image_mso='TableRowsDistribute',
                    screentip="Anzahl der selektierten Reihen",
                    on_change = bkt.Callback(SelectionOps.set_selection_height, areas=True, areas_min=1, areas_max=1),
                    get_text  = bkt.Callback(SelectionOps.get_selection_height, areas=True, areas_min=1, areas_max=1),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    increment = bkt.Callback(lambda areas: SelectionOps._inc_dec_selection_size(areas, 'bottom'), areas=True, areas_min=1, areas_max=1),
                    decrement = bkt.Callback(lambda areas: SelectionOps._inc_dec_selection_size(areas, 'top'), areas=True, areas_min=1, areas_max=1),
                    #big_step = 1,
                    #small_step = 1
                ),
            ]
        ),
        bkt.ribbon.Box(box_style="horizontal",
            children = [
                bkt.ribbon.SpinnerBox(
                    id = 'selection_left',
                    label=u"Links",
                    show_label=False,
                    image_mso='ObjectNudgeRight',
                    screentip="Spalten links der Selektion",
                    on_change = bkt.Callback(SelectionOps.set_selection_left, selection=True, areas=True, areas_min=1),
                    get_text  = bkt.Callback(SelectionOps.get_selection_left, selection=True, areas=True, areas_min=1),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    increment = bkt.Callback(lambda selection, areas: SelectionOps._inc_dec_selection_pos(selection, areas, 'right'), selection=True, areas=True, areas_min=1),
                    decrement = bkt.Callback(lambda selection, areas: SelectionOps._inc_dec_selection_pos(selection, areas, 'left'), selection=True, areas=True, areas_min=1),
                    #big_step = 1,
                    #small_step = 1
                ),
                bkt.ribbon.SpinnerBox(
                    id = 'selection_width',
                    label=u"Spalten",
                    show_label=False,
                    image_mso='TableColumnsDistribute',
                    screentip="Anzahl der selektierten Spalten",
                    on_change = bkt.Callback(SelectionOps.set_selection_width, areas=True, areas_min=1, areas_max=1),
                    get_text  = bkt.Callback(SelectionOps.get_selection_width, areas=True, areas_min=1, areas_max=1),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    increment = bkt.Callback(lambda areas: SelectionOps._inc_dec_selection_size(areas, 'right'), areas=True, areas_min=1, areas_max=1),
                    decrement = bkt.Callback(lambda areas: SelectionOps._inc_dec_selection_size(areas, 'left'), areas=True, areas_min=1, areas_max=1),
                    #big_step = 1,
                    #small_step = 1
                ),
            ]
        ),
        bkt.ribbon.ToggleButton(
            id='selection_move_resize',
            label=u"Verschieben ändert Größe",
            show_label=True,
            image_mso='SizeToControlHeightAndWidth',
            screentip="Verschieben der Selektion verändert auch die Größe",
            on_toggle_action=bkt.Callback(SelectionOps.toggle_move_resize),
            get_pressed=bkt.Callback(lambda: SelectionOps.move_resize),
            get_enabled = bkt.Callback(lambda areas: True, areas=True, areas_min=1, areas_max=1),
        ),
        bkt.ribbon.Separator(),
        bkt.ribbon.Button(
            id = 'save_selection',
            label="Selektion speichern",
            show_label=True,
            image_mso='TableSave',
            supertip="Aktuelle Selektion zwischenspeichern, um diese später wiederherzustellen.",
            on_action=bkt.Callback(SelectionOps.save_selection, selection=True),
            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
        ),
        bkt.ribbon.Button(
            id = 'restore_selection',
            label="Selektion laden",
            show_label=True,
            image_mso='TableSelect',
            supertip="Zwischengespeicherte Selektion auf aktuellem Blatt wiederherstellen.",
            on_action=bkt.Callback(SelectionOps.restore_selection, application=True),
            get_enabled = bkt.Callback(SelectionOps.enabled_restore_selection),
        ),
        bkt.ribbon.Button(
            id = 'selection_address',
            label="Adresse selektieren…",
            show_label=True,
            image_mso='TableSelect',
            supertip="Zeigt Fenster zum Eingeben der Selektionsadresse mit der Adresse der aktuellen Selektion als Standardwert.",
            on_action=bkt.Callback(SelectionOps.selection_address, application=True, selection=True),
            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
        ),

        bkt.ribbon.DialogBoxLauncher(idMso='GoToSpecial')

        #TODOS:
        #Selektion nach Inhalt
        #Selektion nach Farbe/Rahmen
        #Selektion Interval/ jede x-te Zeile/Spalte
    ]
)