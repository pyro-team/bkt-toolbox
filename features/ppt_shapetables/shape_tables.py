# -*- coding: utf-8 -*-
'''
Created on 2017-08-16
@author: Florian Stallmann
'''

from __future__ import absolute_import

#for caching
import time

import bkt
# import bkt.library.powerpoint as pplib

from bkt.library.powerpoint import PositionGallery, pt_to_cm, LocPin, LocpinGallery
from bkt.library.table import TableRecognition

# pt_to_cm_factor = 2.54 / 72;
# def pt_to_cm(pt):
#     return float(pt) * pt_to_cm_factor;
# def cm_to_pt(cm):
#     return float(cm) / pt_to_cm_factor;


class ShapeTables(object):
    resize_cells = False
    fit_cells = False
    alignment_locpin = LocPin()
    # alignment_horizontal = "left"
    # alignment_vertical = "top"
    equal_spacing = False

    # contentarea = {"left": None, "top": None, "width": None, "height": None}
    tr_cache = None
    last_time_tr_cache_changed = 0

    def __init__(self):
        self.position_gallery = PositionGallery(
            label="Tabelle in Bereich einpassen",
            description="Shape-Tabelle auf Größe des gewählten Bereichs anpassen",
            on_position_change = bkt.Callback(self.table_contentarea_fit),
            get_item_supertip = bkt.Callback(self.get_item_supertip)
        )
        self.locpin_gallery = LocpinGallery(
            id="table_alignment",
            label="Shape-Ausrichtung",
            image_mso="ObjectAlignMenu",
            supertip="Legt Ausrichtung der Shapes innerhalb der Tabellenzellen fest.",
            item_height="32",
            item_width="32",
            locpin=self.alignment_locpin,
            item_supertip="Shapes werden bei Shape-Anordnung in Tabellenzellen {} angeordnet.",
        )
    
    @property
    def alignment_horizontal(self):
        return ["left", "center", "right"][self.alignment_locpin.fixation[1]-1]
    @property
    def alignment_vertical(self):
        return ["top", "middle", "bottom"][self.alignment_locpin.fixation[0]-1]

    def get_item_supertip(self, index):
        return 'Die ausgewählten Shapes werden als Tabelle in den Bereich eingepasst.'

    def _prepare_table(self, shapes):
        # Run table recognition only one time in 500ms
        if self.tr_cache is None or time.time() - self.last_time_tr_cache_changed > 0.5:
            self.tr_cache = TableRecognition(shapes)
            self.tr_cache.run()
            self.last_time_tr_cache_changed = time.time()
        return self.tr_cache


    ### ALIGN TABLE / do not consider resize-option ###

    def align_table(self, shapes):
        if bkt.get_key_state(bkt.KeyCodes.SHIFT):
            self.align_table_zero(shapes)
        else:
            tr = self._prepare_table(shapes)
            spac_rows = max(0,tr.min_spacing_rows(max_rows=2))
            spac_cols = max(0,tr.min_spacing_cols(max_cols=2))
            if self.equal_spacing:
                spacing = (spac_rows+spac_cols)/2.0
            else:
                spacing = (spac_rows, spac_cols)
            tr.align(spacing=spacing, fit_cells=self.fit_cells, align_x=self.alignment_horizontal, align_y=self.alignment_vertical)

    def align_table_default(self, shapes):
        tr = self._prepare_table(shapes)
        tr.align(fit_cells=self.fit_cells, align_x=self.alignment_horizontal, align_y=self.alignment_vertical)

    def align_table_median(self, shapes):
        tr = self._prepare_table(shapes)
        tr.align(tr.median_spacing(), fit_cells=self.fit_cells, align_x=self.alignment_horizontal, align_y=self.alignment_vertical)

    def align_table_zero(self, shapes):
        tr = self._prepare_table(shapes)
        tr.align(0, fit_cells=self.fit_cells, align_x=self.alignment_horizontal, align_y=self.alignment_vertical)


    ### SPACING / consider resize option ###

    # def get_spacing(self, shapes):
    #     tr = self._prepare_table(shapes)
    #     res = tr.median_spacing()
    #     return res
    #     # return round(pt_to_cm(res), 2)

    # def set_spacing(self, shapes, value):
    #     # if type(value) == str:
    #     #     value = float(value.replace(',', '.'))
    #     # value = max(0,cm_to_pt(value))
    #     value = max(0, value)

    #     tr = self._prepare_table(shapes)
    #     if self.resize_cells:
    #         bounds = tr.get_bounds()
    #         tr.fit_content(*bounds, spacing=value, fit_cells=self.fit_cells)
    #     else:
    #         tr.align(spacing=value, fit_cells=self.fit_cells, align_x=self.alignment_horizontal, align_y=self.alignment_vertical)


    def get_resize_cells(self):
        if bkt.get_key_state(bkt.KeyCodes.ALT):
            return not self.resize_cells
        else:
            return self.resize_cells

    def enabled_spacing_rows(self, shapes):
        if len(shapes) < 2:
            return False
        tr = self._prepare_table(shapes)
        return tr.dimension[0] > 1

    def get_spacing_rows(self, shapes):
        tr = self._prepare_table(shapes)
        res = tr.min_spacing_rows(max_rows=2)
        return res
        # return round(pt_to_cm(res), 2)

    def set_spacing_rows(self, shapes, value):
        # if type(value) == str:
        #     value = float(value.replace(',', '.'))
        # value = max(0,cm_to_pt(value))
        # value = max(0, value)

        if self.equal_spacing:
            spacing = value
        else:
            spacing = (value, None)

        tr = self._prepare_table(shapes)
        if self.get_resize_cells():
            bounds = tr.get_bounds()
            tr.fit_content(*bounds, spacing=spacing, fit_cells=self.fit_cells)
        else:
            tr.align(spacing=spacing, fit_cells=self.fit_cells, align_x=self.alignment_horizontal, align_y=self.alignment_vertical)

    def enabled_spacing_cols(self, shapes):
        if len(shapes) < 2:
            return False
        tr = self._prepare_table(shapes)
        return tr.dimension[1] > 1

    def get_spacing_cols(self, shapes):
        tr = self._prepare_table(shapes)
        res = tr.min_spacing_cols(max_cols=2)
        return res
        # return round(pt_to_cm(res), 2)

    def set_spacing_cols(self, shapes, value):
        # if type(value) == str:
        #     value = float(value.replace(',', '.'))
        # value = max(0,cm_to_pt(value))
        # value = max(0, value)

        if self.equal_spacing:
            spacing = value
        else:
            spacing = (None, value)

        tr = self._prepare_table(shapes)
        if self.get_resize_cells():
            bounds = tr.get_bounds()
            tr.fit_content(*bounds, spacing=spacing, fit_cells=self.fit_cells)
        else:
            tr.align(spacing=spacing, fit_cells=self.fit_cells, align_x=self.alignment_horizontal, align_y=self.alignment_vertical)


    ### SPECIAL FUNCTIONS / separate method for resizing ###

    def table_transpose_destructive(self,shapes):
        tr = self._prepare_table(shapes)
        spacing = tr.median_spacing()
        tr.transpose()
        tr.align(spacing=spacing, fit_cells=self.fit_cells, align_x=self.alignment_horizontal, align_y=self.alignment_vertical)

    def table_transpose_in_bounds(self,shapes):
        tr = self._prepare_table(shapes)
        spacing = tr.median_spacing()
        bounds = tr.get_bounds()
        tr.transpose()
        tr.transpose_cell_size()
        tr.fit_content(*bounds, spacing=spacing, fit_cells=self.fit_cells)

    def table_info(self,shapes):
        tr = self._prepare_table(shapes)
        msg = u""
        msg += "Tabellengröße: Zeilen=%d, Spalten=%d\r\n" % tr.dimension
        msg += "Median-Abstand: %s cm" % round(pt_to_cm(tr.median_spacing()),2)
        bkt.message(msg)

    def table_info_desc(self,context):
        try:
            if context.selection.Type != 2 or context.selection.ShapeRange.Count < 2:
                raise ValueError("invalid selection for table")

            # shapes = pplib.get_shapes_from_selection(selection)
            tr = self._prepare_table(context.shapes)
            return u"Tabelle: %d Z. \xd7 %d S." % tr.dimension #Zeilen x Spalten
        except:
            return "Tabelle: -"

    ### FIT CELLS ###

    def table_fit_cells_destructive(self,shapes):
        tr = self._prepare_table(shapes)
        spacing = tr.median_spacing()
        tr.align(spacing=spacing, fit_cells=True, align_x=self.alignment_horizontal, align_y=self.alignment_vertical)

    def table_fit_cells_in_bounds(self,shapes):
        tr = self._prepare_table(shapes)
        spacing = tr.median_spacing()
        bounds = tr.get_bounds()
        tr.fit_content(*bounds, spacing=spacing, fit_cells=True)

    def table_contentarea_fit(self,target_frame,shapes):
        if len(shapes) < 2:
            return
        tr = self._prepare_table(shapes)
        spacing = tr.median_spacing()
        tr.fit_content(target_frame.left, target_frame.top, target_frame.width, target_frame.height, spacing)

    ### DISTRIBUTE CELLS ###

    def table_distribute_cols(self, shapes):
        tr = self._prepare_table(shapes)
        spacing = tr.min_spacing_cols()
        bounds = tr.get_bounds()
        tr.fit_content(*bounds, spacing=(None, spacing), fit_cells=True) #equalize spacing in first run
        tr.fit_content(*bounds, spacing=(None, spacing), fit_cells=True, distribute_cols=True) #distribute in second run

    def table_distribute_rows(self, shapes):
        tr = self._prepare_table(shapes)
        spacing = tr.min_spacing_rows()
        bounds = tr.get_bounds()
        tr.fit_content(*bounds, spacing=(spacing, None), fit_cells=True) #equalize spacing in first run
        tr.fit_content(*bounds, spacing=(spacing, None), fit_cells=True, distribute_rows=True) #distribute in second run

    @staticmethod
    def show_dialog(context, shapes):
        from .table_dialog import ShapesAsTableWindow
        ShapesAsTableWindow.create_and_show_dialog(context, shapes)


shape_tables = ShapeTables()


tabellen_gruppe = bkt.ribbon.Group(
    id="bkt_shapetables_group",
    label='Tabelle aus Shapes',
    supertip="Ermöglicht die tabellenförmige Anordnung von Shapes. Das Feature `ppt_shapetables` muss installiert sein.",
    image='align_table',
    children = [
        bkt.ribbon.SplitButton(
            get_enabled = bkt.apps.ppt_shapes_min2_selected,
            size="large",
            children=[
                bkt.ribbon.Button(
                    id = 'align_table',
                    label="Als Tabelle ausrichten",
                    show_label=True,
                    # size="large",
                    image='align_table',
                    screentip="Tabelle ausrichten (Auto)",
                    supertip="Richtet die ausgewählten Shapes als Tabelle aus mit errechnetem Zeilen- und Spaltenabstand. Mit SHIFT-Taste wird Abstand=0 gesetzt.",
                    on_action=bkt.Callback(shape_tables.align_table, shapes=True, shapes_min=2),
                    # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.Menu(label="Menü zum Ausrichten als Tabelle", supertip="Shape-Tabelle mit verschiedenen Abstandsoptionen ausrichten", item_size="large", children=[
                    bkt.ribbon.Button(
                        id = 'align_table2',
                        label="Als Tabelle ausrichten (Standardabstand)",
                        description="Shapes als Tabelle mit Abstand 0,35cm ausrichten",
                        # show_label=True,
                        image='align_table',
                        supertip="Richtet die ausgewählten Shapes als Tabelle aus mit einem Standardabstand von 0,35cm (10pt)",
                        on_action=bkt.Callback(shape_tables.align_table_default, shapes=True, shapes_min=2),
                        # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'align_table_median',
                        label="Als Tabelle ausrichten (Median-Abstand)",
                        description="Shapes als Tabelle mit Median-Abstand ausrichten",
                        # show_label=True,
                        image='align_table_median',
                        supertip="Richtet die ausgewählten Shapes als Tabelle aus mit dem Median-Abstand der Shapes.",
                        on_action=bkt.Callback(shape_tables.align_table_median, shapes=True, shapes_min=2),
                        # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'align_table_zero',
                        label="Als Tabelle ausrichten ohne Abstand [Shift]",
                        description="Shapes als Tabelle ohne Abstand ausrichten.",
                        # show_label=True,
                        image='align_table_zero',
                        supertip="Richtet die ausgewählten Shapes als Tabelle ohne Abstand aus.",
                        on_action=bkt.Callback(shape_tables.align_table_zero, shapes=True, shapes_min=2),
                        # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'align_table_dialog',
                        label="Tabellen-Dialog",
                        # description="Shapes als Tabelle ohne Abstand ausrichten.",
                        # show_label=True,
                        # image='align_table_zero',
                        # supertip="Richtet die ausgewählten Shapes als Tabelle ohne Abstand aus.",
                        on_action=bkt.Callback(ShapeTables.show_dialog, context=True, shapes=True, shapes_min=2),
                        # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    # bkt.ribbon.MenuSeparator(),
                    # bkt.ribbon.Button(
                    #     id = 'table_fit_cells',
                    #     label="Shapes in Zellen einpassen",
                    #     description="Setzt die Größe der Shapes auf die Größe der Tabellenzellen",
                    #     # show_label=True,
                    #     image='table_fit_cells',
                    #     supertip="Setzt die Shape-Größe auf die Größe der Tabellenzelle. Je nach gewähltem Modus werden die Shapes dabei verschoben oder vergrößert bzw. verkleinert.",
                    #     on_action=bkt.Callback(shape_tables.table_fit_cells, shapes=True, shapes_min=2),
                    #     # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    # ),
                ])
            ]
        ),
        shape_tables.locpin_gallery,
        bkt.ribbon.Menu(
            image='shape_table_transpose',
            label="Tabelle anpassen",
            supertip="Shape-Tabelle auf verschiedene Weise anpassen oder transponieren",
            item_size="large",
            get_enabled = bkt.apps.ppt_shapes_min2_selected,
            children=[
                bkt.ribbon.Menu(
                    image='table_fit_cells',
                    label="Shapes in Zellen einpassen",
                    description="Setzt die Größe der Shapes auf die Größe der Tabellenzellen.",
                    item_size="large",
                    children=[
                        bkt.ribbon.Button(
                            id = 'table_fit_cells_in_bounds',
                            label="Einpassen mit gleicher Tabellengröße",
                            description="Setzt die Größe der Shapes auf die Größe der Tabellenzellen und behält dabei die aktuelle Tabellengröße",
                            # show_label=True,
                            image='table_fit_cells_1',
                            supertip="Setzt die Shape-Größe auf die Größe der Tabellenzelle. Dabei werden die Shapes vergrößert bzw. verkleinert, um die aktuelle Tabellengröße nicht zu verändern.",
                            on_action=bkt.Callback(shape_tables.table_fit_cells_in_bounds, shapes=True, shapes_min=2),
                            # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                        ),
                        bkt.ribbon.Button(
                            id = 'table_fit_cells_destructive',
                            label="Einpassen mit angepasster Tabellengröße",
                            description="Setzt die Größe der Shapes auf die Größe der Tabellenzellen und verändert dabei die aktuelle Tabellengröße",
                            # show_label=True,
                            image='table_fit_cells_2',
                            supertip="Setzt die Shape-Größe auf die Größe der Tabellenzelle. Dabei werden die Shapes verschoben und die aktuelle Tabellegröße damit verändert.",
                            on_action=bkt.Callback(shape_tables.table_fit_cells_destructive, shapes=True, shapes_min=2),
                            # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                        ),
                    ]
                ),
                bkt.ribbon.Menu(
                    image_mso='TableColumnsDistribute',
                    label="Zellengrößen angleichen",
                    description="Normalisiert die Breite bzw. Höhe der Zellen.",
                    item_size="large",
                    children=[
                        bkt.ribbon.Button(
                            id = 'table_distribute_cols',
                            label="Spaltenbreite verteilen",
                            description="Die Breite der Spalten gleichmäßig verteilen",
                            # show_label=True,
                            image_mso='TableColumnsDistribute',
                            supertip="Normalisiert die Breite aller Spalten ohne dabei die Tabellengröße zu verändern.",
                            on_action=bkt.Callback(shape_tables.table_distribute_cols, shapes=True, shapes_min=2),
                            # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                        ),
                        bkt.ribbon.Button(
                            id = 'table_distribute_rows',
                            label="Zeilenhöhe verteilen",
                            description="Die Höhe der Zeilen gleichmäßig verteilen",
                            # show_label=True,
                            image_mso='TableRowsDistribute',
                            supertip="Normalisiert die Höhe aller Zeilen ohne dabei die Tabellengröße zu verändern.",
                            on_action=bkt.Callback(shape_tables.table_distribute_rows, shapes=True, shapes_min=2),
                            # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                        ),
                    ]
                ),
                bkt.ribbon.Menu(
                    image='shape_table_transpose',
                    label="Tabelle transponieren",
                    description="Transponiert (d.h. spiegelt) die Tabelle.",
                    item_size="large",
                    children=[
                        bkt.ribbon.Button(
                            id = 'table_transpose_in_bounds',
                            label="Tabelle mit gleicher Tabellengröße transponieren",
                            description="Shape-Tabelle transponieren und dabei Tabellengröße nicht verändern",
                            # show_label=True,
                            image='shape_table_transpose_1',
                            supertip="Transponiert die Tabelle, d.h. spiegelt die Zellen an der Hauptdiagonalen. Dabei wir auch die Größe der Zellen verändert, um die Tabellengröße nicht zu ändern.",
                            on_action=bkt.Callback(shape_tables.table_transpose_in_bounds, shapes=True, shapes_min=2),
                            # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                        ),
                        bkt.ribbon.Button(
                            id = 'table_transpose_destructive',
                            label="Tabelle mit angepasster Tabellengröße transponieren",
                            description="Shape-Tabelle transponieren und dabei Zellengrößen nicht verändern",
                            # show_label=True,
                            image='shape_table_transpose_2',
                            supertip="Transponiert die Tabelle, d.h. spiegelt die Zellen an der Hauptdiagonalen. Dabei wird die Größe der Zellen nicht verändert, sondern nur die Tabellengröße.",
                            on_action=bkt.Callback(shape_tables.table_transpose_destructive, shapes=True, shapes_min=2),
                            # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                        ),
                    ]
                ),
                shape_tables.position_gallery,
            ]
        ),
        bkt.ribbon.Button(
            id = 'table_info',
            # label="Tabelleninfo. zeigen",
            get_label=bkt.Callback(shape_tables.table_info_desc, context=True),
            # get_description=bkt.Callback(shape_tables.table_info_desc, shapes=True, shapes_min=2),
            # description="Informationen über erkannte Tabelle anzeigen",
            # show_label=True,
            image='shape_table_info',
            screentip="Tabelleninformation",
            supertip="Zeigt Informationen über die Tabelle an. Nützlich um vorab herauszufinden, ob eine Tabelle korrekt erkannt wird.",
            on_action=bkt.Callback(shape_tables.table_info, shapes=True, shapes_min=2),
            get_enabled = bkt.apps.ppt_shapes_min2_selected,
        ),
        # bkt.ribbon.Menu(
        #     label=u"Weitere Tabellenfeatures",
        #     image="shape_table_info",
        #     item_size="large",
        #     children = [
        #         # bkt.ribbon.Button(
        #         #     id = 'table_info',
        #         #     label="Tabelleninfo. zeigen",
        #         #     get_description=bkt.Callback(shape_tables.table_info_desc, shapes=True, shapes_min=2),
        #         #     # description="Informationen über erkannte Tabelle anzeigen",
        #         #     # show_label=True,
        #         #     image='shape_table_info',
        #         #     supertip="Zeigt Informationen über die Tabelle an. Nützlich um vorab herauszufinden, ob eine Tabelle korrekt erkannt wird.",
        #         #     on_action=bkt.Callback(shape_tables.table_info, shapes=True, shapes_min=2),
        #         #     get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
        #         # ),
        #         bkt.ribbon.Button(
        #             id = 'table_transpose',
        #             label="Tabelle transponieren",
        #             description="Shape-Tabelle transonieren (d.h. spiegeln)",
        #             # show_label=True,
        #             image='shape_table_transpose',
        #             supertip="Transponiert die Tabelle, d.h. spiegelt die Zellen an der Hauptdiagonalen.",
        #             on_action=bkt.Callback(shape_tables.table_transpose, shapes=True, shapes_min=2),
        #             get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
        #         ),
        #         bkt.ribbon.MenuSeparator(),
        #         shape_tables.position_gallery,
        #         # bkt.ribbon.Button(
        #         #     id = 'table_contentarea_set',
        #         #     label="Inhaltsbereich definieren",
        #         #     show_label=True,
        #         #     image='table_contentarea_set',
        #         #     supertip="Definiert die Position und Größe des ausgewählten Shapes als Inhaltsbereich, um Tabellen in diesem Bereich einzupassen. Das ausgewählte Shape wird dann gelöscht.",
        #         #     on_action=bkt.Callback(shape_tables.table_contentarea_set, presentation=True, shape=True),
        #         #     get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
        #         # ),
        #         # bkt.ribbon.Button(
        #         #     id = 'table_contentarea_unset',
        #         #     label="Inhaltsbereich zurücksetzen",
        #         #     show_label=True,
        #         #     image='table_contentarea_unset',
        #         #     supertip="Löscht den definierten Inhaltsbereich, d.h. es wird wieder der Standard Inhaltsbereich der Folie verwendet.",
        #         #     on_action=bkt.Callback(shape_tables.table_contentarea_unset, presentation=True),
        #         #     get_enabled = bkt.Callback(shape_tables.table_contentarea_defined, presentation=True)
        #         # ),
        #         # bkt.ribbon.Button(
        #         #     id = 'table_contentarea_fit',
        #         #     label="Tabelle in Inhaltsbereich einpassen",
        #         #     show_label=True,
        #         #     image='table_contentarea_fit',
        #         #     supertip="Passt die ausgewählten Shapes als Tabelle in die Dimensionen des Inhaltsbereichs ein. Der Inhaltsbereich kann vorher definiert werden, anderenfalls wird der Standard Inhaltsbereich der Folie verwendet.",
        #         #     on_action=bkt.Callback(shape_tables.table_contentarea_fit, presentation=True, shapes=True, shapes_min=2),
        #         #     #get_enabled = bkt.Callback(shape_tables.table_contentarea_enabled, presentation=True, shapes=True)
        #         #     get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
        #         # ),
        #     ]
        # ),
        bkt.ribbon.Separator(),
        # bkt.ribbon.RoundingSpinnerBox(
        #     id = 'align_table_spacing',
        #     label=u"Abstand",
        #     show_label=False,
        #     image="align_table_spacing",
        #     supertip="Ändert den Abstand der Shapes. Je nach gewähltem Modus werden die Shapes dabei verschoben oder vergrößert bzw. verkleinert.",
        #     on_change = bkt.Callback(shape_tables.set_spacing, shapes=True, shapes_min=2),
        #     get_text  = bkt.Callback(shape_tables.get_spacing, shapes=True, shapes_min=2),
        #     get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
        #     round_cm = True,
        #     convert = 'pt_to_cm'
        # ),
        bkt.ribbon.RoundingSpinnerBox(
            id = 'align_table_spacing_rows',
            label=u"Zeilenabstand",
            show_label=False,
            image_mso="VerticalSpacingIncrease",
            supertip="Ändert den Zeilenabstand der Shapes. [ALT] wechselt zwischen Bewegen und Dehen/Stauchen.",
            on_change = bkt.Callback(shape_tables.set_spacing_rows, shapes=True, shapes_min=2),
            get_text  = bkt.Callback(shape_tables.get_spacing_rows, shapes=True, shapes_min=2),
            get_enabled = bkt.Callback(shape_tables.enabled_spacing_rows, shapes=True),
            round_cm = True,
            convert = 'pt_to_cm'
        ),
        bkt.ribbon.RoundingSpinnerBox(
            id = 'align_table_spacing_cols',
            label=u"Spaltenabstand",
            show_label=False,
            image_mso="HorizontalSpacingIncrease",
            supertip="Ändert den Spaltenabstand der Shapes. [ALT] wechselt zwischen Bewegen und Dehen/Stauchen.",
            on_change = bkt.Callback(shape_tables.set_spacing_cols, shapes=True, shapes_min=2),
            get_text  = bkt.Callback(shape_tables.get_spacing_cols, shapes=True, shapes_min=2),
            get_enabled = bkt.Callback(shape_tables.enabled_spacing_cols, shapes=True),
            round_cm = True,
            convert = 'pt_to_cm'
        ),
        bkt.ribbon.Menu(
            label="Optionen",
            supertip="Einstellungsmöglichkeiten beim Ausrichten der Shape-Tabellen",
            item_size="large",
            children=[
                bkt.ribbon.MenuSeparator(title="Abstand einstellen durch:"),
                bkt.ribbon.ToggleButton(
                    id="toggle_table_move",
                    label="Bewegen",
                    description="Gewünschter Shape-Abstand wird durch Positionierung von Shapes erreicht",
                    image_mso="ObjectNudgeRight",
                    on_toggle_action=bkt.Callback(lambda pressed: setattr(shape_tables, 'resize_cells', False)),
                    get_pressed=bkt.Callback(lambda : not shape_tables.resize_cells)
                ),
                bkt.ribbon.ToggleButton(
                    id="toggle_table_resize",
                    label="Dehnen/Stauchen",
                    description="Gewünschter Shape-Abstand wird durch Verkleinerung/Vergrößerung von Shapes erreicht",
                    image_mso="ShapeWidth",
                    on_toggle_action=bkt.Callback(lambda pressed: setattr(shape_tables, 'resize_cells', True)),
                    get_pressed=bkt.Callback(lambda : shape_tables.resize_cells)
                ),
                bkt.ribbon.MenuSeparator(title="Abstand gleichstellen:"),
                bkt.ribbon.ToggleButton(
                    id="toggle_table_equal_spacing",
                    label="Zeilenabstand = Spaltenabstand",
                    description="Bei Veränderung des Zeilenabstands wird auch der Spaltenabstand geändert und umgekehrt",
                    image="align_table_spacing",
                    on_toggle_action=bkt.Callback(lambda pressed: setattr(shape_tables, 'equal_spacing', pressed)),
                    get_pressed=bkt.Callback(lambda : shape_tables.equal_spacing)
                ),
            ]
        ),
        # bkt.ribbon.ToggleButton(id="toggle_table_resize", label="Größe anp.", show_label=True, supertip="Gewünschte Shape-Anordnung wird durch Verkleinerung/Vergrößerung (anstatt Positionierung) von Shapes erreicht", image_mso="ShapeWidth",   on_toggle_action=bkt.Callback(lambda pressed: setattr(shape_tables, 'resize_cells', pressed)),  get_pressed=bkt.Callback(lambda : shape_tables.resize_cells))
        # bkt.ribbon.Box(box_style="horizontal", children=[
        #     bkt.ribbon.Label(label="Modus: "),
        #     bkt.ribbon.ToggleButton(id="toggle_table_move",   label="Bewegen",         show_label=False, supertip="Gewünschte Shape-Anordnung wird durch Positionierung von Shapes erreicht", image_mso="ObjectNudgeRight",         on_toggle_action=bkt.Callback(lambda pressed: setattr(shape_tables, 'resize_cells', False)), get_pressed=bkt.Callback(lambda : not shape_tables.resize_cells)),
        #     bkt.ribbon.ToggleButton(id="toggle_table_resize", label="Dehnen/Stauchen", show_label=False, supertip="Gewünschte Shape-Anordnung wird durch Verkleinerung/Vergrößerung von Shapes erreicht", image_mso="ShapeWidth",   on_toggle_action=bkt.Callback(lambda pressed: setattr(shape_tables, 'resize_cells', True)),  get_pressed=bkt.Callback(lambda : shape_tables.resize_cells))
        # ]),
    ]
)

bkt.powerpoint.add_tab(bkt.ribbon.Tab(
    id="bkt_powerpoint_toolbox_extensions",
    #id_q="nsBKT:powerpoint_toolbox_extensions",
    #insert_after_q="nsBKT:powerpoint_toolbox_advanced",
    insert_before_mso="TabHome",
    label=u'Toolbox 3/3',
    # get_visible defaults to False during async-startup
    get_visible=bkt.Callback(lambda:True),
    children = [
        tabellen_gruppe,
    ]
), extend=True)