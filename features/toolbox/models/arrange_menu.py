# -*- coding: utf-8 -*-
'''
Created on 19.01.2023

'''

import bkt

import bkt.library.algorithms as algos
import bkt.library.powerpoint as pplib


from ..linkshapes import LinkedShapes


class RepositionGallery(pplib.PositionGallery):
    
    def __init__(self, **kwargs):
        super(RepositionGallery, self).__init__(
            on_position_change = bkt.Callback(self.on_position_change),
            **kwargs
        )
    
    def get_item_supertip(self, index):
        return 'Positioniere die ausgewählten Shapes an der angezeigten Position/Größe.\nNur Position ändern [STRG],\nNur Größe ändern [SHIFT]'
    
    def on_position_change(self, target_frame, selection, shapes):
        if len(shapes) > 1:
            shape = selection.ShapeRange.Group()
            self.change_shape_position(shape, target_frame)
            shape.Ungroup().Select()
        else:
            self.change_shape_position(shapes[0], target_frame)
    
    def change_shape_position(self, shape, target_frame):
        # position shape
        # CTRL = position only
        # SHIFT = size only
        
        if not bkt.get_key_state(bkt.KeyCodes.SHIFT):
            shape.left   = target_frame.left
            shape.top    = target_frame.top
        
        if not bkt.get_key_state(bkt.KeyCodes.CTRL):
            shape.width  = target_frame.width
            shape.height = target_frame.height

reposition_gallery = RepositionGallery(id="positions", label="Shapes positionieren")


class ChartShapes(object):
    chart_dimensions = [None, None] #height, width
    plotarea_dimensions = [None, None, None, None] #top, left, height, width

    @classmethod
    def is_chart_shape(cls, shape):
        try:
            # HasChart throws NotImplementedError for SmartArts
            return shape.HasChart == -1
        except NotImplementedError:
            return False
        # return shape.Type == pplib.MsoShapeType['msoChart'] or shape.Type == pplib.MsoShapeType['msoDiagram']

    @classmethod
    def is_paste_enabled(cls, shapes):
        return cls.chart_dimensions[0] is not None and all(cls.is_chart_shape(shape) for shape in shapes)

    @classmethod
    def copy_dimensions(cls, shape):
        plotarea = shape.Chart.PlotArea
        cls.chart_dimensions = [shape.Height, shape.Width]
        cls.plotarea_dimensions = [plotarea.Top, plotarea.Left, plotarea.Height, plotarea.Width]

    @classmethod
    def paste_dimensions(cls, shapes):
        for shape in shapes:
            plotarea = shape.Chart.PlotArea
            shape.Height, shape.Width = cls.chart_dimensions
            plotarea.Top, plotarea.Left, plotarea.Height, plotarea.Width = cls.plotarea_dimensions


class GroupsMore(object):
    @staticmethod
    def add_into_group(shapes):
        if shapes[0].Type == pplib.MsoShapeType['msoGroup']:
            master = pplib.GroupManager(shapes.pop(0))
            master.add_child_items(shapes).select()
        elif shapes[-1].Type == pplib.MsoShapeType['msoGroup']:
            master = pplib.GroupManager(shapes.pop(-1))
            master.add_child_items(shapes).select()
        else:
            pplib.shapes_to_range(shapes).group().select()

    @staticmethod
    def remove_from_group(shapes):
        master = pplib.GroupManager(shapes[0].ParentGroup)
        master.remove_child_items(shapes)
        pplib.shapes_to_range(shapes).select()
    
    @classmethod
    def visible_add_into_group(cls, shapes):
        return len(shapes) > 1 and cls.contains_group(shapes)

    @staticmethod
    def visible_remove_from_group(shapes):
        return all(pplib.shape_is_group_child(shape) for shape in shapes)
    
    @staticmethod
    def contains_group(shapes):
        return any(shp.Type == pplib.MsoShapeType['msoGroup'] for shp in shapes)

    @staticmethod
    def is_or_within_group(shape):
        return shape.Type == pplib.MsoShapeType['msoGroup'] or pplib.shape_is_group_child(shape)

    @staticmethod
    def recursive_ungroup(shapes):
        for shape in shapes:
            if shape.Type == pplib.MsoShapeType['msoGroup']:
                grp = pplib.GroupManager(shape)
                grp.recursive_ungroup().select(False)

    @staticmethod
    def select_all_groupitems(shape):
        if shape.Type == pplib.MsoShapeType['msoGroup']:
            all_shapes = list(iter(shape.GroupItems))
        elif pplib.shape_is_group_child(shape):
            all_shapes = list(iter(shape.ParentGroup.GroupItems))
        else:
            return
        pplib.shapes_to_range(all_shapes).select()


class ArrangeCenter(object):
    @staticmethod
    def shape_in_center(center_shape, around_shapes):
        midpoint = algos.mid_point_shapes(around_shapes)
        center_shape.left = midpoint[0] - center_shape.width/2.0
        center_shape.top = midpoint[1] - center_shape.height/2.0
    
    @classmethod
    def arrange_shapes(cls, shapes):
        cls.shape_in_center(shapes.pop(-1), shapes)


class PictureFormat(object):
    shape_dimensions = [None, None] #ShapeHeight, ShapeWidth #, ShapeTop, ShapeLeft
    pic_dimensions   = [None, None, None, None] #PictureHeight, PictureWidth, PictureOffsetX, PictureOffsetY

    @classmethod
    def is_pic_shape(cls, shape):
        try:
            return shape.Type == pplib.MsoShapeType["msoPicture"]
        except:
            return False

    @classmethod
    def is_paste_enabled(cls, shapes):
        return cls.shape_dimensions[0] is not None and all(cls.is_pic_shape(shape) for shape in shapes)

    @classmethod
    def copy_dimensions(cls, shape):
        croparea = shape.PictureFormat.crop
        cls.shape_dimensions = [croparea.ShapeHeight, croparea.ShapeWidth] #, croparea.ShapeTop, croparea.ShapeLeft
        cls.pic_dimensions   = [croparea.PictureHeight, croparea.PictureWidth, croparea.PictureOffsetX, croparea.PictureOffsetY]

    @classmethod
    def paste_dimensions(cls, shapes):
        for shape in shapes:
            croparea = shape.PictureFormat.crop
            croparea.ShapeHeight, croparea.ShapeWidth = cls.shape_dimensions
            croparea.PictureHeight, croparea.PictureWidth, croparea.PictureOffsetX, croparea.PictureOffsetY = cls.pic_dimensions


class TableFormat(object):
    col_widths = []
    row_heights = []

    @classmethod
    def is_table_shape(cls, shape):
        try:
            # return shape.Type == pplib.MsoShapeType["msoTable"]
            return shape.HasTable == -1 #also covers tables in placeholders
        except:
            return False

    @classmethod
    def is_paste_enabled(cls, shapes):
        return len(cls.col_widths) > 0 and all(cls.is_table_shape(shape) for shape in shapes)

    @classmethod
    def copy_dimensions(cls, shape):
        cls.col_widths = [col.width for col in shape.table.columns]
        cls.row_heights = [row.height for row in shape.table.rows]

    @classmethod
    def paste_dimensions(cls, shapes):
        for shape in shapes:
            for i, col_width in enumerate(cls.col_widths):
                if shape.table.columns.count-1 < i:
                    break
                shape.table.columns(i+1).width = col_width
            
            for i, row_height in enumerate(cls.row_heights):
                if shape.table.rows.count-1 < i:
                    break
                shape.table.rows(i+1).height = row_height


class EdgeAutoFixer(object):
    threshold  = bkt.settings.get("toolbox.autofixer_threshold", 0.3)
    groupitems = bkt.settings.get("toolbox.autofixer_groupitems", True)
    order_key  = bkt.settings.get("toolbox.autofixer_order_key", "diagonal-down")

    @classmethod
    def settings_setter(cls, name, value):
        bkt.settings["toolbox.autofixer_"+name] = value
        setattr(cls, name, value)

    @classmethod
    def _iterate_all_shapes(cls, shapes, groupitems=True):
        for shape in shapes:
            #shapes that are rotated other than 0, 90, 180 or 270 degree are excluded
            if shape.rotation % 90 != 0:
                continue
            #connected connectors should not be moved
            if shape.Connector and (shape.ConnectorFormat.BeginConnected or shape.ConnectorFormat.EndConnected):
                continue
            
            if groupitems and shape.Type == 6: #pplib.MsoShapeType['msoGroup']
                for gShape in shape.GroupItems:
                    yield gShape
            else:
                yield shape
    
    @classmethod
    def get_image(cls, context):
        if cls.order_key == "diagonal-down":
            return context.python_addin.load_image("autofixer_dd")
        elif cls.order_key == "top-down":
            return context.python_addin.load_image("autofixer_td")
        else:
            return context.python_addin.load_image("autofixer_lr")

    @classmethod
    def autofix_edges_diagonal_down(cls, shapes):
        cls.settings_setter("order_key", "diagonal-down")
        cls.autofix_edges(shapes)
    
    @classmethod
    def autofix_edges_left_right(cls, shapes):
        cls.settings_setter("order_key", "left-right")
        cls.autofix_edges(shapes)
    
    @classmethod
    def autofix_edges_top_down(cls, shapes):
        cls.settings_setter("order_key", "top-down")
        cls.autofix_edges(shapes)

    @classmethod
    def autofix_edges(cls, shapes):
        cls._autofix_edges(shapes, pplib.cm_to_pt(cls.threshold), cls.groupitems, cls.order_key)
    
    @classmethod
    def _autofix_edges(cls, shapes, threshold=None, groupitems=True, order_key="diagonal-down"):
        #TODO: how to handle locked aspect-ratio and autosize? rotated shapes? ojects with 0 height/width? exclude placeholders?

        threshold = threshold or pplib.cm_to_pt(0.3)

        shapes = pplib.wrap_shapes(cls._iterate_all_shapes(shapes, groupitems))

        # shapes.sort(key=lambda shape: (shape.left, shape.top))
        order_keys = {
            "diagonal-down": [lambda shape: shape.visual_x+shape.visual_y, False],
            "diagonal-up":   [lambda shape: shape.visual_x+shape.visual_y, True],
            "left-right": [lambda shape: (shape.visual_x,shape.visual_y), False],
            "top-down":   [lambda shape: (shape.visual_y,shape.visual_x), False],
            "right-left": [lambda shape: (shape.visual_x,shape.visual_y), True],
            "bottom-up":  [lambda shape: (shape.visual_y,shape.visual_x), True],
        }
        shapes.sort(key=order_keys[order_key][0], reverse=order_keys[order_key][1])

        # logging.debug("Autofix: top-left")
        child_shapes = shapes[:]
        for master_shape in shapes:
            child_shapes.remove(master_shape)
            
            for shape in child_shapes:
                # logging.debug("Autofix1: %s x %s", master_shape.name, shape.name)

                #save values before moving shape
                # visual_x1, visual_y1 = shape.visual_x1, shape.visual_y1

                if 1e-4 < abs(shape.visual_x-master_shape.visual_x) < threshold:
                    #resize to left edge
                    delta = master_shape.visual_x - shape.visual_x
                    shape.visual_x += delta
                    shape.visual_width -= delta

                if 1e-4 < abs(shape.visual_y-master_shape.visual_y) < threshold:
                    #resize to top edge
                    delta = master_shape.visual_y - shape.visual_y
                    shape.visual_y += delta
                    shape.visual_height -= delta

                if 1e-4 < abs(shape.visual_x1-master_shape.visual_x1) < threshold:
                    #resize to right edge
                    shape.visual_width = master_shape.visual_x1-shape.visual_x

                if 1e-4 < abs(shape.visual_y1-master_shape.visual_y1) < threshold:
                    #resize to bottom edge
                    shape.visual_height = master_shape.visual_y1-shape.visual_y


arrange_menu = lambda: bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None, 
                children=[
                # bkt.ribbon.MenuSeparator(title="Rotation"),
                # bkt.mso.button.ObjectRotateRight90,
                # bkt.mso.button.ObjectRotateLeft90,
                # bkt.mso.button.ObjectFlipHorizontal,
                # bkt.mso.button.ObjectFlipVertical,
                # bkt.ribbon.MenuSeparator(),
                # bkt.mso.button.ObjectRotationOptionsDialog,
                bkt.ribbon.MenuSeparator(title="Positionieren"),
                bkt.mso.control.ObjectRotateGallery,
                reposition_gallery,
                bkt.ribbon.Menu(
                    label='Dimensionen/Größen übertragen',
                    supertip="Dimensionen von Diagrammen, Bild-Ausschnitten und Tabellen von einem Objekt auf ein anderes kopieren",
                    image_mso='PasteWithColumnWidths',
                    children=[
                        bkt.ribbon.Button(
                            id = 'chart_dimensions_copy',
                            label="Diagramm-Dimensionen kopieren",
                            image_mso="ChartPlotArea",
                            screentip="Größe und Position vom Diagrammbereich kopieren",
                            supertip="Kopiert Höhe und Breite des Diagramms sowie Größe und Position der Zeichnungsfläche, um ein anderes Diagramm anzugleichen.",
                            on_action=bkt.Callback(ChartShapes.copy_dimensions, shape=True),
                            get_enabled = bkt.Callback(ChartShapes.is_chart_shape, shape=True),
                        ),
                        bkt.ribbon.Button(
                            id = 'chart_dimensions_paste',
                            label="Diagramm-Dimensionen einfügen",
                            image_mso="PasteWithColumnWidths",
                            screentip="Größe und Position vom Diagrammbereich einfügen",
                            supertip="Überträgt die kopierte Größe und Position des Diagramms bzw. der Zeichnungsfläche auf das ausgewählte Diagramm.",
                            on_action=bkt.Callback(ChartShapes.paste_dimensions, shapes=True),
                            get_enabled = bkt.Callback(ChartShapes.is_paste_enabled, shapes=True),
                        ),
                        bkt.ribbon.MenuSeparator(),
                        bkt.ribbon.Button(
                            id = 'pic_crop_copy',
                            label="Bild-Zuschnitt kopieren",
                            image_mso="PictureCrop",
                            screentip="Größe und Position des Bildausschnitts kopieren",
                            supertip="Kopiert Höhe und Breite des Ausschnitts bei einem zugeschnittenen Bild, um den Ausschnitt mit einem anderen Bild anzugleichen.",
                            on_action=bkt.Callback(PictureFormat.copy_dimensions, shape=True),
                            get_enabled = bkt.Callback(PictureFormat.is_pic_shape, shape=True),
                        ),
                        bkt.ribbon.Button(
                            id = 'pic_crop_paste',
                            label="Bild-Zuschnitt einfügen",
                            image_mso="PasteWithColumnWidths",
                            screentip="Größe und Position des Bildausschnitts einfügen",
                            supertip="Überträgt die kopierte Größe und Position des Bilde-Ausschnitts auf das ausgewählte Bild.",
                            on_action=bkt.Callback(PictureFormat.paste_dimensions, shapes=True),
                            get_enabled = bkt.Callback(PictureFormat.is_paste_enabled, shapes=True),
                        ),
                        bkt.ribbon.MenuSeparator(),
                        bkt.ribbon.Button(
                            id = 'table_dimensions_copy',
                            label="Tabellengrößen kopieren",
                            image_mso="TableColumnsDistribute",
                            screentip="Breite/Höhe der Tabellenspalten/-zeilen kopieren",
                            supertip="Kopiert Höhe und Breite der Tabellenzeilen bzw. Tabellenspalten, um diese mit einer anderen Tabelle anzugleichen.",
                            on_action=bkt.Callback(TableFormat.copy_dimensions, shape=True),
                            get_enabled = bkt.Callback(TableFormat.is_table_shape, shape=True),
                        ),
                        bkt.ribbon.Button(
                            id = 'table_dimensions_paste',
                            label="Tabellengrößen einfügen",
                            image_mso="PasteWithColumnWidths",
                            screentip="Breite/Höhe der Tabellenspalten/-zeilen einfügen",
                            supertip="Überträgt die kopierten Tabellen-Dimensionen auf die ausgewählte Tabelle.",
                            on_action=bkt.Callback(TableFormat.paste_dimensions, shapes=True),
                            get_enabled = bkt.Callback(TableFormat.is_paste_enabled, shapes=True),
                        ),
                    ]
                ),
                bkt.ribbon.SplitButton(
                    id = 'edge_autofixer_splitbutton',
                    children=[
                        bkt.ribbon.Button(
                            id = 'edge_autofixer',
                            label="Kanten-Autofixer",
                            # image_mso='GridSettings',
                            get_image=bkt.Callback(EdgeAutoFixer.get_image, context=True),
                            supertip="Gleicht minimale Verschiebungen der Kanten der gewählten Shapes aus.",
                            on_action=bkt.Callback(EdgeAutoFixer.autofix_edges, shapes=True),
                            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                        ),
                        bkt.ribbon.Menu(
                            label="Kanten-Autofixer Menü",
                            supertip="Einstellungsmöglichkeiten für den Kanten-Autofixer",
                            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                            children=[
                                bkt.ribbon.Button(
                                    id = 'edge_autofixer-dd',
                                    label="Kanten-Autofixer diagonal von links-oben",
                                    image='autofixer_dd',
                                    supertip="Gleicht minimale Verschiebungen der Kanten der gewählten Shapes aus durch Vergrößerung auf Shapes links-oberhalb der anzupassenden Shapes.",
                                    on_action=bkt.Callback(EdgeAutoFixer.autofix_edges_diagonal_down, shapes=True),
                                    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                                ),
                                bkt.ribbon.Button(
                                    id = 'edge_autofixer-td',
                                    label="Kanten-Autofixer von oben nach unten",
                                    image='autofixer_td',
                                    supertip="Gleicht minimale Verschiebungen der Kanten der gewählten Shapes aus durch Vergrößerung auf Shapes links der anzupassenden Shapes.",
                                    on_action=bkt.Callback(EdgeAutoFixer.autofix_edges_top_down, shapes=True),
                                    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                                ),
                                bkt.ribbon.Button(
                                    id = 'edge_autofixer-lr',
                                    label="Kanten-Autofixer von links nach rechts",
                                    image='autofixer_lr',
                                    supertip="Gleicht minimale Verschiebungen der Kanten der gewählten Shapes aus durch Vergrößerung auf Shapes oberhalb der anzupassenden Shapes.",
                                    on_action=bkt.Callback(EdgeAutoFixer.autofix_edges_left_right, shapes=True),
                                    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                                ),
                                bkt.ribbon.MenuSeparator(),
                                bkt.ribbon.ToggleButton(
                                    label="Gruppen-Elemente einzeln anpassen",
                                    supertip="Gibt an ob Elemente einer Gruppe einzeln betrachtet werden, oder die gesamte Gruppe als Ganzes.",
                                    get_pressed=bkt.Callback(lambda: EdgeAutoFixer.groupitems is True),
                                    on_toggle_action=bkt.Callback(lambda pressed: EdgeAutoFixer.settings_setter("groupitems", pressed)),
                                ),
                                bkt.ribbon.Menu(
                                    label="Toleranz ändern",
                                    screentip="Kanten-Autofixer Toleranz",
                                    supertip="Schwellwert für Kanten-Autofixer anpassen.",
                                    children=[
                                        bkt.ribbon.ToggleButton(
                                            label="Klein 0,1cm",
                                            screentip="Toleranz klein 0,1cm",
                                            supertip="Setzt Toleranz vom Kanten-Autofixer auf klein = 0,1cm",
                                            get_pressed=bkt.Callback(lambda: EdgeAutoFixer.threshold == 0.1),
                                            on_toggle_action=bkt.Callback(lambda pressed: EdgeAutoFixer.settings_setter("threshold", 0.1)),
                                        ),
                                        bkt.ribbon.ToggleButton(
                                            label="Mittel 0,3cm",
                                            screentip="Toleranz mittel 0,3cm",
                                            supertip="Setzt Toleranz vom Kanten-Autofixer auf mittel = 0,3cm",
                                            get_pressed=bkt.Callback(lambda: EdgeAutoFixer.threshold == 0.3),
                                            on_toggle_action=bkt.Callback(lambda pressed: EdgeAutoFixer.settings_setter("threshold", 0.3)),
                                        ),
                                        bkt.ribbon.ToggleButton(
                                            label="Groß 1 cm",
                                            screentip="Toleranz groß 1 cm",
                                            supertip="Setzt Toleranz vom Kanten-Autofixer auf groß = 1cm",
                                            get_pressed=bkt.Callback(lambda: EdgeAutoFixer.threshold == 1.0),
                                            on_toggle_action=bkt.Callback(lambda pressed: EdgeAutoFixer.settings_setter("threshold", 1.0)),
                                        ),
                                    ]
                                ),
                            ]
                        ),
                    ]
                ),
                bkt.ribbon.Button(
                    id = 'arrange_shape_center',
                    label="Shape in Mitte positionieren",
                    image_mso="DiagramRadialInsertClassic",
                    screentip="Letztes Shape in Mittelpunkt positionieren",
                    supertip="Setzt das zuletzt markierte Shape in den gewichteten Mittelpunkt aller anderen markierten Shapes.",
                    on_action=bkt.Callback(ArrangeCenter.arrange_shapes, shapes=True),
                    get_enabled = bkt.apps.ppt_shapes_min2_selected,
                ),
                bkt.ribbon.MenuSeparator(title="Gruppierung"),
                bkt.ribbon.Button(
                    id = 'add_into_group',
                    label="In Gruppe einfügen",
                    image_mso="ObjectsRegroup",
                    screentip="Shapes in Gruppe einfügen",
                    supertip="Sofern das zuerst oder zuletzt markierte Shape eine Gruppe ist, werden alle anderen Shapes in diese Gruppe eingefügt. Anderenfalls werden alle Shapes gruppiert.",
                    on_action=bkt.Callback(GroupsMore.add_into_group, shapes=True),
                    get_enabled = bkt.apps.ppt_shapes_min2_selected,
                ),
                bkt.ribbon.Button(
                    id = 'recursive_ungroup',
                    label="Rekursives Gruppe aufheben",
                    image_mso="ObjectsUngroup",
                    screentip="Gruppe aufheben rekursiv ausführen",
                    supertip="Wendet Gruppe aufheben solange an, bis alle verschachtelten Gruppen aufgelöst sind.",
                    on_action=bkt.Callback(GroupsMore.recursive_ungroup, shapes=True),
                    get_enabled = bkt.Callback(GroupsMore.contains_group, shapes=True),
                ),
                bkt.ribbon.Button(
                    id = 'select_all_groupitems',
                    label="Elemente der Gruppe markieren",
                    image_mso="ObjectsMultiSelect",
                    screentip="Alle Elemente der Gruppe markieren",
                    supertip="Markiert alle Elemente innerhalb der Gruppe.",
                    on_action=bkt.Callback(GroupsMore.select_all_groupitems, shape=True),
                    get_enabled = bkt.Callback(GroupsMore.is_or_within_group, shape=True),
                ),
                bkt.ribbon.Button(
                    id = 'remove_from_group',
                    label="Aus Gruppe lösen",
                    image_mso="ObjectsUngroup",
                    screentip="Shapes aus Gruppe herauslösen",
                    supertip="Die markierten Shapes werden aus der aktuelle Gruppe herausgelöst, ohne die Gruppe dabei zu verändern.",
                    on_action=bkt.Callback(GroupsMore.remove_from_group, shapes=True),
                    get_enabled = bkt.Callback(GroupsMore.visible_remove_from_group, shapes=True),
                ),
                bkt.ribbon.MenuSeparator(title="Verknüpfte Shapes"),
                bkt.ribbon.Button(
                    id = 'shape_copy_to_all',
                    label="Shape auf Folgefolien kopieren und verknüpfen…",
                    image_mso="ShapesDuplicate",
                    screentip="Shape auf Folgefolien duplizieren",
                    supertip="Dupliziert das aktuelle Shapes auf alle Folien hinter der aktuellen Folie und verknüpft diese für zukünftige Operationen.",
                    on_action=bkt.Callback(LinkedShapes.copy_to_all),
                    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                ),
                bkt.ribbon.Button(
                    id = 'shape_find_similar_and_link',
                    label="Shape auf Folgefolien suchen und verknüpfen…",
                    image_mso="FindTag",
                    screentip="Gleiches Shape auf Folgefolien suchen und verknüpfen",
                    supertip="Sucht das aktuelle Shape auf allen Folien hinter der aktuellen Folie anhand Position und Größe und verknüpft diese miteinander.",
                    on_action=bkt.Callback(LinkedShapes.find_similar_and_link),
                    get_enabled = bkt.apps.ppt_shapes_exactly1_selected,
                ),
                bkt.ribbon.MenuSeparator(),
                bkt.ribbon.Button(
                    id = 'link_shapes',
                    label="Ausgewählte Shapes miteinander verknüpfen",
                    image_mso="HyperlinkCreate",
                    screentip="Alle ausgewählte Shapes miteinander verknüpfen",
                    supertip="Die ausgewählten Shapes für zukünftige Operationen verknüpfen. Die Verknüpfung bleibt beim Kopieren der Shapes erhalten.",
                    on_action=bkt.Callback(LinkedShapes.link_shapes, shapes=True),
                    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                ),
                bkt.ribbon.Button(
                    id = 'each_link_shapes',
                    label="Ausgewählte Shapes einzeln in Verknüpfung umwandeln",
                    # image_mso="HyperlinkCreate",
                    screentip="Alle ausgewählte Shapes einzeln verknüpfen",
                    supertip="Die ausgewählten Shapes bekommen jeweils eine interne Verknüpfungs-ID. Die Verknüpfung bleibt beim Kopieren der Shapes erhalten.",
                    on_action=bkt.Callback(LinkedShapes.each_link_shapes, shapes=True),
                    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                ),
                bkt.ribbon.Button(
                    id = 'extend_link_shapes',
                    label="Bestehende Shape-Verknüpfung erweitern",
                    # image_mso="HyperlinkCreate",
                    screentip="Bestehende Shape-Verknüpfung erweitern",
                    supertip="Um die bestehende Shape-Verknüpfung zu erweitern, wird die interne ID zwischengespeichert. Über 'Ausgewählte Shapes zur Verknüpfung hinzufügen' können dann weitere Shapes zur Verknüpfung hinzugefügt werden.",
                    on_action=bkt.Callback(LinkedShapes.extend_link_shapes, shape=True),
                    get_enabled = bkt.Callback(LinkedShapes.is_linked_shape, shape=True),
                ),
                bkt.ribbon.Button(
                    id = 'add_to_link_shapes',
                    label="Ausgewählte Shapes zur Verknüpfung hinzufügen",
                    # image_mso="HyperlinkCreate",
                    screentip="Ausgewählte Shapes zur Verknüpfung hinzufügen",
                    supertip="Ausgewählte Shapes zur zwischengespeicherten ID hinzufügen. Vorher muss eine neue Verknüpfung angelegt oder eine bestehende erweitert werden.",
                    on_action=bkt.Callback(LinkedShapes.add_to_link_shapes, shapes=True),
                    get_enabled = bkt.Callback(LinkedShapes.enabled_add_linked_shapes),
                ),
                bkt.ribbon.MenuSeparator(),
                bkt.ribbon.Button(
                    id = 'unlink_shape',
                    label="Einzelne Shape-Verknüpfung entfernen",
                    image_mso="HyperlinkRemove",
                    screentip="Verknüpfung des ausgewählten Shapes entfernen",
                    supertip="Entfernt die ID zur Verknüpfung vom aktuellen Shape. Alle anderen Shapes mit der gleichen ID bleiben verknüpft.",
                    on_action=bkt.Callback(LinkedShapes.unlink_shape, shape=True),
                    get_enabled = bkt.Callback(LinkedShapes.is_linked_shape),
                ),
                bkt.ribbon.Button(
                    id = 'unlink_all_shapes',
                    label="Gesamte Shape-Verknüpfung auflösen",
                    # image_mso="HyperlinkRemove",
                    screentip="Alle Shape-Verknüpfungen entfernen",
                    supertip="Entfernt die ID zur Verknüpfung vom aktuellen Shape sowie allen verknüpften Shapes mit der gleichen ID.",
                    on_action=bkt.Callback(LinkedShapes.unlink_all_shapes, shape=True, context=True),
                    get_enabled = bkt.Callback(LinkedShapes.is_linked_shape),
                ),

                # bkt.ribbon.MenuSeparator(),
                # bkt.ribbon.Menu(
                #     label='Verknüpfte Shapes',
                #     image_mso='ControlAlignToGrid',
                #     screentip="Operationen auf verknüpfte Shapes",
                #     supertip="Alle verknüpften Shapes löschen oder ausrichten. Optionen stehen auch im Kontextmenü von verknüpften Shapes zur Verfügung.",
                #     get_enabled = bkt.Callback(LinkedShapes.is_linked_shape, shape=True),
                #     children=[
                #         bkt.ribbon.Button(
                #             id = 'linked_shapes_count',
                #             label="Anzahl verknüpfter Shapes",
                #             image_mso="FindDialog",
                #             screentip="Verknüpfte Shapes zählen",
                #             supertip="Zählt die Anzahl der verknüpften Shapes auf allen Folien.",
                #             on_action=bkt.Callback(LinkedShapes.count_link_shapes, shape=True, context=True),
                #             # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                #         ),
                #         bkt.ribbon.Button(
                #             id = 'linked_shapes_next',
                #             label="Nächstes verknüpfte Shape finden",
                #             image_mso="FindNext",
                #             screentip="Zum nächsten verknüpften Shape gehen",
                #             supertip="Sucht nach dem nächste verknüpften Shape. Sollte auf den Folgefolien kein Shape mehr kommen, wird das erste verknüpfte Shape der Präsentation gesucht.",
                #             on_action=bkt.Callback(LinkedShapes.goto_linked_shape, shape=True, context=True),
                #             # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                #         ),
                #         bkt.ribbon.MenuSeparator(),
                #         bkt.ribbon.Button(
                #             id = 'linked_shapes_delete',
                #             label="Andere löschen",
                #             image_mso="HyperlinkRemove",
                #             screentip="Verknüpfte Shapes löschen",
                #             supertip="Alle verknüpften Shapes auf allen Folien löschen.",
                #             on_action=bkt.Callback(LinkedShapes.delete_linked_shapes, shape=True, context=True),
                #             # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                #         ),
                #         bkt.ribbon.Button(
                #             id = 'linked_shapes_replace',
                #             label="Andere mit diesem ersetzen",
                #             image_mso="HyperlinkCreate",
                #             screentip="Verknüpfte Shapes ersetzen",
                #             supertip="Alle verknüpften Shapes auf allen Folien mit ausgewähltem Shape ersetzen.",
                #             on_action=bkt.Callback(LinkedShapes.replace_with_this, shape=True, context=True),
                #             # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                #         ),
                #         bkt.ribbon.MenuSeparator(),
                #         bkt.ribbon.Button(
                #             id = 'linked_shapes_align',
                #             label="Position angleichen",
                #             image_mso="ControlAlignToGrid",
                #             screentip="Position verknüpfter Shapes angleichen",
                #             supertip="Position und Rotation aller verknüpfter Shapes auf Position wie ausgewähltes Shape setzen.",
                #             on_action=bkt.Callback(LinkedShapes.align_linked_shapes, shape=True, context=True),
                #             # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                #         ),
                #         bkt.ribbon.Button(
                #             id = 'linked_shapes_size',
                #             label="Größe angleichen",
                #             image_mso="SizeToControlHeightAndWidth",
                #             screentip="Größe verknüpfter Shapes angleichen",
                #             supertip="Größe aller verknüpfter Shapes auf Größe wie ausgewähltes Shape setzen.",
                #             on_action=bkt.Callback(LinkedShapes.size_linked_shapes, shape=True, context=True),
                #             # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                #         ),
                #         bkt.ribbon.Button(
                #             id = 'linked_shapes_format',
                #             label="Formatierung angleichen",
                #             image_mso="FormatPainter",
                #             screentip="Formatierung verknüpfter Shapes angleichen",
                #             supertip="Formatierung aller verknüpfter Shapes auf Größe wie ausgewähltes Shape setzen.",
                #             on_action=bkt.Callback(LinkedShapes.format_linked_shapes, shape=True, context=True),
                #             # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                #         ),
                #         bkt.ribbon.Button(
                #             id = 'linked_shapes_text',
                #             label="Text angleichen",
                #             image_mso="TextBoxInsert",
                #             screentip="Text verknüpfter Shapes angleichen",
                #             supertip="Text aller verknüpfter Shapes auf Größe wie ausgewähltes Shape setzen.",
                #             on_action=bkt.Callback(LinkedShapes.text_linked_shapes, shape=True, context=True),
                #             # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                #         ),
                #         bkt.ribbon.MenuSeparator(),
                #         bkt.ribbon.Button(
                #             id = 'linked_shapes_all',
                #             label="Alles angleichen",
                #             image_mso="GroupUpdate",
                #             screentip="Alle Eigenschaften verknüpfter Shapes angleichen",
                #             supertip="Alle Eigenschaften aller verknüpfter Shapes wie ausgewähltes Shape setzen.",
                #             on_action=bkt.Callback(LinkedShapes.equalize_linked_shapes, shape=True, context=True),
                #             # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                #         ),
                #     ]
                # )
                ]
        )