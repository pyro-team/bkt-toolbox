# -*- coding: utf-8 -*-
'''
Created on 06.07.2016

@author: rdebeerst
'''



# import logging
import locale

from System import Array

import bkt
import bkt.library.powerpoint as pplib
pt_to_cm = pplib.pt_to_cm
cm_to_pt = pplib.cm_to_pt
get_ambiguity_tuple = bkt.helpers.get_ambiguity_tuple

from bkt.library.algorithms import get_bounding_nodes, mid_point

from bkt import dotnet
Drawing = dotnet.import_drawing()
office = dotnet.import_officecore()

# other toolbox modules
# from .chartlib import shapelib_button
# from .agenda import ToolboxAgenda
# from . import text
from .. import harvey
from .. import stateshapes


class ShapeDialogs(object):
    
    ### DIALOG WINDOWS ###

    @staticmethod
    def shape_split(context, shapes):
        from ..dialogs.shape_split import ShapeSplitWindow
        ShapeSplitWindow.create_and_show_dialog(context, shapes)

    @staticmethod
    def shape_scale(context, shapes):
        from ..dialogs.shape_scale import ShapeScaleWindow
        ShapeScaleWindow.create_and_show_dialog(context, shapes)
    
    @staticmethod
    def show_segmented_circle_dialog(context, slide):
        from ..dialogs.circular_segments import SegmentedCircleWindow
        SegmentedCircleWindow.create_and_show_dialog(context, slide)

    @staticmethod
    def show_process_chevrons_dialog(context, slide):
        from ..dialogs.shape_process import ProcessWindow
        ProcessWindow.create_and_show_dialog(context, slide)

    ### DIRECT CREATE ###

    @staticmethod
    def create_headered_pentagon(slide):
        from .processshapes import Pentagon
        Pentagon.create_headered_pentagon(slide)

    @staticmethod
    def create_headered_chevron(slide):
        from .processshapes import Pentagon
        Pentagon.create_headered_chevron(slide)
    
    @staticmethod
    def create_traffic_light(slide, style):
        from ..popups.traffic_light import Ampel
        Ampel.create(slide, style)


class TrackerShape(object):

    @classmethod
    def generateTracker(cls, shapes, context):
        import uuid
        from ..linkshapes import LinkedShapes

        #shapes to copy formatting
        shapes_count = len(shapes)
        highlight_shape = shapes[0]
        default_shape = shapes[1]
        slide_width = context.app.ActivePresentation.PageSetup.SlideWidth

        #copy and paste shapes (note: shapes can also be part of a group)
        pplib.shapes_to_range(shapes).copy()
        grp = context.slide.shapes.paste().group()

        # format unselected elements
        for shp in grp.GroupItems:
            if shp.HasTextFrame:
                shp.TextFrame.DeleteText()
            default_shape.PickUp()
            shp.Apply()

        # generate unique GUID fpr tracker (items)
        tracker_guid = str(uuid.uuid4())

        # format each selected element and paste tracker as image
        for i in range(1, shapes_count+1):
            highlight_shape.PickUp()
            new_grp = grp.Duplicate()
            
            new_grp.GroupItems(i).Apply()
            new_grp.Copy()

            tracker = context.slide.shapes.PasteSpecial(DataType=6) # ppPastePNG = 6
            tracker.Tags.Add("tracker_id", tracker_guid)

            new_grp.Delete()

            tracker.Height = cm_to_pt(1.5)
            tracker.left = slide_width - cm_to_pt(3.0) - shapes_count*cm_to_pt(1/1.5) + cm_to_pt(i/1.5)
            tracker.top = cm_to_pt(3.0) + cm_to_pt(i/1.5)

        #delete duplicated shapes
        grp.Delete()
        all_trackers = pplib.last_n_shapes_on_slide(context.slide, shapes_count)
        all_trackers_list = list(iter(all_trackers))

        #select all tracker
        all_trackers.select()

        #make trackers linked shapes
        LinkedShapes.link_shapes(all_trackers_list)

        #ask to distribute trackers
        if bkt.message.confirmation("Tracker auf Folgefolien verteilen?"):
            cls.distributeTracker(all_trackers_list, context)
            all_trackers_list[0].select()

    @staticmethod
    def isTracker(shape):
        return pplib.TagHelper.has_tag(shape, "tracker_id")

    @staticmethod
    def alignTracker(shape, context):
        tracker_id = shape.Tags.Item("tracker_id")
        if not tracker_id:
            return
        
        tracker_position_left = shape.left
        tracker_position_top = shape.top
        tracker_rotation = shape.Rotation
        tracker_heigth = shape.Height
        tracker_width = shape.Width
        tracker_lock_ar = shape.LockAspectRatio
        
        for sld in context.app.ActivePresentation.Slides:
            for cShp in sld.shapes:
                if cShp.Tags.Item("tracker_id") == tracker_id:
                    cShp.LockAspectRatio = 0 #msoFalse
                    cShp.left, cShp.top = tracker_position_left, tracker_position_top
                    cShp.Height, cShp.Width = tracker_heigth, tracker_width
                    cShp.Rotation = tracker_rotation
                    cShp.LockAspectRatio = tracker_lock_ar

    @staticmethod
    def removeTracker(shape, context):
        tracker_id = shape.Tags.Item("tracker_id")
        if not tracker_id:
            return
        
        for sld in context.app.ActivePresentation.Slides:
            for cShp in sld.shapes:
                if cShp.Tags.Item("tracker_id") == tracker_id:
                    cShp.Delete()

    @classmethod
    def distributeTracker(cls, shapes, context):
        cur_slide_index = shapes[0].Parent.SlideIndex
        max_index = context.app.ActivePresentation.Slides.Count
        for shape in shapes[1:]:
            cur_slide_index = min(max_index, cur_slide_index+1)
            shape.Cut()
            context.app.ActivePresentation.Slides[cur_slide_index].Shapes.Paste()

        cls.alignTracker(shapes[0], context)


class NumberedShapes(object):
    
    label = "1"                 # 1->1,2,3   a->a,b,c   A->A,B,C   I->I,II,III
    shape_type = "square"       # square, circle
    style = "dark"              # dark, light
    position = "top-left"       # top-left, top-right
    position_offset = True      # True, False
    
    # label_1 = [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26]
    # label_a = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z']
    # label_A = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
    # label_I = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI', 'XII', 'XIII', 'XIV', 'XV', 'XVI', 'XVII', 'XVIII', 'XIX', 'XX', 'XXI', 'XXII', 'XXIII', 'XXIV', 'XXV', 'XXVI']
    
    _count_formatter = None
    @classmethod
    def get_count_formatter(cls):
        if not cls._count_formatter:
            from formatter import AbstractFormatter, DumbWriter
            cls._count_formatter = AbstractFormatter(DumbWriter())
        return cls._count_formatter

    
    @classmethod
    def create_numbers_for_shapes(cls, slide, shapes, **kwargs):
        
        settings = {
            # default settings
            'label': cls.label,
            'shape_type': cls.shape_type,
            'style': cls.style,
            'position': cls.position,
            'position_offset': cls.position_offset
        }
        # default settings are overwritten by key-word-arguments
        settings.update(kwargs)
        
        len_shapes = 0
        for number, shape in enumerate(shapes, start=1):
            cls.create_number_shape(slide, shape, number, **settings)
            len_shapes = number
        
        pplib.last_n_shapes_on_slide(slide, len_shapes).select()
        
        
    
    @classmethod
    def create_number_shape(cls, slide, shape, number, label='1', shape_type='square', style='dark', position='top-left', position_offset=True):
        
        if shape_type == 'square':
            numshape = slide.shapes.addshape( pplib.MsoAutoShapeType['msoShapeRectangle'] , shape.left, shape.top, 14, 14)
        elif shape_type == 'diamond':
            numshape = slide.shapes.addshape( pplib.MsoAutoShapeType['msoShapeDiamond'] , shape.left, shape.top, 14, 14)
        else: #circle
            numshape = slide.shapes.addshape( pplib.MsoAutoShapeType['msoShapeOval'] , shape.left, shape.top, 14, 14)
        
        numshape.LockAspectRatio = -1

        if style == "dark":
            col_background = 13 #msoThemeColorText1
            col_foreground = 14 #msoThemeColorBackground1
        else:
            col_background = 14 #msoThemeColorBackground1
            col_foreground = 13 #msoThemeColorText1

        numshape.Line.Visible = -1
        numshape.Line.ForeColor.ObjectThemeColor = col_foreground
        numshape.Fill.Visible = -1
        numshape.Fill.ForeColor.ObjectThemeColor = col_background

        # if style == "dark":
        #     numshape.line.visible = False
        #     numshape.fill.ForeColor.RGB = 0
        #     numshape.TextFrame.TextRange.Font.Color.rgb = 255 + 255 * 256 + 255 * 256**2
            
        # else: # light
        #     numshape.line.style = 1
        #     numshape.line.weight = 1
        #     numshape.line.ForeColor.RGB = 0
        #     numshape.fill.ForeColor.RGB = 255 + 255 * 256 + 255 * 256**2
        #     numshape.TextFrame.TextRange.Font.Color.rgb = 0
        
        # positions corrections for rounded rectangles and pentagon/chevron-shapes
        pos_correction_l = 0
        pos_correction_r = 0
        if shape.AutoShapeType == pplib.MsoAutoShapeType['msoShapeRoundedRectangle']:
            pos_correction_l = shape.Adjustments.item[1] * min(shape.Height, shape.Width)
            pos_correction_r = pos_correction_r
        if shape.AutoShapeType in [pplib.MsoAutoShapeType['msoShapeChevron'], pplib.MsoAutoShapeType['msoShapePentagon']]:
            pos_correction_r = shape.Adjustments.item[1] * min(shape.Height, shape.Width)
        
        # set position
        if position == "top-right":
            numshape.left = shape.left+shape.width-numshape.width -pos_correction_r
            if position_offset:
                numshape.left += numshape.width/2
                numshape.top -= numshape.height/2
        else: # top-left
            numshape.left += pos_correction_l
            if position_offset:
                numshape.left -= numshape.width/2
                numshape.top -= numshape.height/2
        
        # format shape and text
        # numshape.TextFrame.TextRange.text = getattr(cls, 'label_' + label)[(number-1)%26] #at number 26 start from beginning to avoid IndexError
        textframe = numshape.TextFrame2
        textframe.TextRange.Text = cls.get_count_formatter().format_counter(label, number)
        textframe.TextRange.Font.Size = 12
        textframe.TextRange.Font.Fill.ForeColor.ObjectThemeColor = col_foreground
        textframe.TextRange.ParagraphFormat.Alignment = pplib.PowerPoint.PpParagraphAlignment.ppAlignCenter.value__
        textframe.TextRange.ParagraphFormat.Bullet.Type = 0
        textframe.AutoSize = 0
        textframe.WordWrap = False
        textframe.MarginTop = 0
        textframe.MarginLeft = 0
        textframe.MarginRight = 0
        textframe.MarginBottom = 0
        #textframe.HorizontalAnchor = office.MsoHorizontalAnchor.msoAnchorCenter.value__
        textframe.VerticalAnchor = office.MsoVerticalAnchor.msoAnchorMiddle.value__
        
        return numshape


class NumberShapesGallery(bkt.ribbon.Gallery):
    
    # item-settings for gallery
    items = [ dict(label=l, style=s, shape_type=t) for l in ['1', 'a', 'A', 'i', 'I'] for t in ['circle', 'square', 'diamond'] for s in ['dark', 'light'] ]
    item_cols = 6
    
    position = "top-left"
    position_offset = True
    
    def __init__(self, **kwargs):
        parent_id = kwargs.get('id') or ""
        my_kwargs = dict(
            label = 'Nummerierung',
            columns = self.item_cols,
            screentip="Nummerierungs-Shapes einfügen",
            supertip="Fügt für jedes markierte Shape ein Nummerierungs-Shape ein. Nummerierung und Styling entsprechend der Auswahl. Markierte Shapes werden entsprechend der Selektions-Reihenfolge durchnummeriert.",
            get_image=bkt.Callback(lambda: self.get_item_image(0)),
            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
            item_width=24,
            item_height=24,
            children=[
                bkt.ribbon.Button(id=parent_id + "_pos_left", label="Position links oben", screentip="Nummerierungs-Shapes links-oben",    on_action=bkt.Callback(self.set_pos_top_left), get_image=bkt.Callback(lambda: self.get_toggle_image('pos-top-left')),
                    supertip="Nummerierungs-Shapes links oben auf dem zugehörigen Shape platzieren"),
                bkt.ribbon.Button(id=parent_id + "_pos-right", label="Position rechts oben", screentip="Nummerierungs-Shapes rechts-oben", on_action=bkt.Callback(self.set_pos_top_right), get_image=bkt.Callback(lambda: self.get_toggle_image('pos-top-right')),
                    supertip="Nummerierungs-Shapes rechts oben auf dem zugehörigen Shape platzieren"),
                bkt.ribbon.Button(id=parent_id + "_pos-offset", label="Versetzt positionieren", screentip="Nummerierungs-Shapes versetzt positionieren", on_action=bkt.Callback(self.toggle_pos_offset), get_image=bkt.Callback(lambda: self.get_toggle_image('pos-offset')),
                    supertip="Standardmäßig werden Nummerierungs-Shapes genau am Rand des zugehörigen Shapes ausgerichtet.\n\nIst 'Versetzt positionieren' aktiviert, werden die Nummerierungs-Shapes etwas weiter außerhalb des zugehörigen Shapes plaziert, so dass der Mittelpunkt des Nummerierungs-Shapes auf der Ecke liegt.")
            ],
        )
        my_kwargs.update(kwargs)

        super(NumberShapesGallery, self).__init__(**my_kwargs)
    
    
    def on_action_indexed(self, selected_item, index, slide, shapes):
        ''' create numberd shape according of settings in clicked element '''
        item = self.items[index]
        NumberedShapes.create_numbers_for_shapes(slide, shapes, label=item['label'], shape_type=item['shape_type'], style=item['style'], position=self.position, position_offset=self.position_offset)

                
    def get_item_count(self):
        return len(self.items)
    
    # def get_item_label(self, index):
    #     item = self.items[index]
    #     return "%s" % getattr(NumberedShapes, 'label_' + item['label'])[index%self.columns]
    
    def get_item_screentip(self, index):
        return "Nummerierungs-Shapes einfügen"
        
    def get_item_supertip(self, index):
        return "Fügt für jedes markierte Shape ein Nummerierungs-Shape ein. Nummerierung und Styling entsprechend der Auswahl. Markierte Shapes werden entsprechend der Selektions-Reihenfolge durchnummeriert."
    
    def get_item_image(self, index):
        ''' creates an item image with numberd shape according to settings in the specified item '''
        # retrieve item-settings
        item = self.items[index]
        
        # create bitmap, define pen/brush
        size = 48
        img = Drawing.Bitmap(size, size)
        g = Drawing.Graphics.FromImage(img)
        
        #Draw smooth rectangle/ellipse
        g.SmoothingMode = Drawing.Drawing2D.SmoothingMode.AntiAlias

        if item['style'] == "dark":
            pen_border = Drawing.Pen(Drawing.Color.White,2)
            brush_fill = Drawing.Brushes.Black
            text_brush = Drawing.Brushes.White
        else:
            pen_border = Drawing.Pen(Drawing.Color.Black,2)
            brush_fill = Drawing.Brushes.White
            text_brush = Drawing.Brushes.Black

        if item['shape_type'] == 'circle':
            g.FillEllipse(brush_fill, 2, 2, size-4, size-4) #left, top, width, height
            g.DrawEllipse(pen_border, 2, 2, size-4, size-4) #left, top, width, height
        elif item['shape_type'] == 'diamond':
            diamond_points = [(0,1),(1,2),(2,1),(1,0)]
            size_factor = size/2
            points = Array[Drawing.Point]([Drawing.Point(round(l*size_factor),round(t*size_factor)) for t,l in diamond_points])
            g.FillPolygon(brush_fill, points)
            g.DrawPolygon(pen_border, points)
        else: #fallback shape=1 rectangle
            g.FillRectangle(brush_fill, 2, 2, size-4, size-4) #left, top, width, height
            g.DrawRectangle(pen_border, 2, 2, size-4, size-4) #left, top, width, height
        
        # color_black = Drawing.Color.Black
        # if item['style'] == 'dark':
        #     # create black circle/rectangle
        #     brush = Drawing.SolidBrush(color_black)
        #     text_brush = Drawing.Brushes.White

        #     if item['shape_type'] == 'circle':
        #         g.FillEllipse(brush, 2,2, size-5, size-5)
        #     else: #square
        #         g.FillRectangle(brush, Drawing.Rectangle(2,2, size-5, size-5))

        # else: # light
        #     # create white circle/rectangle
        #     text_brush = Drawing.Brushes.Black
        #     pen = Drawing.Pen(color_black,2)

        #     if item['shape_type'] == 'circle':
        #         g.DrawEllipse(pen, 2,1, size-4, size-4)
        #     else: #square
        #         g.DrawRectangle(pen, Drawing.Rectangle(2,2, size-4, size-4))

        # set string format
        strFormat = Drawing.StringFormat()
        strFormat.Alignment = Drawing.StringAlignment.Center
        strFormat.LineAlignment = Drawing.StringAlignment.Center
        
        # draw string
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAliasGridFit
        # g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAlias
        # g.DrawString(str(getattr(NumberedShapes, 'label_' + item['label'])[index%int(self.columns)]),
        g.DrawString(NumberedShapes.get_count_formatter().format_counter(item['label'], index%int(self.item_cols)+1),
                     Drawing.Font("Arial", 32, Drawing.FontStyle.Bold, Drawing.GraphicsUnit.Pixel), text_brush, 
                     # Drawing.Font("Arial", 7, Drawing.FontStyle.Bold), text_brush, 
                     Drawing.RectangleF(1, 2, size, size-1), 
                     strFormat)
        
        return img
    
    def set_pos_top_left(self):
        self.position = 'top-left'
    
    def set_pos_top_right(self):
        self.position = 'top-right'
    
    def toggle_pos_offset(self):
        self.position_offset = not self.position_offset
    
    def get_toggle_image(self, key):
        if key == 'pos-top-left':
            pressed = (self.position == 'top-left')
        elif key == 'pos-top-right':
            pressed = (self.position == 'top-right')
        elif key == 'pos-offset':
            pressed = self.position_offset
        else:
            pressed = False

        if pressed:
            return self.get_check_image()
        else:
            return self.get_check_image(checked=False)


class ShapeConnectorTags(pplib.BKTTag):
    TAG_NAME = "BKT_SHAPE_CONNECTORS"

class ShapeConnectors(object):
    _default_shape_nodes = dict(top=(0,3), right=(3,2), bottom=(1,2), left=(0,1))
    _special_shape_nodes = {
        pplib.MsoAutoShapeType["msoShapeChevron"]: dict(top=(0,1), right=(1,3), bottom=(3,4), left=(4,0)),
        pplib.MsoAutoShapeType["msoShapePentagon"]: dict(top=(0,1), right=(1,3), bottom=(3,4), left=(4,0)),
        pplib.MsoAutoShapeType["msoShapeHexagon"]: dict(top=(1,2), right=(2,4), bottom=(4,5), left=(5,1)),
        pplib.MsoAutoShapeType["msoShapeOval"]: dict(top=(0,6), right=(3,9), bottom=(6,0), left=(9,3)),
    }

    @staticmethod
    def is_connector(shape):
        return pplib.TagHelper.has_tag(shape, ShapeConnectorTags.TAG_NAME)
        # return shape.Tags.Item(ShapeConnectorTags.TAG_NAME) != '' #FIXME: EnvironmentError for fancy smart-shapes

    @staticmethod
    def _find_shape_by_id(slide, shape_id):
        for shp in slide.shapes:
            if shp.id == shape_id:
                return shp
        else:
            raise IndexError("shape id not found on slide")

    @classmethod
    def _get_shape_connector_nodes(cls, shape, side):
        dummy = None
        try:
            special_nodes = cls._special_shape_nodes[shape.AutoShapeType]
            #convert into freeform by adding and deleting in order to acces points
            dummy = shape.duplicate()
            dummy.left, dummy.top = shape.left, shape.top
            dummy.nodes.insert(1,0,0,0,0)
            dummy.nodes.delete(2)
            shape_nodes = [(node.points[0,0], node.points[0,1]) for node in dummy.nodes]
            shape_p1, shape_p2 = special_nodes[side]
            return shape_nodes[shape_p1], shape_nodes[shape_p2]
        except: #KeyError, or any COM Error
            shape_nodes = get_bounding_nodes(shape)
            shape_p1, shape_p2 = cls._default_shape_nodes[side]
            return shape_nodes[shape_p1], shape_nodes[shape_p2]
        finally:
            if dummy:
                dummy.delete()
    
    @classmethod
    def _set_connector_shape_nodes(cls, shape_connector, shape1, shape2, shape1_side="bottom", shape2_side="top"):
        from math import atan2

        shape1_p1, shape1_p2 = cls._get_shape_connector_nodes(shape1, shape1_side)
        shape2_p1, shape2_p2 = cls._get_shape_connector_nodes(shape2, shape2_side)

        connector_nodes = [shape1_p1, shape1_p2, shape2_p1, shape2_p2]
        #correct ordering is the key to set nodes, here clockwise ordering (left-top, right-top, right-bottom, left-bottom)
        mid_p = mid_point(connector_nodes)
        connector_nodes.sort(key=lambda p: atan2(p[1]-mid_p[1], p[0]-mid_p[0]))

        #convert shape into freeform by adding and deleting node (not sure if this is required)
        shape_connector.Nodes.Insert(1, 0, 0, 0, 0) #msoSegmentLine, msoEditingAuto, x, y
        shape_connector.Nodes.Delete(2)
        # set nodes (rectangle has 5 nodes as start and end node are the same)
        shape_connector.Nodes.SetPosition(1, connector_nodes[0][0], connector_nodes[0][1]) #top-left start node
        shape_connector.Nodes.SetPosition(2, connector_nodes[1][0], connector_nodes[1][1]) #top-right node
        shape_connector.Nodes.SetPosition(3, connector_nodes[2][0], connector_nodes[2][1]) #bottom-right node
        shape_connector.Nodes.SetPosition(4, connector_nodes[3][0], connector_nodes[3][1]) #bottom-left node
        shape_connector.Nodes.SetPosition(5, connector_nodes[0][0], connector_nodes[0][1]) #top-left end node

    @classmethod
    def update_connector_shape(cls, context, shape):
        with ShapeConnectorTags(shape.Tags) as tags:
            slide = context.slide
            try:
                shape1 = cls._find_shape_by_id(slide, tags["shape1_id"])
                shape2 = cls._find_shape_by_id(slide, tags["shape2_id"])
            except IndexError:
                bkt.message.error("Fehler: Verbundenes Shape nicht gefunden!")
            else:
                cls._set_connector_shape_nodes(shape, shape1, shape2, tags["shape1_side"], tags["shape2_side"])

    @classmethod
    def add_connector_shape(cls, slide, shape1, shape2, shape1_side="bottom", shape2_side="top"):
        shp_connector = slide.shapes.AddShape(
            1, #msoShapeRectangle
            1,1, #left-top
            10,10 #width-height
        )

        cls._set_connector_shape_nodes(shp_connector, shape1, shape2, shape1_side, shape2_side)

        # shp_connector.Fill.ForeColor.RGB = 12566463 #193
        shp_connector.Fill.ForeColor.ObjectThemeColor = 16 #Background 2
        # shp_connector.Line.ForeColor.RGB = 8355711 # 127 127 127
        shp_connector.Line.ForeColor.ObjectThemeColor = 15 #Text 2
        # shp_connector.Line.Weight = 0.75
        shp_connector.Line.Visible = -1 #msoTrue

        shp_connector.Name = "[BKT] Connector %s" % shp_connector.id

        with ShapeConnectorTags(shp_connector.Tags) as tags:
            tags["shape1_id"]   = shape1.id
            tags["shape1_side"] = shape1_side
            tags["shape2_id"]   = shape2.id
            tags["shape2_side"] = shape2_side

        return shp_connector


    @classmethod
    def addHorizontalConnector(cls, shapes, context):
        shapes = sorted(shapes, key=lambda shape: shape.Left)

        cls.add_connector_shape(context.slide, shapes[0], shapes[1], "right", "left").select()

        # shpLeft  = shapes[0]
        # shpRight = shapes[1]

        # shpConnector = context.app.ActivePresentation.Slides(context.app.ActiveWindow.View.Slide.SlideIndex).shapes.AddShape(
        #     1, #msoShapeRectangle
        #     shpLeft.Left + shpLeft.Width, shpLeft.Top,
        #     shpRight.Left - shpLeft.Left - shpLeft.width, shpLeft.Height)
        # # node 2: top right
        # shpConnector.Nodes.SetPosition(2, shpRight.Left, shpRight.Top)
        # # node 3: bottom right
        # shpConnector.Nodes.SetPosition(3, shpRight.Left, shpRight.Top + shpRight.Height)
        # shpConnector.Fill.ForeColor.RGB = 12566463 #193
        # shpConnector.Line.ForeColor.RGB = 8355711 # 127 127 127
        # shpConnector.Line.Weight = 0.75
        # shpConnector.Select()

    @classmethod
    def addVerticalConnector(cls, shapes, context):
        shapes = sorted(shapes, key=lambda shape: shape.Top)

        cls.add_connector_shape(context.slide, shapes[0], shapes[1], "bottom", "top").select()

        # shpTop = shapes[0]
        # shpBottom = shapes[1]

        # shpConnector = context.app.ActivePresentation.Slides(context.app.ActiveWindow.View.Slide.SlideIndex).shapes.AddShape(
        #     1, #msoShapeRectangle,
        #     shpTop.Left, shpTop.Top + shpTop.Height,
        #     shpTop.Width, shpBottom.Top - shpTop.Top - shpTop.Height)

        # # node 3: bottom right
        # shpConnector.Nodes.SetPosition(3, shpBottom.Left + shpBottom.width, shpBottom.Top)
        # # node 4: bottom left
        # shpConnector.Nodes.SetPosition(4, shpBottom.Left, shpBottom.Top)
        # shpConnector.Fill.ForeColor.RGB = 12566463 # 193
        # shpConnector.Line.ForeColor.RGB = 8355711 # 127 127 127
        # shpConnector.Line.Weight = 0.75
        # shpConnector.Select()


shapes_interactive_menu = lambda: bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None, 
                children=[
                    bkt.ribbon.MenuSeparator(title="Einfügehilfen"),
                    bkt.ribbon.Button(
                        id = 'segmented_circle',
                        label = "Kreissegmente…",
                        image = "segmented_circle",
                        screentip="Kreissegmente einfügen",
                        supertip="Erstelle Kreis mit Segmenten oder Chevrons.",
                        on_action=bkt.Callback(ShapeDialogs.show_segmented_circle_dialog)
                    ),
                    bkt.ribbon.Button(
                        id='agenda_textbox',
                        label="Agenda-Textbox einfügen",
                        supertip="Standard Agenda-Textbox einfügen, um daraus eine aktualisierbare Agenda zu generieren.",
                        imageMso="TextBoxInsert",
                        on_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "create_agenda_textbox_on_slide", slide=True, context=True)
                    ),
                    NumberShapesGallery(id='number-labels-gallery'),
                    bkt.ribbon.Menu(
                        label='Grafik-Tracker',
                        image = "Tracker",
                        screentip="Tracker erstellen oder ausrichten",
                        supertip="Einen Tracker aus einer Auswahl als Bild erstellen, verteilen und ausrichten.",
                        get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                        children = [
                            bkt.ribbon.Button(
                                id = 'tracker',
                                label = "Tracker aus Auswahl erstellen",
                                #image = "Tracker",
                                screentip="Tracker aus Auswahl erstellen",
                                supertip="Erstelle aus den markierten Shapes einen Tracker.\nDer Shape-Stil für Highlights wird aus dem zuerst markierten Shape (in der Regel oben links) bestimmt. Der Shape-Stil für alle anderen Shapes wird aus dem als zweites markierten Shape bestimmt.",
                                on_action=bkt.Callback(TrackerShape.generateTracker, shapes=True, shapes_min=2, context=True),
                                get_enabled = bkt.apps.ppt_shapes_min2_selected,
                            ),
                            bkt.ribbon.Button(
                                id = 'tracker_distribute',
                                label = "Tracker auf Folien verteilen",
                                #image = "Tracker",
                                screentip="Alle Tracker verteilen",
                                supertip="Verteilen der ausgewählten Tracker auf die Folgefolien und ausrichten.",
                                on_action=bkt.Callback(TrackerShape.distributeTracker, shapes=True, shapes_min=2, context=True),
                                get_enabled = bkt.apps.ppt_shapes_min2_selected,
                            ),
                            bkt.ribbon.MenuSeparator(),
                            bkt.ribbon.Button(
                                id = 'tracker_align',
                                label = "Alle Tracker ausrichten",
                                #image = "Tracker",
                                screentip="Alle Tracker ausrichten",
                                supertip="Ausrichten (Position, Größe, Rotation) aller Tracker (auf allen Folien) anhand des ausgewählten Tracker.",
                                on_action=bkt.Callback(TrackerShape.alignTracker, shape=True, context=True),
                                get_enabled = bkt.Callback(TrackerShape.isTracker, shape=True),
                            ),
                            bkt.ribbon.Button(
                                id = 'tracker_remove',
                                label = "Alle Tracker löschen",
                                #image = "Tracker",
                                screentip="Alle Tracker löschen",
                                supertip="Löschen aller Tracker (auf allen Folien) anhand des ausgewählten Tracker.",
                                on_action=bkt.Callback(TrackerShape.removeTracker, shape=True, context=True),
                                get_enabled = bkt.Callback(TrackerShape.isTracker, shape=True),
                            ),
                        ]
                    ),
                    bkt.ribbon.MenuSeparator(title="Interaktive Formen"),
                    bkt.ribbon.Button(
                        id = 'standard_process',
                        label = "Prozesspfeile…",
                        image = "process_chevrons",
                        screentip="Prozess-Pfeile einfügen",
                        supertip="Erstelle Standard Prozess-Pfeile.",
                        on_action=bkt.Callback(ShapeDialogs.show_process_chevrons_dialog)
                    ),
                    bkt.ribbon.Button(
                        id = 'headered_pentagon',
                        label = "Prozessschritt mit Kopfzeile",
                        image = "headered_pentagon",
                        screentip="Prozess-Schritt-Shape mit Kopfzeile erstellen",
                        supertip="Erstelle einen Prozess-Pfeil mit Header-Shape. Das Header-Shape kann im Prozess-Pfeil über Kontext-Menü des Header-Shapes passend angeordnet werden.",
                        on_action=bkt.Callback(ShapeDialogs.create_headered_pentagon)
                    ),
                    bkt.ribbon.Button(
                        id = 'headered_chevron',
                        label = "2. Prozessschritt mit Kopfzeile",
                        image = "headered_chevron",
                        screentip="Prozess-Schritt-Shape mit Kopfzeile erstellen",
                        supertip="Erstelle einen Prozess-Pfeil mit Header-Shape. Das Header-Shape kann im Prozess-Pfeil über Kontext-Menü des Header-Shapes passend angeordnet werden.",
                        on_action=bkt.Callback(ShapeDialogs.create_headered_chevron)
                    ),
                    harvey.harvey_create_button,
                    bkt.ribbon.Menu(
                        id="traffic_light_menu",
                        label="Ampel",
                        image="traffic_light",
                        screentip='Status-Ampel erstellen',
                        children=[
                            bkt.ribbon.Button(
                                id="traffic_light",
                                label="Ampel vertikal",
                                image="traffic_light",
                                screentip='Status-Ampel vertikal erstellen',
                                supertip="Füge eine Status-Ampel ein. Die Status-Farbe der Ampel kann per Kontext-Dialog konfiguriert werden.",
                                on_action=bkt.Callback(lambda slide: ShapeDialogs.create_traffic_light(slide, "vertical"), slide=True)
                            ),
                            bkt.ribbon.Button(
                                label="Ampel horizontal",
                                image="traffic_light2",
                                screentip='Status-Ampel horizontal erstellen',
                                supertip="Füge eine Status-Ampel ein. Die Status-Farbe der Ampel kann per Kontext-Dialog konfiguriert werden.",
                                on_action=bkt.Callback(lambda slide: ShapeDialogs.create_traffic_light(slide, "horizontal"), slide=True)
                            ),
                            bkt.ribbon.Button(
                                label="Ampel Punkt",
                                image="traffic_light3",
                                screentip='Status-Ampel einfach erstellen',
                                supertip="Füge eine Status-Ampel ein. Die Status-Farbe der Ampel kann per Kontext-Dialog konfiguriert werden.",
                                on_action=bkt.Callback(lambda slide: ShapeDialogs.create_traffic_light(slide, "simple"), slide=True)
                            ),
                        ]
                    ),
                    stateshapes.likert_button,
                    stateshapes.checkbox_button,
                    bkt.ribbon.MenuSeparator(title="Verbindungsflächen"),
                    bkt.ribbon.Button(
                        id = 'connector_h',
                        label = "Horizontale Verbindungsfläche",
                        image = "ConnectorHorizontal",
                        supertip="Erstelle eine horizontale Verbindungsfläche zwischen den vertikalen Seiten (links/rechts) von zwei Shapes.",
                        on_action=bkt.Callback(ShapeConnectors.addHorizontalConnector, context=True, shapes=True, shapes_min=2, shapes_max=2),
                        get_enabled = bkt.apps.ppt_shapes_exactly2_selected,
                    ),
                    bkt.ribbon.Button(
                        id = 'connector_v',
                        label = "Vertikale Verbindungsfläche",
                        image = "ConnectorVertical",
                        supertip="Erstelle eine vertikale Verbindungsfläche zwischen den horizontalen Seiten (oben/unten) von zwei Shapes.",
                        on_action=bkt.Callback(ShapeConnectors.addVerticalConnector, context=True, shapes=True, shapes_min=2, shapes_max=2),
                        get_enabled = bkt.apps.ppt_shapes_exactly2_selected,
                    ),
                    bkt.ribbon.MenuSeparator(),
                    bkt.ribbon.Button(
                        id = 'connector_update',
                        label = "Verbindungsfläche neu verbinden",
                        image = "ConnectorUpdate",
                        supertip="Aktualisiere die Verbindungsfläche nachdem sich die verbundenen Shapes geändert haben.",
                        on_action=bkt.Callback(ShapeConnectors.update_connector_shape, context=True, shape=True),
                        get_enabled = bkt.Callback(ShapeConnectors.is_connector, shape=True),
                    ),
                ]
            )


class ShapeTableGallery(bkt.ribbon.Gallery):
    
    # item-settings for gallery
    #items = [ dict(label=l, style=s, shape_type=t) for l in ['1', 'a', 'A', 'I'] for t in ['circle', 'square'] for s in ['dark', 'light']  ]
    _columns = 6
    _rows = 8
    
    
    def __init__(self, **kwargs):
        self._margin = 0
        parent_id = kwargs.get('id') or ""
        my_kwargs = dict(
            label = 'Shape-Tabelle einfügen',
            columns = ShapeTableGallery._columns,
            image = 'shapetable',
            # image_mso = 'SlidesPerPage4Slides',
            screentip="Shape-Tabelle einfügen",
            supertip="Füge eine Tabelle aus Standard-Shapes ein",
            description="Füge eine Tabelle aus Standard-Shapes ein",
            children=[
                bkt.ribbon.Button(id=parent_id + "_margin0", label="Ohne Abstand", supertip="Abstand bei Shape-Tabelle deaktivieren", on_action=bkt.Callback(lambda: setattr(self, "_margin", 0)), get_image=bkt.Callback(lambda: self.get_toggle_image(0))),
                bkt.ribbon.Button(id=parent_id + "_margin10", label="Kleiner Abstand", supertip="Abstand bei Shape-Tabelle auf klein setzen", on_action=bkt.Callback(lambda: setattr(self, "_margin", 10)), get_image=bkt.Callback(lambda: self.get_toggle_image(10))),
                bkt.ribbon.Button(id=parent_id + "_margin20", label="Großer Abstand", supertip="Abstand bei Shape-Tabelle auf groß setzen", on_action=bkt.Callback(lambda: setattr(self, "_margin", 20)), get_image=bkt.Callback(lambda: self.get_toggle_image(20))),
            ]
        )
        my_kwargs.update(kwargs)

        super(ShapeTableGallery, self).__init__(**my_kwargs)
    
    
    def on_action_indexed(self, selected_item, index, slide):
        ''' create numberd shape according of settings in clicked element '''
        n_rows, n_cols = self.get_rows_cols_from_index(index)
        self.create_shape_table(slide, n_rows, n_cols)
    
    
    def create_shape_table(self, slide, rows, columns):
        
        ref_left,ref_top,ref_width,ref_height = pplib.slide_content_size(slide)
        target_width = ref_width + self._margin
        target_height = ref_height + self._margin
        
        shape_width = target_width/columns
        shape_height = target_height/rows
        
        for r in range(rows):
            for c in range(columns):
                slide.shapes.AddShape(
                    1, #msoShapeRectangle
                    ref_left+c*shape_width, ref_top+r*shape_height,
                    shape_width-self._margin, shape_height-self._margin)
        
        shapes = pplib.last_n_shapes_on_slide(slide, rows*columns)
        shapes.select()
        
    
    def get_rows_cols_from_index(self, index):
        n_cols = index%self._columns
        n_rows = (index-n_cols)//self._columns + 1
        n_cols += 1
        return n_rows, n_cols
    
    def get_item_count(self):
        return self._rows * self._columns
        
    def get_item_label(self, index):
        n_rows, n_cols = self.get_rows_cols_from_index(index)
        return "%sx%s" % (n_cols, n_rows)
    
    def get_item_screentip(self, index):
        return "Shape-Tabelle einfügen"
        
    def get_item_supertip(self, index):
        n_rows, n_cols = self.get_rows_cols_from_index(index)
        return "Füge eine %sx%s-Tabelle aus Standard-Shapes ein (%s Spalten, %s Zeilen)" % (n_cols, n_rows, n_cols, n_rows)
    
    def get_item_image(self, index):
        ''' creates an item image with numberd shape according to settings in the specified item '''
        n_rows, n_cols = self.get_rows_cols_from_index(index)
        
        # create bitmap, define pen/brush
        size_w = 60 #16*3
        size_h = round(size_w/16*9) #9*3
        img = Drawing.Bitmap(size_w, size_h)
        g = Drawing.Graphics.FromImage(img)
        # color_black = Drawing.ColorTranslator.FromOle(0)
        #color_light_grey  = Drawing.ColorTranslator.FromOle(14540253)
        # color_grey  = Drawing.ColorTranslator.FromHtml('#666')
        color_grey  = Drawing.Brushes.Gray
        pen = Drawing.Pen(color_grey,1)
        #brush = Drawing.SolidBrush(color_black)
        
        #Draw smooth rectangle/ellipse
        g.SmoothingMode = Drawing.Drawing2D.SmoothingMode.AntiAlias
        
        #square
        #g.DrawRectangle(pen, Drawing.Rectangle(0,0, size-1, size-1))
        
        width = round(size_w/n_cols-1)
        height = round(size_h/n_rows-1)
        for r in range(n_rows):
            for c in range(n_cols):
                g.DrawRectangle(pen, Drawing.Rectangle(c*width,r*height, width, height))
        
        return img
    
    def get_toggle_image(self, margin):
        if self._margin == margin:
            return self.get_check_image()
        else:
            return self.get_check_image(checked=False)
    

class ChessTableGallery(ShapeTableGallery):
    
    def __init__(self, **kwargs):
        parent_id = kwargs.get('id') or ""
        my_kwargs = dict(
            label = 'Shape-Schachbrett einfügen',
            image = 'shapechessboard',
            screentip="Shape-Schachbrett einfügen",
            supertip="Füge ein Schachbrett aus Standard-Shapes ein",
            description="Füge ein Schachbrett aus Standard-Shapes ein",
        )
        my_kwargs.update(kwargs)
        super(ChessTableGallery, self).__init__(**my_kwargs)

        #overwrite attributes
        self._margin = 10
        #new attributes
        self._insert_textboxes = True
        self.children.append(
            bkt.ribbon.Button(id=parent_id + "_txtboxes", label="Textboxen in Zellen", supertip="Abstand bei Shape-Tabelle auf groß setzen", on_action=bkt.Callback(lambda: setattr(self, "_insert_textboxes", not self._insert_textboxes)), get_image=bkt.Callback(lambda: self.get_check_image(self._insert_textboxes)))
            )
    
    def create_shape_table(self, slide, rows, columns):
        
        ref_left,ref_top,ref_width,ref_height = pplib.slide_content_size(slide)
        target_width = ref_width
        target_height = ref_height
        
        shape_width = (target_width-self._margin)/columns
        shape_height = (target_height-self._margin)/rows
        
        for c in range(columns):
            shp = slide.shapes.AddShape(
                1, #msoShapeRectangle
                ref_left+self._margin+c*shape_width, ref_top,
                shape_width-self._margin, target_height)
            # shp.Fill.Transparency = 0.5
        
        for r in range(rows):
            shp = slide.shapes.AddShape(
                1, #msoShapeRectangle
                ref_left, ref_top+self._margin+r*shape_height,
                target_width, shape_height-self._margin)
            shp.Fill.Transparency = 0.5
        
        num_to_sel = rows+columns

        if self._insert_textboxes:
            for r in range(rows):
                for c in range(columns):
                    shpTxt = slide.shapes.AddTextbox(
                        1, #msoTextOrientationHorizontal
                        ref_left+self._margin+c*shape_width, ref_top+self._margin+r*shape_height,
                        shape_width-self._margin, shape_height-self._margin)
                    shpTxt.TextFrame2.AutoSize = 0 #ppAutoSizeNone
                    shpTxt.TextFrame2.WordWrap = -1 #msoTrue
                    shpTxt.TextFrame2.TextRange.Text = "tbd"
            num_to_sel += rows*columns

        shapes = pplib.last_n_shapes_on_slide(slide, num_to_sel)
        shapes.select()


shapes_table_menu = lambda: bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None, 
                children=[
                    bkt.ribbon.MenuSeparator(title="PowerPoint-Tabelle"),
                    bkt.mso.control.TableInsertGallery,
                    bkt.ribbon.MenuSeparator(title="Shape-Tabelle"),
                    ShapeTableGallery(id="insert_shape_table"),
                    ChessTableGallery(id="insert_shape_chessboard")
                ]
            )



class VisibilityToggleTags(pplib.BKTTag):
    TAG_NAME = "BKT_VISIBILITY_TOGGLE"

class ShapesMore(object):

    @staticmethod
    def show_invisible_shapes(context):
        toggle_shapes = list()
        slide = context.slide
        context.selection.Unselect()
        for shape in slide.shapes:
            if not shape.visible and not pplib.TagHelper.has_tag(shape, "THINKCELLSHAPEDONOTDELETE"):
                shape.visible = True
                shape.Select(replace=False)
                toggle_shapes.append(shape.id)
        return toggle_shapes

    @staticmethod
    def _hide_selected_shapes(context, shapes):
        toggle_shapes = list()
        for shape in shapes:
            shape.visible = False
            toggle_shapes.append(shape.id)
        return toggle_shapes
    
    @staticmethod
    def _hide_saved_shapes(context, shape_ids):
        toggle_shapes = list()
        slide = context.slide
        for shape in slide.shapes:
            if shape.id in shape_ids:
                shape.visible = False
                toggle_shapes.append(shape.id)
        return toggle_shapes

    @classmethod
    def toggle_shapes_visibility(cls, context):
        with VisibilityToggleTags(context.slide.Tags) as tags:

            sel_shapes = context.shapes
            if not sel_shapes:
                shapes_shown = cls.show_invisible_shapes(context)
                if shapes_shown:
                    tags["shape_ids"] = shapes_shown
                else:
                    try:
                        shape_ids = tags["shape_ids"]
                        cls._hide_saved_shapes(context, shape_ids)
                        del tags["shape_ids"]
                    except KeyError:
                        bkt.message.warning("Es sind keine Shapes zum Verstecken ausgewählt, es wurden keine vormals versteckten Shapes gefunden, und es gibt keine versteckten Shapes!")
            
            else:
                cls._hide_selected_shapes(context, sel_shapes)
                try:
                    del tags["shape_ids"]
                except KeyError:
                    pass


    # @staticmethod
    # def hide_shapes(shapes):
    #     for shape in shapes:
    #         shape.visible = False

    # @staticmethod
    # def show_shapes(slide):
    #     slide.Application.ActiveWindow.Selection.Unselect()
    #     for shape in slide.shapes:
    #         if not shape.visible and not pplib.TagHelper.has_tag(shape, "THINKCELLSHAPEDONOTDELETE"):
    #             shape.visible = True
    #             shape.Select(replace=False)
    
    @staticmethod
    def _text_to_shape(shape):
        try:
            return pplib.convert_text_into_shape(shape)
        except:
            logging.exception("Text to shape failed")
    
    @classmethod
    def texts_to_shapes(cls, shapes):
        if pplib.shape_is_group_child(shapes[0]) or any(shape.type == pplib.MsoShapeType["msoGroup"] for shape in shapes):
            bkt.message.error("PowerPoint unterstützt diese Funktion leider nicht für Gruppen.")
            return
        all_shapes = []
        for shape in shapes:
            all_shapes.append( cls._text_to_shape(shape) )
        if len(all_shapes)>0:
            pplib.shapes_to_range(all_shapes).select()


class PlaceholderConverter(object):
    @staticmethod
    def is_text_placeholder(shape):
        # return shape.Type == pplib.MsoShapeType["msoPlaceholder"] and shape.PlaceholderFormat.ContainedType in (pplib.MsoShapeType['msoTextBox'],pplib.MsoShapeType['msoAutoShape'] )
        return shape.Type == pplib.MsoShapeType["msoPlaceholder"]
    
    @classmethod
    def convert_placeholder(cls, shape):
        # new = pplib.replicate_shape(shape)
        new = shape.Duplicate()
        new.top, new.left = shape.top, shape.left
        shape.Delete()
        new.select(False)

    @classmethod
    def convert_shapes(cls, shapes):
        success=False
        for shape in shapes:
            if cls.is_text_placeholder(shape):
                try:
                    cls.convert_placeholder(shape)
                    success = True
                except:
                    logging.exception("placeholder conversion failed")

        if not success:
            bkt.message.warning("Aktuelle Auswahl enthält keine Platzhalter!")



shapes_change_menu = lambda: bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None, 
                children=[
                    bkt.ribbon.MenuSeparator(title="Formen manipulieren"),
                    bkt.ribbon.Button(
                        label="Shapes teilen/vervielfachen…",
                        image="split_horizontal",
                        screentip="Shapes teilen oder vervielfachen",
                        supertip="Shape horizontal/vertikal in mehrere Shapes teilen oder verfielfachen.",
                        on_action=bkt.Callback(ShapeDialogs.shape_split),
                        get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                    ),
                    bkt.ribbon.Button(
                        label="Shapes skalieren…",
                        image_mso="DiagramScale",
                        screentip="Shapes skalieren",
                        supertip="Shape-Größe inkl. aller Elemente/Eigenschaften (Schriftgröße, Konturen, etc.) gleichmäßig ändern.",
                        on_action=bkt.Callback(ShapeDialogs.shape_scale),
                        get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                    ),
                    bkt.ribbon.Button(
                        label="Platzhalter in Textbox umwandeln",
                        image_mso="ConvertTableToText",
                        supertip="Wandelt alle markierten Text-Platzhalter in echte Textboxen um, die u.A. eine Gruppierung erlauben.",
                        on_action=bkt.Callback(PlaceholderConverter.convert_shapes),
                        get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                    ),
                    bkt.mso.control.ObjectEditPoints,
                    bkt.ribbon.Button(
                        label="Text/Symbol zu Shapes umwandeln",
                        image_mso="TextEffectTransformGallery",
                        screentip="Texte bzw. Symbole werden in Standardshapes umgewandelt",
                        supertip="Ersetzt den Text einer Textbox in Shapes. Damit kann man bspw. einen Icon-Font in echte Icons umwandeln.",
                        on_action=bkt.Callback(ShapesMore.texts_to_shapes),
                        get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                    ),
                    bkt.ribbon.MenuSeparator(title="Formen zusammenführen"),
                    bkt.mso.control.ShapesUnion,
                    bkt.mso.control.ShapesCombine,
                    bkt.mso.control.ShapesFragment,
                    bkt.mso.control.ShapesIntersect,
                    bkt.mso.control.ShapesSubtract
                ]
            )


shapes_more_menu = lambda: bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None, 
                children=[
                    bkt.ribbon.MenuSeparator(title="Bilder und Objekte"),
                    bkt.mso.control.PictureInsertFromFilePowerPoint,
                    bkt.mso.control.OleObjectctInsert,
                    bkt.mso.control.ClipArtInsertDialog,
                    bkt.mso.control.SmartArtInsert,
                    bkt.mso.control.ChartInsert,
                    # bkt.mso.control.IconInsertFromFile, #only available in Office 2016 with 365 subscription
                    bkt.ribbon.MenuSeparator(title="Text & Beschriftungen"),
                    bkt.mso.control.HeaderFooterInsert,
                    bkt.mso.control.DateAndTimeInsert,
                    bkt.mso.control.NumberInsert,
                    bkt.mso.control.InsertNewComment,
                    bkt.ribbon.MenuSeparator(title="Ein-/Ausblenden"),
                    bkt.ribbon.Button(
                        id = 'toggle_shapes_visibility',
                        label = "Shapes vertecken/einblenden",
                        image_mso="ShapesSubtract",
                        supertip="Wenn Shapes ausgewählt sind, verstecke alle markierten Shapes (visible=False), anderen falls mache Shapes wieder sichtbar (visible=True). Gibt es keine unsichtbaren Shapes, werden die zuletzt versteckten Shapes erneut versteckt.",
                        on_action=bkt.Callback(ShapesMore.toggle_shapes_visibility),
                        # get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                    ),
                    # bkt.ribbon.Button(
                    #     id = 'hide_shape',
                    #     label = u"Shapes verstecken",
                    #     image_mso="ShapesSubtract",
                    #     supertip="Verstecke alle markierten Shapes (visible=False).",
                    #     on_action=bkt.Callback(ShapesMore.hide_shapes),
                    #     get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                    # ),
                    bkt.ribbon.Button(
                        id = 'show_shapes',
                        label = "Alle versteckten Shapes einblenden",
                        image_mso="VisibilityVisible",
                        supertip="Mache alle versteckten Shapes (visible=False) wieder sichtbar.",
                        on_action=bkt.Callback(ShapesMore.show_invisible_shapes)
                    ),
                ]
            )