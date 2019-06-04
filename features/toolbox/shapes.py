# -*- coding: utf-8 -*-
'''
Created on 06.07.2016

@author: rdebeerst
'''

import bkt
import bkt.library.powerpoint as pplib
from bkt.library.powerpoint import pt_to_cm, cm_to_pt

import logging

# other toolbox modules
from chartlib import shapelib_button
from agenda import ToolboxAgenda
import text
import harvey
import stateshapes

#import popups
import popups.traffic_light as traffic_light

#import System
from System import Guid, Array


from bkt import dotnet
Drawing = dotnet.import_drawing()
office = dotnet.import_officecore()





class PositionSize(object):

    @classmethod
    def set_top(cls, shapes, value):
        bkt.apply_delta_on_ALT_key(
            lambda shape, value: setattr(shape, 'top', value), 
            lambda shape: shape.top, 
            shapes, value)
    
    @classmethod
    def get_top(cls, shapes):
        return [shape.top for shape in shapes] #shapes[0].top
    
    
    @classmethod
    def set_left(cls, shapes, value):
        bkt.apply_delta_on_ALT_key(
            lambda shape, value: setattr(shape, 'left', value), 
            lambda shape: shape.left, 
            shapes, value)

    @classmethod
    def get_left(cls, shapes):
        return [shape.left for shape in shapes] #shapes[0].left


    @classmethod
    def set_height(cls, shapes, value):
        bkt.apply_delta_on_ALT_key(
            lambda shape, value: setattr(shape, 'height', value), 
            lambda shape: shape.height, 
            shapes, value)

    @staticmethod
    def get_height(shapes):
        return [shape.height for shape in shapes] #shapes[0].height
    
    
    @classmethod
    def set_width(cls, shapes, value):
        bkt.apply_delta_on_ALT_key(
            lambda shape, value: setattr(shape, 'width', value), 
            lambda shape: shape.width, 
            shapes, value)

    @staticmethod
    def get_width(shapes):
        return [shape.width for shape in shapes] #shapes[0].width


    @staticmethod
    def set_zorder(shapes, value):
        delta = int(value) - shapes[0].ZOrderPosition
        shapes = sorted(shapes, key=lambda shape: shape.ZOrderPosition, reverse=True if delta > 0 else False)
        for shape in shapes:
            pplib.set_shape_zorder(shape, delta=delta)
        # Normal behavior too confusing for users:
        # bkt.apply_delta_on_ALT_key(
        #     PositionSize._set_shape_zorder, 
        #     lambda shape: shape.ZOrderPosition, 
        #     shapes, int(value))

    @staticmethod
    def get_zorder(shapes):
        if len(shapes) == 1:
            return shapes[0].ZOrderPosition
        else:
            return [shapes[0].ZOrderPosition, None] #force ambiguous mode


    @staticmethod
    def shape_lock_aspect_ratio(shapes, pressed):
        for shape in shapes:
            shape.LockAspectRatio = -1 if pressed else 0


spinner_top = bkt.ribbon.RoundingSpinnerBox(
    id="pos_size_spinner_top",
    image_mso='ObjectNudgeDown',
    label="Position von oben",
    show_label=False,
    screentip="Position von oben",
    supertip="Änderung der Position von oben.\n\nBei gedrückter STRG-Taste Veränderung um 0,1 cm statt 0,2 cm.\n\nBei gedrückter ALT-Taste Veränderung relativ je Shape (wenn mehrere Shapes ausgewählt sind).",
    round_cm=True,
    on_change=bkt.Callback(PositionSize.set_top, shapes=True, wrap_shapes=True),
    get_text=bkt.Callback(PositionSize.get_top, shapes=True, wrap_shapes=True),
    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
    convert="pt_to_cm",
    image_element=pplib.LocpinGallery(image_mso='ObjectNudgeDown')
)

spinner_left = bkt.ribbon.RoundingSpinnerBox(
    id="pos_size_spinner_left",
    image_mso='ObjectNudgeRight',
    label="Position von links",
    show_label=False,
    screentip="Position von links",
    supertip="Änderung der Position von links.\n\nBei gedrückter STRG-Taste Veränderung um 0,1 cm statt 0,2 cm.\n\nBei gedrückter ALT-Taste Veränderung relativ je Shape (wenn mehrere Shapes ausgewählt sind).",
    round_cm=True,
    on_change=bkt.Callback(PositionSize.set_left, shapes=True, wrap_shapes=True),
    get_text=bkt.Callback(PositionSize.get_left, shapes=True, wrap_shapes=True),
    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
    convert="pt_to_cm",
    image_element=pplib.LocpinGallery(image_mso='ObjectNudgeRight')
)

spinner_height = bkt.ribbon.RoundingSpinnerBox(
    id="pos_size_spinner_height",
    image_mso='ShapeHeight',
    label="Höhe",
    show_label=False,
    screentip="Höhe",
    supertip="Änderung der Höhe.\n\nBei gedrückter STRG-Taste Veränderung um 0,1 cm statt 0,2 cm.\n\nBei gedrückter ALT-Taste Veränderung relativ je Shape (wenn mehrere Shapes ausgewählt sind).",
    round_cm=True,
    on_change=bkt.Callback(PositionSize.set_height, shapes=True, wrap_shapes=True),
    get_text=bkt.Callback(PositionSize.get_height, shapes=True, wrap_shapes=True),
    get_enabled=bkt.apps.ppt_shapes_or_text_selected,
    convert="pt_to_cm",
    image_element=pplib.LocpinGallery(image_mso='ShapeHeight')
)

spinner_width = bkt.ribbon.RoundingSpinnerBox(
    id="pos_size_spinner_width",
    image_mso='ShapeWidth',
    label="Breite",
    show_label=False,
    screentip="Breite",
    supertip="Änderung der Breite.\n\nBei gedrückter STRG-Taste Veränderung um 0,1 cm statt 0,2 cm.\n\nBei gedrückter ALT-Taste Veränderung relativ je Shape (wenn mehrere Shapes ausgewählt sind).",
    round_cm=True,
    on_change=bkt.Callback(PositionSize.set_width, shapes=True, wrap_shapes=True),
    get_text=bkt.Callback(PositionSize.get_width, shapes=True, wrap_shapes=True),
    get_enabled=bkt.apps.ppt_shapes_or_text_selected,
    convert="pt_to_cm",
    image_element=pplib.LocpinGallery(image_mso='ShapeWidth')
)

spinner_zorder = bkt.ribbon.RoundingSpinnerBox(
    id="pos_size_spinner_zorder",
    image_mso='ObjectBringForward',
    label="Z-Order",
    show_label=False,
    screentip="Z-Order",
    supertip="Änderung der Z-Order, also der Reihenfolge der Shapes auf der Folie.",
    on_change=bkt.Callback(PositionSize.set_zorder, shapes=True),
    get_text=bkt.Callback(PositionSize.get_zorder, shapes=True),
    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
    round_int=True,
    small_step=1,
    big_step=1,
    image_element=bkt.ribbon.Menu(
        children=[
            bkt.mso.control.ObjectBringToFront,
            bkt.mso.control.ObjectSendToBack,
        ],
    ),
)

#button_lock_aspect_ratio = bkt.ribbon.CheckBox(
button_lock_aspect_ratio = dict(
    #id = 'shape_lock_aspect_ratio',
    label="Seitenverhält.",
    screentip="Seitenverhältnis sperren",
    supertip="Wenn das Kontrollkästchen Seitenverhältnis sperren aktiviert ist, ändern sich die Einstellungen von Höhe und Breite im Verhältnis zueinander.",
    on_toggle_action = bkt.Callback(PositionSize.shape_lock_aspect_ratio, shapes=True),
    get_pressed = bkt.Callback(lambda shapes: shapes[0].LockAspectRatio == -1, shapes=True),
    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
)

size_group = bkt.ribbon.Group(
    id="bkt_size_group",
    label='Größe',
    image_mso='GroupSizeAndPosition',
    children =[
        #spinner_height,
        #spinner_width,
        bkt.mso.control.ShapeHeight(show_label=False),
        bkt.mso.control.ShapeWidth(show_label=False),
        bkt.ribbon.CheckBox(id="shape_lock_aspect_ratio1", **button_lock_aspect_ratio),
        bkt.ribbon.DialogBoxLauncher(idMso='ObjectSizeAndPositionDialog')
    ]
)

# pos_group = bkt.ribbon.Group(
#     label='Position',
#     image_mso='GroupSizeAndPosition',
#     children =[
#         spinner_top,
#         spinner_left,
#         spinner_zorder,
#         bkt.ribbon.DialogBoxLauncher(idMso='ObjectSizeAndPositionDialog')
#     ]
# )

pos_size_group = bkt.ribbon.Group(
    id="bkt_possize_group",
    label='Position/Größe',
    image_mso='GroupSizeAndPosition',
    children =[
        spinner_height,
        spinner_width,
        bkt.ribbon.CheckBox(id="shape_lock_aspect_ratio2", **button_lock_aspect_ratio),
        spinner_top,
        spinner_left,
        spinner_zorder,
        bkt.ribbon.DialogBoxLauncher(idMso='ObjectSizeAndPositionDialog')
    ]
)



class ShapesMore(object):

    @staticmethod
    def generateTracker(shapes, context):
        sld = context.app.ActivePresentation.Slides(context.app.ActiveWindow.View.Slide.SlideIndex)
        
        shapeCount = len(shapes)
        highlightShape = shapes[0]
        defaultShape = shapes[1]
        slideWidth = context.app.ActivePresentation.PageSetup.SlideWidth
        
        shpCounter = 0
        selShapes = Array.CreateInstance(str, shapeCount)
        for shp in shapes:
            selShapes[shpCounter] = shp.Name
            shpCounter += 1
        
        grp = sld.Shapes.Range(selShapes).Group()
        
        # duplicate group of shapes
        alterGrp = grp.duplicate()
        grp.ungroup()
        
        # format unselected elements
        for shp in alterGrp:
            shp.TextFrame.TextRange.Text = ""
            defaultShape.PickUp()
            shp.Apply()
        
        # generate unique GUID fpr tracker (items)
        tracker_guid = str(Guid.NewGuid())
        
        # format each selected element and paste tracker as image
        for curPosition in range(1, shapeCount+1):
            highlightShape.PickUp()
            curGrp = alterGrp.duplicate()
            curGrp.GroupItems(curPosition).Apply()
            curGrp.Copy()
            
            tracker = sld.shapes.PasteSpecial(DataType=6) # ppPastePNG = 6
            tracker.Tags.Add("tracker_id", tracker_guid)
            
            curGrp.delete()
            
            tracker.Height = cm_to_pt(1.5)
            tracker.left = slideWidth - cm_to_pt(3.0) + cm_to_pt(curPosition/1.5)
            tracker.top = cm_to_pt(3.0) + cm_to_pt(curPosition/1.5)
        
        alterGrp.delete()


    @staticmethod
    def alignTracker(shape, context):
        tracker_id = shape.Tags.Item("tracker_id")
        
        if tracker_id != "":            
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
    def distributeTracker(shapes, context):
        cur_slide_index = shapes[0].Parent.SlideIndex
        max_index = context.app.ActivePresentation.Slides.Count
        for shape in shapes[1:]:
            cur_slide_index = min(max_index, cur_slide_index+1)
            shape.Cut()
            context.app.ActivePresentation.Slides[cur_slide_index].Shapes.Paste()

        ShapesMore.alignTracker(shapes[0], context)

    

    @staticmethod
    def addHorizontalConnector(shapes, context):
        shapes = sorted(shapes, key=lambda shape: shape.Left)
        shpLeft  = shapes[0]
        shpRight = shapes[1]

        shpConnector = context.app.ActivePresentation.Slides(context.app.ActiveWindow.View.Slide.SlideIndex).shapes.AddShape(
            1, #msoShapeRectangle
            shpLeft.Left + shpLeft.Width, shpLeft.Top,
            shpRight.Left - shpLeft.Left - shpLeft.width, shpLeft.Height)
        # node 2: top right
        shpConnector.Nodes.SetPosition(2, shpRight.Left, shpRight.Top)
        # node 3: bottom right
        shpConnector.Nodes.SetPosition(3, shpRight.Left, shpRight.Top + shpRight.Height)
        shpConnector.Fill.ForeColor.RGB = 12566463 #193
        shpConnector.Line.ForeColor.RGB = 8355711 # 127 127 127
        shpConnector.Line.Weight = 0.75
        shpConnector.Select()

    @staticmethod
    def addVerticalConnector(shapes, context):
        shapes = sorted(shapes, key=lambda shape: shape.Top)
        shpTop = shapes[0]
        shpBottom = shapes[1]

        shpConnector = context.app.ActivePresentation.Slides(context.app.ActiveWindow.View.Slide.SlideIndex).shapes.AddShape(
            1, #msoShapeRectangle,
            shpTop.Left, shpTop.Top + shpTop.Height,
            shpTop.Width, shpBottom.Top - shpTop.Top - shpTop.Height)

        # node 3: bottom right
        shpConnector.Nodes.SetPosition(3, shpBottom.Left + shpBottom.width, shpBottom.Top)
        # node 4: bottom left
        shpConnector.Nodes.SetPosition(4, shpBottom.Left, shpBottom.Top)
        shpConnector.Fill.ForeColor.RGB = 12566463 # 193
        shpConnector.Line.ForeColor.RGB = 8355711 # 127 127 127
        shpConnector.Line.Weight = 0.75
        shpConnector.Select()
    
    @staticmethod
    def hide_shapes(shapes):
        for shape in shapes:
            shape.visible = False

    @staticmethod
    def show_shapes(slide):
        slide.Application.ActiveWindow.Selection.Unselect()
        for shape in slide.shapes:
            if not shape.visible:
                shape.visible = True
                shape.Select(replace=False)

    @staticmethod
    def paste_to_slides(slides):
        for slide in slides:
            slide.Shapes.Paste()
    
    @staticmethod
    def text_to_shape(shape):
        try:
            return pplib.convert_text_into_shape(shape)
        except Exception as e:
            logging.error("Text to shape failed with error {}".format(e))
    
    @classmethod
    def texts_to_shapes(cls, shapes):
        last_shape=None
        for shape in shapes:
            last_shape=cls.text_to_shape(shape)
        if last_shape:
            last_shape.select()



class Pentagon(bkt.FeatureContainer):
    
    @classmethod
    def create_headered_pentagon(cls, slide):
        ''' creates a headered pentagon on the given slide '''
        shapeCount = slide.shapes.count
        # shapes erstellen
        pentagon = slide.shapes.addshape( pplib.MsoAutoShapeType['msoShapePentagon'] , 100, 100, 400,200)
        header = slide.shapes.addshape( pplib.MsoAutoShapeType['msoShapeRectangle'], 100, 100, 400,30)

        pentagon.TextFrame.TextRange.Text = "Content"
        header.TextFrame.TextRange.Text = "Header"

        pentagon.Fill.ForeColor.ObjectThemeColor = pplib.MsoThemeColorIndex['msoThemeColorBackground1']
        header.Fill.ForeColor.ObjectThemeColor   = pplib.MsoThemeColorIndex['msoThemeColorText1']
        #header.Fill.ForeColor.ObjectThemeColor   = pplib.MsoThemeColorIndex['msoThemeColorBackground2']
        pentagon.TextFrame.TextRange.Font.Color.ObjectThemeColor = pplib.MsoThemeColorIndex['msoThemeColorText1']
        header.TextFrame.TextRange.Font.Color.ObjectThemeColor   = pplib.MsoThemeColorIndex['msoThemeColorBackground1']
        #header.TextFrame.TextRange.Font.Color.ObjectThemeColor   = pplib.MsoThemeColorIndex['msoThemeColorText2']

        pentagon.Line.ForeColor.ObjectThemeColor   = pplib.MsoThemeColorIndex['msoThemeColorText1']
        header.Line.ForeColor.ObjectThemeColor   = pplib.MsoThemeColorIndex['msoThemeColorText1']

        pentagon.Adjustments.item[1] = 0.2

        # margin top
        pentagon.textFrame.MarginTop = 36

        # align top/left
        pentagon.TextFrame.VerticalAnchor = 1 # Top
        pentagon.TextFrame.TextRange.ParagraphFormat.Alignment = 1 # Left
        header.TextFrame.TextRange.ParagraphFormat.Alignment = 1 # Left
        
        # gruppieren/selektieren
        grp = slide.Shapes.Range(Array[int]([shapeCount+1, shapeCount+2])).group()
        grp.select()

        #cls.update_pentagon_group(grp)
        cls.update_pentagon_header(pentagon, header)

    @classmethod
    def create_headered_chevron(cls, slide):
        ''' creates a headered pentagon on the given slide '''
        shapeCount = slide.shapes.count
        # shapes erstellen
        pentagon = slide.shapes.addshape( pplib.MsoAutoShapeType['msoShapeChevron'] , 100, 100, 400,200)
        header = slide.shapes.addshape( pplib.MsoAutoShapeType['msoShapeRectangle'], 100, 100, 400,30)

        pentagon.TextFrame.TextRange.Text = "Content"
        header.TextFrame.TextRange.Text = "Header"

        pentagon.Fill.ForeColor.ObjectThemeColor = pplib.MsoThemeColorIndex['msoThemeColorBackground1']
        header.Fill.ForeColor.ObjectThemeColor   = pplib.MsoThemeColorIndex['msoThemeColorText1']
        #header.Fill.ForeColor.ObjectThemeColor   = pplib.MsoThemeColorIndex['msoThemeColorBackground2']
        pentagon.TextFrame.TextRange.Font.Color.ObjectThemeColor = pplib.MsoThemeColorIndex['msoThemeColorText1']
        header.TextFrame.TextRange.Font.Color.ObjectThemeColor   = pplib.MsoThemeColorIndex['msoThemeColorBackground1']
        #header.TextFrame.TextRange.Font.Color.ObjectThemeColor   = pplib.MsoThemeColorIndex['msoThemeColorText2']

        pentagon.Line.ForeColor.ObjectThemeColor   = pplib.MsoThemeColorIndex['msoThemeColorText1']
        header.Line.ForeColor.ObjectThemeColor   = pplib.MsoThemeColorIndex['msoThemeColorText1']

        pentagon.Adjustments.item[1] = 0.2

        # margin top
        pentagon.textFrame.MarginTop = 36

        #margin left
        header.textFrame.MarginLeft = 16

        # align top/left
        pentagon.TextFrame.VerticalAnchor = 1 # Top
        pentagon.TextFrame.TextRange.ParagraphFormat.Alignment = 1 # Left
        header.TextFrame.TextRange.ParagraphFormat.Alignment = 1 # Left
        
        # gruppieren/selektieren
        grp = slide.Shapes.Range(Array[int]([shapeCount+1, shapeCount+2])).group()
        grp.select()

        #cls.update_pentagon_group(grp)
        cls.update_header(pentagon, header)
    
    
    @classmethod
    def update_pentagon_group(cls, shape):
        ''' updates the header of a group-shape (header + pentagon-body) '''
        body, header = cls.get_body_and_header_from_group(shape)
        if body:
            cls.update_header(body, header)
    
    
    @classmethod
    def update_header(cls, body, header):
        if body.AutoShapeType == pplib.MsoAutoShapeType['msoShapePentagon']:
            cls.update_pentagon_header(body, header)
        elif body.AutoShapeType == pplib.MsoAutoShapeType['msoShapeChevron']:
            cls.update_chevron_header(body, header)
    
    @classmethod
    def update_pentagon_header(cls, pentagon, header):
        ''' updates the header of the given pentagon '''
        offset = pentagon.Adjustments.item[1] * min(pentagon.width, pentagon.height)

        # header punkt links oben / links unten
        header.left = pentagon.left
        header.top = pentagon.top
        # header punkt rechts oben
        header.Nodes.SetPosition(2, pentagon.left + pentagon.width - offset, pentagon.top)
        # header punkt rechts unten
        header.Nodes.SetPosition(3, pentagon.left + pentagon.width - offset + ( header.height/(pentagon.height/2) * offset), pentagon.top + header.height)

    @classmethod
    def update_chevron_header(cls, chevron, header):
        ''' updates the header of the given pentagon '''
        cls.update_pentagon_header(chevron, header)
        
        # header punkt links unten
        offset = chevron.Adjustments.item[1] * min(chevron.width, chevron.height)
        header.Nodes.SetPosition(4, chevron.left + ( header.height/(chevron.height/2) * offset), chevron.top + header.height)
        
        

    @classmethod
    def is_headered_group(cls, shape):
        ''' returns true for group-shapes (header+body) '''
        pentagon, header = cls.get_body_and_header_from_group(shape)
        return pentagon != None

    @classmethod
    def is_header_shape(cls, shape):
        ''' returns true for header-shapes (Freeforms) '''
        return shape.Type == pplib.MsoShapeType['msoFreeform'] or (shape.Type == pplib.MsoShapeType['msoGraphic'] and shape.AutoShapeType == pplib.MsoAutoShapeType['msoShapeNotPrimitive'])
    
    @classmethod
    def is_body_shape(cls, shape):
        ''' returns true for body-shapes (Pentagon, Chevron, ...) '''
        return shape.AutoShapeType in [pplib.MsoAutoShapeType['msoShapePentagon'], pplib.MsoAutoShapeType['msoShapeChevron']]

    @classmethod
    def search_body_and_update_header(cls, shapes, shape):
        ''' for the pentagon represented by the given shape (header, body, or group header+body), the header position and size are updated '''
        header = shape
        body = cls.find_corresponding_body_shape(shapes, header)
        cls.update_header(body, header)
        

    @classmethod
    def find_corresponding_body_shape(cls, shapes, header):
        ''' given a shape-list and a header, the body-shape corresponding to the header in the list is returned
            the body shape is identified by its AutoShapeType
            if multiple possible body shapes are found, the body shape is choosen by its position,
            i.e. header-top-left-corner must lie inside the body shape
        '''
        possible_shapes = []
        # find body shapes
        for shape in shapes:
            if cls.is_body_shape(shape):
                possible_shapes.append(shape)
        # choose element
        if len(possible_shapes) == 0:
            return None
        elif len(possible_shapes) == 1:
            return possible_shapes[0]
        else:
            # choose element with smallest distance of top-left corners (roughly)
            distances = []
            for shape in possible_shapes:
                distances.append(abs(shape.top-header.top) + abs(shape.left-header.left))
            return possible_shapes[distances.index(min(distances))]
            


    @classmethod
    def get_body_and_header_from_group(cls, shape):
        ''' for a given group-shape (header + body-shape), the corresponding header and body are retured '''
        if not shape.Type == pplib.MsoShapeType['msoGroup']:
            return None, None
        if not shape.GroupItems.Count == 2:
            return None, None

        if cls.is_body_shape(shape.GroupItems(1)) and cls.is_header_shape(shape.GroupItems(2)):
            return shape.GroupItems(1), shape.GroupItems(2)
        elif cls.is_body_shape(shape.GroupItems(2)) and cls.is_header_shape(shape.GroupItems(1)):
            return shape.GroupItems(2), shape.GroupItems(1)
        else:
            return None, None



class ShapeDialogs(object):
    @staticmethod
    def show_segmented_circle_dialog(slide):
        from dialogs.circular_segments import SegmentedCircleWindow
        SegmentedCircleWindow.create_and_show_dialog(slide)

    @staticmethod
    def show_process_chevrons_dialog(slide):
        from dialogs.shape_process import ProcessWindow
        ProcessWindow.create_and_show_dialog(slide)




class NumberedShapes(object):
    
    label = "1"                 # 1->1,2,3   a->a,b,c   A->A,B,C   I->I,II,III
    shape_type = "square"       # square, circle
    style = "dark"              # dark, light
    position = "top-left"       # top-left, top-right
    position_offset = True      # True, False
    
    label_1 = [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26]
    label_a = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z']
    label_A = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
    label_I = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI', 'XII', 'XIII', 'XIV', 'XV', 'XVI', 'XVII', 'XVIII', 'XIX', 'XX', 'XXI', 'XXII', 'XXIII', 'XXIV', 'XXV', 'XXVI']
    
    
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
        
        number = 1
        num_shapes = []
        for shape in shapes:
            num_shape = cls.create_number_shape(slide, shape, number, **settings)
            num_shapes.append(num_shape)
            number += 1
        
        bkt.library.powerpoint.last_n_shapes_on_slide(slide, len(num_shapes)).select()
        
        
    
    @classmethod
    def create_number_shape(cls, slide, shape, number, label='1', shape_type='square', style='dark', position='top-left', position_offset=True):
        
        if shape_type == 'square':
            numshape = slide.shapes.addshape( pplib.MsoAutoShapeType['msoShapeRectangle'] , shape.left, shape.top, 14, 14)
        else: #circle
            numshape = slide.shapes.addshape( pplib.MsoAutoShapeType['msoShapeOval'] , shape.left, shape.top, 14, 14)
        
        if style == "dark":
            numshape.line.visible = False
            numshape.fill.ForeColor.RGB = 0
            numshape.TextFrame.TextRange.Font.Color.rgb = 255 + 255 * 256 + 255 * 256**2
            
        else: # light
            numshape.line.style = 1
            numshape.line.weight = 1
            numshape.line.ForeColor.RGB = 0
            numshape.fill.ForeColor.RGB = 255 + 255 * 256 + 255 * 256**2
            numshape.TextFrame.TextRange.Font.Color.rgb = 0
        
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
        numshape.TextFrame.TextRange.text = getattr(cls, 'label_' + label)[(number-1)%26] #at number 26 start from beginning to avoid IndexError
        numshape.TextFrame.TextRange.font.size = 12
        numshape.TextFrame.TextRange.ParagraphFormat.Alignment = pplib.PowerPoint.PpParagraphAlignment.ppAlignCenter.value__
        numshape.TextFrame.TextRange.ParagraphFormat.Bullet.Type = 0
        numshape.TextFrame.AutoSize = 0
        numshape.TextFrame.WordWrap = False
        numshape.TextFrame.MarginTop = 0
        numshape.TextFrame.MarginLeft = 0
        numshape.TextFrame.MarginRight = 0
        numshape.TextFrame.MarginBottom = 0
        #numshape.TextFrame.HorizontalAnchor = office.MsoHorizontalAnchor.msoAnchorCenter.value__
        numshape.TextFrame.VerticalAnchor = office.MsoVerticalAnchor.msoAnchorMiddle.value__
        
        return numshape
    
    
    
class NumberShapesGallery(bkt.ribbon.Gallery):
    
    # item-settings for gallery
    items = [ dict(label=l, style=s, shape_type=t) for l in ['1', 'a', 'A', 'I'] for t in ['circle', 'square'] for s in ['dark', 'light']  ]
    columns = 4
    
    position = "top-left"
    position_offset = True
    
    def __init__(self, **kwargs):
        parent_id = kwargs.get('id') or ""
        super(NumberShapesGallery, self).__init__(
            label = 'Nummerierung',
            columns = 4,
            screentip="Nummerierungs-Shapes einfügen",
            supertip="Fügt für jedes markierte Shape ein Nummerierungs-Shape ein. Nummerierung und Styling entsprechend der Auswahl. Markierte Shapes werden entsprechend der Selektions-Reihenfolge durchnummeriert.",
            get_image=bkt.Callback(lambda: self.get_item_image(0) ),
            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
            children=[
                bkt.ribbon.Button(id=parent_id + "_pos_left", label="Position links oben",    on_action=bkt.Callback(self.set_pos_top_left), get_image=bkt.Callback(lambda: self.get_toggle_image('pos-top-left')),
                    supertip="Nummerierungs-Shapes links oben auf dem zugehörigen Shape platzieren"),
                bkt.ribbon.Button(id=parent_id + "_pos-right", label="Position rechts oben",   on_action=bkt.Callback(self.set_pos_top_right), get_image=bkt.Callback(lambda: self.get_toggle_image('pos-top-right')),
                    supertip="Nummerierungs-Shapes rechts oben auf dem zugehörigen Shape platzieren"),
                bkt.ribbon.Button(id=parent_id + "_pos-offset", label="Versetzt positionieren", on_action=bkt.Callback(self.toggle_pos_offset), get_image=bkt.Callback(lambda: self.get_toggle_image('pos-offset')),
                    supertip="Standardmäßig werden Nummerierungs-Shapes genau am Rand des zugehörigen Shapes ausgerichtet.\n\nIst \"Versetzt positionieren\" aktiviert, werden die Nummerierungs-Shapes etwas weiter außerhalb des zugehörigen Shapes plaziert, so dass der Mittelpunkt des Nummerierungs-Shapes auf der Ecke liegt.")
            ],
            **kwargs
        )
    
    
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
        size = 30
        img = Drawing.Bitmap(size, size)
        g = Drawing.Graphics.FromImage(img)
        color_black = Drawing.ColorTranslator.FromOle(0)
        pen = Drawing.Pen(color_black,1)
        brush = Drawing.SolidBrush(color_black)
        
        #Draw smooth rectangle/ellipse
        g.SmoothingMode = Drawing.Drawing2D.SmoothingMode.AntiAlias
        
        if item['style'] == 'dark':
            # create black circle/rectangle
            brush = Drawing.SolidBrush(color_black)
            text_brush = Drawing.Brushes.White

            if item['shape_type'] == 'circle':
                g.FillEllipse(brush, 2,2, size-5, size-5)
            else: #square
                g.FillRectangle(brush, Drawing.Rectangle(2,2, size-5, size-5))

        else: # light
            # create white circle/rectangle
            text_brush = Drawing.Brushes.Black
            pen = Drawing.Pen(color_black,1)

            if item['shape_type'] == 'circle':
                g.DrawEllipse(pen, 2,1, size-4, size-4)
            else: #square
                g.DrawRectangle(pen, Drawing.Rectangle(2,2, size-4, size-4))

        # set string format
        strFormat = Drawing.StringFormat()
        strFormat.Alignment = Drawing.StringAlignment.Center
        strFormat.LineAlignment = Drawing.StringAlignment.Center
        
        # draw string
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAliasGridFit
        # g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAlias
        g.DrawString(str(getattr(NumberedShapes, 'label_' + item['label'])[index%int(self.columns)]),
                     Drawing.Font("Arial", 18, Drawing.FontStyle.Bold, Drawing.GraphicsUnit.Pixel), text_brush, 
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
            return None

    def get_check_image(self):
        size = 16
        img = Drawing.Bitmap(size, size)
        g = Drawing.Graphics.FromImage(img)
        
        text_brush = Drawing.Brushes.Black
        strFormat = Drawing.StringFormat()
        strFormat.Alignment = Drawing.StringAlignment.Center
        strFormat.LineAlignment = Drawing.StringAlignment.Center
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAlias
        g.DrawString('',
                     Drawing.Font("Wingdings", 14, Drawing.GraphicsUnit.Pixel), text_brush,
                     Drawing.RectangleF(2, 3, size, size),
                     strFormat)
        return img



class SplitShapes(object):
    default_row_sep = cm_to_pt(0.2)
    default_col_sep = cm_to_pt(0.2)
    default_rows = 6
    default_cols = 6
    
    @classmethod
    def split_shapes(cls, shapes, rows, cols, row_sep, col_sep):
        for shape in shapes:
            cls.split_shape(shape, rows, cols, row_sep, col_sep)
    
    @classmethod
    def split_shape(cls, shape, rows, cols, row_sep, col_sep):
        shape_width = (shape.width - (cols-1)*col_sep)/cols
        shape_height = (shape.height - (rows-1)*row_sep)/rows
        
        #shape.width = shape_width
        #shape.height = shape_height
        
        for row_idx in range(rows):
            for col_idx in range(cols):
                if row_idx == 0 and col_idx == 0:
                    shape_copy = shape
                else:
                    shape_copy = shape.duplicate()
                shape_copy.left = shape.left + col_idx*(shape_width+col_sep)
                shape_copy.top = shape.top + row_idx*(shape_height+row_sep)
                shape_copy.width = shape_width
                shape_copy.height = shape_height
                shape_copy.select(False)
        #shape.Delete()



#FIXME: no dependency to circular wanted here
#from circular import CircularArrangement


class MultiplyShapes(object):
    
    @classmethod
    def multiply_shapes(cls, shapes, rows, cols, row_sep, col_sep):
        for shape in shapes:
            cls.multiply_shape(shape, rows, cols, row_sep, col_sep)
    
    @classmethod
    def multiply_shape(cls, shape, rows, cols, row_sep, col_sep):
        shape_width = shape.width
        shape_height = shape.height
        
        for row_idx in range(rows):
            for col_idx in range(cols):
                if row_idx == 0 and col_idx == 0:
                    continue
                shape_copy = shape.duplicate()
                shape_copy.left = shape.left + col_idx*(shape_width+col_sep)
                shape_copy.top = shape.top + row_idx*(shape_height+row_sep)
                shape_copy.width = shape_width
                shape_copy.height = shape_height
                shape_copy.select(False)
    
    # @classmethod
    # def multiply_shapes_cyclic(cls, shapes, number, sep):
    #     for shape in shapes:
    #         cls.multiply_shaps_cyclic(shape, number, sep)
    #
    # @classmethod
    # def multiply_shaps_cyclic(cls, shape, number, sep):
    #     shapes = [shape]
    #
    #     # distance circle-midpoint to shape-midpoint is given by sep
    #     height = 2*(max(shape.height, shape.width)/2 + sep)
    #     width = height
    #
    #     midpoint = [ shape.left + shape.width/2 , shape.top + shape.height/2 + height/2 ]
    #
    #     # generate shapes and arrange
    #     for row_idx in range(number-1):
    #         shape_copy = shape.duplicate()
    #         shape_copy.select(False)
    #         shapes.append(shape_copy)
    #     CircularArrangement.arrange_circular_wargs(shapes, midpoint, width, height)
        


split_shapes_group = bkt.ribbon.Group(
    id="bkt_splitshapes_group",
    label="Teilen/Vervielfachen",
    image_mso='TableRowsDistribute',
    children=[
        #bkt.ribbon.Label(label="Zeilen"),
        bkt.ribbon.Box(
            box_style="horizontal",
            children = [
                bkt.ribbon.Button(
                    id='shape_split_horizontal',
                    label="Horizontal teilen",
                    show_label=False,
                    image="split_horizontal",
                    screentip="Horizontal teilen",
                    supertip="Shape horizontal in mehrere Shapes teilen, entsprechend der angegebenen Anzahl und mit angegebenem Abstand zwischen den Shapes.",
                    on_action = bkt.Callback(lambda shapes: SplitShapes.split_shapes(shapes, SplitShapes.default_rows, 1, SplitShapes.default_row_sep, 0 )),
                    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                ),
                bkt.ribbon.Button(
                    id='shape_split_vertical',
                    label="Vertikal teilen",
                    show_label=False,
                    image="split_vertical",
                    screentip="Vertikal teilen",
                    supertip="Shape vertikal in mehrere Shapes teilen, entsprechend der angegebenen Anzahl und mit angegebenem Abstand zwischen den Shapes.",
                    on_action = bkt.Callback(lambda shapes: SplitShapes.split_shapes(shapes, 1, SplitShapes.default_cols, 0, SplitShapes.default_col_sep )),
                    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                ),
                bkt.ribbon.Button(
                    id='shape_mult_vertical',
                    label="Vertikal vervielfachen",
                    show_label=False,
                    image="multiply_vertical",
                    screentip="Vertikal vervielfachen",
                    supertip="Shape mehrfach dublizieren, entsprechend der angegebenen Anzahl. Shapes werden untereinander angeordnet mit dem angegebenem Abstand zwischen den Shapes.",
                    on_action = bkt.Callback(lambda shapes: MultiplyShapes.multiply_shapes(shapes, SplitShapes.default_rows, 1, SplitShapes.default_row_sep, 0 )),
                    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                ),
                bkt.ribbon.Button(
                    id='shape_mult_horizontal',
                    label="Horizontal vervielfachen",
                    show_label=False,
                    image="multiply_horizontal",
                    screentip="Horizontal vervielfachen",
                    supertip="Shape mehrfach dublizieren, entsprechend der angegebenen Anzahl. Shapes werden nebeneinander angeordnet mit dem angegebenem Abstand zwischen den Shapes.",
                    on_action = bkt.Callback(lambda shapes: MultiplyShapes.multiply_shapes(shapes, 1, SplitShapes.default_cols, 0, SplitShapes.default_col_sep )),
                    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                ),
            ]
        ),
        bkt.ribbon.RoundingSpinnerBox(
            id = 'shape_slit_rows',
            label=u"Anzahl Zeilen/Spalten",
            supertip="Angestrebte Shapeanzahl für das Teilen/Vervielfachen von Shapes.",
            show_label=False,
            imageMso="TableRowsDistribute",
            #TableRowsDistribute, TableStyleBandedRowsWord, TableRowsSelect
            on_change = bkt.Callback(lambda value: [setattr(SplitShapes, 'default_rows', max(0, int(value))), setattr(SplitShapes, 'default_cols', max(0, int(value)))]),
            get_text  = bkt.Callback(lambda: SplitShapes.default_rows),
            big_step = 1,
            small_step = 1,
            round_at = 0
        ),
        bkt.ribbon.RoundingSpinnerBox(
            id = 'shape_slit_row_sep',
            label=u"Zeilen-/Spaltenabstand",
            supertip="Abstand zwischen Shapes zur Berücksichtigung beim Teilen/Vervielfachen von Shapes.\n\nBei Kreisanordnung wird hiermit der vertikale/horizontale Abstand zum Mittelpunkt angegeben.",
            show_label=False,
            image_mso="RowHeight",
            on_change = bkt.Callback(lambda value: [setattr(SplitShapes, 'default_row_sep', cm_to_pt(value)), setattr(SplitShapes, 'default_col_sep', cm_to_pt(value))]),
            get_text  = bkt.Callback(lambda: round(pt_to_cm(SplitShapes.default_row_sep),2)),
            round_cm = True
        ),
        # bkt.ribbon.Button(
        #     id='shape_mult_circular',
        #     #label="mult.",
        #     image="multiply_circular",
        #     screentip="Shape kreisförmig vervielfachen",
        #     supertip="Shape vervielfachen und kreisförmig anordnen.",
        #     on_action = bkt.Callback(lambda shapes: MultiplyShapes.multiply_shapes_cyclic(shapes, SplitShapes.default_rows, SplitShapes.default_row_sep )),
        #     get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
        # )
    ]
)




# Context menu if multiple connectors are selected
class CtxVerbinder(object):
    @staticmethod
    def ctx_connectors_reroute_enabled(shapes):
        return all(shape.Connector == -1 and shape.ConnectorFormat.BeginConnected == -1 and shape.ConnectorFormat.EndConnected == -1 for shape in shapes)

    @staticmethod
    def ctx_connectors_visible(shapes):
        return all(shape.Connector == -1 for shape in shapes)

    @staticmethod
    def set_connector_type(shapes, con_type):
        for shape in shapes:
            if shape.Connector == -1: #msoTrue
                shape.ConnectorFormat.Type = con_type

    @staticmethod
    def reroute_connectors(shapes):
        for shape in shapes:
            if shape.Connector == -1 and shape.ConnectorFormat.BeginConnected == -1 and shape.ConnectorFormat.EndConnected == -1: #msoTrue
                shape.RerouteConnections()

    @staticmethod
    def invert_direction(shapes):
        for shape in shapes:
            if shape.Connector == -1: #msoTrue
                #Swap begin and end styles
                shape.Line.BeginArrowheadLength, shape.Line.EndArrowheadLength = shape.Line.EndArrowheadLength, shape.Line.BeginArrowheadLength
                shape.Line.BeginArrowheadStyle, shape.Line.EndArrowheadStyle = shape.Line.EndArrowheadStyle, shape.Line.BeginArrowheadStyle
                shape.Line.BeginArrowheadWidth, shape.Line.EndArrowheadWidth = shape.Line.EndArrowheadWidth, shape.Line.BeginArrowheadWidth
    



class ShapeFormats(object):
    transparencies = range(0, 110, 10)

    @classmethod
    def _attr_setter(cls, shape, value, shp_object, attribute):
        try:
            if attribute == "Transparency":
                value = min(max(0, value/100),100)
            else:
                value = max(0, value)
            shp_object = getattr(shape, shp_object)
            setattr(shp_object, "visible", -1)
            setattr(shp_object, attribute, value)
        except:
            logging.debug("Setting {} attribute {} to value {} failed!".format(shp_object, attribute, value))
    @classmethod
    def _attr_getter(cls, shape, shp_object, attribute):
        try:
            shp_object = getattr(shape, shp_object)
            value = max(0, getattr(shp_object, attribute))
            if attribute == "Transparency":
                value = value*100
            return value
        except:
            logging.debug("Getting {} attribute {} failed!".format(shp_object, attribute))
            return 0

    ### Fill properties ###
    @classmethod
    def get_fill_enabled(cls, context):
        #TESTME: is fill implemented for all shape types? (see also problem with line)
        # shape = next(pplib.iterate_shape_subshapes(shapes))
        # return shape.Fill.visible == -1
        
        # copy enabled status of fill-button
        return context.app.commandbars.GetEnabledMso("ShapeFillColorPicker")

    @classmethod
    def get_fill_transparency(cls, shapes):
        shapes = pplib.iterate_shape_subshapes(shapes)
        for shape in shapes:
            try:
                return max(0, round(shape.fill.transparency*100))
            except:
                continue
        return None
    
    @classmethod
    def set_fill_transparency(cls, shapes, value):
        value = min(max(0, value),100) #min=0, max=100
        shapes = list(pplib.iterate_shape_subshapes(shapes))
        bkt.apply_delta_on_ALT_key(
            # lambda shape, value: setattr(shape.Fill, 'Transparency', min(max(0, value/100),100)), 
            cls._attr_setter,
            cls._attr_getter,
            shapes, value, shp_object="Fill", attribute="Transparency")

    ### Line properties ###
    @classmethod
    def get_line_enabled(cls, context):
        # return len(cls._line_filter(shapes)) > 0
        # shape = next(pplib.iterate_shape_subshapes(shapes))
        # try:
        #     return hasattr(shape.line, "visible")
        # except ValueError:
        #     return False

        # copy enabled status of line-button
        return context.app.commandbars.GetEnabledMso("ShapeOutlineColorPicker")

    @classmethod
    def get_line_transparency(cls, shapes):
        shapes = pplib.iterate_shape_subshapes(shapes)
        for shape in shapes:
            try:
                return max(0, round(shape.line.transparency*100))
            except:
                continue
        return None
    
    @classmethod
    def set_line_transparency(cls, shapes, value):
        value = min(max(0, value),100) #min=0, max=100
        shapes = pplib.iterate_shape_subshapes(shapes)
        bkt.apply_delta_on_ALT_key(
            # lambda shape, value: setattr(shape.Line, 'Transparency', min(max(0, value/100),100)), 
            cls._attr_setter,
            cls._attr_getter,
            shapes, value, shp_object="Line", attribute="Transparency")

    @classmethod
    def get_line_weight(cls, shapes):
        shapes = pplib.iterate_shape_subshapes(shapes)
        for shape in shapes:
            try:
                return max(0, shape.line.weight)
            except:
                continue
        return None
    
    @classmethod
    def set_line_weight(cls, shapes, value):
        value = max(0, value)
        shapes = list(pplib.iterate_shape_subshapes(shapes))
        bkt.apply_delta_on_ALT_key(
            # lambda shape, value: setattr(shape.Line, 'weight', max(0, value)), 
            cls._attr_setter,
            cls._attr_getter,
            shapes, value, shp_object="Line", attribute="weight")

    ### GALLERY ###
    @classmethod
    def get_item_count(cls):
        return len(cls.transparencies)

    @classmethod
    def fill_on_action_indexed(cls, selected_item, index, shapes):
        value = float(cls.transparencies[index])
        cls.set_fill_transparency(shapes, value)
    
    @classmethod
    def fill_get_selected_item_index(cls, context):
        try:
            return cls.transparencies.index(cls.get_fill_transparency(context.shapes))
        except:
            return -1

    @classmethod
    def line_on_action_indexed(cls, selected_item, index, shapes):
        value = float(cls.transparencies[index])
        cls.set_line_transparency(shapes, value)
    
    @classmethod
    def line_get_selected_item_index(cls, context):
        try:
            return cls.transparencies.index(cls.get_line_transparency(context.shapes))
        except:
            return -1


format_group = bkt.ribbon.Group(
    id="bkt_format_group",
    label="Format",
    image_mso='BehindText',
    children=[
        bkt.ribbon.RoundingSpinnerBox(
            id = 'fill_transparency',
            label=u"Transparenz Hintergrund",
            show_label=False,
            round_int = True,
            image="fill_transparency",
            on_change = bkt.Callback(ShapeFormats.set_fill_transparency, shapes=True),
            get_text  = bkt.Callback(ShapeFormats.get_fill_transparency, shapes=True),
            get_enabled = bkt.Callback(ShapeFormats.get_fill_enabled),
        ),
        bkt.ribbon.RoundingSpinnerBox(
            id = 'line_transparency',
            label=u"Transparenz Linie/Rahmen",
            show_label=False,
            round_int = True,
            image="line_transparency",
            on_change = bkt.Callback(ShapeFormats.set_line_transparency, shapes=True),
            get_text  = bkt.Callback(ShapeFormats.get_line_transparency, shapes=True),
            get_enabled = bkt.Callback(ShapeFormats.get_line_enabled),
        ),
        bkt.ribbon.RoundingSpinnerBox(
            id = 'line_weight',
            label=u"Dicke Linie/Rahmen",
            show_label=False,
            round_pt = True,
            rounding_factor=0.25,
            huge_step=1,
            big_step=0.5,
            small_step=0.25,
            image_mso="LineThickness",
            on_change = bkt.Callback(ShapeFormats.set_line_weight, shapes=True),
            get_text  = bkt.Callback(ShapeFormats.get_line_weight, shapes=True),
            get_enabled = bkt.Callback(ShapeFormats.get_line_enabled),
        ),
        bkt.ribbon.DialogBoxLauncher(idMso='ObjectFormatDialog')
    ]
)

fill_transparency_gallery = bkt.ribbon.Gallery(
    id="bkt_fill_transparency_menu",
    label="Transparenz Hintergrund",
    supertip="Setzt die Hintergrund-Transparenz auf den gewählten Wert.",
    show_label=False,
    show_item_label=True,
    image="fill_transparency",
    columns="1",
    get_enabled = bkt.Callback(ShapeFormats.get_fill_enabled),
    on_action_indexed = bkt.Callback(ShapeFormats.fill_on_action_indexed, shapes=True),
    get_selected_item_index = bkt.Callback(ShapeFormats.fill_get_selected_item_index, context=True),
    children=[
        bkt.ribbon.Item(label="%s%%" % transp, image="transp_%s" % transp)
        for transp in ShapeFormats.transparencies
    ]
)

line_transparency_gallery = bkt.ribbon.Gallery(
    id="bkt_line_transparency_menu",
    label="Transparenz Linie/Rahmen",
    supertip="Setzt die Linien-Transparenz auf den gewählten Wert.",
    show_label=False,
    show_item_label=True,
    image="line_transparency",
    columns="1",
    get_enabled = bkt.Callback(ShapeFormats.get_line_enabled),
    on_action_indexed = bkt.Callback(ShapeFormats.line_on_action_indexed, shapes=True),
    get_selected_item_index = bkt.Callback(ShapeFormats.line_get_selected_item_index, context=True),
    children=[
        bkt.ribbon.Item(label="%s%%" % transp, image="transp_%s" % transp)
        for transp in ShapeFormats.transparencies
    ]
)


class PictureFormat(object):
    @staticmethod
    def make_img_transparent(slide, shapes, transparency=0.5):
        if not bkt.helpers.confirmation("Das bestehende Bild wird dabei ersetzt. Fortfahren?"):
            return

        import tempfile, os
        filename = tempfile.gettempdir() + "\\bktimgtransp.png"

        for shape in shapes:
            if shape.Type != pplib.MsoShapeType["msoPicture"]:
                continue

            shape.Export(filename, 2) #2=ppShapeFormatPNG

            pic_shape = slide.Shapes.AddShape(
                shape.AutoShapeType,
                shape.Left, shape.Top,
                shape.Width, shape.Height
                )
            pic_shape.LockAspectRatio = -1
            pic_shape.Rotation = shape.Rotation
            pplib.set_shape_zorder(pic_shape, value=shape.ZOrderPosition)
            shape.PickUp()
            pic_shape.Apply()
            pic_shape.line.visible = shape.line.visible # line is not properly transferred by pickup-apply

            pic_shape.fill.UserPicture(filename)
            pic_shape.fill.transparency = transparency
            pic_shape.Select(replace=False)

            shape.Delete()
            os.remove(filename)


class ShapeDialogActions(object):
    
    @staticmethod
    def shape_split(shapes):
        from dialogs.shape_split import ShapeSplitWindow
        ShapeSplitWindow.create_and_show_dialog(shapes)



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
                bkt.ribbon.Button(id=parent_id + "_margin0", label="Ohne Abstand", on_action=bkt.Callback(lambda: setattr(self, "_margin", 0)), get_image=bkt.Callback(lambda: self.get_toggle_image(0))),
                bkt.ribbon.Button(id=parent_id + "_margin10", label="Kleiner Abstand", on_action=bkt.Callback(lambda: setattr(self, "_margin", 10)), get_image=bkt.Callback(lambda: self.get_toggle_image(10))),
                bkt.ribbon.Button(id=parent_id + "_margin20", label="Großer Abstand", on_action=bkt.Callback(lambda: setattr(self, "_margin", 20)), get_image=bkt.Callback(lambda: self.get_toggle_image(20))),
            ]
        )
        my_kwargs.update(kwargs)

        super(ShapeTableGallery, self).__init__(**my_kwargs)
    
    
    def on_action_indexed(self, selected_item, index, slide):
        ''' create numberd shape according of settings in clicked element '''
        n_rows, n_cols = self.get_rows_cols_from_index(index)
        self.create_shape_table(slide, n_rows, n_cols)
    
    
    def create_shape_table(self, slide, rows, columns):
        
        ref_left,ref_top,ref_width,ref_height = pplib.slide_content_size(slide.parent)
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
        n_rows = (index-n_cols)/self._columns + 1
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
        size_h = size_w/16*9 #9*3
        img = Drawing.Bitmap(size_w, size_h)
        g = Drawing.Graphics.FromImage(img)
        # color_black = Drawing.ColorTranslator.FromOle(0)
        #color_light_grey  = Drawing.ColorTranslator.FromOle(14540253)
        color_grey  = Drawing.ColorTranslator.FromHtml('#666')
        pen = Drawing.Pen(color_grey,1)
        #brush = Drawing.SolidBrush(color_black)
        
        #Draw smooth rectangle/ellipse
        g.SmoothingMode = Drawing.Drawing2D.SmoothingMode.AntiAlias
        
        #square
        #g.DrawRectangle(pen, Drawing.Rectangle(0,0, size-1, size-1))
        
        width = size_w/n_cols-1
        height = size_h/n_rows-1
        for r in range(n_rows):
            for c in range(n_cols):
                g.DrawRectangle(pen, Drawing.Rectangle(0+c*width,0+r*height, width, height))
        
        return img
    
    def get_toggle_image(self, margin):
        if self._margin == margin:
            return self.get_check_image()
        else:
            return None

    def get_check_image(self):
        size = 16
        img = Drawing.Bitmap(size, size)
        g = Drawing.Graphics.FromImage(img)
        
        text_brush = Drawing.Brushes.Black
        strFormat = Drawing.StringFormat()
        strFormat.Alignment = Drawing.StringAlignment.Center
        strFormat.LineAlignment = Drawing.StringAlignment.Center
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAlias
        g.DrawString('',
                     Drawing.Font("Wingdings", 14, Drawing.GraphicsUnit.Pixel), text_brush,
                     Drawing.RectangleF(2, 3, size, size),
                     strFormat)
        return img
    

class ChessTableGallery(ShapeTableGallery):
    
    def __init__(self, **kwargs):
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
    
    def create_shape_table(self, slide, rows, columns):
        
        ref_left,ref_top,ref_width,ref_height = pplib.slide_content_size(slide.parent)
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

        shapes = pplib.last_n_shapes_on_slide(slide, rows+columns)
        shapes.select()



picture_format_tab = bkt.ribbon.Tab(
    idMso = "TabPictureToolsFormat",
    children = [
        bkt.ribbon.Group(
            id="bkt_pictureformat_group",
            label="Format",
            insert_after_mso="GroupPictureTools",
            children = [
                bkt.ribbon.Button(
                    id = 'make_img_transparent',
                    label="Transparenz ermöglichen",
                    supertip="Ersetzt das Bild durch ein Shape mit Bildfüllung, welches nativ transparent gemacht werden kann. Dabei wird das bestehende Bild exportiert und dann gelöscht, d.h. etwaige zugeschnittene Bereiche gehen verloren und Bildformate können nicht rückgängig gemacht werden.",
                    size="large",
                    show_label=True,
                    image_mso='PictureRecolorWashout',
                    on_action=bkt.Callback(PictureFormat.make_img_transparent),
                    # get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                ),
            ]
        )
    ]
)







shapes_group = bkt.ribbon.Group(
    id="bkt_shapes_group",
    label='Formen',
    image_mso='ShapesInsertGallery',
    children = [
        bkt.mso.control.ShapesInsertGallery,
        text.text_splitbutton,
        bkt.ribbon.Menu(
            image_mso='TableInsertGallery',
            screentip="Tabelle einfügen",
            supertip="Einfügen von Standard- oder Shape-Tabellen",
            item_size="large",
            children=[
                bkt.ribbon.MenuSeparator(title="PowerPoint-Tabelle"),
                bkt.mso.control.TableInsertGallery,
                bkt.ribbon.MenuSeparator(title="Shape-Tabelle"),
                ShapeTableGallery(id="inesrt_shape_table"),
                ChessTableGallery(id="inesrt_shape_chessboard")
            ]
        ),
        
        #bkt.mso.control.PictureInsertFromFilePowerPoint,
        shapelib_button,
        text.symbol_insert_splitbutton,
        bkt.ribbon.Menu(
            label='Spezialformen',
            show_label=False,
            image_mso='SmartArtInsert',
            screentip="Spezielle und Interaktive Formen ",
            supertip="Bilder, Objekte, Spezial-Shapes, Text-Zerlegung, Shapes verstecken, ...",
            children = [
                bkt.ribbon.MenuSeparator(title="Einfügehilfen"),
                bkt.ribbon.Button(
                    id = 'standard_process',
                    label = u"Prozesspfeile…",
                    image = "process_chevrons",
                    screentip="Prozess-Pfeile einfügen",
                    supertip="Erstelle Standard Prozess-Pfeile.",
                    on_action=bkt.Callback(ShapeDialogs.show_process_chevrons_dialog)
                ),
                bkt.ribbon.Button(
                    id = 'segmented_circle',
                    label = u"Kreissegmente…",
                    image = "segmented_circle",
                    screentip="Kreissegmente einfügen",
                    supertip="Erstelle Kreis mit Segmenten oder Chevrons.",
                    on_action=bkt.Callback(ShapeDialogs.show_segmented_circle_dialog)
                ),
                bkt.ribbon.Button(
                    id='agenda_textbox',
                    label="Agenda-Textbox einfügen",
                    screentip="Standard Agenda-Textbox einfügen.",
                    imageMso="TextBoxInsert",
                    on_action=bkt.Callback(ToolboxAgenda.create_agenda_textbox_on_slide)
                ),
                NumberShapesGallery(id='number-labels-gallery'),
                bkt.ribbon.Menu(
                    label='Tracker',
                    image = "Tracker",
                    screentip="Tracker erstellen oder ausrichten",
                    supertip="Einen Tracker aus einer Auswahl erstellen, verteilen und ausrichten.",
                    children = [
                        bkt.ribbon.Button(
                            id = 'tracker',
                            label = u"Tracker aus Auswahl erstellen",
                            #image = "Tracker",
                            screentip="Tracker aus Auswahl erstellen",
                            supertip="Erstelle aus den markierten Shapes einen Tracker.\nDer Shape-Stil für Highlights wird aus dem zuerst markierten Shape (in der Regel oben links) bestimmt. Der Shape-Stil für alle anderen Shapes wird aus dem als zweites markierten Shape bestimmt.",
                            on_action=bkt.Callback(ShapesMore.generateTracker, shapes=True, shapes_min=2, context=True),
                            get_enabled = bkt.apps.ppt_shapes_min2_selected,
                        ),
                        bkt.ribbon.Button(
                            id = 'tracker_distribute',
                            label = u"Tracker auf Folien verteilen",
                            #image = "Tracker",
                            screentip="Alle Tracker verteilen",
                            supertip="Verteilen der ausgewählten Tracker auf die Folgefolien und ausrichten.",
                            on_action=bkt.Callback(ShapesMore.distributeTracker, shapes=True, shapes_min=2, context=True),
                            get_enabled = bkt.apps.ppt_shapes_min2_selected,
                        ),
                        bkt.ribbon.Button(
                            id = 'tracker_align',
                            label = u"Alle Tracker ausrichten",
                            #image = "Tracker",
                            screentip="Alle Tracker ausrichten",
                            supertip="Ausrichten (Position, Größe, Rotation) aller Tracker (auf allen Folien) anhand des ausgewählten Tracker.",
                            on_action=bkt.Callback(ShapesMore.alignTracker, shape=True, context=True),
                            get_enabled = bkt.apps.ppt_shapes_exactly1_selected,
                        ),
                    ]
                ),
                bkt.ribbon.MenuSeparator(title="Interaktive Formen"),
                bkt.ribbon.Button(
                    id = 'headered_pentagon',
                    label = u"Prozessschritt mit Kopfzeile",
                    image = "headered_pentagon",
                    screentip="Prozess-Schritt-Shape mit Kopfzeile erstellen",
                    supertip="Erstelle einen Prozess-Pfeil mit Header-Shape. Das Header-Shape kann im Prozess-Pfeil über Kontext-Menü des Header-Shapes passend angeordnet werden.",
                    on_action=bkt.Callback(Pentagon.create_headered_pentagon)
                ),
                bkt.ribbon.Button(
                    id = 'headered_chevron',
                    label = u"2. Prozessschritt mit Kopfzeile",
                    image = "headered_chevron",
                    screentip="Prozess-Schritt-Shape mit Kopfzeile erstellen",
                    supertip="Erstelle einen Prozess-Pfeil mit Header-Shape. Das Header-Shape kann im Prozess-Pfeil über Kontext-Menü des Header-Shapes passend angeordnet werden.",
                    on_action=bkt.Callback(Pentagon.create_headered_chevron)
                ),
                harvey.harvey_create_button,
                traffic_light.traffic_light_create_button,
                stateshapes.likert_button,
                bkt.ribbon.MenuSeparator(title="Verbindungsflächen"),
                bkt.ribbon.Button(
                    id = 'connector_h',
                    label = u"Horizontale Verbindungsfläche",
                    image = "ConnectorHorizontal",
                    supertip="Erstelle eine horizontale Verbindungsfläche zwischen den vertikalen Seiten (links/rechts) von zwei Shapes.",
                    on_action=bkt.Callback(ShapesMore.addHorizontalConnector, context=True, shapes=True, shapes_min=2, shapes_max=2),
                    get_enabled = bkt.apps.ppt_shapes_exactly2_selected,
                ),
                bkt.ribbon.Button(
                    id = 'connector_v',
                    label = u"Vertikale Verbindungsfläche",
                    image = "ConnectorVertical",
                    supertip="Erstelle eine vertikale Verbindungsfläche zwischen den horizontalen Seiten (oben/unten) von zwei Shapes.",
                    on_action=bkt.Callback(ShapesMore.addVerticalConnector, context=True, shapes=True, shapes_min=2, shapes_max=2),
                    get_enabled = bkt.apps.ppt_shapes_exactly2_selected,
                ),
            ]
        ),
        bkt.mso.control.ShapeChangeShapeGallery,
        bkt.ribbon.Menu(
            image_mso='CombineShapesMenu',
            label="Shape verändern",
            show_label=False,
            children=[
                bkt.ribbon.MenuSeparator(title="Formen manipulieren"),
                bkt.ribbon.Button(
                    label="Shapes teilen/vervielfachen…",
                    image="split_horizontal",
                    screentip="Shapes teilen oder vervielfachen",
                    supertip="Shape horizontal/vertikal in mehrere Shapes teilen oder verfielfachen.",
                    on_action=bkt.Callback(ShapeDialogActions.shape_split),
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
        ),
        bkt.ribbon.Menu(
            label='Mehr',
            show_label=False,
            image_mso='TableDesign',
            screentip="Weitere Funktionen",
            supertip="Bilder, Objekte, Spezial-Shapes, Text-Zerlegung, Shapes verstecken, ...",
            children = [
                bkt.ribbon.MenuSeparator(title="Bilder und Objekte"),
                bkt.mso.control.PictureInsertFromFilePowerPoint,
                bkt.mso.control.OleObjectctInsert,
                bkt.mso.control.ClipArtInsertDialog,
                bkt.mso.control.SmartArtInsert,
                bkt.mso.control.ChartInsert,
                bkt.ribbon.MenuSeparator(title="Text && Beschriftungen"),
                bkt.mso.control.HeaderFooterInsert,
                bkt.mso.control.DateAndTimeInsert,
                bkt.mso.control.NumberInsert,
                bkt.mso.control.InsertNewComment,
                bkt.ribbon.MenuSeparator(title="Ein-/Ausblenden"),
                bkt.ribbon.Button(
                    id = 'hide_shape',
                    label = u"Shapes verstecken",
                    image_mso="ShapesSubtract",
                    supertip="Verstecke alle markierten Shapes (visible=False).",
                    on_action=bkt.Callback(ShapesMore.hide_shapes),
                    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                ),
                bkt.ribbon.Button(
                    id = 'show_shapes',
                    label = u"Versteckte Shapes einblenden",
                    supertip="Mache alle versteckten Shapes (visible=False) wieder sichtbar.",
                    on_action=bkt.Callback(ShapesMore.show_shapes)
                ),

            ]
        ),
    ]
)
