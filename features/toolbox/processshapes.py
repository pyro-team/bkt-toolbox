# -*- coding: utf-8 -*-
'''
Created on 08.07.2019

@author: fstallmann
'''

import bkt

import bkt.library.powerpoint as pplib
pt_to_cm = pplib.pt_to_cm
cm_to_pt = pplib.cm_to_pt

import os.path



class ProcessChevrons(object):
    BKT_DIALOG_TAG = "BKT_PROCESS_CHEVRONS"

    @classmethod
    def is_convertible(cls, shape):
        try:
            allowed_types = [pplib.MsoAutoShapeType['msoShapePentagon'], pplib.MsoAutoShapeType['msoShapeChevron']]
            return shape.Type == pplib.MsoShapeType['msoGroup'] and \
                not cls.is_process_chevrons(shape) and \
                all(shp.AutoShapeType in allowed_types for shp in shape.GroupItems)
        except:
            return False
    
    @classmethod
    def convert_to_process_chevrons(cls, shape):
        cls._add_tags(shape)

    @classmethod
    def is_process_chevrons(cls, shape):
        try:
            return shape.Tags(bkt.contextdialogs.BKT_CONTEXTDIALOG_TAGKEY) == cls.BKT_DIALOG_TAG
        except:
            return False

    @classmethod
    def _add_tags(cls, shape):
        shape.Tags.Add(bkt.contextdialogs.BKT_CONTEXTDIALOG_TAGKEY, cls.BKT_DIALOG_TAG)

    @classmethod
    def create_process(cls, slide, num_steps=3, first_pentagon=True, spacing=5):
        ref_left,ref_top,ref_width,ref_height = pplib.slide_content_size(slide.parent)

        width=(ref_width+spacing)/num_steps-spacing
        height=50
        top=ref_top
        left=ref_left

        if first_pentagon:
            slide.shapes.addshape( pplib.MsoAutoShapeType['msoShapePentagon'] , left, top, width, height)
        else:
            slide.shapes.addshape( pplib.MsoAutoShapeType['msoShapeChevron'] , left, top, width, height)
        
        for i in range(num_steps-1):
            left += width+spacing
            slide.shapes.addshape( pplib.MsoAutoShapeType['msoShapeChevron'] , left, top, width, height)
        
        shapes = pplib.last_n_shapes_on_slide(slide, num_steps)
        shapes.Adjustments[1] = 0.28346 #0.5cm
        shapes.Textframe2.TextRange.ParagraphFormat.Bullet.Type = 0
        shapes.Textframe2.TextRange.ParagraphFormat.LeftIndent = 0
        shapes.Textframe2.TextRange.ParagraphFormat.FirstLineIndent = 0
        shapes.Textframe2.VerticalAnchor = 3 #middle
        grp = shapes.group()
        cls._add_tags(grp)
        grp.select()


    @classmethod
    def _refresh_group(cls, shape):
        #ungroup+group required so new shapes are properly added to group and counted in groupitems
        grp = shape.Ungroup()
        grp = grp.Group()
        cls._add_tags(grp)
        return grp

    @classmethod
    def add_chevron(cls, shapes):
        for shape in shapes:
            cls._add_chevron(shape)

    @classmethod
    def _add_chevron(cls, shape):
        slide = shape.parent

        group = pplib.GroupManager(shape, additional_attrs=["width"])
        group.prepare_ungroup()

        # cur_size = shape.width
        # cur_rot  = shape.rotation
        # cur_name = shape.name

        # shape.rotation = 0
        # group_shapes = sorted(iter(shape.GroupItems), key=lambda s: s.left)
        group_shapes = group.child_items
        group_shapes.sort(key=lambda s: s.left)
        ref_shape = group_shapes[-1]
        dis_shape = group_shapes[-2]

        new_shape = ref_shape.Duplicate()

        new_shape.top = ref_shape.top
        distance = ref_shape.left - dis_shape.left-dis_shape.width
        new_shape.left = ref_shape.left+ref_shape.width+distance

        # grp = cls._refresh_group(shape)
        # grp.width    = cur_size
        # grp.rotation = cur_rot
        # grp.name     = cur_name
        # grp.select()
        group.refresh()
        group.select(False)
        return group.shape

    @classmethod
    def remove_chevron(cls, shapes):
        for shape in shapes:
            cls._remove_chevron(shape)

    @classmethod
    def _remove_chevron(cls, shape):
        if shape.GroupItems.Count < 3:
            return

        group = pplib.GroupManager(shape, additional_attrs=["width"])
        group.prepare_ungroup()

        # cur_size = shape.width
        # cur_rot  = shape.rotation
        # cur_name = shape.name

        # shape.rotation = 0
        # group_shapes = sorted(iter(shape.GroupItems), key=lambda s: s.left)
        group_shapes = group.child_items
        group_shapes.sort(key=lambda s: s.left)
        ref_shape = group_shapes[-1]
        
        ref_shape.delete()

        # grp = cls._refresh_group(shape)
        # grp.width    = cur_size
        # grp.rotation = cur_rot
        # grp.name     = cur_name
        # grp.select()
        group.refresh()
        group.select(False)
        return group.shape


class ProcessChevronsPopup(bkt.ui.WpfWindowAbstract):
    _filename = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'popups', 'process_shapes.xaml')
    '''
    class representing a popup-dialog for a linked shape
    '''
    
    def __init__(self, context=None):
        self.IsPopup = True
        self._context = context

        super(ProcessChevronsPopup, self).__init__()

    def btnplus(self, sender, event):
        try:
            ProcessChevrons.add_chevron(list(iter(self._context.selection.ShapeRange)))
        except:
            bkt.helpers.message("Funktion aus unbekannten Gründen fehlgeschlagen.")
            # bkt.helpers.exception_as_message()

    def btnminus(self, sender, event):
        try:
            ProcessChevrons.remove_chevron(list(iter(self._context.selection.ShapeRange)))
        except:
            bkt.helpers.message("Funktion aus unbekannten Gründen fehlgeschlagen.")
            # bkt.helpers.exception_as_message()

#initialization function called by contextdialogs.py
def create_window(context):
    return ProcessChevronsPopup(context)



class Pentagon(object):
    
    @classmethod
    def create_headered_pentagon(cls, slide):
        ''' creates a headered pentagon on the given slide '''
        # shapeCount = slide.shapes.count
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
        # grp = slide.Shapes.Range(Array[int]([shapeCount+1, shapeCount+2])).group()
        grp = pplib.last_n_shapes_on_slide(slide, 2).group()
        grp.select()

        #cls.update_pentagon_group(grp)
        cls.update_pentagon_header(pentagon, header)

    @classmethod
    def create_headered_chevron(cls, slide):
        ''' creates a headered pentagon on the given slide '''
        # shapeCount = slide.shapes.count
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
        # grp = slide.Shapes.Range(Array[int]([shapeCount+1, shapeCount+2])).group()
        grp = pplib.last_n_shapes_on_slide(slide, 2).group()
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
