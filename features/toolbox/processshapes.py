# -*- coding: utf-8 -*-
'''
Created on 08.07.2019

@author: fstallmann
'''

from __future__ import absolute_import

from contextlib import contextmanager #for flip and rotation correction

import bkt

import bkt.library.powerpoint as pplib
pt_to_cm = pplib.pt_to_cm
cm_to_pt = pplib.cm_to_pt


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
        return pplib.TagHelper.has_tag(shape, bkt.contextdialogs.BKT_CONTEXTDIALOG_TAGKEY, cls.BKT_DIALOG_TAG)

    @classmethod
    def _add_tags(cls, shape):
        shape.Tags.Add(bkt.contextdialogs.BKT_CONTEXTDIALOG_TAGKEY, cls.BKT_DIALOG_TAG)

    @classmethod
    def create_process(cls, slide, num_steps=3, first_pentagon=True, spacing=5):
        ref_left,ref_top,ref_width,_ = pplib.slide_content_size(slide.parent)

        width=(ref_width+spacing)/num_steps-spacing
        height=50
        top=ref_top
        left=ref_left

        if first_pentagon:
            slide.shapes.addshape( pplib.MsoAutoShapeType['msoShapePentagon'] , left, top, width, height)
        else:
            slide.shapes.addshape( pplib.MsoAutoShapeType['msoShapeChevron'] , left, top, width, height)
        
        for _ in range(num_steps-1):
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
    def add_chevron(cls, shapes):
        for shape in shapes:
            cls._add_chevron(shape)

    @classmethod
    def _add_chevron(cls, shape):
        group = pplib.GroupManager(shape, additional_attrs=["width"])
        group.prepare_ungroup()

        # group_shapes = sorted(iter(shape.GroupItems), key=lambda s: s.left)
        group_shapes = group.child_items
        group_shapes.sort(key=lambda s: s.left)
        ref_shape = group_shapes[-1]
        dis_shape = group_shapes[-2]

        new_shape = ref_shape.Duplicate()

        new_shape.top = ref_shape.top
        distance = ref_shape.left - dis_shape.left-dis_shape.width
        new_shape.left = ref_shape.left+ref_shape.width+distance

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

        # group_shapes = sorted(iter(shape.GroupItems), key=lambda s: s.left)
        group_shapes = group.child_items
        group_shapes.sort(key=lambda s: s.left)
        ref_shape = group_shapes[-1]
        
        ref_shape.delete()

        group.refresh()
        group.select(False)
        return group.shape



@contextmanager
def flip_and_rotation_correction(body, header):
    # NOTE: is this flipping correction useful for any other bkt function?
    # following situations can happen:
    #   group   | header    | flipping correction
    #   --------|-----------|----------------------
    #   0       | 0         | no correction
    #   -1      | 0         | flip group
    #   0       | -1        | flip group
    #  -1       | -1        | no correction
    #  (none)   | 0         | no correction
    #  (none)   | -1        | flip header, align edge at the end
    #

    try:
        flip_header = False
        flip_body = False
        stored_rotation = 0 #FIXME: works only for groups

        is_group_child = pplib.shape_is_group_child(header)
        group_is_flipped = is_group_child and header.ParentGroup.HorizontalFlip

        #set rotation to 0 for rotated groups
        if is_group_child and header.ParentGroup.Rotation != 0:
            stored_rotation = header.ParentGroup.Rotation
            header.ParentGroup.Rotation = 0

        #check if flip correction needs to be applied
        if group_is_flipped != header.HorizontalFlip: #XOR
            if is_group_child:
                flip_body = True
                header.ParentGroup.Flip(0) #msoFlipHorizontal
            else:
                flip_header = True
                header.Flip(0) #msoFlipHorizontal

        yield body, header #contextmanager requires a yield
    finally:
        #restore flip for groups
        if flip_body:
            header.ParentGroup.Flip(0) #msoFlipHorizontal
        #restore flip and correct edge for header without group
        elif flip_header:
            header.Flip(0) #msoFlipHorizontal
            header.left = body.left + body.width - header.width
        #restore group rotation
        if stored_rotation != 0:
            header.ParentGroup.Rotation = stored_rotation


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

        with flip_and_rotation_correction(pentagon, header):
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
        
        with flip_and_rotation_correction(chevron, header):
            cls.update_pentagon_header(chevron, header)
            # header punkt links unten
            offset = chevron.Adjustments.item[1] * min(chevron.width, chevron.height)
            header.Nodes.SetPosition(4, chevron.left + ( header.height/(chevron.height/2) * offset), chevron.top + header.height)
        
        

    @classmethod
    def is_headered_group(cls, shape):
        ''' returns true for group-shapes (header+body) '''
        pentagon, _ = cls.get_body_and_header_from_group(shape)
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
    def search_body_and_update_header(cls, context, shape):
        ''' for the pentagon represented by the given shape (header, body, or group header+body), the header position and size are updated '''
        header = shape
        if pplib.shape_is_group_child(header):
            shapes = list(iter(shape.ParentGroup.GroupItems))
        else:
            shapes = list(iter(context.slide.shapes))
        body = cls.find_corresponding_body_shape(shapes, header)
        if not body:
            bkt.helpers.error("Fehler: Zugehöriges Prozess-Shape nicht gefunden!")
        else:
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
