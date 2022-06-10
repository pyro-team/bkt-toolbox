# -*- coding: utf-8 -*-
'''
Created on 08.07.2019

@author: fstallmann
'''

from __future__ import absolute_import

import logging

from contextlib import contextmanager #for flip and rotation correction

import bkt

import bkt.library.powerpoint as pplib
pt_to_cm = pplib.pt_to_cm
cm_to_pt = pplib.cm_to_pt


class ProcessChevrons(object):
    BKT_DIALOG_TAG = "BKT_PROCESS_CHEVRONS"
    BKT_ROW_TAG = "BKT_PROCESS_CHEVRONS_ROW"

    ### TAGS ###

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
        from uuid import uuid4
        cls._add_tags(shape, str(uuid4()))

    @classmethod
    def is_process_chevrons(cls, shape):
        return pplib.TagHelper.has_tag(shape, bkt.contextdialogs.BKT_CONTEXTDIALOG_TAGKEY, cls.BKT_DIALOG_TAG)

    @classmethod
    def _add_tags(cls, shape, uuid):
        shape.Tags.Add(bkt.contextdialogs.BKT_CONTEXTDIALOG_TAGKEY, cls.BKT_DIALOG_TAG)
        shape.Tags.Add(cls.BKT_DIALOG_TAG, uuid)

    @classmethod
    def _add_tags_row(cls, shape, uuid):
        shape.Tags.Add(bkt.contextdialogs.BKT_CONTEXTDIALOG_TAGKEY, cls.BKT_ROW_TAG)
        shape.Tags.Add(cls.BKT_ROW_TAG, uuid)

    @classmethod
    def _get_uuid(cls, shape):
        uuid = pplib.TagHelper.get_tag(shape, cls.BKT_DIALOG_TAG)
        if not uuid:
            raise ValueError("not a process shape")
        return uuid

    ### ALIGNING PROCESS ###
    
    @classmethod
    def align_process(cls, slide, shapes):
        for shape in shapes:
            try:
                cls._align_process_shapes(slide, shape)
                cls._align_row_shapes(slide, shape)
            except:
                logging.exception("error aligning process")

    @classmethod
    def _align_process_shapes(cls, slide, process_shape, ref_width=None):
        group_shapes = list(iter(process_shape.GroupItems))
        num_steps = len(group_shapes)
        first_shape = group_shapes[0]

        if not ref_width:
            ref_width = process_shape.width

        adj_value = first_shape.Adjustments[1]*min(first_shape.width,first_shape.height)
        dis_value = group_shapes[1].left - first_shape.left-first_shape.width

        left = first_shape.left
        new_width = (ref_width+dis_value)/num_steps - dis_value
        for shape in group_shapes:
            shape.left = left
            shape.width = new_width
            shape.Adjustments[1] = adj_value/min(shape.width,shape.height)
            left += shape.width+dis_value

    ### HELPER ###

    @classmethod
    def _determine_from_process(cls, shape):
        group_shapes = list(iter(shape.GroupItems))
        num_steps = len(group_shapes)
        first_shape = group_shapes[0]

        first_pentagon = first_shape.AutoShapeType == pplib.MsoAutoShapeType['msoShapePentagon']

        adj_value = first_shape.Adjustments[1]*min(first_shape.width,first_shape.height)
        dis_value = group_shapes[1].left - first_shape.left-first_shape.width
        spacing = dis_value+adj_value

        return num_steps, first_pentagon, spacing, shape.height

    @classmethod
    def _find_row_shapes_for_process(cls, slide, shape):
        return cls._find_row_shapes(slide, cls._get_uuid(shape))

    @classmethod
    def _find_row_shapes(cls, slide, uuid):
        result = []
        for shape in slide.shapes:
            if pplib.TagHelper.has_tag(shape, cls.BKT_ROW_TAG, uuid):
                result.append(shape)
        return sorted(result, key=lambda s: s.top)
    
    ### ALIGNING ROWS ###
    
    @classmethod
    def _align_row_shapes(cls, slide, process_shape):
        group_shapes = list(iter(process_shape.GroupItems))
        len_process = len(group_shapes)
        first_shape = group_shapes[0]

        adj_value = first_shape.Adjustments[1]*min(first_shape.width,first_shape.height)

        for row in cls._find_row_shapes_for_process(slide, process_shape):
            group_row = pplib.GroupManager(row)
            group_row.prepare_ungroup()
            row_shapes = sorted(group_row.child_items, key=lambda s:s.left)
            for i in range(max(len_process, len(row_shapes))):
                try:
                    p_shape = group_shapes[i]
                except IndexError:
                    logging.debug("Process: Too many row shapes, deleting shape %s", i)
                    row_shapes[i].Delete()
                else:
                    try:
                        shape = row_shapes[i]
                    except IndexError:
                        logging.debug("Process: Too few row shapes, duplicating shape %s", i)
                        shape = row_shapes[-1].Duplicate()
                        shape.top = row_shapes[-1].top
                    
                    shape.left = p_shape.left
                    shape.width = p_shape.width - adj_value
            group_row.refresh()
    
    @classmethod
    def _distribute_row_shapes(cls, slide, process_shape, ref_height=None):
        rows = sorted(cls._find_row_shapes_for_process(slide, process_shape), key=lambda r:r.top)

        if not rows:
            return
        
        if not ref_height:
            # ref_height = rows[-1].top+rows[-1].height-rows[0].top
            ref_height = rows[-1].top+rows[-1].height - process_shape.top-process_shape.height
        
        _, _, spacing, height = cls._determine_from_process(process_shape)

        new_height = ref_height/len(rows) - spacing
        top = process_shape.top+height+spacing

        for row in rows:
            row.top = top
            row.height = new_height
            top += new_height+spacing

    ### ADD/REMOVE ROWS

    @classmethod
    def _create_rows_for_process(cls, slide, shape, num_rows=1):
        uuid = cls._get_uuid(shape)

        _,_,_,ref_height = pplib.slide_content_size(slide)

        num_steps, _, spacing, height = cls._determine_from_process(shape)
        first_shape = shape.GroupItems[1]
        adj_value = first_shape.Adjustments[1]*min(first_shape.width,first_shape.height)

        #create row shapes
        rect_top    = shape.top+height+spacing
        rect_height = max(1,(ref_height-height)/num_rows-spacing)
        rect_width  = first_shape.width-adj_value

        for _ in range(num_rows):
            left = shape.left
            for _ in range(num_steps):
                slide.shapes.addshape( pplib.MsoAutoShapeType['msoShapeRectangle'] , left, rect_top, rect_width, rect_height)
                left += rect_width+spacing
            shapes = pplib.last_n_shapes_on_slide(slide, num_steps)
            grp = shapes.group()
            ProcessChevrons._add_tags_row(grp, uuid)
            grp.Name = "[BKT] Process-Row %s.%s" % (shape.id, grp.id)
            # grp.select(False)
            rect_top += rect_height+spacing

    @classmethod
    def _add_row(cls, slide, shape, n=1):
        rows = sorted(cls._find_row_shapes_for_process(slide, shape), key=lambda r:r.top)

        if not rows:
            return cls._create_rows_for_process(slide, shape, n)

        # ref_height = rows[-1].top+rows[-1].height-rows[0].top
        ref_height = rows[-1].top+rows[-1].height - shape.top-shape.height

        ref = rows[-1]
        for _ in range(n):
            new = ref.Duplicate()
            new.Name = "[BKT] Process-Row %s.%s" % (shape.id, new.id)
            new.Left = ref.Left
        
        cls._distribute_row_shapes(slide, shape, ref_height)

    @classmethod
    def _remove_row(cls, slide, shape, n=1):
        rows = sorted(cls._find_row_shapes_for_process(slide, shape), key=lambda r:r.top)

        # ref_height = rows[-1].top+rows[-1].height-rows[0].top
        ref_height = rows[-1].top+rows[-1].height - shape.top-shape.height

        for _ in range(n):
            if len(rows) == 0:
                break
            row = rows.pop(-1)
            row.delete()
        
        if rows:
            cls._distribute_row_shapes(slide, shape, ref_height)


    ### ADD/REMOVE CHEVRONS ###

    @classmethod
    def add_chevron(cls, slide, shapes):
        for shape in shapes:
            try:
                cls._add_chevron(slide, shape)
            except:
                logging.exception("error adding chevron to process")

    @classmethod
    def _add_chevron(cls, slide, shape, n=1):
        ref_width = shape.width

        logging.info("_add_chevron =%s", n)

        # group = pplib.GroupManager(shape, additional_attrs=["width"])
        group = pplib.GroupManager(shape)
        group.prepare_ungroup()

        group_shapes = sorted(group.child_items, key=lambda s:s.left)
        ref_shape = group_shapes[-1]

        for _ in range(n):
            new_shape = ref_shape.Duplicate()
            new_shape.top = ref_shape.top

        # group_shapes.append(new_shape) #this is not really required...
        group.refresh()

        cls._align_process_shapes(slide, group.shape, ref_width)
        cls._align_row_shapes(slide, group.shape)

        group.select(False)

    @classmethod
    def remove_chevron(cls, slide, shapes):
        for shape in shapes:
            try:
                cls._remove_chevron(slide, shape)
            except:
                logging.exception("error removing chevron from process")

    @classmethod
    def _remove_chevron(cls, slide, shape, n=1):
        if shape.GroupItems.Count < 3:
            return

        logging.info("_remove_chevron =%s", n)

        ref_width = shape.width

        # group = pplib.GroupManager(shape, additional_attrs=["width"])
        group = pplib.GroupManager(shape)
        group.prepare_ungroup()

        group_shapes = sorted(group.child_items, key=lambda s:s.left)
        for _ in range(n):
            if len(group_shapes) < 3:
                break
            ref_shape = group_shapes.pop(-1)
            ref_shape.delete()

        group.refresh()

        cls._align_process_shapes(slide, group.shape, ref_width)
        cls._align_row_shapes(slide, group.shape)

        group.select(False)



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
        grp.Name = "[BKT] Headered Pentagon %s" % grp.id
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
        grp.Name = "[BKT] Headered Chevron %s" % grp.id
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
            bkt.message.error("Fehler: Zugehöriges Prozess-Shape nicht gefunden!")
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
