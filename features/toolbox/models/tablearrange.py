# -*- coding: utf-8 -*-
'''
Created on 19.01.2023

'''

import bkt
import bkt.library.powerpoint as pplib


class TableArrange(object):
    ARRANGE_HAUTO = -1 #auto: for pars=none, for table=center
    ARRANGE_HNONE = 0
    ARRANGE_LEFT = 1
    ARRANGE_HCENTER = 2
    ARRANGE_RIGHT = 3

    ARRANGE_VAUTO = -1 #auto: for pars=line-center, for table=center
    ARRANGE_VNONE = 0
    ARRANGE_TOP = 1
    ARRANGE_VCENTER = 2
    ARRANGE_BOTTOM = 3
    ARRANGE_LCENTER = 4 #center on first line
    
    horizontal_arrangement = ARRANGE_HAUTO
    vertical_arrangement = ARRANGE_VAUTO
    
    
    @classmethod
    def arrange_shapes_on_table(cls, shapes, table):
        # all width/height of table cols/rows
        widths = [col.width for col in table.table.columns]
        heights = [row.height for row in table.table.rows]
        
        # left/top of cols/rows
        agg_widths = [ sum(widths[0:i])  for i in range(len(widths)+1)]
        agg_heights = [ sum(heights[0:i])  for i in range(len(heights)+1)]
        col_lefts = [table.x +w  for w in agg_widths]
        row_tops = [table.y +h for h in agg_heights]
        
        for shape in shapes:
            shape_midpoint = [ shape.center_x, shape.center_y ]
            
            try:
                # determine target-cell and target-rect
                try:
                    col = [c - shape_midpoint[0] >=0 for c in col_lefts].index(True)
                    row = [r - shape_midpoint[1] >=0 for r in row_tops].index(True)
                except ValueError:
                    #no col/row found, shape outside bottom-right boundaries of table
                    continue
                if col == 0 or row == 0:
                    #shape outside top-left boundaries of table
                    continue
                target_rect = [ col_lefts[col-1], row_tops[row-1], widths[col-1], heights[row-1]   ]

                # determine target-midpoint from arrangement-setting
                target_midpoint = [0,0]
                if cls.horizontal_arrangement == cls.ARRANGE_HNONE:
                    target_midpoint[0] = shape_midpoint[0] #no change in position
                elif cls.horizontal_arrangement == cls.ARRANGE_LEFT:
                    target_midpoint[0] = target_rect[0] + shape.width/2
                elif cls.horizontal_arrangement == cls.ARRANGE_RIGHT:
                    target_midpoint[0] = target_rect[0] + target_rect[2] - shape.width/2
                else: # ARRANGE_HCENTER or ARRANGE_HAUTO
                    target_midpoint[0] = target_rect[0] + target_rect[2]/2

                if cls.vertical_arrangement == cls.ARRANGE_VNONE:
                    target_midpoint[1] = shape_midpoint[1] #no change in position
                elif cls.vertical_arrangement == cls.ARRANGE_TOP:
                    target_midpoint[1] = target_rect[1] + shape.height/2
                elif cls.vertical_arrangement == cls.ARRANGE_BOTTOM:
                    target_midpoint[1] = target_rect[1] + target_rect[3] - shape.height/2
                elif cls.vertical_arrangement == cls.ARRANGE_LCENTER:
                    textframe = table.table.cell(row, col).shape.textframe2
                    if textframe.HasText == -1: #cell has text
                        first_line_height = textframe.textrange.lines[1].boundheight
                    else: #cell has no text
                        first_line_height = textframe.textrange.boundheight
                    target_midpoint[1] = target_rect[1] + first_line_height/2
                else: # ARRANGE_VCENTER or ARRANGE_VAUTO
                    target_midpoint[1] = target_rect[1] + target_rect[3]/2
                
                # move shape
                shape.x += target_midpoint[0] - shape_midpoint[0]
                shape.y  += target_midpoint[1] - shape_midpoint[1]
            
            except:
                bkt.helpers.exception_as_message()
    
    @classmethod
    def arrange_table_shapes(cls, shapes):
        shapes = pplib.wrap_shapes(shapes)
        # determine table in shapes-list
        # tables = [s for s in shapes if s.Type == pplib.MsoShapeType['msoTable']]
        tables = [s for s in shapes if s.HasTable == -1]
        shapes = [s for s in shapes if s not in tables]
        # for each table call arrange_shapes_on_table width shapes
        for table in tables:
            cls.arrange_shapes_on_table(shapes, table)
    
    
    @classmethod
    def arrange_shapes_on_paragraph(cls, shapes, textshape):
        try:
            paragraphs = [textshape.TextFrame2.TextRange.Paragraphs(idx) for idx in range(1,textshape.TextFrame2.TextRange.Paragraphs().Count+1)]
            paragraph_v_bounds = [ p.boundtop for p in paragraphs]
            # bound below paragraphs
            paragraph_v_bounds.append( paragraphs[-1].boundtop+paragraphs[-1].boundheight  )

            # FIXME: find solution that works with columns within textframe
            # if textshape.TextFrame2.Column.Number > 1:
            #     colstart = textshape.left+textshape.TextFrame2.MarginLeft
            #     colsize  = (textshape.width - textshape.TextFrame2.MarginLeft - textshape.TextFrame2.MarginRight) / textshape.TextFrame2.Column.Number
            #     cols = [ colstart + x*colsize for x in range(1,textshape.TextFrame2.Column.Number)]
            
            #list of paragraphs already used to assign a shape
            par_idx_used = []
            
            for shape in shapes:
                shape_midpoint = [ shape.center_x, shape.center_y ]
                # par_idx in [1...len(paragraphs)]
                try:
                    par_idx = [ v_bound - shape_midpoint[1]  >=0 for v_bound in paragraph_v_bounds].index(True)
                except ValueError:
                    # shape is below every paragraph, use last
                    par_idx = len(paragraphs)
                if par_idx ==0:
                    # shape is over every paragraph, use first
                    par_idx = 1

                #prevent multiple shapes on the same paragraph, e.g. on top of each other
                if par_idx in par_idx_used:
                    #FIXME: better behavior to increase par_idx until "empty slot" is found?
                    continue
                else:
                    par_idx_used.append(par_idx)
                    
                #FIXME: SpaceBefore/After give lines (not points) when LineRuleBefore/After, but should be a special case as this cannot be defined via GUI

                # target_box: left , top , width, height
                target_box = [ paragraphs[par_idx-1].boundleft, paragraphs[par_idx-1].boundtop + paragraphs[par_idx-1].paragraphFormat.SpaceBefore, 
                    paragraphs[par_idx-1].boundwidth, paragraphs[par_idx-1].boundheight - paragraphs[par_idx-1].paragraphFormat.SpaceBefore - paragraphs[par_idx-1].paragraphFormat.SpaceAfter   ]

                #alignment on first line
                align_lcenter = False
                lines = paragraphs[par_idx-1].Lines
                if cls.vertical_arrangement in [cls.ARRANGE_LCENTER, cls.ARRANGE_VAUTO] and lines().Count > 0:
                    align_lcenter = True
                    target_box[3] = lines[1].boundheight - paragraphs[par_idx-1].paragraphFormat.SpaceBefore - paragraphs[par_idx-1].paragraphFormat.SpaceAfter
                    
                    if paragraphs[par_idx-1].paragraphFormat.LineRuleWithin == -1: #msoTrue
                        linefactor = paragraphs[par_idx-1].paragraphFormat.SpaceWithin
                    else:
                        linefactor = paragraphs[par_idx-1].paragraphFormat.SpaceWithin / 1.2 / lines[1].Font.Size #1.2 seems to be an arbitrary constant of ppt


                if par_idx == 1:
                    # this is the first paragraph
                    # here boundtop and boundheight do not consider SpaceBefore
                    target_box[1] = target_box[1] - paragraphs[0].paragraphFormat.SpaceBefore
                    target_box[3] = target_box[3] + paragraphs[0].paragraphFormat.SpaceBefore


                #the following line doesnt always works (if last paragraph is single line and has space after):
                #is_last_par = textshape.TextFrame2.TextRange.Boundtop + textshape.TextFrame2.TextRange.Boundheight == paragraphs[par_idx-1].boundtop + paragraphs[par_idx-1].boundheight
                is_last_par = par_idx == len(paragraphs) #TESTME: really this simple?
                if is_last_par or (align_lcenter and lines().Count > 1):
                    # this is the last paragraph containing text OR a paragraph with more than one line and center to line is selected
                    # here boundheight does not include SpaceAfter / boundheight must be removed from height by adding it again
                    target_box[3] = target_box[3] + paragraphs[par_idx-1].paragraphFormat.SpaceAfter
                    
                    if lines().Count == 1:
                        # this is the last paragraph containing exactly one line
                        # in this case the boundwidth is a little bit smaller as new line character is not counted
                        target_box[2] = target_box[2] + 5 #5 is randomly chosen here as it is not possible to determine width of new line character
                    
                        # if align_lcenter:
                        #     # this is the last paragraph containing exactly one line and center to line is selected
                        #     # SpaceWithin under the text is not considered in some rare cases (e.g. when font size changed after line spacing is defined)
                        #     target_box[3] = target_box[3] * 1.2


                # determine target-midpoint from arrangement-setting
                target_midpoint = [0,0]
                if cls.horizontal_arrangement in [cls.ARRANGE_HNONE, cls.ARRANGE_HAUTO]:
                    target_midpoint[0] = shape_midpoint[0] #no change in position
                elif cls.horizontal_arrangement == cls.ARRANGE_LEFT:
                    target_midpoint[0] = target_box[0] - shape.width - 5 + shape.width/2 #5 is randomly chosen to give a little bit of spacing
                elif cls.horizontal_arrangement == cls.ARRANGE_RIGHT:
                    target_midpoint[0] = target_box[0] + target_box[2] + shape.width - shape.width/2
                else: # ARRANGE_HCENTER
                    target_midpoint[0] = target_box[0] + target_box[2]/2

                if cls.vertical_arrangement == cls.ARRANGE_VNONE:
                    target_midpoint[1] = shape_midpoint[1] #no change in position
                elif cls.vertical_arrangement == cls.ARRANGE_TOP:
                    target_midpoint[1] = target_box[1] + shape.height/2
                elif cls.vertical_arrangement == cls.ARRANGE_BOTTOM:
                    target_midpoint[1] = target_box[1] + target_box[3] - shape.height/2
                elif align_lcenter: #includes ARRANGE_VAUTO
                    linefactor = 1 + (linefactor-1)/10 if linefactor < 2 else 1 + linefactor/10
                    target_midpoint[1] = target_box[1] + target_box[3]/linefactor/2 + (target_box[3] - target_box[3]/linefactor)
                else: # ARRANGE_VCENTER
                    target_midpoint[1] = target_box[1] + target_box[3]/2
            
                shape.x += target_midpoint[0] - shape_midpoint[0]
                shape.y += target_midpoint[1] - shape_midpoint[1]
        except:
            bkt.helpers.exception_as_message()
            
    
    @classmethod
    def arrange_paragraph_shapes(cls, shapes):
        shapes = pplib.wrap_shapes(shapes)
        for master in shapes:
            # master has no paragraphs -> skip master
            if master.HasTextFrame != -1 or master.TextFrame2.TextRange.Paragraphs().Count == 0:
                continue

            # shapes in selection, complete in master
            #inner_shapes = [ s  for s in shapes if (s!=master and s.left>=master.left and s.top>=master.top and s.top+s.height<=master.top+master.height and s.left+s.width<=master.left+master.width )]
            # shapes in selection, being smaller and midpoint in master
            inner_shapes = [
                s  for s in shapes
                if s!=master and cls._is_shape_within(master, s)
                ]
            
            if len(inner_shapes) > 0:
                cls.arrange_shapes_on_paragraph(inner_shapes, master)
    
    @classmethod
    def arrange_shapes_on_shapes(cls, shapes, background_shapes):
        for shape in shapes:
            if cls.horizontal_arrangement == cls.ARRANGE_LEFT:
                shape.x = max(s.x for s in background_shapes)
            elif cls.horizontal_arrangement == cls.ARRANGE_RIGHT:
                shape.x1 = min(s.x1 for s in background_shapes)
            # elif cls.horizontal_arrangement == cls.ARRANGE_HCENTER or (cls.horizontal_arrangement == cls.ARRANGE_HAUTO and background_shape.height >= background_shape.width):
            elif cls.horizontal_arrangement in [cls.ARRANGE_HCENTER, cls.ARRANGE_HAUTO]:
                if cls.horizontal_arrangement != cls.ARRANGE_HAUTO or len(background_shapes) > 1 or background_shapes[0].height >= background_shapes[0].width:
                    shape.center_x = min(background_shapes, key=lambda s: s.width).center_x

            if cls.vertical_arrangement == cls.ARRANGE_TOP:
                shape.y = max(s.y for s in background_shapes)
            elif cls.vertical_arrangement == cls.ARRANGE_BOTTOM:
                shape.y1 = min(s.y1 for s in background_shapes)
            # elif cls.vertical_arrangement in [cls.ARRANGE_VCENTER,cls.ARRANGE_LCENTER] or (cls.vertical_arrangement == cls.ARRANGE_VAUTO and background_shape.width >= background_shape.height):
            elif cls.vertical_arrangement in [cls.ARRANGE_VCENTER, cls.ARRANGE_LCENTER, cls.ARRANGE_VAUTO]:
                if cls.vertical_arrangement != cls.ARRANGE_VAUTO or len(background_shapes) > 1 or background_shapes[0].width >= background_shapes[0].height:
                    shape.center_y = min(background_shapes, key=lambda s: s.height).center_y


    @classmethod
    def arrange_shapes_shapes(cls, shapes):
        shapes = pplib.wrap_shapes(shapes)
        for child in shapes:
            background_shapes = [
                s for s in shapes 
                if s!=child and s.ZOrderPosition < child.ZOrderPosition and
                cls._is_shape_within(s, child)
                ]

            if len(background_shapes) > 0:
                cls.arrange_shapes_on_shapes([child], background_shapes)
        # for master in shapes:
        #     inner_shapes = [
        #         s for s in shapes 
        #         if s!=master and s.ZOrderPosition > master.ZOrderPosition and
        #         cls._is_shape_within(master, s)
        #         ]

        #     if len(inner_shapes) > 0:
        #         cls.arrange_shapes_on_shapes(inner_shapes, [master])


    @classmethod
    def _is_shape_within(cls, outer_s, inner_s):
        #test if center point of inner_s is within bounds of outer_s
        return inner_s.width<=outer_s.width and inner_s.height<=outer_s.height and outer_s.x <= inner_s.center_x <= outer_s.x1 and outer_s.y <= inner_s.center_y <= outer_s.y1


    @classmethod
    def arrange_overlay_shapes(cls, shapes):
        from itertools import chain #chain allows to concatenate lists and return a generator

        shapes = pplib.wrap_shapes(shapes) #all functions support/require wrapped shapes

        table_shapes = []
        table_childs = []
        par_shapes = []
        par_childs = []
        remaining_shapes = []
        shape_shapes = []

        #step 1: seperate tables, paragraph shapes, and all the remaining shapes
        for s in shapes:
            #test if shape is a table
            if s.HasTable == -1:
                table_shapes.append(s)
            
            #test if shape has a paragraph
            elif s.HasTextFrame == -1 and s.TextFrame2.TextRange.Paragraphs().Count > 1:
                par_shapes.append(s)

            else:
                remaining_shapes.append(s)
        
        #step 2: from remaining shapes, find shapes within tables and paragraph shapes from step 1
        for s in remaining_shapes:
            #get all shapes within tables
            if any(cls._is_shape_within(o, s) for o in table_shapes):
                table_childs.append(s)
            
            #get all shapes within paragraphs
            elif any(cls._is_shape_within(o, s) for o in par_shapes):
                par_childs.append(s)
            
            else:
                shape_shapes.append(s)

        #arrange on table
        cls.arrange_table_shapes(chain(table_shapes, table_childs))
        #arrange on paragraph
        cls.arrange_paragraph_shapes(chain(par_shapes, par_childs))
        #arrange shapes on shapes
        cls.arrange_shapes_shapes(shape_shapes)


class TALocPin(pplib.LocPin):
    def __init__(self):
        super(TALocPin, self).__init__()
        self.locpins = [
            (TableArrange.ARRANGE_VAUTO,  TableArrange.ARRANGE_HAUTO), (TableArrange.ARRANGE_VNONE,  TableArrange.ARRANGE_LEFT), (TableArrange.ARRANGE_VNONE,  TableArrange.ARRANGE_HCENTER), (TableArrange.ARRANGE_VNONE,  TableArrange.ARRANGE_RIGHT),
            (TableArrange.ARRANGE_TOP,    TableArrange.ARRANGE_HNONE), (TableArrange.ARRANGE_TOP,    TableArrange.ARRANGE_LEFT), (TableArrange.ARRANGE_TOP,    TableArrange.ARRANGE_HCENTER), (TableArrange.ARRANGE_TOP,    TableArrange.ARRANGE_RIGHT),
            (TableArrange.ARRANGE_LCENTER,TableArrange.ARRANGE_HNONE), (TableArrange.ARRANGE_LCENTER,TableArrange.ARRANGE_LEFT), (TableArrange.ARRANGE_LCENTER,TableArrange.ARRANGE_HCENTER), (TableArrange.ARRANGE_LCENTER,TableArrange.ARRANGE_RIGHT),
            (TableArrange.ARRANGE_VCENTER,TableArrange.ARRANGE_HNONE), (TableArrange.ARRANGE_VCENTER,TableArrange.ARRANGE_LEFT), (TableArrange.ARRANGE_VCENTER,TableArrange.ARRANGE_HCENTER), (TableArrange.ARRANGE_VCENTER,TableArrange.ARRANGE_RIGHT),
            (TableArrange.ARRANGE_BOTTOM, TableArrange.ARRANGE_HNONE), (TableArrange.ARRANGE_BOTTOM, TableArrange.ARRANGE_LEFT), (TableArrange.ARRANGE_BOTTOM, TableArrange.ARRANGE_HCENTER), (TableArrange.ARRANGE_BOTTOM, TableArrange.ARRANGE_RIGHT),
        ]