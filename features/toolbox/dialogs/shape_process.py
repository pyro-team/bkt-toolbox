# -*- coding: utf-8 -*-
'''
Created on 2018-05-29
@author: Florian Stallmann
'''

from __future__ import absolute_import, division

import logging
from uuid import uuid4

import bkt.ui
notify_property = bkt.ui.notify_property

import bkt.library.powerpoint as pplib
pt_to_cm = pplib.pt_to_cm
cm_to_pt = pplib.cm_to_pt

from ..models.processshapes import ProcessChevrons


class ProcessChevronsModel(object):

    @staticmethod
    def determine_from_shape(process_shape):
        try:
            num_rows = len(ProcessChevrons._find_row_shapes_for_process(process_shape.parent, process_shape))
        except ValueError:
            num_rows = 0

        num_steps, first_pentagon, spacing, height = ProcessChevrons._determine_from_process(process_shape)

        return num_steps, first_pentagon, spacing, height, num_rows

    @classmethod
    def create_process(cls, slide, num_steps=3, first_pentagon=True, spacing=5, height=50, num_rows=2):
        ref_left,ref_top,ref_width,_ = pplib.slide_content_size(slide)

        p_spacing = spacing - cm_to_pt(0.5)
        width   = (ref_width+p_spacing)/num_steps-p_spacing
        top     = ref_top
        left    = ref_left
        adj     = cm_to_pt(0.5)/min(width, height)

        uuid = str(uuid4())
        
        #create process shapes
        first = True
        for _ in range(num_steps):
            if first:
                st = pplib.MsoAutoShapeType['msoShapePentagon'] if first_pentagon else pplib.MsoAutoShapeType['msoShapeChevron']
                first = False
            else:
                st = pplib.MsoAutoShapeType['msoShapeChevron']

            s=slide.shapes.addshape( st , left, top, width, height)
            s.Adjustments[1] = adj
            left += width+p_spacing
        
        #text formatting
        shapes = pplib.last_n_shapes_on_slide(slide, num_steps)
        shapes.Textframe2.TextRange.ParagraphFormat.Bullet.Type = 0
        shapes.Textframe2.TextRange.ParagraphFormat.LeftIndent = 0
        shapes.Textframe2.TextRange.ParagraphFormat.FirstLineIndent = 0
        shapes.Textframe2.VerticalAnchor = 3 #middle
        p_grp = shapes.group()
        p_grp.Name = "[BKT] Process %s" % p_grp.id
        ProcessChevrons._add_tags(p_grp, uuid)

        if num_rows:
            ProcessChevrons._create_rows_for_process(slide, p_grp, num_rows)

        #select process only
        p_grp.select()

    @classmethod
    def update_process(cls, shape, slide, num_steps=3, first_pentagon=True, spacing=5, height=50, num_rows=2):
        ref_width = shape.width

        group_shapes = list(iter(shape.GroupItems))
        cur_steps = len(group_shapes)
        
        first_shape = group_shapes[0]
        adj_value = first_shape.Adjustments[1]

        first_shape.AutoShapeType = pplib.MsoAutoShapeType['msoShapePentagon'] if first_pentagon else pplib.MsoAutoShapeType['msoShapeChevron']
        shape.height = height

        first_shape.Adjustments[1] = adj_value

        p_spacing = spacing - adj_value*min(first_shape.width,first_shape.height)
        group_shapes[1].left = first_shape.left+first_shape.width+p_spacing

        try:
            rows = ProcessChevrons._find_row_shapes_for_process(slide, shape)
        except ValueError:
            ProcessChevrons.convert_to_process_chevrons(shape)
            rows = []
        delta_rows = num_rows-len(rows)

        if delta_rows > 0:
            ProcessChevrons._add_row(slide, shape, delta_rows)
        elif delta_rows < 0:
            ProcessChevrons._remove_row(slide, shape, abs(delta_rows))
        else:
            ProcessChevrons._distribute_row_shapes(slide, shape)
        
        delta_steps = num_steps-cur_steps
        if delta_steps > 0:
            ProcessChevrons._add_chevron(slide, shape, delta_steps)
        elif delta_steps < 0:
            ProcessChevrons._remove_chevron(slide, shape, abs(delta_steps))
        else:
            ProcessChevrons._align_process_shapes(slide, shape, ref_width)
            ProcessChevrons._align_row_shapes(slide, shape)



class ViewModel(bkt.ui.ViewModelSingleton):
    
    def __init__(self):
        super(ViewModel, self).__init__()
        
        self._update_enabled = False
        self._num_steps = 3
        self._num_rows  = 2
        self._spacing   = 0.2
        self._height    = 2.0
        self._first_pentagon = True
    
    def set_values_based_on_shape(self, shape):
        if not ProcessChevrons.is_process_chevrons(shape):
            raise ValueError("not a process shape")
        self.num_steps, first_pentagon, self._spacing, self._height, self.num_rows = ProcessChevronsModel.determine_from_shape(shape)
        self.spacing = pt_to_cm(self._spacing)
        self.height = pt_to_cm(self._height)
        if first_pentagon:
            self.first_pentagon = True
        else:
            self.first_chevron = True
        self.update_enabled = True


    @notify_property
    def update_enabled(self):
        return self._update_enabled

    @update_enabled.setter
    def update_enabled(self, value):
        self._update_enabled = value

    @notify_property
    def num_steps(self):
        return self._num_steps

    @num_steps.setter
    def num_steps(self, value):
        self._num_steps = value

    @notify_property
    def spacing(self):
        return self._spacing

    @spacing.setter
    def spacing(self, value):
        self._spacing = value

    @notify_property
    def height(self):
        return self._height

    @height.setter
    def height(self, value):
        self._height = value
    
    
    @notify_property
    def num_rows(self):
        return self._num_rows

    @num_rows.setter
    def num_rows(self, value):
        self._num_rows = value

    @notify_property
    def first_pentagon(self):
        return self._first_pentagon

    @first_pentagon.setter
    def first_pentagon(self, value):
        self._first_pentagon = True

    @notify_property
    def first_chevron(self):
        return not self._first_pentagon

    @first_chevron.setter
    def first_chevron(self, value):
        self._first_pentagon = False



class ProcessWindow(bkt.ui.WpfWindowAbstract):
    _xamlname = 'shape_process'
    _vm_class = ViewModel
    
    def __init__(self, context, slide):
        super(ProcessWindow, self).__init__(context)

        self._slide = slide
        self._shape = None

        try:
            self._vm.update_enabled = False
            self._vm.set_values_based_on_shape(context.shape)
            self._shape = context.shape
        except:
            pass #e.g. nothing selected

    def cancel(self, sender, event):
        self.Close()
    
    def update_process(self, sender, event):
        try:
            if self._shape:
                ProcessChevronsModel.update_process(self._shape, self._slide, self._vm.num_steps, self._vm.first_pentagon, cm_to_pt(self._vm.spacing), cm_to_pt(self._vm.height), self._vm.num_rows)
            else:
                ProcessChevronsModel.create_process(self._slide, self._vm.num_steps, self._vm.first_pentagon, cm_to_pt(self._vm.spacing), cm_to_pt(self._vm.height), self._vm.num_rows)
        except:
            logging.exception("Dialog action failed")
        finally:
            self.Close()
    
    def create_process(self, sender, event):
        try:
            ProcessChevronsModel.create_process(self._slide, self._vm.num_steps, self._vm.first_pentagon, cm_to_pt(self._vm.spacing), cm_to_pt(self._vm.height), self._vm.num_rows)
        except:
            logging.exception("Dialog action failed")
        finally:
            self.Close()
