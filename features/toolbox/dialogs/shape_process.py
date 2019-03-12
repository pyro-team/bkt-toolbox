# -*- coding: utf-8 -*-
'''
Created on 2018-05-29
@author: Florian Stallmann
'''

import os.path
import bkt.ui
notify_property = bkt.ui.notify_property

import bkt.library.powerpoint as pplib
pt_to_cm = pplib.pt_to_cm
cm_to_pt = pplib.cm_to_pt
# from bkt.library.powerpoint import pt_to_cm, cm_to_pt


class ProcessChevrons(object):
    @staticmethod
    def create_process(slide, num_steps=3, first_pentagon=True, spacing=5):
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
        shapes.group().select()



class ViewModel(bkt.ui.ViewModelSingleton):
    
    def __init__(self):
        super(ViewModel, self).__init__()
        
        self._num_steps = 3
        self._spacing   = 0.2
        self._first_pentagon = True
    
    
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
    _filename = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'shape_process.xaml')
    _vm_class = ViewModel
    
    def __init__(self, slide):
        self._slide = slide
        super(ProcessWindow, self).__init__()

    def cancel(self, sender, event):
        self.Close()
    
    def create_process(self, sender, event):
        ProcessChevrons.create_process(self._slide, self._vm.num_steps, self._vm.first_pentagon, cm_to_pt(self._vm.spacing))
        self.Close()
