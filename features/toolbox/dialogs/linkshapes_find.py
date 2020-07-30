# -*- coding: utf-8 -*-
'''
Created on 2018-05-29
@author: Florian Stallmann
'''

from __future__ import absolute_import, division

import bkt.ui
notify_property = bkt.ui.notify_property


class ViewModel(bkt.ui.ViewModelAsbtract):
    default_shape_keys = {'type': False, 'x': True, 'y': True, 'x1': True, 'y1': True, 'center_x': False, 'center_y': False, 'width': False, 'height': False, 'rotation': False, 'name': False}
    default_threshold  = 0.0

    def __init__(self, context):
        super(ViewModel, self).__init__()
        
        slide = context.slide

        cur_slideno = slide.slideindex
        max_slideno = context.presentation.slides.count
        initial_num_slides = self._get_last_slideindex_in_section(context) - cur_slideno
        
        self._num_slides = max(0, initial_num_slides)
        self._findmode_all = True
        self._findmode_num = False

        #difference between slide number and slide index (slide number may begin with 0 or any value >0, so diff can also be negative)
        diff_to_real_slide_number = slide.slidenumber - cur_slideno

        self.max_slides = max(0, max_slideno-cur_slideno)

        self.cur_slideno = cur_slideno + diff_to_real_slide_number
        self.max_slideno = max_slideno + diff_to_real_slide_number

        self._threshold = ViewModel.default_threshold
        self._shape_keys = ViewModel.default_shape_keys
    
    def _get_last_slideindex_in_section(self, context):
        sections = context.presentation.sectionProperties
        if sections.Count == 0:
            return context.presentation.slides.count
        cur_section = context.slide.sectionIndex
        return sections.FirstSlide(cur_section) + sections.SlidesCount(cur_section) - 1


    @notify_property
    def threshold(self):
        return self._threshold*100
    @threshold.setter
    def threshold(self, value):
        ViewModel.default_threshold = value/100
        self._threshold = value/100
        

    @notify_property
    def attr_left(self):
        return self._shape_keys["x"]
    @attr_left.setter
    def attr_left(self, value):
        self._shape_keys["x"] = value
        self.OnPropertyChanged('okay_enabled')

    @notify_property
    def attr_top(self):
        return self._shape_keys["y"]
    @attr_top.setter
    def attr_top(self, value):
        self._shape_keys["y"] = value
        self.OnPropertyChanged('okay_enabled')

    @notify_property
    def attr_center(self):
        return self._shape_keys["center_x"]
    @attr_center.setter
    def attr_center(self, value):
        self._shape_keys["center_x"] = value
        self._shape_keys["center_y"] = value
        self.OnPropertyChanged('okay_enabled')

    @notify_property
    def attr_right(self):
        return self._shape_keys["x1"]
    @attr_right.setter
    def attr_right(self, value):
        self._shape_keys["x1"] = value
        self.OnPropertyChanged('okay_enabled')

    @notify_property
    def attr_bottom(self):
        return self._shape_keys["y1"]
    @attr_bottom.setter
    def attr_bottom(self, value):
        self._shape_keys["y1"] = value
        self.OnPropertyChanged('okay_enabled')

    @notify_property
    def attr_type(self):
        return self._shape_keys["type"]
    @attr_type.setter
    def attr_type(self, value):
        self._shape_keys["type"] = value
        self.OnPropertyChanged('okay_enabled')

    @notify_property
    def attr_width(self):
        return self._shape_keys["width"]
    @attr_width.setter
    def attr_width(self, value):
        self._shape_keys["width"] = value
        self.OnPropertyChanged('okay_enabled')

    @notify_property
    def attr_height(self):
        return self._shape_keys["height"]
    @attr_height.setter
    def attr_height(self, value):
        self._shape_keys["height"] = value
        self.OnPropertyChanged('okay_enabled')

    @notify_property
    def attr_rotation(self):
        return self._shape_keys["rotation"]
    @attr_rotation.setter
    def attr_rotation(self, value):
        self._shape_keys["rotation"] = value
        self.OnPropertyChanged('okay_enabled')

    @notify_property
    def attr_name(self):
        return self._shape_keys["name"]
    @attr_name.setter
    def attr_name(self, value):
        self._shape_keys["name"] = value
        self.OnPropertyChanged('okay_enabled')
    

    @notify_property
    def num_slides(self):
        return self._num_slides
    @num_slides.setter
    def num_slides(self, value):
        self._num_slides = value
        self._findmode_all = False
        self._findmode_num = True
        self.OnPropertyChanged('slide_no')
        self.OnPropertyChanged('findmode_all')
        self.OnPropertyChanged('findmode_num')
        self.OnPropertyChanged('okay_enabled')
        self.OnPropertyChanged('search_description')
    
    @notify_property
    def slide_no(self):
        return self.cur_slideno + self._num_slides
    @slide_no.setter
    def slide_no(self, value):
        self._num_slides = value - self.cur_slideno
        self._findmode_all = False
        self._findmode_num = True
        self.OnPropertyChanged('num_slides')
        self.OnPropertyChanged('findmode_all')
        self.OnPropertyChanged('findmode_num')
        self.OnPropertyChanged('okay_enabled')
        self.OnPropertyChanged('search_description')
    
    @notify_property
    def findmode_all(self):
        return self._findmode_all
    @findmode_all.setter
    def findmode_all(self, value):
        self._findmode_all = value
        self.OnPropertyChanged('okay_enabled')
        self.OnPropertyChanged('search_description')
    
    @notify_property
    def findmode_num(self):
        return self._findmode_num
    @findmode_num.setter
    def findmode_num(self, value):
        self._findmode_num = value
        self.OnPropertyChanged('okay_enabled')
        self.OnPropertyChanged('search_description')
    
    @property
    def num_searchslides(self):
        return self._num_slides if self._findmode_num else self.max_slides

    @notify_property
    def search_description(self):
        num_searchslides = self.num_searchslides
        return "Suche auf {} Folien von Foliennummer {} bis {}.".format(num_searchslides, self.cur_slideno, self.cur_slideno+num_searchslides)
    
    @notify_property
    def okay_enabled(self):
        return self.num_searchslides > 0 and any(self._shape_keys.values())


class FindWindow(bkt.ui.WpfWindowAbstract):
    _xamlname = 'linkshapes_find'
    # _vm_class = ViewModel

    def __init__(self, model, context, shape):
        self.shape = shape

        self._model = model
        self._vm = ViewModel(context)

        super(FindWindow, self).__init__(context)

    def _link_shapes(self, dry_run=False):
        shape_keys = [k for k,v in self._vm._shape_keys.items() if v]
        num_slides = None if self._vm.findmode_all else self._vm._num_slides
        self._model.find_similar_shapes_and_link(self.shape, self._context, shape_keys, self._vm._threshold, num_slides, dry_run)
    
    def select_all(self, sender, event):
        self._vm.attr_bottom    = True
        self._vm.attr_center    = True
        self._vm.attr_height    = True
        self._vm.attr_left      = True
        self._vm.attr_name      = True
        self._vm.attr_right     = True
        self._vm.attr_rotation  = True
        self._vm.attr_top       = True
        self._vm.attr_type      = True
        self._vm.attr_width     = True

    def select_none(self, sender, event):
        self._vm.attr_bottom    = False
        self._vm.attr_center    = False
        self._vm.attr_height    = False
        self._vm.attr_left      = False
        self._vm.attr_name      = False
        self._vm.attr_right     = False
        self._vm.attr_rotation  = False
        self._vm.attr_top       = False
        self._vm.attr_type      = False
        self._vm.attr_width     = False
    
    def linkshapes_find(self, sender, event):
        self.Close()
        self._link_shapes()
    
    def linkshapes_dryrun(self, sender, event):
        self._link_shapes(True)