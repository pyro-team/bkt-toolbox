# -*- coding: utf-8 -*-
'''
Created on 2018-05-29
@author: Florian Stallmann
'''

import os.path
import bkt.ui
notify_property = bkt.ui.notify_property


class ViewModel(bkt.ui.ViewModelAsbtract):
    def __init__(self, model, cur_slideno, max_slideno, initial_num_slides=1):
        super(ViewModel, self).__init__()
        
        self._num_slides = max(0, initial_num_slides)
        self._copymode_all = True
        self._copymode_num = False

        self.max_slides = max(0, max_slideno-cur_slideno)

        self.cur_slideno = cur_slideno
        self.max_slideno = max_slideno
    
    @notify_property
    def num_slides(self):
        return self._num_slides
    @num_slides.setter
    def num_slides(self, value):
        self._num_slides = value
        self._copymode_all = False
        self._copymode_num = True
        self.OnPropertyChanged('slide_no')
        self.OnPropertyChanged('copymode_all')
        self.OnPropertyChanged('copymode_num')
        self.OnPropertyChanged('okay_enabled')
        self.OnPropertyChanged('copy_description')
    
    @notify_property
    def slide_no(self):
        return self.cur_slideno + self._num_slides
    @slide_no.setter
    def slide_no(self, value):
        self._num_slides = value - self.cur_slideno
        self._copymode_all = False
        self._copymode_num = True
        self.OnPropertyChanged('num_slides')
        self.OnPropertyChanged('copymode_all')
        self.OnPropertyChanged('copymode_num')
        self.OnPropertyChanged('okay_enabled')
        self.OnPropertyChanged('copy_description')
    
    @notify_property
    def copymode_all(self):
        return self._copymode_all
    @copymode_all.setter
    def copymode_all(self, value):
        self._copymode_all = value
        self.OnPropertyChanged('okay_enabled')
        self.OnPropertyChanged('copy_description')
    
    @notify_property
    def copymode_num(self):
        return self._copymode_num
    @copymode_num.setter
    def copymode_num(self, value):
        self._copymode_num = value
        self.OnPropertyChanged('okay_enabled')
        self.OnPropertyChanged('copy_description')
    
    @property
    def num_copies(self):
        return self._num_slides if self._copymode_num else self.max_slides

    @notify_property
    def copy_description(self):
        num_copies = self.num_copies
        return "Kopiere {} mal von Foliennummer {} bis {}.".format(num_copies, self.cur_slideno, self.cur_slideno+num_copies)
    
    @notify_property
    def okay_enabled(self):
        return self.num_copies > 0


class CopyWindow(bkt.ui.WpfWindowAbstract):
    _filename = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'linkshapes_copy.xaml')
    # _vm_class = ViewModel

    def __init__(self, model, context, shape):
        self.context = context
        self.shape = shape

        self._model = model
        cur_slide = context.slide.slideindex
        self._vm = ViewModel(self, cur_slide, context.presentation.slides.count, self._get_last_slideindex_in_section(context)-cur_slide)

        super(CopyWindow, self).__init__()
    
    def _get_last_slideindex_in_section(self, context):
        sections = context.presentation.sectionProperties
        if sections.Count == 0:
            return context.presentation.slides.count
        cur_section = context.slide.sectionIndex
        return sections.FirstSlide(cur_section) + sections.SlidesCount(cur_section) - 1

    def cancel(self, sender, event):
        self.Close()
    
    def linkshapes_copy(self, sender, event):
        self.Close()
        if self._vm.copymode_num:
            self._model.copy_shapes_to_slides([self.shape], self.context, self._vm.num_slides)
        else:
            self._model.copy_shapes_to_slides([self.shape], self.context)