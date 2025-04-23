# -*- coding: utf-8 -*-
'''
Created on 2018-05-29
@author: Florian Stallmann
'''



import bkt.ui
notify_property = bkt.ui.notify_property


class ViewModel(bkt.ui.ViewModelAsbtract):
    def __init__(self, context):
        super(ViewModel, self).__init__()
        
        slide = context.slide

        cur_slideno = slide.slideindex
        max_slideno = context.presentation.slides.count
        initial_num_slides = self._get_last_slideindex_in_section(context) - cur_slideno

        self._num_slides = max(0, initial_num_slides)
        self._copymode_all = True
        self._copymode_num = False

        #difference between slide number and slide index (slide number may begin with 0 or any value >0, so diff can also be negative)
        diff_to_real_slide_number = slide.slidenumber - cur_slideno

        self.max_slides = max(0, max_slideno-cur_slideno)

        self.cur_slideno = cur_slideno + diff_to_real_slide_number
        self.max_slideno = max_slideno + diff_to_real_slide_number
    
    def _get_last_slideindex_in_section(self, context):
        sections = context.presentation.sectionProperties
        if sections.Count == 0:
            return context.presentation.slides.count
        cur_section = context.slide.sectionIndex
        return sections.FirstSlide(cur_section) + sections.SlidesCount(cur_section) - 1


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
    _xamlname = 'linkshapes_copy'
    # _vm_class = ViewModel

    def __init__(self, model, context, shape):
        # self.context = context
        self.shape = shape

        self._model = model
        self._vm = ViewModel(context)

        super(CopyWindow, self).__init__(context)
    
    def linkshapes_copy(self, sender, event):
        self.Close()
        if self._vm.copymode_num:
            self._model.copy_shapes_to_slides([self.shape], self._context, self._vm.num_slides)
        else:
            self._model.copy_shapes_to_slides([self.shape], self._context)