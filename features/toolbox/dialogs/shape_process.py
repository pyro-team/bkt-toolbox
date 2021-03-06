# -*- coding: utf-8 -*-
'''
Created on 2018-05-29
@author: Florian Stallmann
'''

from __future__ import absolute_import

import bkt.ui
notify_property = bkt.ui.notify_property

import bkt.library.powerpoint as pplib
pt_to_cm = pplib.pt_to_cm
cm_to_pt = pplib.cm_to_pt



class ViewModel(bkt.ui.ViewModelSingleton):
    
    def __init__(self):
        super(ViewModel, self).__init__()
        
        self._num_steps = 3
        self._num_rows  = 2
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
    
    def __init__(self, context, slide, model):
        self._slide = slide
        self._model = model
        super(ProcessWindow, self).__init__(context)

    def cancel(self, sender, event):
        self.Close()
    
    def create_process(self, sender, event):
        self._model.create_process(self._slide, self._vm.num_steps, self._vm.first_pentagon, cm_to_pt(self._vm.spacing), self._vm.num_rows)
        self.Close()
