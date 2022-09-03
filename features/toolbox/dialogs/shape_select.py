# -*- coding: utf-8 -*-
'''
Created on 2018-05-29
@author: Florian Stallmann
'''

from __future__ import absolute_import

import logging

import bkt.ui
notify_property = bkt.ui.notify_property

class ViewModel(bkt.ui.ViewModelSingleton):
    selectors = {
        'shape_all':    ["shape_type", "shape_width", "shape_height"],
        'pos_all':      ["pos_left", "pos_top", "pos_right", "pos_bottom", "pos_rotation"],
        'fill_all':     ["fill_type", "fill_color", "fill_transp"],
        'line_all':     ["line_weight", "line_style", "line_color", "line_begin", "line_end"],
        'font_all':     ["font_name", "font_size", "font_color", "font_style"],
        'content_all':  ["content_len", "content_text"],
    }

    def __init__(self, model, context):
        super(ViewModel, self).__init__()

        self._model = model
        self._context = context

        self._shape_keys = {k: False for k in model.key_functions.keys()}
    
    def __getattr__(self, name):
        return self._shape_keys[name[3:]]
    
    def __setattr__(self, name, value):
        if name.startswith("sk_"):
            self._shape_keys[name[3:]] = value
            self.OnPropertyChanged(name)
        else:
            super(ViewModel, self).__setattr__(name, value)

    # @notify_property
    # def shape_keys(self):
    #     return self._shape_keys
    # @shape_keys.setter
    # def shape_keys(self, value):
    #     self._shape_keys = value


class SelectWindow(bkt.ui.WpfWindowAbstract):
    _xamlname = 'shape_select'
    # _vm_class = ViewModel

    def __init__(self, model, context):
        self._model = model
        self._vm = ViewModel(model, context)
        self._master_shapes = context.shapes[:] #copy of list

        super(SelectWindow, self).__init__(context)

    def cancel(self, sender, event):
        self.Close()
        #undo preview by selecting only master shapes
        self._model.selectShapes(self._context, self._master_shapes)

    def select_all(self, sender, event):
        tag = sender.Tag
        keys = self._vm.selectors.get(tag, [])
        for key in keys:
            setattr(self._vm, "sk_"+key, True)

    def select_none(self, sender, event):
        tag = sender.Tag
        keys = self._vm.selectors.get(tag, [])
        for key in keys:
            setattr(self._vm, "sk_"+key, False)
    
    def shapes_select(self, sender, event):
        logging.debug("SelectWindow.shapes_select")
        keys = [k for k,v in self._vm._shape_keys.iteritems() if v==True]
        self._model.selectByKeys(self._context, keys, self._master_shapes, True)
    
    def shapes_select_close(self, sender, event):
        self.Close()
        self.shapes_select(sender, event)