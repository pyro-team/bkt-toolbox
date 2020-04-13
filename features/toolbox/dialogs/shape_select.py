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
    def __init__(self, model, context):
        super(ViewModel, self).__init__()

        self._model = model
        self._context = context

        self._shape_keys = {k: False for k in model.key_functions.keys()}

        # self._shape_all = False
        # self._pos_all = False
        # self._fill_all = False
        # self._line_all = False
        # self._font_all = False
        # self._content_all = False


    # @notify_property
    # def shape_all(self):
    #     return self._shape_all
    # @shape_all.setter
    # def shape_all(self, value):
    #     self._shape_all = value
    #     self.shape_keys["shape_type"] = value
    #     self.shape_keys["shape_width"] = value
    #     self.shape_keys["shape_height"] = value
    #     self.OnPropertyChanged('shape_keys')

    # @notify_property
    # def pos_all(self):
    #     return self._pos_all
    # @pos_all.setter
    # def pos_all(self, value):
    #     self._pos_all = value
    #     self.shape_keys["pos_left"] = value
    #     self.shape_keys["pos_top"] = value
    #     self.shape_keys["pos_rotation"] = value
    #     self.OnPropertyChanged('shape_keys')

    # @notify_property
    # def fill_all(self):
    #     return self._fill_all
    # @fill_all.setter
    # def fill_all(self, value):
    #     self._fill_all = value
    #     self.shape_keys["fill_type"] = value
    #     self.shape_keys["fill_color"] = value
    #     self.OnPropertyChanged('shape_keys')

    # @notify_property
    # def line_all(self):
    #     return self._line_all
    # @line_all.setter
    # def line_all(self, value):
    #     self._line_all = value
    #     self.shape_keys["line_weight"] = value
    #     self.shape_keys["line_style"] = value
    #     self.shape_keys["line_color"] = value
    #     self.shape_keys["line_begin"] = value
    #     self.shape_keys["line_end"] = value
    #     self.OnPropertyChanged('shape_keys')

    # @notify_property
    # def font_all(self):
    #     return self._font_all
    # @font_all.setter
    # def font_all(self, value):
    #     self._font_all = value
    #     self.shape_keys["font_name"] = value
    #     self.shape_keys["font_color"] = value
    #     self.shape_keys["font_style"] = value
    #     self.OnPropertyChanged('shape_keys')

    # @notify_property
    # def content_all(self):
    #     return self._content_all
    # @content_all.setter
    # def content_all(self, value):
    #     self._content_all = value
    #     self.shape_keys["content_len"] = value
    #     self.shape_keys["content_text"] = value
    #     self.OnPropertyChanged('shape_keys')

    @notify_property
    def shape_keys(self):
        return self._shape_keys
    @shape_keys.setter
    def shape_keys(self, value):
        self._shape_keys = value


class SelectWindow(bkt.ui.WpfWindowAbstract):
    _xamlname = 'shape_select'
    # _vm_class = ViewModel

    def __init__(self, model, context):
        self._model = model
        self._vm = ViewModel(model,context)
        self._context = context
        self._master_shapes = context.shapes[:]

        super(SelectWindow, self).__init__()

    def cancel(self, sender, event):
        self.Close()
        #undo preview by selecting only master shapes
        self._model.selectShapes(self._context, self._master_shapes)
    
    def shapes_select(self, sender, event):
        logging.debug("SelectWindow.shapes_select")
        keys = [k for k,v in self._vm._shape_keys.iteritems() if v==True]
        self._model.selectByKeys(self._context, keys, self._master_shapes, True)
    
    def shapes_select_close(self, sender, event):
        self.Close()
        self.shapes_select(sender, event)