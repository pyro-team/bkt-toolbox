# -*- coding: utf-8 -*-
'''
Created on 2018-05-29
@author: Florian Stallmann
'''



import logging

from collections import OrderedDict

import bkt.ui
notify_property = bkt.ui.notify_property


class ViewModel(bkt.ui.ViewModelSingleton):
    selectors = {
        'contents_all':     ['remove_hidden_slides', 'remove_slide_notes', 'remove_slide_comments', 'remove_author', 'break_links'],
        'animation_all':    ['remove_transitions', 'remove_animations'],
        'format_all':       ['blackwhite_gray_scale', 'remove_doublespaces', 'remove_empty_placeholders'],
        'master_all':       ['remove_unused_masters', 'remove_unused_designs'],
    }

    def __init__(self):
        super(ViewModel, self).__init__()

        self._settings = OrderedDict.fromkeys(
            ViewModel.selectors['contents_all'] +
            ViewModel.selectors['animation_all'] +
            ViewModel.selectors['format_all'] +
            ViewModel.selectors['master_all'], True
            )
    
    def __getattr__(self, name):
        return self._settings[name[3:]]
    
    def __setattr__(self, name, value):
        if name.startswith("cl_"):
            self._settings[name[3:]] = value
            self.OnPropertyChanged(name)
        else:
            super(ViewModel, self).__setattr__(name, value)

    @notify_property
    def settings_all(self):
        return all(self._settings.values())
    
    @settings_all.setter
    def settings_all(self, value):
        for k in self._settings:
            self._settings[k] = value
            self.OnPropertyChanged('cl_'+k)


class SlideCleanWindow(bkt.ui.WpfWindowAbstract):
    _xamlname = 'slides_clean'
    _vm_class = ViewModel

    def __init__(self, model, context):
        self._model = model

        super(SlideCleanWindow, self).__init__(context)

    def select_all(self, sender, event):
        tag = sender.Tag
        keys = self._vm.selectors.get(tag, [])
        for key in keys:
            setattr(self._vm, "cl_"+key, True)

    def select_none(self, sender, event):
        tag = sender.Tag
        keys = self._vm.selectors.get(tag, [])
        for key in keys:
            setattr(self._vm, "cl_"+key, False)
    
    def cleanup(self, sender, event):
        self.Close()
        for k,v in self._vm._settings.items():
            if v:
                try:
                    logging.debug("Slide clean-up: %s", k)
                    getattr(self._model, k)(self._context)
                except:
                    logging.exception("error in slide clean-up")
