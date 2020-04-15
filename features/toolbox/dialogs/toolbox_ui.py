# -*- coding: utf-8 -*-
'''
Created on 2018-05-29
@author: Florian Stallmann
'''

from __future__ import absolute_import

import sys
import os.path
import logging

import System

import bkt
import bkt.ui
notify_property = bkt.ui.notify_property


class ViewModel(bkt.ui.ViewModelAsbtract):
    def __init__(self, toolboxui):
        super(ViewModel, self).__init__()

        # settings = bkt.settings.get("toolboxui.settings", {})
        s_dict = {'0': False, '1': False, '2': False}
        h_dict = {'0': "", '1': "", '2': ""}
        
        self._resource_path = os.path.normpath(os.path.join(os.path.dirname(os.path.realpath(__file__)), "..", "resources", "toolboxui"))

        for key in toolboxui.get_all_keys():
            bool_dict = s_dict.copy()
            bool_dict[str(toolboxui.get_setting(key))] = True

            header_dict = h_dict.copy()
            header_dict[str(toolboxui.get_theme_setting(key))] = "*"

            img_path = System.Uri(os.path.join(self._resource_path, key+".png"))

            setattr(self, key, bool_dict)
            setattr(self, key+"_header", header_dict)
            setattr(self, key+"_url", img_path)

        # self.size_group = s_dict.copy()
        # self.size_group[str(toolboxui.get_setting('size_group'))] = True
        # self.size_group_header[str(toolboxui.get_theme_setting('size_group'))] = "*"
        # self.size_group_url = System.Uri(os.path.join(self._resource_path, "size_group.png"))


class ToolboxUiWindow(bkt.ui.WpfWindowAbstract):
    _xamlname = 'toolbox_ui'
    # _vm_class = ViewModel

    def __init__(self, model, context):
        self._model = model
        self._vm = ViewModel(model)

        super(ToolboxUiWindow, self).__init__(context)

    def cancel(self, sender, event):
        self.Close()
    
    def _value2key(self, sdir):
        if sdir['0']:
            return 0
        elif sdir['1']:
            return 1
        elif sdir['2']:
            return 2
        raise ValueError('invalid settings value')

    def reset_settings(self, sender, event):
        self._model.reset_to_defaults()

        self.Close()
        if bkt.helpers.confirmation("Soll die BKT nun neu geladen werden?"):
            self._reload_bkt()

    def save_settings(self, sender, event):
        self._model.set_setting("size_group", self._value2key(self._vm.size_group))
        self._model.set_setting("pos_size_group", self._value2key(self._vm.pos_size_group))

        self._model.set_setting("arrange_mini_group", self._value2key(self._vm.arrange_mini_group))
        self._model.set_setting("arrange_euclid_group", self._value2key(self._vm.arrange_euclid_group))
        self._model.set_setting("arrange_adv_group", self._value2key(self._vm.arrange_adv_group))
        self._model.set_setting("arrange_adv_easy_group", self._value2key(self._vm.arrange_adv_easy_group))
        self._model.set_setting("arrange_dist_rota_group", self._value2key(self._vm.arrange_dist_rota_group))

        self._model.set_setting("text_padding_group", self._value2key(self._vm.text_padding_group))
        self._model.set_setting("text_par_group", self._value2key(self._vm.text_par_group))
        self._model.set_setting("text_parindent_group", self._value2key(self._vm.text_parindent_group))

        self._model.set_setting("adjustments_group", self._value2key(self._vm.adjustments_group))
        self._model.set_setting("format_group", self._value2key(self._vm.format_group))
        self._model.set_setting("split_group", self._value2key(self._vm.split_group))
        self._model.set_setting("language_group", self._value2key(self._vm.language_group))
        self._model.set_setting("stateshape_group", self._value2key(self._vm.stateshape_group))
        self._model.set_setting("iconsearch_group", self._value2key(self._vm.iconsearch_group))

        self.Close()
        if bkt.helpers.confirmation("Soll die BKT nun neu geladen werden?"):
            self._reload_bkt()
    
    def _reload_bkt(self):
        from modules.settings import BKTReload
        BKTReload.reload_bkt(self._context)
