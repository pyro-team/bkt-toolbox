# -*- coding: utf-8 -*-
'''
Created on 2018-05-29
@author: Florian Stallmann
'''

import os.path
import bkt.ui
notify_property = bkt.ui.notify_property


class ViewModel(bkt.ui.ViewModelAsbtract):
    def __init__(self, model):
        super(ViewModel, self).__init__()
        
        self.settings = model.settings
        self.settings_all = all(self.settings.values())
    
    @notify_property
    def settings_fill(self):
        return self.settings["fill"]
    @settings_fill.setter
    def settings_fill(self, value):
        self.settings["fill"] = value
        self.OnPropertyChanged('settings_all')
    
    @notify_property
    def settings_type(self):
        return self.settings["type"]
    @settings_type.setter
    def settings_type(self, value):
        self.settings["type"] = value
        self.OnPropertyChanged('settings_all')
    
    @notify_property
    def settings_line(self):
        return self.settings["line"]
    @settings_line.setter
    def settings_line(self, value):
        self.settings["line"] = value
        self.OnPropertyChanged('settings_all')
    
    @notify_property
    def settings_textframe2(self):
        return self.settings["textframe2"]
    @settings_textframe2.setter
    def settings_textframe2(self, value):
        self.settings["textframe2"] = value
        self.OnPropertyChanged('settings_all')
    
    @notify_property
    def settings_paragraphformat(self):
        return self.settings["paragraphformat"]
    @settings_paragraphformat.setter
    def settings_paragraphformat(self, value):
        self.settings["paragraphformat"] = value
        self.OnPropertyChanged('settings_all')
    
    @notify_property
    def settings_font(self):
        return self.settings["font"]
    @settings_font.setter
    def settings_font(self, value):
        self.settings["font"] = value
        self.OnPropertyChanged('settings_all')
    
    @notify_property
    def settings_shadow(self):
        return self.settings["shadow"]
    @settings_shadow.setter
    def settings_shadow(self, value):
        self.settings["shadow"] = value
        self.OnPropertyChanged('settings_all')
    
    @notify_property
    def settings_size(self):
        return self.settings["size"]
    @settings_size.setter
    def settings_size(self, value):
        self.settings["size"] = value
        self.OnPropertyChanged('settings_all')
    
    @notify_property
    def settings_position(self):
        return self.settings["position"]
    @settings_position.setter
    def settings_position(self, value):
        self.settings["position"] = value
        self.OnPropertyChanged('settings_all')


    @notify_property
    def settings_all(self):
        return all(self.settings.values())
    
    @settings_all.setter
    def settings_all(self, value):
        for k in self.settings:
            self.settings[k] = value
        self.OnPropertyChanged('settings_fill')
        self.OnPropertyChanged('settings_type')
        self.OnPropertyChanged('settings_line')
        self.OnPropertyChanged('settings_textframe2')
        self.OnPropertyChanged('settings_paragraphformat')
        self.OnPropertyChanged('settings_font')
        self.OnPropertyChanged('settings_shadow')
        self.OnPropertyChanged('settings_size')
        self.OnPropertyChanged('settings_position')


class ApplyWindow(bkt.ui.WpfWindowAbstract):
    _filename = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'apply_style.xaml')
    # _vm_class = ViewModel

    def __init__(self, model, index, context):
        self.settings = model.custom_styles[index]["style_settings"].copy()
        self.index = index
        self.context = context

        self._model = model
        self._vm = ViewModel(self)

        super(ApplyWindow, self).__init__()

    def cancel(self, sender, event):
        self.Close()
    
    def apply_style(self, sender, event):
        self.Close()
        self._model.custom_styles[self.index]["style_settings"] = self.settings
        self._model.save_to_config()
        self._model.apply_custom_style(self.index, self.context, self.settings)