# -*- coding: utf-8 -*-
'''
Created on 2018-05-29
@author: Florian Stallmann
'''

import os.path
import bkt.ui
notify_property = bkt.ui.notify_property

from System.Windows import Visibility


class ViewModel(bkt.ui.ViewModelAsbtract):
    def __init__(self, settings, name=None):
        super(ViewModel, self).__init__()

        self.settings = settings
        if name is not None:
            self.title_text = "Style {} bearbeiten".format(name)
            self.show_delete = Visibility.Visible
        else:
            self.title_text = "Neuen Style anlegen"
            self.show_delete = Visibility.Collapsed
    
    @notify_property
    def settings_fill(self):
        return self.settings["Fill"]
    @settings_fill.setter
    def settings_fill(self, value):
        self.settings["Fill"] = value
    
    @notify_property
    def settings_type(self):
        return self.settings["Type"]
    @settings_type.setter
    def settings_type(self, value):
        self.settings["Type"] = value
    
    @notify_property
    def settings_line(self):
        return self.settings["Line"]
    @settings_line.setter
    def settings_line(self, value):
        self.settings["Line"] = value
    
    @notify_property
    def settings_textframe(self):
        return self.settings["TextFrame"]
    @settings_textframe.setter
    def settings_textframe(self, value):
        self.settings["TextFrame"] = value
    
    @notify_property
    def settings_paragraphformat(self):
        return self.settings["ParagraphFormat"]
    @settings_paragraphformat.setter
    def settings_paragraphformat(self, value):
        self.settings["ParagraphFormat"] = value
    
    @notify_property
    def settings_font(self):
        return self.settings["Font"]
    @settings_font.setter
    def settings_font(self, value):
        self.settings["Font"] = value
    
    @notify_property
    def settings_shadow(self):
        return self.settings["Shadow"]
    @settings_shadow.setter
    def settings_shadow(self, value):
        self.settings["Shadow"] = value
        self.settings["Glow"] = value
        self.settings["SoftEdge"] = value
        self.settings["Reflection"] = value
    
    @notify_property
    def settings_size(self):
        return self.settings["Size"]
    @settings_size.setter
    def settings_size(self, value):
        self.settings["Size"] = value
    
    @notify_property
    def settings_position(self):
        return self.settings["Position"]
    @settings_position.setter
    def settings_position(self, value):
        self.settings["Position"] = value


class PickupWindow(bkt.ui.WpfWindowAbstract):
    _filename = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'pickup_style.xaml')
    # _vm_class = ViewModel

    def __init__(self, model, style_setting, shape=None, index=None):
        self._model = model
        self.shape = shape
        self.index = index
        name = None if index is None else index+1
        self._vm = ViewModel(style_setting.copy(), name)

        super(PickupWindow, self).__init__()

    def cancel(self, sender, event):
        self.Close()
    
    def pickup_style(self, sender, event):
        self.Close()
        if self.shape:
            #pickup
            self._model.pickup_custom_style(self.shape, self._vm.settings)
        elif self.index is not None:
            #edit
            self._model.edit_custom_style(self.index, self._vm.settings)
    
    def delete_style(self, sender, event):
        self.Close()
        if self.index is not None:
            self._model.delete_custom_style(self.index)