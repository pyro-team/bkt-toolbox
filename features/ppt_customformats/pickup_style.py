# -*- coding: utf-8 -*-
'''
Created on 2018-05-29
@author: Florian Stallmann
'''

import os.path
import bkt.ui
notify_property = bkt.ui.notify_property

from System.Windows import Visibility
from collections import namedtuple

class ViewModel(bkt.ui.ViewModelAsbtract):
    def __init__(self, settings, mode, name=None):
        super(ViewModel, self).__init__()

        self.settings = settings
        # self.settings_all = all(self.settings.values())
        
        self.show_delete = Visibility.Collapsed
        if mode == "edit" and name is not None:
            self.title_text = "Style {} bearbeiten".format(name)
            self.show_delete = Visibility.Visible
        elif mode == "new":
            self.title_text = "Neuen Style anlegen"
        else: #apply
            self.title_text = "Style anwenden"


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
        self.OnPropertyChanged('settings_textframe')
        self.OnPropertyChanged('settings_paragraphformat')
        self.OnPropertyChanged('settings_font')
        self.OnPropertyChanged('settings_shadow')
        self.OnPropertyChanged('settings_size')
        self.OnPropertyChanged('settings_position')


    @notify_property
    def settings_fill(self):
        return self.settings["Fill"]
    @settings_fill.setter
    def settings_fill(self, value):
        self.settings["Fill"] = value
        self.OnPropertyChanged('settings_all')
    
    @notify_property
    def settings_type(self):
        return self.settings["Type"]
    @settings_type.setter
    def settings_type(self, value):
        self.settings["Type"] = value
        self.OnPropertyChanged('settings_all')
    
    @notify_property
    def settings_line(self):
        return self.settings["Line"]
    @settings_line.setter
    def settings_line(self, value):
        self.settings["Line"] = value
        self.OnPropertyChanged('settings_all')
    
    @notify_property
    def settings_textframe(self):
        return self.settings["TextFrame"]
    @settings_textframe.setter
    def settings_textframe(self, value):
        self.settings["TextFrame"] = value
        self.OnPropertyChanged('settings_all')
    
    @notify_property
    def settings_paragraphformat(self):
        return self.settings["ParagraphFormat"]
    @settings_paragraphformat.setter
    def settings_paragraphformat(self, value):
        self.settings["ParagraphFormat"] = value
        self.OnPropertyChanged('settings_all')
    
    @notify_property
    def settings_font(self):
        return self.settings["Font"]
    @settings_font.setter
    def settings_font(self, value):
        self.settings["Font"] = value
        self.OnPropertyChanged('settings_all')
    
    @notify_property
    def settings_shadow(self):
        return self.settings["Shadow"]
    @settings_shadow.setter
    def settings_shadow(self, value):
        self.settings["Shadow"] = value
        self.settings["Glow"] = value
        self.settings["SoftEdge"] = value
        self.settings["Reflection"] = value
        self.OnPropertyChanged('settings_all')
    
    @notify_property
    def settings_size(self):
        return self.settings["Size"]
    @settings_size.setter
    def settings_size(self, value):
        self.settings["Size"] = value
        self.OnPropertyChanged('settings_all')
    
    @notify_property
    def settings_position(self):
        return self.settings["Position"]
    @settings_position.setter
    def settings_position(self, value):
        self.settings["Position"] = value
        self.OnPropertyChanged('settings_all')




class PickupWindow(bkt.ui.WpfWindowAbstract):
    _filename = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'pickup_style.xaml')
    # _vm_class = ViewModel

    def __init__(self, model, style_setting, shape=None, index=None): #modes: new, edit, apply
        self._model = model

        self.shape = shape
        self.index = index
        self.result = None
        
        if shape:
            #new
            self._vm = ViewModel(style_setting.copy(), "new")
        elif index is not None:
            #edit
            self._vm = ViewModel(style_setting.copy(), "edit", model.get_custom_style_name(index))
        else:
            #apply
            self._vm = ViewModel(style_setting.copy(), "apply")


        # name = None if index is None else index+1

        super(PickupWindow, self).__init__()

    def cancel(self, sender, event):
        self.Close()

    def pickup_style(self, sender, event):
        self.Close()
        if self.shape:
            #new: call pickup method
            self._model.pickup_custom_style(self.shape, self._vm.settings)
            self.result = True
        elif self.index is not None:
            #edit: call edit method
            self._model.edit_custom_style(self.index, self._vm.settings)
            self.result = True
        else:
            #apply: only save settings to result
            self.result = self._vm.settings
    
    def delete_style(self, sender, event):
        self.Close()
        if self.index is not None:
            self._model.delete_custom_style(self.index)