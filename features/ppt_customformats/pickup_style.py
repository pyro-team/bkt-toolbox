# -*- coding: utf-8 -*-
'''
Created on 2018-05-29
@author: Florian Stallmann
'''

import os.path
import bkt.ui
notify_property = bkt.ui.notify_property


class ViewModel(bkt.ui.ViewModelAsbtract):
    def __init__(self, model, buttonindex):
        super(ViewModel, self).__init__()
        
        self._model = model

        self.settings = self._model.style_settings
        self.buttonindex = buttonindex
    
    @notify_property
    def settings(self):
        return self._settings
    
    @settings.setter
    def settings(self, value):
        self._settings = value
    

    @notify_property
    def buttonindexA(self):
        return self.buttonindex == 0
    @buttonindexA.setter
    def buttonindexA(self, value):
        if value:
            self.buttonindex = 0
    
    @notify_property
    def buttonindexB(self):
        return self.buttonindex == 1
    @buttonindexB.setter
    def buttonindexB(self, value):
        if value:
            self.buttonindex = 1
    
    @notify_property
    def buttonindexC(self):
        return self.buttonindex == 2
    @buttonindexC.setter
    def buttonindexC(self, value):
        if value:
            self.buttonindex = 2
    
    @notify_property
    def buttonindexD(self):
        return self.buttonindex == 3
    @buttonindexD.setter
    def buttonindexD(self, value):
        if value:
            self.buttonindex = 3
    
    @notify_property
    def buttonindexE(self):
        return self.buttonindex == 4
    @buttonindexE.setter
    def buttonindexE(self, value):
        if value:
            self.buttonindex = 4

    
    @notify_property
    def imgButtonA(self):
        return bkt.ui.convert_bitmap_to_bitmapsource(self._model.get_image_by_index(0, 32))
    @notify_property
    def imgButtonB(self):
        return bkt.ui.convert_bitmap_to_bitmapsource(self._model.get_image_by_index(1, 32))
    @notify_property
    def imgButtonC(self):
        return bkt.ui.convert_bitmap_to_bitmapsource(self._model.get_image_by_index(2, 32))
    @notify_property
    def imgButtonD(self):
        return bkt.ui.convert_bitmap_to_bitmapsource(self._model.get_image_by_index(3, 32))
    @notify_property
    def imgButtonE(self):
        return bkt.ui.convert_bitmap_to_bitmapsource(self._model.get_image_by_index(4, 32))


class PickupWindow(bkt.ui.WpfWindowAbstract):
    _filename = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'pickup_style.xaml')
    # _vm_class = ViewModel

    def __init__(self, model, shape, buttonindex=None):
        self._model = model
        if buttonindex is None:
            try:
                buttonindex = self._model.custom_styles.index(None)
            except ValueError:
                buttonindex = 0
        self._vm = ViewModel(model, buttonindex)
        self.shape = shape

        super(PickupWindow, self).__init__()

    def cancel(self, sender, event):
        self.Close()
    
    def pickup_style(self, sender, event):
        self.Close()
        self._model.pickup_custom_style(self._vm.buttonindex, self.shape)