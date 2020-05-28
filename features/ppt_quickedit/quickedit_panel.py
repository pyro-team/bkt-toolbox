# -*- coding: utf-8 -*-
'''
Created on 2018-05-29
@author: Florian Stallmann
'''

from __future__ import absolute_import, division

import logging
# import os.path

from time import time

from System.Collections.ObjectModel import ObservableCollection
from System.Windows.Controls import Orientation
from System.Windows.Media import Colors, SolidColorBrush
from System.Windows import Visibility

import bkt.ui
notify_property = bkt.ui.notify_property

from bkt.helpers import BitwiseValueAccessor
from bkt.callbacks import WpfActionCallback
from .quickedit_model import QuickEdit, QEColorButton, QEColorButtons, QECatalog


# class ColorButton(bkt.ui.NotifyPropertyChangedBase):
#     def __init__(self, index, color):
#         self._qecolor = color
#         self._Tag = index
#         super(ColorButton, self).__init__()

#     @property
#     def Tag(self):
#         return self._Tag

#     @property
#     def Label(self):
#         return self._qecolor.get_label()
    
#     @property
#     def Color(self):
#         return self._qecolor.get_color_html()
    
#     @property
#     def Checked(self):
#         return self._qecolor.get_checked()
#     @Checked.setter
#     def Checked(self, value):
#         # enforce onPropertyChange to ensure correct checked state
#         self.OnPropertyChanged("Checked")


class ViewModel(bkt.ui.ViewModelSingleton):
    def __init__(self, orientation_mode, window_left, window_top, docking_edge):
        super(ViewModel, self).__init__()
        self.init_buttons()
        self._orientation_mode = orientation_mode
        self._left = window_left
        self._top  = window_top
        self._docking_edge = docking_edge
        
        self._viewstate = BitwiseValueAccessor(settings_key="quickedit.viewstate", attributes=["recent_hidden", "docking_to_slide", "dark_theme"])

        self._editmode  = False

        self._brush_dark = SolidColorBrush(Colors.DimGray)
        self._brush_light = SolidColorBrush(Colors.WhiteSmoke)
        
        self._catalogs = ObservableCollection[QECatalog]()
        for cat in QuickEdit._catalogs:
            self._catalogs.Add(cat)

        # self.image_pickup  = bkt.ui.load_bitmapimage("qe_pickup")
        # self.image_nocolor = bkt.ui.load_bitmapimage("qe_nocolor")
        # self.image_edit    = bkt.ui.load_bitmapimage("qe_edit")
    
    def init_buttons(self):
        self._colors_theme = ObservableCollection[QEColorButton]()
        for color in QuickEdit._colors:
            self._colors_theme.Add(color)

        self._colors_recent = ObservableCollection[QEColorButton]()
        for color in QuickEdit._recent:
            self._colors_recent.Add(color)

        self._colors_own = ObservableCollection[QEColorButton]()
        for color in QuickEdit._userdefined:
            self._colors_own.Add(color)

        # self._colors_theme = ObservableCollection[ColorButton]()
        # for i,color in enumerate(model._colors):
        #     self._colors_theme.Add(ColorButton(i, color))

        # self._colors_recent = ObservableCollection[ColorButton]()
        # for i,color in enumerate(model._recent):
        #     self._colors_recent.Add(ColorButton(i, color))

        # self._colors_own = ObservableCollection[ColorButton]()
        # for i,color in enumerate(model._userdefined):
        #     self._colors_own.Add(ColorButton(i, color))

    def update_buttons(self):
        for btn in self._colors_theme:
            btn.OnPropertyChanged("Color")
            btn.OnPropertyChanged("Checked")
        for btn in self._colors_recent:
            btn.OnPropertyChanged("Color")
            btn.OnPropertyChanged("Checked")
        for btn in self._colors_own:
            btn.OnPropertyChanged("Color")
            btn.OnPropertyChanged("Checked")

    @notify_property
    def docking_edge(self):
        return self._docking_edge
    @docking_edge.setter
    def docking_edge(self, value):
        self._docking_edge = value

    @notify_property
    def colors_theme(self):
        return self._colors_theme
    @colors_theme.setter
    def colors_theme(self, value):
        self._colors_theme = value
    
    @notify_property
    def colors_recent(self):
        return self._colors_recent
    @colors_recent.setter
    def colors_recent(self, value):
        self._colors_recent = value
    
    @notify_property
    def colors_own(self):
        return self._colors_own
    @colors_own.setter
    def colors_own(self, value):
        self._colors_own = value
    
    @notify_property
    def current_orientation(self):
        return "%s/4" % (self._orientation_mode+1)
    
    @notify_property
    def orientation_mode(self):
        return self._orientation_mode
    @orientation_mode.setter
    def orientation_mode(self, value):
        self._orientation_mode = value
        self.OnPropertyChanged("current_orientation")
        self.OnPropertyChanged("outer_orientation")
        self.OnPropertyChanged("inner_orientation")
    
    @notify_property
    def outer_orientation(self):
        if self._orientation_mode in [0,2]:
            return Orientation.Horizontal
        else:
            return Orientation.Vertical
    
    @notify_property
    def inner_orientation(self):
        if self._orientation_mode in [0,1]:
            return Orientation.Horizontal
        else:
            return Orientation.Vertical
    
    @notify_property
    def recent_visibility(self):
        if self.recent_visible:
            return Visibility.Visible
        else:
            return Visibility.Collapsed
    
    @notify_property
    def color_background(self):
        if self.dark_theme:
            return self._brush_dark
        else:
            return self._brush_light
    
    @notify_property
    def color_foreground(self):
        if self.dark_theme:
            return self._brush_light
        else:
            return self._brush_dark

    @notify_property
    def recent_visible(self):
        return not self._viewstate.recent_hidden
    @recent_visible.setter
    def recent_visible(self, value):
        self._viewstate.recent_hidden = not value
        self.OnPropertyChanged("recent_visibility")

    @notify_property
    def docking_to_slide(self):
        return self._viewstate.docking_to_slide
    @docking_to_slide.setter
    def docking_to_slide(self, value):
        self._viewstate.docking_to_slide = value
        # self.OnPropertyChanged("window_left")
        # self.OnPropertyChanged("window_top")

    @notify_property
    def dark_theme(self):
        return self._viewstate.dark_theme
    @dark_theme.setter
    def dark_theme(self, value):
        self._viewstate.dark_theme = value
        self.OnPropertyChanged("color_background")
        self.OnPropertyChanged("color_foreground")

    @notify_property
    def window_left(self):
        return self._left
    @window_left.setter
    def window_left(self, value):
        self._left = value

    @notify_property
    def window_top(self):
        return self._top
    @window_top.setter
    def window_top(self, value):
        self._top = value

    @notify_property
    def editmode(self):
        return self._editmode
    @editmode.setter
    def editmode(self, value):
        self._editmode = value

    @notify_property
    def auto_start(self):
        return bkt.settings.get("quickedit.restore_panel", False)
    @auto_start.setter
    def auto_start(self, value):
        bkt.settings["quickedit.restore_panel"] = value
    
    # @notify_property
    # def change_orientation(self):
    #     return False
    #     # return self._orientation == Orientation.Horizontal
    # @change_orientation.setter
    # def change_orientation(self, value):
    #     self._orientation = Orientation.Horizontal if self._orientation == Orientation.Vertical else Orientation.Vertical
    #     self.OnPropertyChanged("window_orientation")


    @notify_property
    def catalogs(self):
        return self._catalogs
    @catalogs.setter
    def catalogs(self, value):
        self._catalogs = value
    
    # def update_files(self):
    #     for cat in self._catalogs:
    #         cat.OnPropertyChanged("Checked")


class QuickEditPanel(bkt.ui.WpfWindowAbstract):
    # _filename = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'quickedit_panel.xaml')
    _xamlname = 'quickedit_panel'
    # _vm_class = ViewModel

    def __init__(self, context):
        # self._model = model
        # self._context = context
        
        self.IsToolbar = True

        QuickEdit.update_colors(context)

        self._vm = ViewModel( 
            context.settings.get("quickedit.orientation_mode", 0),
            context.settings.get("quickedit.window_left", 300),
            context.settings.get("quickedit.window_top", 300),
            context.settings.get("quickedit.docking_edge", None),
            # context.settings.get("quickedit.viewstate", 0)
        )

        # first start detection
        if "quickedit.viewstate" not in context.settings:
            if bkt.message.confirmation("Dies scheint dein erster Start von QuickEdit zu sein. Soll die Anleitung (PDF) geöffnet werden?"):
                QuickEdit.show_help()

        self._last_catalog_change = 0

        super(QuickEditPanel, self).__init__(context)

    def _get_color(self, button):
        return QEColorButtons.get(button.Tag)


    def cancel(self, sender, event):
        self.Close()

    def determine_docking(self, sender=None, event=None):
        try:
            window = self._context.app.ActiveWindow
            if window.WindowState == 3 and window.ViewType == 9: #docking only if ppWindowMaximized and ppViewNormal
                window.Panes(2).Activate() #active ppViewSlide pane
                #remove current setting to trigger determine edge
                self._vm.docking_edge = None
                self.update_docking(window=window)
            else:
                bkt.message("Docking funktioniert nur bei maximiertem Fenster und bei normaler Folienansicht!")
        except:
            logging.exception("QUICKEDIT DOCKINGERROR")


    def update_docking(self, sender=None, event=None, window=None):
        try:
            if not self._vm.docking_to_slide:
                return
            if not window:
                window = self._context.app.ActiveWindow
            if window.WindowState == 3 and window.ViewType == 9 and window.ActivePane.ViewType == 1: #ppWindowMaximized and ppViewNormal and ppViewSlide
                slidem = window.View.Slide.Master
                left, top = window.PointsToScreenPixelsX(0), window.PointsToScreenPixelsY(0)
                right, bottom = window.PointsToScreenPixelsX(slidem.Width), window.PointsToScreenPixelsY(slidem.Height)

                if self._vm.docking_edge is None:
                    #determine edge
                    rect = self.GetDeviceRect()
                    mid_rect = (rect.Left + rect.Width/2, rect.Top + rect.Height/2)
                    mid_slide = (left + (right-left)/2, top + (bottom-top)/2)
                    vec = (mid_rect[0] - mid_slide[0], mid_rect[1] - mid_slide[1])
                    if vec[0] > 0 and vec[1] > 0:
                        self._vm.docking_edge = 4
                    elif vec[0] < 0 and vec[1] > 0:
                        self._vm.docking_edge = 3
                    elif vec[0] > 0 and vec[1] < 0:
                        self._vm.docking_edge = 2
                    else: #if vec[0] < 0 and vec[1] < 0:
                        self._vm.docking_edge = 1
                
                if self._vm.docking_edge == 1:
                    ###top-left docking
                    if self._vm.orientation_mode <= 1: #horizontal
                        self.SetDevicePosition(deviceLeft=left, deviceBottom=top-5)
                    else:
                        self.SetDevicePosition(deviceRight=left-5, deviceTop=top)
                elif self._vm.docking_edge == 2:
                    ###top-right docking
                    if self._vm.orientation_mode <= 1: #horizontal
                        self.SetDevicePosition(deviceRight=right, deviceBottom=top-5)
                    else:
                        self.SetDevicePosition(deviceLeft=right+5, deviceTop=top)
                elif self._vm.docking_edge == 3:
                    ###bottom-left docking
                    if self._vm.orientation_mode <= 1: #horizontal
                        self.SetDevicePosition(deviceLeft=left, deviceTop=bottom+5)
                    else:
                        self.SetDevicePosition(deviceRight=left-5, deviceBottom=bottom)
                else:
                    ###bottom-right docking
                    if self._vm.orientation_mode <= 1: #horizontal
                        self.SetDevicePosition(deviceRight=right, deviceTop=bottom+5)
                    else:
                        self.SetDevicePosition(deviceLeft=right+5, deviceBottom=bottom)

        except:
            #PointsToScreenPixelsX illegal value if ActivePane != 1 (e.g. slide thumbnails selected)
            logging.exception("QUICKEDIT DOCKINGERROR")
    
    def change_orientation(self, sender, event):
        self._vm.orientation_mode = (self._vm.orientation_mode+1) %4
        # self.Width, self.Height = self.Height, self.Width
    
    def change_file(self, sender, event):
        try:
            file = event.OriginalSource.Tag
            QuickEdit.read_from_config(file)
            QuickEdit.update_colors(self._context)
            QuickEdit.update_pressed(self._context)
        except:
            pass
    
    def Catalog_Wheel(self, sender, event):
        try:
            if time() - self._last_catalog_change < 0.2: #change catalog only after 200ms
                raise RuntimeError("too many scroll activities")
            current_index = [c.Checked for c in QuickEdit._catalogs].index(True)
            direction = 1 if event.Delta < 0 else -1
            next_index = (current_index+direction) % len(QuickEdit._catalogs)
            QuickEdit.read_from_config(QuickEdit._catalogs[next_index].File)
            QuickEdit.update_colors(self._context)
            QuickEdit.update_pressed(self._context)
            self._last_catalog_change = time()
        except:
            pass

    @WpfActionCallback
    def ForceReload(self, sender, event):
        QuickEdit.update_colors(self._context)
        self._vm.update_buttons()
        #FIXME: I think we can delete this

    @WpfActionCallback
    def ColorThemeButton(self, sender, event):
        # print("button theme clicked: %s" % sender.Tag)
        QuickEdit.action(self._get_color(sender), self._context)

    @WpfActionCallback
    def ColorRecentButton(self, sender, event):
        # print("button recent clicked: %s" % sender.Tag)
        # QuickEdit.action(QuickEdit._recent[int(sender.Tag)], self._context)
        QuickEdit.action(self._get_color(sender), self._context)

    @WpfActionCallback
    def ColorOwnButton(self, sender, event):
        # print("button own clicked: %s" % sender.Tag)
        if self._vm.editmode:
            QuickEdit.pickup_own_color(self._get_color(sender), self._context)
        else:
            QuickEdit.action(self._get_color(sender), self._context)

    @WpfActionCallback
    def ColorNone(self, sender, event):
        QuickEdit.action_no_fill(self._context)
        
    @WpfActionCallback
    def ColorNone_Wheel(self, sender, event):
        # print("mouse wheel %s" % event.Delta)
        delta = 0.1 if event.Delta < 0 else -0.1
        QuickEdit.action_transparency(self._context, delta)
        # sender.ToolTip.IsOpen = True

    @WpfActionCallback
    def PickupRecent(self, sender, event):
        QuickEdit.pickup_recent_color(self._context)

    # @WpfActionCallback
    # def PickupOwn(self, sender, event):
    #     QuickEdit.pickup_own_color(QuickEdit._userdefined[0], self._context)

    @WpfActionCallback
    def ResetOwnColors(self, sender, event):
        QuickEdit.reset_own_colors()
    

    def ShowHelp(self, sender, event):
        QuickEdit.show_help()

    def Window_MouseLeftButtonDown(self, sender, event):
        # self._vm.docking_to_slide = False
        if event.ClickCount >= 2:
            self.determine_docking()
        else:
            self.DragMove()

    def Window_Closing(self, sender, event):
        # print("window closing")
        self._context.settings["quickedit.orientation_mode"] = self._vm.orientation_mode
        self._context.settings["quickedit.window_left"] = self._vm.window_left
        self._context.settings["quickedit.window_top"] = self._vm.window_top
        self._context.settings["quickedit.docking_edge"] = self._vm.docking_edge
    
    # def show_dialog(self, modal=True):
    #     #TODO: Save and restore position and size of window
    #     left, top = self._context.app.activewindow.PointsToScreenPixelsX(0), self._context.app.activewindow.PointsToScreenPixelsY(0)
    #     self.SetDevicePosition(left, top)

    #     return self.Show()