# -*- coding: utf-8 -*-
'''
Created on 2018-05-29
@author: Florian Stallmann
'''

from __future__ import absolute_import, division

import logging

from collections import deque
from System.Windows.Media import Colors, SolidColorBrush

import bkt.ui
notify_property = bkt.ui.notify_property


class ViewModel(bkt.ui.ViewModelSingleton):
    def __init__(self, window_left, window_top):
        super(ViewModel, self).__init__()
        self._left = window_left
        self._top  = window_top

        self._brush_dark = SolidColorBrush(Colors.DimGray)
        self._brush_light = SolidColorBrush(Colors.WhiteSmoke)
    

    @notify_property
    def dark_theme(self):
        return False
    # @dark_theme.setter
    # def dark_theme(self, value):
    #     self._viewstate.dark_theme = value
    #     self.OnPropertyChanged("color_background")
    #     self.OnPropertyChanged("color_foreground")

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

class SlideNavigator(bkt.ui.WpfWindowAbstract):
    _xamlname = 'slide_navigator'

    def __init__(self, context):
        self.IsToolbar = True

        self._vm = ViewModel(
            context.settings.get("slide_navigator.window_left", 300),
            context.settings.get("slide_navigator.window_top", 300),
            )

        super(SlideNavigator, self).__init__(context)
    
    def go_backward(self, sender, event):
        logging.info("go back")
        try:
            if len(SlideNavigator.slide_history1) > 1:
                SlideNavigator.slide_history2.appendleft(SlideNavigator.slide_history1.pop())
                sld = SlideNavigator.slide_history1.pop()
                self._context.app.ActiveWindow.View.GotoSlide(sld)
        except:
            logging.exception("error going back")
    
    def go_forward(self, sender, event):
        logging.info("go forth")
        try:
            if SlideNavigator.slide_history2:
                # SlideNavigator.slide_history1.append(SlideNavigator.slide_history2.popleft())
                sld = SlideNavigator.slide_history2.pop()
                self._context.app.ActiveWindow.View.GotoSlide(sld)
        except:
            logging.exception("error going forth")


    def Window_MouseLeftButtonDown(self, sender, event):
        self.DragMove()

    def Window_Closing(self, sender, event):
        self._context.settings["slide_navigator.window_left"] = self._vm.window_left
        self._context.settings["slide_navigator.window_top"] = self._vm.window_top


    navigation_ongoing = False
    slide_history1 = deque(maxlen=20)
    slide_history2 = deque()

    @classmethod
    def slide_change(cls, slide_range):
        logging.info("slide changed")
        if slide_range.Count == 1:
            SlideNavigator.slide_history1.append(slide_range.SlideIndex)


class SlideNavigatorManager(object):
    panel_windows = {}

    @staticmethod
    def _create_panel(context):
        return SlideNavigator(context)

    @classmethod
    def get_panel_for_active_window(cls, context):
        logging.debug("get panel for active window")
        windowid = context.addin.GetWindowHandle()
        if windowid in cls.panel_windows:
            return cls.panel_windows[windowid]
        else:
            return None

    @classmethod
    def show_panel_for_active_window(cls, context):
        logging.debug("show panel for active window")

        windowid = context.addin.GetWindowHandle()
        if windowid in cls.panel_windows:
            if cls.panel_windows[windowid].IsLoaded:
                #ensure that window is on the screen
                cls.panel_windows[windowid].ShiftWindowOntoScreen()
                return #active panel window already exists
            else:
                cls._close_panel(windowid)

        cls._show_panel(context, windowid)

    @classmethod
    def close_panel_for_active_window(cls, context, presentation):
        logging.debug("close panel for active window")
        if cls._is_windowed_presentation(context, presentation):
            windowid = context.addin.GetWindowHandle()
            cls._close_panel(windowid)

    @classmethod
    def _show_panel(cls, context, windowid):
        logging.debug("show panel for window %s", windowid)
        try:
            panel = cls._create_panel(context)
            panel.SetOwner(windowid)
            panel.Show()
            panel.ShiftWindowOntoScreen() #ensure that window is on the screen
            cls.panel_windows[windowid] = panel
        except:
            logging.exception("panel activation failed")
            if bkt.config.show_exception:
                bkt.helpers.exception_as_message()
            else:
                bkt.message("Unbekannter Fehler beim Anzeigen des Panels!")

    @classmethod
    def _close_panel(cls, windowid):
        logging.debug("close panel for window %s", windowid)
        try:
            cls.panel_windows[windowid].Close()
            del cls.panel_windows[windowid]
        except:
            pass
    
    @staticmethod
    def _is_windowed_presentation(context, presentation):
        try:
            #only show if at least one window exists
            return presentation.Windows.Count > 0
            #ALTERNATIVE: only show if opened presentation equals active presentation (not the case if opened without window)
            # return presentation.FullName == context.presentation.FullName
        except: #COMException
            return False

    @classmethod
    def close_all_panels(cls):
        logging.debug("close all panels")
        for windowid in cls.panel_windows.keys():
            cls._close_panel(windowid)


bkt.AppEvents.presentation_close += bkt.Callback(SlideNavigatorManager.close_panel_for_active_window, context=True)
bkt.AppEvents.bkt_unload         += bkt.Callback(SlideNavigatorManager.close_all_panels)

bkt.AppEvents.slide_selection_changed += bkt.Callback(SlideNavigator.slide_change)