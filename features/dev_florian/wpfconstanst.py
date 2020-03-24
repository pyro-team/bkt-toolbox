# -*- coding: utf-8 -*-
'''
Created on 2017-07-24
@author: Florian Stallmann
'''

import win32con

constants = {
    "WS_OVERLAPPED":             win32con.WS_OVERLAPPED,
    "WS_POPUP":                 win32con.WS_POPUP,
    "WS_CHILD":                win32con.WS_CHILD,
    "WS_MINIMIZE":           win32con.WS_MINIMIZE,
    "WS_VISIBLE":            win32con.WS_VISIBLE,
    "WS_DISABLED":           win32con.WS_DISABLED,
    "WS_CLIPSIBLINGS":           win32con.WS_CLIPSIBLINGS,
    "WS_CLIPCHILDREN":           win32con.WS_CLIPCHILDREN,
    "WS_MAXIMIZE":           win32con.WS_MAXIMIZE,
    "WS_BORDER":             win32con.WS_BORDER,
    "WS_DLGFRAME":           win32con.WS_DLGFRAME,
    "WS_VSCROLL":            win32con.WS_VSCROLL,
    "WS_HSCROLL":            win32con.WS_HSCROLL,
    "WS_SYSMENU":            win32con.WS_SYSMENU,
    "WS_THICKFRAME":             win32con.WS_THICKFRAME,
    "WS_GROUP":          win32con.WS_GROUP,
    "WS_TABSTOP":            win32con.WS_TABSTOP,
    "WS_MINIMIZEBOX":            win32con.WS_MINIMIZEBOX,
    "WS_MAXIMIZEBOX":            win32con.WS_MAXIMIZEBOX,
    "WS_CAPTION":            win32con.WS_CAPTION,
    "WS_TILED":          win32con.WS_TILED,
    "WS_ICONIC":             win32con.WS_ICONIC,
    "WS_SIZEBOX":            win32con.WS_SIZEBOX,
    "WS_TILEDWINDOW":            win32con.WS_TILEDWINDOW,
    "WS_OVERLAPPEDWINDOW":           win32con.WS_OVERLAPPEDWINDOW,
    "WS_POPUPWINDOW":            win32con.WS_POPUPWINDOW,
    "WS_CHILDWINDOW":            win32con.WS_CHILDWINDOW,
}


def check(input):
    result = ""
    for key,value in constants.iteritems():
        if (input & value) == abs(value):
            result += key + "\r\n"
    return result

exconstants = {
    "WS_EX_DLGMODALFRAME":       win32con.WS_EX_DLGMODALFRAME,
    "WS_EX_NOPARENTNOTIFY":      win32con.WS_EX_NOPARENTNOTIFY,
    "WS_EX_TOPMOST":         win32con.WS_EX_TOPMOST,
    "WS_EX_ACCEPTFILES":         win32con.WS_EX_ACCEPTFILES,
    "WS_EX_TRANSPARENT":         win32con.WS_EX_TRANSPARENT,
    "WS_EX_MDICHILD":        win32con.WS_EX_MDICHILD,
    "WS_EX_TOOLWINDOW":      win32con.WS_EX_TOOLWINDOW,
    "WS_EX_WINDOWEDGE":      win32con.WS_EX_WINDOWEDGE,
    "WS_EX_CLIENTEDGE":      win32con.WS_EX_CLIENTEDGE,
    "WS_EX_CONTEXTHELP":         win32con.WS_EX_CONTEXTHELP,
    "WS_EX_RIGHT":       win32con.WS_EX_RIGHT,
    "WS_EX_LEFT":        win32con.WS_EX_LEFT,
    "WS_EX_RTLREADING":      win32con.WS_EX_RTLREADING,
    "WS_EX_LTRREADING":      win32con.WS_EX_LTRREADING,
    "WS_EX_LEFTSCROLLBAR":       win32con.WS_EX_LEFTSCROLLBAR,
    "WS_EX_RIGHTSCROLLBAR":      win32con.WS_EX_RIGHTSCROLLBAR,
    "WS_EX_CONTROLPARENT":       win32con.WS_EX_CONTROLPARENT,
    "WS_EX_STATICEDGE":      win32con.WS_EX_STATICEDGE,
    "WS_EX_APPWINDOW":       win32con.WS_EX_APPWINDOW,
    "WS_EX_LAYERED":         win32con.WS_EX_LAYERED,
    "WS_EX_COMPOSITED":      win32con.WS_EX_COMPOSITED,
    "WS_EX_NOACTIVATE":      win32con.WS_EX_NOACTIVATE,
    "WS_EX_NOINHERITLAYOUT":         win32con.WS_EX_NOINHERITLAYOUT,
    "WS_EX_NOPARENTNOTIFY":      win32con.WS_EX_NOPARENTNOTIFY,
    "WS_EX_OVERLAPPEDWINDOW":        win32con.WS_EX_OVERLAPPEDWINDOW,
    "WS_EX_PALETTEWINDOW":       win32con.WS_EX_PALETTEWINDOW,
}


def check2(input):
    result = ""
    for key,value in exconstants.iteritems():
        if (input & value) == abs(value):
            result += key + "\r\n"
    return result