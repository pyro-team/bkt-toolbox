# -*- coding: utf-8 -*-
'''
Created on 2017-07-24
@author: Florian Stallmann
'''

from __future__ import absolute_import

import os.path
import ctypes

# import clr
# clr.AddReference('PresentationCore')
# clr.AddReference("PresentationFramework")
# clr.AddReference("PresentationFramework.Aero")
# clr.AddReference('Microsoft.Dynamic')
# clr.AddReference('Microsoft.Scripting')
# clr.AddReference('System')
# clr.AddReference('IronPython')
# clr.AddReference('IronPython.Modules')
# clr.AddReference('IronPython.Wpf')

# from IronPython.Modules import Wpf as wpf

import System.Windows
Window = System.Windows.Window
StackPanel = System.Windows.Controls.StackPanel
Button = System.Windows.Controls.Button

import bkt
import bkt.dotnet
wpf = bkt.dotnet.import_wpf()


class MyWindow(Window):
    getColorR = System.Windows.Media.Brushes.Red

    def __init__(self):
        folder = os.path.dirname(os.path.realpath(__file__))
        wpf.LoadComponent( self, os.path.join(folder, 'MyWindow.xaml') )
        self.txtDescription.Text = "XXX"
        self.txtInput.Height=100
        self.getColorR = System.Windows.Media.Brushes.Red

    def btnClickMe_Click(self, sender, e):
        bkt.message("test")

    def btnClose_Click(self, sender, e):
        self.Close()

    @property
    def GetColorR(self):
        return self.getColorR

    @GetColorR.setter
    def GetColorR(self, value):
        self.getColorR = value

    # def GetColor(self, sender, e):
    #     print "im am here"

    # def OnSourceInitialized(self, event):
    #     # print("quick edit source initialized")
    #     import win32con as wc
    #     import wpfconstanst

    #     GWL_STYLE   = (-16)
    #     GWL_EXSTYLE = (-20)
    #     wnd_wih = System.Windows.Interop.WindowInteropHelper(self)
    #     style   = int(ctypes.windll.user32.GetWindowLongW(wnd_wih.Handle, GWL_STYLE))
    #     exstyle = int(ctypes.windll.user32.GetWindowLongW(wnd_wih.Handle, GWL_EXSTYLE))

    #     print("OLD Style: %s" % style)
    #     print(wpfconstanst.check(style))
    #     print("OLD Ex Style: %s" % exstyle)
    #     print(wpfconstanst.check2(exstyle))

    #     # style = style | ADD & ~REMOVE
    #     # exstyle = exstyle | ADD & ~REMOVE
    #     style = style | wc.WS_POPUP
    #     # exstyle = exstyle | wc.WS_EX_NOACTIVATE

    #     print("NEW Style: %s" % style)
    #     print(wpfconstanst.check(style))
    #     print("NEW Ex Style: %s" % exstyle)
    #     print(wpfconstanst.check2(exstyle))

    #     ctypes.windll.user32.SetWindowLongW(wnd_wih.Handle, GWL_STYLE, style )
    #     ctypes.windll.user32.SetWindowLongW(wnd_wih.Handle, GWL_EXSTYLE, exstyle )


class ViewModel(bkt.ui.ViewModelSingleton):
    def __init__(self):
        super(ViewModel, self).__init__()

class BktWindow(bkt.ui.WpfWindowAbstract):
    _filename = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'BktWindow.xaml')
    _vm_class = ViewModel

    def __init__(self, context):
        super(BktWindow, self).__init__(context)

    def cancel(self, sender, event):
        self.Close()
    
    def get_rect(self, sender, event):
        window = self._context.app.ActiveWindow
        slidem = window.View.Slide.Master

        left, top = window.PointsToScreenPixelsX(0), window.PointsToScreenPixelsY(0)
        right, bottom = window.PointsToScreenPixelsX(slidem.Width), window.PointsToScreenPixelsY(slidem.Height)
        bkt.message("X {}, Y {}\nW {}, H {}".format(left, top, right-left, bottom-top))

        rect = self.GetDeviceRect()
        bkt.message("X {}, Y {}\nW {}, H {}".format(rect.X, rect.Y, rect.Width, rect.Height))

    def set_pos_tl(self, sender, event):
        window = self._context.app.ActiveWindow
        left, top = window.PointsToScreenPixelsX(0), window.PointsToScreenPixelsY(0)
        self.SetDevicePosition(left, top)

    def set_pos_br(self, sender, event):
        window = self._context.app.ActiveWindow
        slidem = window.View.Slide.Master
        right, bottom = window.PointsToScreenPixelsX(slidem.Width), window.PointsToScreenPixelsY(slidem.Height)
        self.SetDevicePosition(deviceRight=right, deviceBottom=bottom)

    def set_size(self, sender, event):
        window = self._context.app.ActiveWindow
        slidem = window.View.Slide.Master
        left, top = window.PointsToScreenPixelsX(0), window.PointsToScreenPixelsY(0)
        right, bottom = window.PointsToScreenPixelsX(slidem.Width), window.PointsToScreenPixelsY(slidem.Height)
        self.SetDeviceSize(right-left, bottom-top)



class WPFTest(object):
    @staticmethod
    def show_xaml(context):
        # def _get_hwnd():
        #     return ctypes.windll.user32.GetForegroundWindow()
        form = MyWindow()
        # wih = System.Windows.Interop.WindowInteropHelper(form)
        # wih.Owner = clr.Reference[System.IntPtr](ctypes.windll.user32.GetForegroundWindow())
        # form.owner = _get_hwnd()
        System.Windows.Interop.WindowInteropHelper(form).Owner = System.IntPtr(ctypes.windll.user32.GetForegroundWindow())
        form.ShowDialog()

    @staticmethod
    def show_wpf():
        my_window = Window()
        my_window.Title = 'Welcome to IronPython'

        my_stack = StackPanel()
        my_window.Content = my_stack

        my_button = Button()
        my_button.Content = 'Push Me'
        my_stack.Children.Add (my_button)

        my_window.ShowDialog()

    @staticmethod
    def show_bkt_window(context):
        BktWindow.create_and_show_dialog(context)

    @staticmethod
    def show_msgbox():
        bkt.message("Standard BKT message box with ctypes", "BKT")
        System.Windows.MessageBox.Show("WPF message box", "BKT")
        System.Windows.Forms.MessageBox.Show("WinForms message box", "BKT")


xamltest_gruppe = bkt.ribbon.Group(
    label='XAML',
    image_mso='HappyFace',
    children = [
        bkt.ribbon.Button(
            label="XAML Window",
            show_label=True,
            size="large",
            image_mso='HappyFace',
            on_action=bkt.Callback(WPFTest.show_xaml),
            # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
        ),
        bkt.ribbon.Button(
            label="WPF Window",
            show_label=True,
            size="large",
            image_mso='HappyFace',
            on_action=bkt.Callback(WPFTest.show_wpf),
            # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
        ),
        bkt.ribbon.Button(
            label="BKT Window",
            show_label=True,
            size="large",
            image_mso='HappyFace',
            on_action=bkt.Callback(WPFTest.show_bkt_window),
            # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
        ),
        bkt.ribbon.Button(
            label="Messageboxes",
            show_label=True,
            size="large",
            image_mso='HappyFace',
            on_action=bkt.Callback(WPFTest.show_msgbox),
            # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
        ),
    ]
)