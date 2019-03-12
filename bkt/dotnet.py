# -*- coding: utf-8 -*-
'''
Created on 23.11.2014

@author: cschmitt
'''

from __future__ import absolute_import, division, print_function

def import_linq():
    import clr
    clr.AddReference('System.Xml.Linq')
    import System.Xml.Linq as xml
    return xml

def import_forms():
    import clr
    clr.AddReference('System.Windows.Forms')
    import System.Windows.Forms as forms
    return forms

def import_drawing():
    import clr
    clr.AddReference('System.Drawing')
    import System.Drawing as drawing
    return drawing

def import_win32():
    import Microsoft.Win32 as win32
    return win32

def import_powerpoint():
    import clr
    clr.AddReference('Microsoft.Office.Interop.PowerPoint')
    import Microsoft.Office.Interop.PowerPoint as powerpoint
    return powerpoint

def import_excel():
    import clr
    clr.AddReference('Microsoft.Office.Interop.Excel')
    import Microsoft.Office.Interop.Excel as excel
    return excel

def import_outlook():
    import clr
    clr.AddReference('Microsoft.Office.Interop.Outlook')
    import Microsoft.Office.Interop.Outlook as outlook
    return outlook

def import_officecore():
    import clr
    clr.AddReference('Office')
    import Microsoft.Office.Core as officecore
    return officecore

def import_wpf():
    import clr
    clr.AddReference("IronPython.Wpf")
    import wpf as wpf
    return wpf

def import_bkt():
    import clr
    clr.AddReference("BKT")
    import BKT as bkt_addin
    return bkt_addin
