# -*- coding: utf-8 -*-
'''
Load System or Office specific .Net functions

Created on 23.11.2014
@author: cschmitt
'''



import clr

from functools import wraps

cache = {}
def memoize(func):
    @wraps(func)
    def memoizer():
        try:
            return cache[func]
        except KeyError:
            result = cache[func] = func()
            return result
    return memoizer

@memoize
def import_linq():
    clr.AddReference('System.Xml.Linq')
    import System.Xml.Linq as xml
    return xml

@memoize
def import_forms():
    clr.AddReference('System.Windows.Forms')
    import System.Windows.Forms as forms
    return forms

@memoize
def import_drawing():
    clr.AddReference('System.Drawing')
    import System.Drawing as drawing
    return drawing

@memoize
def import_win32():
    import Microsoft.Win32 as win32
    return win32

@memoize
def import_powerpoint():
    clr.AddReference('Microsoft.Office.Interop.PowerPoint')
    import Microsoft.Office.Interop.PowerPoint as powerpoint
    return powerpoint

@memoize
def import_excel():
    clr.AddReference('Microsoft.Office.Interop.Excel')
    import Microsoft.Office.Interop.Excel as excel
    return excel

@memoize
def import_outlook():
    clr.AddReference('Microsoft.Office.Interop.Outlook')
    import Microsoft.Office.Interop.Outlook as outlook
    return outlook

@memoize
def import_visio():
    clr.AddReference('Microsoft.Office.Interop.Visio')
    import Microsoft.Office.Interop.Visio as visio
    return visio

@memoize
def import_officecore():
    clr.AddReference('Office')
    import Microsoft.Office.Core as officecore
    return officecore

@memoize
def import_wpf():
    clr.AddReference("IronPython.Wpf")
    import wpf
    return wpf

@memoize
def import_bkt():
    clr.AddReference("BKT")
    import BKT as bkt_addin
    return bkt_addin
