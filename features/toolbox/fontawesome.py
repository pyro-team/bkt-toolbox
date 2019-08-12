# -*- coding: utf-8 -*-
'''
Created on 10.02.2017

@author: rdebeerst
'''

import os.path
import imp

import bkt
from bkt.library.powerpoint import PPTSymbolsGallery


# import fontsymbols
#
fontsettings = [
    # module-name, font-filename
    ('fontawesome4', 'FontAwesome'),
    ('fontawesome5', 'Font Awesome 5 Free Regular'),
    ('segoeui', 'Segoe UI'),
    ('segoemdl2', 'Segoe MDL2 Assets'),
    ('materialicons', 'Material Icons'),
    # ('foobar', 'Non-existing test font'),
]


# helper to check system-fonts
import System.Drawing.Text
font_collection = System.Drawing.Text.InstalledFontCollection()
all_fonts = [font.Name for font in font_collection.Families]

def font_exists(fontname):
    return fontname in all_fonts
    

# initialize galleries 
symbol_galleries = []

for fontsetting in fontsettings:
    font_module, font_name = fontsetting
    
    # check if font exists
    if font_exists(font_name):
        # import the corresponding font-symbol-module from 'fontsymbols'-folder
        base_folder = os.path.dirname(os.path.realpath(__file__))
        fontsymbolmodule = imp.load_source(font_module,"%s\\fontsymbols\\%s.py" % (base_folder, font_module))
        
        # add menu seperator with title
        if fontsymbolmodule.menu_title:
            symbol_galleries += [
                bkt.ribbon.MenuSeparator(title="" + fontsymbolmodule.menu_title),
            ]
        else:
            symbol_galleries += [
                bkt.ribbon.MenuSeparator(),
            ]
        
        # add font-symbol-galleries
        symbol_galleries += fontsymbolmodule.menus
    else:
        symbol_galleries += [
            bkt.ribbon.MenuSeparator(title=font_name),
            bkt.ribbon.Button(
                label="Font nicht installiert",
                enabled=False
            )
        ]
    

