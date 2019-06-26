# -*- coding: utf-8 -*-
'''
Created on 2019-06-26
@author: Florian Stallmann
'''

import bkt
import bkt.library.excel.helpers as xllib


### Room for testing and playing around:

def set_cell_color(cell, color_index, brightness):
    cell.Interior.ThemeColor = color_index
    cell.Interior.TintAndShade = brightness

def set_cell_color2(cell, color):
    cell.Interior.Color = color

def get_cell_color(cell):
    return [cell.Interior.ThemeColor, round(cell.Interior.TintAndShade,2), cell.Interior.Color]

testgruppe = bkt.ribbon.Group(
    label="TESTGRUPPE",
    children=[
        bkt.ribbon.ColorGallery(
            label="Testgallery",
            size="large",
            color_helper=xllib.ColorHelper,
            on_theme_color_change=bkt.Callback(set_cell_color, cell=True),
            on_rgb_color_change=bkt.Callback(set_cell_color2, cell=True),
            get_selected_color=bkt.Callback(get_cell_color, cell=True),
        )
    ]
)

bkt.excel.add_tab(bkt.ribbon.Tab(
    id='bkt_excel_demo',
    #id_q='nsBKT:excel_toolbox_advanced',
    label=u'BKT DEMO',
    insert_before_mso="TabHome",
    get_visible=bkt.Callback(lambda: True),
    children = [
        testgruppe,
    ]
))