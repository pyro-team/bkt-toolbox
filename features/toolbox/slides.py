# -*- coding: utf-8 -*-
'''
Created on 06.07.2016

@author: rdebeerst
'''


import bkt

# for ui composition
from .chartlib import chartlib_button



slides_group = bkt.ribbon.Group(
    id="bkt_slide_group",
    label='Folien',
    image_mso='SlideNewGallery',
    children=[
        bkt.mso.splitbutton.SlideNewGallery,
        #bkt.mso.control.SlideLayoutGallery,
        chartlib_button,
        bkt.ribbon.DynamicMenu(
            label="Mehr",
            show_label=False,
            image_mso='TableDesign',
            screentip="Weitere Slide-Funktionen",
            supertip="Agenda, Foliennummerierung, Slidedeck aufräumen, und viele weitere Folien-bezogene Funktionen",
            get_content=bkt.CallbackLazy("toolbox.models.slides_menu", "slides_menu")
        )
    ]
)
