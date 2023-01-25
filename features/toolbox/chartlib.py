# -*- coding: utf-8 -*-
'''
Created on 04.05.2016

@author: rdebeerst
'''

import bkt


chartlib_button = bkt.ribbon.DynamicMenu(
    id='menu-add-chart',
    label="Templatefolie einfügen",
    show_label=False,
    screentip="Folie aus Slide-Library einfügen",
    supertip="Aus den hinterlegten Slide-Templates kann ein Template als neue Folie eingefügt werden.",
    image_mso="BibliographyGallery",
    # image_mso="SlideMasterInsertLayout",
    #image_mso="CreateFormBlankForm",
    get_content = bkt.CallbackLazy("toolbox.models.chartlib", "charts", "get_root_menu")
)
shapelib_button = bkt.ribbon.DynamicMenu(
    id='menu-add-shape',
    label="Personal Shape Library",
    show_label=False,
    screentip="Shape aus Shape-Library einfügen",
    supertip="Aus den hinterlegten Shape-Templates kann ein Shape auf die aktuelle Folie eingefügt werden.",
    image_mso="ActionInsert",
    #image_mso="ShapesInsertGallery",
    #image_mso="OfficeExtensionsGallery",
    get_content = bkt.CallbackLazy("toolbox.models.chartlib", "shapes", "get_root_menu")
)

# chartlibgroup = bkt.ribbon.Group(
#     label="chartlib",
#     children=[ chartlib_button, shapelib_button]
# )

# bkt.powerpoint.add_tab(
#     bkt.ribbon.Tab(
#         label="chartlib",
#         children = [
#             chartlibgroup
#         ]
#     )
# )





