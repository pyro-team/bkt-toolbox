# -*- coding: utf-8 -*-
'''
Created on 24.07.2014

@author: rdebeerst
'''


import bkt

# import toolbox modules with ui
import arrange
import info
import language
import text
import shape_adjustments
import stateshapes
import slides
import shapes as mod_shapes
import shape_selection




# define context-menus and context-tabs
import context_menus




# ==============================
# = Definition of Toolbox-Tabs =
# ==============================

bkt.powerpoint.add_tab(bkt.ribbon.Tab(
    id="bkt_powerpoint_toolbox",
    label=u'Toolbox 1/3',
    insert_before_mso="TabHome",
    # get_visible defaults to False during async-startup
    get_visible=bkt.Callback(lambda:True),
    children = [
        shape_selection.clipboard_group,
        slides.slides_group,
        bkt.mso.group.GroupFont,
        bkt.mso.group.GroupParagraph,
        mod_shapes.shapes_group,
        styles_group,
        mod_shapes.size_group,
        arrange.arrange_group,
        arrange.distance_rotation_group,
        text.innenabstand_gruppe,
        text.paragraph_group,
        shape_adjustments.adjustments_group,
        info.info_group
    ]
))

bkt.powerpoint.add_tab(bkt.ribbon.Tab(
    id="bkt_powerpoint_toolbox_advanced",
    label=u'Toolbox 2/3',
    insert_before_mso="TabHome",
    # get_visible defaults to False during async-startup
    get_visible=bkt.Callback(lambda:True),
    children = [
        mod_shapes.pos_size_group,
        arrange.arrange_advanced_group,
        arrange.arrange_adv_easy_group,
        arrange.euclid_angle_group,
        mod_shapes.format_group,
        text.paragraph_indent_group,
        mod_shapes.split_shapes_group,
        language.sprachen_gruppe,
        stateshapes.stateshape_gruppe,
    ]
))





