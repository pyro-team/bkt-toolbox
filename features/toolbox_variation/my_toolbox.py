# -*- coding: utf-8 -*-
'''
Created on 24.07.2014

@author: rdebeerst
'''


import bkt

# import toolbox modules with ui
from toolbox import arrange
from toolbox import harvey
from toolbox import info
from toolbox import language
# from toolbox import placeholder
from toolbox import shape_adjustments
from toolbox import shape_selection
from toolbox import shapes as mod_shapes
from toolbox import slides
from toolbox import stateshapes
from toolbox import text


# FIXME: this should be easier
import os
bkt.helpers.Resources.root_folders.append(os.path.join(os.path.dirname(os.path.realpath(__file__)), '..', 'toolbox', 'resources'))



# define context-menus and context-tabs
from toolbox import context_menus


# default ui for shape styling
styles_group = bkt.ribbon.Group(
    id="bkt_style_group",
    label='Stile',
    image_mso='ShapeFillColorPicker',
    children = [
        bkt.mso.control.ShapeFillColorPicker,
        bkt.mso.control.ShapeOutlineColorPicker,
        bkt.mso.control.ShapeEffectsMenu,
        bkt.mso.control.TextFillColorPicker,
        bkt.mso.control.TextOutlineColorPicker,
        bkt.mso.control.TextEffectsMenu,
        bkt.mso.control.OutlineWeightGallery,
        bkt.mso.control.OutlineDashesGallery,
        bkt.mso.control.ArrowStyleGallery,
        bkt.ribbon.DialogBoxLauncher(idMso='ObjectFormatDialog')
    ]
)


# ==============================
# = Definition of Toolbox-Tabs =
# ==============================

bkt.powerpoint.add_tab(bkt.ribbon.Tab(
    id="bkt_powerpoint_toolbox",
    label=u'My-Toolbox 1/3',
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
        ##### Added group from Tab 2 here
        arrange.arrange_advanced_group,
        #####
        info.info_group
    ]
))

bkt.powerpoint.add_tab(bkt.ribbon.Tab(
    id="bkt_powerpoint_toolbox_advanced",
    label=u'My-Toolbox 2/3',
    insert_before_mso="TabHome",
    # get_visible defaults to False during async-startup
    get_visible=bkt.Callback(lambda:True),
    children = [
        mod_shapes.pos_size_group,
        mod_shapes.format_group,
        mod_shapes.split_shapes_group,
        text.paragraph_indent_group,
        language.sprachen_gruppe,
        stateshapes.stateshape_gruppe,
    ]
))





