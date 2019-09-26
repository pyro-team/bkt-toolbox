# -*- coding: utf-8 -*-
'''
Created on 24.07.2014

@author: fstsallmann
'''


import bkt

# import toolbox modules with ui
from toolbox import arrange, info, language, text, shape_adjustments, stateshapes, slides, shapes as mod_shapes, shape_selection


# FIXME: this should be easier
from os.path import dirname, join, normpath, realpath
bkt.apps.Resources.root_folders.append(normpath(join(dirname(realpath(__file__)), '..', 'toolbox', 'resources')))



# define context-menus and context-tabs
from toolbox import context_menus

# tab id for tab activator workaround
info.TabActivator.tab_id = "my_bkt_powerpoint_toolbox"




# ==============================
# = Definition of Toolbox-Tabs =
# ==============================

bkt.powerpoint.add_tab(bkt.ribbon.Tab(
    id="my_bkt_powerpoint_toolbox",
    label=u'BKT 1/2',
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
        #mod_shapes.size_group,
        mod_shapes.pos_size_group,
        arrange.arrange_group,
        arrange_advanced_small_group,
        arrange.distance_rotation_group,
        text.innenabstand_gruppe,
        text.paragraph_group,
        text.paragraph_indent_group,
        shape_adjustments.adjustments_group,
        info.info_group
    ]
))

bkt.powerpoint.add_tab(bkt.ribbon.Tab(
    id="bkt_powerpoint_toolbox_extensions",
    #id_q="nsBKT:powerpoint_toolbox_extensions",
    #insert_after_q="nsBKT:powerpoint_toolbox_advanced",
    insert_before_mso="TabHome",
    label=u'BKT 2/2',
    # get_visible defaults to False during async-startup
    get_visible=bkt.Callback(lambda:True),
    children = [
        arrange.arrange_advanced_group,
        arrange.euclid_angle_group,
        mod_shapes.format_group,
        #arrange.arrange_adv_easy_group,
        #mod_shapes.split_shapes_group,
        #language.sprachen_gruppe,
        stateshapes.stateshape_gruppe,
    ]
), extend=True)



