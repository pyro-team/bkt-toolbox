# -*- coding: utf-8 -*-
'''
Created on 24.07.2014

@author: fstsallmann
'''

from __future__ import absolute_import


import bkt

# FIXME: this should be easier
from os.path import dirname, join, normpath, realpath
bkt.apps.Resources.root_folders.append(normpath(join(dirname(realpath(__file__)), '..', 'toolbox', 'resources')))
# ###

from toolbox import toolboxui, info


# tab id for tab activator workaround
info.TabActivator.tab_id = "my_bkt_powerpoint_toolbox"


### default settings
default_settings = {
    "tab_name": "BKT", #Tab Name
    "size_group": 0, #off
    "pos_size_group": 1, #page no 1
    "arrange_mini_group": 1, #page no 1
    "arrange_adv_easy_group": 0, #off
    "text_parindent_group": 1, #page no 1
}

### render toolbox ui
toolbox_ui = toolboxui.ToolboxUi(default_settings, 2)
toolbox_ui.render_pages()
toolbox_ui.render_contextmenus()


# ==============================
# = Definition of Toolbox-Tabs =
# ==============================

bkt.powerpoint.add_tab(bkt.ribbon.Tab(
    id="my_bkt_powerpoint_toolbox",
    label=toolbox_ui.get_page_name(1),
    insert_before_mso="TabHome",
    # get_visible defaults to False during async-startup
    get_visible=bkt.Callback(lambda:True),
    children = toolbox_ui.get_page(1),
))

bkt.powerpoint.add_tab(bkt.ribbon.Tab(
    id="bkt_powerpoint_toolbox_extensions",
    #id_q="nsBKT:powerpoint_toolbox_extensions",
    #insert_after_q="nsBKT:powerpoint_toolbox_advanced",
    insert_before_mso="TabHome",
    label=toolbox_ui.get_page_name(2),
    # get_visible defaults to False during async-startup
    get_visible=bkt.Callback(lambda:True),
    children = toolbox_ui.get_page(2),
), extend=True)



