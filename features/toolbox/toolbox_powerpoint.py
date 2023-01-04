# -*- coding: utf-8 -*-
'''
Created on 24.07.2014

@author: rdebeerst
'''



import bkt

# load function to get toolbox pages based on settings
from . import toolboxui



### default settings
default_settings = {
    "split_group": 2, #page no 2
    "language_group": 2, #page no 2
}

### render toolbox ui
toolbox_ui = toolboxui.ToolboxUi(default_settings, 3)
toolbox_ui.render_pages()
toolbox_ui.render_contextmenus()


# ==============================
# = Definition of Toolbox-Tabs =
# ==============================

bkt.powerpoint.add_tab(bkt.ribbon.Tab(
    id="bkt_powerpoint_toolbox",
    label=toolbox_ui.get_page_name(1),
    insert_before_mso="TabHome",
    # get_visible defaults to False during async-startup
    get_visible=bkt.Callback(lambda:True),
    children = toolbox_ui.get_page(1),
))

bkt.powerpoint.add_tab(bkt.ribbon.Tab(
    id="bkt_powerpoint_toolbox_advanced",
    label=toolbox_ui.get_page_name(2),
    insert_before_mso="TabHome",
    # get_visible defaults to False during async-startup
    get_visible=bkt.Callback(lambda:True),
    children = toolbox_ui.get_page(2),
))

bkt.powerpoint.add_tab(bkt.ribbon.Tab(
    id="bkt_powerpoint_toolbox_extensions",
    label=toolbox_ui.get_page_name(3),
    insert_before_mso="TabHome",
    # get_visible defaults to False during async-startup
    get_visible=bkt.Callback(lambda:True),
), extend=True)

