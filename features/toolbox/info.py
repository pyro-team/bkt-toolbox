# -*- coding: utf-8 -*-
'''
Created on 06.02.2018

@author: rdebeerst
'''

import sys
import bkt

# reuse settings-menu from bkt-framework
import modules.settings as settings

version_short = 'v2.4'
version_long  = 'Powerpoint Toolbox v2.4 / r18-03-29'


# Workaround to activate Tab when new shape is added instead of auto switching to "Format" contextual tab
class TabActivator(object):
    activated = False
    context = None
    tab_id = "bkt_powerpoint_toolbox"
    shapes_on_slide = 0

    @classmethod
    def activate_tab_on_new_shape(cls, selection):
        count_shapes = selection.SlideRange(1).Shapes.Count
        #FIXME: fires also when shape is copy-pasted, but should only fire for real new shapes
        try:
            if selection.type == 2 and count_shapes > cls.shapes_on_slide and selection.ShapeRange[1].Type != 6: #ppSelectionShape, shapes increased, no group
                #bkt.helpers.message("shape added")
                cls.context.ribbon.ActivateTab(cls.tab_id)
                print "tab activator: default tab activated"
        except:
            pass
        cls.shapes_on_slide = count_shapes

    @classmethod
    def enable(cls, context):
        if not cls.activated and bkt.config.ppt_activate_tab_on_new_shape:
            cls.context = context
            #FIXME: event is not unassigned on reload/unload of addin
            context.app.WindowSelectionChange += cls.activate_tab_on_new_shape
            print "tab activator: workaround enabled"
        cls.activated = True
        return True


class ProtectedView(object):
    @staticmethod
    def get_visible(application):
        return application.ProtectedViewWindows.Count > 0
    
    @staticmethod
    def show_warning():
        message = '''At least one open presentation in protected view detected. Even if the protected view window is in the background, PowerPoint might show unexpected behavior such as keyboard input lags or shapes are glued to the cursor on selection.

If you continue editing in PowerPoint it is highly recommended to open all presentations in editing mode or close all protected view windows. This is not a BKT bug but a PowerPoint bug.'''
        bkt.helpers.message(message, "Protected view window detected!")



class FormatTab(object):
    ppt_hide_format_tab = bkt.config.ppt_hide_format_tab is True

    @classmethod
    def get_visible(cls):
        return not cls.ppt_hide_format_tab
    
    @classmethod
    def get_config(cls):
        return cls.ppt_hide_format_tab

    @classmethod
    def set_config(cls, context, pressed):
        cls.ppt_hide_format_tab = pressed
        bkt.config.set_smart("ppt_hide_format_tab", cls.ppt_hide_format_tab)
        # context.ribbon.InvalidateControlMso("TabDrawingToolsFormat")
        # context.ribbon.InvalidateControlMso("TabSetDrawingTools")


class ToolbarVariations(object):
    #FIXME: very hard-coded, should be more flexible and allow multiple variations

    @classmethod
    def get_pressed(cls):
        return "toolbox_widescreen" in sys.modules

    @classmethod
    def change_variation(cls, context, pressed):
        from os.path import dirname, realpath, normpath, join
        folders = context.config.feature_folders or []
        folder = join(dirname(realpath(__file__)), "..")
        # print(normpath(join(folder,"toolbox")))
        # remove both folders just in case
        try:
            folders.remove(normpath(join(folder,"toolbox")))
        except ValueError:
            pass
        try:
            folders.remove(normpath(join(folder,"toolbox_widescreen")))
        except ValueError:
            pass
        if pressed:
            folders.insert(0, normpath(join(folder,"toolbox_widescreen")))
        else:
            folders.insert(0, normpath(join(folder,"toolbox")))
        context.config.set_smart("feature_folders", folders)
        #reload bkt using settings module
        sys.modules["modules"].settings.BKTReload.reload_bkt(context)


settings.settings_menu.children.extend([
    bkt.ribbon.ToggleButton(
        label="Format-Tab ausblenden",
        get_pressed=bkt.Callback(FormatTab.get_config),
        on_toggle_action=bkt.Callback(FormatTab.set_config, context=True)
    ),
    bkt.ribbon.ToggleButton(
        label="Profi-Widescreen-Theme",
        get_pressed=bkt.Callback(ToolbarVariations.get_pressed),
        on_toggle_action=bkt.Callback(ToolbarVariations.change_variation, context=True)
    )
])


# Workaround is enabled via "get_visible" of info group:
info_group = bkt.ribbon.Group(
    id="bkt_settings_group",
    label="Settings",
    image_mso="AddInManager",
    get_visible=bkt.Callback(TabActivator.enable, context=True),
    children=[
        settings.settings_menu,
        bkt.ribbon.Button(label=version_short, screentip="Toolbox", supertip=version_long + "\n" + bkt.full_version, on_action=bkt.Callback(settings.BKTInfos.show_debug_message)),
    ] + settings.get_task_pane_button_list(id='toolbox-taskpane-toggler') + [
        bkt.ribbon.Button(
            label="BKT Warning",
            size="large",
            image_mso="CancelRequest",
            screentip="Protected window warning",
            supertip="At least one open presentation in protected view detected. Unexpected PowerPoint behavior may occur.",
            get_visible=bkt.Callback(ProtectedView.get_visible, application=True),
            on_action=bkt.Callback(ProtectedView.show_warning)
        ),
    ]
)

# Workaround to maintain focus on BKT tab
context_format_tab = bkt.ribbon.Tab(
    idMso = "TabDrawingToolsFormat",
    get_visible=bkt.Callback(FormatTab.get_visible),
)