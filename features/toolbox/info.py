# -*- coding: utf-8 -*-
'''
Created on 06.02.2018

@author: rdebeerst
'''

from __future__ import absolute_import

import sys

import bkt

# reuse settings-menu from bkt-framework
import modules.settings as settings

version_short = bkt.version_tag_name
version_long  = 'Powerpoint Toolbox v{}'.format(bkt.version_tag_name)


# Workaround to activate Tab when new shape is added instead of auto switching to "Format" contextual tab
class TabActivator(object):
    activated = False
    context = None
    tab_id = "bkt_powerpoint_toolbox"
    shapes_on_slide = 0

    @classmethod
    def activate_tab_on_new_shape(cls, selection):
        #FIXME: fires also when shape is copy-pasted, but should only fire for real new shapes
        try:
            count_shapes = selection.SlideRange[1].Shapes.Count
            if selection.type == 2 and count_shapes > cls.shapes_on_slide and selection.ShapeRange[1].Type != 6: #ppSelectionShape, shapes increased, no group
                #bkt.helpers.message("shape added")
                cls.context.ribbon.ActivateTab(cls.tab_id)
                # print("tab activator: default tab activated")
            cls.shapes_on_slide = count_shapes
        except:
            pass
            # print("tab activator: failed activating tab")

    @classmethod
    def enable(cls, context):
        if not cls.activated and bkt.config.ppt_activate_tab_on_new_shape:
            cls.context = context
            #FIXME: event is not unassigned on reload/unload of addin
            # context.app.WindowSelectionChange += cls.activate_tab_on_new_shape
            bkt.AppEvents.selection_changed += bkt.Callback(cls.activate_tab_on_new_shape, selection=True)
            # print("tab activator: workaround enabled")
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
        bkt.helpers.warning(message, "BKT: Protected view window detected!")



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


class PopupConfig(object):
    @staticmethod
    def get_config(context):
        return context.app_ui.use_contextdialogs is False

    @staticmethod
    def set_config(context, pressed):
        context.app_ui.use_contextdialogs = not pressed
        bkt.config.set_smart("ppt_use_contextdialogs", not pressed)



class ToolbarVariations(object):
    #FIXME: very hard-coded, should be more flexible and allow multiple variations

    @classmethod
    def get_pressed_default(cls):
        return "toolbox_widescreen" not in sys.modules
    
    @classmethod
    def get_pressed_wide(cls):
        return "toolbox_widescreen" in sys.modules
        
    @classmethod
    def change_to_default(cls, context, pressed):
        cls.change_variation(context, "toolbox")
        
    @classmethod
    def change_to_wide(cls, context, pressed):
        cls.change_variation(context, "toolbox_widescreen")

    @classmethod
    def change_variation(cls, context, variation):
        from os.path import join
        folders = context.config.feature_folders or []
        folder = bkt.helpers.file_base_path_join(__file__, "..")
        # print(join(folder,"toolbox"))
        # remove both folders just in case
        try:
            folders.remove(join(folder,"toolbox"))
        except ValueError:
            pass
        try:
            folders.remove(join(folder,"toolbox_widescreen"))
        except ValueError:
            pass
        folders.insert(0, join(folder, variation))
        context.config.set_smart("feature_folders", folders)

        #reload bkt using settings module
        if bkt.helpers.confirmation("Soll die BKT nun neu geladen werden?"):
            settings.BKTReload.reload_bkt(context)
    
    @classmethod
    def show_uisettings(cls, context):
        from .toolboxui import ToolboxUi
        ToolboxUi.get_instance().show_settings_editor(context)


settings.settings_menu.children.extend([
    bkt.ribbon.ToggleButton(
        label="Format-Tab ausblenden",
        get_pressed=bkt.Callback(FormatTab.get_config),
        on_toggle_action=bkt.Callback(FormatTab.set_config, context=True),
        supertip="Format-Tab wird ausgeblendet, um automatischen Wechsel zum Tab beim Anlegen von Shapes zu verhindern."
    ),
    bkt.ribbon.ToggleButton(
        label="Popup-Dialog deaktivieren",
        get_pressed=bkt.Callback(PopupConfig.get_config),
        on_toggle_action=bkt.Callback(PopupConfig.set_config, context=True),
        supertip="Deaktiviert die Popup-Dialoge von BKT-Shapes wie Harvey-Balls, verknüpfte Shapes, etc.",
    ),
    bkt.ribbon.Menu(
        label="Toolbox-Tabs anpassen",
        image_mso="PageSettings",
        supertip="Tabs und Gruppen der PowerPoint-Toolbox auf individuelle Bedürfnisse anpassen",
        children=[
            bkt.ribbon.ToggleButton(
                label="Standard (3-seitig)",
                supertip="Drei Tabs für die Toolbox mit allen erweiterten Features auf einer separaten Seite 3",
                get_pressed=bkt.Callback(ToolbarVariations.get_pressed_default),
                on_toggle_action=bkt.Callback(ToolbarVariations.change_to_default, context=True)
            ),
            bkt.ribbon.ToggleButton(
                label="Widescreen (2-seitig)",
                supertip="Zwei Tabs für die Toolbox mit allen erweiterten Features gemeinsam auf Seite 2.",
                get_pressed=bkt.Callback(ToolbarVariations.get_pressed_wide),
                on_toggle_action=bkt.Callback(ToolbarVariations.change_to_wide, context=True)
            ),
            bkt.ribbon.MenuSeparator(),
            bkt.ribbon.Button(
                label="Theme-Einstellungen anpassen",
                supertip="Festlegung der Seite je Gruppe und Ausblenden von Gruppen.",
                on_action=bkt.Callback(ToolbarVariations.show_uisettings),
            ),
        ]
    ),
])


# Workaround is enabled via "get_visible" of info group:
info_group = bkt.ribbon.Group(
    id="bkt_settings_group",
    label="Settings",
    image_mso="AddInManager",
    get_visible=bkt.Callback(TabActivator.enable, context=True),
    children=[
        settings.settings_menu,
        bkt.ribbon.Button(label=version_short, screentip="Toolbox", supertip=version_long + "\n" + bkt.full_version, on_action=bkt.Callback(settings.BKTInfos.show_version_dialog)),
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