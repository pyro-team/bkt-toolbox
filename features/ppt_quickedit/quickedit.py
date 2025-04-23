# -*- coding: utf-8 -*-
'''
Created on 2018-01-10
@author: Florian Stallmann
'''



import logging

import bkt
import bkt.library.powerpoint as pplib


class QuickEditPanelManager(object):
    panel_windows = {}

    @staticmethod
    def _create_panel(context):
        from .quickedit_panel import QuickEditPanel
        return QuickEditPanel(context)
    
    @staticmethod
    def show_help():
        from .quickedit_model import QuickEdit
        QuickEdit.show_help()
    
    @staticmethod
    def autostart_pressed():
        return bkt.settings.get("quickedit.restore_panel", False)
    
    @classmethod
    def autostart_toggle(cls, pressed):
        bkt.settings["quickedit.restore_panel"] = pressed
        if cls.panel_windows:
            panel = next(iter(cls.panel_windows.values()))
            panel._vm.OnPropertyChanged("auto_start")

    @classmethod
    def get_panel_for_active_window(cls, context):
        logging.debug("get panel for active window")
        windowid = context.addin.GetWindowHandle()
        if windowid in cls.panel_windows:
            return cls.panel_windows[windowid]
        else:
            return None

    @classmethod
    def autoshow_panel_for_active_window(cls, context, presentation):
        logging.debug("auto show panel for active window")
        autoshow = bkt.settings.get("quickedit.restore_panel", False)
        if autoshow and cls._is_windowed_presentation(context, presentation):
            cls.show_panel_for_active_window(context)

    @classmethod
    def show_panel_for_active_window(cls, context):
        logging.debug("show panel for active window")

        windowid = context.addin.GetWindowHandle()
        if windowid in cls.panel_windows:
            if cls.panel_windows[windowid].IsLoaded:
                #ensure that window is on the screen
                cls.panel_windows[windowid].ShiftWindowOntoScreen()
                return #active panel window already exists
            else:
                cls._close_panel(windowid)

        cls._show_panel(context, windowid)

    @classmethod
    def close_panel_for_active_window(cls, context, presentation):
        logging.debug("close panel for active window")
        if cls._is_windowed_presentation(context, presentation):
            windowid = context.addin.GetWindowHandle()
            cls._close_panel(windowid)

    @classmethod
    def _show_panel(cls, context, windowid):
        logging.debug("show panel for window %s" % windowid)
        try:
            panel = cls._create_panel(context)
            panel.SetOwner(windowid)
            panel.Show()
            panel.ShiftWindowOntoScreen() #ensure that window is on the screen
            cls.panel_windows[windowid] = panel
            panel.update_docking()
        except:
            logging.exception("panel activation failed")
            if bkt.config.show_exception:
                bkt.helpers.exception_as_message()
            else:
                bkt.message("Unbekannter Fehler beim Anzeigen des QuickEdit-Panels!")

    @classmethod
    def _close_panel(cls, windowid):
        logging.debug("close panel for window %s" % windowid)
        try:
            cls.panel_windows[windowid].Close()
            del cls.panel_windows[windowid]
        except:
            pass
    
    @staticmethod
    def _is_windowed_presentation(context, presentation):
        try:
            #only show if at least one window exists
            return presentation.Windows.Count > 0
            #ALTERNATIVE: only show if opened presentation equals active presentation (not the case if opened without window)
            # return presentation.FullName == context.presentation.FullName
        except: #COMException
            return False

    @classmethod
    def close_all_panels(cls):
        logging.debug("close all panels")
        for windowid in list(cls.panel_windows.keys()):
            cls._close_panel(windowid)

    @classmethod
    def update_panel_position_for_presentation(cls, presentation, window, context):
        logging.debug("update panel position for active window")

        if cls.panel_windows:
            windowid = context.addin.GetWindowHandle(window)
            if windowid in cls.panel_windows and cls.panel_windows[windowid].IsLoaded:
                cls.panel_windows[windowid].update_docking(window=window)

    @classmethod
    def update_panel_position_for_slide(cls, slide_range, context):
        logging.debug("update panel position for active window")

        if cls.panel_windows:
            windowid = context.addin.GetWindowHandle()
            if windowid in cls.panel_windows and cls.panel_windows[windowid].IsLoaded:
                cls.panel_windows[windowid].update_docking()


bkt.AppEvents.after_new_presentation  += bkt.Callback(QuickEditPanelManager.autoshow_panel_for_active_window, context=True)
bkt.AppEvents.after_presentation_open += bkt.Callback(QuickEditPanelManager.autoshow_panel_for_active_window, context=True)

bkt.AppEvents.presentation_close += bkt.Callback(QuickEditPanelManager.close_panel_for_active_window, context=True)
bkt.AppEvents.bkt_unload         += bkt.Callback(QuickEditPanelManager.close_all_panels)

bkt.AppEvents.window_activate    += bkt.Callback(QuickEditPanelManager.update_panel_position_for_presentation, context=True)
bkt.AppEvents.slide_selection_changed    += bkt.Callback(QuickEditPanelManager.update_panel_position_for_slide, context=True)


color_selector_gruppe = bkt.ribbon.Group(
    id="bkt_quickedit_group",
    label='QuickEdit',
    supertip="Aktiviert eine kleine freischwebende Mini-Toolbar mit einer interaktiven Farbauswahl. Das Feature `ppt_quickedit` muss installiert sein.",
    image_mso='SmartArtChangeColorsGallery',
    children = [
        bkt.ribbon.SplitButton(
            id="quickedit_splitbutton",
            size="large",
            children=[
                bkt.ribbon.Button(
                    id="quickedit_button",
                    image="qe_icon",
                    label="QuickEdit Panel",
                    supertip="Blendet das QuickEdit Panel mit der interaktiven Farbauswahl ein.",
                    on_action=bkt.Callback(QuickEditPanelManager.show_panel_for_active_window, context=True)
                ),
                bkt.ribbon.Menu(
                    label="QuickEdit Panel Menü",
                    children=[
                        bkt.ribbon.Button(
                            id="quickedit_button2",
                            image="qe_icon",
                            label="Für dieses Fenster anzeigen",
                            supertip="Blendet das QuickEdit Panel mit der interaktiven Farbauswahl ein.",
                            on_action=bkt.Callback(QuickEditPanelManager.show_panel_for_active_window, context=True)
                        ),
                        bkt.ribbon.Button(
                            label="Für dieses Fenster schließen",
                            supertip="Schließt das QuickEdit Panel für das aktuelle PowerPoint-Fenster.",
                            on_action=bkt.Callback(QuickEditPanelManager.close_panel_for_active_window, context=True, presentation=True)
                        ),
                        bkt.ribbon.Button(
                            label="Alle schließen",
                            supertip="Schließt alle QuickEdit Panels für alle PowerPoint-Fenster.",
                            on_action=bkt.Callback(QuickEditPanelManager.close_all_panels)
                        ),
                        bkt.ribbon.MenuSeparator(),
                        bkt.ribbon.Button(
                            label="Hilfe anzeigen",
                            image_mso="Help",
                            supertip="Öffnet die Hilfedatei.",
                            on_action=bkt.Callback(QuickEditPanelManager.show_help)
                        ),
                        bkt.ribbon.MenuSeparator(title="Optionen"),
                        bkt.ribbon.ToggleButton(
                            label="Autostart ein/aus",
                            supertip="Anzeige des QuickEdit Panels bei PowerPoint-Start ein- und ausschalten.",
                            get_pressed=bkt.Callback(QuickEditPanelManager.autostart_pressed),
                            on_toggle_action=bkt.Callback(QuickEditPanelManager.autostart_toggle),
                        ),
                    ]
                )
            ]
        ),
    ]
)



##############################################################
###  ORIGINAL RIBBON  ########################################
##############################################################

# def color_toggle_button(c):
#     return bkt.ribbon.ToggleButton(
#                 id="quickedit_color_%s" % c.get_identifier(),
#                 label="QuickEdit %s" % c.get_label(),
#                 show_label=False,
#                 supertip="Setzt oder selektiert die ausgewählte Farbe für Hintergrund, Linie oder Text, abhängig von gedrückter STRG, SHIFT und ALT-Taste.\n\nIst der Button markiert, wird die Farbe in der aktuellen Auswahl als Hintergrund und/oder Linie verwendet.",
#                 tag=c.get_identifier(),
#                 get_image=bkt.Callback(c.get_image),
#                 # get_image=bkt.Callback(QuickEdit.get_image_by_control, current_control=True, context=True),
#                 on_toggle_action=bkt.Callback(QuickEdit.action_by_control, current_control=True, context=True),
#                 get_pressed=bkt.Callback(c.get_checked),
#                 # get_pressed=bkt.Callback(QuickEdit.get_pressed_by_control, current_control=True, context=True),
#                 # get_pressed=bkt.Callback(lambda context: QuickEdit.get_pressed(context, c), context=True),
#                 # get_enabled=bkt.Callback(QuickEdit.get_enabled, current_control=True, context=True, cache=False),
#             )

# color_selector_gruppe = bkt.ribbon.Group(
#     id="bkt_quickedit_group",
#     label='QuickEdit',
#     image_mso='SmartArtChangeColorsGallery',
#     get_visible=bkt.Callback(QuickEdit.initialize, context=True),
#     children = [
#         bkt.ribbon.Box(box_style="vertical", children=[
#             bkt.ribbon.Label(label="Theme: "),
#             bkt.ribbon.Label(label="Recent: "),
#             bkt.ribbon.Label(label="Own: "),
#         ]),
#         bkt.ribbon.ButtonGroup(id="bkt_quickedit_colors", children=[
#             bkt.ribbon.Button(
#                 id="quickedit_color_none",
#                 label="QuickEdit No Fill",
#                 show_label=False,
#                 supertip="Setzt oder selektiert Shapes ohne Fülling bei Hintergrund, Linie oder Text, abhängig von gedrückter STRG, SHIFT und ALT-Taste.\n\nIst der Button markiert, wird die Farbe in der aktuellen Auswahl als Hintergrund und/oder Linie verwendet.",
#                 image_mso="TableDivideUp",
#                 on_action=bkt.Callback(QuickEdit.action_no_fill, context=True),
#                 # get_pressed=bkt.Callback(QuickEdit.get_pressed_no_fill, context=True, shapes=True),
#             ),
#         ] + [
#             color_toggle_button(c)
#             for c in QuickEdit._colors
#         ]),
#         bkt.ribbon.ButtonGroup(id="bkt_quickedit_recent", children=[
#             bkt.ribbon.Button(
#                 id="quickedit_recent_add",
#                 label="QuickEdit Add Recent Color",
#                 show_label=False,
#                 image_mso="PickUpStyle",
#                 supertip="Hintergrundfarbe des ausgewählten Shapes zu zuletzt verwendeten Farben hinzufügen.",
#                 on_action=bkt.Callback(QuickEdit.pickup_recent_color, context=True),
#                 # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
#             ),
#         ] + [
#             color_toggle_button(c)
#             for c in QuickEdit._recent
#         ]),
#         bkt.ribbon.ButtonGroup(id="bkt_quickedit_own", children=[
#             bkt.ribbon.Button(
#                 id="quickedit_own_add",
#                 label="QuickEdit Add Own Color",
#                 show_label=False,
#                 image_mso="PickUpStyle",
#                 supertip="Hintergrundfarbe des ausgewählten Shapes zu eigenen Farben hinzufügen.",
#                 on_action=bkt.Callback(QuickEdit.pickup_own_color, context=True),
#                 # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
#             ),
#         ] + [
#             color_toggle_button(c)
#             for c in QuickEdit._userdefined
#         ]),
#         bkt.ribbon.Separator(),
#         bkt.ribbon.Box(box_style="vertical", children=[
#             bkt.ribbon.Label(label="[Shift]: Auswahl"),
#             bkt.ribbon.Label(label="[Strg]: Linie"),
#             bkt.ribbon.Label(label="[Alt]: Text"),
#         ]),
#         bkt.ribbon.Box(box_style="horizontal", children=[
#             bkt.ribbon.Button(image_mso="Help", label="Hilfe", show_label=False, on_action=bkt.Callback(QuickEdit.show_help)),
#             bkt.ribbon.Label(label=u"[Shift: Auswahl | Strg: Linie | Alt: Text]"),
#         ]),
#     ]
# )


bkt.powerpoint.add_tab(bkt.ribbon.Tab(
    id="bkt_powerpoint_toolbox_extensions",
    #id_q="nsBKT:powerpoint_toolbox_extensions",
    #insert_after_q="nsBKT:powerpoint_toolbox_advanced",
    insert_before_mso="TabHome",
    label='Toolbox 3/3',
    # get_visible defaults to False during async-startup
    get_visible=bkt.Callback(lambda:True),
    children = [
        color_selector_gruppe,
    ]
), extend=True)



##############################################################
###  TASKPANE TEST  ##########################################
##############################################################

# tpbuttons = []

# def color_taskpane_button(c):
#     newr = bkt.taskpane.Wpf.Rectangle(
#                                     Fill="Green",
#                                     Height="16", Width="16",
#                                 )
#     newb = bkt.taskpane.Button(
#                 id="quickedit_tp_color_%s" % c,
#                 header="QuickEdit %s" % QuickEdit._get_label(c),
#                 size="small",
#                 tag=str(c),
#                 # get_image=bkt.Callback(QuickEdit.get_image_by_control, current_control=True, context=True),
#                 on_action=bkt.Callback(QuickEdit.action, current_control=True, context=True),
#                 prop1 = bkt.taskpane.Icon(children=[
#                                 newr
#                             ])
#             )
#     tpbuttons.append(newr)
#     return newb

# def recolor():
#     tpbuttons[1]["Fill"] = "Red"
#     # tpbuttons[1]["Fill"] = System.Windows.Media.Brushes.Red
#     # tpbuttons[1].attributes.Fill = System.Windows.Media.Brushes.Red

# # qe_taskpane = bkt.taskpane.Expander(auto_wrap=True, is_expanded=True, header="QuickEdit",
# #     children=[
# qe_taskpane = bkt.taskpane.Wpf.WrapPanel(
#     # Initialized = bkt.Callback(lambda: bkt.message("test1")),
#     # Loaded = bkt.Callback(lambda: bkt.message("test2")),
#     children=[
#         bkt.taskpane.Group(auto_wrap=True, show_separator=False,
#             children=[
#                 bkt.taskpane.Button(
#                     id="qe_col1",
#                     header="color1",
#                     size="small",
#                     image="settings",
#                     on_action = bkt.Callback(recolor),
#                 ),
#                 bkt.taskpane.Button(
#                     id="qe_col2",
#                     header="color2",
#                     size="small",
#                     on_action = bkt.Callback(lambda shapes: bkt.message(str(len(shapes))), shapes=True),
#                     prop1 = bkt.taskpane.Icon(children=[
#                                     bkt.taskpane.Wpf.Rectangle(
#                                         Fill="Green",
#                                         Height="16", Width="16",
#                                         RadiusX="2", RadiusY="2",
#                                     )
#                                 ])
#                 ),
#             ]
#         ),
#         bkt.taskpane.Group(auto_wrap=True, show_separator=False,
#             children=[
#                 color_taskpane_button(c)
#                 for c in QuickEdit._buttons1
#             ]
#         ),
#     ]
# )

# # print qe_taskpane.wpf_xml()

# bkt.powerpoint.add_taskpane_control(qe_taskpane)