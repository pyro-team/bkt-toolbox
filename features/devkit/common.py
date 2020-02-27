# -*- coding: utf-8 -*-
'''
Created on 26.02.2020

@author: fstallmann
'''


import bkt
import modules.settings as settings


class DevGroup(object):
    
    @staticmethod
    def show_console(context):
        import bkt.console as co
        co.console.Visible = True
        co.console.scroll_down()
        co.console.BringToFront()  # @UndefinedVariable
        co.console._globals['context'] = context

    @staticmethod
    def reload_bkt(context):
        settings.BKTReload.reload_bkt(context)
        # import bkt.console
        # try:
        #     addin = context.app.COMAddIns["BKT.AddIn"]
        #     addin.Connect = False
        #     addin.Connect = True
        # except Exception, e:
        #     bkt.console.show_message(str(e))
    
    @staticmethod
    def show_config(context):
        import bkt.console
        def _iter_lines():
            cfg = dict(context.config.items("BKT"))
            for k in sorted(cfg):
                yield "{:30} = {}".format(k, str(getattr(context.config, k)))
            yield ''
        
        bkt.console.show_message('\r\n'.join(_iter_lines()))
    
    @staticmethod
    def show_settings(context):
        import bkt.console
        def _iter_lines():
            for k in sorted(context.settings):
                yield "{:35} = {}".format(k, context.settings.get(k, "ERROR"))
            yield ''
        
        bkt.console.show_message('\r\n'.join(_iter_lines()))

    @staticmethod
    def show_ribbon_xml(python_addin, ribbon_id):
        import bkt.console
        bkt.console.show_message(python_addin.get_custom_ui(ribbon_id))
        
    @staticmethod
    def toggle_show_exception(pressed):
        bkt.config.set_smart("show_exception", pressed)
        
    @staticmethod
    def toggle_log_write_file(pressed):
        bkt.config.set_smart("log_write_file", pressed)
    
    @staticmethod
    def change_log_level(new_level):
        bkt.config.set_smart("log_level", new_level)




common_group = bkt.ribbon.Group(
    id="bkt_common_dev_group",
    label="Common development tools",
    image_mso="ControlsGallery",
    children=[
        bkt.ribbon.Button(
            label="Interactive Console",
            size="large",
            image_mso="WatchWindow",
            on_action=bkt.Callback(DevGroup.show_console, context=True, transaction=False),
        ),
        bkt.ribbon.Button(
            label="Reload BKT",
            size="large",
            image_mso="AccessRefreshAllLists",
            on_action=bkt.Callback(DevGroup.reload_bkt, context=True, transaction=False),
        ),
        bkt.ribbon.Separator(),
        bkt.ribbon.Menu(
            label="Config",
            size="large",
            image_mso="WebPageComponent",
            children=[
                bkt.ribbon.Button(
                    label="Open BKT folder",
                    image_mso="Folder",
                    on_action=bkt.Callback(settings.BKTInfos.open_folder, transaction=False)
                ),
                bkt.ribbon.MenuSeparator(),
                bkt.ribbon.Button(
                    label="Open config.txt",
                    image_mso="NewNotepadTool",
                    on_action=bkt.Callback(settings.BKTInfos.open_config, transaction=False)
                ),
                bkt.ribbon.Button(
                    label="Show config",
                    image_mso="WebPageComponent",
                    on_action=bkt.Callback(DevGroup.show_config, context=True, transaction=False),
                ),
                bkt.ribbon.MenuSeparator(),
                bkt.ribbon.Button(
                    label="Show app settings",
                    image_mso="WebPartProperties",
                    on_action=bkt.Callback(DevGroup.show_settings, context=True, transaction=False),
                ),
                bkt.ribbon.MenuSeparator(),
                bkt.ribbon.ToggleButton(
                    label="Show exceptions on/off",
                    get_pressed=bkt.Callback(lambda: bkt.config.show_exception or False),
                    on_toggle_action=bkt.Callback(DevGroup.toggle_show_exception, transaction=False)
                ),
                bkt.ribbon.ToggleButton(
                    label="Write log file on/off",
                    get_pressed=bkt.Callback(lambda: bkt.config.log_write_file or False),
                    on_toggle_action=bkt.Callback(DevGroup.toggle_log_write_file, transaction=False)
                ),
                bkt.ribbon.Menu(
                    label="Change log-level",
                    children=[
                        bkt.ribbon.ToggleButton(
                            label="DEBUG",
                            get_pressed=bkt.Callback(lambda: bkt.config.log_level == "DEBUG"),
                            on_toggle_action=bkt.Callback(lambda pressed: DevGroup.change_log_level("DEBUG"), transaction=False)
                        ),
                        bkt.ribbon.ToggleButton(
                            label="INFO",
                            get_pressed=bkt.Callback(lambda: bkt.config.log_level == "INFO"),
                            on_toggle_action=bkt.Callback(lambda pressed: DevGroup.change_log_level("INFO"), transaction=False)
                        ),
                        bkt.ribbon.ToggleButton(
                            label="WARNING",
                            get_pressed=bkt.Callback(lambda: bkt.config.log_level == "WARNING"),
                            on_toggle_action=bkt.Callback(lambda pressed: DevGroup.change_log_level("WARNING"), transaction=False)
                        ),
                        bkt.ribbon.ToggleButton(
                            label="ERROR",
                            get_pressed=bkt.Callback(lambda: bkt.config.log_level == "ERROR"),
                            on_toggle_action=bkt.Callback(lambda pressed: DevGroup.change_log_level("ERROR"), transaction=False)
                        ),
                        bkt.ribbon.ToggleButton(
                            label="CRITICAL",
                            get_pressed=bkt.Callback(lambda: bkt.config.log_level == "CRITICAL"),
                            on_toggle_action=bkt.Callback(lambda pressed: DevGroup.change_log_level("CRITICAL"), transaction=False)
                        ),
                    ]
                )
            ]
        ),
        bkt.ribbon.Button(
            label="Show Ribbon XML",
            size="large",
            image="xml",
            on_action=bkt.Callback(DevGroup.show_ribbon_xml, python_addin=True, ribbon_id=True, transaction=False),
        ),
        #TODO: create new feature folder, clear all caches, show/hide contextmenu ids
        #ICONS: ControlsPane
    ]
)


class ImageMso(object):
    all_images = None

    @classmethod
    def load_json(cls):
        import os.path
        import io
        import json

        file = os.path.join(os.path.dirname(os.path.realpath(__file__)), "imagemso.json")
        with io.open(file, 'r', encoding='utf-8') as json_file:
            cls.all_images = json.load(json_file)




common_groups = [common_group]