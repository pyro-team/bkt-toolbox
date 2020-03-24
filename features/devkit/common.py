# -*- coding: utf-8 -*-
'''
Created on 26.02.2020

@author: fstallmann
'''

from __future__ import absolute_import

import logging

import bkt
import modules.settings as settings



class DevGroup(object):
    log_level = None
    
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
    def toggle_legacy_syntax(pressed):
        bkt.config.set_smart("enable_legacy_syntax", pressed)
    
    @staticmethod
    def change_log_level(pressed, current_control):
        bkt.config.set_smart("log_level", current_control["tag"])
    
    @classmethod
    def get_log_level(cls, current_control):
        if cls.log_level is None:
            logger = logging.getLogger()
            cls.log_level = logging.getLevelName(logger.level)
        return cls.log_level == current_control["tag"]




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
                bkt.ribbon.Button(
                    label="Open Cache folder",
                    image_mso="Folder",
                    on_action=bkt.Callback(lambda: settings.BKTInfos.open_folder(bkt.helpers.get_cache_folder()), transaction=False)
                ),
                bkt.ribbon.Button(
                    label="Open Settings folder",
                    image_mso="Folder",
                    on_action=bkt.Callback(lambda: settings.BKTInfos.open_folder(bkt.helpers.get_settings_folder()), transaction=False)
                ),
                bkt.ribbon.Button(
                    label="Open Favorites folder",
                    image_mso="Folder",
                    on_action=bkt.Callback(lambda: settings.BKTInfos.open_folder(bkt.helpers.get_fav_folder()), transaction=False)
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
                bkt.ribbon.ToggleButton(
                    label="Legacy Syntax on/off",
                    get_pressed=bkt.Callback(lambda: bkt.config.enable_legacy_syntax or False),
                    on_toggle_action=bkt.Callback(DevGroup.toggle_legacy_syntax, transaction=False)
                ),
                bkt.ribbon.Menu(
                    label="Change log-level",
                    children=[
                        bkt.ribbon.ToggleButton(
                            label="DEBUG",
                            tag="DEBUG",
                            get_pressed=bkt.Callback(DevGroup.get_log_level),
                            on_toggle_action=bkt.Callback(DevGroup.change_log_level, transaction=False)
                        ),
                        bkt.ribbon.ToggleButton(
                            label="INFO",
                            tag="INFO",
                            get_pressed=bkt.Callback(DevGroup.get_log_level),
                            on_toggle_action=bkt.Callback(DevGroup.change_log_level, transaction=False)
                        ),
                        bkt.ribbon.ToggleButton(
                            label="WARNING",
                            tag="WARNING",
                            get_pressed=bkt.Callback(DevGroup.get_log_level),
                            on_toggle_action=bkt.Callback(DevGroup.change_log_level, transaction=False)
                        ),
                        bkt.ribbon.ToggleButton(
                            label="ERROR",
                            tag="ERROR",
                            get_pressed=bkt.Callback(DevGroup.get_log_level),
                            on_toggle_action=bkt.Callback(DevGroup.change_log_level, transaction=False)
                        ),
                        bkt.ribbon.ToggleButton(
                            label="CRITICAL",
                            tag="CRITICAL",
                            get_pressed=bkt.Callback(DevGroup.get_log_level),
                            on_toggle_action=bkt.Callback(DevGroup.change_log_level, transaction=False)
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
        #TODO: create new feature folder, clear all caches
        #ICONS: ControlsPane
    ]
)


class ImageMso(object):
    search_limit = 100

    search_engine = None
    search_term = ""
    search_results = None

    @classmethod
    def load_json(cls, search_engine):
        import os.path
        import io
        import json

        search_writer = search_engine.writer()
        file = os.path.join(os.path.dirname(os.path.realpath(__file__)), "imagemso.json")
        with io.open(file, 'r', encoding='utf-8') as json_file:
            for icon in json.load(json_file):
                search_writer.add_document(module="imagemso", name=icon, keywords=icon)
            search_writer.commit()

    @classmethod
    def get_search_engine(cls):
        if cls.search_engine:
            return cls.search_engine

        from bkt.search import get_search_engine
        cls.search_engine = get_search_engine("imagemso")
        # initialize search index on first use of engine
        cls.load_json(cls.search_engine)
        return cls.search_engine
    
    @classmethod
    def _perform_search(cls):
        if len(cls.search_term) > 0:
            cls.search_results = None
            engine = cls.get_search_engine()
            with engine.searcher() as searcher:
                    cls.search_results = searcher.search(cls.search_term)
        else:
            cls.search_results = None

    @classmethod
    def set_search_term(cls, value):
        cls.search_term = value
        cls._perform_search()

    @classmethod
    def get_search_term(cls):
        return cls.search_term
    
    @classmethod
    def get_symbol_galleries(cls):
        if not cls.search_results or len(cls.search_results) == 0:
            image_msos = [
                bkt.ribbon.Button(
                    label="Keine Ergebnisse für »{}«".format(cls.search_term),
                    enabled=False
                )
            ]
        
        else:
            image_msos = [
                bkt.ribbon.MenuSeparator(title="{} Ergebnisse für »{}«".format(len(cls.search_results), cls.search_term))
            ]
            if len(cls.search_results) > cls.search_limit:
                image_msos.append(
                    bkt.ribbon.Button(label="Zeige die ersten {} Ergebnisse".format(cls.search_limit), enabled=False)
                )
            for icon in cls.search_results.limit(cls.search_limit):
                image_msos.append(
                    bkt.ribbon.Button(
                        label=icon.name,
                        image_mso=icon.name,
                        on_action=bkt.Callback(cls.copy_to_clipboard)
                    )
                )
        
        return bkt.ribbon.Menu(
                    xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                    id=None,
                    children=image_msos
                )

    @classmethod
    def get_enabled_results(cls):
        return cls.search_results is not None
    
    @classmethod
    def get_results_label(cls):
        if cls.search_results is not None:
            return "{} Icons".format(len(cls.search_results))
        else:
            return "Ergebnis"

    @classmethod
    def copy_to_clipboard(cls, current_control):
        import bkt.dotnet as dotnet
        Forms = dotnet.import_forms() #required to read clipboard
        Forms.Clipboard.SetText( current_control['label'] )



iconsearch_group = bkt.ribbon.Group(
    id="bkt_common_dev_iconsearch_group",
    label="ImageMso",
    image_mso='SearchUI',
    children=[
        bkt.ribbon.Label(
            label="Suchwort:",
        ),
        bkt.ribbon.EditBox(
            label="Suchwort",
            show_label=False,
            sizeString = '#########',
            get_text = bkt.Callback(ImageMso.get_search_term),
            on_change = bkt.Callback(ImageMso.set_search_term),
            supertip="Suchwort eingeben und ENTER klicken",
        ),
        bkt.ribbon.DynamicMenu(
            get_label=bkt.Callback(ImageMso.get_results_label),
            get_content=bkt.Callback(ImageMso.get_symbol_galleries),
            get_enabled=bkt.Callback(ImageMso.get_enabled_results),
            screentip="Suchergebnisse",
            supertip="Zeigt die Suchergebnisse der Icon-Suche nach dem gewünschten Suchwort an",
        ),
    ]
)


common_groups = [common_group, iconsearch_group]