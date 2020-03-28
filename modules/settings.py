# -*- coding: utf-8 -*-
'''
Created on 08.12.2016

@author: rdebeerst
'''

from __future__ import absolute_import

import os.path

import bkt


CONFIG_FOLDERS = "feature_folders"
UPDATE_URL = "https://api.github.com/repos/pyro-team/bkt-toolbox/releases/latest"

class FolderSetup(object):
    @classmethod
    def add_folder_by_dialog(cls, context):
        F = bkt.dotnet.import_forms()
        
        dialog = F.FolderBrowserDialog()
        # select feature folder
        cur_folder = os.path.dirname(os.path.realpath(__file__))
        feature_folder = os.path.realpath(os.path.join(cur_folder, "..", "features"))
        if os.path.isdir(feature_folder):
            dialog.SelectedPath = feature_folder
        else:
            dialog.SelectedPath = cur_folder
        # dialog.Description = "Please choose an additional folder with BKT-features"
        dialog.Description = "Bitte einen BKT Feature-Ordner auswählen"
        
        if (dialog.ShowDialog(None) == F.DialogResult.OK):
            cls.add_folder(context, dialog.SelectedPath)
    
    @staticmethod
    def add_folder(context, folder):
        folders = context.config.feature_folders or []
        folders.append(folder)
        context.config.set_smart(CONFIG_FOLDERS, folders)
        BKTReload.reload_bkt(context)
    
    @staticmethod
    def delete_folder(context, folder):
        folders = context.config.feature_folders or []
        folders.remove(folder)
        context.config.set_smart(CONFIG_FOLDERS, folders)
        BKTReload.reload_bkt(context)


class BKTReload(object):
    @staticmethod
    def reload_bkt(context):
        try:
            addin = context.app.COMAddIns["BKT.AddIn"]
            addin.Connect = False
            addin.Connect = True
        except Exception as e:
            pass

    @staticmethod
    def invalidate(context):
        try:
            context.addin.invalidate_ribbon()
        except Exception as e:
            pass


class BKTInfos(object):
    @staticmethod
    def open_website():
        import webbrowser
        webbrowser.open('https://www.bkt-toolbox.de')
    
    @staticmethod
    def check_for_update():
        import json
        import urllib2

        try:
            response = urllib2.urlopen(UPDATE_URL, timeout=4).read()
            data = json.loads(response)
            version_string = data["tag_name"]
            version = tuple(int(x) for x in version_string.split("."))
            current_version = tuple(int(x) for x in bkt.version_tag_name.split("."))
            if version > current_version:
                bkt.helpers.message("Aktualisiert verfügbar auf v{}. \nInstallierte Version ist v{}.".format(version_string, bkt.version_tag_name))
            else:
                bkt.helpers.message("Keine Aktualisierung verfügbar. Aktuelle Version ist v{}.".format(version_string))
        except Exception as e:
            bkt.helpers.message("Fehler im Aufruf der Aktualisierungs-URL: {}".format(e))
    
    @staticmethod
    def show_debug_message(context):
        import sys
        import bkt.console

        # https://docs.microsoft.com/de-de/office/troubleshoot/reference/numbering-scheme-for-product-guid

        winver = sys.getwindowsversion()
        debug_info = '''--- DEBUG INFORMATION ---

BKT-Framework Version:  {} (v{})
Operating System:       {} ({}.{}.{})
Office Version:         {} {}.{} ({})
IPY-Version:            {}
'''.format(
        bkt.full_version, bkt.version_tag_name,
        context.app.OperatingSystem, winver.major, winver.minor, winver.build,
        context.app.name, context.app.Version, context.app.Build, context.app.ProductCode,
        sys.version,
        )
        bkt.console.show_message(bkt.ui.endings_to_windows(debug_info))

    @staticmethod
    def get_bkt_folder_path():
        return os.path.normpath(os.path.join(os.path.dirname(os.path.realpath(__file__)), ".."))

    @staticmethod
    def get_bkt_config_path():
        return os.path.join(os.path.dirname(os.path.realpath(__file__)), "..", "config.txt")
        
    @classmethod
    def open_folder(cls, path=None):
        from os import startfile
        folder_to_open=path or cls.get_bkt_folder_path()
        if os.path.isdir(folder_to_open):
            startfile(folder_to_open)
    
    @classmethod
    def open_config(cls):
        from os import startfile
        config_filename=cls.get_bkt_config_path()
        if os.path.exists(config_filename):
            startfile(config_filename)



class SettingsMenu(bkt.ribbon.Menu):
    def __init__(self, idtag="", **kwargs):
        postfix = ("-" if idtag else "") + idtag

        if ((bkt.config.task_panes or False)):
            taskpanebutton = [
                bkt.ribbon.ToggleButton(
                id='setting-toggle-bkttaskpane' + postfix,
                label='Task Pane',
                show_label=False,
                image_mso='MenuToDoBar',
                supertip="BKT Task Pane (Seitenleiste) anzeigen/verstecken",
                tag='BKT Task Pane',
                get_pressed='GetPressed_TaskPaneToggler',
                on_action='OnAction_TaskPaneToggler')
            ]
        else:
            taskpanebutton = []
        
        super(SettingsMenu, self).__init__(
            id='bkt-settings' + postfix,
            image='bkt_logo', 
            supertip="BKT-Einstellungen verwalten, BKT neu laden, Website aufrufen, etc.",
            children=[
                bkt.ribbon.Button(
                    id='settings-website' + postfix,
                    label="Website: bkt-toolbox.de",
                    supertip="BKT-Webseite im Browser öffnen",
                    image_mso="HyperlinkInsert",
                    on_action=bkt.Callback(BKTInfos.open_website, transaction=False)
                ),
                bkt.ribbon.Button(
                    id='settings-version' + postfix,
                    label="{} v{}".format(bkt.full_version, bkt.version_tag_name),
                    image_mso="Info",
                    supertip="Erweiterte Versionsinformationen anzeigen",
                    on_action=bkt.Callback(BKTInfos.show_debug_message, context=True, transaction=False)
                ),
                bkt.ribbon.Button(
                    id='settings-updatecheck' + postfix,
                    label="Auf neue Version prüfen",
                    supertip="Überprüfen, ob neue BKT-Version verfügbar ist",
                    image_mso="SyncStatusUpToDate",
                    on_action=bkt.Callback(BKTInfos.check_for_update, transaction=False)
                ),
                bkt.ribbon.MenuSeparator(),
                bkt.ribbon.DynamicMenu(
                    label='Feature-Ordner',
                    supertip="Feature-Ordner hinzufügen oder entfernen",
                    image_mso='ModuleInsert',
                    get_content = bkt.Callback(lambda context: self.get_folder_menu(context, postfix), context=True)
                ),
                #bkt.ribbon.MenuSeparator(),
                bkt.ribbon.Button(
                    id='settings-reload-addin' + postfix,
                    label="Addin neu laden",
                    supertip="BKT-Addin beenden und neu laden (ähnlich PowerPoint-Neustart)",
                    image_mso="AccessRefreshAllLists",
                    on_action=bkt.Callback(BKTReload.reload_bkt, context=True, transaction=False)
                ),
                bkt.ribbon.Button(
                    id='settings-invalidate' + postfix,
                    label="Ribbon aktualisieren",
                    supertip="Oberfläche aktualisieren und alle Werte neu laden (sog. Invalidate ausführen)",
                    image_mso="AccessRefreshAllLists",
                    on_action=bkt.Callback(BKTReload.invalidate, context=True, transaction=False)
                ),
                bkt.ribbon.MenuSeparator(),
                bkt.ribbon.Button(
                    id='settings-open-folder' + postfix,
                    label="Öffne BKT-Ordner",
                    supertip="Öffne Ordner mit BKT-Framework und Konfigurationsdatei",
                    image_mso="Folder",
                    on_action=bkt.Callback(BKTInfos.open_folder, transaction=False)
                ),
                bkt.ribbon.Button(
                    id='settings-open-config' + postfix,
                    label="Öffne config.txt",
                    supertip="Öffne Konfigurationsdatei im Standardeditor",
                    image_mso="NewNotepadTool",
                    on_action=bkt.Callback(BKTInfos.open_config, transaction=False)
                ),
                bkt.ribbon.MenuSeparator(),
                bkt.ribbon.ToggleButton(
                    id='key-mouse-hook-toggle' + postfix,
                    label='Key-/Mouse-Hooks an/aus',
                    supertip="Tastatur-/Maus-Events für aktuelle Sitzung ein- oder ausschalten",
                    get_pressed='GetMouseKeyHookActivated',
                    on_action='ToggleMouseKeyHookActivation'
                )
            ] + taskpanebutton,
            **kwargs
        )
        
    def info_delete_button_for_folder(self, label, folder):
        return bkt.ribbon.Button(
            label=label,
            supertip="Feature-Ordner »{}« aus BKT-Konfiguration entfernen".format(folder),
            image_mso='DeleteThisFolder',
            on_action=bkt.Callback(lambda context: FolderSetup.delete_folder(context, folder))
        )

    def get_folder_menu(self, context, postfix):
        import importlib
        
        buttons = []

        for folder in bkt.config.feature_folders:
            module_name = os.path.basename(folder)
            try:
                module = importlib.import_module(module_name + '.__bkt_init__')
                buttons.append(
                    self.info_delete_button_for_folder(module.BktFeature.name, folder)
                )
            except:
                buttons.append(
                    self.info_delete_button_for_folder(module_name, folder)
                )

        return bkt.ribbon.Menu(
            xmlns="http://schemas.microsoft.com/office/2009/07/customui",
            id=None,
            children=[
                bkt.ribbon.Button(
                    id='setting_add_folder' + postfix,
                    label='Feature-Ordner hinzufügen',
                    supertip="Einen BKT Feature-Ordner auswählen und hinzufügen",
                    image_mso='ModuleInsert',
                    on_action=bkt.Callback(FolderSetup.add_folder_by_dialog)
                ),
                bkt.ribbon.MenuSeparator()
            ] + buttons
        )


settings_menu = SettingsMenu("duplicate", label="Settings", show_label=False)

settings_home_tab = bkt.ribbon.Tab(
    id_mso="TabHome",
    children=[
        bkt.ribbon.Group(
            id="bkt_tabhome_settings_group",
            label="BKT",
            image="bkt_logo",
            children = [SettingsMenu("tabhome", size="large", label="Settings")]
        )
    ] 
)

bkt.powerpoint.add_tab(settings_home_tab)
bkt.word.add_tab(settings_home_tab)
bkt.excel.add_tab(settings_home_tab)

