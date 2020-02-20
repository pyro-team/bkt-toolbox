# -*- coding: utf-8 -*-
'''
Created on 08.12.2016

@author: rdebeerst
'''

import bkt
import os.path


CONFIG_FOLDERS = "feature_folders"
UPDATE_URL = "https://api.github.com/repos/pyro-team/bkt-toolbox/releases/latest"

class FolderSetup(object):
    @classmethod
    def add_folder_by_dialog(cls, context):
        F = bkt.dotnet.import_forms()
        
        dialog = F.FolderBrowserDialog()
        dialog.SelectedPath = os.path.dirname(os.path.realpath(__file__))
        # dialog.Description = "Please choose an additional folder with BKT-features"
        dialog.Description = "Bitte einen BKT Feature-Ordner hinzufügen"
        
        if (dialog.ShowDialog(None) == F.DialogResult.OK):
            cls.add_folder(context, dialog.SelectedPath)
    
    @classmethod
    def add_folder(cls, context, folder):
        folders = context.config.feature_folders or []
        folders.append(folder)
        context.config.set_smart(CONFIG_FOLDERS, folders)
        BKTReload.reload_bkt(context)
    
    @classmethod
    def delete_folder(cls, context, folder):
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
        except Exception, e:
            pass

    @staticmethod
    def invalidate(context):
        try:
            context.addin.invalidate_ribbon()
        except Exception, e:
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
                bkt.helpers.message("Update available to v{}. \nYour version is v{}.".format(version_string, bkt.version_tag_name))
            else:
                bkt.helpers.message("No update available. Current version is v{}.".format(version_string))
        except Exception as e:
            bkt.helpers.message("Error calling and parsing update URL: {}".format(e))
    
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



class SettingsMenu(bkt.ribbon.Menu):
    def __init__(self, idtag="", **kwargs):
        postfix = ("-" if idtag else "") + idtag
        
        # if (bkt.config.use_keymouse_hooks or False):
        keymouse_hook_buttons = [
            bkt.ribbon.MenuSeparator(),
            bkt.ribbon.ToggleButton(
                id='key-mouse-hook-toggle' + postfix,
                label='Key-/Mouse-Hooks',
                get_pressed='GetMouseKeyHookActivated',
                on_action='ToggleMouseKeyHookActivation'
            )
        ]
        # else:
            # keymouse_hook_buttons = []
        
        def open_folder():
            from os import startfile
            bkt_folder=os.path.normpath(os.path.join(os.path.dirname(os.path.realpath(__file__)), "..", ".."))
            if os.path.isdir(bkt_folder):
                startfile(bkt_folder)
        
        def open_config():
            from os import startfile
            config_filename=os.path.join(os.path.dirname(os.path.realpath(__file__)), "..", "config.txt")
            if os.path.exists(config_filename):
                os.startfile(config_filename)
        
        super(SettingsMenu, self).__init__(
            id='bkt-settings' + postfix,
            image='bkt_logo', 
            children=[
                bkt.ribbon.Button(
                    id='settings-website' + postfix,
                    label="Website: bkt-toolbox.de",
                    image_mso="HyperlinkInsert",
                    on_action=bkt.Callback(BKTInfos.open_website)
                ),
                bkt.ribbon.Button(
                    id='settings-version' + postfix,
                    label="{} v{}".format(bkt.full_version, bkt.version_tag_name),
                    image_mso="Info",
                    on_action=bkt.Callback(BKTInfos.show_debug_message, context=True)
                ),
                bkt.ribbon.Button(
                    id='settings-updatecheck' + postfix,
                    label="Check for updates",
                    image_mso="SyncStatusUpToDate",
                    on_action=bkt.Callback(BKTInfos.check_for_update)
                ),
                bkt.ribbon.MenuSeparator(),
                bkt.ribbon.DynamicMenu(
                    label='Feature-Ordner',
                    image_mso='ModuleInsert',
                    get_content = bkt.Callback(lambda: self.get_folder_menu(postfix))
                ),
                #bkt.ribbon.MenuSeparator(),
                bkt.ribbon.Button(
                    id='settings-reload-addin' + postfix,
                    label="Addin neu laden",
                    image_mso="AccessRefreshAllLists",
                    on_action=bkt.Callback(BKTReload.reload_bkt)
                ),
                bkt.ribbon.Button(
                    id='settings-invalidate' + postfix,
                    label="Ribbon aktualisieren",
                    image_mso="AccessRefreshAllLists",
                    on_action=bkt.Callback(BKTReload.invalidate)
                ),
                # FIXME: idQ-Referenz funktioniert nicht, control wird nicht angezeigt
                # bkt.ribbon.Button(
                #     #id='reload-addin',
                #     idQ='nsBKT:ppt__043a3c86-6596-4e3d-9d92-727b870cfbf7',
                #     label="Addin neu laden",
                #     image_mso="Refresh",
                #     visible=True
                #     #on_action=bkt.Callback(settings. )
                # ),
                #bkt.ribbon.MenuSeparator(),
                #bkt.ribbon.Button(label='Ctrl=kleine Schritte'),
                bkt.ribbon.MenuSeparator(),
                bkt.ribbon.Button(
                    id='settings-open-folder' + postfix,
                    label="Öffne BKT-Ordner",
                    image_mso="Folder",
                    on_action=bkt.Callback(open_folder)
                ),
                bkt.ribbon.Button(
                    id='settings-open-config' + postfix,
                    label="Öffne config.txt",
                    image_mso="NewNotepadTool",
                    on_action=bkt.Callback(open_config)
                ),
            ] + keymouse_hook_buttons,
            **kwargs
        )
        
    def info_delete_button_for_folder(self, folder, postfix):
        return bkt.ribbon.Button(
            label= folder,
            image_mso='DeleteThisFolder',
            on_action=bkt.Callback(lambda context: FolderSetup.delete_folder(context, folder))
        )

    def get_folder_menu(self, postfix):
        return bkt.ribbon.Menu(
            xmlns="http://schemas.microsoft.com/office/2009/07/customui",
            id=None,
            children=[
                bkt.ribbon.Button(
                    id='setting_add_folder' + postfix,
                    label='Feature-Ordner hinzufügen',
                    #image_mso='Folder',
                    image_mso='ModuleInsert',
                    on_action=bkt.Callback(FolderSetup.add_folder_by_dialog)
                ),
                bkt.ribbon.MenuSeparator()
            ] + [
                self.info_delete_button_for_folder(folder, postfix)
                for folder in bkt.config.feature_folders
            ]
        )


#def get_task_pane_button(id='setting-toggle-bkttaskpane'):
def get_task_pane_button_list(id='setting-toggle-bkttaskpane'):
    if ((bkt.config.task_panes or False)):
        return [bkt.ribbon.ToggleButton(
            id=id,
            label='Task Pane',
            show_label=False,
            image_mso='MenuToDoBar',
            screentip="Show/Hide BKT task pane",
            tag='BKT Task Pane',
            get_pressed='GetPressed_TaskPaneToggler',
            on_action='OnAction_TaskPaneToggler')]
    else:
        return []


settings_menu = SettingsMenu("duplicate", label="Settings", show_label=False)

settings_home_tab = bkt.ribbon.Tab(
    id_mso="TabHome",
    children=[
        bkt.ribbon.Group(label="BKT", image="bkt_logo", children =[SettingsMenu("tabhome", size="large", label="Settings")] + get_task_pane_button_list())
    ] 
)

bkt.powerpoint.add_tab(settings_home_tab)
bkt.word.add_tab(settings_home_tab)
bkt.excel.add_tab(settings_home_tab)

