# -*- coding: utf-8 -*-
'''
Created on 04.05.2016

@author: rdebeerst
'''

from __future__ import absolute_import

import os #for os.path, listdir and makedirs
import time

import logging

from threading import Thread #used for non-blocking gallery thumbnails refresh
from contextlib import contextmanager #used for opening and closing presentations
from string import maketrans #for cleansing names

from System import Array

import bkt
import bkt.library.powerpoint as pplib
import bkt.library.graphics as glib

from bkt import dotnet

Drawing = dotnet.import_drawing() #required for getting png format from clipboard
Forms = dotnet.import_forms() #required for clipboard functions


@contextmanager
def open_presentation_without_window(context, filename, readonly=True):
    ''' opens and returns presentation file '''
    logging.debug("open_presentation_without_window: %s", filename)
    presentation = None
    try:
        presentation = context.app.Presentations.Open(filename, readonly, False, False) #readonly, untitled, withwindow
        yield presentation
    finally:
        if presentation:
            presentation.Saved = True
            presentation.Close()


# TODO
# ChartLib --> ChartLibMenu
# ChartLib class with
#    file-buttons
#    copy shape/slide action

THUMBNAIL_POSTFIX = '_thumbnails'
FILETYPES_PPT = ('.pptx', '.ppt')
FILETYPES_POT = ('.potx', '.pot')



class ChartLib(object):
    
    def __init__(self, copy_shapes=False):
        '''Constructor. Configures options and slide_action'''

        # configuration of library
        self.library_folders = []
        self.library_files = []
        
        # option: show gallery or menu for presentations
        self.show_gallery = True
        # callback method for menu-actions
        self.slide_action = None
        
        # caches
        self.cached_presentation_menus = {}
        self.cached_presentation_galleries = {}

        # folder for favorites
        # self.fav_folder = os.path.dirname(os.path.realpath(__file__))
        # self.fav_folder = bkt.helpers.get_fav_folder()
        self.fav_folder = None
        
        # init slide_action
        self.copy_shapes_setting = copy_shapes
        # if copy_shapes:
        #     # copy shapes
        #     self.slide_action = self.copy_shapes_callback
        #     self.fav_folder = os.path.join(self.fav_folder, "shapelib")
        # else:
        #     # copy full slide
        #     self.slide_action = self.copy_slide_callback
        #     self.fav_folder = os.path.join(self.fav_folder, "chartlib")
        
        # # add favorite folder as first folder
        # self.library_folders.insert(0, {'title': "Favoriten", 'folder': self.fav_folder} )
    

    def init_chartlib(self):
        logging.debug('initializing chartlib')

        if self.copy_shapes_setting:
            # copy shapes
            subfolder = "shapelib"
            self.slide_action = self.copy_shapes_callback

            library_folders = bkt.config.shape_library_folders or []
            self.library_files.extend( bkt.config.shape_libraries or [] )

        else:
            # copy full slide
            subfolder = "chartlib"
            self.slide_action = self.copy_slide_callback

            library_folders = bkt.config.chart_library_folders or []
            self.library_files.extend( bkt.config.chart_libraries or [] )

        # add from library_folders
        for folder in library_folders:
            self.library_folders.append( {'title':os.path.basename(os.path.realpath(folder)), 'folder':folder} )

        # add from feature-folders
        for folder in bkt.config.feature_folders:
            chartlib_folder = os.path.join(folder, subfolder)
            if os.path.exists(chartlib_folder):
                self.library_folders.append( {'title':os.path.basename(os.path.realpath(folder)), 'folder':chartlib_folder} )
        
        # sort by name
        self.library_files.sort()
        self.library_folders.sort(key=lambda f:f['title'])

        # add favorite folder as first folder
        self.fav_folder = bkt.helpers.get_fav_folder(subfolder)
        self.library_folders.insert(0, {'title': "Favoriten", 'folder': self.fav_folder} )
    
    
    # ===========
    # = Helpers =
    # ===========
    
    # @classmethod
    # def open_presentation_file(cls, context, file):
    #     ''' opens and returns presentation file '''
    #     # if self.presentations.has_key(file):
    #     #     presentation = self.presentations[file]
    #     # else:
    #     #     # # Parameter: schreibgeschützt, ohne Titel, kein Fenster
    #     #     # presentation = context.app.Presentations.Open(file, True, False, False)
    #     #     # Parameter: rw-access, ohne Titel, kein Fenster
    #     #     presentation = context.app.Presentations.Open(file, False, False, False)
    #     #     self.presentations[file] = presentation
    #     return context.app.Presentations.Open(file, True, False, False)
    
    @classmethod
    def create_or_open_presentation(cls, context, file):
        if os.path.exists(file):
            return context.app.Presentations.Open(file, False, False, False)
        else:
            newpres = context.app.Presentations.Add(False) #WithWindow=False
            if not os.path.isdir(os.path.dirname(file)):
                os.makedirs(os.path.dirname(file))
            newpres.SaveAs(file)
            return newpres
    
    @staticmethod
    def _save_paste(obj, *args, **kwargs):
        for _ in range(3):
            try:
                return obj.paste(*args, **kwargs)
            except EnvironmentError:
                logging.debug("chartlib pasting error, waiting for 50ms")
                #wait some time to avoid EnvironmentError (run ahead bug if clipboard is busy, see https://stackoverflow.com/questions/54028910/vba-copy-paste-issues-pptx-creation)
                time.sleep(.50)
                #FIXME: maybe better way to check if clipboard actually contains "something"
        else:
            raise EnvironmentError("pasting not successfull")

    def _open_in_explorer(self, path):
        from os import startfile
        if os.path.isdir(path):
            startfile(path)

    def _check_folder_in_lib(self, folder):
        for f in self.library_folders:
            if type(f) is dict:
                f = f["folder"]
            if folder.startswith(f):
                return True
        else:
            return False

    def _check_file_in_lib(self, file):
        if file in self.library_files:
            return True
        for f in self.library_folders:
            if type(f) is dict:
                f = f["folder"]
            if file.startswith(f):
                return True
        else:
            return False

    def _add_files_to_config(self, files):
        if self.copy_shapes_setting:
            conf = "shape_libraries"
        else:
            conf = "chart_libraries"
        
        cur_files = getattr(bkt.config, conf) or []
        cur_files.extend(files)
        bkt.config.set_smart(conf, cur_files)

        self.reset_cashes()
        self.library_files = []
        self.library_folders = []
        self.init_chartlib()

    def _add_folder_to_config(self, folder):
        if self.copy_shapes_setting:
            conf = "shape_library_folders"
        else:
            conf = "chart_library_folders"
        
        cur_folders = getattr(bkt.config, conf) or []
        cur_folders.append(folder)
        bkt.config.set_smart(conf, cur_folders)

        self.reset_cashes()
        self.library_files = []
        self.library_folders = []
        self.init_chartlib()

    def _is_template_file(self, filename):
        return filename.endswith(FILETYPES_POT)
    
    def _is_valid_powerpoint(self, filename):
        if self.copy_shapes_setting:
            return filename.endswith(FILETYPES_PPT) and not filename.startswith("~$")
        else:
            return filename.endswith(FILETYPES_PPT+FILETYPES_POT) and not filename.startswith("~$")
    
    def _get_control_for_file(self, filename):
        if not self.copy_shapes_setting and self._is_template_file(filename):
            return self.get_template_file_control(filename)
        elif self.show_gallery:
            return self.get_chartlib_gallery_from_file(filename)
        else:
            return self.get_dynamic_file_menu(filename)
    
    # ====================================
    # = Menus for folders and subfolders =
    # ====================================
    
    ##Overview of UI parts for chart library
    ###Chartlib-Files
    # - fixed chartlib-menu (Menu) from file/presentation  ***Will open presentation***
    #       get_chartlib_menu_from_file                    ***caches***
    #       get_chartlib_menu_from_presentation
    # - dynamic chartlib-menu (DynamicMenu)                ***Will only open presentation if menu is opened***
    #       get_dynamic_file_menu
    #       --> maps to get_chartlib_menu_from_file / get_chartlib_menu_callback
    # - dynamic chartlib-gallery (Gallery)                 ***Will only open presentation if gallery is opened***
    #       get_chartlib_gallery_from_file
    #       --> uses ChartLibGallery
    ###Chartlib-Folders
    # - fixed folder-menu (Menu) from folder
    #       get_folder_menu
    #       --> uses get_dynamic_file_menu for files
    #       --> uses get_dynamic_folder_menu for subfolders
    # - dynamic folder-menu (DynamicMenu)
    #       get_dynamic_folder_menu
    #       --> maps to get_folder_menu / get_folder_menu_callback
    # - root library menu (Menu)
    #       get_root_menu
    #       --> uses get_folder_menu for library folders to create menu-sections
    #       --> uses get_dynamic_file_menu for library files
    #
    
    def get_root_menu(self):
        '''
        Returns static menu-control for root-chartlib-dir
        with entries for every root-directory and root-file.
        If just one root-directory is configured, it lists its contents directly instead.
        
        Menu
          --- root-folder 1 ---
          file 1
          file 2
          ...
          --- root-folder 2 ---
          ...
          ---------------------
          root-file 1
          root-file 2
        '''
        
        #initialize chartlib on first call
        if not self.fav_folder:
            self.init_chartlib()

        # create menu-items for chartlibrary-directories
        if not self.library_folders:
            # create empty menu
            menu = bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None
            )
        elif len(self.library_folders) == 1:
            # create menu with contents of directory
            folder = self.library_folders[0]
            if type(folder) is dict:
                folder = folder['folder']
            menu = self.get_folder_menu(folder, id=None)
            menu.label=None
        else:
            # create menu with sections for each directory
            children = []
            for folder in self.library_folders:
                if type(folder) is dict:
                    title = folder['title']
                    directory = folder['folder']
                else:
                    title = os.path.basename(folder)
                    directory = folder
                if not os.path.isdir(directory):
                    logging.warning("chartlib: %s is not a folder", directory)
                    continue
                children.append( self.get_folder_menu(directory, label=title) )

            # old version (shorter but less readable and more loops):
            # folders = [ folder if type(folder) == dict else  {'folder':folder, 'title':os.path.basename(folder)}   for folder in self.library_folders]
            # logging.debug("ChartLib root menu with folders: %s", folders)
            # children = [ self.get_folder_menu(folder['folder'], label=folder['title']) for folder in folders]

            # make list flat
            flat_children = []
            for folder_menu in children:
                if folder_menu.children:
                    flat_children.append( bkt.ribbon.MenuSeparator(title=folder_menu['label'] or None) )
                    flat_children.extend( folder_menu.children )
            
            menu = bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None,
                children=flat_children
            )
            
            
        # create menu-items for chartlibrary-files
        if self.library_files:
            menu.children.append( bkt.ribbon.MenuSeparator(title="Einzel-Libraries") )
            for filename in self.library_files:
                if not os.path.isfile(filename):
                    logging.warning("chartlib: %s is not a file", filename)
                    continue
                menu.children.append( self._get_control_for_file(filename) )
        
        # buttons
        menu.children += [
            bkt.ribbon.MenuSeparator(title="Settings"),
            bkt.ribbon.Button(label="Markierte Shapes zu Favoriten" if self.copy_shapes_setting else "Markierte Folien zu Favoriten",
                screentip="Aktuelle Auswahl zu Favoriten-Library hinzufügen",
                supertip="Ausgewählte Slides/Shapes zu Standard-Favoriten-Library hinzufügen. Falls diese Library noch nicht existiert, wird diese neu angelegt.",
                image_mso='AddToFavorites',
                on_action=bkt.Callback(self.add_chart_to_lib, context=True)
            ),
            bkt.ribbon.Button(label="Library erneut indizieren",
                screentip="Library erneut indizieren",
                supertip="Alle Cashes werden zurückgesetzt. Library-Order und -Dateien werden jeweils beim nächsten Öffnen erneut indiziert.\n\nEs werden nur Thumbnails bereits geöffneter Libraries aktualisiert. Eine Neugenerierung der Thumbnails kann auch für jede Shape-Library-Datei separat angestoßen werden (im jeweiligen Menü).",
                image_mso='AccessRefreshAllLists',
                on_action=bkt.Callback(self.update_thumbnails_and_reset_cashes)
            ),
            bkt.ribbon.MenuSeparator(),
            bkt.ribbon.DynamicMenu(
                label="Library verwalten",
                supertip="Library-Dateien oder Ordner hinzufügen oder entfernen.",
                image_mso='AddToolGallery',
                get_content=bkt.Callback(self.get_libraries_menu),
            ),
        ]
        
        logging.debug("ChartLib root menu: %s", menu.xml())
        return menu
    
    def get_root_menu_xml(self):
        ''' returns xml of root-chartlib-dir menu'''
        return self.get_root_menu().xml_string()
    
    def get_folder_menu_callback(self, current_control):
        ''' callback for dynamic menu, returns menu-control for folder represented by control '''
        folder = current_control['tag']
        if not self.cached_presentation_menus.has_key(folder):
            folder_menu = self.get_folder_menu(folder, id=None)
            self.cached_presentation_menus[folder] = folder_menu
        # else:
        #     logging.debug('get_folder_menu_callback: reuse menu for %s', folder)
        return self.cached_presentation_menus[folder]
    
    def get_folder_menu(self, folder, **kwargs):
        ''' returns static menu-control for given folder. menu contains entries for subfolders and pptx-files. '''

        files = []
        subfolders = []
        
        if os.path.isdir(folder):
            for filename in os.listdir(folder):
                # find pptx-Files
                if self._is_valid_powerpoint(filename):
                    files.append(filename)
    
                # find subfolders
                if os.path.isdir(os.path.join(folder, filename)):
                    if not filename.endswith(THUMBNAIL_POSTFIX):
                        subfolders.append(filename)

            logging.debug('get_folder_menu files: %s', files)
            logging.debug('get_folder_menu folders: %s', subfolders)
            
        else:
            logging.warning("Chartlib. No such directory: %s", folder)
        
        # FIXME: no folders which belong to files
        
        children = []
        # create DynamicMenus / Galleries
        for filename in files:
            children.append( self._get_control_for_file(os.path.join(folder, filename)) )
        if files:
            #only add separator if files >0
            children.append(bkt.ribbon.MenuSeparator())
        for subfolder in subfolders:
            children.append( self.get_dynamic_folder_menu(os.path.join(folder, subfolder)) )
            
            # old version (shorter but less readable and more loops):
            # children = [
            #     self.get_chartlib_gallery_from_file(folder + "\\" + filename) if self.show_gallery else self.get_dynamic_file_menu(folder + "\\" + filename) 
            #     for filename in files
            # ] + ([bkt.ribbon.MenuSeparator()] if len(files) > 0 else [] )+ [
            #     self.get_dynamic_folder_menu(folder + "\\" + subfolder)
            #     for subfolder in subfolders
            # ]
        
        # return Menu
        return bkt.ribbon.Menu(
            xmlns="http://schemas.microsoft.com/office/2009/07/customui",
            children = children,
            **kwargs
        )
    
    def get_dynamic_folder_menu(self, folder):
        ''' returns dynamic menu for folder. if menu unfolds, menu content is obtained by get_folder_menu '''
        basename = os.path.basename(folder)
        return bkt.ribbon.DynamicMenu(label=basename, tag=folder, get_content=bkt.Callback(self.get_folder_menu_callback, current_control=True))
    
    def get_dynamic_file_menu(self, filename):
        ''' returns dynamic menu for file. if menu unfolds, menu content is obtained by get_chartlib_menu_from_file '''
        file_basename = os.path.splitext(os.path.basename(filename))[0]
        return bkt.ribbon.DynamicMenu(
            label=file_basename, tag=filename,
            get_content=bkt.Callback(self.get_chartlib_menu_callback, current_control=True, context=True)
        )

    def get_template_file_control(self, filename):
        ''' return menu for template file '''
        file_basename = os.path.splitext(os.path.basename(filename))[0]
        return bkt.ribbon.Menu(
            label=file_basename,
            image_mso="SlideMasterMasterLayout",
            children=[
                bkt.ribbon.Button(
                    label="Design auf markierte Folien anwenden",
                    supertip="Folienmaster auf die ausgewählten Folien anwenden.",
                    tag=filename,
                    on_action=bkt.Callback(self.apply_template_to_slides, current_control=True, context=True)
                ),
                bkt.ribbon.Button(
                    label="Design auf ganze Präsentation anwenden",
                    supertip="Folienmaster auf die ganze Präsentation anwenden.",
                    tag=filename,
                    on_action=bkt.Callback(self.replace_template_in_presentation, current_control=True, context=True)
                ),
                bkt.ribbon.Button(
                    label="Design zu Präsentation hinzufügen",
                    supertip="Folienmaster zu der aktuellen Präsentation hinzufügen.",
                    tag=filename,
                    on_action=bkt.Callback(self.apply_template_to_slides, current_control=True, context=True)
                ),
                bkt.ribbon.Button(
                    label="Neue Präsentation mit Design erstellen",
                    supertip="Neue leere Präsentation auf Basis des Folienmasters erstellen.",
                    tag=filename,
                    on_action=bkt.Callback(self.new_presentation_with_template, current_control=True, context=True)
                ),
                bkt.ribbon.MenuSeparator(),
                bkt.ribbon.Button(label="Datei öffnen und bearbeiten",
                    image_mso='FileSaveAsPowerPointPptx',
                    tag=filename,
                    on_action=bkt.Callback(self.open_file, context=True, current_control=True)
                ),
            ]
        )
    
    def update_thumbnails_and_reset_cashes(self, context):
        if not bkt.message.confirmation("Dieser Vorgang kann bei vielen Libraries einige Minuten dauern und nicht abgebrochen werden. Trotzdem fortsetzen?", "BKT: ChartLib"):
            return

        def loop(worker):
            try:
                worker.ReportProgress(1, "Lade Ordner und Dateien")
                #get all galleries
                galleries = []
                for file in self.library_files:
                    if worker.CancellationPending:
                        break
                    galleries.append( self.get_chartlib_gallery_from_file(file) )

                for folder in self.library_folders:
                    if worker.CancellationPending:
                        break
                    if type(folder) is dict:
                        folder = folder["folder"]
                    for root, _, files in os.walk(folder):
                        if root.endswith(THUMBNAIL_POSTFIX):
                            continue
                        for file in files:
                            if self._is_valid_powerpoint(file) and not file.endswith(FILETYPES_POT): #no need to create thumbnails from potx files
                                galleries.append( self.get_chartlib_gallery_from_file(os.path.join(root, file)) )

                total = len(galleries)+1
                current = 1.0
                # for gal in self.cached_presentation_galleries.values():
                for gal in galleries:
                    if worker.CancellationPending:
                        break
                    logging.info("Updating library %s", gal.filename)
                    if len(gal.filename) > 50:
                        worker.ReportProgress(current/total*100, "..." + gal.filename[-50:])
                    else:
                        worker.ReportProgress(current/total*100, gal.filename)
                    
                    gal.reset_gallery_items(context)
                    current += 1.0
                worker.ReportProgress(100, "Cache löschen")
            except:
                logging.exception("Error on refreshing chartlib libraries")
            finally:
                self.reset_cashes()
        
        bkt.ui.execute_with_progress_bar(loop, context, modal=False) #modal=False important so main thread can handle app events and all presentations close properly

    def reset_cashes(self):
        ''' cashes for library menus and galleries are deleted '''
        self.cached_presentation_menus = {}
        self.cached_presentation_galleries = {}
        

    # ==============================================
    # = Manage library, add files/folder to config =
    # ==============================================

    def get_libraries_menu(self):
        children = []
        if self.library_files:
            children.append( bkt.ribbon.MenuSeparator(title="Einzel-Libraries") )
            for filename in self.library_files:
                children.append(
                    bkt.ribbon.Button(
                        label=os.path.basename(filename),
                        supertip=filename,
                        tag=filename,
                        image_mso="Delete",
                        on_action=bkt.Callback(self.delete_library),
                    )
                )
        
        if self.library_folders:
            folders_lib = []
            folder_feat = []
            for folder in self.library_folders:
                if type(folder) is dict:
                    folder_feat.append(
                        bkt.ribbon.Button(
                            label=folder["title"],
                            supertip=folder["folder"],
                            tag="ff|"+folder["folder"],
                            image_mso="Delete",
                            on_action=bkt.Callback(self.delete_library),
                        )
                    )
                else:
                    folders_lib.append(
                        bkt.ribbon.Button(
                            label=os.path.basename(folder),
                            supertip=folder,
                            tag=folder,
                            image_mso="Delete",
                            on_action=bkt.Callback(self.delete_library),
                        )
                    )

            if folders_lib:
                children.append(bkt.ribbon.MenuSeparator(title="Ordner-Libraries"))
                children.extend(folders_lib)
            if folder_feat:
                children.append(bkt.ribbon.MenuSeparator(title="Über Feature-Folders"))
                children.extend(folder_feat)

        children.extend([
            bkt.ribbon.MenuSeparator(title="Hinzufügen"),
            bkt.ribbon.Button(
                label="Datei zu Library hinzufügen",
                supertip="Library-Dateien auswählen und zur Konfiguration dauerhaft hinzufügen.",
                image_mso='FilesToolAddFiles',
                on_action=bkt.Callback(self.add_files_to_config)
            ),
            bkt.ribbon.Button(
                label="Ordner zu Library hinzufügen",
                supertip="Library-Ordner auswählen und zur Konfiguration dauerhaft hinzufügen.",
                image_mso='Folder',
                on_action=bkt.Callback(self.add_folders_to_config)
            ),
            bkt.ribbon.Button(
                label="Neue Library-Datei erstellen",
                supertip="Library-Ordner auswählen und zur Konfiguration dauerhaft hinzufügen.",
                image_mso='FileSaveAsPowerPointPptx',
                on_action=bkt.Callback(self.create_new_library)
            ),
        ])
        
        return bkt.ribbon.Menu(
            xmlns="http://schemas.microsoft.com/office/2009/07/customui",
            id=None,
            children=children
        )
    
    def delete_library(self, current_control):
        lib = current_control["tag"]
        if lib.startswith("ff|"):
            lib = lib[3:]
            
            if lib == self.fav_folder:
                bkt.message.warning("Der Favoriten-Ordner selbst kann nicht gelöscht werden, aber einzelne Dateien können manuell gelöscht werden. Der Ordner wird nun im Explorer geöffnet.")
                return self._open_in_explorer(lib)
            
            if not bkt.message.confirmation("Soll der Feature-Folder aus der Library entfernt werden?\n\nAchtung: Feature-Folder können gleichzeitig ChartLibs, ShapeLibs und Funktionen enthalten, die dann nicht mehr verfügbar sind!"):
                return
            conf = "feature_folders"
        elif os.path.isfile(lib):
            if self.copy_shapes_setting:
                conf = "shape_libraries"
            else:
                conf = "chart_libraries"
        elif os.path.isdir(lib):
            if self.copy_shapes_setting:
                conf = "shape_library_folders"
            else:
                conf = "chart_library_folders"
        else:
            bkt.message.error("Library nicht gefunden!")
        
        current = getattr(bkt.config, conf) or []
        current.remove(lib)
        bkt.config.set_smart(conf, current)

        self.reset_cashes()
        self.library_files = []
        self.library_folders = []
        self.init_chartlib()

    def add_files_to_config(self):
        fileDialog = Forms.OpenFileDialog()
        fileDialog.Filter = "PowerPoint (*.pptx;*.ppt;*.pot;*.potx)|*.pptx;*.ppt;*.pot;*.potx|Alle Dateien (*.*)|*.*"
        fileDialog.Title = "PowerPoint-Dateien auswählen"
        fileDialog.Multiselect = True

        if not fileDialog.ShowDialog() == Forms.DialogResult.OK:
            return

        to_add = []
        skipped = []
        for file in fileDialog.FileNames:
            # if self._check_file_in_lib(file):
            #     skipped.append(file)
            # else:
            #     to_add.append(file)
            if os.path.isfile(file):
                to_add.append(file)

        if skipped:
            bkt.message.warning("Einige Dateien sind bereits in der Library und werden nicht hinzugefügt!")

        self._add_files_to_config(to_add)

    def add_folders_to_config(self):
        dialog = Forms.FolderBrowserDialog()
        dialog.Description = "Bitte einen Ordner mit PowerPoint-Dateien auswählen"
        
        if dialog.ShowDialog() == Forms.DialogResult.OK:
            folder = dialog.SelectedPath
            if os.path.isdir(folder):
                if self._check_folder_in_lib(folder):
                    return bkt.message.warning("Der Ordner (oder ein Überordner) ist bereits in der Library und wird nicht hinzugefügt!")

                self._add_folder_to_config(folder)
    
    def create_new_library(self, context):
        name = bkt.ui.show_user_input("Bitte Name der Library eingeben:", "Name", "BKT Lib")
        if not name:
            return
        
        if not name.endswith(".pptx"):
            name += ".pptx"
        
        file = os.path.join(self.fav_folder, name)
        if os.path.isfile(file):
            return bkt.message.error("Library existiert schon!")
        
        # self._add_files_to_config([file])

        pres = self.create_or_open_presentation(context, file)
        pres.NewWindow()
        
    
    # ==========================
    # = Quick add to favorites =
    # ==========================

    # def add_fav_to_config(self):
    #     if not self.copy_shapes_setting:
    #         folders = bkt.config.chart_library_folders or []
    #         if self.fav_folder not in folders:
    #             folders.append(self.fav_folder)
    #             self.library_folders.insert(0, self.fav_folder)
    #         bkt.config.set_smart("chart_library_folders", folders)
    #     else:
    #         folders = bkt.config.shape_library_folders or []
    #         if self.fav_folder not in folders:
    #             folders.append(self.fav_folder)
    #             self.library_folders.insert(0, self.fav_folder)
    #         bkt.config.set_smart("shape_library_folders", folders)

    @classmethod
    def add_slides_to_lib(cls, context, presentation):
        #Copy slides
        context.selection.SlideRange.Copy()
        # pres.Slides.Paste()
        cls._save_paste(presentation.Slides)

    @classmethod
    def add_shapes_to_lib(cls, context, presentation):
        #Copy each shape individually
        for shape in context.shapes:
            title = bkt.ui.show_user_input("Bitte Shape-Titel eingeben:", "Shape-Titel", shape.Name)
            if title is None:
                break
            shape.Copy()
            slide = presentation.Slides.Add(presentation.Slides.Count+1, 11) #11=ppTitleOnly
            
            #set slide background color
            slide.FollowMasterBackground = False
            orig_bg = context.slides[0].Background.Fill.ForeColor
            if orig_bg.Type == pplib.MsoColorType['msoColorTypeScheme'] and orig_bg.ObjectThemeColor > 0:
                slide.Background.Fill.ForeColor.ObjectThemeColor = orig_bg.ObjectThemeColor
                slide.Background.Fill.ForeColor.Brightness = orig_bg.Brightness
            else:
                slide.Background.Fill.ForeColor.RGB = orig_bg.RGB
            
            # new_shp = slide.Shapes.Paste()
            new_shp = cls._save_paste(slide.Shapes)
            new_shp.Left = (presentation.PageSetup.SlideWidth - new_shp.Width)*0.5
            new_shp.Top  = (presentation.PageSetup.SlideHeight - new_shp.Height)*0.5
            slide.Shapes.Title.Textframe.TextRange.Text = title


    # def add_chart_to_file(self, context, current_control):
    #     '''Add current shape/slide to selected library'''
    #     if self.copy_shapes_setting and len(context.shapes) == 0:
    #         return bkt.message("Keine Shapes ausgewählt!")

    #     filename = current_control['tag']
    #     self.add_chart_to_lib(context, filename)

    def add_chart_to_lib(self, context):
        '''Add current shape/slide to favorites library'''
        if self.copy_shapes_setting and len(context.shapes) == 0:
            return bkt.message("Keine Shapes ausgewählt!")

        #Open default file
        #FIXME: read list of files in fav folder and ask user to select file
        file = os.path.join(self.fav_folder, "Favorites.pptx")
        
        pres = self.create_or_open_presentation(context, file)
        
        try:
            if not self.copy_shapes_setting:
                self.add_slides_to_lib(context, pres)
            else:
                self.add_shapes_to_lib(context, pres)

            pres.Save()
        except:
            logging.exception("error adding chart to library")
            bkt.message.error("Fehler beim Hinzufügen zur Library", "BKT: ChartLib")
            # bkt.helpers.exception_as_message()
        finally:
            pres.Saved = True
            pres.Close()

        #Regenerate thumbnails
        gallery = self.get_chartlib_gallery_from_file(file)
        gallery.reset_gallery_items(context)
        self.reset_cashes()


    
    # ====================================
    # = Menu-items for ChartLib-sections =
    # ====================================
    
    def reset_chartlib_menu_callback(self, context, current_control):
        ''' callback to reload menu for presentation-file represented by control '''
        return self.reset_chartlib_menu_from_file(context, current_control['tag'])
        
    def reset_chartlib_menu_from_file(self, context, filename):
        ''' reloads menu for presentation-file. resets caches and relads chartlib-menu '''
        self.cached_presentation_menus.pop(filename, None)
        self.get_chartlib_menu_from_file(context, filename)
        return None
    
    def get_chartlib_menu_callback(self, context, current_control):
        ''' callback for dynamic menu, return Menu-control for presentation represented by control '''
        return self.get_chartlib_menu_from_file(context, current_control['tag'])
    
    def get_chartlib_menu_from_file(self, context, filename):
        ''' returns static menu for presentation-file. uses cached menu or generates menu using get_chartlib_menu_from_presentation '''
        if not self.cached_presentation_menus.has_key(filename):
            # presentation = self.open_presentation_file(context, filename)
            with open_presentation_without_window(context, filename) as presentation:
                menu = self.get_chartlib_menu_from_presentation(presentation)
                # presentation.Close()
                self.cached_presentation_menus[filename] = menu
        # else:
        #     logging.debug('get_chartlib_menu_from_presentation: reuse menu for %s', filename)
        return self.cached_presentation_menus[filename]
    
    def get_chartlib_menu_from_presentation(self, presentation):
        ''' returns static menu for presentation. 
            Opens prensentation and creates menu-buttons for every slide.
            If presentation has sections, these are reused to structure menu:
            either by menu-sections or sub-menus; latter if presentation has 40 slides or more.
        '''
        children = [ ]
        
        # items per section
        num_sections = presentation.sectionProperties.Count
        if num_sections == 0:
            # only one section, list slide-Buttons
            children += self.get_section_menu(presentation.slides, 1, presentation.slides.count)
        elif presentation.Slides.Count < 40:
            # list-seperator per section, with slide-Buttons
            for idx in range(1,num_sections+1):
                children += [ bkt.ribbon.MenuSeparator(title = presentation.sectionProperties.Name(idx)) ]
                children += self.get_section_menu(presentation.slides,
                         presentation.sectionProperties.FirstSlide(idx),
                         presentation.sectionProperties.SlidesCount(idx) )
        else:
            # list Menu per section, with slide-Buttons
            children += [
                bkt.ribbon.Menu(
                    label = presentation.sectionProperties.Name(idx),
                    children = self.get_section_menu(presentation.slides,
                         presentation.sectionProperties.FirstSlide(idx),
                         presentation.sectionProperties.SlidesCount(idx) )
                )
                for idx in range(1,num_sections+1)
            ]

        # open-file-Button
        children += self.get_file_buttons(presentation.FullName)
        
        # return Menu
        menu = bkt.ribbon.Menu(
            #label=presentation.Name,
            xmlns="http://schemas.microsoft.com/office/2009/07/customui",
            id=None,
            children = children
        )
        return menu
    
    
    def get_chartlib_gallery_from_file(self, filename):
        ''' returns dynamic gallery for presentation-file. uses cached gallery or generates gallery using ChartLibGallery class '''
        if not self.cached_presentation_galleries.has_key(filename):
            gallery = ChartLibGallery(filename, copy_shapes=self.copy_shapes_setting)
            self.cached_presentation_galleries[filename] = gallery
        return self.cached_presentation_galleries[filename]
    
    
    
        
    # ==================================
    # = Menu-items for ChartLib-slides =
    # ==================================
    
    def get_section_menu(self, slides, offset, count):
        ''' return menu-buttons for given slide selection '''
        
        # control chars are removed from labels
        control_chars = dict.fromkeys(range(32))
        
        return [
            bkt.ribbon.Button(
                label = slides.item(idx).Shapes.Title.Textframe.TextRange.text[:40].translate(control_chars) if slides.item(idx).Shapes.Title.Textframe.TextRange.text else "slide" + str(idx),
                tag=slides.parent.FullName + "|" + str(idx),
                on_action=bkt.Callback(self.slide_action, context=True, current_control=True)
            )
            for idx in range(offset, offset+count)
            if slides.item(idx).shapes.hastitle
        ]
    


    
    # ===========
    # = actions =
    # ===========
    
    def get_file_buttons(self, filename, show_menu_separator=True):
        return ([bkt.ribbon.MenuSeparator(title = "Library")] if show_menu_separator else []) + [
            bkt.ribbon.Button(label="Datei öffnen und Library bearbeiten",
                image_mso='FileSaveAsPowerPointPptx',
                tag=filename,
                on_action=bkt.Callback(self.open_file, context=True, current_control=True)
            ),
            bkt.ribbon.Button(label="Datei erneut indizieren",
                image_mso='AccessRefreshAllLists',
                tag=filename,
                on_action=bkt.Callback(self.reset_chartlib_menu_callback, context=True, current_control=True)
            )
        ]
    
    @classmethod
    def open_file(cls, context, current_control):
        ''' Open library file with window and write access '''
        filename = current_control['tag']
        # Parameter: readonly, untitled, withwindow
        context.app.Presentations.Open(filename, False, False, True)
    
    @classmethod
    def apply_template_to_slides(cls, context, current_control):
        ''' Apply template to presentation '''
        filename = current_control['tag']
        context.selection.SlideRange.ApplyTemplate(filename)
    
    @classmethod
    def replace_template_in_presentation(cls, context, current_control):
        ''' Apply template to presentation '''
        filename = current_control['tag']
        context.presentation.ApplyTemplate(filename)
    
    @classmethod
    def add_template_to_presentation(cls, context, current_control):
        ''' Apply template to presentation '''
        filename = current_control['tag']
        context.presentation.Designs.Load(filename)
    
    @classmethod
    def new_presentation_with_template(cls, context, current_control):
        ''' Open presentation untitled '''
        filename = current_control['tag']
        # Parameter: readonly, untitled, withwindow
        context.app.Presentations.Open(filename, False, True, True)
    
    @classmethod
    def copy_slide_callback(cls, context, current_control):
        filename, slide_index = current_control['tag'].split("|")
        cls.copy_slide(context, filename, slide_index)
        
    @classmethod
    def copy_slide(cls, context, filename, slide_index):
        ''' Copy slide from chart lib '''
        # active_window = context.app.ActiveWindow
        # # open presentation
        # # template_presentation = cls.open_presentation_file(context, filename)
        # with open_presentation_without_window(context, filename) as template_presentation:
        #     # copy slide
        #     orig_slide = template_presentation.slides.item(int(slide_index))
        #     orig_slide.copy()
        #     # paste slide
        #     position = active_window.View.Slide.SlideIndex
        #     # active_window.presentation.slides.paste(position+1)
        #     slide = cls._save_paste(active_window.presentation.slides, position+1)
        #     # template_presentation.Close()
        
        active_window = context.app.ActiveWindow
        try:
            position = active_window.View.Slide.SlideIndex
        except:
            #fallback is not slide in view, e.g. selection within two slides in sorter
            position = 0
        slide_index = int(slide_index)
        active_window.presentation.slides.InsertFromFile(filename, position, slide_index, slide_index)
        active_window.View.GotoSlide(position+1)
    
    @classmethod
    def copy_shapes_callback(cls, context, current_control):
        filename, slide_index = current_control['tag'].split("|")
        cls.copy_shapes(context, filename, slide_index)
        
    @classmethod
    def copy_shapes(cls, context, filename, slide_index):
        ''' Copy shape from shape lib '''
        # open presentation
        # template_presentation = cls.open_presentation_file(context, filename)
        with open_presentation_without_window(context, filename) as template_presentation:
            template_slide = template_presentation.slides.item(int(slide_index))
            # find relevant shapes
            shape_indices = []
            shape_index = 1
            for shape in template_slide.shapes:
                if shape.type != 14 and shape.visible == -1:
                    # shape is not a placeholder and visible
                    shape_indices.append(shape_index)
                shape_index+=1
            # select and copy shapes
            template_slide.shapes.Range(Array[int](shape_indices)).copy()
            # current slide
            # cur_slide = context.app.activeWindow.View.Slide
            cur_slides = context.slides
            do_select = len(cur_slides) == 1
            for cur_slide in cur_slides:
                shape_count = cur_slide.shapes.count
                # cur_slide.shapes.paste()
                cls._save_paste(cur_slide.shapes)
                new_shape_count = cur_slide.shapes.count
                # group+select shapes
                if new_shape_count - shape_count > 1:
                    shape = cur_slide.shapes.Range(Array[int](range(shape_count+1, new_shape_count+1))).group()
                    shape.select()
                elif do_select:
                    cur_slide.shapes.item(new_shape_count).select()
            # template_presentation.Close()




class ChartLibGallery(bkt.ribbon.Gallery):
    
    # size of items
    # item_width should not be used in chart lib (copy_shapes=False); will be overwritten with respect to slide-format
    item_height = 150
    item_width = 200

    #translation table
    transtab = None

    @classmethod
    def filename_as_ui_id(cls, filename):
        ''' creates customui-id from filename by removing unsupported characters '''
        if not cls.transtab:
            # characters to remove
            chr_numbers = range(45) + range(46,48) + range(58,65) + range(91,95) + [96] + range(123,256)
            # translation tab
            cls.transtab = maketrans("".join( chr(i) for i in chr_numbers ), '|'*len(chr_numbers))
        
        return 'id_' + filename.translate(cls.transtab).replace('|', '__')
    
    
    def __init__(self, filename, copy_shapes=False, **user_kwargs):
        '''Constructor
           Initializes Gallery for chart/shape-library
        '''
        self.filename = filename
        self.thumb_dict = os.path.splitext(filename)[0] + THUMBNAIL_POSTFIX
        self.items_initialized = False
        
        self.labels = []
        self.slide_indices = []
        
        # parent_id = user_kwargs.get('id') or ""
        
        self.this_id = ChartLibGallery.filename_as_ui_id(filename) + "--" + str(copy_shapes)
        
        self.cache = bkt.helpers.caches.get("chartlib")
        
        # default settings
        kwargs = dict(
            #label = u"Test Gallery",
            id = self.this_id,
            label = os.path.splitext(os.path.basename(filename))[0],
            show_label=False,
            #screentip="",
            #supertip="",
            get_image=bkt.Callback(lambda: self.get_chartlib_item_image(0)),
        )
        
        self.copy_shapes = copy_shapes
        if self.copy_shapes:
            # shape lib settings
            kwargs.update(dict(
                item_height=50,
                item_width=50,
                columns=3
            ))
        else:
            # slide lib settings
            kwargs.update(dict(
                item_height=100,
                #item_width=177, item_height=100, # 16:9
                #item_width=133, item_height=100, # 4:3
                columns=2
            ))
        
        # user kwargs
        if len(user_kwargs) > 0:
            kwargs.update(user_kwargs)
        
        # initialize gallery
        super(ChartLibGallery, self).__init__(
            children = self.get_file_buttons(),
            **kwargs
        )
    
    
    
    # ================
    # = ui callbacks =
    # ================
    
    def on_action_indexed(self, selected_item, index, context, current_control):
        '''CustomUI-callback: copy shape/slide from library to current presentation'''
        if self.copy_shapes:
            # copy shapes / shape library
            ChartLib.copy_shapes(context, self.filename, self.slide_indices[index])
        else:
            # copy slide / slide library
            ChartLib.copy_slide(context, self.filename, self.slide_indices[index])
    
    def get_item_count(self, context):
        '''CustomUI-callback'''
        if not self.items_initialized:
            self.init_gallery_items(context, closing_gallery_workaround=True)
        return len(self.labels)

    def get_item_height(self):
        '''CustomUI-callback: return item-height
           item-height is initialized by init_gallery_items
        '''
        return self.item_height

    def get_item_width(self):
        '''CustomUI-callback: return item-width
           item-width is initialized by init_gallery_items
        '''
        return self.item_width
    
    def get_item_id(self, index):
        '''CustomUI-callback: returns corresponding item id'''
        return self.this_id + "--" + str(index)
        
    def get_item_label(self, index, context):
        '''CustomUI-callback: returns corresponding item label
           labels are initialized by init_gallery_items
        '''
        if not self.items_initialized:
            self.init_gallery_items(context, closing_gallery_workaround=True)
        return self.labels[index][:40]
    
    def get_item_screentip(self, index):
        '''CustomUI-callback'''
        #return "Shape aus Shape-library einfügen"
        if self.copy_shapes:
            # copy shapes / shape library
            return "Shape »" + self.labels[index]  + "« aus Shape-Library einfügen"
        else:
            # copy slide / slide library
            return "Folie »" + self.labels[index]  + "« aus Chart-Library einfügen"
        
    # def get_item_supertip(self, index):
    #     return "tbd"
    
    def get_item_image(self, index, context):
        '''CustomUI-callback: calls get_chartlib_item_image'''
        if not self.items_initialized:
            self.init_gallery_items(context, closing_gallery_workaround=True)
        return self.get_chartlib_item_image(index)
    
    
    
    # ===========
    # = methods =
    # ===========
    
    def get_file_buttons(self):
        return [
            bkt.ribbon.Button(label="Datei öffnen und Library bearbeiten",
                image_mso='FileSaveAsPowerPointPptx',
                tag=self.filename,
                on_action=bkt.Callback(ChartLib.open_file, context=True, current_control=True)
            ),
            bkt.ribbon.Button(label="Markierte Shapes zu Datei hinzufügen" if self.copy_shapes else "Markierte Folien zu Datei hinzufügen",
                image_mso='AddToFavorites',
                on_action=bkt.Callback(self.add_charts, context=True)
            ),
            bkt.ribbon.Button(label="Datei neu indizieren und Thumbnails aktualisieren",
                image_mso='AccessRefreshAllLists',
                on_action=bkt.Callback(self.reset_gallery_items, context=True)
            )
        ]
    
    def add_charts(self, context):
        '''Add selected slides/shapes to this gallery'''
        if self.copy_shapes and len(context.shapes) == 0:
            return bkt.message("Keine Shapes ausgewählt!")

        with open_presentation_without_window(context, self.filename, False) as presentation:
            try:
                if not self.copy_shapes:
                    ChartLib.add_slides_to_lib(context, presentation)
                else:
                    ChartLib.add_shapes_to_lib(context, presentation)

                presentation.Save()
            except:
                logging.exception("error adding chart to library")
                bkt.message.error("Fehler beim Hinzufügen zur Library", "BKT: ChartLib")

        #Regenerate thumbnails
        self.reset_gallery_items(context)
    
    def reset_gallery_items(self, context):
        '''Forces Gallery to re-initialize and generate thumbnail-images'''
        # reset gallery in a separate thread so core thread is still able to handle the app-events (i.e. presentation open/close).
        # otherwise when this function is called in a loop, some presentations remain open and block ppt process from being quit.
        #NOTE: Use of with-statement is very important so comrelease does not release any com objects while stuff is going on in separate thread!
        # with context.app:
        t = Thread(target=self.init_gallery_items, args=(context, True))
        t.start()
        t.join()
        # self.init_gallery_items(context, force_thumbnail_generation=True)
    
    def init_gallery_items(self, context, force_thumbnail_generation=False, closing_gallery_workaround=False):
        try:
            if force_thumbnail_generation:
                raise KeyError("Thumbnail generation not possible from cache")
            self.init_gallery_items_from_cache()
        except KeyError as e:
            if str(e) == "CACHE_FILEMTIME_INVALID":
                force_thumbnail_generation = True
            with open_presentation_without_window(context, self.filename) as presentation:
                try:
                    self.init_gallery_items_from_presentation(presentation, force_thumbnail_generation=force_thumbnail_generation, closing_gallery_workaround=closing_gallery_workaround)
                except:
                    logging.exception('error initializing gallery')
        
        # try:
        #     presentation = Charlib.open_presentation_file(context, self.filename)
        #     # presentation = context.app.Presentations.Open(self.filename, True, False, True) #filename, readonly, untitled, withwindow
        #     self.init_gallery_items_from_presentation(presentation, force_thumbnail_generation=force_thumbnail_generation)
        # except:
        #     logging.exception('error initializing gallery')
        # finally:
        #     # logging.debug('closing presentation: %s', self.filename)
        #     presentation.Saved = True
        #     presentation.Close()
        
    def init_gallery_items_from_cache(self):
        ''' initialize gallery items from cache if file has not been modified.
        '''
        cache = self.cache[self.filename]
        if os.path.getmtime(self.filename) != cache["file_mtime"]:
            raise KeyError("CACHE_FILEMTIME_INVALID")
        self.item_width     = cache["item_width"]
        self.slide_indices  = cache["slide_indices"]
        self.labels         = cache["labels"]
        self.item_count     = cache["item_count"]
        self.items_initialized = True
    
    def init_gallery_items_from_presentation(self, presentation, force_thumbnail_generation=False, closing_gallery_workaround=False):
        ''' initialize gallery items (count, labels, item-widht, item-height... ).
            Also generates thumbnail-image-files if needed.
        '''
        
        # item width should respect slide-format
        if not self.copy_shapes:
            self.item_width = int(self.item_height * presentation.SlideMaster.Width / presentation.SlideMaster.Height)
    
        # init labels
        control_chars = dict.fromkeys(range(32))
        self.slide_indices = []
        self.labels = []
        for slide in presentation.slides:
            if slide.shapes.hastitle:
                self.slide_indices.append(slide.SlideIndex)

                label = slide.Shapes.Title.Textframe.TextRange.text
                if not label:
                    label = "slide %s" % slide.SlideIndex
                else:
                    label = label.translate(control_chars)
                self.labels.append(label)

        # pres_slides = presentation.Slides
        # item_count = pres_slides.Count
        # self.slide_indices = [
        #     idx
        #     for idx in range(1, 1+item_count)
        #     if pres_slides.item(idx).shapes.hastitle != False
        # ]
        # self.labels = [
        #     # pres_slides.item(idx).Shapes.Title.Textframe.TextRange.text[:40].translate(control_chars) 
        #     pres_slides.item(idx).Shapes.Title.Textframe.TextRange.text.translate(control_chars) 
        #     if pres_slides.item(idx).Shapes.Title.Textframe.TextRange.text != "" else "slide" + str(idx)
        #     for idx in self.slide_indices
        # ]
        # logging.debug("labels %s", self.labels)

        # init items
        self.item_count = len(self.labels)

        # items are initialized and can be displayed
        self.items_initialized = True

        #create cache
        self.cache[self.filename] = dict(
            # cache_time      = time.time(),
            file_mtime      = os.path.getmtime(self.filename),
            item_width      = self.item_width,
            slide_indices   = self.slide_indices,
            labels          = self.labels,
            item_count      = self.item_count,
        )
        self.cache.sync()

        # init images, if first thumbnail does not exist
        image_filename = self.get_image_filename(1)
        if force_thumbnail_generation or not os.path.exists(image_filename):
            if self.copy_shapes:
                self.generate_gallery_images_from_shapes(presentation)
            else:
                self.generate_gallery_images_from_slides(presentation, closing_gallery_workaround)
        
        # # cache items for next call on library
        # # image loading still necessary
        # self.children = [
        #     bkt.ribbon.Item(
        #         id=self.get_item_id(idx),
        #         label=self.get_item_label(idx),
        #         #image=self.get_item_image(idx),
        #     )
        #     for idx in range(self.get_item_count(context=None))
        # ] + self.children
        
    
    
    def get_image_filename(self, index, postfix=""):
        ''' returns path of thumbnail for given chart-lib slide '''
        return os.path.join(self.thumb_dict, ''.join([str(index),postfix,'.png']))
    
    def get_image_filename_for_index(self, index, postfix=""):
        ''' returns path of thumbnail for given chart-lib slide '''
        return self.get_image_filename(self.slide_indices[index], postfix)
    
    
    def generate_gallery_images_from_slides(self, presentation, closing_gallery_workaround=False):
        ''' generate thumbnail images for all chart-lib items '''
        # filename = presentation.FullName
        # item_count = presentation.Slides.Count
        # control_chars = dict.fromkeys(range(32))
        
        # make sure, directory exists
        # directory = os.path.split( self.get_image_filename(1) )[0]
        directory = self.thumb_dict
        if not os.path.exists(directory):
            os.makedirs(directory)
        
        for slide in presentation.slides:
            if slide.shapes.hastitle != False:
                image_filename = self.get_image_filename(slide.SlideIndex)
                try:
                    #NOTE: slide.Export() closes the chartlib menu, so every time the images are generated, the user needs to re-open the chartlib. As workaround we use the clipboard.
                    if closing_gallery_workaround:
                        logging.debug("Creation of thumbnail image via clipboard workaround")
                        slide.Copy()
                        Forms.Clipboard.GetImage().Save(image_filename, Drawing.Imaging.ImageFormat.Png)
                    else:
                        logging.debug("Creation of thumbnail image via export")
                        slide.Export(image_filename, 'PNG', 600) 
                except:
                    logging.exception('Creation of thumbnail image failed: %s', image_filename)
        if closing_gallery_workaround:
            Forms.Clipboard.Clear()
    
    
    def generate_gallery_images_from_shapes(self, presentation, with_placeholders=False):
        ''' generate thumbnail images for all shape-lib items '''
        # filename = presentation.FullName
        # item_count = presentation.Slides.Count
        # control_chars = dict.fromkeys(range(32))
        
        # make sure, directory exists
        # directory = os.path.split( self.get_image_filename(1) )[0]
        directory = self.thumb_dict
        if not os.path.exists(directory):
            os.makedirs(directory)
        
        for slide in presentation.slides:
            if slide.shapes.hastitle:
                # find relevant shapes
                shape_indices = []
                shape_index = 1
                for shape in slide.shapes:
                    if shape.visible == -1 and (with_placeholders or shape.type != 14):
                        # shape is visible and not a placeholder (or placeholder are allowed)
                        shape_indices.append(shape_index)
                    shape_index+=1
                # select shapes
                shape_range = slide.shapes.Range(Array[int](shape_indices))
                image_filename = self.get_image_filename(slide.SlideIndex)
                # WAS: image_filename = os.path.join(os.path.splitext(self.filename)[0], str(slide.SlideIndex) + '.png')
                
                try:
                    # export image as PNG, 
                    # ppShapeFormatGIF = 0, ppShapeFormatJPG = 1, ppShapeFormatPNG = 2, ppShapeFormatBMP = 3, ppShapeFormatWMF = 4, ppShapeFormatEMF = 5;
                    shape_range.Export(image_filename, 2)
                except:
                    logging.exception('Creation of thumbnail image failed: %s', image_filename)
                
                # resize thumbnail image to square
                if os.path.exists(image_filename):
                    try:
                        transparent = slide.Background.Fill.Transparency == 1.0
                        if not transparent:
                            background_color = slide.Background.Fill.ForeColor.RGB
                        else:
                            background_color = None

                        cropped_image_filename = self.get_image_filename(slide.SlideIndex, postfix="-cropped")
                        original_image_filename = self.get_image_filename(slide.SlideIndex, postfix="-original")

                        glib.make_thumbnail(image_filename, 100, 100, cropped_image_filename, background_color)

                        # move files
                        if os.path.exists(original_image_filename):
                            os.remove(original_image_filename)
                        os.rename(image_filename, original_image_filename)
                        os.rename(cropped_image_filename, image_filename)
                    except:
                        logging.exception('Creation of croped thumbnail image failed: %s', image_filename)
    
    
    def get_chartlib_item_image(self, index):
        ''' load item image from corresponding image-file and return Bitmap-object '''
        #logging.debug('get_chartlib_item_image %s -- %s', filename, index)
        try:
            image_filename = self.get_image_filename_for_index(index)
        except IndexError:
            #IndexError occurs for gallery image before init_gallery was called
            image_filename = self.get_image_filename(index+1)
        #logging.debug('get_chartlib_item_image %s', image_filename)
        if os.path.exists(image_filename):
            # return empty from file
            #return Bitmap.FromFile(image_filename)
            
            #version that should not lock the file, which prevents updating of thumbnails:
            return glib.open_bitmap_nonblocking(image_filename)
        else:
            # return empty white image
            return glib.empty_image(50,50)





charts = ChartLib()
# charts.library_folders.extend( bkt.config.chart_library_folders or [] )
# charts.library_files.extend( bkt.config.chart_libraries or [] )

shapes = ChartLib( copy_shapes=True )
# shapes.library_folders.extend( bkt.config.shape_library_folders or [] )
# shapes.library_files.extend( bkt.config.shape_libraries or [] )


# add from feature-folders
# for folder in bkt.config.feature_folders:
#     chartlib_folder = os.path.join(folder, "chartlib")
#     charts.library_folders.append( { 'title':os.path.basename(os.path.realpath(folder)), 'folder':chartlib_folder})
#     shapelib_folder = os.path.join(folder, "shapelib")
#     shapes.library_folders.append( { 'title':os.path.basename(os.path.realpath(folder)), 'folder':shapelib_folder})



chartlib_button = bkt.ribbon.DynamicMenu(
    id='menu-add-chart',
    label="Templatefolie einfügen",
    show_label=False,
    screentip="Folie aus Slide-Library einfügen",
    supertip="Aus den hinterlegten Slide-Templates kann ein Template als neue Folie eingefügt werden.",
    image_mso="BibliographyGallery",
    # image_mso="SlideMasterInsertLayout",
    #image_mso="CreateFormBlankForm",
    get_content = bkt.Callback(
        charts.get_root_menu
    )
)
shapelib_button = bkt.ribbon.DynamicMenu(
    id='menu-add-shape',
    label="Personal Shape Library",
    show_label=False,
    screentip="Shape aus Shape-Library einfügen",
    supertip="Aus den hinterlegten Shape-Templates kann ein Shape auf die aktuelle Folie eingefügt werden.",
    image_mso="ActionInsert",
    #image_mso="ShapesInsertGallery",
    #image_mso="OfficeExtensionsGallery",
    get_content = bkt.Callback(
        shapes.get_root_menu
    )
)

# chartlibgroup = bkt.ribbon.Group(
#     label="chartlib",
#     children=[ chartlib_button, shapelib_button]
# )

# bkt.powerpoint.add_tab(
#     bkt.ribbon.Tab(
#         label="chartlib",
#         children = [
#             chartlibgroup
#         ]
#     )
# )





