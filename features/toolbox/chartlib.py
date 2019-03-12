# -*- coding: utf-8 -*-
'''
Created on 04.05.2016

@author: rdebeerst
'''


import bkt
import bkt.library.powerpoint as pplib

import os
# import os.path
import System
        
import logging
import traceback

from bkt import dotnet
Drawing = dotnet.import_drawing()
Bitmap = Drawing.Bitmap
ColorTranslator = Drawing.ColorTranslator

Forms = dotnet.import_forms() #required for clipboard functions



# TODO
# ChartLib --> ChartLibMenu
# ChartLib class with
#    file-buttons
#    copy shape/slide action

THUMBNAIL_POSTFIX = '_thumbnails'


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
        self.fav_folder = bkt.helpers.get_fav_folder()
        
        # init slide_action
        self.copy_shapes_setting = copy_shapes
        if copy_shapes:
            # copy shapes
            self.slide_action = self.copy_shapes_callback
            self.fav_folder = os.path.join(self.fav_folder, "shapelib")
        else:
            # copy full slide
            self.slide_action = self.copy_slide_callback
            self.fav_folder = os.path.join(self.fav_folder, "chartlib")
        
        # add favorite folder as first folder
        self.library_folders.insert(0, {'title': "Favoriten", 'folder': self.fav_folder} )
    
    
    
    # ===========
    # = Helpers =
    # ===========
    
    @classmethod
    def open_presentation_file(cls, context, file):
        ''' opens and returns presentation file '''
        # if self.presentations.has_key(file):
        #     presentation = self.presentations[file]
        # else:
        #     # # Parameter: schreibgeschützt, ohne Titel, kein Fenster
        #     # presentation = context.app.Presentations.Open(file, True, False, False)
        #     # Parameter: rw-access, ohne Titel, kein Fenster
        #     presentation = context.app.Presentations.Open(file, False, False, False)
        #     self.presentations[file] = presentation
        return context.app.Presentations.Open(file, True, False, False)
    
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
        
        # create menu-items for chartlibrary-directories
        if len(self.library_folders) == 0:
            # create empty menu
            menu = bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None
            )
        elif len(self.library_folders) == 1:
            # create menu with contents of directory
            folder = self.library_folders[0]
            if type(folder) == dict:
                folder = folder['folder']
            menu = self.get_folder_menu(folder, id=None)
            menu.label=None
        else:
            # create menu with sections for each directory
            folders = [ folder if type(folder) == dict else  {'folder':folder, 'title':os.path.basename(folder)}   for folder in self.library_folders]
            logging.debug("ChartLib root menu with folders: %s" % folders)
            
            children = [ self.get_folder_menu(folder['folder'], label=folder['title']) for folder in folders]
            # make list flat 
            flat_children = []
            for folder_menu in children:
                if len(folder_menu.children) > 0:
                    flat_children.append( bkt.ribbon.MenuSeparator(title=folder_menu['label'] or None) )
                    flat_children += folder_menu.children
            
            menu = bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None,
                children=flat_children
            )
            
            
        # create menu-items for chartlibrary-files
        menu.children.append( bkt.ribbon.MenuSeparator() )
        for filename in self.library_files:
            menu.children.append( self.get_dynamic_file_menu(filename) )
        
        # buttons
        menu.children += [
            bkt.ribbon.MenuSeparator(title="Settings"),
            bkt.ribbon.Button(label="Zu Favoriten hinzufügen",
                screentip="Chart zu Favoriten-Library hinzufügen",
                supertip="Ausgewählte Slides/Shapes zu Standard-Favoriten-Library hinzufügen. Falls diese Library noch nicht existiert, wird diese neu angelegt.",
                image_mso='SourceControlCheckIn',
                on_action=bkt.Callback(self.add_chart_to_lib)
            ),
            bkt.ribbon.Button(label="Library erneut indizieren",
                screentip="Library erneut indizieren",
                supertip="Alle Cashes werden zurückgesetzt. Library-Order und -Dateien werden jeweils beim nächsten Öffnen erneut indiziert.\n\nEs werden nur Thumbnails bereits geöffneter Libraries aktualisiert. Eine Neugenerierung der Thumbnails kann auch für jede Shape-Library-Datei separat angestoßen werden (im jeweiligen Menü).",
                image_mso='AccessRefreshAllLists',
                on_action=bkt.Callback(self.update_thumbnails_and_reset_cashes)
            )
        ]
        
        logging.debug("ChartLib root menu: %s" % menu.xml())
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
        #     logging.debug('get_folder_menu_callback: reuse menu for %s' % folder)
        return self.cached_presentation_menus[folder]
    
    def get_folder_menu(self, folder, **kwargs):
        ''' returns static menu-control for given folder. menu contains entries for subfolders and pptx-files. '''

        files = []
        subfolders = []
        
        if os.path.isdir(folder):
            # find pptx-Files
            for filename in os.listdir(folder):
                if filename.endswith(".pptx") and not filename.startswith("~$"):
                    files.append(filename)
            logging.debug('get_folder_menu files: %s' % files)
        
            # find subfolders
            for filename in os.listdir(folder):
                if os.path.isdir(folder + "\\" + filename) == True:
                    if not filename.endswith(THUMBNAIL_POSTFIX):
                        subfolders.append(filename)
            logging.debug('get_folder_menu folders: %s' % subfolders)
            
        else:
            logging.warning("Chartlib. No such directory: %s" % folder)
        
        # FIXME: no folders which belong to files
        
        if len(files) + len(subfolders) == 0:
            children = []
        else:
            # create DynamicMenus / Galleries
            children = [
                self.get_chartlib_gallery_from_file(folder + "\\" + filename) if self.show_gallery else self.get_dynamic_file_menu(folder + "\\" + filename) 
                for filename in files
            ] + ([bkt.ribbon.MenuSeparator()] if len(files) > 0 else [] )+ [
                self.get_dynamic_folder_menu(folder + "\\" + subfolder)
                for subfolder in subfolders
            ]
        
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
    
    def update_thumbnails_and_reset_cashes(self, context):
        if not bkt.helpers.confirmation("Dieser Vorgang kann bei vielen Libraries einige Minuten dauern und nicht abgebrochen werden. Trotzdem fortsetzen?"):
            return
        
        for gal in self.cached_presentation_galleries.itervalues():
            gal.reset_gallery_items(context)
        self.reset_cashes()

    def reset_cashes(self):
        ''' cashes for library menus and galleries are deleted '''
        self.cached_presentation_menus = {}
        self.cached_presentation_galleries = {}
        
        
    
    # ====================================
    # = Quick add to favorites =
    # ====================================

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

    def add_chart_to_lib(self, context):
        #Open default file
        #FIXME: read list of files in fav folder and ask user to select file
        file = os.path.join(self.fav_folder, "Favorites.pptx")
        pres = self.create_or_open_presentation(context, file)
        #Ensure fav folder is in config
        # self.add_fav_to_config()
        
        try:
            if not self.copy_shapes_setting:
                #Copy slides
                context.selection.SlideRange.Copy()
                pres.Slides.Paste()
            else:
                #Copy each shape individually
                for shape in context.shapes:
                    title = bkt.ui.show_user_input("Bitte Shape-Titel eingeben:", "Shape-Titel")
                    if title == None:
                        break
                    shape.Copy()
                    slide = pres.Slides.Add(pres.Slides.Count+1, 11) #11=ppTitleOnly
                    
                    #set slide background color
                    slide.FollowMasterBackground = False
                    orig_bg = context.slides[0].Background.Fill.ForeColor
                    if orig_bg.Type == pplib.MsoColorType['msoColorTypeScheme'] and orig_bg.ObjectThemeColor > 0:
                        slide.Background.Fill.ForeColor.ObjectThemeColor = orig_bg.ObjectThemeColor
                        slide.Background.Fill.ForeColor.Brightness = orig_bg.Brightness
                    else:
                        slide.Background.Fill.ForeColor.RGB = orig_bg.RGB
                    
                    new_shp = slide.Shapes.Paste()
                    new_shp.Left = (pres.PageSetup.SlideWidth - new_shp.Width)*0.5
                    new_shp.Top  = (pres.PageSetup.SlideHeight - new_shp.Height)*0.5
                    slide.Shapes.Title.Textframe.TextRange.Text = title

            pres.Save()
        except:
            bkt.helpers.exception_as_message()
        finally:
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
            presentation = self.open_presentation_file(context, filename)
            menu = self.get_chartlib_menu_from_presentation(presentation)
            presentation.Close()
            self.cached_presentation_menus[filename] = menu
        # else:
        #     logging.debug('get_chartlib_menu_from_presentation: reuse menu for %s' % filename)
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
                label = slides.item(idx).Shapes.Title.Textframe.TextRange.text[:40].translate(control_chars) if slides.item(idx).Shapes.Title.Textframe.TextRange.text != "" else "slide" + str(idx),
                tag=slides.parent.FullName + "|" + str(idx),
                on_action=bkt.Callback(self.slide_action, context=True, current_control=True)
            )
            for idx in range(offset, offset+count)
            if slides.item(idx).shapes.hastitle != False
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
        # Parameter: rw-access, ohne Titel, mit Fenster
        presentation = context.app.Presentations.Open(filename, False, False, True)
    
    @classmethod
    def copy_slide_callback(cls, context, current_control):
        filename, slide_index = current_control['tag'].split("|")
        cls.copy_slide(context, filename, slide_index)
        
    @classmethod
    def copy_slide(cls, context, filename, slide_index):
        ''' Copy slide from chart lib '''
        # open presentation
        template_presentation = cls.open_presentation_file(context, filename)
        # copy slide
        template_presentation.slides.item(int(slide_index)).copy()
        # paste slide
        position = context.app.activeWindow.View.Slide.SlideIndex
        context.app.activeWindow.presentation.slides.paste(position+1)
        template_presentation.Close()
    
    @classmethod
    def copy_shapes_callback(cls, context, current_control):
        filename, slide_index = current_control['tag'].split("|")
        cls.copy_shapes(context, filename, slide_index)
        
    @classmethod
    def copy_shapes(cls, context, filename, slide_index):
        ''' Copy shape from shape lib '''
        # open presentation
        template_presentation = cls.open_presentation_file(context, filename)
        template_slide = template_presentation.slides.item(int(slide_index))
        # current slide
        cur_slide = context.app.activeWindow.View.Slide
        shape_count = cur_slide.shapes.count
        # find relevant shapes
        shape_indices = []
        shape_index = 1
        for shape in template_slide.shapes:
            if shape.type != 14 and shape.visible == -1:
                # shape is not a placeholder and visible
                shape_indices.append(shape_index)
            shape_index+=1
        # select and copy shapes
        template_slide.shapes.Range(System.Array[int](shape_indices)).copy()
        cur_slide.shapes.paste()
        
        # group+select shapes
        if cur_slide.shapes.count - shape_count > 1:
            cur_slide.shapes.Range(System.Array[int](range(shape_count+1, cur_slide.shapes.count+1))).group().select()
        else:
            cur_slide.shapes.item(cur_slide.shapes.count).select()
        template_presentation.Close()
    
    
    


def filename_as_ui_id(filename):
    from string import maketrans
    ''' creates customui-id from filename by removing unsupported characters '''
    # characters to remove
    chr_numbers = range(45) + range(46,48) + range(58,65) + range(91,95) + [96] + range(123,256)
    # translation tab
    transtab = maketrans("".join( chr(i) for i in chr_numbers ), '|'*len(chr_numbers))
    
    return 'id_' + filename.translate(transtab).replace('|', '__')



class ChartLibGallery(bkt.ribbon.Gallery):
    
    # size of items
    # item_width should not be used in chart lib (copy_shapes=False); will be overwritten with respect to slide-format
    item_height = 150
    item_width = 200
    
    def __init__(self, filename, copy_shapes=False, **user_kwargs):
        '''Constructor
           Initializes Gallery for chart/shape-library
        '''
        self.filename = filename
        self.items_initialized = False
        
        self.labels = []
        self.slide_indices = []
        
        parent_id = user_kwargs.get('id') or ""
        
        self.this_id = filename_as_ui_id(filename) + "--" + str(copy_shapes)
        
        
        
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
            self.init_gallery_items(context)
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
        
    def get_item_label(self, index):
        '''CustomUI-callback: returns corresponding item label
           labels are initialized by init_gallery_items
        '''
        if not self.items_initialized:
            self.init_gallery_items(context)
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
    
    def get_item_image(self, index):
        '''CustomUI-callback: calls get_chartlib_item_image'''
        if not self.items_initialized:
            self.init_gallery_items(context)
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
            bkt.ribbon.Button(label="Datei erneut indizieren und Thumbnails aktualisieren",
                image_mso='AccessRefreshAllLists',
                on_action=bkt.Callback(self.reset_gallery_items, context=True)
            )
        ]
    
    def reset_gallery_items(self, context):
        '''Forces Gallery to re-initialize and generate thumbnail-images'''
        self.init_gallery_items(context, force_thumbnail_generation=True)
    
    def init_gallery_items(self, context, force_thumbnail_generation=False):
        # FIXME: use general presentation_open method
        try:
            presentation = context.app.Presentations.Open(self.filename, True, False, False)
            self.init_gallery_items_from_presentation(presentation, force_thumbnail_generation=force_thumbnail_generation)
        except:
            logging.error('error initializing gallery')
            logging.debug(traceback.format_exc())
        finally:
            presentation.Close()
        
    
    def init_gallery_items_from_presentation(self, presentation, force_thumbnail_generation=False):
        ''' initialize gallery items (count, labels, item-widht, item-height... ).
            Also generates thumbnail-image-files if needed.
        '''
        
        # item width should respect slide-format
        if not self.copy_shapes:
            self.item_width = int(self.item_height * presentation.SlideMaster.Width / presentation.SlideMaster.Height)
    
        # init labels
        control_chars = dict.fromkeys(range(32))
        item_count = presentation.Slides.Count
        self.slide_indices = [
            idx
            for idx in range(1, 1+item_count)
            if presentation.slides.item(idx).shapes.hastitle != False
        ]
        self.labels = [
            # presentation.slides.item(idx).Shapes.Title.Textframe.TextRange.text[:40].translate(control_chars) 
            presentation.slides.item(idx).Shapes.Title.Textframe.TextRange.text.translate(control_chars) 
            if presentation.slides.item(idx).Shapes.Title.Textframe.TextRange.text != "" else "slide" + str(idx)
            for idx in self.slide_indices
        ]
        # logging.debug("labels %s" % self.labels)

        # init items
        self.item_count = len(self.labels)

        # items are initialized and can be displayed
        self.items_initialized = True

        # init images, if first thumbnail does not exist
        image_filename = self.get_image_filename(1)
        if force_thumbnail_generation or not os.path.exists(image_filename):
            if self.copy_shapes:
                self.generate_gallery_images_from_shapes(presentation)
            else:
                self.generate_gallery_images_from_slides(presentation)
        
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
        return os.path.join(os.path.splitext(self.filename)[0] + THUMBNAIL_POSTFIX, str(index) + postfix + '.png')
    
    
    def generate_gallery_images_from_slides(self, presentation):
        ''' generate thumbnail images for all chart-lib items '''
        filename = presentation.FullName
        item_count = presentation.Slides.Count
        control_chars = dict.fromkeys(range(32))
        
        # make sure, directory exists
        directory = os.path.split( self.get_image_filename(1) )[0]
        if not os.path.exists(directory):
            os.makedirs(directory)
        
        for slide in presentation.slides:
            if slide.shapes.hastitle != False:
                # select shapes
                # slide_range = presentation.slides.Range(System.Array[int]([ slide.SlideIndex ]))
                image_filename = self.get_image_filename(slide.SlideIndex)
                try:
                    # export image as PNG,  2 = ppShapeFormatPNG, 0 = ppShapeFormatGIF
                    # width 600, auto height (450 for 4:3)
                    # slide_range.Export(image_filename, 'PNG', 600)
                    # slide.Export(image_filename, 'PNG', 600) #FIXME: this line closes the chartlib menu, so every time the images are generated, the user needs to re-open the chartlib
                    slide.Copy()
                    Forms.Clipboard.GetImage().Save(image_filename, Drawing.Imaging.ImageFormat.Png)
                except:
                    logging.warning('Creation of thumbnail image failed: %s' % image_filename)
        Forms.Clipboard.Clear()
    
    
    def generate_gallery_images_from_shapes(self, presentation, with_placeholders=False):
        ''' generate thumbnail images for all shape-lib items '''
        filename = presentation.FullName
        item_count = presentation.Slides.Count
        control_chars = dict.fromkeys(range(32))
        
        # make sure, directory exists
        directory = os.path.split( self.get_image_filename(1) )[0]
        if not os.path.exists(directory):
            os.makedirs(directory)
        
        for slide in presentation.slides:
            if slide.shapes.hastitle != False:
                # find relevant shapes
                shape_indices = []
                shape_index = 1
                for shape in slide.shapes:
                    if shape.visible == -1 and (with_placeholders or shape.type != 14):
                        # shape is visible and not a placeholder (or placeholder are allowed)
                        shape_indices.append(shape_index)
                    shape_index+=1
                # select shapes
                shape_range = slide.shapes.Range(System.Array[int](shape_indices))
                image_filename = self.get_image_filename(slide.SlideIndex)
                # WAS: image_filename = os.path.join(os.path.splitext(self.filename)[0], str(slide.SlideIndex) + '.png')
                
                try:
                    # export image as PNG, 
                    # ppShapeFormatGIF = 0, ppShapeFormatJPG = 1, ppShapeFormatPNG = 2, ppShapeFormatBMP = 3, ppShapeFormatWMF = 4, ppShapeFormatEMF = 5;
                    shape_range.Export(image_filename, 2) 
                except:
                    logging.warning('Creation of thumbnail image failed: %s' % image_filename)
                
                # resize thumbnail image to square
                if os.path.exists(image_filename):
                    try:
                        # init croped image
                        width = 100
                        height = 100
                        image = Bitmap(image_filename)
                        bmp = Bitmap(width, height)
                        graph = Drawing.Graphics.FromImage(bmp)
                        #set background color of thumbnails
                        background_color = Drawing.ColorTranslator.FromOle(slide.Background.Fill.ForeColor.RGB)
                        if background_color != pplib.PowerPoint.XlRgbColor.rgbWhite.value__:
                            graph.Clear(background_color)
                        # compute scale
                        scale = min(float(width) / image.Width, float(height) / image.Height)
                        scaleWidth = int(image.Width * scale)
                        scaleHeight = int(image.Height * scale)
                        # set quality
                        graph.InterpolationMode  = Drawing.Drawing2D.InterpolationMode.High
                        graph.CompositingQuality = Drawing.Drawing2D.CompositingQuality.HighQuality
                        graph.SmoothingMode      = Drawing.Drawing2D.SmoothingMode.AntiAlias
                        # redraw and save
                        # logging.debug('crop image from %sx%s to %sx%s. rect %s.%s-%sx%s' % (image.Width, image.Height, width, height, int((width - scaleWidth)/2), int((height - scaleHeight)/2), scaleWidth, scaleHeight))
                        graph.DrawImage(image, Drawing.Rectangle(int((width - scaleWidth)/2), int((height - scaleHeight)/2), scaleWidth, scaleHeight))
                        cropped_image_filename = self.get_image_filename(slide.SlideIndex, postfix="-cropped")
                        bmp.Save(cropped_image_filename)
                        # close files
                        image.Dispose()
                        bmp.Dispose()
                        # move files
                        original_image_filename = self.get_image_filename(slide.SlideIndex, postfix="-original")
                        if os.path.exists(original_image_filename):
                            os.remove(original_image_filename)
                        os.rename(image_filename, original_image_filename)
                        os.rename(cropped_image_filename, image_filename)
                        
                    except:
                        logging.error('Creation of croped thumbnail image failed: %s' % image_filename)
                        logging.debug(traceback.format_exc())
                    finally:
                        if image:
                            image.Dispose()
                        if bmp:
                            bmp.Dispose()
    
    
    def get_chartlib_item_image(self, index):
        ''' load item image from corresponding image-file and return Bitmap-object '''
        #logging.debug('get_chartlib_item_image %s -- %s' % (filename, index))
        image_filename = self.get_image_filename(index+1)
        #logging.debug('get_chartlib_item_image %s' % (image_filename))
        if os.path.exists(image_filename):
            # return empty from file
            #return Bitmap.FromFile(image_filename)
            
            #version that should not lock the file, which prevents updating of thumbnails:
            with Bitmap.FromFile(image_filename) as img:
                new_img = Bitmap(img)
                img.Dispose()
                return new_img
        else:
            # return empty white image
            img = Bitmap(50, 50)
            color = ColorTranslator.FromHtml('#ffffff00')
            img.SetPixel(0, 0, color);
            return img





charts = ChartLib()
#charts.root_dir = "S:\\Tooling\\Toolbox-git\\_personal\\chartlib"
#charts.root_dir=os.path.join(os.path.dirname(os.path.realpath(__file__)), "..", "..", "_personal", "chartlib")
charts.library_folders.extend( bkt.config.chart_library_folders or [] )
charts.library_files.extend( bkt.config.chart_libraries or [] )

shapes = ChartLib( copy_shapes=True )
#shapes.root_dir = "S:\\Tooling\\Toolbox-git\\_personal\\shapelib"
#shapes.root_dir=os.path.join(os.path.dirname(os.path.realpath(__file__)), "..", "..", "_personal", "shapelib")
shapes.library_folders.extend( bkt.config.shape_library_folders or [] )
shapes.library_files.extend( bkt.config.shape_libraries or [] )


# add from feature-folders
for folder in bkt.config.feature_folders:
    chartlib_folder = os.path.join(folder, "chartlib")
    charts.library_folders.append( { 'title':os.path.basename(os.path.realpath(folder)), 'folder':chartlib_folder})
    shapelib_folder = os.path.join(folder, "shapelib")
    shapes.library_folders.append( { 'title':os.path.basename(os.path.realpath(folder)), 'folder':shapelib_folder})


#bkt.helpers.message(charts.library_folders)
#bkt.helpers.message(shapes.library_folders)




chartlib_button = bkt.ribbon.DynamicMenu(
    id='menu-add-chart',
    label="Templatefolie einfügen",
    show_label=False,
    screentip="Folie aus Slide-Library einfügen",
    supertip="Aus den hinterlegten Slide-Templates kann ein Template als neue Folie eingefügt werden.",
    image_mso="SlideMasterInsertLayout",
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

chartlibgroup = bkt.ribbon.Group(
    label="chartlib",
    children=[ chartlib_button, shapelib_button]
)

# bkt.powerpoint.add_tab(
#     bkt.ribbon.Tab(
#         label="chartlib",
#         children = [
#             chartlibgroup
#         ]
#     )
# )





