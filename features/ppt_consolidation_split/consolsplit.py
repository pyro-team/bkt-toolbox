# -*- coding: utf-8 -*-
'''
Created on 2017-11-09
@author: Florian Stallmann
'''

import bkt
import bkt.library.powerpoint as pplib

import logging
import os
from string import maketrans
from System import Array #SlideRange

class ConsolSplit(object):
    trans_table = maketrans('\t\n\r\f\v', '     ')

    @classmethod
    def consolidate_ppt_slides(cls, application, presentation):
        fileDialog = application.FileDialog(3) #msoFileDialogFilePicker
        fileDialog.Filters.Add("PowerPoint", "*.ppt; *.pptx; *.pptm", 1)
        fileDialog.Filters.Add("Alle Dateien", "*.*", 2)
        if presentation.Path:
            fileDialog.InitialFileName = presentation.Path + '\\'
        fileDialog.title = "PowerPoint-Dateien auswählen"
        fileDialog.AllowMultiSelect = True

        if fileDialog.Show() == 0: #msoFalse
            return

        for file in list(iter(fileDialog.SelectedItems)):
            if file == presentation.FullName:
                continue
            try:
                presentation.Slides.InsertFromFile(file, presentation.Slides.Count)
            except:
                pass


    @classmethod
    def export_slide(cls, application, slides, full_name):
        slides[0].Parent.SaveCopyAs(full_name)
        newPres = application.Presentations.Open(full_name, False, False, False) #readonly, untitled, withwindow

        slideIds = [slide.SlideIndex for slide in slides]
        removeIds = list(set(range(1,newPres.Slides.Count+1)) - set(slideIds))
        removeIds.sort()
        removeIds.reverse()
        for slideId in removeIds:
            newPres.Slides(slideId).Delete()

        #remove all sections
        sections = newPres.SectionProperties
        for i in reversed(range(sections.count)):
            sections.Delete(i+1, 0) #index, deleteSlides=False

        newPres.Save()
        newPres.Close()
    
    @classmethod
    def _get_safe_filename(cls, title):
        title = title.encode('ascii', 'ignore') #remove unicode characters
        title = title.translate(cls.trans_table, '\/:*?"<>|') #replace special whitespace chacaters with space, also delete not allowed characters
        title = title[:32] #max 32 characters of title
        title = title.strip() #remove whitespaces at beginning and end
        return title
    
    @classmethod
    def split_slides_to_ppt(cls, context, slides):
        # if not presentation.Path:
        #     bkt.helpers.message("Bitte erst Datei speichern.")
        #     return

        application = context.app
        presentation = context.presentation

        # save_pattern = "[slidenumber]_[slidetitle]"
        
        fileDialog = application.FileDialog(4) #msoFileDialogFolderPicker
        if presentation.Path:
            fileDialog.InitialFileName = presentation.Path + '\\'
        fileDialog.title = "Ordner zum Speichern auswählen"

        if fileDialog.Show() == 0: #msoFalse
            return
        folder = fileDialog.SelectedItems(1)
        if not os.path.isdir(folder):
            return

        save_pattern = bkt.ui.show_user_input("Bitte Dateinamen-Pattern eingeben:", "Dateiname eingeben", "[slidenumber]_[slidetitle]")
        if save_pattern is None:
            return

        def _get_name(slide):
            try:
                title = cls._get_safe_filename(slide.Shapes.Title.TextFrame.TextRange.Text)
            except:
                title = "UNKNOWN"
            
            filename = save_pattern.replace("[slidenumber]", str(slide.SlideIndex)).replace("[slidetitle]", title)
            return os.path.join(folder, filename + ".pptx") #FIXME: file ending according to current presentation

        def loop(worker):
            error = False
            slides_current = 1.0
            slides_total = presentation.Slides.Count
            worker.ReportProgress(0)
            for slide in presentation.Slides:
                if worker.CancellationPending:
                    break
                worker.ReportProgress(slides_current/slides_total*100)
                slides_current += 1.0
                try:
                    cls.export_slide(application, [slide], _get_name(slide))
                except Exception as e:
                    logging.error("split_slides_to_ppt error %r" % e)
                    error = True

            worker.ReportProgress(100)
            if worker.CancellationPending:
                bkt.helpers.message("Export durch Nutzer abgebrochen")
            elif error:
                bkt.helpers.message("Export mit Fehlern abgeschlossen")
            else:
                bkt.helpers.message("Export erfolgreich abgeschlossen")
            os.startfile(folder)

        bkt.ui.execute_with_progress_bar(loop, context, modal=False) #modal=False important so main thread can handle app events and all presentations close properly
    
    @classmethod
    def split_sections_to_ppt(cls, context, slides):
        # if not presentation.Path:
        #     bkt.helpers.message("Bitte erst Datei speichern.")
        #     return

        application = context.app
        presentation = context.presentation

        sections = presentation.SectionProperties
        if sections.count < 2:
            bkt.helpers.message("Präsentation hat weniger als 2 Abschnitte!")
            return

        # save_pattern = "[sectionnumber]_[sectiontitle]"
        
        fileDialog = application.FileDialog(4) #msoFileDialogFolderPicker
        if presentation.Path:
            fileDialog.InitialFileName = presentation.Path + '\\'
        fileDialog.title = "Ordner zum Speichern auswählen"

        if fileDialog.Show() == 0: #msoFalse
            return
        folder = fileDialog.SelectedItems(1)
        if not os.path.isdir(folder):
            return

        save_pattern = bkt.ui.show_user_input("Bitte Dateinamen-Pattern eingeben:", "Dateiname eingeben", "[sectionnumber]_[sectiontitle]")
        if save_pattern is None:
            return

        def _get_name(index):
            try:
                title = cls._get_safe_filename(sections.Name(index))
            except:
                title = "UNKNOWN"
            
            filename = save_pattern.replace("[sectionnumber]", str(index)).replace("[sectiontitle]", title)
            return os.path.join(folder, filename + ".pptx") #FIXME: file ending according to current presentation

        def loop(worker):
            error = False
            sections_current = 1.0
            sections_total = sections.count
            worker.ReportProgress(0)
            for i in range(sections.count):
                if worker.CancellationPending:
                    break
                worker.ReportProgress(sections_current/sections_total*100)
                sections_current += 1.0
                try:
                    start = sections.FirstSlide(i+1)
                    if start == -1:
                        continue #empty section
                    count = sections.SlidesCount(i+1)
                    slides = list(iter( presentation.Slides.Range( Array[int](range(start, start+count)) ) ))
                    cls.export_slide(application, slides, _get_name(i+1))
                except Exception as e:
                    logging.error("split_sections_to_ppt error %r" % e)
                    error = True
                    continue

            worker.ReportProgress(100)
            if worker.CancellationPending:
                bkt.helpers.message("Export durch Nutzer abgebrochen")
            elif error:
                bkt.helpers.message("Export mit Fehlern abgeschlossen")
            else:
                bkt.helpers.message("Export erfolgreich abgeschlossen")
            os.startfile(folder)

        bkt.ui.execute_with_progress_bar(loop, context, modal=False) #modal=False important so main thread can handle app events and all presentations close properly


class FolderToSlides(object):
    filetypes = [".jpg", ".png", ".emf"]

    @classmethod
    def _files_to_slides(cls, context, all_files):
        master_slide_new = False
        try:
            master_slide = context.slide
        except EnvironmentError:
            #nothing appropriate selected
            master_slide = context.presentation.slides.add(1, 11) #ppLayoutTitleOnly
            master_slide_new = True
        
        ref_frame = pplib.BoundingFrame(master_slide, contentarea=True)

        # all_files.sort(reverse=True)
        for root, full_path in reversed(all_files):
            #create slide
            new_slide = master_slide.Duplicate()
            # new_slide.Layout = 11 #ppLayoutTitleOnly
            #add title
            if new_slide.Shapes.HasTitle:
                new_slide.Shapes.Title.Textframe.TextRange.Text = root
            #paste slide
            try:
                new_pic = new_slide.Shapes.AddPicture(full_path, 0, -1, ref_frame.left, ref_frame.top) #FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height
            except:
                new_slide.Delete()
            else:
                if new_pic.width > ref_frame.width:
                    new_pic.width = ref_frame.width
                if new_pic.height > ref_frame.height:
                    new_pic.height = ref_frame.height
        
        if master_slide_new:
            master_slide.Delete()

    @classmethod
    def folder_to_slides(cls, context):
        fileDialog = context.app.FileDialog(4) #msoFileDialogFolderPicker
        if context.presentation.Path:
            fileDialog.InitialFileName = context.presentation.Path + '\\'
        fileDialog.title = "Ordner mit Bildern auswählen"

        if fileDialog.Show() == 0: #msoFalse
            return
        folder = fileDialog.SelectedItems(1)
        if not os.path.isdir(folder):
            return
        
        all_files = []

        for file in os.listdir(folder):
            full_path = os.path.join(folder, file)
            root,ext = os.path.splitext(file)
            if ext in cls.filetypes:
                all_files.append((root, full_path))

        cls._files_to_slides(context, sorted(all_files))

    @classmethod
    def pictures_to_slides(cls, context):
        fileDialog = context.app.FileDialog(3) #msoFileDialogFilePicker
        fileDialog.Filters.Add("Bilder", "; ".join(["*"+f for f in cls.filetypes]), 1)
        fileDialog.Filters.Add("SVG", "*.svg", 2)
        fileDialog.Filters.Add("Alle Dateien", "*.*", 3)
        if context.presentation.Path:
            fileDialog.InitialFileName = context.presentation.Path + '\\'
        fileDialog.title = "Bild-Dateien auswählen"
        fileDialog.AllowMultiSelect = True

        if fileDialog.Show() == 0: #msoFalse
            return

        def get_root(path):
            _,file = os.path.split(path)
            root,_ = os.path.splitext(file)
            return root

        all_files = [(get_root(file), file) for file in fileDialog.SelectedItems]
        cls._files_to_slides(context, all_files)


# consolsplit_gruppe = bkt.ribbon.Group(
#     label='Konsolidieren & Teilen',
#     image_mso='ThemeBrowseForThemes',
#     children = [
#         bkt.ribbon.Button(
#             id = 'consolidate_ppt_slides',
#             label="Folien aus Dateien anfügen",
#             show_label=True,
#             size="large",
#             image_mso='ThemeBrowseForThemes',
#             supertip="Alle Folien aus mehreren PowerPoint-Dateien an diese Präsentation anfügen.",
#             on_action=bkt.Callback(ConsolSplit.consolidate_ppt_slides, application=True, presentation=True),
#         ),
#         bkt.ribbon.Menu(
#             id = 'split_slides_to_ppt',
#             image_mso='ThemeSaveCurrent',
#             label="Folien einzeln speichern",
#             supertip="Folien in einzelne PowerPoint-Dateien speichern.",
#             show_label=True,
#             size="large",
#             item_size="large",
#             children=[
#                 bkt.ribbon.Button(
#                     # id = 'split_slides_to_ppt',
#                     label="Folien einzeln speichern",
#                     image_mso='ThemeSaveCurrent',
#                     description="Jede Folie als einzelne Datei im gewählten Ordner speichern. Die Dateien werden mit Foliennummer nummeriert und nach Folientitel benannt.",
#                     on_action=bkt.Callback(ConsolSplit.split_slides_to_ppt, application=True, presentation=True, slides=True),
#                 ),
#                 bkt.ribbon.Button(
#                     # id = 'split_slides_to_ppt',
#                     label="Abschnitte einzeln speichern",
#                     image_mso='ThemeSaveCurrent',
#                     description="Jeden Abschnitt als einzelne Datei im gewählten Ordner speichern. Die Dateien werden nummeriert und nach Abschnittstitel benannt.",
#                     on_action=bkt.Callback(ConsolSplit.split_sections_to_ppt, application=True, presentation=True, slides=True),
#                 ),
#             ]
#         )
#     ]
# )

# bkt.powerpoint.add_tab(bkt.ribbon.Tab(
#     id="bkt_powerpoint_toolbox_extensions",
#     insert_before_mso="TabHome",
#     label=u'Toolbox 3/3',
#     # get_visible defaults to False during async-startup
#     get_visible=bkt.Callback(lambda:True),
#     children = [
#         consolsplit_gruppe,
#     ]
# ), extend=True)



bkt.powerpoint.add_backstage_control(
    bkt.ribbon.Tab(
        label="Konsol./Teilen",
        title="BKT - Dateien Konsolidieren & Teilen",
        insertAfterMso="TabPublish", #http://youpresent.co.uk/customising-powerpoint-2016-backstage-view/
        columnWidthPercent="50",
        children=[
            bkt.ribbon.FirstColumn(children=[
                bkt.ribbon.Group(label="Folien aus Dateien anfügen", children=[
                    bkt.ribbon.PrimaryItem(children=[
                        bkt.ribbon.Button(
                            label="Folien aus Dateien anfügen",
                            image_mso='ThemeBrowseForThemes',
                            on_action=bkt.Callback(ConsolSplit.consolidate_ppt_slides, application=True, presentation=True),
                            is_definitive=True,
                        ),
                    ]),
                    bkt.ribbon.TopItems(children=[
                        bkt.ribbon.Label(label="Alle Folien aus mehreren PowerPoint-Dateien an diese Präsentation anfügen."),
                        bkt.ribbon.Label(label="Dieser Vorgang kann bei großen Dateien und vielen Folien einige Zeit in Anspruch nehmen!"),
                    ]),
                ]),
                bkt.ribbon.Group(label="Folien aus Bildern erstellen", children=[
                    bkt.ribbon.PrimaryItem(children=[
                        bkt.ribbon.Menu(
                            label="Folien aus Bildern erstellen",
                            image_mso='PhotoGalleryProperties',
                            children=[
                                bkt.ribbon.MenuGroup(
                                    item_size="large",
                                    children=[
                                        bkt.ribbon.Button(
                                            label="Bild-Dateien auswählen",
                                            image_mso='PhotoGalleryProperties',
                                            description="Alle Bild-Dateien zum Einfügen einzeln auswählen.",
                                            on_action=bkt.Callback(FolderToSlides.pictures_to_slides, context=True),
                                            is_definitive=True,
                                        ),
                                        bkt.ribbon.Button(
                                            label="Ordner mit Bildern auswählen",
                                            image_mso='OpenFolder',
                                            description="Ordner mit Bild-Dateien zum Einfügen auswählen.",
                                            on_action=bkt.Callback(FolderToSlides.folder_to_slides, context=True),
                                            is_definitive=True,
                                        ),
                                    ]
                                ),
                            ]
                        ),
                    ]),
                    bkt.ribbon.TopItems(children=[
                        bkt.ribbon.Label(label="Mehrere Bild-Dateien (jpg, png, emf) auf jeweils eine Folie einfügen."),
                        bkt.ribbon.Label(label="Die ausgewählte Folie wird für jede Bild-Datei dupliziert und der Dateiname als Titel gesetzt."),
                    ]),
                ]),
                bkt.ribbon.Group(label="Folien einzeln speichern", children=[
                    bkt.ribbon.PrimaryItem(children=[
                        bkt.ribbon.Menu(
                            label="Folien einzeln speichern",
                            image_mso='ThemeSaveCurrent',
                            children=[
                                bkt.ribbon.MenuGroup(
                                    item_size="large",
                                    children=[
                                        bkt.ribbon.Button(
                                            label="Folien einzeln speichern",
                                            image_mso='ThemeSaveCurrent',
                                            description="Jede Folie einzeln speichern. Die Dateien werden mit Foliennummer nummeriert und nach Folientitel benannt.",
                                            on_action=bkt.Callback(ConsolSplit.split_slides_to_ppt, context=True, slides=True),
                                            is_definitive=True,
                                        ),
                                        bkt.ribbon.Button(
                                            label="Abschnitte einzeln speichern",
                                            image_mso='SectionAdd',
                                            description="Jeden Abschnitt einzeln speichern. Die Dateien werden nummeriert und nach Abschnittstitel benannt.",
                                            on_action=bkt.Callback(ConsolSplit.split_sections_to_ppt, context=True, slides=True),
                                            is_definitive=True,
                                        ),
                                    ]
                                ),
                            ]
                        ),
                    ]),
                    bkt.ribbon.TopItems(children=[
                        bkt.ribbon.Label(label="Alle Folien in einzelne PowerPoint-Dateien im gewählten Ordner speichern."),
                        bkt.ribbon.Label(label="Dieser Vorgang kann bei großen Dateien und vielen Folien einige Zeit in Anspruch nehmen!"),
                    ]),
                ]),
            ])
        ]
    )
)