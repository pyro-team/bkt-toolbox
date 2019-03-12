# -*- coding: utf-8 -*-
'''
Created on 2017-11-09
@author: Florian Stallmann
'''

import bkt
import os
from System import Array #SlideRange

class ConsolSplit(object):

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
        newPres = application.Presentations.Open(full_name)

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
    def split_slides_to_ppt(cls, application, presentation, slides):
        # if not presentation.Path:
        #     bkt.helpers.message("Bitte erst Datei speichern.")
        #     return

        save_pattern = "[slidenumber]_[slidetitle]"
        
        fileDialog = application.FileDialog(4) #msoFileDialogFolderPicker
        if presentation.Path:
            fileDialog.InitialFileName = presentation.Path + '\\'
        fileDialog.title = "Ordner zum Speichern auswählen"

        if fileDialog.Show() == 0: #msoFalse
            return
        folder = fileDialog.SelectedItems(1)
        if not os.path.isdir(folder):
            return

        def _get_name(slide):
            try:
                title = slide.Shapes.Title.TextFrame.TextRange.Text
                for c in '\/:*?"<>|': #delete not allowd characters
                    title = title.replace(c, '')
                title = title.strip()
            except:
                title = "UNKNOWN"
            
            filename = save_pattern.replace("[slidenumber]", str(slide.SlideIndex)).replace("[slidetitle]", title[:32]) #max 32 characters of title
            return os.path.join(folder, filename + ".pptx") #FIXME: file ending according to current presentation

        # for slide in slides:
        for slide in presentation.Slides:
            cls.export_slide(application, [slide], _get_name(slide))

        bkt.helpers.message("Export abgeschlossen")
    
    @classmethod
    def split_sections_to_ppt(cls, application, presentation, slides):
        # if not presentation.Path:
        #     bkt.helpers.message("Bitte erst Datei speichern.")
        #     return

        sections = presentation.SectionProperties
        if sections.count < 2:
            bkt.helpers.message("Präsentation hat weniger als 2 Abschnitte!")
            return

        save_pattern = "[sectionnumber]_[sectiontitle]"
        # save_pattern = "[sectiontitle]"
        
        fileDialog = application.FileDialog(4) #msoFileDialogFolderPicker
        if presentation.Path:
            fileDialog.InitialFileName = presentation.Path + '\\'
        fileDialog.title = "Ordner zum Speichern auswählen"

        if fileDialog.Show() == 0: #msoFalse
            return
        folder = fileDialog.SelectedItems(1)
        if not os.path.isdir(folder):
            return

        def _get_name(index):
            try:
                title = sections.Name(index)
                for c in '\/:*?"<>|': #delete not allowd characters
                    title = title.replace(c, '')
                title = title.strip()
            except:
                title = "UNKNOWN"
            
            filename = save_pattern.replace("[sectionnumber]", str(index)).replace("[sectiontitle]", title[:32]) #max 32 characters of title
            return os.path.join(folder, filename + ".pptx") #FIXME: file ending according to current presentation

        for i in range(sections.count):
            start = sections.FirstSlide(i+1)
            if start == -1:
                continue #empty section
            count = sections.SlidesCount(i+1)
            slides = list(iter( presentation.Slides.Range( Array[int](range(start, start+count)) ) ))
            cls.export_slide(application, slides, _get_name(i+1))

        bkt.helpers.message("Export abgeschlossen")


# consolsplit_gruppe = bkt.ribbon.Group(
#     label='Konsolidieren && Teilen',
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
        insertAfterMso="TabInfo",
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
                                            on_action=bkt.Callback(ConsolSplit.split_slides_to_ppt, application=True, presentation=True, slides=True),
                                            is_definitive=True,
                                        ),
                                        bkt.ribbon.Button(
                                            label="Abschnitte einzeln speichern",
                                            image_mso='ThemeSaveCurrent',
                                            description="Jeden Abschnitt einzeln speichern. Die Dateien werden nummeriert und nach Abschnittstitel benannt.",
                                            on_action=bkt.Callback(ConsolSplit.split_sections_to_ppt, application=True, presentation=True, slides=True),
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
                ])
            ])
        ]
    )
)