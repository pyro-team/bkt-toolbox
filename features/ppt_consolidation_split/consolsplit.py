# -*- coding: utf-8 -*-
'''
Created on 2017-11-09
@author: Florian Stallmann
'''

import bkt

MODEL_MODULE = __package__ + ".consolsplit_model"
MODEL_CONTAINER = "ConsolSplit"

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
                bkt.ribbon.Group(id="bkt_consolsplit_consolidate_group", label="Folien aus Dateien anfügen", children=[
                    bkt.ribbon.PrimaryItem(children=[
                        bkt.ribbon.Button(
                            label="Folien aus Dateien anfügen",
                            supertip="Alle Folien aus mehreren PowerPoint-Dateien an diese Präsentation anfügen",
                            image_mso='ThemeBrowseForThemes',
                            on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "consolidate_ppt_slides", application=True, presentation=True),
                            is_definitive=True,
                        ),
                    ]),
                    bkt.ribbon.TopItems(children=[
                        bkt.ribbon.Label(label="Alle Folien aus mehreren PowerPoint-Dateien an diese Präsentation anfügen."),
                        bkt.ribbon.Label(label="Dieser Vorgang kann bei großen Dateien und vielen Folien einige Zeit in Anspruch nehmen!"),
                    ]),
                ]),
                bkt.ribbon.Group(id="bkt_consolsplit_pic2slides_group", label="Folien aus Bildern erstellen", children=[
                    bkt.ribbon.PrimaryItem(children=[
                        bkt.ribbon.Menu(
                            label="Folien aus Bildern erstellen",
                            supertip="Mehrere Bild-Dateien (jpg, png, emf) auf jeweils eine Folie einfügen",
                            image_mso='PhotoGalleryProperties',
                            children=[
                                bkt.ribbon.MenuGroup(
                                    item_size="large",
                                    children=[
                                        bkt.ribbon.Button(
                                            label="Bild-Dateien auswählen",
                                            image_mso='PhotoGalleryProperties',
                                            description="Alle Bild-Dateien zum Einfügen einzeln auswählen.",
                                            on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "pictures_to_slides", context=True),
                                            is_definitive=True,
                                        ),
                                        bkt.ribbon.Button(
                                            label="Ordner mit Bildern auswählen",
                                            image_mso='OpenFolder',
                                            description="Ordner mit Bild-Dateien zum Einfügen auswählen.",
                                            on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "folder_to_slides", context=True),
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
                bkt.ribbon.Group(id="bkt_consolsplit_split_group", label="Folien einzeln speichern", children=[
                    bkt.ribbon.PrimaryItem(children=[
                        bkt.ribbon.Menu(
                            label="Folien einzeln speichern",
                            supertip="Alle Folien in einzelne PowerPoint-Dateien im gewählten Ordner speichern",
                            image_mso='ThemeSaveCurrent',
                            children=[
                                bkt.ribbon.MenuGroup(
                                    item_size="large",
                                    children=[
                                        bkt.ribbon.Button(
                                            label="Folien einzeln speichern",
                                            image_mso='ThemeSaveCurrent',
                                            description="Jede Folie einzeln speichern. Die Dateien werden mit Foliennummer nummeriert und nach Folientitel benannt.",
                                            on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "split_slides_to_ppt", context=True, slides=True),
                                            is_definitive=True,
                                        ),
                                        bkt.ribbon.Button(
                                            label="Abschnitte einzeln speichern",
                                            image_mso='SectionAdd',
                                            description="Jeden Abschnitt einzeln speichern. Die Dateien werden nummeriert und nach Abschnittstitel benannt.",
                                            on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "split_sections_to_ppt", context=True, slides=True),
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