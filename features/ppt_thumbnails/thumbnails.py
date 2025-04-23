# -*- coding: utf-8 -*-
'''
Created on 2017-07-24
@author: Florian Stallmann
'''

import bkt
import bkt.library.powerpoint as pplib

import bkt.dotnet as dotnet
Forms = dotnet.import_forms() #required to read clipboard and open file dialogs


MODEL_MODULE = __package__ + ".thumbnails_model"
MODEL_CONTAINER = "Thumbnailer"


BKT_THUMBNAIL = "BKT_THUMBNAIL"


class ThumbnailerUi(object):

    @classmethod
    def has_clipboard_data(cls):
        return Forms.Clipboard.ContainsData(BKT_THUMBNAIL) or (Forms.Clipboard.ContainsData("PowerPoint 12.0 Internal Slides") and Forms.Clipboard.ContainsData("Link Source")) #"PowerPoint 14.0 Slides Package"
        # return Forms.Clipboard.ContainsData(BKT_THUMBNAIL)

    @classmethod
    def enabled_paste(cls):
        return cls.has_clipboard_data()
        #return Forms.Clipboard.ContainsImage()

    @classmethod
    def is_thumbnail(cls, shape):
        return pplib.TagHelper.has_tag(shape, BKT_THUMBNAIL)



thumbnail_gruppe = bkt.ribbon.Group(
    id="bkt_slidethumbnails_group",
    label='Folien-Thumbnails',
    supertip="Ermöglicht das Einfügen von aktualisierbaren Folien-Thumbnails. Das Feature `ppt_thumbnails` muss installiert sein.",
    image_mso='PasteAsPicture',
    children = [
        bkt.ribbon.Button(
            id = 'slide_copy',
            label="Folie(n) als Thumbnail kopieren",
            show_label=True,
            image_mso='Copy',
            supertip="Aktuelle Folie zum Erstellen vom aktualisierbaren Folien-Thumbnails kopieren.",
            on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "slides_copy", presentation=True, slides=True),
        ),
        # bkt.ribbon.Button(
        #     id = 'shape_copy',
        #     label="Shape als Thumbnail kopieren",
        #     show_label=True,
        #     image_mso='Copy',
        #     supertip="Aktuelle Folie zum Erstellen vom aktualisierbaren Folien-Thumbnails kopieren.",
        #     on_action=bkt.Callback(Thumbnailer.shape_copy, presentation=True, slide=True, shape=True),
        # ),
        bkt.ribbon.SplitButton(
            get_enabled = bkt.Callback(ThumbnailerUi.enabled_paste),
            children=[
                bkt.ribbon.Button(
                    id = 'slide_paste',
                    label="Folien-Thumbnail einfügen",
                    show_label=True,
                    image_mso='PasteAsPicture',
                    supertip="Kopierte Folie als aktualisierbares Thumbnail mit Referenz auf Originalfolie einfügen.\n\nIst die Originalfolie aus einer anderen Datei im gleichen Verzeichnis, wird nur der Dateiname hinterlegt, anderenfalls wird der absolute Pfad hinterlegt und die Originaldatei darf nicht verschoben werden.",
                    on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "slide_paste", application=True),
                    # get_enabled = bkt.Callback(Thumbnailer.enabled_paste),
                ),
                bkt.ribbon.Menu(label="Einfügen-Menü", supertip="Einfüge-Optionen für aktualisierbare Folien-Thumbnails", children=[
                    bkt.ribbon.Button(
                        id = 'slide_paste_png',
                        label="Folien-Thumbnail als PNG einfügen",
                        show_label=True,
                        #image_mso='PasteAsPicture',
                        supertip="Kopierte Folie als aktualisierbares Thumbnail im PNG-Format mit Referenz auf Originalfolie einfügen.",
                        on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "slide_paste_png", application=True),
                        # get_enabled = bkt.Callback(Thumbnailer.enabled_paste),
                    ),
                    bkt.ribbon.Button(
                        id = 'slide_paste_btm',
                        label="Folien-Thumbnail als Bitmap einfügen",
                        show_label=True,
                        image_mso='PasteAsPicture',
                        supertip="Kopierte Folie als aktualisierbares Thumbnail im Bitmap-Format mit Referenz auf Originalfolie einfügen.",
                        on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "slide_paste_btm", application=True),
                        # get_enabled = bkt.Callback(Thumbnailer.enabled_paste),
                    ),
                    bkt.ribbon.Button(
                        id = 'slide_paste__emf',
                        label="Folien-Thumbnail als Vektor (EMF) einfügen",
                        show_label=True,
                        #image_mso='PasteAsPicture',
                        supertip="Kopierte Folie als aktualisierbares Thumbnail im Vektor-Format (Enhanced Metafile) mit Referenz auf Originalfolie einfügen.",
                        on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "slide_paste_emf", application=True),
                        # get_enabled = bkt.Callback(Thumbnailer.enabled_paste),
                    ),
                ])
            ]
        ),
        bkt.ribbon.SplitButton(
            children=[
                bkt.ribbon.Button(
                    id = 'slide_refresh',
                    label="Alle Thumbnails aktualisieren",
                    show_label=True,
                    image_mso='PictureChange',
                    supertip="Alle Folien-Thumbnails auf den ausgewählten Folien aktualisieren. Das Thumbnail muss vorher mit dieser Funktion eingefügt worden sein. Stammt die Folie aus einer anderen Datei, wird diese automatisch kurzzeitig geöffnet.",
                    on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "slide_refresh", application=True, slides=True),
                ),
                bkt.ribbon.Menu(label="Aktualisieren-Menü", supertip="Aktualisierung der Folien-Thumbnails auf dieser Folie oder in der ganzen Präsentation", item_size="large", children=[
                    bkt.ribbon.Button(
                        id = 'slide_refresh2',
                        label="Thumbnails auf Folie/n aktualisieren",
                        description="Alle Thumbnails auf aktueller bzw. ausgewählten Folie/n aktualisieren",
                        # show_label=True,
                        image_mso='PictureChange',
                        supertip="Alle Folien-Thumbnails auf den ausgewählten Folien aktualisieren. Das Thumbnail muss vorher mit dieser Funktion eingefügt worden sein. Stammt die Folie aus einer anderen Datei, wird diese automatisch kurzzeitig geöffnet.",
                        on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "slide_refresh", application=True, slides=True),
                    ),
                    bkt.ribbon.MenuSeparator(),
                    bkt.ribbon.Button(
                        id = 'presentation_refresh',
                        label="Thumbnails in Präsentation aktualisieren",
                        description="Alle Thumbnails in der gesamten Präsentation aktualisieren",
                        # show_label=True,
                        #image_mso='PictureChange',
                        supertip="Alle Folien-Thumbnails in der Präsentation aktualisieren. Das Thumbnail muss vorher mit dieser Funktion eingefügt worden sein. Stammt die Folie aus einer anderen Datei, wird diese automatisch kurzzeitig geöffnet.",
                        on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "presentation_refresh", application=True, presentation=True),
                    ),
                    bkt.ribbon.Button(
                        id = 'presentation_unset',
                        label="Thumbnails in Präsentation umwandeln",
                        description="Folien-Referenz aller Thumbnails in der Präsentation löschen und Thumbnails in Bilder konvertieren",
                        # show_label=True,
                        #image_mso='PictureChange',
                        supertip="Alle Folien-Thumbnails in der Präsentation in normale Bilder konvertieren, die sich nicht mehr aktualisieren lassen.",
                        on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "presentation_unset", presentation=True),
                    ),
                ])
            ]
        ),
    ]
)


bkt.powerpoint.add_tab(bkt.ribbon.Tab(
    id="bkt_powerpoint_toolbox_extensions",
    #id_q="nsBKT:powerpoint_toolbox_extensions",
    #insert_after_q="nsBKT:powerpoint_toolbox_advanced",
    insert_before_mso="TabHome",
    label='Toolbox 3/3',
    # get_visible defaults to False during async-startup
    get_visible=bkt.Callback(lambda:True),
    children = [
        thumbnail_gruppe,
    ]
), extend=True)


bkt.powerpoint.add_context_menu(
    bkt.ribbon.ContextMenu(id_mso='ContextMenuPicture', children=[
        bkt.ribbon.Button(
            id='context-thumbnail-refresh',
            label="Thumbnail aktualisieren",
            supertip="Ausgewähltes Folien-Thumbnail aktualisieren",
            insertBeforeMso='Cut',
            image_mso='PictureChange',
            on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "shape_refresh", shape=True, application=True),
            get_visible=bkt.Callback(ThumbnailerUi.is_thumbnail, shape=True),
        ),
        bkt.ribbon.DynamicMenu(
            id='context-thumbnail-settings',
            label="Thumbnail-Einstellungen",
            supertip="Einstellungen des gewählten Folien-Thumbnails ändern",
            image_mso='PictureSharpenSoftenGallery',
            insertBeforeMso='Cut',
            get_visible=bkt.Callback(ThumbnailerUi.is_thumbnail, shape=True),
            get_content=bkt.CallbackLazy(MODEL_MODULE, "context_settings")
        ),
        bkt.ribbon.DynamicMenu(
            id='context-thumbnail-reference',
            label="Folien-Referenz",
            supertip="Referenz des gewählten Folien-Thumbnails öffnen oder ändern",
            image_mso='PictureInsertFromFile',
            insertBeforeMso='Cut',
            get_visible=bkt.Callback(ThumbnailerUi.is_thumbnail, shape=True),
            get_content=bkt.CallbackLazy(MODEL_MODULE, "context_reference")
        ),
        bkt.ribbon.MenuSeparator(insertBeforeMso='Cut')
    ])
)


bkt.powerpoint.add_context_menu(
    bkt.ribbon.ContextMenu(id_mso='ContextMenuThumbnail', children=[
        bkt.ribbon.Button(
            id='context-thumbnail-slide-copy',
            label="Als Folien-Thumbnail kopieren",
            supertip="Gewählte Folie als aktualisierbares Thumbnail kopieren",
            insertAfterMso='Copy',
            image_mso='Copy',
            on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "slides_copy", presentation=True, slides=True),
            #get_visible=bkt.Callback(Thumbnailer.is_thumbnail, shape=True),
        ),
    ])
)

bkt.powerpoint.add_context_menu(
    bkt.ribbon.ContextMenu(id_mso='ContextMenuFrame', children=[
        bkt.ribbon.SplitButton(
            insertAfterMso='PasteGalleryMini',
            get_enabled=bkt.Callback(ThumbnailerUi.enabled_paste),
            children=[
                bkt.ribbon.Button(
                    id='context-thumbnail-slide-paste',
                    label="Als Folien-Thumbnail einfügen",
                    supertip="Als aktualisierbares Folien-Thumbnail im PNG-Format einfügen",
                    image_mso='PasteAsPicture',
                    on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "slide_paste", application=True),
                    #get_visible=bkt.Callback(Thumbnailer.is_thumbnail, shape=True),
                    get_enabled=bkt.Callback(ThumbnailerUi.enabled_paste),
                ),
                bkt.ribbon.Menu(label="Als Folien-Thumbnail einfügen Menü", supertip="Format zum Einfügen des Thumbnails auswählen", children=[
                    bkt.ribbon.Button(
                        label="Als PNG einfügen (Standard)",
                        on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "slide_paste_png", application=True),
                    ),
                    bkt.ribbon.Button(
                        label="Als Bitmap einfügen",
                        on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "slide_paste_btm", application=True),
                    ),
                    bkt.ribbon.Button(
                        label="Als Vektor (EMF) einfügen",
                        on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "slide_paste_emf", application=True),
                    ),
                ])
            ]
        ),
    ])
)


# register dialog
bkt.powerpoint.context_dialogs.register_dialog(
    bkt.contextdialogs.ContextDialog(
        id=BKT_THUMBNAIL,
        module="ppt_thumbnails.thumbnails_popup"
    )
)