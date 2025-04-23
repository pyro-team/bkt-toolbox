# -*- coding: utf-8 -*-
'''
Created on 29.03.2017

@author: tweuffel
'''

import bkt

MODEL_MODULE = __package__ + ".notes_model"
MODEL_CONTAINER = "EditModeShapes"


notes_gruppe = bkt.ribbon.Group(
    id="bkt_notes_group",
    label='Folien-Notizen',
    supertip="Ermöglicht das Einfügen von Bearbeitungsnotizen auf Folien. Das Feature `ppt_notes` muss installiert sein.",
    image='noteAdd',
    children = [
        bkt.ribbon.Button(
            id = 'notes_add',
            label='Erstellen', screentip='Notiz hinzufügen',
            supertip="Fügt eine Bearbeitungsnotiz oben rechts auf der Folie ein inkl. Autor und Datum.",
            image='noteAdd',
            on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "addNote", slide=True, context=True)
        ),
        bkt.ribbon.Button(
            id = 'notes_toggle',
            label='An/Aus', screentip='Notizen auf Folie ein-/ausblenden',
            supertip="Alle Notizen der aktuellen Folie temporär ausblenden und wieder einblenden.",
            image='noteToggle',
            on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "toogleNotesOnSlides", slides=True)
        ),
        bkt.ribbon.Button(
            id = 'notes_remove',
            label='Löschen', screentip='Notizen auf Folie löschen',
            supertip="Alle Notizen der aktuellen Folie entfernen.",
            image='noteRemove',
            on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "removeNotesOnSlides", slides=True)
        ),
        bkt.ribbon.Button(
            id = 'notes_toggle_all',
            label='Alle an/aus', screentip='Alle Notizen ein-/ausblenden',
            supertip="Alle Notizen auf allen Folien temporär ausblenden und wieder einblenden.",
            image='noteToggleAll',
            on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "toggleNotesOnAllSlides", slide=True)
        ),
        bkt.ribbon.Button(
            id = 'notes_remove_all',
            label='Alle löschen', screentip='Alle Notizen löschen',
            supertip="Alle Notizen auf allen Folien entfernen.",
            image='noteRemoveAll',
            on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "removeNotesOnAllSlides", slide=True)
        ),
        bkt.ribbon.ColorGallery(
            id = 'notes_color',
            label='Farbe',
            screentip='Notizen-Farbe ändern',
            supertip="Hintergrundfarbe für neue Bearbeitungsnotizen ändern.",
            on_rgb_color_change = bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "set_color_rgb"),
            on_theme_color_change = bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "set_color_theme"),
            get_selected_color = bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "get_color"),
            children=[
                bkt.ribbon.Button(
                    id="notes_color_default",
                    label="Standardfarbe",
                    supertip="Hintergrundfarbe für Bearbeitungsnotizen auf Standard zurücksetzen.",
                    on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "set_color_default"),
                    image_mso="ColorTeal",
                )
            ]
            # get_enabled = bkt.apps.ppt_shapes_or_text_selected,
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
        notes_gruppe,
    ]
), extend=True)


