# -*- coding: utf-8 -*-
'''
Created on 2018-05-29
@author: Florian Stallmann
'''

import bkt

MODEL_MODULE = __package__ + ".model"
MODEL_CONTAINER = "CustomQuickEdit"


class FormatLibGallery(bkt.ribbon.Gallery):
    '''
    This is the gallery element to show custom format styles.
    '''
    
    def __init__(self, **kwargs):
        parent_id = kwargs.get('id') or ""
        my_kwargs = dict(
            label = 'Styles anzeigen',
            columns = 6,
            image_mso = 'ShapeQuickStylesHome',
            show_item_label=False,
            screentip="Custom-Styles Gallerie",
            supertip="Zeigt Übersicht über alle Custom-Styles im aktuellen Katalog.",
            item_height=64, item_width=64,
            children=[
                bkt.ribbon.Button(id=parent_id + "_pickup", label="Neuen Style aufnehmen", supertip="Nimmt Format vom gewählten Shape neu in die Gallerie auf.", image_mso="PickUpStyle", on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "show_pickup_window", context=True, shape=True), get_enabled = bkt.apps.ppt_shapes_exactly1_selected,),
                bkt.ribbon.Button(id=parent_id + "_help1", label="[STRG]+Klick für Bearbeiten und Löschen", supertip="Bei Klick auf ein Custom-Style mit gedrückter STRG-Taste öffnet sich ein Fenster zur Bearbeitung und Löschung dieses Styles.", enabled = False),
                bkt.ribbon.Button(id=parent_id + "_help2", label="[SHIFT]+Klick für Anlage neues Shape", supertip="Bei Klick auf ein Custom-Style mit gedrückter SHIFT-Taste wird immer ein neues Shapes in gewähltem Style angelegt.", enabled = False),
                bkt.ribbon.Button(id=parent_id + "_help3", label="Einschränkungen durch PowerPoint-Bugs", supertip="Liste von funktionalen Einschränkungen durch interne PowerPoint-Bugs anzeigen", image_mso="Risks", on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "show_caveats")),
            ]
        )
        my_kwargs.update(kwargs)

        super(FormatLibGallery, self).__init__(**my_kwargs)

    def on_action_indexed(self, selected_item, index, context):
        from .model import CustomQuickEdit
        CustomQuickEdit.apply_custom_style(index, context)
    
    def get_item_count(self):
        from .model import CustomFormatCatalog
        return CustomFormatCatalog.get_count()
        
    def get_item_label(self, index):
        from .model import CustomFormatCatalog
        return CustomFormatCatalog.get_label(index)
    
    def get_item_screentip(self, index):
        from .model import CustomFormatCatalog
        return CustomFormatCatalog.get_screentip(index)
        
    def get_item_supertip(self, index):
        from .model import CustomFormatCatalog
        return CustomFormatCatalog.get_supertip(index)
    
    def get_item_image(self, index):
        from .model import CustomFormatCatalog
        return CustomFormatCatalog.get_image(index)


customformats_group = bkt.ribbon.Group(
    id="bkt_customformats_group",
    label='Styles',
    supertip="Ermöglicht die Speicherung von eigenen Formatierungen/Styles in Katalogen zur späteren Wiederverwendung. Das Feature `ppt_customformats` muss installiert sein.",
    image_mso='SmartArtChangeColorsGallery',
    children = [
        FormatLibGallery(id="customformats_gallery", size="large"),
        bkt.ribbon.DynamicMenu(
            id="quickedit_config_menu",
            label="Styles konfiguration",
            supertip="Style-Katalog laden oder neuen Katalog anlegen.",
            image_mso="ShapeReports",
            show_label=False,
            # size="large",
            get_content = bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "get_styles"),
        ),
        bkt.ribbon.Button(
            id="quickedit_temp_apply",
            label='Format anwenden',
            image_mso='PasteApplyStyle',
            supertip="Ausgewählte Formate aus Zwischenspeicher auf selektierte Shapes anwenden.\n\nMit STRG kann die Auswahl der Formate erneut bearbeitet werden.",
            show_label=False,
            on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "temp_apply", shapes=True, context=True),
            get_enabled = bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "temp_enabled", selection=True),
        ),
        bkt.ribbon.Button(
            id="quickedit_temp_pickup",
            label='Format aufnehmen',
            image_mso='PickUpStyle',
            supertip="Format aus selektiertem Shape in Zwischenspeicher aufnehmen.",
            show_label=False,
            on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "temp_pickup", shape=True),
            get_enabled = bkt.apps.ppt_shapes_exactly1_selected,
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
        customformats_group,
    ]
), extend=True)


bkt.powerpoint.add_lazy_replacement("ShapeQuickStylesHome", FormatLibGallery(id="customformats_gallery-rep", show_label=False), )
