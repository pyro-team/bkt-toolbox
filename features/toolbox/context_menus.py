# -*- coding: utf-8 -*-
'''
Created on 06.02.2018

@author: rdebeerst
'''

from __future__ import absolute_import

import bkt

#FIXME: would be nice to have less dependencies and more lazy loading of modules on callback
from . import text
from . import arrange
from . import harvey
from . import shapes as mod_shapes
from . import shape_selection
from . import info
from . import agenda
from . import linkshapes
from . import processshapes
from . import language
from . import slides


# =================
# = CONTEXT MENUS =
# =================


### Context menu for multiple shapes or grouped shape

bkt.powerpoint.add_context_menu(
    bkt.ribbon.ContextMenu(id_mso='ContextMenuObjectsGroup', children=[
        ### Any shapes
        bkt.ribbon.Button(id='context-format-sync', label="Format angleichen", insertBeforeMso='Cut', image_mso="FormatPainter",
            supertip="Alle Shapes so formatieren wie das Shape, welches beim Öffnen des Kontextmenüs unter dem Cursor ist",
            on_action=bkt.Callback(shape_selection.FormatPainter.cm_sync_shapes, shapes=True, context=True),
        ),
        ### Chevron with header
        bkt.ribbon.Button(id='context-arrange-header-group', label="Überschrift anordnen", insertBeforeMso='Cut', image="headered_pentagon",
            supertip="Kopfzeile (Überschrift) wieder richtig auf dem Prozessschritt-Shape positionieren",
            on_action=bkt.Callback(processshapes.Pentagon.update_pentagon_group, shape=True),
            get_visible=bkt.Callback(processshapes.Pentagon.is_headered_group, shape=True)
        ),
        ### Updatable process chevrons
        bkt.ribbon.Button(id='context-convert-process', label="In Prozess konvertieren", insertBeforeMso='Cut', image="process_chevrons",
            supertip="Ausgewählte Prozess-Shapes in eine interaktive Prozess-Gruppe umwandeln, um einfach Prozesschritte hinzuzufügen",
            on_action=bkt.Callback(processshapes.ProcessChevrons.convert_to_process_chevrons, shape=True),
            get_visible=bkt.Callback(processshapes.ProcessChevrons.is_convertible, shape=True)
        ),
        bkt.ribbon.MenuSeparator(insertBeforeMso='Cut'),
        ### Harvey
        harvey.harvey_size_gallery(
            insert_before_mso='Cut',
            id='ctx_harvey_ball_size_gallery',
            get_visible=bkt.Callback(harvey.harvey_balls.change_harvey_enabled, shapes=True)
        ),
        harvey.harvey_color_gallery(
            insert_before_mso='Cut',
            id='ctx_harvey_ball_color_gallery',
            get_visible=bkt.Callback(harvey.harvey_balls.change_harvey_enabled, shapes=True)
        ),
        bkt.ribbon.MenuSeparator(insertBeforeMso='Cut'),
        ### Connector functions (basically "copy" standard functions to multi-selection of connectors)
        bkt.ribbon.Menu(
            id='context-connectors-type',
            label="Verbindungstypen",
            supertip="Verbindungstyp für alle ausgewählten Verbinder ändern",
            image_mso='ShapeConnectorStyleMenu',
            insertBeforeMso='ObjectsGroupMenu',
            get_visible=bkt.Callback(mod_shapes.CtxVerbinder.ctx_connectors_visible, shapes=True),
            children=[
                bkt.ribbon.Button(
                    id='context-connectors-type-straight',
                    label="Gerader Verbinder",
                    image_mso='ShapeConnectorStyleStraight',
                    on_action=bkt.Callback(lambda shapes: mod_shapes.CtxVerbinder.set_connector_type(shapes, 1), shapes=True),
                ),
                bkt.ribbon.Button(
                    id='context-connectors-type-elbow',
                    label="Gewinkelte Verbindung",
                    image_mso='ShapeConnectorStyleElbow',
                    on_action=bkt.Callback(lambda shapes: mod_shapes.CtxVerbinder.set_connector_type(shapes, 2), shapes=True),
                ),
                bkt.ribbon.Button(
                    id='context-connectors-type-curved',
                    label="Gekrümmte Verbindung",
                    image_mso='ShapeConnectorStyleCurved',
                    on_action=bkt.Callback(lambda shapes: mod_shapes.CtxVerbinder.set_connector_type(shapes, 3), shapes=True),
                ),
            ]
        ),
        bkt.ribbon.Button(
            id='context-connectors-reroute',
            label="Verbindungen neu erstellen",
            supertip="Alle ausgewählten Verbinder neu erstellen",
            insertBeforeMso='ObjectsGroupMenu',
            image_mso='ShapeRerouteConnectors',
            on_action=bkt.Callback(mod_shapes.CtxVerbinder.reroute_connectors, shapes=True),
            get_visible=bkt.Callback(mod_shapes.CtxVerbinder.ctx_connectors_visible, shapes=True),
            get_enabled=bkt.Callback(mod_shapes.CtxVerbinder.ctx_connectors_reroute_enabled, shapes=True),
        ),
        bkt.ribbon.Button(
            id='context-connectors-invert',
            label="Pfeilrichtung umdrehen",
            supertip="Pfeilrichtung des Verbinders umkehren",
            insertBeforeMso='ObjectsGroupMenu',
            image_mso='ArrowStyleGallery',
            on_action=bkt.Callback(mod_shapes.CtxVerbinder.invert_direction, shapes=True),
            get_visible=bkt.Callback(mod_shapes.CtxVerbinder.ctx_connectors_visible, shapes=True),
        ),
        bkt.ribbon.MenuSeparator(insertBeforeMso='ObjectsGroupMenu'),
        ### Language setting
        bkt.ribbon.DynamicMenu(
            id="context-lang-change-shapes",
            label="Sprache ändern",
            supertip="Sprache der Rechtschreibkorrektur für ausgewählte Shapes anpassen",
            image_mso="GroupLanguage",
            insertAfterMso='ObjectFormatDialog',
            get_content=bkt.Callback(language.LangSetter.get_dynamicmenu_content),
        ),
        ### Text operations
        bkt.ribbon.Button(
            id = 'text_in_shape-context',
            label = u"Text in Shape kombinieren",
            supertip="Kopiere den Text eines Text-Shapes in das zweite markierte Shape und löscht das Text-Shape.",
            image_mso = "TextBoxInsert",
            on_action=bkt.Callback(text.TextOnShape.textIntoShape, shapes=True, shapes_min=2),
            get_visible = bkt.Callback(text.TextOnShape.is_mergable, shapes=True),
            insertBeforeMso='ObjectsGroupMenu',
        ),
        bkt.ribbon.Button(
            id = 'compose_text-context',
            label = u"Shape-Text zusammenführen",
            supertip="Führe die markierten Shapes in ein Shape zusammen. Der Text aller Shapes wird übernommen und aneinandergehängt.",
            image_mso = "TracePrecedents",
            on_action=bkt.Callback(text.SplitTextShapes.joinShapesWithText, shapes=True),
            get_visible = bkt.Callback(text.SplitTextShapes.is_joinable, shapes=True),
            insertBeforeMso='ObjectsGroupMenu',
        ),
        bkt.ribbon.Button(
            id = 'text_truncate-context',
            label="Shape-Texte löschen",
            supertip="Führe die markierten Shapes in ein Shape zusammen. Der Text aller Shapes wird übernommen und aneinandergehängt.",
            image_mso='ReviewDeleteMarkup',
            on_action=bkt.Callback(text.TextPlaceholder.text_truncate, shapes=True),
            get_visible = bkt.Callback(text.SplitTextShapes.is_joinable, shapes=True), #reuse callback from SplitTextShapes
            insertBeforeMso='ObjectsGroupMenu',
        ),
        bkt.ribbon.Button(
            id = 'text_replace-context',
            label="Shape-Texte ersetzen…",
            supertip="Text aller gewählten Shapes mit im Dialogfeld eingegebenen Text ersetzen.",
            image_mso='ReplaceDialog',
            on_action=bkt.Callback(text.TextPlaceholder.text_replace, shapes=True),
            get_visible = bkt.Callback(text.SplitTextShapes.is_joinable, shapes=True), #reuse callback from SplitTextShapes
            insertBeforeMso='ObjectsGroupMenu',
        ),
        bkt.ribbon.MenuSeparator(insertBeforeMso='ObjectsGroupMenu'),
        # Grouping functions
        bkt.ribbon.Button(
            id='add_into_group-context',
            label="In Gruppe einfügen",
            supertip="Sofern das zuerst oder zuletzt markierte Shape eine Gruppe ist, werden alle anderen Shapes in diese Gruppe eingefügt. Anderenfalls werden alle Shapes gruppiert.",
            image_mso="ObjectsRegroup",
            on_action=bkt.Callback(arrange.GroupsMore.add_into_group, shapes=True),
            get_visible = bkt.Callback(arrange.GroupsMore.visible_add_into_group, shapes=True),
            insertAfterMso='ObjectsGroupMenu',
        ),
        ### Clipboard operations
        bkt.ribbon.Button(
            id='paste_and_replace-context-shapes',
            label="Mit Zwischenablage ersetzen",
            supertip="Markiertes Shape mit dem Inhalt der Zwischenablage ersetzen und dabei Größe und Position erhalten.",
            insertAfterMso='PasteGalleryMini',
            image_mso='PasteSingleCellExcelTableDestinationFormatting',
            on_action=bkt.Callback(shape_selection.SlidesMore.paste_and_replace, slide=True, shape=True),
            get_enabled=bkt.Callback(lambda context: context.app.commandbars.GetEnabledMso("Paste"), context=True),
            get_visible=bkt.apps.ppt_shapes_exactly1_selected,
        ),
        ### Linked shapes
    ] + linkshapes.linked_shapes_context_menu('context-shapes'))
)


### Context menu for freeform shape type

bkt.powerpoint.add_context_menu(
    bkt.ribbon.ContextMenu(id_mso='ContextMenuShapeFreeform', children=[
        ### Chevron with header
        bkt.ribbon.Button(id='context-arrange-header', label="Überschrift anordnen", insertBeforeMso='Cut', image="headered_pentagon",
            supertip="Kopfzeile (Überschrift) wieder richtig auf dem Prozessschritt-Shape positionieren",
            on_action=bkt.Callback(processshapes.Pentagon.search_body_and_update_header, shape=True, context=True),
            get_visible=bkt.Callback(processshapes.Pentagon.is_header_shape, shape=True)
        ),
        ### Shape connectors
        bkt.ribbon.Button(
            id = 'connector_update-context',
            label = "Verbindungsfläche neu verbinden",
            supertip="Aktualisiere die Verbindungsfläche nachdem sich die verbundenen Shapes geändert haben.",
            image = "ConnectorUpdate",
            on_action=bkt.Callback(mod_shapes.ShapeConnectors.update_connector_shape, context=True, shape=True),
            get_visible = bkt.Callback(mod_shapes.ShapeConnectors.is_connector, shape=True),
            insertBeforeMso='Cut',
        ),
        bkt.ribbon.MenuSeparator(insertBeforeMso='Cut'),
        ### Clipboard operations
        bkt.ribbon.Button(
            id='paste_and_replace-context-freeform',
            label="Mit Zwischenablage ersetzen",
            supertip="Markiertes Shape mit dem Inhalt der Zwischenablage ersetzen und dabei Größe und Position erhalten.",
            insertAfterMso='PasteGalleryMini',
            image_mso='PasteSingleCellExcelTableDestinationFormatting',
            on_action=bkt.Callback(shape_selection.SlidesMore.paste_and_replace, slide=True, shape=True),
            get_enabled=bkt.Callback(lambda context: context.app.commandbars.GetEnabledMso("Paste"), context=True),
        ),
        ### Linked shapes
    ] + linkshapes.linked_shapes_context_menu('context-freeform'))
)


### Context menu for single "standard" shape

bkt.powerpoint.add_context_menu(
    bkt.ribbon.ContextMenu(id_mso='ContextMenuShape', children=[
        ### Language setting
        bkt.ribbon.DynamicMenu(
            id="context-lang-change-shp",
            label="Sprache ändern",
            supertip="Sprache der Rechtschreibkorrektur für ausgewähltes Shape anpassen",
            image_mso="GroupLanguage",
            insertAfterMso='ObjectFormatDialog',
            get_content=bkt.Callback(language.LangSetter.get_dynamicmenu_content),
        ),
        ### Text operations
        bkt.ribbon.Button(
            id = 'decompose_text-context',
            label = "Shape-Text zerlegen",
            supertip="Zerlege die markierten Shapes anhand der Text-Absätze in mehrere Shapes. Pro Absatz wird ein Shape mit dem entsprechenden Text angelegt.",
            image_mso = "TraceDependents",
            on_action=bkt.Callback(text.SplitTextShapes.splitShapesByParagraphs, shapes=True, context=True),
            get_visible = bkt.Callback(text.SplitTextShapes.is_splitable, shape=True),
            insertAfterMso='ObjectEditPoints',
        ),
        bkt.ribbon.Button(
            id = 'text_on_shape-context',
            label = "Text auf Shape zerlegen",
            supertip="Überführe jeweils den Textinhalt der markierten Shapes in ein separates Text-Shape.",
            image_mso = "TableCellCustomMarginsDialog",
            on_action=bkt.Callback(text.TextOnShape.textOutOfShape, shapes=True, context=True),
            get_visible = bkt.Callback(text.TextOnShape.is_outable, shape=True),
            insertAfterMso='ObjectEditPoints',
        ),
        ### Clipboard operations
        bkt.ribbon.Button(
            id='paste_and_replace-context-shp',
            label="Mit Zwischenablage ersetzen",
            supertip="Markiertes Shape mit dem Inhalt der Zwischenablage ersetzen und dabei Größe und Position erhalten.",
            insertAfterMso='PasteGalleryMini',
            image_mso='PasteSingleCellExcelTableDestinationFormatting',
            on_action=bkt.Callback(shape_selection.SlidesMore.paste_and_replace, slide=True, shape=True),
            get_enabled=bkt.Callback(lambda context: context.app.commandbars.GetEnabledMso("Paste"), context=True),
        ),
        ### Linked shapes
    ] + linkshapes.linked_shapes_context_menu('context-shp'))
)


### Conext menu for text selection

bkt.powerpoint.add_context_menu(
    bkt.ribbon.ContextMenu(id_mso='ContextMenuTextEdit', children=[
        ### Language setting
        bkt.ribbon.DynamicMenu(
            id="context-lang-change-txt",
            label="Sprache ändern",
            supertip="Sprache der Rechtschreibkorrektur ausgewählten Text anpassen",
            image_mso="GroupLanguage",
            insertAfterMso='ObjectFormatDialog',
            get_content=bkt.Callback(language.LangSetter.get_dynamicmenu_content),
        ),
    ])
)


### Context menu for spell correction

bkt.powerpoint.add_context_menu(
    bkt.ribbon.ContextMenu(id_mso='ContextMenuSpell', children=[
        ### Language setting
        bkt.ribbon.DynamicMenu(
            id="context-lang-change-spell",
            label="Sprache ändern",
            supertip="Sprache der Rechtschreibkorrektur ausgewählten Text anpassen",
            image_mso="GroupLanguage",
            insertAfterMso='ObjectFormatDialog',
            get_content=bkt.Callback(language.LangSetter.get_dynamicmenu_content),
        ),
    ])
)


### Context menu for picture

bkt.powerpoint.add_context_menu(
    bkt.ribbon.ContextMenu(id_mso='ContextMenuPicture', children=[
        ### Clipboard operations
        bkt.ribbon.Button(
            id='paste_and_replace-context-pic',
            label="Mit Zwischenablage ersetzen",
            supertip="Markiertes Shape mit dem Inhalt der Zwischenablage ersetzen und dabei Größe und Position erhalten.",
            insertAfterMso='PasteGalleryMini',
            image_mso='PasteSingleCellExcelTableDestinationFormatting',
            on_action=bkt.Callback(shape_selection.SlidesMore.paste_and_replace, slide=True, shape=True),
            get_enabled=bkt.Callback(lambda context: context.app.commandbars.GetEnabledMso("Paste"), context=True),
        ),
        ### Linked shapes
    ] + linkshapes.linked_shapes_context_menu('context-pic'))
)


### Context menu for connector

bkt.powerpoint.add_context_menu(
    bkt.ribbon.ContextMenu(id_mso='ContextMenuShapeConnector', children=[
        ### Connector functions
        bkt.ribbon.Button(
            id='context-connector-invert',
            label="Pfeilrichtung umdrehen",
            supertip="Pfeilrichtung des Verbinders umkehren",
            insertAfterMso='ShapeRerouteConnectors',
            image_mso='ArrowStyleGallery',
            on_action=bkt.Callback(mod_shapes.CtxVerbinder.invert_direction, shapes=True),
            #get_visible=bkt.Callback(mod_shapes.CtxVerbinder.ctx_connectors_visible, shapes=True),
        ),
    ])
)


### Context menu for empty slide area

bkt.powerpoint.add_context_menu(
    bkt.ribbon.ContextMenu(id_mso='ContextMenuFrame', children=[
        ### nothing so far
    ])
)


### Context menu for slide thumbnails

bkt.powerpoint.add_context_menu(
    bkt.ribbon.ContextMenu(id_mso='ContextMenuThumbnail', children=[
        ### Copy to slides
        bkt.ribbon.Button(
            id='context-paste-to-slides',
            label="Auf ausgewählte Folien einfügen",
            supertip="Zwischenablage auf allen ausgewählten Folien einfügen",
            insertAfterMso='PasteGalleryMini',
            image_mso='Paste',
            on_action=bkt.Callback(shape_selection.SlidesMore.paste_to_slides, slides=True),
            get_enabled=bkt.Callback(lambda context: context.app.commandbars.GetEnabledMso("Paste"), context=True),
        ),
        ### Language setting
        bkt.ribbon.DynamicMenu(
            id="context-lang-change-slides",
            label="Sprache ändern",
            supertip="Sprache der Rechtschreibkorrektur für ausgewählte Folien anpassen",
            image_mso="GroupLanguage",
            insertAfterMso='SlideBackgroundFormatDialog',
            get_content=bkt.Callback(language.LangSetter.get_dynamicmenu_content),
        ),
        ### Export (send, save) selected slides
        bkt.ribbon.Menu(
            id="context-export-slides",
            label="Ausgewählte Folien exportieren",
            supertip="Ausgewählte Folien als eigene Präsentation exportieren oder als E-Mail versenden",
            image_mso="SaveSelectionToTextBoxGallery",
            insertAfterMso='SlideBackgroundFormatDialog',
            children=[
                bkt.ribbon.Button(
                    id = 'context-export-save_slides',
                    label='Speichern',
                    image_mso='SaveSelectionToTextBoxGallery',
                    supertip="Speichert die ausgewählten Folien in einer neuen Präsentation.",
                    on_action=bkt.Callback(slides.SlideMenu.save_slides_dialog)
                ),
                bkt.ribbon.Button(
                    id = 'context-export-send_slides',
                    label='Senden',
                    image_mso='FileSendAsAttachment',
                    supertip="Sendet die ausgewählten Folien als E-Mail Anhang.",
                    on_action=bkt.Callback(slides.SlideMenu.send_slides_dialog)
                ),
            ]
        ),
        bkt.ribbon.MenuSeparator(insertAfterMso='SlideBackgroundFormatDialog'),
    ])
)


### Context menu for slide thumbnails in sort view

bkt.powerpoint.add_context_menu(
    bkt.ribbon.ContextMenu(id_mso='ContextMenuSlideSorter', children=[
        ### Export (send, save) selected slides
        bkt.ribbon.Menu(
            id="context-export2-slides",
            label="Folien exportieren",
            supertip="Ausgewählte Folien als eigene Präsentation exportieren oder als E-Mail versenden",
            image_mso="SaveSelectionToTextBoxGallery",
            insertAfterMso='SlideBackgroundFormatDialog',
            children=[
                bkt.ribbon.Button(
                    id = 'context-export2-save_slides',
                    label='Speichern',
                    image_mso='SaveSelectionToTextBoxGallery',
                    supertip="Speichert die ausgewählten Folien in einer neuen Präsentation.",
                    on_action=bkt.Callback(slides.SlideMenu.save_slides_dialog)
                ),
                bkt.ribbon.Button(
                    id = 'context-export2-send_slides',
                    label='Senden',
                    image_mso='FileSendAsAttachment',
                    supertip="Sendet die ausgewählten Folien als E-Mail Anhang.",
                    on_action=bkt.Callback(slides.SlideMenu.send_slides_dialog)
                ),
            ]
        ),
        bkt.ribbon.MenuSeparator(insertAfterMso='SlideBackgroundFormatDialog'),
    ])
)




# ==================
# = CONTEXTUAL TAB =
# ==================

# bkt.powerpoint.add_contextual_tab(
#     "TabSetDrawingTools",
#     harvey.harvey_ball_tab
# )
#use standard tab instead of contextual tab as contextual tab is not reliably shown (e.g. if PPT Format tab is hidden)
bkt.powerpoint.add_tab(harvey.harvey_ball_tab)

bkt.powerpoint.add_tab(agenda.agenda_tab)

bkt.powerpoint.add_tab(linkshapes.linkshapes_tab)

bkt.powerpoint.add_contextual_tab(
    "TabSetPictureTools",
    mod_shapes.picture_format_tab
)

bkt.powerpoint.add_contextual_tab(
    "TabSetDrawingTools",
    info.context_format_tab
)





# ==========
# = POPUPS =
# ==========


bkt.powerpoint.context_dialogs.register("BKT_DIALOG_AMPEL3", "toolbox.popups.traffic_light") #traffic light
bkt.powerpoint.context_dialogs.register("BKT_DIALOG_STATESHAPE", "toolbox.popups.stateshapes") #stateshapes, e.g. likert scale

bkt.powerpoint.context_dialogs.register(processshapes.ProcessChevrons.BKT_DIALOG_TAG, "toolbox.popups.processshapes") #process chevrons
bkt.powerpoint.context_dialogs.register(harvey.HarveyBalls.BKT_HARVEY_DIALOG_TAG, "toolbox.popups.harvey") #harvey balls
bkt.powerpoint.context_dialogs.register(linkshapes.BKT_LINK_UUID, "toolbox.popups.linkshapes") #linked shapesD
bkt.powerpoint.context_dialogs.register(agenda.TOOLBOX_AGENDA_POPUP, "toolbox.popups.agenda") #linked shapesD