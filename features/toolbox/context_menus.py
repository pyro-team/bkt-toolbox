# -*- coding: utf-8 -*-
'''
Created on 06.02.2018

@author: rdebeerst
'''

import bkt

import arrange
import harvey
import shapes as mod_shapes
import info
import agenda
import linkshapes
import processshapes


# =================
# = CONTEXT MENUS =
# =================

bkt.powerpoint.add_context_menu(
    bkt.ribbon.ContextMenu(id_mso='ContextMenuObjectsGroup', children=[
        bkt.ribbon.Button(id='context-arrange-header-group', label="Überschrift anordnen", insertBeforeMso='Cut', image="headered_pentagon",
            on_action=bkt.Callback(lambda shape: processshapes.Pentagon.update_pentagon_group(shape), shapes=True, shape=True),
            get_visible=bkt.Callback(lambda shape: processshapes.Pentagon.is_headered_group(shape), shapes=True, shape=True)
        ),
        bkt.ribbon.MenuSeparator(insertBeforeMso='Cut')
    ])
)

bkt.powerpoint.add_context_menu(
    bkt.ribbon.ContextMenu(id_mso='ContextMenuShapeFreeform', children=[
        bkt.ribbon.Button(id='context-arrange-header', label="Überschrift anordnen", insertBeforeMso='Cut', image="headered_pentagon",
            on_action=bkt.Callback(lambda shape, context: processshapes.Pentagon.search_body_and_update_header(list(iter(context.app.activewindow.view.slide.shapes)), shape), shapes=True, shape=True, context=True),
            get_visible=bkt.Callback(lambda shape: processshapes.Pentagon.is_header_shape(shape), shapes=True, shape=True)
        ),
        bkt.ribbon.MenuSeparator(insertBeforeMso='Cut')
    ])
)

bkt.powerpoint.add_context_menu(
    bkt.ribbon.ContextMenu(id_mso='ContextMenuShape', children=
        linkshapes.linked_shapes_context_menu('context1')
    )
)
bkt.powerpoint.add_context_menu(
    bkt.ribbon.ContextMenu(id_mso='ContextMenuShapeFreeform', children=
        linkshapes.linked_shapes_context_menu('context2')
    )
)
bkt.powerpoint.add_context_menu(
    bkt.ribbon.ContextMenu(id_mso='ContextMenuPicture', children=
        linkshapes.linked_shapes_context_menu('context3')
    )
)
bkt.powerpoint.add_context_menu(
    bkt.ribbon.ContextMenu(id_mso='ContextMenuObjectsGroup', children=
        linkshapes.linked_shapes_context_menu('context4')
    )
)

bkt.powerpoint.add_context_menu(
    bkt.ribbon.ContextMenu(id_mso='ContextMenuObjectsGroup', children=[
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
        bkt.ribbon.MenuSeparator(insertBeforeMso='Cut')
    ])
)

bkt.powerpoint.add_context_menu(
    bkt.ribbon.ContextMenu(id_mso='ContextMenuObjectsGroup', children=[
        bkt.ribbon.Menu(
            id='context-connectors-type',
            label="Verbindungstypen",
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
            insertBeforeMso='ObjectsGroupMenu',
            image_mso='ShapeRerouteConnectors',
            on_action=bkt.Callback(mod_shapes.CtxVerbinder.reroute_connectors, shapes=True),
            get_visible=bkt.Callback(mod_shapes.CtxVerbinder.ctx_connectors_visible, shapes=True),
            get_enabled=bkt.Callback(mod_shapes.CtxVerbinder.ctx_connectors_reroute_enabled, shapes=True),
        ),
        bkt.ribbon.Button(
            id='context-connectors-invert',
            label="Pfeilrichtung umdrehen",
            insertBeforeMso='ObjectsGroupMenu',
            image_mso='ArrowStyleGallery',
            on_action=bkt.Callback(mod_shapes.CtxVerbinder.invert_direction, shapes=True),
            get_visible=bkt.Callback(mod_shapes.CtxVerbinder.ctx_connectors_visible, shapes=True),
        ),
        bkt.ribbon.MenuSeparator(insertBeforeMso='ObjectsGroupMenu')
    ])
)

bkt.powerpoint.add_context_menu(
    bkt.ribbon.ContextMenu(id_mso='ContextMenuShapeConnector', children=[
        bkt.ribbon.Button(
            id='context-connector-invert',
            label="Pfeilrichtung umdrehen",
            insertAfterMso='ShapeRerouteConnectors',
            image_mso='ArrowStyleGallery',
            on_action=bkt.Callback(mod_shapes.CtxVerbinder.invert_direction, shapes=True),
            #get_visible=bkt.Callback(mod_shapes.CtxVerbinder.ctx_connectors_visible, shapes=True),
        ),
    ])
)

bkt.powerpoint.add_context_menu(
    bkt.ribbon.ContextMenu(id_mso='ContextMenuThumbnail', children=[
        bkt.ribbon.Button(
            id='context-paste-to-slides',
            label="Auf ausgewählte Folien einfügen",
            insertAfterMso='PasteGalleryMini',
            image_mso='Paste',
            on_action=bkt.Callback(mod_shapes.ShapesMore.paste_to_slides, slides=True),
            get_enabled=bkt.Callback(lambda context: context.app.commandbars.GetEnabledMso("Paste"), context=True),
        ),
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


bkt.powerpoint.context_dialogs.register("BKT_PROCESS_CHEVRONS", "toolbox.processshapes") #process chevrons
bkt.powerpoint.context_dialogs.register("BKT_DIALOG_AMPEL3", "toolbox.popups.traffic_light") #traffic light
bkt.powerpoint.context_dialogs.register("BKT_DIALOG_STATESHAPE", "toolbox.stateshapes") #stateshapes, e.g. likert scale