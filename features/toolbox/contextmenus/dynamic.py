# -*- coding: utf-8 -*-
'''
Created on 29.04.2021

@author: fstallmann
'''

from __future__ import absolute_import

import bkt

from .. import arrange
from .. import shapes as mod_shapes
from .. import text

from ..models import processshapes

from .stateshapes import ContextStateShapes
from .linkshapes import ContextLinkedShapes


class ObjectsGroup(object):
    @staticmethod
    def get_children(shapes):
        '''
        TODO:
        if   State Shape
        else Process Shape
        else Harvey
        else Linked Shapes (Create or Align)
        else Text Functions+Add to group
        '''


        return bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None, 
                children= ContextStateShapes.get_buttons(shapes) + ContextLinkedShapes.get_buttons(shapes) + [
        bkt.ribbon.MenuSeparator(),
        ### Chevron with header
        bkt.ribbon.Button(id='context-arrange-header-group', label="Überschrift anordnen", image="headered_pentagon",
            supertip="Kopfzeile (Überschrift) wieder richtig auf dem Prozessschritt-Shape positionieren",
            on_action=bkt.Callback(processshapes.Pentagon.update_pentagon_group, shape=True),
            get_visible=bkt.Callback(processshapes.Pentagon.is_headered_group, shape=True)
        ),
        ### Updatable process chevrons
        bkt.ribbon.Button(id='context-convert-process', label="In Prozess konvertieren", image="process_chevrons",
            supertip="Ausgewählte Prozess-Shapes in eine interaktive Prozess-Gruppe umwandeln, um einfach Prozesschritte hinzuzufügen",
            on_action=bkt.Callback(processshapes.ProcessChevrons.convert_to_process_chevrons, shape=True),
            get_visible=bkt.Callback(processshapes.ProcessChevrons.is_convertible, shape=True)
        ),
        bkt.ribbon.MenuSeparator(),
        ### Text operations
        bkt.ribbon.Button(
            id = 'text_in_shape-context',
            label = u"Text in Shape kombinieren",
            supertip="Kopiere den Text eines Text-Shapes in das zweite markierte Shape und löscht das Text-Shape.",
            image_mso = "TextBoxInsert",
            on_action=bkt.Callback(text.TextOnShape.textIntoShape, shapes=True, shapes_min=2),
            # get_enabled = bkt.Callback(text.TextOnShape.is_mergable, shapes=True),
            get_enabled = bkt.apps.ppt_shapes_min2_selected,
        ),
        bkt.ribbon.Button(
            id = 'compose_text-context',
            label = u"Shape-Text zusammenführen",
            supertip="Führe die markierten Shapes in ein Shape zusammen. Der Text aller Shapes wird übernommen und aneinandergehängt.",
            image_mso = "TracePrecedents",
            on_action=bkt.Callback(text.SplitTextShapes.joinShapesWithText, shapes=True, shapes_min=2),
            # get_enabled = bkt.Callback(text.SplitTextShapes.is_joinable, shapes=True),
            get_enabled = bkt.apps.ppt_shapes_min2_selected,
        ),
        bkt.ribbon.Button(
            id = 'text_truncate-context',
            label="Shape-Texte löschen",
            supertip="Führe die markierten Shapes in ein Shape zusammen. Der Text aller Shapes wird übernommen und aneinandergehängt.",
            image_mso='ReviewDeleteMarkup',
            on_action=bkt.Callback(text.TextPlaceholder.text_truncate, shapes=True),
            # get_enabled = bkt.Callback(text.SplitTextShapes.is_joinable, shapes=True), #reuse callback from SplitTextShapes
            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
        ),
        bkt.ribbon.Button(
            id = 'text_replace-context',
            label="Shape-Texte ersetzen…",
            supertip="Text aller gewählten Shapes mit im Dialogfeld eingegebenen Text ersetzen.",
            image_mso='ReplaceDialog',
            on_action=bkt.Callback(text.TextPlaceholder.text_replace, shapes=True, presentation=True),
            # get_enabled = bkt.Callback(text.SplitTextShapes.is_joinable, shapes=True), #reuse callback from SplitTextShapes
            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
        ),
        bkt.ribbon.MenuSeparator(),
        # Grouping functions
        bkt.ribbon.Button(
            id='add_into_group-context',
            label="In Gruppe einfügen",
            supertip="Sofern das zuerst oder zuletzt markierte Shape eine Gruppe ist, werden alle anderen Shapes in diese Gruppe eingefügt. Anderenfalls werden alle Shapes gruppiert.",
            image_mso="ObjectsRegroup",
            on_action=bkt.Callback(arrange.GroupsMore.add_into_group, shapes=True),
            get_enabled = bkt.Callback(arrange.GroupsMore.visible_add_into_group, shapes=True),
        ),
        bkt.ribbon.Button(
            id='remove_from_group-context',
            label="Aus Gruppe lösen",
            supertip="Die markierten Shapes werden aus der aktuelle Gruppe herausgelöst, ohne die Gruppe dabei zu verändern.",
            image_mso="ObjectsUngroup",
            on_action=bkt.Callback(arrange.GroupsMore.remove_from_group, shapes=True),
            get_visible = bkt.Callback(arrange.GroupsMore.visible_remove_from_group, shapes=True),
        ),
                    ]
            )


class ShapeFreeform(object):
    @staticmethod
    def get_children():
        return bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None, 
                children=[
        ### Chevron with header
        bkt.ribbon.Button(id='context-arrange-header', label="Überschrift anordnen", image="headered_pentagon",
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
            get_enabled = bkt.Callback(mod_shapes.ShapeConnectors.is_connector, shape=True),
        ),
        bkt.ribbon.MenuSeparator(),
        ### Text operations
        bkt.ribbon.Button(
            id = 'decompose_text-context',
            label = "Shape-Text zerlegen",
            supertip="Zerlege die markierten Shapes anhand der Text-Absätze in mehrere Shapes. Pro Absatz wird ein Shape mit dem entsprechenden Text angelegt.",
            image_mso = "TraceDependents",
            on_action=bkt.Callback(text.SplitTextShapes.splitShapesByParagraphs, shapes=True, context=True),
            get_enabled = bkt.Callback(text.SplitTextShapes.is_splitable, shape=True),
        ),
        bkt.ribbon.Button(
            id = 'text_on_shape-context',
            label = "Text auf Shape zerlegen",
            supertip="Überführe jeweils den Textinhalt der markierten Shapes in ein separates Text-Shape.",
            image_mso = "TableCellCustomMarginsDialog",
            on_action=bkt.Callback(text.TextOnShape.textOutOfShape, shapes=True, slide=True),
            get_enabled = bkt.Callback(text.TextOnShape.is_outable, shape=True),
        ),
                    ]
            )




class Shape(object):
    @staticmethod
    def get_children(shape):
        shapes = [shape]
        return bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None, 
                children= ContextStateShapes.get_buttons(shapes) + ContextLinkedShapes.get_buttons(shapes) + [
        # Grouping functions
        bkt.ribbon.MenuSeparator(),
        bkt.ribbon.Button(
            id='remove_from_group-context',
            label="Aus Gruppe lösen",
            supertip="Die markierten Shapes werden aus der aktuelle Gruppe herausgelöst, ohne die Gruppe dabei zu verändern.",
            image_mso="ObjectsUngroup",
            on_action=bkt.Callback(arrange.GroupsMore.remove_from_group, shapes=True),
            get_visible = bkt.Callback(arrange.GroupsMore.visible_remove_from_group, shapes=True),
        ),
                ]
        )
