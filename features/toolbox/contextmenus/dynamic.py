# -*- coding: utf-8 -*-
'''
Created on 29.04.2021

@author: fstallmann
'''

from __future__ import absolute_import

import bkt

from .. import arrange
from .. import shapes as mod_shapes

from ..models import processshapes

from .stateshapes import ContextStateShapes
from .linkshapes import ContextLinkedShapes
from .harvey import ContextHarveyShapes
from .text import ContextTextShapes
from .process import ContextProcessShapes


class ObjectsGroup(object):
    @staticmethod
    def get_children(shapes):

        return bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None, 
                children=   ContextStateShapes.get_buttons(shapes) + 
                            ContextHarveyShapes.get_buttons(shapes) + 
                            ContextProcessShapes.get_buttons(shapes) + 
                            ContextLinkedShapes.get_buttons(shapes) + 
                            ContextTextShapes.get_buttons(shapes) + [
            # Grouping functions
            bkt.ribbon.MenuSeparator(title="Gruppierung"),
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
    def get_children(shape):
        shapes = [shape]
        return bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None, 
                children= [
        ### Shape connectors
        bkt.ribbon.Button(
            id = 'connector_update-context',
            label = "Verbindungsfläche neu verbinden",
            supertip="Aktualisiere die Verbindungsfläche nachdem sich die verbundenen Shapes geändert haben.",
            image = "ConnectorUpdate",
            on_action=bkt.Callback(mod_shapes.ShapeConnectors.update_connector_shape, context=True, shape=True),
            get_enabled = bkt.Callback(mod_shapes.ShapeConnectors.is_connector, shape=True),
        ),
        ] + ContextLinkedShapes.get_buttons(shapes) + 
            ContextProcessShapes.get_buttons(shapes) + 
            ContextTextShapes.get_buttons(shapes)
        )




class Shape(object):
    @staticmethod
    def get_children(shape):
        shapes = [shape]
        return bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None, 
                children=   ContextStateShapes.get_buttons(shapes) + 
                            ContextLinkedShapes.get_buttons(shapes) + 
                            ContextTextShapes.get_buttons(shapes) + [
            # Grouping functions
            bkt.ribbon.MenuSeparator(title="Gruppierung"),
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




class Picture(object):
    @staticmethod
    def get_children(shape):
        shapes = [shape]
        return bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None, 
                children= ContextLinkedShapes.get_buttons(shapes)
        )