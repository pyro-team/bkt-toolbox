# -*- coding: utf-8 -*-
'''
Created on 29.04.2021

@author: fstallmann
'''

from __future__ import absolute_import

import bkt

from ..linkshapes import LinkedShapes


class ContextLinkedShapes(object):
    @staticmethod
    def get_buttons(shapes):
        return [
            bkt.ribbon.MenuSeparator(title="Verknüpfte Shapes"),
            bkt.ribbon.Button(
                label="Shape auf Folgefolien suchen und verknüpfen…",
                image_mso="ShapesDuplicate",
                on_action=bkt.Callback(LinkedShapes.find_similar_and_link, shape=True, context=True),
                get_visible=bkt.Callback(LinkedShapes.not_is_linked_shape, shape=True),
            ),
            bkt.ribbon.Button(
                label="Shape auf Folgefolien kopieren und verknüpfen…",
                image_mso="FindTag",
                on_action=bkt.Callback(LinkedShapes.copy_to_all, shape=True, context=True),
                get_visible=bkt.Callback(LinkedShapes.not_is_linked_shape, shape=True),
            ),
            bkt.ribbon.SplitButton(
                get_visible=bkt.Callback(LinkedShapes.is_linked_shape, shape=True),
                children=[
                    bkt.ribbon.Button(
                        label="Verknüpfte Shapes angleichen",
                        image_mso="HyperlinkCreate",
                        on_action=bkt.Callback(LinkedShapes.equalize_linked_shapes, shapes=True, context=True),
                    ),
                    bkt.ribbon.Menu(children=[
                        bkt.ribbon.Button(
                            label="Alles angleichen",
                            # image_mso="GroupUpdate",
                            image_mso='HyperlinkCreate',
                            on_action=bkt.Callback(LinkedShapes.equalize_linked_shapes, shapes=True, context=True),
                        ),
                        bkt.ribbon.MenuSeparator(),
                        bkt.ribbon.Button(
                            label="Position angleichen",
                            image_mso="ControlAlignToGrid",
                            on_action=bkt.Callback(LinkedShapes.align_linked_shapes, shapes=True, context=True),
                        ),
                        bkt.ribbon.Button(
                            label="Größe angleichen",
                            image_mso="SizeToControlHeightAndWidth",
                            on_action=bkt.Callback(LinkedShapes.size_linked_shapes, shapes=True, context=True),
                        ),
                        bkt.ribbon.Button(
                            label="Formatierung angleichen",
                            image_mso="FormatPainter",
                            on_action=bkt.Callback(LinkedShapes.format_linked_shapes, shapes=True, context=True),
                        ),
                        bkt.ribbon.Button(
                            label="Text angleichen",
                            image_mso="TextBoxInsert",
                            on_action=bkt.Callback(LinkedShapes.text_linked_shapes, shapes=True, context=True),
                        ),
                        bkt.ribbon.MenuSeparator(),
                        bkt.ribbon.Button(
                            label="In den Vordergrund",
                            image_mso="ObjectBringToFront",
                            on_action=bkt.Callback(LinkedShapes.linked_shapes_tofront, shapes=True, context=True),
                        ),
                        bkt.ribbon.Button(
                            label="In den Hintergrund",
                            image_mso="ObjectSendToBack",
                            on_action=bkt.Callback(LinkedShapes.linked_shapes_toback, shapes=True, context=True),
                        ),
                        bkt.ribbon.MenuSeparator(),
                        bkt.ribbon.Button(
                            label="Andere löschen",
                            image_mso="HyperlinkRemove",
                            on_action=bkt.Callback(LinkedShapes.delete_linked_shapes, shapes=True, context=True),
                        ),
                        bkt.ribbon.Button(
                            label="Andere mit diesem ersetzen",
                            image_mso="HyperlinkCreate",
                            on_action=bkt.Callback(LinkedShapes.replace_with_this, shapes=True, context=True),
                        ),
                    ])
                ]
            ),
        ]

    @staticmethod
    def get_children_create():
        return bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None, 
                children=[
            bkt.ribbon.Button(
                label="Ähnliche Shapes suchen…",
                on_action=bkt.Callback(LinkedShapes.find_similar_and_link, shape=True, context=True),
                get_visible=bkt.Callback(LinkedShapes.not_is_linked_shape, shape=True),
            ),
            bkt.ribbon.Button(
                label="Dieses Shape kopieren…",
                on_action=bkt.Callback(LinkedShapes.copy_to_all, shape=True, context=True),
                get_visible=bkt.Callback(LinkedShapes.not_is_linked_shape, shape=True),
            ),
                    ]
            )
            
    @staticmethod
    def get_children_align():
        return bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None, 
                children=[
            bkt.ribbon.Button(
                label="Alles angleichen",
                # image_mso="GroupUpdate",
                image_mso='HyperlinkCreate',
                on_action=bkt.Callback(LinkedShapes.equalize_linked_shapes, shapes=True, context=True),
            ),
            bkt.ribbon.MenuSeparator(),
            bkt.ribbon.Button(
                label="Position angleichen",
                image_mso="ControlAlignToGrid",
                on_action=bkt.Callback(LinkedShapes.align_linked_shapes, shapes=True, context=True),
            ),
            bkt.ribbon.Button(
                label="Größe angleichen",
                image_mso="SizeToControlHeightAndWidth",
                on_action=bkt.Callback(LinkedShapes.size_linked_shapes, shapes=True, context=True),
            ),
            bkt.ribbon.Button(
                label="Formatierung angleichen",
                image_mso="FormatPainter",
                on_action=bkt.Callback(LinkedShapes.format_linked_shapes, shapes=True, context=True),
            ),
            bkt.ribbon.Button(
                label="Text angleichen",
                image_mso="TextBoxInsert",
                on_action=bkt.Callback(LinkedShapes.text_linked_shapes, shapes=True, context=True),
            ),
            bkt.ribbon.MenuSeparator(),
            bkt.ribbon.Button(
                label="In den Vordergrund",
                image_mso="ObjectBringToFront",
                on_action=bkt.Callback(LinkedShapes.linked_shapes_tofront, shapes=True, context=True),
            ),
            bkt.ribbon.Button(
                label="In den Hintergrund",
                image_mso="ObjectSendToBack",
                on_action=bkt.Callback(LinkedShapes.linked_shapes_toback, shapes=True, context=True),
            ),
            bkt.ribbon.MenuSeparator(),
            bkt.ribbon.Button(
                label="Andere löschen",
                image_mso="HyperlinkRemove",
                on_action=bkt.Callback(LinkedShapes.delete_linked_shapes, shapes=True, context=True),
            ),
            bkt.ribbon.Button(
                label="Andere mit diesem ersetzen",
                image_mso="HyperlinkCreate",
                on_action=bkt.Callback(LinkedShapes.replace_with_this, shapes=True, context=True),
            ),
                    ]
            )