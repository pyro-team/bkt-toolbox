# -*- coding: utf-8 -*-
'''
Created on 29.04.2021

@author: fstallmann
'''



import bkt

from .. import stateshapes
from ..models.stateshapes import StateShape


class ContextStateShapes(object):
    # cb_visible = bkt.Callback(stateshapes.StateShape.are_state_shapes)

    @classmethod
    def get_buttons(cls, shapes):
        if not stateshapes.StateShapeUi.are_state_shapes(shapes):
            return []
        return [
            bkt.ribbon.MenuSeparator(title="Wechselshapes"),
            bkt.ribbon.Button(
                image_mso="PreviousResource",
                label='Vorheriges',
                supertip="Wechselt zum vorherigen Status (d.h. Shape in der Gruppe) des Wechsel-Shapes.",
                on_action=bkt.Callback(StateShape.previous_state),
                # get_visible=cls.cb_visible,
            ),
            bkt.ribbon.Button(
                image_mso="NextResource",
                label="Nächstes",
                supertip="Wechselt zum nächsten Status (d.h. Shape in der Gruppe) des Wechsel-Shapes.",
                on_action=bkt.Callback(StateShape.next_state),
                # get_visible=cls.cb_visible,
            ),
            bkt.ribbon.MenuSeparator(),
            stateshapes.stateshape_fill1_gallery(
                # get_visible=cls.cb_visible,
            ),
            stateshapes.stateshape_fill2_gallery(
                # get_visible=cls.cb_visible,
            ),
            stateshapes.stateshape_line_gallery(
                # get_visible=cls.cb_visible,
            ),
        ]