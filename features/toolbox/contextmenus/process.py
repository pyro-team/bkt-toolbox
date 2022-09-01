# -*- coding: utf-8 -*-
'''
Created on 01.08.2022

@author: fstallmann
'''

from __future__ import absolute_import

import bkt

from ..models import processshapes
from ..models import segmentedcircle

class ContextProcessShapes(object):
    @classmethod
    def get_buttons(cls, shapes):
        buttons = []
        ### Chevron with header
        if processshapes.Pentagon.are_headered_groups(shapes):
            buttons.append(
                bkt.ribbon.Button(id='context-arrange-header-group', label="Überschrift anordnen", image="headered_pentagon",
                    supertip="Kopfzeile (Überschrift) wieder richtig auf dem Prozessschritt-Shape positionieren",
                    on_action=bkt.Callback(processshapes.Pentagon.update_pentagon_groups, shapes=True),
                )
            )
        ### Updatable process chevrons
        if len(shapes) == 1 and processshapes.ProcessChevrons.is_convertible(shapes[0]):
            buttons.append(
                bkt.ribbon.Button(id='context-convert-process', label="In Prozess konvertieren", image="process_chevrons",
                    supertip="Ausgewählte Prozess-Shapes in eine interaktive Prozess-Gruppe umwandeln, um einfach Prozesschritte hinzuzufügen",
                    on_action=bkt.Callback(processshapes.ProcessChevrons.convert_to_process_chevrons, shape=True),
                )
            )
        if len(shapes) == 1 and processshapes.ProcessChevrons.is_process_chevrons(shapes[0]):
            buttons.append(
                bkt.ribbon.Button(id='context-edit-process', label="Prozess bearbeiten", image="process_chevrons",
                    supertip="Ausgewählte Prozess-Gruppe bearbeiten und anpassen",
                    on_action=bkt.Callback(cls.show_process_chevrons_dialog, context=True, slide=True),
                )
            )
        ### Updatable segmented circle
        if len(shapes) == 1 and segmentedcircle.SegmentedCircle.is_segmented_circle(shapes[0]):
            buttons.append(
                bkt.ribbon.Button(id='context-edit-circle', label="Kreissegmente bearbeiten", image="segmented_circle",
                    supertip="Ausgewählte Kreissegmente bearbeiten und anpassen",
                    on_action=bkt.Callback(cls.show_segmented_circle_dialog, context=True, slide=True),
                )
            )

        if buttons:
            buttons.insert(0, bkt.ribbon.MenuSeparator(title="Prozess und Kreissegmente"))
        return buttons

        # return [
        #     bkt.ribbon.MenuSeparator(title="Prozess und Kreissegmente"),
        #     ### Chevron with header
        #     bkt.ribbon.Button(id='context-arrange-header-group', label="Überschrift anordnen", image="headered_pentagon",
        #         supertip="Kopfzeile (Überschrift) wieder richtig auf dem Prozessschritt-Shape positionieren",
        #         on_action=bkt.Callback(processshapes.Pentagon.update_pentagon_groups, shapes=True),
        #         get_visible=bkt.Callback(processshapes.Pentagon.are_headered_groups, shapes=True)
        #     ),
        #     ### Updatable process chevrons
        #     bkt.ribbon.Button(id='context-convert-process', label="In Prozess konvertieren", image="process_chevrons",
        #         supertip="Ausgewählte Prozess-Shapes in eine interaktive Prozess-Gruppe umwandeln, um einfach Prozesschritte hinzuzufügen",
        #         on_action=bkt.Callback(processshapes.ProcessChevrons.convert_to_process_chevrons, shape=True),
        #         get_visible=bkt.Callback(processshapes.ProcessChevrons.is_convertible, shape=True)
        #     ),
        #     bkt.ribbon.Button(id='context-edit-process', label="Prozess bearbeiten", image="process_chevrons",
        #         supertip="Ausgewählte Prozess-Gruppe bearbeiten und anpassen",
        #         on_action=bkt.Callback(cls.show_process_chevrons_dialog, context=True, slide=True),
        #         get_visible=bkt.Callback(processshapes.ProcessChevrons.is_process_chevrons, shape=True)
        #     ),
        #     ### Updatable segmented circle
        #     bkt.ribbon.Button(id='context-edit-circle', label="Kreissegmente bearbeiten", image="segmented_circle",
        #         supertip="Ausgewählte Kreissegmente bearbeiten und anpassen",
        #         on_action=bkt.Callback(cls.show_segmented_circle_dialog, context=True, slide=True),
        #         get_visible=bkt.Callback(segmentedcircle.SegmentedCircle.is_segmented_circle, shape=True)
        #     ),
        # ]
    
    @staticmethod
    def show_process_chevrons_dialog(context, slide):
        from ..dialogs.shape_process import ProcessWindow
        ProcessWindow.create_and_show_dialog(context, slide)
    
    @staticmethod
    def show_segmented_circle_dialog(context, slide):
        from ..dialogs.circular_segments import SegmentedCircleWindow
        SegmentedCircleWindow.create_and_show_dialog(context, slide)