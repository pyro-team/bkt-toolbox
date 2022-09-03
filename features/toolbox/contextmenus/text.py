# -*- coding: utf-8 -*-
'''
Created on 01.08.2022

@author: fstallmann
'''

from __future__ import absolute_import

import bkt

from .. import text

class ContextTextShapes(object):
    @staticmethod
    def get_buttons(shapes):
        if len(shapes) == 1:
            return [
                ### Text operations
                bkt.ribbon.MenuSeparator(title="Textoperationen"),
                bkt.ribbon.Button(
                    id = 'text_on_shape-context',
                    label = "Text auf Shape zerlegen",
                    supertip="Überführe jeweils den Textinhalt der markierten Shapes in ein separates Text-Shape.",
                    image_mso = "TableCellCustomMarginsDialog",
                    on_action=bkt.Callback(text.TextOnShape.textOutOfShape, shapes=True, slide=True),
                    get_enabled = bkt.Callback(text.TextOnShape.is_outable, shape=True),
                ),
                bkt.ribbon.Button(
                    id = 'decompose_text-context',
                    label = "Shape-Text zerlegen",
                    supertip="Zerlege die markierten Shapes anhand der Text-Absätze in mehrere Shapes. Pro Absatz wird ein Shape mit dem entsprechenden Text angelegt.",
                    image_mso = "TraceDependents",
                    on_action=bkt.Callback(text.SplitTextShapes.splitShapesByParagraphs, shapes=True, context=True),
                    get_enabled = bkt.Callback(text.SplitTextShapes.is_splitable, shape=True),
                ),
            ]
        else:
            return [
                ### Text operations
                bkt.ribbon.MenuSeparator(title="Textoperationen"),
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
                    id = 'text_on_shape-context',
                    label = "Text auf Shape zerlegen",
                    supertip="Überführe jeweils den Textinhalt der markierten Shapes in ein separates Text-Shape.",
                    image_mso = "TableCellCustomMarginsDialog",
                    on_action=bkt.Callback(text.TextOnShape.textOutOfShape, shapes=True, slide=True),
                    # get_enabled = bkt.Callback(text.TextOnShape.is_outable, shape=True),
                    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                ),
                bkt.ribbon.Button(
                    id = 'decompose_text-context',
                    label = "Shape-Text zerlegen",
                    supertip="Zerlege die markierten Shapes anhand der Text-Absätze in mehrere Shapes. Pro Absatz wird ein Shape mit dem entsprechenden Text angelegt.",
                    image_mso = "TraceDependents",
                    on_action=bkt.Callback(text.SplitTextShapes.splitShapesByParagraphs, shapes=True, context=True),
                    # get_enabled = bkt.Callback(text.SplitTextShapes.is_splitable, shape=True),
                    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
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
            ]