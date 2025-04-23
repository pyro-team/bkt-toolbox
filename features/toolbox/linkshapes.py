# -*- coding: utf-8 -*-
'''
Created on 06.09.2018

@author: fstallmann
'''

import bkt
import bkt.library.powerpoint as pplib


BKT_LINK_UUID = "BKT_LINK_UUID"


class LinkedShapesUi(object):
    ### Enabled/Visible callbacks ###

    @staticmethod
    def is_linked_shape(shape):
        return pplib.TagHelper.has_tag(shape, BKT_LINK_UUID)

    @classmethod
    def not_is_linked_shape(cls, shape):
        return not cls.is_linked_shape(shape)

    @classmethod
    def are_linked_shapes(cls, shapes):
        return all(cls.is_linked_shape(shape) for shape in shapes)

    @classmethod
    def enabled_add_linked_shapes(cls):
        return cls.current_link_guid != None



linkshapes_tab = bkt.ribbon.Tab(
    id = "bkt_context_tab_linkshapes",
    label = "[BKT] Verknüpfte Shapes",
    get_visible=bkt.Callback(LinkedShapesUi.are_linked_shapes, shapes=True),
    children = [
        bkt.ribbon.Group(
            id="bkt_linkshapes_find_group",
            label = "Verknüpfte Shapes finden",
            get_visible=bkt.apps.ppt_shapes_exactly1_selected,
            children = [
                bkt.ribbon.Box(box_style="horizontal", children=[
                    bkt.ribbon.Button(
                        id = 'linked_shapes_first',
                        label="Erstes verknüpfte Shape finden",
                        show_label=False,
                        image_mso="MailMergeGoToFirstRecord",
                        screentip="Zum ersten verknüpften Shape gehen",
                        supertip="Sucht nach dem ersten verknüpften Shape.",
                        on_action=bkt.CallbackLazy("toolbox.models.linkshapes", "LinkedShapes", "goto_first_shape", shape=True, context=True),
                    ),
                    bkt.ribbon.Button(
                        id = 'linked_shapes_previous',
                        label="Vorheriges verknüpfte Shape finden",
                        show_label=False,
                        image_mso="MailMergeGoToPreviousRecord",
                        screentip="Zum vorherigen verknüpften Shape gehen",
                        supertip="Sucht nach dem vorherigen verknüpften Shape. Sollte auf den vorherigen Folien kein Shape mehr kommen, wird das letzte verknüpfte Shape der Präsentation gesucht.",
                        on_action=bkt.CallbackLazy("toolbox.models.linkshapes", "LinkedShapes", "goto_previous_shape", shape=True, context=True),
                    ),
                    bkt.ribbon.Label(
                        label="Gehe zu",
                    ),
                    bkt.ribbon.Button(
                        id = 'linked_shapes_next',
                        label="Nächstes verknüpfte Shape finden",
                        show_label=False,
                        image_mso="MailMergeGoToNextRecord",
                        screentip="Zum nächsten verknüpften Shape gehen",
                        supertip="Sucht nach dem nächste verknüpften Shape. Sollte auf den Folgefolien kein Shape mehr kommen, wird das erste verknüpfte Shape der Präsentation gesucht.",
                        on_action=bkt.CallbackLazy("toolbox.models.linkshapes", "LinkedShapes", "goto_next_shape", shape=True, context=True),
                    ),
                    bkt.ribbon.Button(
                        id = 'linked_shapes_last',
                        label="Letztes verknüpfte Shape finden",
                        show_label=False,
                        image_mso="MailMergeGotToLastRecord",
                        screentip="Zum letzten verknüpften Shape gehen",
                        supertip="Sucht nach dem letzten verknüpften Shape.",
                        on_action=bkt.CallbackLazy("toolbox.models.linkshapes", "LinkedShapes", "goto_last_shape", shape=True, context=True),
                    ),
                ]),
                bkt.ribbon.Button(
                    id = 'linked_shapes_count',
                    label="Shapes zählen",
                    image_mso="FormattingUnique",
                    screentip="Alle verknüpften Shapes zählen",
                    supertip="Zählt die Anzahl der verknüpften Shapes auf allen Folien.",
                    on_action=bkt.CallbackLazy("toolbox.models.linkshapes", "LinkedShapes", "count_link_shapes", shape=True, context=True),
                ),
                bkt.ribbon.Button(
                    id = 'linked_shapes_select',
                    label="Folien anzeigen",
                    image_mso="SlideTransitionApplyToAll",
                    screentip="Alle Foliennummern mit verknüpften Shapes anzeigen",
                    supertip="Zeigt alle Foliennummern die zugehörige verknüpfte Shapes enthalten.",
                    on_action=bkt.CallbackLazy("toolbox.models.linkshapes", "LinkedShapes", "select_link_shapes_slides", shape=True, context=True),
                ),
            ]
        ),
        bkt.ribbon.Group(
            id="bkt_linkshapes_align_group",
            label = "Verknüpfte Shapes angleichen",
            children = [
                bkt.ribbon.DynamicMenu(
                    id = 'linked_shapes_master',
                    label="Referenz wählen",
                    image_mso="CircularReferences",
                    size="large",
                    screentip="Referenzshape auswählen",
                    supertip="Auswählen, ob selektiertes, erstes oder letztes Shape als Referenz für alle Angleichungsfunktionen verwendet werden soll. Standard ist das aktuell ausgewählte Shape.",
                    get_content=bkt.CallbackLazy("toolbox.models.linkshapes", "reference_menu")
                ),
                bkt.ribbon.Button(
                    id = 'linked_shapes_all',
                    label="Alles angleichen",
                    image_mso="GroupUpdate",
                    size="large",
                    screentip="Alle Eigenschaften verknüpfter Shapes angleichen",
                    supertip="Alle Eigenschaften aller verknüpfter Shapes wie ausgewähltes Shape setzen.",
                    on_action=bkt.CallbackLazy("toolbox.models.linkshapes", "LinkedShapes", "equalize_linked_shapes", shapes=True, context=True),
                ),
                bkt.ribbon.Separator(),
                bkt.ribbon.Button(
                    id = 'linked_shapes_align',
                    label="Position angleichen",
                    image_mso="ControlAlignToGrid",
                    screentip="Position verknüpfter Shapes angleichen",
                    supertip="Position und Rotation aller verknüpfter Shapes auf Position wie ausgewähltes Shape setzen.",
                    on_action=bkt.CallbackLazy("toolbox.models.linkshapes", "LinkedShapes", "align_linked_shapes", shapes=True, context=True),
                ),
                bkt.ribbon.Button(
                    id = 'linked_shapes_size',
                    label="Größe angleichen",
                    image_mso="SizeToControlHeightAndWidth",
                    screentip="Größe verknüpfter Shapes angleichen",
                    supertip="Größe aller verknüpfter Shapes auf Größe wie ausgewähltes Shape setzen.",
                    on_action=bkt.CallbackLazy("toolbox.models.linkshapes", "LinkedShapes", "size_linked_shapes", shapes=True, context=True),
                ),
                bkt.ribbon.Button(
                    id = 'linked_shapes_format',
                    label="Formatierung angleichen",
                    image_mso="FormatPainter",
                    screentip="Formatierung verknüpfter Shapes angleichen",
                    supertip="Formatierung aller verknüpfter Shapes auf Größe wie ausgewähltes Shape setzen.",
                    on_action=bkt.CallbackLazy("toolbox.models.linkshapes", "LinkedShapes", "format_linked_shapes", shapes=True, context=True),
                ),
                bkt.ribbon.Button(
                    id = 'linked_shapes_text',
                    label="Text angleichen",
                    image_mso="TextBoxInsert",
                    screentip="Text verknüpfter Shapes angleichen",
                    supertip="Text aller verknüpfter Shapes auf Größe wie ausgewähltes Shape setzen.",
                    on_action=bkt.CallbackLazy("toolbox.models.linkshapes", "LinkedShapes", "text_linked_shapes", shapes=True, context=True),
                ),
                bkt.ribbon.DynamicMenu(
                    id="linked_shapes_actions",
                    label="Aktion ausführen",
                    supertip="Diverse Aktionen auf alle verknüpften Shapes ausführen",
                    image_mso="ObjectBringToFront",
                    get_content=bkt.CallbackLazy("toolbox.models.linkshapes", "action_menu")
                ),
                bkt.ribbon.DynamicMenu(
                    id="linked_shapes_properties",
                    label="Eigenschaft angleichen",
                    supertip="Eine einzelne Eigenschaft auf alle verknüpften Shapes übertragen",
                    image_mso="ObjectNudgeRight",
                    get_content=bkt.CallbackLazy("toolbox.models.linkshapes", "properties_menu")
                ),
                bkt.ribbon.Separator(),
                bkt.ribbon.Button(
                    id = 'linked_shapes_delete',
                    label="Andere Shapes löschen",
                    image_mso="HyperlinkRemove",
                    screentip="Verknüpfte Shapes löschen",
                    supertip="Alle verknüpften Shapes auf allen Folien löschen.",
                    on_action=bkt.CallbackLazy("toolbox.models.linkshapes", "LinkedShapes", "delete_linked_shapes", shapes=True, context=True),
                ),
                bkt.ribbon.Button(
                    id = 'linked_shapes_replace',
                    label="Mit Referenz ersetzen",
                    image_mso="HyperlinkCreate",
                    screentip="Verknüpfte Shapes ersetzen",
                    supertip="Alle verknüpften Shapes auf allen Folien mit Referenz-Shape (standardmäßig das ausgwählte Shape) ersetzen.",
                    on_action=bkt.CallbackLazy("toolbox.models.linkshapes", "LinkedShapes", "replace_with_this", shapes=True, context=True),
                ),
                bkt.ribbon.Button(
                    id = 'linked_shapes_search',
                    label="Weitere Shapes suchen",
                    image_mso="FindTag",
                    screentip="Gleiches Shape auf Folgefolien suchen und verknüpfen",
                    supertip="Erneut nach Shapes anhand Position und Größe suche, um weitere Shapes zu dieser Verknüpfung hinzuzufügen.",
                    get_enabled=bkt.apps.ppt_shapes_exactly1_selected,
                    on_action=bkt.CallbackLazy("toolbox.models.linkshapes", "LinkedShapes", "find_similar_and_link", shape=True, context=True),
                ),
                ### Custom action
                # bkt.ribbon.Button(
                #     id = 'linked_shapes_xyz',
                #     label="Custom Action",
                #     image_mso="HappyFace",
                #     on_action=bkt.CallbackLazy("toolbox.models.linkshapes", "LinkedShapes", "linked_shapes_xyz", shape=True, context=True),
                #     # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                # ),
            ]
        ),
        bkt.ribbon.Group(
            id="bkt_linkshapes_unlink_group",
            label = "Verknüpfung aufheben",
            children = [
                bkt.ribbon.Button(
                    id = 'linked_shapes_unlink',
                    label="Einzelne Shape-Verknüpfung entfernen",
                    image_mso="HyperlinkRemove",
                    screentip="Verknüpfung des ausgewählten Shapes entfernen",
                    supertip="Entfernt die ID zur Verknüpfung vom aktuellen Shape. Alle anderen Shapes mit der gleichen ID bleiben verknüpft.",
                    on_action=bkt.CallbackLazy("toolbox.models.linkshapes", "LinkedShapes", "unlink_shapes", shapes=True),
                ),
                bkt.ribbon.Button(
                    id = 'linked_shapes_unlink_all',
                    label="Gesamte Shape-Verknüpfung auflösen",
                    image_mso="HyperlinkRemove",
                    screentip="Alle Shape-Verknüpfungen entfernen",
                    supertip="Entfernt die ID zur Verknüpfung vom aktuellen Shape sowie allen verknüpften Shapes mit der gleichen ID.",
                    on_action=bkt.CallbackLazy("toolbox.models.linkshapes", "LinkedShapes", "unlink_all_shapes", shapes=True, context=True),
                ),
            ]
        ),
    ]
)