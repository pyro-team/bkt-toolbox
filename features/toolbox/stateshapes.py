# -*- coding: utf-8 -*-
'''
Created on 21.12.2017

@author: fstallmann
'''


import bkt
import bkt.library.powerpoint as pplib


class StateShapeUi(object):
    BKT_DIALOG_TAG = 'BKT_DIALOG_STATESHAPE'

    @classmethod
    def is_convertable_to_state_shape(cls, shapes):
        try:
            if len(shapes) > 1:
                return not any(cls.is_state_shape(s) for s in shapes)
            else:
                shape = shapes[0]
                return shape.Type == pplib.MsoShapeType['msoGroup'] and not cls.is_state_shape(shape)
        except:
            return False

    @classmethod
    def is_state_shape(cls, shape):
        return pplib.TagHelper.has_tag(shape, bkt.contextdialogs.BKT_CONTEXTDIALOG_TAGKEY, cls.BKT_DIALOG_TAG)
        # return shape.Type == pplib.MsoShapeType['msoGroup']
    
    @classmethod
    def are_state_shapes(cls, shapes):
        return all(cls.is_state_shape(s) for s in shapes)



def stateshape_fill1_gallery(**kwargs):
    return bkt.ribbon.ColorGallery(
                    label = 'Farbe 1 (Hintergrund) ändern',
                    image_mso = 'ShapeFillColorPicker',
                    screentip="Hintergrundfarbe eines Wechsel-Shapes ändern",
                    supertip="Passt die Hintergrundfarbe aller Shapes im Wechsel-Shape an. Die Hintergrundfarbe ist die Farbe des zuerst gefundenen Shapes.",
                    on_rgb_color_change   = bkt.CallbackLazy("toolbox.models.stateshapes", "StateShape", "set_color_fill_rgb1", shapes=True),
                    on_theme_color_change = bkt.CallbackLazy("toolbox.models.stateshapes", "StateShape", "set_color_fill_theme1", shapes=True),
                    # get_selected_color    = bkt.CallbackLazy("toolbox.models.stateshapes", "StateShape", "get_selected_color1", shapes=True),
                    get_enabled           = bkt.Callback(StateShapeUi.are_state_shapes, shapes=True),
                    children=[
                        bkt.ribbon.Button(
                            label="Kein Hintergrund",
                            supertip="Wechsel-Shape Hintergrundfarbe auf transparent setzen",
                            on_action=bkt.CallbackLazy("toolbox.models.stateshapes", "StateShape", "set_color_fill_none1", shapes=True),
                        ),
                    ],
                    **kwargs
                )

def stateshape_fill2_gallery(**kwargs):
    return bkt.ribbon.ColorGallery(
                    label = 'Farbe 2 (Vordergrund) ändern',
                    image_mso = 'ShapeFillColorPicker',
                    screentip="Vordergrundfarbe eines Wechsel-Shapes ändern",
                    supertip="Passt die Vordergrundfarbe aller Shapes im Wechsel-Shape an. Die Vordergrundfarbe ist jede Farbe ungleich der Hintergrundfarbe.",
                    on_rgb_color_change   = bkt.CallbackLazy("toolbox.models.stateshapes", "StateShape", "set_color_fill_rgb2", shapes=True),
                    on_theme_color_change = bkt.CallbackLazy("toolbox.models.stateshapes", "StateShape", "set_color_fill_theme2", shapes=True),
                    # get_selected_color    = bkt.CallbackLazy("toolbox.models.stateshapes", "StateShape", "get_selected_color2", shapes=True),
                    get_enabled           = bkt.Callback(StateShapeUi.are_state_shapes, shapes=True),
                    **kwargs
                )

def stateshape_line_gallery(**kwargs):
    return bkt.ribbon.ColorGallery(
                    label = 'Linie ändern',
                    image_mso = 'ShapeOutlineColorPicker',
                    screentip="Linie eines Wechsel-Shapes ändern",
                    supertip="Passt die Linienfarbe aller Shapes im Wechsel-Shape an, die der ersten gefundenen Linienfarbe entsprechen.",
                    on_rgb_color_change   = bkt.CallbackLazy("toolbox.models.stateshapes", "StateShape", "set_color_line_rgb", shapes=True),
                    on_theme_color_change = bkt.CallbackLazy("toolbox.models.stateshapes", "StateShape", "set_color_line_theme", shapes=True),
                    # get_selected_color    = bkt.CallbackLazy("toolbox.models.stateshapes", "StateShape", "get_selected_line", shapes=True),
                    get_enabled           = bkt.Callback(StateShapeUi.are_state_shapes, shapes=True),
                    children=[
                        bkt.ribbon.Button(
                            label="Keine Linie",
                            supertip="Wechsel-Shape Linie auf transparent setzen",
                            on_action=bkt.CallbackLazy("toolbox.models.stateshapes", "StateShape", "set_color_line_none", shapes=True),
                        ),
                    ],
                    **kwargs
                )



stateshape_gruppe = bkt.ribbon.Group(
    id="bkt_stateshape_group",
    label='Wechsel-Shapes',
    image_mso='GroupSmartArtQuickStyles',
    children = [
        bkt.ribbon.SplitButton(
            id="stateshape_convert_splitbutton",
            # size="large",
            children=[
                bkt.ribbon.Button(
                    id="stateshape_convert",
                    label="Konvertieren",
                    image_mso='GroupSmartArtQuickStyles',
                    screentip="Gruppierte Shapes in ein Wechselshape konvertieren",
                    supertip="Bei gruppierten Shapes (Wechsel-Shapes) kann zwischen den Shapes innerhalb der Gruppe gewechselt werden, d.h. es ist immer nur ein Shape der Gruppe sichtbar. Dies ist bspw. nützlich für Ampeln, Skalen, etc.",
                    on_action=bkt.CallbackLazy("toolbox.models.stateshapes", "StateShape", "convert_to_state_shape", shapes=True),
                    get_enabled=bkt.Callback(StateShapeUi.is_convertable_to_state_shape, shapes=True),
                ),
                bkt.ribbon.Menu(
                    label="Wechselshapes-Menü",
                    supertip="In Wechselshapes konvertieren oder wieder alle Shapes sichtbar machen",
                    children=[
                        bkt.ribbon.Button(
                            id="stateshape_convert2",
                            label="In Wechselshape konvertieren",
                            image_mso='GroupSmartArtQuickStyles',
                            screentip="Gruppierte Shapes in ein Wechselshape konvertieren",
                            supertip="Bei gruppierten Shapes (Wechsel-Shapes) kann zwischen den Shapes innerhalb der Gruppe gewechselt werden, d.h. es ist immer nur ein Shape der Gruppe sichtbar. Dies ist bspw. nützlich für Ampeln, Skalen, etc.",
                            on_action=bkt.CallbackLazy("toolbox.models.stateshapes", "StateShape", "convert_to_state_shape", shapes=True),
                            get_enabled=bkt.Callback(StateShapeUi.is_convertable_to_state_shape, shapes=True),
                        ),
                        bkt.ribbon.MenuSeparator(),
                        # bkt.ribbon.ToggleButton(
                        bkt.ribbon.Button(
                            id="stateshape_show_all",
                            label="Alle Shapes wieder anzeigen",
                            screentip="Alle Shapes sichtbar machen",
                            supertip="Mit diesem Button können die Shapes innerhalb der Wechselshape-Gruppe eingeblendet werden.",
                            # image_mso='GroupSmartArtQuickStyles',
                            # get_pressed=bkt.Callback(StateShape.get_show_all),
                            # on_toggle_action=bkt.Callback(StateShape.toggle_show_all),
                            on_action=bkt.CallbackLazy("toolbox.models.stateshapes", "StateShape", "show_all", shape=True),
                            get_enabled=bkt.Callback(StateShapeUi.is_state_shape, shape=True),
                        ),
                    ]
                )
            ]
        ),
        # bkt.ribbon.Separator(),
        # bkt.ribbon.LabelControl(label="Wechsel: "),
        bkt.ribbon.Box(box_style="horizontal", children=[
            bkt.ribbon.Button(
                id="stateshape_reset",
                image_mso="Undo",
                label="Zurücksetzen",
                show_label=False,
                screentip="Auf erstes Shape zurücksetzen",
                supertip="Setzt alle Wechsel-Shapes auf den ersten Status, d.h. das erste Shape der Gruppe zurück.",
                on_action=bkt.CallbackLazy("toolbox.models.stateshapes", "StateShape", "reset_state", shapes=True),
                get_enabled=bkt.Callback(StateShapeUi.are_state_shapes, shapes=True),
            ),
            bkt.ribbon.Button(
                id="stateshape_prev",
                image_mso="PreviousResource",
                label='Vorheriges',
                show_label=False,
                screentip="Vorheriges Shape",
                supertip="Wechselt zum vorherigen Status (d.h. Shape in der Gruppe) des Wechsel-Shapes.",
                on_action=bkt.CallbackLazy("toolbox.models.stateshapes", "StateShape", "previous_state", shapes=True),
                get_enabled=bkt.Callback(StateShapeUi.are_state_shapes, shapes=True),
            ),
            # bkt.ribbon.EditBox(
            #     id="stateshape_set",
            #     label="Position",
            #     show_label=False,
            #     size_string="#",
            #     on_change=bkt.CallbackLazy("toolbox.models.stateshapes", "StateShape", "set_state"),
            #     get_enabled=bkt.Callback(StateShapeUi.are_state_shapes),
            #     get_text=bkt.Callback(lambda: None),
            # ),
            bkt.ribbon.Button(
                id="stateshape_next",
                image_mso="NextResource",
                label="Nächstes",
                # show_label=False,
                screentip="Nächstes Shape",
                supertip="Wechselt zum nächsten Status (d.h. Shape in der Gruppe) des Wechsel-Shapes.",
                on_action=bkt.CallbackLazy("toolbox.models.stateshapes", "StateShape", "next_state", shapes=True),
                get_enabled=bkt.Callback(StateShapeUi.are_state_shapes, shapes=True),
            )
        ]),
        bkt.ribbon.Menu(
            id="stateshape_color_menu",
            label="Farbe ändern",
            supertip="Die Farben von Wechselshapes anpassen",
            image_mso="RecolorColorPicker",
            children=[
                stateshape_fill1_gallery(id="stateshape_color_fill1"),
                stateshape_fill2_gallery(id="stateshape_color_fill2"),
                stateshape_line_gallery(id="stateshape_color_line"),
            ]
        ),
        # bkt.ribbon.Button(
        #     id="stateshape_help",
        #     image_mso="Help",
        #     label=u"Anleitung",
        #     on_action=bkt.Callback(StateShape.show_help),
        #     # get_enabled=bkt.Callback(StateShape.are_state_shapes),
        # ),
        # likert_button,
    ]
)
