# -*- coding: utf-8 -*-
'''
Created on 06.07.2016

@author: rdebeerst
'''

import bkt
import bkt.library.powerpoint as powerpoint


class HarveyBallsUi(object):
    BKT_HARVEY_DIALOG_TAG = "BKT_SHAPE_HARVEY"
    BKT_HARVEY_DENOM_TAG = "BKT_HARVEY_DENOM_TAG"
    BKT_HARVEY_VERSION = "BKT_HARVEY_V2"

    BKT_HARVEY_LEGACY_VERSION = ("BKT_HARVEY_V1")

    # =====================================
    # = Feature-Logik und Hilfsfunktionen =
    # =====================================
    
    def is_harvey_group(self, shape):
        # "new" method via tags
        if powerpoint.TagHelper.has_tag(shape, self.BKT_HARVEY_DIALOG_TAG):
            return True
        # "old" method via shape types
        pie, _ = self.get_pie_circ(shape)
        return pie != None

    def get_pie_circ(self, shape):
        if not shape.Type == powerpoint.MsoShapeType['msoGroup']:
            return None, None
        if not shape.GroupItems.Count == 2:
            return None, None

        pie_types = (powerpoint.MsoAutoShapeType['msoShapePie'],powerpoint.MsoAutoShapeType['msoShapeArc'],powerpoint.MsoAutoShapeType['msoShapeBlockArc'])

        if shape.GroupItems(1).AutoShapeType in pie_types:
            return shape.GroupItems(1), shape.GroupItems(2)
        elif shape.GroupItems(2).AutoShapeType in pie_types:
            return shape.GroupItems(2), shape.GroupItems(1)
        else:
            return None, None

    def change_harvey_enabled(self, shapes):
        return self.is_harvey_group(shapes[0])

harveyui = HarveyBallsUi()


def harvey_color_gallery(**kwargs):
    return bkt.ribbon.ColorGallery(
        label = 'Farbe ändern',
        #image_mso = 'RecolorColorPicker',
        image='harvey ball color',
        screentip="Farbe eines Harvey-Balls ändern",
        supertip="Passe die Farbe eines Harvey-Balls entsprechend der Auswahl an.\n\nEin Harvey-Ball-Shape ist eine Gruppe aus Kreis- und Torten-Shape.",
        on_rgb_color_change   = bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "color_gallery_action", shapes=True),
        on_theme_color_change = bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "color_gallery_theme_color_change", shapes=True),
        get_selected_color    = bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "get_selected_color", shapes=True),
        get_enabled           = bkt.Callback(harveyui.change_harvey_enabled, shapes=True),
        item_width=16, item_height=16,
        **kwargs
    )

def harvey_background_gallery(**kwargs):
    return bkt.ribbon.ColorGallery(
        label = 'Hintergrund ändern',
        #image_mso = 'RecolorColorPicker',
        image='harvey ball background',
        screentip="Hintergrund eines Harvey-Balls ändern",
        supertip="Passe die Hintergrund-Farbe eines Harvey-Balls entsprechend der Auswahl an.\n\nEin Harvey-Ball-Shape ist eine Gruppe aus Kreis- und Torten-Shape.",
        on_rgb_color_change   = bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "background_gallery_action", shapes=True),
        on_theme_color_change = bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "background_gallery_theme_color_change", shapes=True),
        get_selected_color    = bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "get_selected_background", shapes=True),
        get_enabled           = bkt.Callback(harveyui.change_harvey_enabled, shapes=True),
        children=[
            bkt.ribbon.Button(
                label='Hintergrund aus',
                supertip="Harvey-Ball Hintergrund auf transparent setzen",
                #get_image=bkt.Callback(lambda: harvey_balls.get_harvey_image(0.6, 64)),
                image='harvey ball background',
                on_action=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "harvey_background_off", shapes=True),
            ),
        ],
        item_width=16, item_height=16,
        **kwargs
    )

def harvey_line_gallery(**kwargs):
    return bkt.ribbon.ColorGallery(
        label = 'Linienfarbe ändern',
        #image_mso = 'RecolorColorPicker',
        image='harvey ball line',
        screentip="Linienfarbe eines Harvey-Balls ändern",
        supertip="Passe die Linienfarbe eines Harvey-Balls entsprechend der Auswahl an.\n\nEin Harvey-Ball-Shape ist eine Gruppe aus Kreis- und Torten-Shape.",
        on_rgb_color_change   = bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "line_gallery_action", shapes=True),
        on_theme_color_change = bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "line_gallery_theme_color_change", shapes=True),
        get_selected_color    = bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "get_selected_line", shapes=True),
        get_enabled           = bkt.Callback(harveyui.change_harvey_enabled, shapes=True),
        children=[
            bkt.ribbon.Button(
                label='Linie aus',
                supertip="Harvey-Ball Linie ausblenden",
                #get_image=bkt.Callback(lambda: harvey_balls.get_harvey_image(0.6, 64)),
                image='harvey ball line off',
                on_action=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "harvey_line_off", shapes=True),
            ),
            bkt.ribbon.Button(
                label='Linie nur außen ein/aus',
                supertip="Harvey-Ball Linie wird nur entweder um den ganzen Kuchen oder nur um den äußeren Kreis angezeigt",
                #get_image=bkt.Callback(lambda: harvey_balls.get_harvey_image(0.6, 64)),
                image='harvey ball line outside',
                on_action=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "harvey_line_outside_only", shapes=True),
            ),
        ],
        item_width=16, item_height=16,
        **kwargs
    )

def harvey_size_gallery(**kwargs):
    return bkt.ribbon.Gallery(
        label = 'Füllstand ändern',
        image = 'harvey ball size',
        #get_image=bkt.Callback(lambda: harvey_balls.get_harvey_image(0.6, 64)),
        screentip="Füllstand eines Harvey-Balls ändern",
        supertip="Passe den Füllstand eines Harvey-Balls entsprechend der Auswahl an.\n\nEin Harvey-Ball-Shape ist eine Gruppe aus Kreis- und Torten-Shape.",
        columns="9", #9=harvey_columns
        on_action_indexed = bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "change_harvey", shapes=True),
        get_item_count    = bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "get_item_count"),
        get_item_label    = bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "get_item_label"),
        get_item_screentip = bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "get_item_screentip"),
        get_item_supertip = bkt.Callback(lambda index: "Passe den Füllstand eines Harvey-Balls entsprechend der Auswahl an."),
        get_enabled       = bkt.Callback(harveyui.change_harvey_enabled, shapes=True),
        get_item_image    = bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "get_harvey_item_image"),
        item_width=16, item_height=16,
        **kwargs
    )


harvey_create_button = bkt.ribbon.Button(
    id='create_harvey_ball',
    label='Harvey Ball',
    screentip='Harvey Ball erstellen',
    image='harvey ball',
    on_action=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "create_harvey_ball", context=True, slide=True),
    supertip="Füge ein Harvey-Ball ein, welcher sich bzgl. Farbe/Füllstand konfigurieren lässt.\n\nFarbe und Füllstand lassen sich über Kontext-Menü und Kontext-Tab konfigurieren, im Tab auch Prozent-Angaben möglich.\n\nEin Harvey-Ball-Shape ist eine Gruppe aus Kreis- und Torten-Shape."
)


harvey_ball_group = bkt.ribbon.Group(
    id="bkt_harvey_group",
    label = "Harvey Balls",
    children = [
        bkt.ribbon.Button(
            id='harvey_ball_create',
            size='large',
            label='Neuer Harvey Ball',
            screentip='Harvey Ball erstellen',
            image='harvey ball',
            on_action=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "create_harvey_ball", context=True, slide=True),
            supertip="Füge ein Harvey-Ball ein, welcher sich bzgl. Farbe/Füllstand konfigurieren lässt.\n\nFarbe und Füllstand lassen sich über Kontext-Menü und Kontext-Tab konfigurieren, im Tab auch Prozent-Angaben möglich.\n\nEin Harvey-Ball-Shape ist eine Gruppe aus Kreis- und Torten-Shape."
        ),
        bkt.ribbon.Button(
            id='harvey_ball_duplicate',
            size='large',
            label='Harvey Ball duplizieren',
            screentip='Harvey Ball duplizieren',
            image='harvey ball duplicate',
            on_action=bkt.Callback(lambda selection: selection.ShapeRange.Duplicate().Select()),
            supertip="Dupliziert den aktuell ausgewählten Harvey-Ball."
        ),
        bkt.ribbon.Separator(),

        #bkt.ribbon.SplitButton(show_label=False, children=[
            # bkt.ribbon.Button(
            #     id='create_harvey_ball',
            #     label='Harvey Ball erstellen',
            #     screentip='Harvey Ball erstellen',
            #     image='harvey ball',
            #     on_action=bkt.Callback(harvey_balls.create_harvey_ball)
            # ),
            # bkt.ribbon.Menu(label='menu',
            #     children = [
        harvey_size_gallery(id='harvey_ball_size_gallery', size="large"),
        harvey_color_gallery(id='harvey_ball_color_gallery', size="large"),
        #         ]
        #     )
        # ]),

        harvey_background_gallery(id='harvey_ball_background', size="large"),
        harvey_line_gallery(id='harvey_ball_line', size="large"),

        bkt.ribbon.Separator(),

        bkt.ribbon.Button(
            id='harvey_ball_style_classic',
            size='large',
            label='Style klassisch',
            supertip="Harvey-Ball im klassischen Style ohne zusätzlichem Rand darstellen.",
            image='harvey ball classic',
            on_action=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "harvey_change_style_classic", shapes=True),
        ),

        bkt.ribbon.Button(
            id='harvey_ball_style_modern',
            size='large',
            label='Style modern',
            supertip="Harvey-Ball im modernen Style mit weißem Rand darstellen.",
            image='harvey ball modern',
            on_action=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "harvey_change_style_modern", shapes=True),
        ),

        bkt.ribbon.Button(
            id='harvey_ball_style_chart',
            size='large',
            label='Style Diagramm',
            supertip="Harvey-Ball im Diagramm-Style mit hervorgehobenem Füllstand.",
            image='harvey ball diagram',
            on_action=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "harvey_change_style_chart", shapes=True),
        ),

        bkt.ribbon.ToggleButton(
            id='harvey_ball_flip',
            size='large',
            label='Gegen Uhrzeigersinn',
            supertip="Harvey-Ball spiegel, um Füllstand mit oder gegen den Uhrzeigersinn anzuzeigen.",
            image='harvey ball flip',
            get_pressed=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "harvey_fliph_pressed", shapes=True),
            on_toggle_action=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "harvey_fliph", shapes=True),
        ),

        bkt.ribbon.Separator(),
        #bkt.ribbon.LabelControl(label="Füllstand:"),
        
        bkt.ribbon.SpinnerBox(label='Füllstand in %', size_string='33,33',
            id = 'harvey_spinner',
            screentip="Füllstand eines Harvey-Balls ändern",
            supertip="Passe den Füllstand eines Harvey-Balls entsprechend der Auswahl an.\n\nEin Harvey-Ball-Shape ist eine Gruppe aus Kreis- und Torten-Shape.",
            on_change = bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "harvey_percent_setter", shapes=True),
            get_text  = bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "get_harvey_percent", shapes=True),
            increment = bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "harvey_percent_inc", shapes=True),
            decrement = bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "harvey_percent_dec", shapes=True)
        ),
        bkt.ribbon.LabelControl(label="   mit SHIFT: Schrittweite +/-25"),
        bkt.ribbon.LabelControl(label="   mit ALT: Delta je Harvey Ball"),

        bkt.ribbon.Separator(),

        bkt.ribbon.Button(
            id='harvey_ball_0',
            size='large',
            label='0%',
            screentip="Harvey-Ball auf 0%",
            supertip="Setzt alle gewählten Harvey-Balls auf 0%",
            get_image=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "get_harvey_image_by_control", current_control=True),
            on_action=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "set_harvey_by_control", shapes=True, current_control=True),
            tag="0",
        ),
        bkt.ribbon.Button(
            id='harvey_ball_25',
            size='large',
            label='25%',
            screentip="Harvey-Ball auf 25%",
            supertip="Setzt alle gewählten Harvey-Balls auf 25%",
            get_image=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "get_harvey_image_by_control", current_control=True),
            on_action=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "set_harvey_by_control", shapes=True, current_control=True),
            tag="25",
        ),
        bkt.ribbon.Button(
            id='harvey_ball_33',
            size='large',
            label='33%',
            screentip="Harvey-Ball auf 33%",
            supertip="Setzt alle gewählten Harvey-Balls auf 33%",
            get_image=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "get_harvey_image_by_control", current_control=True),
            on_action=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "set_harvey_by_control", shapes=True, current_control=True),
            tag="33.3",
        ),
        bkt.ribbon.Button(
            id='harvey_ball_50',
            size='large',
            label='50%',
            screentip="Harvey-Ball auf 50%",
            supertip="Setzt alle gewählten Harvey-Balls auf 50%",
            get_image=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "get_harvey_image_by_control", current_control=True),
            on_action=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "set_harvey_by_control", shapes=True, current_control=True),
            tag="50",
        ),
        bkt.ribbon.Button(
            id='harvey_ball_66',
            size='large',
            label='66%',
            screentip="Harvey-Ball auf 66%",
            supertip="Setzt alle gewählten Harvey-Balls auf 66%",
            get_image=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "get_harvey_image_by_control", current_control=True),
            on_action=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "set_harvey_by_control", shapes=True, current_control=True),
            tag="66.6",
        ),
        bkt.ribbon.Button(
            id='harvey_ball_75',
            size='large',
            label='75%',
            screentip="Harvey-Ball auf 75%",
            supertip="Setzt alle gewählten Harvey-Balls auf 75%",
            get_image=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "get_harvey_image_by_control", current_control=True),
            on_action=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "set_harvey_by_control", shapes=True, current_control=True),
            tag="75",
        ),
        bkt.ribbon.Button(
            id='harvey_ball_100',
            size='large',
            label='100%',
            screentip="Harvey-Ball auf 100%",
            supertip="Setzt alle gewählten Harvey-Balls auf 100%",
            get_image=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "get_harvey_image_by_control", current_control=True),
            on_action=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "set_harvey_by_control", shapes=True, current_control=True),
            tag="100",
        ),

        # bkt.ribbon.Separator(),

        # bkt.ribbon.Button(
        #     id='harvey_legacy_upgrade',
        #     size='large',
        #     label='Version aktualisieren',
        #     screentip="Harvey-Ball auf neuste Version aktualisieren",
        #     supertip="Harvey-Ball wird auf die neueste Version aktualisiert, wodurch er optisch etwas besser aussieht und die Linienfarbe angepasst werden kann. Allerdings ist er dann nicht mehr mit älteren Versionen kompatibel.",
        #     image_mso="UpdateIcon",
        #     get_visible=bkt.Callback(harvey_balls.is_legacy_any),
        #     on_action=bkt.Callback(harvey_balls.upgrade_all),
        # ),
    ]
)

harvey_ball_tab = bkt.ribbon.Tab(
    id = "bkt_context_tab_harvey",
    label = "[BKT] Harvey Balls",
    get_visible=bkt.Callback(harveyui.change_harvey_enabled, shapes=True),
    children = [
        # Harvey Balls
        harvey_ball_group
    ]
)