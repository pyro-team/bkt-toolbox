# -*- coding: utf-8 -*-
'''

@author: rdebeerst
'''

import bkt

MODEL_MODULE = __package__ + ".circular_model"
MODEL_CONTAINER = "CircularArrangement"


group_circlify = bkt.ribbon.Group(
    id="bkt_circlify_group",
    label="Kreisanordnung",
    image="circlify",
    supertip="Ermöglicht die kreisförmige Anordnung von Shapes. Das Feature `ppt_circlify` muss installiert sein.",
    children=[
        bkt.ribbon.SplitButton(
            id="circlify_splitbutton",
            size='large',
            children=[
                bkt.ribbon.Button(
                    id="circlify_button",
                    label="Kreisförmig anordnen",
                    image="circlify", #image_mso="DiagramRadialInsertClassic",
                    # size='large',
                    supertip="Ausgewählte Shapes werden Kreis-förmig angeordnet, entsprechend der eingestellten Breite/Höhe.\nDie Reihenfolge der Shapes ist abhängig von der Selektionsreihenfolge: das zuerst selektierte Shape wird auf 12 Uhr positioniert, die weiteren Shapes folgen im Urzeigersinn.",
                    on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "arrange_circular", shapes=True, shapes_min=3),
                    get_enabled="PythonGetEnabled"
                ),
                bkt.ribbon.Menu(
                    label="Kreisanordnung Optionen",
                    supertip="Einstellungen zur kreisförmigen Ausrichtung von Shapes",
                    item_size="large",
                    children=[
                        bkt.ribbon.MenuSeparator(title="Optionen:"),
                        bkt.ribbon.ToggleButton(
                            label="Shape-Rotation an/aus",
                            image_mso="ObjectRotateFree",
                            description="Objekte in der Kreisanordnung entsprechend ihrer Position im Kreis rotieren",
                            on_toggle_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "arrange_circular_rotated"),
                            get_pressed=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "arrange_circular_rotated_pressed")
                        ),
                        bkt.ribbon.ToggleButton(
                            label="Kreis (Breite = Höhe) an/aus",
                            description="Bei Veränderung der Höhe wird auch die Breite geändert und umgekehrt",
                            image_mso="ShapeDonut",
                            on_toggle_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "arrange_circular_fixed"),
                            get_pressed=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "arrange_circular_fixed_pressed")
                        ),
                        bkt.ribbon.ToggleButton(
                            label="Erstes Shapes in Mitte",
                            description="Das zuerst selektierte Shape wird in den Kreis-Mittelpunkt gesetzt",
                            image_mso="DiagramTargetInsertClassic",
                            on_toggle_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "arrange_circular_centerpoint"),
                            get_pressed=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "arrange_circular_centerpoint_pressed")
                        ),
                        bkt.ribbon.MenuSeparator(title="Funktionen:"),
                        bkt.ribbon.Button(
                            label="Aktuelle Parameter interpolieren",
                            description="Es wird versucht den aktuellen Radius, Anfangswinkel und die Optionen der ausgewählten Shapes näherungsweise zu bestimmen",
                            image_mso="DiagramRadialInsertClassic",
                            on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "determine_ellipse_params", shapes=True, shapes_min=3),
                            get_enabled="PythonGetEnabled",
                        ),
                    ]
                ),
            ]
        ),
        bkt.ribbon.RoundingSpinnerBox(
            label="Breite",
            round_cm=True,
            convert = 'pt_to_cm',
            image_mso="ShapeWidth",
            show_label=False,
            supertip="Breite der Ellipse (Diagonale) für die Kreisanordnung",
            on_change=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "set_circ_width", shapes=True, shapes_min=3),
            get_enabled="PythonGetEnabled",
            get_text=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "get_circ_width", shapes=True, shapes_min=3),
        ),
        bkt.ribbon.RoundingSpinnerBox(
            label="Höhe",
            round_cm=True,
            convert = 'pt_to_cm',
            image_mso="ShapeHeight",
            show_label=False,
            supertip="Höhe der Ellipse (Diagonale) für die Kreisanordnung",
            on_change=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "set_circ_height", shapes=True, shapes_min=3),
            get_enabled="PythonGetEnabled",
            get_text=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "get_circ_height", shapes=True, shapes_min=3),
        ),
        bkt.ribbon.RoundingSpinnerBox(
            label="Drehung",
            round_int = True,
            huge_step = 45,
            image_mso="DiagramCycleInsertClassic",
            show_label=False,
            supertip="Winkel des ersten Shapes gibt die Drehung der Kreisanornung an.",
            on_change=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "set_segment_start", shapes=True, shapes_min=3),
            get_enabled="PythonGetEnabled",
            get_text=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "get_segment_start", shapes=True, shapes_min=3),
        ),
    ]
)



bkt.powerpoint.add_tab(bkt.ribbon.Tab(
    id="bkt_powerpoint_toolbox_extensions",
    insert_before_mso="TabHome",
    label='Toolbox 3/3',
    # get_visible defaults to False during async-startup
    get_visible=bkt.Callback(lambda:True),
    children = [
        group_circlify,
        # group_segmented_circle
    ]
), extend=True)

