# -*- coding: utf-8 -*-
'''
Created on 17.03.2017

@author: fstallmann
'''

import bkt
import math
import logging

#from bkt.library import visio


#FIXME: This whole file does currently not consider Shape Protection (LockWidth, LockHeight, LockMoveX, LockMoveY, LockRotate, etc.) -> should this even be in VisioWrapper?


class ObjektAbstand(object):
    
    @staticmethod
    def set_shape_sep_vertical(shapes, value):
        master = shapes[0]
        apply_to_upper = [shape for shape in shapes[1:] if shape._bottom > master._bottom]
        apply_to_lower = [shape for shape in shapes[1:] if shape._bottom <= master._bottom]
        apply_to_upper.sort(key=lambda shape: shape._bottom)
        apply_to_lower.sort(key=lambda shape: shape._bottom, reverse=True)
        if type(value) == str:
            value = float(value.replace(',', '.'))
        top = master._bottom
        for shape in apply_to_lower:
            top -= shape._height + value
            shape._bottom = top
        top = master._bottom + master._height + value
        for shape in apply_to_upper:
            shape._bottom = top
            top += shape._height + value

    @staticmethod
    def get_shape_sep_vertical(shapes):
        shapes = sorted(shapes, key=lambda shape: (shape._bottom, shape._left))
        return round(shapes[1]._bottom-shapes[0]._bottom-shapes[0]._height,2)
        # return round(pt_to_cm(shapes[1]._bottom-shapes[0]._bottom-shapes[0].height),2)

    @staticmethod
    def set_shape_sep_horizontal(shapes, value):
        master = shapes[0]
        apply_to_righthand = [shape for shape in shapes[1:] if shape._left > master._left]
        apply_to_lefthand = [shape for shape in shapes[1:] if shape._left <= master._left]
        apply_to_righthand.sort(key=lambda shape: shape._left)
        apply_to_lefthand.sort(key=lambda shape: shape._left, reverse=True)
        if type(value) == str:
            value = float(value.replace(',', '.'))
        top = master._left
        for shape in apply_to_lefthand:
            top -= shape._width + value
            shape._left = top
        top = master._left + master._width + value
        for shape in apply_to_righthand:
            shape._left = top
            top += shape._width + value

    @staticmethod
    def get_shape_sep_horizontal(shapes):
        shapes = sorted(shapes, key=lambda shape: (shape._left, shape._bottom))
        return round(shapes[1]._left-shapes[0]._left-shapes[0]._width,2)
        # return round(pt_to_cm(shapes[1]._left-shapes[0]._left-shapes[0].width),2)

    @staticmethod
    def set_rotation(shapes, value):
        value = math.radians(value)
        for shape in shapes:
            shape.angle = value

    @staticmethod
    def get_rotation(shapes):
        return round(math.degrees(shapes[0].angle),1)

    @staticmethod
    def enabled_rotation(shapes):
        return all(shape.shape.OneD == 0 for shape in shapes)


objektabstand_gruppe = bkt.ribbon.Group(
    label="Abstand/Rotation",
    image_mso='VerticalSpacingIncrease',
    children=[
        bkt.ribbon.RoundingSpinnerBox(
            id = 'shape_sep_v',
            label=u"Objektabstand vertikal",
            show_label=False,
            image_mso='VerticalSpacingIncrease',
            on_change = bkt.Callback(ObjektAbstand.set_shape_sep_vertical, shapes=True, shapes_min=2),
            get_text  = bkt.Callback(ObjektAbstand.get_shape_sep_vertical, shapes=True, shapes_min=2),
            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            round_cm = True
        ),

        bkt.ribbon.RoundingSpinnerBox(
            id = 'shape_sep_h',
            label=u"Objektabstand horizontal",
            show_label=False,
            image_mso='HorizontalSpacingIncrease',
            on_change = bkt.Callback(ObjektAbstand.set_shape_sep_horizontal, shapes=True, shapes_min=2),
            get_text  = bkt.Callback(ObjektAbstand.get_shape_sep_horizontal, shapes=True, shapes_min=2),
            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            round_cm = True
        ),

        bkt.ribbon.RoundingSpinnerBox(
            id = 'rotation',
            label=u"Rotation",
            show_label=False,
            image_mso='RotationTool',
            on_change = bkt.Callback(ObjektAbstand.set_rotation, shapes=True, shapes_min=1),
            get_text  = bkt.Callback(ObjektAbstand.get_rotation, shapes=True, shapes_min=1),
            get_enabled = bkt.Callback(ObjektAbstand.enabled_rotation, shapes=True, shapes_min=1),
            round_int = True
        )
    ]
)

class PositionSize(object):

    @staticmethod
    def set_y(shapes, value):
        # value = cm_to_pt(value)
        bkt.apply_delta_on_ALT_key(
            lambda shape, value: setattr(shape, '_y', value), 
            lambda shape: shape._y, 
            shapes, value)
    
    @staticmethod
    def get_y(shapes):
        return round(shapes[0]._y,2)
        # return round(pt_to_cm(shapes[0].y),2)
    
    
    @staticmethod
    def set_x(shapes, value):
        # value = cm_to_pt(value)
        bkt.apply_delta_on_ALT_key(
            lambda shape, value: setattr(shape, '_x', value), 
            lambda shape: shape._x, 
            shapes, value)

    @staticmethod
    def get_x(shapes):
        return round(shapes[0]._x,2)
        # return round(pt_to_cm(shapes[0].x),2)


    @staticmethod
    def set_height(shapes, value):
        # value = cm_to_pt(value)
        bkt.apply_delta_on_ALT_key(
            lambda shape, value: setattr(shape, '_height', value), 
            lambda shape: shape._height, 
            shapes, value)

    @staticmethod
    def get_height(shapes):
        return round(shapes[0]._height,2)
        # return round(pt_to_cm(shapes[0].height),2)
    
    
    @staticmethod
    def set_width(shapes, value):
        # value = cm_to_pt(value)
        bkt.apply_delta_on_ALT_key(
            lambda shape, value: setattr(shape, '_width', value), 
            lambda shape: shape._width, 
            shapes, value)

    @staticmethod
    def get_width(shapes):
        return round(shapes[0]._width,2)
        # return round(pt_to_cm(shapes[0].width),2)
    
    @staticmethod
    def round_position(shapes):
        for shape in shapes:
            shape._x = round(shape._x)
            shape._y = round(shape._y)
    
    @staticmethod
    def round_size(shapes):
        for shape in shapes:
            shape._width = round(shape._width)
            shape._height = round(shape._height)


class LocPin(object):
    @staticmethod
    def get_pins_state(shapes, left=0, bottom=0):
        return all(round(shape.localpinx) == round(left*shape.width) and round(shape.localpiny) == round(bottom*shape.height) for shape in shapes)

    @classmethod
    def set_pins_state(cls, shapes, left=0, bottom=0):
        for shape in shapes:
            if shape.shape.OneD == -1: #1D shapes (connectors) do not support locpins
                continue
            #store old locpin values
            old_x = shape._left
            old_y = shape._bottom
            #change pin
            shape.localpinx_formula = "Width*"+str(left)
            shape.localpiny_formula = "Height*"+str(bottom)
            #restore position of shape
            shape._left = old_x
            shape._bottom = old_y

    @staticmethod
    def get_enabled(shapes):
        return all(shape.shape.OneD == 0 for shape in shapes)


pos_size_group = bkt.ribbon.Group(
    label='Position/Größe',
    image_mso='ShapeWidth',
    children =[
        bkt.ribbon.Box(
            box_style="vertical",
            children =[
                bkt.ribbon.RoundingSpinnerBox(
                    id="spinner_pos_bottom",
                    round_cm=True,
                    image_mso='ObjectNudgeUp', 
                    screentip="Position von unten",
                    supertip="Änderung der Position von unten.\n\nBei gedrückter STRG-Taste Veränderung um 0,1 cm statt 0,2 cm.\n\nBei gedrückter ALT-Taste Veränderung relativ je Shape (wenn mehrere Shapes ausgewählt sind).",
                    on_change=bkt.Callback(PositionSize.set_y, shapes=True, shapes_min=1),
                    get_text=bkt.Callback(PositionSize.get_y, shapes=True, shapes_min=1),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name
                ),
                bkt.ribbon.RoundingSpinnerBox(
                    id="spinner_pos_left",
                    round_cm=True,
                    screentip="Position von links",
                    supertip="Änderung der Position von links.\n\nBei gedrückter STRG-Taste Veränderung um 0,1 cm statt 0,2 cm.\n\nBei gedrückter ALT-Taste Veränderung relativ je Shape (wenn mehrere Shapes ausgewählt sind).",
                    image_mso='ObjectNudgeRight',
                    on_change=bkt.Callback(PositionSize.set_x, shapes=True, shapes_min=1),
                    get_text=bkt.Callback(PositionSize.get_x, shapes=True, shapes_min=1),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name
                ),
                bkt.ribbon.Button(
                    id = 'round_position',
                    label="Pos. runden",
                    show_label=True,
                    image_mso='ObjectsAlignToGridOutlook',
                    screentip="Shape-Koordinaten auf ganze Zahlen runden",
                    on_action=bkt.Callback(PositionSize.round_position, shapes=True, shapes_min=1),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name
                ),
            ]
        ),
        bkt.ribbon.Box(
            box_style="vertical",
            children =[
                bkt.ribbon.RoundingSpinnerBox(
                    id="spinner_size_height",
                    round_cm=True,
                    image_mso='ShapeHeight',
                    screentip="Höhe",
                    supertip="Änderung der Höhe.\n\nBei gedrückter STRG-Taste Veränderung um 0,1 cm statt 0,2 cm.\n\nBei gedrückter ALT-Taste Veränderung relativ je Shape (wenn mehrere Shapes ausgewählt sind).",
                    on_change=bkt.Callback(PositionSize.set_height, shapes=True, shapes_min=1),
                    get_text=bkt.Callback(PositionSize.get_height, shapes=True, shapes_min=1),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name
                ),
                bkt.ribbon.RoundingSpinnerBox(
                    id="spinner_size_width",
                    round_cm=True,
                    image_mso='ShapeWidth',
                    screentip="Breite",
                    supertip="Änderung der Breite.\n\nBei gedrückter STRG-Taste Veränderung um 0,1 cm statt 0,2 cm.\n\nBei gedrückter ALT-Taste Veränderung relativ je Shape (wenn mehrere Shapes ausgewählt sind).",
                    on_change=bkt.Callback(PositionSize.set_width, shapes=True, shapes_min=1),
                    get_text=bkt.Callback(PositionSize.get_width, shapes=True, shapes_min=1),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name
                ),
                bkt.ribbon.Button(
                    id = 'round_size',
                    label="Größe runden",
                    show_label=True,
                    image_mso='SizeToGridOutlook',
                    screentip="Shape-Größe auf ganze Zahlen runden",
                    on_action=bkt.Callback(PositionSize.round_size, shapes=True, shapes_min=1),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name
                ),
            ]
        ),
        bkt.ribbon.Separator(),
        bkt.ribbon.Box(
            box_style="horizontal",
            children =[
                bkt.ribbon.ToggleButton(
                    id="fix_loc_pins_0_1",
                    label="Bezugspos. oben/links",
                    show_label=False,
                    image="fix_loc_pins_0_1",
                    # size="large",
                    on_toggle_action=bkt.Callback(lambda shapes, pressed: LocPin.set_pins_state(shapes, 0, 1), shapes=True, shapes_min=1),
                    get_pressed=bkt.Callback(lambda shapes: LocPin.get_pins_state(shapes, 0, 1), shapes=True, shapes_min=1),
                    get_enabled = bkt.Callback(LocPin.get_enabled, shapes=True, shapes_min=1),
                    # screentip="Die Bezugs-Position gewählter Shapes auf linke untere Ecke ändern"
                ),
                bkt.ribbon.ToggleButton(
                    id="fix_loc_pins_05_1",
                    label="Bezugspos. oben/mitte",
                    show_label=False,
                    image="fix_loc_pins_05_1",
                    # size="large",
                    on_toggle_action=bkt.Callback(lambda shapes, pressed: LocPin.set_pins_state(shapes, 0.5, 1), shapes=True, shapes_min=1),
                    get_pressed=bkt.Callback(lambda shapes: LocPin.get_pins_state(shapes, 0.5, 1), shapes=True, shapes_min=1),
                    get_enabled = bkt.Callback(LocPin.get_enabled, shapes=True, shapes_min=1),
                    # screentip="Die Bezugs-Position gewählter Shapes auf linke untere Ecke ändern"
                ),
                bkt.ribbon.ToggleButton(
                    id="fix_loc_pins_1_1",
                    label="Bezugspos. oben/rechts",
                    show_label=False,
                    image="fix_loc_pins_1_1",
                    # size="large",
                    on_toggle_action=bkt.Callback(lambda shapes, pressed: LocPin.set_pins_state(shapes, 1, 1), shapes=True, shapes_min=1),
                    get_pressed=bkt.Callback(lambda shapes: LocPin.get_pins_state(shapes, 1, 1), shapes=True, shapes_min=1),
                    get_enabled = bkt.Callback(LocPin.get_enabled, shapes=True, shapes_min=1),
                    # screentip="Die Bezugs-Position gewählter Shapes auf linke untere Ecke ändern"
                ),
            ]
        ),
        bkt.ribbon.Box(
            box_style="horizontal",
            children =[
                bkt.ribbon.ToggleButton(
                    id="fix_loc_pins_0_05",
                    label="Bezugspos. mitte/links",
                    show_label=False,
                    image="fix_loc_pins_0_05",
                    # size="large",
                    on_toggle_action=bkt.Callback(lambda shapes, pressed: LocPin.set_pins_state(shapes, 0, 0.5), shapes=True, shapes_min=1),
                    get_pressed=bkt.Callback(lambda shapes: LocPin.get_pins_state(shapes, 0, 0.5), shapes=True, shapes_min=1),
                    get_enabled = bkt.Callback(LocPin.get_enabled, shapes=True, shapes_min=1),
                    # screentip="Die Bezugs-Position gewählter Shapes auf linke untere Ecke ändern"
                ),
                bkt.ribbon.ToggleButton(
                    id="fix_loc_pins_05_05",
                    label="Bezugspos. mitte/mitte",
                    show_label=False,
                    image="fix_loc_pins_05_05",
                    # size="large",
                    on_toggle_action=bkt.Callback(lambda shapes, pressed: LocPin.set_pins_state(shapes, 0.5, 0.5), shapes=True, shapes_min=1),
                    get_pressed=bkt.Callback(lambda shapes: LocPin.get_pins_state(shapes, 0.5, 0.5), shapes=True, shapes_min=1),
                    get_enabled = bkt.Callback(LocPin.get_enabled, shapes=True, shapes_min=1),
                    # screentip="Die Bezugs-Position gewählter Shapes auf linke untere Ecke ändern"
                ),
                bkt.ribbon.ToggleButton(
                    id="fix_loc_pins_1_05",
                    label="Bezugspos. mitte/rechts",
                    show_label=False,
                    image="fix_loc_pins_1_05",
                    # size="large",
                    on_toggle_action=bkt.Callback(lambda shapes, pressed: LocPin.set_pins_state(shapes, 1, 0.5), shapes=True, shapes_min=1),
                    get_pressed=bkt.Callback(lambda shapes: LocPin.get_pins_state(shapes, 1, 0.5), shapes=True, shapes_min=1),
                    get_enabled = bkt.Callback(LocPin.get_enabled, shapes=True, shapes_min=1),
                    # screentip="Die Bezugs-Position gewählter Shapes auf linke untere Ecke ändern"
                ),
            ]
        ),
        bkt.ribbon.Box(
            box_style="horizontal",
            children =[
                bkt.ribbon.ToggleButton(
                    id="fix_loc_pins_0_0",
                    label="Bezugspos. unten/links",
                    show_label=False,
                    image="fix_loc_pins_0_0",
                    # size="large",
                    on_toggle_action=bkt.Callback(lambda shapes, pressed: LocPin.set_pins_state(shapes, 0, 0), shapes=True, shapes_min=1),
                    get_pressed=bkt.Callback(lambda shapes: LocPin.get_pins_state(shapes, 0, 0), shapes=True, shapes_min=1),
                    get_enabled = bkt.Callback(LocPin.get_enabled, shapes=True, shapes_min=1),
                    # screentip="Die Bezugs-Position gewählter Shapes auf linke untere Ecke ändern"
                ),
                bkt.ribbon.ToggleButton(
                    id="fix_loc_pins_05_0",
                    label="Bezugspos. unten/mitte",
                    show_label=False,
                    image="fix_loc_pins_05_0",
                    # size="large",
                    on_toggle_action=bkt.Callback(lambda shapes, pressed: LocPin.set_pins_state(shapes, 0.5, 0), shapes=True, shapes_min=1),
                    get_pressed=bkt.Callback(lambda shapes: LocPin.get_pins_state(shapes, 0.5, 0), shapes=True, shapes_min=1),
                    get_enabled = bkt.Callback(LocPin.get_enabled, shapes=True, shapes_min=1),
                    # screentip="Die Bezugs-Position gewählter Shapes auf linke untere Ecke ändern"
                ),
                bkt.ribbon.ToggleButton(
                    id="fix_loc_pins_1_0",
                    label="Bezugspos. unten/rechts",
                    show_label=False,
                    image="fix_loc_pins_1_0",
                    # size="large",
                    on_toggle_action=bkt.Callback(lambda shapes, pressed: LocPin.set_pins_state(shapes, 1, 0), shapes=True, shapes_min=1),
                    get_pressed=bkt.Callback(lambda shapes: LocPin.get_pins_state(shapes, 1, 0), shapes=True, shapes_min=1),
                    get_enabled = bkt.Callback(LocPin.get_enabled, shapes=True, shapes_min=1),
                    # screentip="Die Bezugs-Position gewählter Shapes auf linke untere Ecke ändern"
                ),
            ]
        ),
        bkt.ribbon.DialogBoxLauncher(idMso='RulerAndGridDialog')
    ]
)

class Swap(object):
    @staticmethod
    def swap(shapes):
        s1, s2 = shapes
        s1._x, s2._x = s2._x, s1._x
        s1._y, s2._y = s2._y, s1._y
    
    @staticmethod
    def multi_swap(shapes):
        l,t = shapes[-1]._x, shapes[-1]._y
        for i in range(len(shapes)-2, -1, -1):
            shapes[i+1]._x, shapes[i+1]._y = shapes[i]._x, shapes[i]._y
        shapes[0]._x, shapes[0]._y = l, t
    
    @staticmethod
    def multi_swap_pos_size(shapes):
        l,t = shapes[-1]._x, shapes[-1]._y
        w,h = shapes[-1]._width, shapes[-1]._height
        for i in range(len(shapes)-2, -1, -1):
            shapes[i+1]._x, shapes[i+1]._y = shapes[i]._x, shapes[i]._y
            shapes[i+1]._width, shapes[i+1]._height = shapes[i]._width, shapes[i]._height
        shapes[0]._x, shapes[0]._y = l, t
        shapes[0]._width, shapes[0]._height = w, h


class EqualSize(object):

    @staticmethod
    def equal_height(shapes):
        master = shapes[0]
        apply_to = shapes[1:]
        for shape in apply_to:
            shape._height = master._height

    @staticmethod
    def equal_width(shapes):
        master = shapes[0]
        apply_to = shapes[1:]
        for shape in apply_to:
            shape._width = master._width


    @staticmethod
    def equal_height_func(shapes, func):
        heights = []
        for shape in shapes:
            heights.append(shape._height)
        
        sel_height = func(heights)
        for shape in shapes:
            shape._height = sel_height

    @staticmethod
    def equal_width_func(shapes, func):
        widths = []
        for shape in shapes:
            widths.append(shape._width)
        
        sel_width = func(widths)
        for shape in shapes:
            shape._width = sel_width


class ObjektAusrichtung(object):

    @staticmethod
    def align_left(selection, shapes):
        selection.Align(1,0,0)

    @staticmethod
    def align_right(selection, shapes):
        selection.Align(3,0,0)

    @staticmethod
    def align_top(selection, shapes):
        selection.Align(0,1,0)

    @staticmethod
    def align_bottom(selection, shapes):
        selection.Align(0,3,0)

    @staticmethod
    def vertical_centering(selection, shapes):
        selection.Align(0,2,0)

    @staticmethod
    def horizontal_centering(selection, shapes):
        selection.Align(2,0,0)


class ObjektVerteilung(object):

    @staticmethod
    def dist_horiz(selection, shapes):
        selection.Distribute(0,0)

    @staticmethod
    def dist_horiz_left_edge(selection, shapes):
        selection.Distribute(1,0)

    @staticmethod
    def dist_horiz_center(selection, shapes):
        selection.Distribute(2,0)

    @staticmethod
    def dist_horiz_right_edge(selection, shapes):
        selection.Distribute(3,0)

    @staticmethod
    def dist_vert(selection, shapes):
        selection.Distribute(4,0)

    @staticmethod
    def dist_vert_top_edge(selection, shapes):
        selection.Distribute(5,0)

    @staticmethod
    def dist_vert_center(selection, shapes):
        selection.Distribute(6,0)

    @staticmethod
    def dist_vert_bottom_edge(selection, shapes):
        selection.Distribute(7,0)


anordnen_gruppe = bkt.ribbon.Group(
    label='Anordnen',
    image_mso='ObjectsArrangeMenu',
    children = [
        bkt.ribbon.Box(box_style="vertical",
            children = [
                bkt.ribbon.SplitButton(
                    show_label=False,
                    children=[
                        bkt.ribbon.Button(
                            id = 'equal_height',
                            label="Gleiche Höhe",
                            show_label=False,
                            image_mso='ShapeHeight',
                            screentip="Höhe der Shapes einheitlich (Ausrichtung am Master-Shape)",
                            on_action=bkt.Callback(EqualSize.equal_height, shapes=True, shapes_min=2),
                            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name
                        ),
                        bkt.ribbon.Menu(children=[
                            bkt.ribbon.Button(
                                label="Gleiche Höhe wie Master-Shape",
                                image_mso='ShapeHeight',
                                screentip="Höhe der Shapes einheitlich (Ausrichtung am Master-Shape)",
                                on_action=bkt.Callback(EqualSize.equal_height, shapes=True, shapes_min=2),
                                get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name
                            ),
                            bkt.ribbon.MenuSeparator(),
                            bkt.ribbon.Button(
                                label="Gleiche Höhe wie höchstes Shape",
                                # image_mso='ShapeHeight',
                                screentip="Höhe der Shapes einheitlich (Ausrichtung am höchsten Shape)",
                                on_action=bkt.Callback(lambda shapes: EqualSize.equal_height_func(shapes, max), shapes=True, shapes_min=2),
                                get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name
                            ),
                            bkt.ribbon.Button(
                                label="Gleiche Höhe wie niedrigstes Shape",
                                # image_mso='ShapeHeight',
                                screentip="Höhe der Shapes einheitlich (Ausrichtung am niedrigsten Shape)",
                                on_action=bkt.Callback(lambda shapes: EqualSize.equal_height_func(shapes, min), shapes=True, shapes_min=2),
                                get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name
                            ),
                        ])
                    ]
                ),
                bkt.ribbon.SplitButton(
                    show_label=False,
                    children=[
                        bkt.ribbon.Button(
                            id = 'equal_width',
                            label="Gleiche Breite",
                            show_label=False,
                            image_mso='ShapeWidth',
                            screentip="Breite der Shapes einheitlich (Ausrichtung am Master-Shape)",
                            on_action=bkt.Callback(EqualSize.equal_width, shapes=True, shapes_min=2),
                            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name
                        ),
                        bkt.ribbon.Menu(children=[
                            bkt.ribbon.Button(
                                label="Gleiche Breite wie Master-Shape",
                                image_mso='ShapeWidth',
                                screentip="Breite der Shapes einheitlich (Ausrichtung am Master-Shape)",
                                on_action=bkt.Callback(EqualSize.equal_width, shapes=True, shapes_min=2),
                                get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name
                            ),
                            bkt.ribbon.MenuSeparator(),
                            bkt.ribbon.Button(
                                label="Gleiche Breite wie breitestes Shape",
                                # image_mso='ShapeWidth',
                                screentip="Breite der Shapes einheitlich (Ausrichtung am breitesten Shape)",
                                on_action=bkt.Callback(lambda shapes: EqualSize.equal_width_func(shapes, max), shapes=True, shapes_min=2),
                                get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name
                            ),
                            bkt.ribbon.Button(
                                label="Gleiche Breite wie schmalstes Shape",
                                # image_mso='ShapeWidth',
                                screentip="Breite der Shapes einheitlich (Ausrichtung am schmalsten Shape)",
                                on_action=bkt.Callback(lambda shapes: EqualSize.equal_width_func(shapes, min), shapes=True, shapes_min=2),
                                get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name
                            ),
                        ])
                    ]
                ),
                bkt.ribbon.SplitButton(
                    show_label=False,
                    children=[
                        bkt.ribbon.Button(
                            id = 'swap',
                            label="Tausche Position",
                            show_label=False,
                            image_mso='MailMergeMatchFields',
                            screentip="Shape-Position tauschen",
                            on_action=bkt.Callback(Swap.multi_swap, shapes=True, shapes_min=2),
                            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name
                        ),
                        bkt.ribbon.Menu(children=[
                            bkt.ribbon.Button(
                                label="Tausche Position",
                                image_mso='MailMergeMatchFields',
                                screentip="Shape-Position tauschen",
                                on_action=bkt.Callback(Swap.multi_swap, shapes=True, shapes_min=2),
                                get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name
                            ),
                            bkt.ribbon.MenuSeparator(),
                            bkt.ribbon.Button(
                                label="Tausche Position und Größe",
                                # image_mso='MailMergeMatchFields',
                                screentip="Shape-Position und -größe tauschen",
                                on_action=bkt.Callback(Swap.multi_swap_pos_size, shapes=True, shapes_min=2),
                                get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name
                            )
                        ])
                    ]
                ),
            ]
        ),
        bkt.ribbon.Box(box_style="vertical",
            children = [
                bkt.ribbon.Button(
                    id = 'align_left',
                    label="Links ausrichten",
                    show_label=False,
                    image_mso='ObjectsAlignLeft',
                    screentip="Shapes links ausrichten",
                    on_action=bkt.Callback(ObjektAusrichtung.align_left, selection=True, shapes=True, shapes_min=2),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name
                ),
                bkt.ribbon.Button(
                    id = 'align_right',
                    label="Rechts ausrichten",
                    show_label=False,
                    image_mso='ObjectsAlignRight',
                    screentip="Shapes rechts ausrichten",
                    on_action=bkt.Callback(ObjektAusrichtung.align_right, selection=True, shapes=True, shapes_min=2),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name
                ),
                bkt.ribbon.Button(
                    id = 'vertical_centering',
                    label="Vertikal mittig ausrichten",
                    show_label=False,
                    image_mso='ObjectsAlignMiddleVertical',
                    screentip="Shapes mittig ausrichten",
                    on_action=bkt.Callback(ObjektAusrichtung.vertical_centering, selection=True, shapes=True, shapes_min=2),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name
                ),
                bkt.ribbon.Button(
                    id = 'align_top',
                    label="Oben ausrichten",
                    show_label=False,
                    image_mso='ObjectsAlignTop',
                    screentip="Shapes oben ausrichten",
                    on_action=bkt.Callback(ObjektAusrichtung.align_top, selection=True, shapes=True, shapes_min=2),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name
                ),
                bkt.ribbon.Button(
                    id = 'align_bottom',
                    label="Unten ausrichten",
                    show_label=False,
                    image_mso='ObjectsAlignBottom',
                    screentip="Shapes unten ausrichten",
                    on_action=bkt.Callback(ObjektAusrichtung.align_bottom, selection=True, shapes=True, shapes_min=2),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name
                ),
                bkt.ribbon.Button(
                    id = 'horizontal_centering',
                    label="Horizontal zentriert ausrichten",
                    show_label=False,
                    image_mso='ObjectsAlignCenterHorizontal',
                    screentip="Shapes zentriert ausrichten",
                    on_action=bkt.Callback(ObjektAusrichtung.horizontal_centering, selection=True, shapes=True, shapes_min=2),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name
                ),
                bkt.ribbon.SplitButton(show_label=False, children=[
                    bkt.ribbon.Button(
                        id = 'dist_horiz',
                        label="Horizontal verteilen",
                        show_label=False,
                        image_mso='AlignDistributeHorizontally',
                        screentip="Shapes horizontal verteilen",
                        on_action=bkt.Callback(ObjektVerteilung.dist_horiz, selection=True, shapes=True, shapes_min=2),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name
                    ),
                    bkt.ribbon.Menu(children=[
                        bkt.ribbon.Button(
                            label="Horizontal verteilen",
                            image_mso='AlignDistributeHorizontally',
                            screentip="Shapes horizontal verteilen",
                            on_action=bkt.Callback(ObjektVerteilung.dist_horiz, selection=True, shapes=True, shapes_min=2),
                            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name
                        ),
                        bkt.ribbon.MenuSeparator(),
                        bkt.ribbon.Button(
                            label="Horizontal an linker Kante verteilen",
                            screentip="Shapes horizontal an linker Kante verteilen",
                            on_action=bkt.Callback(ObjektVerteilung.dist_horiz_left_edge, selection=True, shapes=True, shapes_min=2),
                            get_enabled=bkt.CallbackTypes.get_enabled.dotnet_name,
                        ),
                        bkt.ribbon.Button(
                            label="Horizontal an Mittelpunkt verteilen",
                            screentip="Shapes horizontal an Shape-Mittelpunkt verteilen",
                            on_action=bkt.Callback(ObjektVerteilung.dist_horiz_center, selection=True, shapes=True, shapes_min=2),
                            get_enabled=bkt.CallbackTypes.get_enabled.dotnet_name,
                        ),
                        bkt.ribbon.Button(
                            label="Horizontal an rechter Kante verteilen",
                            screentip="Shapes horizontal an rechter Kante verteilen",
                            on_action=bkt.Callback(ObjektVerteilung.dist_horiz_right_edge, selection=True, shapes=True, shapes_min=2),
                            get_enabled=bkt.CallbackTypes.get_enabled.dotnet_name,
                        ),
                    ])
                ]),
                bkt.ribbon.SplitButton(show_label=False, children=[
                    bkt.ribbon.Button(
                        id = 'dist_vert',
                        label="Vertikal verteilen",
                        show_label=False,
                        image_mso='AlignDistributeVertically',
                        screentip="Shapes vertikal verteilen",
                        on_action=bkt.Callback(ObjektVerteilung.dist_vert, selection=True, shapes=True, shapes_min=2),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name
                    ),
                    bkt.ribbon.Menu(children=[
                        bkt.ribbon.Button(
                            label="Vertikal verteilen",
                            image_mso='AlignDistributeVertically',
                            screentip="Shapes vertikal verteilen",
                            on_action=bkt.Callback(ObjektVerteilung.dist_vert, selection=True, shapes=True, shapes_min=2),
                            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name
                        ),
                        bkt.ribbon.MenuSeparator(),
                        bkt.ribbon.Button(
                            label="Vertikal an oberer Kante verteilen",
                            screentip="Shapes vertikal an oberer Kante verteilen",
                            on_action=bkt.Callback(ObjektVerteilung.dist_vert_top_edge, selection=True, shapes=True, shapes_min=2),
                            get_enabled=bkt.CallbackTypes.get_enabled.dotnet_name,
                        ),
                        bkt.ribbon.Button(
                            label="Vertikal an Mittelpunkt verteilen",
                            screentip="Shapes vertikal an Shape-Mittelpunkt verteilen",
                            on_action=bkt.Callback(ObjektVerteilung.dist_vert_center, selection=True, shapes=True, shapes_min=2),
                            get_enabled=bkt.CallbackTypes.get_enabled.dotnet_name,
                        ),
                        bkt.ribbon.Button(
                            label="Vertikal an unterer Kante verteilen",
                            screentip="Shapes vertikal an unterer Kante verteilen",
                            on_action=bkt.Callback(ObjektVerteilung.dist_vert_bottom_edge, selection=True, shapes=True, shapes_min=2),
                            get_enabled=bkt.CallbackTypes.get_enabled.dotnet_name,
                        ),
                    ])
                ]),
                bkt.mso.control.ObjectRotateGallery,
                bkt.mso.control.ObjectsGroup,
                bkt.mso.control.ObjectsUngroup,
                bkt.mso.control.ObjectsGroupMenu,
                # bkt.mso.control.AlignGallery,
                bkt.mso.control.ObjectBringToFrontMenu,
                bkt.mso.control.ObjectSendToBackMenu,
                bkt.mso.control.PositionMenu
            ]
        ),
        bkt.ribbon.DialogBoxLauncher(idMso='AlignDialog')
    ]
)