# -*- coding: utf-8 -*-
'''
Created on 06.07.2016

@author: rdebeerst
'''

import bkt

import bkt.library.algorithms as algos
import bkt.library.powerpoint as pplib
from bkt.library.powerpoint import pt_to_cm, cm_to_pt

import math
import os.path
from heapq import nsmallest, nlargest

from linkshapes import LinkedShapes
from shapes import PositionSize

class Swap(object):
    locpin = pplib.LocPin()

    # @staticmethod
    # def swap(shapes):
    #     s1, s2 = shapes
    #     s1.Left, s2.Left = s2.Left, s1.Left
    #     s1.Top, s2.Top = s2.Top, s1.Top

    # @staticmethod
    # def swap_left(shapes):
    #     s1, s2 = shapes
    #     s1.Left, s2.Left = s2.Left, s1.Left

    # @staticmethod
    # def swap_top(shapes):
    #     s1, s2 = shapes
    #     s1.Top, s2.Top = s2.Top, s1.Top
    
    @classmethod
    def multi_swap(cls, shapes):
        if bkt.library.system.get_key_state(bkt.library.system.key_code.SHIFT):
            # change position *and size*
            return Swap.multi_swap_pos_size(shapes)

        shapes = pplib.wrap_shapes(shapes, cls.locpin)
        l,t = shapes[-1].left, shapes[-1].top
        for i in range(len(shapes)-2, -1, -1):
            shapes[i+1].left, shapes[i+1].top = shapes[i].left, shapes[i].top
        shapes[0].left, shapes[0].top = l, t

    @classmethod
    def multi_swap_left(cls, shapes):
        l = shapes[-1].left
        for i in range(len(shapes)-2, -1, -1):
            shapes[i+1].left = shapes[i].left
        shapes[0].left = l

    @classmethod
    def multi_swap_top(cls, shapes):
        t = shapes[-1].top
        for i in range(len(shapes)-2, -1, -1):
            shapes[i+1].top = shapes[i].top
        shapes[0].top = t

    @classmethod
    def multi_swap_zorder(cls, shapes):
        all_zorders = [s.ZOrderPosition for s in shapes]
        for i in range(len(shapes)-2, -1, -1):
            pplib.set_shape_zorder(shapes[i+1], value=all_zorders[i])
        pplib.set_shape_zorder(shapes[0], value=all_zorders[-1])
    
    @classmethod
    def multi_swap_pos_size(cls, shapes):
        shapes = pplib.wrap_shapes(shapes, cls.locpin)

        l,t = shapes[-1].left, shapes[-1].top
        w,h = shapes[-1].width, shapes[-1].height
        for i in range(len(shapes)-2, -1, -1):
            shapes[i+1].left, shapes[i+1].top = shapes[i].left, shapes[i].top
            shapes[i+1].width, shapes[i+1].height = shapes[i].width, shapes[i].height
        shapes[0].left, shapes[0].top = l, t
        shapes[0].width, shapes[0].height = w, h


    # @classmethod
    # def swap_format(cls, shapes):
    #     s1, s2 = shapes
    #     stemp = s2.Duplicate()
    #     s1.PickUp()
    #     s2.Apply()
    #     stemp.PickUp()
    #     s1.Apply()
    #     stemp.Delete()

    @classmethod
    def multi_swap_format(cls, shapes):
        temp = shapes[-1].Duplicate()
        try:
            for i in range(len(shapes)-2, -1, -1):
                shapes[i].PickUp()
                shapes[i+1].Apply()
            temp.PickUp()
            shapes[0].Apply()
        except:
            # bkt.helpers.exception_as_message()
            pass
        temp.Delete()


    @staticmethod
    def replace_keep_size(shapes):
        new = shapes[0]
        ref = shapes[1]
        new.top = ref.top
        new.left = ref.left
        new.width = ref.width
        new.height = ref.height
        new.rotation = ref.rotation
        pplib.set_shape_zorder(new, value=ref.ZOrderPosition)
        ref.Delete()



swap_button = bkt.ribbon.SplitButton(
    show_label=False,
    get_enabled=bkt.apps.ppt_shapes_min2_selected,
    children=[
        bkt.ribbon.Button(
            id = 'swap',
            label="Tauschen",
            image_mso='MailMergeMatchFields',
            screentip="Shape-Position tauschen",
            supertip="Tausche die Position (left/top) der markierten Shapes.",
            on_action=bkt.Callback(Swap.multi_swap, shapes=True, shapes_min=2),
            # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
        ),
        bkt.ribbon.Menu(label="Tauschen", children=[
            bkt.ribbon.MenuSeparator(title="Gewählte Shapes tauschen"),
            bkt.ribbon.Button(
                id = 'swap2',
                label="Tausche Position",
                image_mso='MailMergeMatchFields',
                screentip="Shape-Position tauschen",
                supertip="Tausche die Position (left/top) der markierten Shapes.",
                on_action=bkt.Callback(Swap.multi_swap, shapes=True, shapes_min=2),
                # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            ),
            bkt.ribbon.Button(
                id = 'swap_pos_and_size',
                label="Tausche Position und Größe [Shift]",
                show_label=True,
                #image_mso='MailMergeMatchFields',
                screentip="Shape-Position tauschen",
                supertip="Tausche die Position (left/top) und Größe der markierten Shapes.",
                on_action=bkt.Callback(Swap.multi_swap_pos_size, shapes=True, shapes_min=2),
                # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            ),
            pplib.LocpinGallery(
                label="Ankerpunkt beim Tauschen",
                screentip="Ankerpunkt beim Tauschen festlegen",
                supertip="Legt den Punkt fest, der beim Tauschen der Shapes fixiert sein soll.",
                locpin=Swap.locpin,
            ),
            bkt.ribbon.MenuSeparator(),
            bkt.ribbon.Button(
                id = 'swap_left',
                label="Tausche x-Position",
                show_label=True,
                screentip="Tausche x-Position",
                supertip="Tausche die x-Position (Abstand vom linken Folienrand) der markierten Shapes.",
                on_action=bkt.Callback(Swap.multi_swap_left, shapes=True, shapes_min=2),
                # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            ),
            bkt.ribbon.Button(
                id = 'swap_top',
                label="Tausche y-Position",
                show_label=True,
                screentip="Tausche y-Position",
                supertip="Tausche die y-Position (Abstand vom oberen Folienrand) der markierten Shapes.",
                on_action=bkt.Callback(Swap.multi_swap_top, shapes=True, shapes_min=2),
                # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            ),
            bkt.ribbon.Button(
                id = 'swap_zorder',
                label="Tausche Z-Order",
                show_label=True,
                screentip="Tausche Z-Order-Position",
                supertip="Tausche die Z-Order-Position der markierten Shapes.",
                on_action=bkt.Callback(Swap.multi_swap_zorder, shapes=True, shapes_min=2),
                # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            ),
            bkt.ribbon.MenuSeparator(),
            bkt.ribbon.Button(
                id = 'swap_format',
                label="Tausche Formatierung",
                show_label=True,
                #image_mso='MailMergeMatchFields',
                screentip="Shape-Formatierung tauschen",
                supertip="Tausche die Formatierung (Farbe, Rahmen, Schrift, ...) der markierten Shapes.",
                on_action=bkt.Callback(Swap.multi_swap_format, shapes=True, shapes_min=2),
                # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            ),
            bkt.ribbon.Button(
                id = 'replace_keep_size',
                label="Ersetzen und Größe erhalten",
                show_label=True,
                #size='large',
                image='replace_keep_size',
                supertip="Zuletzt gewähltes Shape mit zuerst gewähltem Shape ersetzen und dabei die Größe vom Original-Shape erhalten.",
                on_action=bkt.Callback(Swap.replace_keep_size, shapes=True, shapes_min=2, shapes_max=2),
                get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            ),
        ])
    ]
)




class EqualSize(object):
    @classmethod
    def _get_func(cls):
        if bkt.library.system.get_key_state(bkt.library.system.key_code.SHIFT):
            return min
        elif bkt.library.system.get_key_state(bkt.library.system.key_code.CTRL):
            if ArrangeAdvanced.master == "FIRST":
                return lambda l: l.pop(0)
            else:
                return lambda l: l.pop()
        else:
            return max


    @classmethod
    def equal_height_master(cls, shapes):
        if ArrangeAdvanced.master == "FIRST":
            func = lambda l: l.pop(0)
        else:
            func = lambda l: l.pop()
        
        cls.equal_height(shapes, func)

    @classmethod
    def equal_height(cls, shapes, func=None):
        func = func or cls._get_func()

        heights = []
        for shape in shapes:
            if shape.rotation == 90 or shape.rotation == 270:
                heights.append(shape.width)
            else:
                heights.append(shape.height)

        sel_height = func(heights)
        for shape in shapes:
            if shape.rotation == 90 or shape.rotation == 270:
                shape.width = sel_height
            else:
                shape.height = sel_height


    @classmethod
    def equal_width_master(cls, shapes):
        if ArrangeAdvanced.master == "FIRST":
            func = lambda l: l.pop(0)
        else:
            func = lambda l: l.pop()
        
        cls.equal_width(shapes, func)

    @classmethod
    def equal_width(cls, shapes, func=None):
        func = func or cls._get_func()

        widths = []
        for shape in shapes:
            if shape.rotation == 90 or shape.rotation == 270:
                widths.append(shape.height)
            else:
                widths.append(shape.width)

        sel_width = func(widths)
        for shape in shapes:
            if shape.rotation == 90 or shape.rotation == 270:
                shape.height = sel_width
            else:
                shape.width = sel_width



equal_height_button = bkt.ribbon.SplitButton(
    show_label=False,
    get_enabled=bkt.apps.ppt_shapes_min2_selected,
    children=[
        bkt.ribbon.Button(
            id = 'equal_height',
            label="Gleiche Höhe",
            image_mso='SizeToTallest',
            screentip="Gleiche Höhe wie höchstes Shape",
            supertip="Normiere die Höhe der markierten Shapes, entsprechend der Höhe des höchsten Shapes.\nBei gedrückter Shift-Taste entsprechend des niedrigsten Shapes.\nBei gedrückter Strg-Taste entsprechend des zuletzt markierten Shapes.",
            on_action=bkt.Callback(EqualSize.equal_height, shapes=True, shapes_min=2),
            # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
        ),
        bkt.ribbon.Menu(label="Höhe angleichen", children=[
            bkt.ribbon.MenuSeparator(title="Ausrichten an Shape-Auswahl"),
            bkt.ribbon.Button(
                id = 'equal_height2',
                label="Gleiche Höhe wie höchstes Shape",
                image_mso='SizeToTallest',
                supertip="Normiere die Höhe der markierten Shapes, entsprechend der Höhe des höchsten Shapes.",
                on_action=bkt.Callback(lambda shapes: EqualSize.equal_height(shapes, max), shapes=True, shapes_min=2),
                # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            ),
            bkt.ribbon.MenuSeparator(),
            bkt.ribbon.Button(
                id = 'equal_height_min',
                label="Gleiche Höhe wie niedrigstes Shape [Shift]",
                show_label=True,
                image_mso='SizeToShortest',
                supertip="Normiere die Höhe der markierten Shapes, entsprechend der Höhe des niedrigsten Shapes.",
                on_action=bkt.Callback(lambda shapes: EqualSize.equal_height(shapes, min), shapes=True, shapes_min=2),
                # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            ),
            bkt.ribbon.Button(
                id = 'equal_height_median',
                label="Gleiche Höhe wie Median der Shape-Höhen",
                show_label=True,
                #image_mso='ShapeHeight',
                supertip="Normiere die Höhe der markierten Shapes, entsprechend des Medians der Höhen der Shapes.",
                on_action=bkt.Callback(lambda shapes: EqualSize.equal_height(shapes, bkt.library.algorithms.median), shapes=True, shapes_min=2),
                # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            ),
            bkt.ribbon.Button(
                id = 'equal_height_mean',
                label="Gleiche Höhe wie Mittelwert der Shape-Höhen",
                show_label=True,
                #image_mso='ShapeHeight',
                supertip="Normiere die Höhe der markierten Shapes, entsprechend des Mittelwers der Höhen der Shapes.",
                on_action=bkt.Callback(lambda shapes: EqualSize.equal_height(shapes, bkt.library.algorithms.mean), shapes=True, shapes_min=2),
                # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            ),
            bkt.ribbon.Button(
                id = 'equal_height_last_sel',
                label="Gleiche Höhe wie Master-Shape [Strg]",
                show_label=True,
                #image_mso='ShapeWidth',
                supertip="Normiere die Höhe der markierten Shapes, entsprechend der Höhe des Master-Shapes (also zuletzt bzw. zuerst markiertes Shape).",
                on_action=bkt.Callback(EqualSize.equal_height_master, shapes=True, shapes_min=2),
                # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            ),
        ])
    ]
)

equal_width_button = bkt.ribbon.SplitButton(
    show_label=False,
    get_enabled=bkt.apps.ppt_shapes_min2_selected,
    children=[
        bkt.ribbon.Button(
            id = 'equal_width',
            label="Gleiche Breite",
            image_mso='SizeToWidest',
            screentip="Gleiche Breite wie breitestes Shape",
            supertip="Normiere die Breite der markierten Shapes, entsprechend der Breite des breitesten Shapes.\nBei gedrückter Shift-Taste entsprechend des schmalsten Shapes.\nBei gedrückter Strg-Taste entsprechend des zuletzt markierten Shapes.",
            on_action=bkt.Callback(EqualSize.equal_width, shapes=True, shapes_min=2),
            # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
        ),
        bkt.ribbon.Menu(label="Breite angleichen", children=[
            bkt.ribbon.MenuSeparator(title="Ausrichten an Shape-Auswahl"),
            bkt.ribbon.Button(
                id = 'equal_width2',
                label="Gleiche Breite wie breitestes Shape",
                image_mso='SizeToWidest',
                supertip="Normiere die Breite der markierten Shapes, entsprechend der Breite des breitesten Shapes.",
                on_action=bkt.Callback(lambda shapes: EqualSize.equal_width(shapes, max), shapes=True, shapes_min=2),
                # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            ),
            bkt.ribbon.MenuSeparator(),
            bkt.ribbon.Button(
                id = 'equal_width_min',
                label="Gleiche Breite wie schmalstes Shape [Shift]",
                show_label=True,
                image_mso='SizeToNarrowest',
                supertip="Normiere die Breite der markierten Shapes, entsprechend der Breite des schmalsten Shapes.",
                on_action=bkt.Callback(lambda shapes: EqualSize.equal_width(shapes, min), shapes=True, shapes_min=2),
                # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            ),
            bkt.ribbon.Button(
                id = 'equal_width_median',
                label="Gleiche Breite wie Median der Shape-Breiten",
                show_label=True,
                #image_mso='ShapeWidth',
                supertip="Normiere die Breite der markierten Shapes, entsprechend des Medians der Breiten der Shapes.",
                on_action=bkt.Callback(lambda shapes: EqualSize.equal_width(shapes, bkt.library.algorithms.median), shapes=True, shapes_min=2),
                # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            ),
            bkt.ribbon.Button(
                id = 'equal_width_mean',
                label="Gleiche Breite wie Mittelwert der Shape-Breiten",
                show_label=True,
                #image_mso='ShapeWidth',
                supertip="Normiere die Breite der markierten Shapes, entsprechend des Mittelwerts der Breiten der Shapes.",
                on_action=bkt.Callback(lambda shapes: EqualSize.equal_width(shapes, bkt.library.algorithms.mean), shapes=True, shapes_min=2),
                # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            ),
            bkt.ribbon.Button(
                id = 'equal_width_last_sel',
                label="Gleiche Breite wie Master-Shape [Strg]",
                show_label=True,
                #image_mso='ShapeWidth',
                supertip="Normiere die Breite der markierten Shapes, entsprechend der Breite des Master-Shapes (also zuletzt bzw. zuerst markiertes Shape).",
                on_action=bkt.Callback(EqualSize.equal_width_master, shapes=True, shapes_min=2),
                # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            ),
        ])
    ]
)



class ShapeDistance(object):
    default_sep = 0.2
    vertical_edges   = bkt.settings.get("toolbox.shapedis.vertical_edges",   "distance") #other options: visual, top, center, bottom
    horizontal_edges = bkt.settings.get("toolbox.shapedis.horizontal_edges", "distance") #other options: visual, left, center, right
    vertical_fix   = bkt.settings.get("toolbox.shapedis.vertical_fix",   "top") #other options: bottom
    horizontal_fix = bkt.settings.get("toolbox.shapedis.horizontal_fix", "left") #other options: right

    ### only for euclid distance and angle:
    shape1_index  = 0 #center-shape-index is either 0 for first selected shape or -1 for last selected shape
    shape1_locpin = pplib.LocPin(4) #center point as initial locpin
    shape2_locpin = pplib.LocPin(4) #center point as initial locpin
    shape_rotate_with_angle = False #rotate shape if angle is changed
    euclid_multi_shape_mode = "centric" #Options: centric, delta, distribute

    @classmethod
    def change_settings(cls, name, value):
        # change setting and save value
        setattr(cls, name, value)
        bkt.settings["toolbox.shapedis."+name] = value

    @classmethod
    def _get_locpin(cls):
        locpin = pplib.LocPin()
        fix_v, fix_h = 1, 1 #visual and distance will use this locpin, but they will use x/x1/y/y1 properties anyway

        if cls.vertical_edges == "center":
            fix_v = 2
        elif cls.vertical_edges == "bottom":
            fix_v = 3
        
        if cls.horizontal_edges == "center":
            fix_h = 2
        elif cls.horizontal_edges == "right":
            fix_h = 3
        
        locpin.fixation = (fix_v, fix_h)
        return locpin

    @classmethod
    def get_enabled_min2_group(cls, shapes):
        if len(shapes) > 1:
            return True
        try:
            #test if shape is a group
            shapes[0].GroupItems
            return True
        except:
            return False
    
    @classmethod
    def get_wrapped_shapes(cls, shapes):
        # if only 1 shape is provided, it must be a group (secured by get_enabled)
        if len(shapes) == 1:
            try:
                shapes = list(iter(shapes[0].GroupItems))
            except:
                raise bkt.context.InappropriateContextError
        return pplib.wrap_shapes(shapes, cls._get_locpin())
    
    @classmethod
    def get_vertical_fix(cls):
        if bkt.library.system.get_key_state(bkt.library.system.key_code.ALT):
            if cls.vertical_fix == "bottom":
                return "top"
            else:
                return "bottom"
        else:
            return cls.vertical_fix
        
    @classmethod
    def set_shape_sep_vertical(cls, shapes, value):
        try:
            shapes = cls.get_wrapped_shapes(shapes)
        except:
            return

        vertical_fix = cls.get_vertical_fix()
        if vertical_fix=="bottom":
            shapes.sort(key=lambda shape: (shape.y1, shape.x1), reverse=True)
        else:
            shapes.sort(key=lambda shape: (shape.y, shape.x))

        if cls.vertical_edges == "visual" and vertical_fix=="bottom":
            cur_y = shapes[0].visual_y1
        elif cls.vertical_edges == "visual": #vertical_fix=top
            cur_y = shapes[0].visual_y
        
        elif cls.vertical_edges == "distance" and vertical_fix=="bottom":
            cur_y = shapes[0].y1
        elif cls.vertical_edges == "distance": #vertical_fix=top
            cur_y = shapes[0].y

        else:
            cur_y = shapes[0].top
        
        for shape in shapes:
            if cls.vertical_edges == "visual":
                if vertical_fix=="bottom":
                    shape.visual_y1 = cur_y
                else:
                    shape.visual_y = cur_y
                delta = shape.visual_height + value
            elif cls.vertical_edges == "distance":
                if vertical_fix=="bottom":
                    shape.y1 = cur_y
                else:
                    shape.y = cur_y
                delta = shape.height + value
            else:
                shape.top = cur_y
                delta = value
            if vertical_fix=="bottom":
                cur_y -= delta
            else:
                cur_y += delta

    @classmethod
    def get_shape_sep_vertical(cls, shapes):
        try:
            shapes = cls.get_wrapped_shapes(shapes)
        except:
            return
        
        if cls.vertical_fix=="bottom":
            # shapes.sort(key=lambda shape: (shape.y1, shape.x1))
            # shapes = shapes[-2:]
            shapes = nlargest(2, shapes, key=lambda shape: (shape.y1, shape.x1))
            shapes.reverse()
        else:
            # shapes.sort(key=lambda shape: (shape.y, shape.x))
            # shapes = shapes[:2]
            shapes = nsmallest(2, shapes, key=lambda shape: (shape.y, shape.x))
        
        if cls.vertical_edges == "distance":
            dis = shapes[1].y-shapes[0].y1
        elif cls.vertical_edges == "visual":
            dis = shapes[1].visual_y-shapes[0].visual_y-shapes[0].visual_height
        else:
            dis = shapes[1].top-shapes[0].top
        
        return dis

    @classmethod
    def get_horizontal_fix(cls):
        if bkt.library.system.get_key_state(bkt.library.system.key_code.ALT):
            if cls.horizontal_fix == "right":
                return "left"
            else:
                return "right"
        else:
            return cls.horizontal_fix

    @classmethod
    def set_shape_sep_horizontal(cls, shapes, value):
        try:
            shapes = cls.get_wrapped_shapes(shapes)
        except:
            return

        horizontal_fix = cls.get_horizontal_fix()
        if horizontal_fix=="right":
            shapes.sort(key=lambda shape: (shape.x1, shape.y1), reverse=True)
        else:
            shapes.sort(key=lambda shape: (shape.x, shape.y))

        if cls.horizontal_edges == "visual" and horizontal_fix=="right":
            cur_x = shapes[0].visual_x1
        elif cls.horizontal_edges == "visual": #horizontal_fix=left
            cur_x = shapes[0].visual_x
        
        elif cls.horizontal_edges == "distance" and horizontal_fix=="right":
            cur_x = shapes[0].x1
        elif cls.horizontal_edges == "distance": #horizontal_fix=left
            cur_x = shapes[0].x

        else:
            cur_x = shapes[0].left
        
        for shape in shapes:
            if cls.horizontal_edges == "visual":
                if horizontal_fix=="right":
                    shape.visual_x1 = cur_x
                else:
                    shape.visual_x = cur_x
                delta = shape.visual_width + value
            elif cls.horizontal_edges == "distance":
                if horizontal_fix=="right":
                    shape.x1 = cur_x
                else:
                    shape.x = cur_x
                delta = shape.width + value
            else:
                shape.left = cur_x
                delta = value
            if horizontal_fix=="right":
                cur_x -= delta
            else:
                cur_x += delta

    @classmethod
    def get_shape_sep_horizontal(cls, shapes):
        try:
            shapes = cls.get_wrapped_shapes(shapes)
        except:
            return
        
        if cls.horizontal_fix=="right":
            # shapes.sort(key=lambda shape: (shape.x1, shape.y1))
            # shapes = shapes[-2:]
            shapes = nlargest(2, shapes, key=lambda shape: (shape.x1, shape.y1))
            shapes.reverse()
        else:
            # shapes.sort(key=lambda shape: (shape.x, shape.y))
            # shapes = shapes[:2]
            shapes = nsmallest(2, shapes, key=lambda shape: (shape.x, shape.y))

        if cls.horizontal_edges == "distance":
            dis = shapes[1].x-shapes[0].x1
        elif cls.horizontal_edges == "visual":
            dis = shapes[1].visual_x-shapes[0].visual_x-shapes[0].visual_width
        else:
            dis = shapes[1].left-shapes[0].left
        
        return dis


    @classmethod
    def set_shape_sep_vertical_zero(cls, shapes):
        if bkt.library.system.get_key_state(bkt.library.system.key_code.SHIFT):
            cls.set_shape_sep_vertical(shapes, cls.default_sep)
        else:
            cls.set_shape_sep_vertical(shapes, 0)

    @classmethod
    def set_shape_sep_horizontal_zero(cls, shapes):
        if bkt.library.system.get_key_state(bkt.library.system.key_code.SHIFT):
            cls.set_shape_sep_horizontal(shapes, cls.default_sep)
        else:
            cls.set_shape_sep_horizontal(shapes, 0)


    ### Euclidian distance and angle methods ###

    @classmethod
    def is_mode_centric(cls):
        return cls.euclid_multi_shape_mode == "centric"

    @classmethod
    def is_mode_delta(cls):
        alt = bkt.library.system.get_key_state(bkt.library.system.key_code.ALT)
        return alt or cls.euclid_multi_shape_mode == "delta"

    @classmethod
    def is_mode_distribute(cls):
        return cls.euclid_multi_shape_mode == "distribute"

    @classmethod
    def get_shape_sep_euclid(cls, shapes):
        '''
        get euclidian distance from center shape to second shape
        '''
        shape1 = pplib.wrap_shape(shapes[cls.shape1_index], cls.shape1_locpin)
        shape2 = pplib.wrap_shape(shapes[cls.shape1_index+1], cls.shape2_locpin)
        shape1_x, shape1_y = shape1.left, shape1.top
        shape2_x, shape2_y = shape2.left, shape2.top

        # return math.sqrt( (shape2_y-shape1_y)**2 + (shape2_x-shape1_x)**2 )
        return math.hypot(shape2_x-shape1_x, shape2_y-shape1_y)
    
    @classmethod
    def set_shape_sep_euclid(cls, shapes, value):
        '''
        set euclidian distance from center shape to all other shapes
        '''
        def _get_current_distance(shape1, shape2):
            vector = [shape2.left-shape1.left, shape2.top-shape1.top]
            return math.hypot( *vector )

        def _get_new_shape_coords(shape1, shape2, delta_distance):
            vector = [shape2.left-shape1.left, shape2.top-shape1.top]
            cur_dis = math.hypot( *vector )
            if cur_dis == 0:
                raise ValueError("current distance is 0")
            uni_vector = [vector[0]/cur_dis, vector[1]/cur_dis]
            distance = cur_dis + delta_distance
            new_vector = [uni_vector[0]*distance, uni_vector[1]*distance]
            shape_rotation = (360-round(-180/math.pi * math.atan2(new_vector[1], new_vector[0]), 1)) % 360
            return new_vector, shape_rotation


        shape1 = pplib.wrap_shape(shapes[cls.shape1_index], cls.shape1_locpin)
        shape1_x, shape1_y = shape1.left, shape1.top

        shapes = pplib.wrap_shapes(shapes[cls.shape1_index+1:], cls.shape2_locpin)
        # shape2_x, shape2_y = shapes[0].left, shapes[0].top

        # alt = bkt.library.system.get_key_state(bkt.library.system.key_code.ALT)

        # if cls.is_mode_centric() or cls.is_mode_distribute():
        if not cls.is_mode_delta():
            for i, shape in enumerate(shapes):
                try:
                    if cls.is_mode_centric():
                        delta_distance = value-_get_current_distance(shape1, shape)
                    else: #is_mode_distribute
                        delta_distance = (i+1)*value-_get_current_distance(shape1, shape)
                    new_vector, shape_rotation = _get_new_shape_coords(shape1, shape, delta_distance)
                except ValueError:
                    continue
                shape.left = shape1_x+new_vector[0]
                shape.top  = shape1_y+new_vector[1]
                # set shape rotation (without using wrapper function)
                if cls.shape_rotate_with_angle:
                    shape.shape.rotation = shape_rotation
        
        else: #is_mode_delta
            delta_distance = value-_get_current_distance(shape1, shapes[0])
            for shape in shapes:
                try:
                    new_vector, shape_rotation = _get_new_shape_coords(shape1, shape, delta_distance)
                except ValueError:
                    continue
                shape.left = shape1_x+new_vector[0]
                shape.top  = shape1_y+new_vector[1]
                # set shape rotation (without using wrapper function)
                if cls.shape_rotate_with_angle:
                    shape.shape.rotation = shape_rotation

    @classmethod
    def get_shape_angle(cls, shapes):
        '''
        get euclidian angle from center shape to second shaope
        '''
        shape1 = pplib.wrap_shape(shapes[cls.shape1_index], cls.shape1_locpin)
        shape2 = pplib.wrap_shape(shapes[cls.shape1_index+1], cls.shape2_locpin)
        shape1_x, shape1_y = shape1.left, shape1.top
        shape2_x, shape2_y = shape2.left, shape2.top

        vector = [shape2_x-shape1_x, shape2_y-shape1_y]
        return round(-180/math.pi * math.atan2(vector[1], vector[0]),1)

    @classmethod
    def set_shape_angle(cls, shapes, value):
        '''
        set euclidian angle from center shape to all other shapes
        '''
        def _get_current_angle(shape1, shape2):
            vector = [shape2.left-shape1.left, shape2.top-shape1.top]
            return round(-180/math.pi * math.atan2(vector[1], vector[0]), 1)

        def _get_new_shape_coords(shape1, shape2, delta_angle):
            vector = [shape2.left-shape1.left, shape2.top-shape1.top]
            new_vector = algos.rotate_point(vector[0], vector[1], 0, 0, delta_angle)
            shape_rotation = (360-round(-180/math.pi * math.atan2(new_vector[1], new_vector[0]), 1)) % 360
            return new_vector, shape_rotation


        shape1 = pplib.wrap_shape(shapes[cls.shape1_index], cls.shape1_locpin)
        shape1_x, shape1_y = shape1.left, shape1.top

        shapes = pplib.wrap_shapes(shapes[cls.shape1_index+1:], cls.shape2_locpin)
        # shape_rotation = (360-value) % 360

        # alt = bkt.library.system.get_key_state(bkt.library.system.key_code.ALT)

        # if cls.is_mode_centric() or cls.is_mode_distribute():
        if not cls.is_mode_delta():
            for i, shape in enumerate(shapes):
                if cls.is_mode_centric():
                    delta_angle = -(_get_current_angle(shape1, shape) - value)
                else: #is_mode_distribute
                    delta_angle = -(_get_current_angle(shape1, shape) - value*(i+1))
                new_vector, shape_rotation = _get_new_shape_coords(shape1, shape, delta_angle)
                shape2_x, shape2_y = shape.left, shape.top

                # vector = [shape2_x-shape1_x, shape2_y-shape1_y]
                # cur_angle = round(-180/math.pi * math.atan2(vector[1], vector[0]), 1)
                # new_vector = algos.rotate_point(vector[0], vector[1], 0, 0, -(cur_angle-value))

                shape.left = shape1_x+new_vector[0]
                shape.top  = shape1_y+new_vector[1]
                # set shape rotation (without using wrapper function)
                if cls.shape_rotate_with_angle:
                    shape.shape.rotation = shape_rotation
        
        else: #is_mode_delta
            delta_angle = _get_current_angle(shape1, shapes[0])-value
            for shape in shapes:
                new_vector, shape_rotation = _get_new_shape_coords(shape1, shape, -delta_angle)

                shape.left = shape1_x+new_vector[0]
                shape.top  = shape1_y+new_vector[1]
                # set shape rotation (without using wrapper function)
                if cls.shape_rotate_with_angle:
                    shape.shape.rotation = shape_rotation





class ShapeRotation(object):
    locpin = pplib.LocPin(4, "toolbox.rotation_locpin") #center point as initial locpin

    @staticmethod
    def get_enabled(selection):
        try:
            return selection.Type in [2,3] and not selection.ShapeRange.HasTable in [-2, -1]
            #tables do not support rotation
        except:
            return False

    @classmethod
    def set_rotation(cls, shapes, value):
        shapes = pplib.wrap_shapes(shapes, cls.locpin)
        bkt.apply_delta_on_ALT_key(
            lambda shape, value: setattr(shape, 'rotation', value), 
            lambda shape: shape.rotation, 
            shapes, value)

    @classmethod
    def get_rotation(cls, shapes):
        shape = pplib.wrap_shape(shapes[0], cls.locpin)
        return shape.rotation

    @classmethod
    def set_rotation_zero(cls, shapes):
        if bkt.library.system.get_key_state(bkt.library.system.key_code.SHIFT):
            cls.set_rotation(shapes, 180)
        else:
            cls.set_rotation(shapes, 0)

    @staticmethod
    def get_pressed_flipv(shapes):
        return shapes[0].VerticalFlip == -1

    @staticmethod
    def set_flipv(pressed, shapes):
        pressed = -1 if pressed else 0
        for shape in shapes:
            if shape.VerticalFlip != pressed:
                shape.Flip(1) #msoFlipVertical

    @staticmethod
    def get_pressed_fliph(shapes):
        return shapes[0].HorizontalFlip == -1

    @staticmethod
    def set_fliph(pressed, shapes):
        pressed = -1 if pressed else 0
        for shape in shapes:
            if shape.HorizontalFlip != pressed:
                shape.Flip(0) #msoFlipHorizontal




distance_rotation_group = bkt.ribbon.Group(
    id="bkt_shapesep_group",
    label="Abstand/Rotation",
    image_mso='VerticalSpacingIncrease',
    children=[
        bkt.ribbon.RoundingSpinnerBox(
            id = 'shape_sep_v',
            image_mso='VerticalSpacingIncrease',
            # image_button=True,
            label=u"Objektabstand vertikal",
            show_label=False,
            screentip="Vertikalen Objektabstand",
            supertip="Ändere den vertikalen Objektabstand auf das angegebene Maß (in cm).",
            on_change = bkt.Callback(ShapeDistance.set_shape_sep_vertical, shapes=True),
            get_text  = bkt.Callback(ShapeDistance.get_shape_sep_vertical, shapes=True),
            # on_action = bkt.Callback(ShapeDistance.set_shape_sep_vertical_zero, shapes=True),
            get_enabled = bkt.Callback(ShapeDistance.get_enabled_min2_group, shapes=True),
            # get_enabled = bkt.apps.ppt_shapes_min2_selected,
            round_cm = True,
            convert="pt_to_cm",
            image_element=bkt.ribbon.SplitButton(show_label=False, children=[
                bkt.ribbon.Button(
                    on_action = bkt.Callback(ShapeDistance.set_shape_sep_vertical_zero, shapes=True),
                ),
                bkt.ribbon.Menu(label="Vertikaler Abstand Einstellungen", item_size="large", children=[
                    bkt.ribbon.MenuSeparator(title="Vertikaler Abstand"),
                    bkt.ribbon.ToggleButton(
                        label="Shape-Abstand (Standard)",
                        image="shapedis_vdistance",
                        description="Abstand von unterer Kante zu oberer Kante anzeigen.",
                        get_pressed=bkt.Callback(lambda: ShapeDistance.vertical_edges=="distance"),
                        on_toggle_action=bkt.Callback(lambda pressed: ShapeDistance.change_settings("vertical_edges", "distance")),
                    ),
                    bkt.ribbon.ToggleButton(
                        label="Visueller Abstand (bei Rotation)",
                        image="shapedis_vvisual",
                        description="Abstand von visueller unterer Kante zu visueller oberer Kante anzeigen. Hilfreich bei rotierten Shapes.",
                        get_pressed=bkt.Callback(lambda: ShapeDistance.vertical_edges=="visual"),
                        on_toggle_action=bkt.Callback(lambda pressed: ShapeDistance.change_settings("vertical_edges", "visual")),
                    ),
                    bkt.ribbon.MenuSeparator(),
                    bkt.ribbon.ToggleButton(
                        label="Abstand zu oberen Kanten",
                        image="shapedis_top",
                        description="Abstand von oberer Kante zu oberer Kante anzeigen.",
                        get_pressed=bkt.Callback(lambda: ShapeDistance.vertical_edges=="top"),
                        on_toggle_action=bkt.Callback(lambda pressed: ShapeDistance.change_settings("vertical_edges", "top")),
                    ),
                    bkt.ribbon.ToggleButton(
                        label="Abstand zu Mittelpunkten",
                        image="shapedis_vcenter",
                        description="Abstand von den jeweiligen Mittelpunkten anzeigen.",
                        get_pressed=bkt.Callback(lambda: ShapeDistance.vertical_edges=="center"),
                        on_toggle_action=bkt.Callback(lambda pressed: ShapeDistance.change_settings("vertical_edges", "center")),
                    ),
                    bkt.ribbon.ToggleButton(
                        label="Abstand zu unteren Kanten",
                        image="shapedis_bottom",
                        description="Abstand von unterer Kante zu unterer Kante anzeigen.",
                        get_pressed=bkt.Callback(lambda: ShapeDistance.vertical_edges=="bottom"),
                        on_toggle_action=bkt.Callback(lambda pressed: ShapeDistance.change_settings("vertical_edges", "bottom")),
                    ),
                    bkt.ribbon.MenuSeparator(title="Bewegungsrichtung"),
                    bkt.ribbon.ToggleButton(
                        label="Nach unten",
                        image="shapedis_movedown",
                        description="Ändert die Distanz ausgehend vom obersten Shape und schiebt alle anderen nach unten. Temporärer Wechsel auf 'Nach oben' durch [ALT].",
                        get_pressed=bkt.Callback(lambda: ShapeDistance.vertical_fix=="top"),
                        on_toggle_action=bkt.Callback(lambda pressed: ShapeDistance.change_settings("vertical_fix", "top")),
                    ),
                    bkt.ribbon.ToggleButton(
                        label="Nach oben",
                        image="shapedis_moveup",
                        description="Ändert die Distanz ausgehend vom untersten Shape und schiebt alle anderen nach oben. Temporärer Wechsel auf 'Nach unten' durch [ALT].",
                        get_pressed=bkt.Callback(lambda: ShapeDistance.vertical_fix=="bottom"),
                        on_toggle_action=bkt.Callback(lambda pressed: ShapeDistance.change_settings("vertical_fix", "bottom")),
                    ),
                ])
            ]),
        ),

        bkt.ribbon.RoundingSpinnerBox(
            id = 'shape_sep_h',
            image_mso='HorizontalSpacingIncrease',
            # image_button=True,
            label=u"Objektabstand horizontal",
            show_label=False,
            screentip="Horizontalen Objektabstand",
            supertip="Ändere den horizontalen Objektabstand auf das angegebene Maß (in cm).",
            on_change = bkt.Callback(ShapeDistance.set_shape_sep_horizontal, shapes=True),
            get_text  = bkt.Callback(ShapeDistance.get_shape_sep_horizontal, shapes=True),
            # on_action = bkt.Callback(ShapeDistance.set_shape_sep_horizontal_zero, shapes=True),
            # get_enabled = bkt.apps.ppt_shapes_min2_selected,
            get_enabled = bkt.Callback(ShapeDistance.get_enabled_min2_group, shapes=True),
            round_cm = True,
            convert="pt_to_cm",
            image_element=bkt.ribbon.SplitButton(show_label=False, children=[
                bkt.ribbon.Button(
                    on_action = bkt.Callback(ShapeDistance.set_shape_sep_horizontal_zero, shapes=True),
                ),
                bkt.ribbon.Menu(label="Horizontaler Abstand Einstellungen", item_size="large", children=[
                    bkt.ribbon.MenuSeparator(title="Horizontaler Abstand"),
                    bkt.ribbon.ToggleButton(
                        label="Shape-Abstand (Standard)",
                        image="shapedis_hdistance",
                        description="Abstand von rechter Kante zu linker Kante anzeigen.",
                        get_pressed=bkt.Callback(lambda: ShapeDistance.horizontal_edges=="distance"),
                        on_toggle_action=bkt.Callback(lambda pressed: ShapeDistance.change_settings("horizontal_edges", "distance")),
                    ),
                    bkt.ribbon.ToggleButton(
                        label="Visueller Abstand (bei Rotation)",
                        image="shapedis_hvisual",
                        description="Abstand von visueller rechter Kante zu visueller linker Kante. Hilfreich bei rotierten Shapes.",
                        get_pressed=bkt.Callback(lambda: ShapeDistance.horizontal_edges=="visual"),
                        on_toggle_action=bkt.Callback(lambda pressed: ShapeDistance.change_settings("horizontal_edges", "visual")),
                    ),
                    bkt.ribbon.MenuSeparator(),
                    bkt.ribbon.ToggleButton(
                        label="Abstand zu linken Kanten",
                        image="shapedis_left",
                        description="Abstand von linker Kante zu linker Kante anzeigen.",
                        get_pressed=bkt.Callback(lambda: ShapeDistance.horizontal_edges=="left"),
                        on_toggle_action=bkt.Callback(lambda pressed: ShapeDistance.change_settings("horizontal_edges", "left")),
                    ),
                    bkt.ribbon.ToggleButton(
                        label="Abstand zu Mittelpunkten",
                        image="shapedis_hcenter",
                        description="Abstand von den jeweiligen Mittelpunkten anzeigen.",
                        get_pressed=bkt.Callback(lambda: ShapeDistance.horizontal_edges=="center"),
                        on_toggle_action=bkt.Callback(lambda pressed: ShapeDistance.change_settings("horizontal_edges", "center")),
                    ),
                    bkt.ribbon.ToggleButton(
                        label="Abstand zu rechten Kanten",
                        image="shapedis_right",
                        description="Abstand von rechter Kante zu rechter Kante anzeigen.",
                        get_pressed=bkt.Callback(lambda: ShapeDistance.horizontal_edges=="right"),
                        on_toggle_action=bkt.Callback(lambda pressed: ShapeDistance.change_settings("horizontal_edges", "right")),
                    ),
                    bkt.ribbon.MenuSeparator(title="Bewegungsrichtung"),
                    bkt.ribbon.ToggleButton(
                        label="Nach rechts",
                        image="shapedis_moveright",
                        description="Ändert die Distanz ausgehend vom linken Shape und schiebt alle anderen nach rechts. Temporärer Wechsel auf 'Nach links' durch [ALT].",
                        get_pressed=bkt.Callback(lambda: ShapeDistance.horizontal_fix=="left"),
                        on_toggle_action=bkt.Callback(lambda pressed: ShapeDistance.change_settings("horizontal_fix", "left")),
                    ),
                    bkt.ribbon.ToggleButton(
                        label="Nach links",
                        image="shapedis_moveleft",
                        description="Ändert die Distanz ausgehend vom rechten Shape und schiebt alle anderen nach links. Temporärer Wechsel auf 'Nach rechts' durch [ALT].",
                        get_pressed=bkt.Callback(lambda: ShapeDistance.horizontal_fix=="right"),
                        on_toggle_action=bkt.Callback(lambda pressed: ShapeDistance.change_settings("horizontal_fix", "right")),
                    ),
                ])
            ]),
        ),

        bkt.ribbon.RoundingSpinnerBox(
            id = 'rotation',
            label=u"Rotation",
            show_label=False,
            image_mso='Repeat',
            # image_button=True,
            screentip="Form-Rotation",
            supertip="Ändere die Rotation des Shapes auf das angegebene Maß (in Grad).",
            on_change = bkt.Callback(ShapeRotation.set_rotation, shapes=True),
            get_text  = bkt.Callback(ShapeRotation.get_rotation, shapes=True),
            # on_action = bkt.Callback(ShapeRotation.set_rotation_zero, shapes=True),
            get_enabled = bkt.Callback(ShapeRotation.get_enabled, selection=True),
            # get_enabled = bkt.apps.ppt_shapes_or_text_selected,
            round_int = True,
            huge_step = 45,
            image_element=bkt.ribbon.SplitButton(show_label=False, children=[
                bkt.ribbon.Button(
                    on_action = bkt.Callback(ShapeRotation.set_rotation_zero, shapes=True),
                ),
                bkt.ribbon.Menu(label="Rotation und Spiegelung Menü", children=[
                    bkt.ribbon.MenuSeparator(title="Rotation"),
                    pplib.LocpinGallery(
                        label="Ankerpunkt beim Rotieren",
                        screentip="Ankerpunkt beim Rotieren festlegen",
                        supertip="Legt den Punkt fest, der beim Rotieren der Shapes fixiert sein soll.",
                        locpin=ShapeRotation.locpin,
                    ),
                    bkt.ribbon.Button(
                        label="Auf 0 Grad setzen",
                        screentip="Shape-Rotation auf 0 Grad setzen",
                        on_action=bkt.Callback(lambda shapes: ShapeRotation.set_rotation(shapes, 0), shapes=True),
                    ),
                    bkt.ribbon.Button(
                        label="Auf 180 Grad setzen",
                        screentip="Shape-Rotation auf 180 Grad setzen",
                        on_action=bkt.Callback(lambda shapes: ShapeRotation.set_rotation(shapes, 180), shapes=True),
                    ),
                    # bkt.mso.button.ObjectRotateRight90,
                    # bkt.mso.button.ObjectRotateLeft90,
                    bkt.ribbon.MenuSeparator(title="Spieglung"),
                    bkt.ribbon.ToggleButton(
                        label="Horizontal gespiegelt an/aus",
                        screentip="Shape ist horizontal gespiegelt",
                        image_mso="ObjectFlipHorizontal",
                        get_pressed=bkt.Callback(ShapeRotation.get_pressed_fliph, shapes=True),
                        on_toggle_action=bkt.Callback(ShapeRotation.set_fliph, shapes=True),
                    ),
                    bkt.ribbon.ToggleButton(
                        label="Vertikal gespiegelt an/aus",
                        screentip="Shape ist vertikal gespiegelt",
                        image_mso="ObjectFlipVertical",
                        get_pressed=bkt.Callback(ShapeRotation.get_pressed_flipv, shapes=True),
                        on_toggle_action=bkt.Callback(ShapeRotation.set_flipv, shapes=True),
                    ),
                    # bkt.ribbon.MenuSeparator(),
                    # bkt.mso.button.ObjectFlipHorizontal,
                    # bkt.mso.button.ObjectFlipVertical,
                    bkt.ribbon.MenuSeparator(title="Optionen"),
                    bkt.mso.button.ObjectRotationOptionsDialog,
                ])
            ]),
        ),
    ]
)


euclid_angle_group = bkt.ribbon.Group(
    id="bkt_shapesep_advanced",
    label="Erw. Pos.",
    image_mso='VerticalSpacingIncrease',
    children=[
        bkt.ribbon.RoundingSpinnerBox(
            id = 'shape_sep_euclid',
            image='shape_euclid_distance',
            #image_button=True,
            label=u"Objektabstand euklidisch",
            show_label=False,
            screentip="Euklidischer Objektabstand",
            supertip="Ändert den euklidischen Objektabstand auf das angegebene Maß (in cm). Gemessen wird jeweils vom definierten Referenzpunkt der beiden zuerst selektierten Shapes.\n\nMit ALT-Taste wird jedes Shapes im gleichen Delta-Abstand zum ersten Shape ausgerichtet.",
            on_change = bkt.Callback(ShapeDistance.set_shape_sep_euclid, shapes=True, shapes_min=2),
            get_text  = bkt.Callback(ShapeDistance.get_shape_sep_euclid, shapes=True, shapes_min=2),
            get_enabled = bkt.apps.ppt_shapes_min2_selected,
            round_cm = True,
            convert="pt_to_cm",
        ),
        bkt.ribbon.RoundingSpinnerBox(
            id = 'shape_angle',
            image='shape_angle',
            #image_button=True,
            label=u"Winkel",
            show_label=False,
            screentip="Winkel der Shapes zueinander",
            supertip="Ändert den Winkel der Shapes zur X-Achse (in Grad). Gemessen wird jeweils vom definierten Referenzpunkt der beiden zuerst selektierten Shapes.\n\nMit ALT-Taste wird jedes Shapes um den gleichen Delta-Winkel zum ersten Shape verschoben.",
            on_change = bkt.Callback(ShapeDistance.set_shape_angle, shapes=True, shapes_min=2),
            get_text  = bkt.Callback(ShapeDistance.get_shape_angle, shapes=True, shapes_min=2),
            get_enabled = bkt.apps.ppt_shapes_min2_selected,
            round_int = True,
            huge_step = 45,
        ),
        bkt.ribbon.Menu(
            label="Ref.pkt.",
            image="fix_locpin_mm",
            screentip="Referenzpunkte festlegen",
            supertip="Referenz- bzw. Ankerpunkte für Distanz und Winkel für erstes Shape und die weitere Shapes festlegen",
            children=[
                bkt.ribbon.MenuSeparator(title="Zentrum-Shape festlegen"),
                bkt.ribbon.MenuSeparator(title="Referenzpunkte festlegen"),
                pplib.LocpinGallery(
                    label="Innerhalb Zentrum-Shape",
                    screentip="Referenz- bzw. Ankerpunkt für Distanz und Winkel",
                    supertip="Legt den Punkt innerhalb des Zentrum-Shapes fest, von dem der Abstand bzw. Winkel der Shapes gemessen werden soll.",
                    locpin=ShapeDistance.shape1_locpin,
                ),
                pplib.LocpinGallery(
                    label="Innerhalb weiterer Shapes",
                    screentip="Referenz- bzw. Ankerpunkt für Distanz und Winkel",
                    supertip="Legt den Punkt innerhalb aller weiteren Shapes fest, von dem der Abstand bzw. Winkel der Shapes gemessen werden soll.",
                    locpin=ShapeDistance.shape2_locpin,
                ),
                bkt.ribbon.MenuSeparator(title="Erweiterte Einstellungen festlegen"),
                bkt.ribbon.Menu(
                    label="Zentrum-Shape auswählen",
                    screentip="Shape im Zentrum auswählen",
                    supertip="Legt fest welches Shape innerhalb der Selektion im Zentrum stehen sollen.",
                    children=[
                        bkt.ribbon.ToggleButton(
                            label="Zuerst selektiertes Shape (Standard)",
                            get_pressed=bkt.Callback(lambda: ShapeDistance.shape1_index==0),
                            on_toggle_action=bkt.Callback(lambda pressed: setattr(ShapeDistance, "shape1_index", 0)),
                        ),
                        bkt.ribbon.ToggleButton(
                            label="Zuletzt selektiertes Shape",
                            get_pressed=bkt.Callback(lambda: ShapeDistance.shape1_index==-1),
                            on_toggle_action=bkt.Callback(lambda pressed: setattr(ShapeDistance, "shape1_index", -1)),
                        ),
                    ]
                ),
                bkt.ribbon.Menu(
                    label="Verhalten für >2 Shapes",
                    screentip="Verhalten bei Auswahl von mehr als 2 Shapes festlegen",
                    supertip="Bei Auswahl von mehr als 2 Shapes können unterschiedliche Optionen zur Berechnung zwischen Zentrum-Shape und allen weitere Shapes gewählt werden.",
                    children=[
                        bkt.ribbon.ToggleButton(
                            label="Abstand/Winkel einzeln vom Zentrum zu Shapes (Standard)",
                            get_pressed=bkt.Callback(lambda: ShapeDistance.euclid_multi_shape_mode == "centric"),
                            on_toggle_action=bkt.Callback(lambda pressed: setattr(ShapeDistance, "euclid_multi_shape_mode", "centric")),
                        ),
                        bkt.ribbon.ToggleButton(
                            label="Abstände/Winkel immer um gleiche Differenz ändern [ALT]",
                            get_pressed=bkt.Callback(lambda: ShapeDistance.euclid_multi_shape_mode == "delta"),
                            on_toggle_action=bkt.Callback(lambda pressed: setattr(ShapeDistance, "euclid_multi_shape_mode", "delta")),
                        ),
                        bkt.ribbon.ToggleButton(
                            label="Abstände/Winkel gleichmäßig zwischen Shapes verteilen",
                            get_pressed=bkt.Callback(lambda: ShapeDistance.euclid_multi_shape_mode == "distribute"),
                            on_toggle_action=bkt.Callback(lambda pressed: setattr(ShapeDistance, "euclid_multi_shape_mode", "distribute")),
                        ),
                    ]
                ),
                bkt.ribbon.ToggleButton(
                    label="Shape-Rotation angleichen",
                    screentip="Shape-Rotation an Winkel angeichen an/aus",
                    supertip="Passt die Shape-Rotation an den Abstandsvektor bzw. den Winkel an.",
                    get_pressed=bkt.Callback(lambda: getattr(ShapeDistance, "shape_rotate_with_angle")),
                    on_toggle_action=bkt.Callback(lambda pressed: setattr(ShapeDistance, "shape_rotate_with_angle", pressed)),
                ),
            ]
        )
    ]
)


class MasterShapeIndicator(bkt.ui.WpfWindowAbstract):
    _filename = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'master_shape.xaml')
    IsPopup = True

    # def __init__(self, context=None):
    #     self.IsPopup = True
    #     self._context = context

    #     super(MasterShapeIndicator, self).__init__()

class MasterShapeDialog(bkt.contextdialogs.ContextDialog):
    def __init__(self, arranger):
        super(MasterShapeDialog, self).__init__("MASTER")
        self.wnd = None
        self.arranger = arranger
    
    def create_window(self, context):
        # if not self.wnd:
        #     self.wnd = MasterShapeIndicator(context)
        # return self.wnd
        return MasterShapeIndicator(context)

    def get_master_shape(self, shapes):
        return self.arranger.get_master_for_indicator(shapes)



class ArrangeAdvanced(object):
    #class variables:
    # FIXME: master should be an instance-variable, other classes should get an ArrangeAdvanced-instance by dependency injection
    master = "LAST"
    fallback_first_last = "LAST"
    master_indicator = True

    #instance variables:
    margin = 0
    resize = False
    ref_shape = None
    ref_frame = None
    

    # def _create_control(self, control, children, callbacks):
    #     control.children = [
    #         children['arrange_left_at_left'],
    #         children['arrange_middle_at_left'],
    #         children['arrange_right_at_left'],
    #         children['arrange_left_at_middle'],
    #         children['arrange_middle_at_middle'],
    #         children['arrange_right_at_middle'],
    #         children['arrange_left_at_right'],
    #         children['arrange_middle_at_right'],
    #         children['arrange_right_at_right'],
    #
    #         children['arrange_top_at_top'],
    #         children['arrange_middle_at_top'],
    #         children['arrange_bottom_at_top'],
    #         children['arrange_top_at_middle'],
    #         children['arrange_vmiddle_at_vmiddle'],
    #         children['arrange_bottom_at_middle'],
    #         children['arrange_top_at_bottom'],
    #         children['arrange_middle_at_bottom'],
    #         children['arrange_bottom_at_bottom'],
    #
    #         children['set_margin'],
    #         bkt.ribbon.Menu(label="Optionen", children=[
    #             children['set_moving'],
    #             children['set_resizing']
    #         ]),
    #         bkt.ribbon.SplitButton(children=[
    #             bkt.ribbon.ToggleButton(image_mso="PositionAbsoluteMarks"),
    #             bkt.ribbon.Menu(label='Ausrichtung an', children=[
    #                 #bkt.ribbon.MenuSeparator(title='Form aus Selektion'),
    #                 children['set_master_first'],
    #                 children['set_master_last'],
    #                 #bkt.ribbon.MenuSeparator(title='Spezifizierte Form'),
    #                 bkt.ribbon.MenuSeparator(),
    #                 children['set_master_specified'],
    #                 #children['set_master_slide'],
    #                 children['specify_shape']
    #             ])
    #         ])
    #     ]
    #     return control

    def __init__(self):
        #instance variables:
        margin = 0
        resize = False
        ref_shape = None
        ref_frame = None
        
        #save preference for master shape (first or last selected) into class variables
        ArrangeAdvanced.fallback_first_last = bkt.settings.get("arrange_advanced.default", "LAST")
        ArrangeAdvanced.master = ArrangeAdvanced.fallback_first_last
        #save preference to show master shape indicator into class variables
        ArrangeAdvanced.master_indicator = bkt.settings.get("arrange_advanced.master_indicator", True)
        if ArrangeAdvanced.master_indicator:
            self._register_dialog()

        # self.position_gallery = pplib.PositionGallery(
        #     label="Benutzerdef. Bereich wählen",
        #     on_position_change = bkt.Callback(self.specify_master_customarea),
        #     get_item_supertip = bkt.Callback(self.get_item_supertip)
        # )
    
    def get_item_supertip(self, index):
        return 'Verwende angezeigten Position/Größe als Master.'
    

    ### arrange at master's left side ###

    def arrange_left_at_left(self, shapes):
        ''' arrange left side of shapes at masters's left side '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                new_size = self.get_shape_left(shape) + self.get_shape_width(shape) - ( self.get_shape_left(master) + self.margin )
                if self.use_resizing() and new_size>=0:
                    self.set_shape_width(shape, new_size)
                self.set_shape_left( shape, self.get_shape_left(master) + self.margin )


    def arrange_middle_at_left(self, shapes):
        ''' arrange midpoint of shapes horizontally at masters's left side '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                offset = self.get_shape_left(master) - self.margin
                max_distance = max( abs(self.get_shape_left(shape) + self.get_shape_width(shape) - offset), abs(self.get_shape_left(shape) - offset )    )
                new_size = 2*max_distance
                #new_size = 2*(self.get_shape_left(shape) + self.get_shape_width(shape) - self.get_shape_left(master) + self.margin)
                if self.use_resizing() and new_size >=0:
                    self.set_shape_width(shape, new_size)
                self.set_shape_left(shape, self.get_shape_left(master) - self.get_shape_width(shape)/2 - self.margin)



    def arrange_right_at_left(self, shapes):
        ''' arrange right side of shapes at masters's left side '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                new_size = self.get_shape_left(master) - self.get_shape_left(shape) - self.margin
                if self.use_resizing() and new_size >=0:
                    self.set_shape_width(shape, new_size)
                self.set_shape_left(shape, self.get_shape_left(master) - self.get_shape_width(shape) - self.margin)


    ### arrange at master's middle ###

    def arrange_left_at_middle(self, shapes):
        ''' arrange left side of shapes horizontally at masters's midpoint '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                new_size = self.get_shape_left(shape) + self.get_shape_width(shape) - self.get_shape_left(master) - self.get_shape_width(master)/2 - self.margin
                if self.use_resizing() and new_size >=0:
                    self.set_shape_width(shape, new_size)
                self.set_shape_left(shape, self.get_shape_left(master) + self.get_shape_width(master)/2 + self.margin)


    def arrange_middle_at_middle(self, shapes):
        ''' arrange midpoint of shapes horizontally at masters's midpoint '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                offset = self.get_shape_left(master) + self.get_shape_width(master)/2
                max_distance = max( abs(self.get_shape_left(shape) + self.get_shape_width(shape) - offset), abs(self.get_shape_left(shape) - offset )    )
                new_size = 2*max_distance
                if self.use_resizing() and new_size >=0:
                    self.set_shape_width(shape, new_size)
                self.set_shape_left(shape, self.get_shape_left(master) + self.get_shape_width(master)/2 - self.get_shape_width(shape)/2)


    def arrange_right_at_middle(self, shapes):
        ''' arrange right side of shapes horizontally at masters's midpoint '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                new_size = self.get_shape_left(master) + self.get_shape_width(master)/2 - self.get_shape_left(shape) - self.margin
                if self.use_resizing() and new_size >=0:
                    self.set_shape_width(shape, new_size)
                self.set_shape_left(shape, self.get_shape_left(master) + self.get_shape_width(master)/2 - self.get_shape_width(shape) - self.margin)


    ### arrange at master's right side ###

    def arrange_left_at_right(self, shapes):
        ''' arrange left side of shapes at masters's right side '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                new_size = self.get_shape_left(shape) + self.get_shape_width(shape) - self.get_shape_left(master) - self.get_shape_width(master) - self.margin
                if self.use_resizing() and new_size >=0:
                    self.set_shape_width(shape, new_size)
                self.set_shape_left(shape, self.get_shape_left(master) + self.get_shape_width(master) + self.margin)


    def arrange_middle_at_right(self, shapes):
        ''' arrange midpoint of shapes horizontally at masters's right side '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                offset = self.get_shape_left(master) + self.get_shape_width(master) + self.margin
                max_distance = max( abs(self.get_shape_left(shape) + self.get_shape_width(shape) - offset), abs(self.get_shape_left(shape) - offset )    )
                new_size = 2*max_distance
                if self.use_resizing() and new_size >=0:
                    self.set_shape_width(shape, new_size)
                self.set_shape_left(shape, self.get_shape_left(master) + self.get_shape_width(master) - self.get_shape_width(shape)/2 + self.margin)


    def arrange_right_at_right(self, shapes):
        ''' arrange right side of shapes at masters's right side '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                new_size = self.get_shape_left(master) + self.get_shape_width(master) - self.get_shape_left(shape) - self.margin
                if self.use_resizing() and new_size >=0:
                    self.set_shape_width(shape, new_size)
                self.set_shape_left(shape, self.get_shape_left(master) + self.get_shape_width(master) - self.get_shape_width(shape) - self.margin)


    ############

    ### arrange at master's top side ###

    def arrange_top_at_top(self, shapes):
        ''' arrange top side of shapes at masters's top side '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                new_size = self.get_shape_top(shape) + self.get_shape_height(shape) - ( self.get_shape_top(master) + self.margin )
                if self.use_resizing() and new_size >=0:
                    self.set_shape_height(shape, new_size)
                self.set_shape_top(shape, self.get_shape_top(master) + self.margin)


    def arrange_middle_at_top(self, shapes):
        ''' arrange midpoint of shapes at masters's top side '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                new_size = 2*(self.get_shape_top(shape) + self.get_shape_height(shape) - self.get_shape_top(master) + self.margin)
                if self.use_resizing() and new_size >=0:
                    self.set_shape_height(shape, new_size)
                self.set_shape_top(shape, self.get_shape_top(master) - self.get_shape_height(shape)/2 - self.margin)


    def arrange_bottom_at_top(self, shapes):
        ''' arrange bottom side of shapes at masters's top side '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                new_size = self.get_shape_top(master) - self.get_shape_top(shape) - self.margin
                if self.use_resizing() and new_size >=0:
                    self.set_shape_height(shape, new_size)
                self.set_shape_top(shape, self.get_shape_top(master) - self.get_shape_height(shape) - self.margin)


    ### arrange at master's vertical middle ###

    def arrange_top_at_middle(self, shapes):
        ''' arrange top side of shapes vertically at masters's midpoint '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                new_size = self.get_shape_top(shape) + self.get_shape_height(shape) - self.get_shape_top(master) - self.get_shape_height(master)/2 - self.margin
                if self.use_resizing() and new_size >=0:
                    self.set_shape_height(shape, new_size)
                self.set_shape_top(shape, self.get_shape_top(master) + self.get_shape_height(master)/2 + self.margin)


    def arrange_vmiddle_at_vmiddle(self, shapes):
        ''' arrange midpoint of shapes vertically at masters's midpoint '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                new_size = 2*(self.get_shape_top(shape) + self.get_shape_height(shape) - self.get_shape_top(master) - self.get_shape_height(master)/2)
                if self.use_resizing() and new_size >=0:
                    self.set_shape_height(shape, new_size)
                self.set_shape_top(shape, self.get_shape_top(master) + self.get_shape_height(master)/2 - self.get_shape_height(shape)/2)


    def arrange_bottom_at_middle(self, shapes):
        ''' arrange bottom side of shapes vertically at masters's midpoint '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                new_size = self.get_shape_top(master) + self.get_shape_height(master)/2 - self.get_shape_top(shape) - self.margin
                if self.use_resizing() and new_size >=0:
                    self.set_shape_height(shape, new_size)
                self.set_shape_top(shape, self.get_shape_top(master) + self.get_shape_height(master)/2 - self.get_shape_height(shape) - self.margin)


    ### arrange at master's right side ###

    def arrange_top_at_bottom(self, shapes):
        ''' arrange top side of shapes at masters's bottom side '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                new_size = self.get_shape_top(shape) + self.get_shape_height(shape) - self.get_shape_top(master) - self.get_shape_height(master) - self.margin
                if self.use_resizing() and new_size >=0:
                    self.set_shape_height(shape, new_size)
                self.set_shape_top(shape, self.get_shape_top(master) + self.get_shape_height(master) + self.margin)


    def arrange_middle_at_bottom(self, shapes):
        ''' arrange midpoint of shapes vertically at masters's bottom side '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                new_size = 2*(self.get_shape_top(shape) + self.get_shape_height(shape) - self.get_shape_top(master) - self.get_shape_height(master)-self.margin)
                if self.use_resizing() and new_size >=0:
                    self.set_shape_height(shape, new_size)
                self.set_shape_top(shape, self.get_shape_top(master) + self.get_shape_height(master) - self.get_shape_height(shape)/2 + self.margin)

    def arrange_bottom_at_bottom(self, shapes):
        ''' arrange bottom side of shapes at masters's bottom side '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                new_size = self.get_shape_top(master) + self.get_shape_height(master) - self.get_shape_top(shape) - self.margin
                if self.use_resizing() and new_size >=0:
                    self.set_shape_height(shape, new_size)
                self.set_shape_top(shape, self.get_shape_top(master) + self.get_shape_height(master) - self.get_shape_height(shape) - self.margin)


    ### quick arrange ###

    def arrange_quick_position(self, shapes):
        ''' position shapes at master's position '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                self.set_shape_left( shape, self.get_shape_left(master) )
                self.set_shape_top( shape, self.get_shape_top(master) )

    def arrange_quick_size(self, shapes):
        ''' change size of shapes according to master's size '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                self.set_shape_width( shape, self.get_shape_width(master) )
                self.set_shape_height( shape, self.get_shape_height(master) )



    def enabled(self, shapes):
        #return (self.master in ["FIXED-SHAPE", "FIXED-SLIDE", "FIXED-CONTENTAREA", "FIXED-CUSTOMAREA"] and len(shapes) > 0) or len(shapes) > 1
        return len(shapes) > 0


    ### detect master shape ###

    def get_master_from_shapes(self, shapes):
        def _test_ref_or_use_fallback(ref):
            try:
                ref.left #test if ref still exists
                return ref
            except:
                bkt.helpers.message("Fehler: Referenz wurde nicht gefunden. Fallback zu Ausrichtung am Mastershape innerhalb der Selektion.")
                ArrangeAdvanced.master = self.fallback_first_last
                return self.get_master_from_shapes(shapes)

        ''' obtain master shape from given shapes according to master-setting '''
        if ArrangeAdvanced.master == "FIXED-SHAPE" and self.ref_shape != None:
            return _test_ref_or_use_fallback(self.ref_shape)
        elif ArrangeAdvanced.master == "FIXED-SLIDE" and self.ref_frame !=None:
            return _test_ref_or_use_fallback(self.ref_frame)
        elif ArrangeAdvanced.master == "FIXED-CONTENTAREA" and self.ref_frame !=None:
            return _test_ref_or_use_fallback(self.ref_frame)
        elif ArrangeAdvanced.master == "FIXED-CUSTOMAREA" and self.ref_frame !=None:
            return _test_ref_or_use_fallback(self.ref_frame)
            
        elif len(shapes) == 1:
            ## fallback if only one shape in selection
            return pplib.BoundingFrame(shapes[0].parent, contentarea=True)

        elif ArrangeAdvanced.master == "PPTDEFAULT":
            return pplib.BoundingFrame.from_shapes(shapes)
            
        elif ArrangeAdvanced.master == "FIRST":
            return shapes[0]
        else:
            # default: master == "LAST"
            return shapes[-1]
    
    
    ### configure margin, resizing-option ###

    def set_margin(self, value, context):
        ''' callback to set margin-value '''
        self.margin = value

    def get_margin(self):
        ''' returns current margin-value '''
        return self.margin

    def set_moving(self, pressed):
        ''' callback to switch between moving/resizing-option '''
        self.resize=(pressed==False)

    def get_moving(self):
        ''' returns if moving-option is used '''
        return (self.resize==False)

    def set_resizing(self, pressed):
        ''' callback to switch between resizing/moving-option '''
        self.resize=(pressed==True)

    def get_resizing(self):
        ''' returns if resizing-option is set '''
        return (self.resize==True)
    
    def use_resizing(self):
        ''' returns if resizing-option should be used. Depends on setting for resize-option, and SHIFT-key-state '''
        if bkt.library.system.get_key_state(bkt.library.system.key_code.SHIFT):
            return True
        else:
            return self.get_resizing()
    
    
    ### configure master in selection ###

    def set_master_first(self, pressed):
        ''' callback: set master-setting to use first shape in selection '''
        ArrangeAdvanced.fallback_first_last="FIRST"
        ArrangeAdvanced.master="FIRST"
        bkt.settings["arrange_advanced.default"] = "FIRST"

    def get_master_first(self):
        ''' returns whether master-setting is set to first shape in selection '''
        return (ArrangeAdvanced.master=="FIRST")


    def set_master_last(self, pressed):
        ''' callback: set master-setting to use last shape in selection '''
        ArrangeAdvanced.fallback_first_last="LAST"
        ArrangeAdvanced.master="LAST"
        bkt.settings["arrange_advanced.default"] = "LAST"

    def get_master_last(self):
        ''' returns whether master-setting is set to last shape in selection '''
        return (ArrangeAdvanced.master=="LAST")


    def set_master_pptdefault(self, pressed):
        ''' callback: set master-setting to use outermost shape in selection (as powerpoint default aligning) '''
        ArrangeAdvanced.fallback_first_last="PPTDEFAULT"
        ArrangeAdvanced.master="PPTDEFAULT"
        bkt.settings["arrange_advanced.default"] = "PPTDEFAULT"

    def get_master_pptdefault(self):
        ''' returns whether master-setting is set to outermost shape in selection (as powerpoint default aligning) '''
        return (ArrangeAdvanced.master=="PPTDEFAULT")


    def set_master_indicator(self, pressed):
        ''' callback: set whether master shape indicator is shown '''
        if not pressed:
            ArrangeAdvanced.master_indicator = False
            bkt.settings["arrange_advanced.master_indicator"] = False
            self._unregister_dialog()
        else:
            ArrangeAdvanced.master_indicator = True
            bkt.settings["arrange_advanced.master_indicator"] = True
            self._register_dialog()

    def get_master_indicator(self):
        ''' returns whether master shape indicator is shown '''
        return self.master_indicator

    @classmethod
    def _register_dialog(cls):
        # register dialog
        bkt.powerpoint.context_dialogs.register_dialog(
            MasterShapeDialog(cls)
        )
    
    @classmethod
    def _unregister_dialog(cls):
        # unregister dialog
        bkt.powerpoint.context_dialogs.unregister("MASTER")

    @classmethod
    def get_master_for_indicator(cls, shapes):
        if cls.master=="FIRST":
            return shapes[0]
        elif cls.master=="LAST":
            return shapes[-1]
        else:
            return None



    ### configure fixed master-shape ###
    
    def specify_shape(self, shape, pressed):
        ''' callback: set master to selected shape. On deactivation fallback to master in selection '''
        if pressed:
            self.ref_shape = shape
            ArrangeAdvanced.master = "FIXED-SHAPE"
        else:
            ArrangeAdvanced.master = self.fallback_first_last

    def specify_master_slide(self, slide, pressed):
        ''' callback: set master to slide. On deactivation fallback to master in selection '''
        if pressed:
            self.ref_frame = pplib.BoundingFrame(slide)
            ArrangeAdvanced.master = "FIXED-SLIDE"
        else:
            ArrangeAdvanced.master = self.fallback_first_last

    def specify_master_contentarea(self, slide, pressed):
        ''' callback: set master to content-area of slide. On deactivation fallback to master in selection '''
        if pressed:
            self.ref_frame = pplib.BoundingFrame(slide, contentarea=True)
            ArrangeAdvanced.master = "FIXED-CONTENTAREA"
        else:
            ArrangeAdvanced.master = self.fallback_first_last
            
    def specify_master_customarea(self, target_frame):
        ''' set master to given target-frame '''
        ArrangeAdvanced.master = "FIXED-CUSTOMAREA"
        self.ref_frame = target_frame
    
    
    def specify_wiz(self, pressed, context):
        ''' callback for master-wizzard: sets master to selected shape, content-area or master in selection - 
            depending on given shape-selection and current master-setting
        '''
        if not pressed:
            ArrangeAdvanced.master = self.fallback_first_last
            
        else:
            # pressed == True
            shapes = self.get_shapes_from_context(context)
            
            if len(shapes) == 0:
                # use contentarea as default
                self.specify_master_contentarea(context.app.activewindow.selection.SlideRange.Item(1), pressed=True)
            
            else:
                # use shape
                self.specify_shape(shapes[0], pressed=True)
                # FIXME: message if multiple shapes are selected?
    
    
    def is_slide_or_shape_specified(self):
        ''' returns whether a fixed master is used '''
        return ArrangeAdvanced.master in ["FIXED-SHAPE", "FIXED-SLIDE", "FIXED-CONTENTAREA", "FIXED-CUSTOMAREA"]
        
    def is_shape_specified(self):
        ''' returns whether master is a fixed shape '''
        return ArrangeAdvanced.master == "FIXED-SHAPE"
    
    def is_slide_specified(self):
        ''' returns whether master is set to slide '''
        return ArrangeAdvanced.master == "FIXED-SLIDE"

    def is_contentarea_specified(self):
        ''' returns whether master is set to slide content area '''
        return ArrangeAdvanced.master == "FIXED-CONTENTAREA"

    def is_customarea_specified(self):
        ''' returns whether master is set to custom area '''
        return ArrangeAdvanced.master == "FIXED-CUSTOMAREA"
    
    def is_shape_specified_or_shape_specifiable(self, context):
        ''' callback: returns whether master can be changed to fixed shape.
            Can either use a shape from selection or the last fixed shape.
        '''
        if self.is_shape_specified():
            return True
        else:
            shapes = self.get_shapes_from_context(context)
            return len(shapes) > 0
    
    
    
    ### helper function for shapes ###
    
    def get_shapes_from_context(self, context):
        ''' retrieves shapes from context, depending on selection-type (text, shapes, group) '''
        try:
            if context.app.ActiveWindow.ViewType != 9:
               return []

            selection = context.app.ActiveWindow.Selection
        except:
            return []
        
        return pplib.get_shapes_from_selection(selection)

        # # ShapeRange accessible if shape or text selected
        # shapes = []
        # if selection.Type == 2 or selection.Type == 3:
        #     try:
        #         if selection.HasChildShapeRange:
        #             # shape selection inside grouped shapes
        #             shapes = list(iter(selection.ChildShapeRange))
        #         else:
        #             shapes = list(iter(selection.ShapeRange))
        #     except:
        #         shapes = []
        
        # return shapes
        
    

    ### helper functions to compute/change actual left/top/width/height of rotated shapes ###

    ### actual shape's left ###

    def get_shape_left(self, shape):
        ''' return actual shape's left boundary considering the shape's rotation '''
        if isinstance(shape, pplib.BoundingFrame):
            return shape.left
        shapes_left = min( p[0] for p in self.get_shape_bounding_nodes(shape) )
        return shapes_left

    def set_shape_left(self, shape, value):
        ''' set shape's left considering the shape's rotation '''
        delta = shape.left - self.get_shape_left(shape)
        shape.left = value + delta

    ### actual shape's width ###

    def get_shape_width(self, shape):
        ''' return actual shape's width considering the shape's rotation '''
        if isinstance(shape, pplib.BoundingFrame):
            return shape.width
        return max( p[0] for p in self.get_shape_bounding_nodes(shape) ) - self.get_shape_left(shape)

    def set_shape_width(self, shape, value):
        ''' set shape's width considering the shape's rotation '''
        if shape.rotation == 0 or shape.rotation == 180:
            shape.width = value
        elif shape.rotation == 90 or shape.rotation == 270:
            shape.height = value
        else:
            delta = value - self.get_shape_width(shape)
            # delta_vector (delta-width, 0) um shape-rotation drehen
            # delta_vector = self.rotate_point(delta, 0, 0, 0, shape.rotation)
            delta_vector = algos.rotate_point_by_shape_rotation(delta, 0, shape)
            # vorzeichen beibehalten (entweder vergrößern oder verkleinern - nicht beides)
            vorzeichen = 1 if delta > 0 else -1
            delta_vector = [vorzeichen * abs(delta_vector[0]), vorzeichen * abs(delta_vector[1]) ]
            # shape anpassen
            shape.width += delta_vector[0]
            shape.height += delta_vector[1]

    ### actual shape's top ###

    def get_shape_top(self, shape):
        ''' return actual shape's top boundary considering the shape's rotation '''
        if isinstance(shape, pplib.BoundingFrame):
            return shape.top
        return min( p[1] for p in self.get_shape_bounding_nodes(shape) )

    def set_shape_top(self, shape, value):
        ''' set shape's top considering the shape's rotation '''
        delta = shape.top - self.get_shape_top(shape)
        shape.top = value + delta

    ### actual shape's height ###

    def get_shape_height(self, shape):
        ''' return actual shape's height considering the shape's rotation '''
        if isinstance(shape, pplib.BoundingFrame):
            return shape.height
        return max( p[1] for p in self.get_shape_bounding_nodes(shape) ) - self.get_shape_top(shape)

    def set_shape_height(self, shape, value):
        ''' set shape's height considering the shape's rotation '''
        if shape.rotation == 0 or shape.rotation == 180:
            shape.height = value
        elif shape.rotation == 90 or shape.rotation == 270:
            shape.width = value
        else:
            delta = value - self.get_shape_height(shape)
            # delta_vector (0, delta-height) um shape-rotation drehen
            # delta_vector = self.rotate_point(0, delta, 0, 0, shape.rotation)
            delta_vector = algos.rotate_point_by_shape_rotation(0, delta, shape)
            # vorzeichen beibehalten (entweder vergrößern oder verkleinern - nicht beides)
            vorzeichen = 1 if delta > 0 else -1
            delta_vector = [vorzeichen * abs(delta_vector[0]), vorzeichen * abs(delta_vector[1]) ]
            # shape anpassen
            shape.width += delta_vector[0]
            shape.height += delta_vector[1]


    ### helper fuctions to compute rotated points ###

    def get_shape_bounding_nodes(self, shape):
        ''' compute bounding nodes (surrounding-square) for rotated shapes '''
        return algos.get_bounding_nodes(shape)
        # points = [ [ shape.left, shape.top ], [shape.left, shape.top+shape.height], [shape.left+shape.width, shape.top+shape.height], [shape.left+shape.width, shape.top] ]

        # x0 = shape.left + shape.width/2
        # y0 = shape.top + shape.height/2

        # rotated_points = [
        #     list(algos.rotate_point(p[0], p[1], x0, y0, 360-shape.rotation))
        #     for p in points
        # ]
        # return rotated_points



    ### ui generation functions ###
    def get_master_button(self, postfix="", **kwargs):
        return bkt.ribbon.SplitButton(id="arrange_master_splitbutton"+postfix, children=[
            
            bkt.ribbon.ToggleButton(
                id="arrange_use_master"+postfix,
                label="Master",
                image_mso="PositionAbsoluteMarks", on_toggle_action=bkt.Callback(self.specify_wiz), get_pressed=bkt.Callback(self.is_slide_or_shape_specified),
                screentip="Ausrichtung an Master (Shape/Slide) ein-/ausschalten", 
                supertip="Bei Aktivierung werden anschließend alle Shapes am selektierten Shape (Mastershape) ausgerichtet. Ist bei Aktivierung kein Shape selektiert, erfolgt die Ausrichtung am Inhaltsbereich."
            ),
            
            bkt.ribbon.Menu(label='Auswahl der Ausrichtung', children=[
                
                bkt.ribbon.Button(
                    id="arrange_master_auto"+postfix,
                    label="Automatische Masterwahl",
                    image_mso="PositionAbsoluteMarks", 
                    supertip="Macht das zuerst ausgewählte Shape zum Mastershape. Ist kein Shape markiert, wird der Inhaltsbereich als Master verwendet.",
                    on_action=bkt.Callback(lambda context:self.specify_wiz(pressed=True, context=context))
                ),
                
                bkt.ribbon.MenuSeparator(title="Ausrichtung innerhalb Selektion…"),
                bkt.ribbon.ToggleButton(
                    id="arrange_use_first_shape"+postfix,
                    label="…an erstem Shape",  on_toggle_action=bkt.Callback(self.set_master_first), get_pressed=bkt.Callback(self.get_master_first),
                    screentip="Ausrichtung am ersten Shape innerhalb der Selektion",
                    supertip="Shapes werden am zuerst selektierten Shape in der Selektion ausgerichtet."
                ),
                bkt.ribbon.ToggleButton(
                    id="arrange_use_last_shape"+postfix,
                    label="…an letztem Shape", on_toggle_action=bkt.Callback(self.set_master_last),  get_pressed=bkt.Callback(self.get_master_last),
                    screentip="Ausrichtung am letzten Shape innerhalb der Selektion",
                    supertip="Shapes werden am zuletzt selektierten Shape in der Selektion ausgerichtet."
                ),
                bkt.ribbon.ToggleButton(
                    id="arrange_use_pptdefault_shape"+postfix,
                    label="…an äußerstem Shape (PPT-Standard)", on_toggle_action=bkt.Callback(self.set_master_pptdefault),  get_pressed=bkt.Callback(self.get_master_pptdefault),
                    screentip="Ausrichtung am äußerstem Shape innerhalb der Selektion",
                    supertip="Shapes werden am äußersten selektierten Shape in der Selektion ausgerichtet."
                ),
                bkt.ribbon.MenuSeparator(),
                bkt.ribbon.ToggleButton(
                    id="arrange_show_master_shape"+postfix,
                    label="Indikator an Master-Shape", on_toggle_action=bkt.Callback(self.set_master_indicator),  get_pressed=bkt.Callback(self.get_master_indicator),
                    screentip="Indikator an dem Master-Shape innerhalb der Selektion anzeigen",
                    supertip="Werden mind. zwei Shapes ausgewählt wird ein kleiner Indikator mit dem Text 'Master' an der unteren linken Ecke des Master-Shape (erstes bzw. letztes) angezeigt."
                ),
                
                bkt.ribbon.MenuSeparator(title="Ausrichtung an Master"),
                bkt.ribbon.ToggleButton(
                    id="arrange_use_shape"+postfix,
                    label="Shape", on_toggle_action=bkt.Callback(self.specify_shape),         get_pressed=bkt.Callback(self.is_shape_specified), get_enabled=bkt.Callback(self.is_shape_specified_or_shape_specifiable),
                    screentip="Ausrichtung am selektierten Shape (Mastershape)",
                    supertip="Das selektierten Shape wird als Mastershape festgelegt. Shapes werden am Mastershape ausgerichtet. "
                ),
                bkt.ribbon.ToggleButton(
                    id="arrange_use_slide"+postfix,
                    label="Folie", on_toggle_action=bkt.Callback(self.specify_master_slide),  get_pressed=bkt.Callback(self.is_slide_specified),
                    screentip="Ausrichtung an der Folie",
                    supertip="Shapes werden an der Folie ausgerichtet."
                ),
                bkt.ribbon.ToggleButton(
                    id="arrange_use_contentarea"+postfix,
                    label="Inhaltsbereich",
                    on_toggle_action=bkt.Callback(self.specify_master_contentarea),
                    get_pressed=bkt.Callback(self.is_contentarea_specified),
                    screentip="Ausrichtung an Inhaltsbereich",
                    supertip="Shapes werden am Inhaltsbereich ausgerichtet.\n\nDer Inhaltsbereich entspricht der Fläche des Text-Platzhalters auf dem Master-Slide.",
                ),
                bkt.ribbon.ToggleButton(
                    id="arrange_use_customarea"+postfix,
                    label="Benutzerdef. Bereich",
                    #on_toggle_action=bkt.Callback(self.specify_master_contentarea),
                    get_pressed=bkt.Callback(self.is_customarea_specified),
                    get_enabled=bkt.Callback(self.is_customarea_specified),
                    screentip="Ausrichtung an benutzerdefiniertem Bereich",
                    supertip="Shapes werden an einem festgelegten Bereich ausgerichtet, der zuvor durch den Benutzer definiert wird.",
                ),
                bkt.ribbon.MenuSeparator(),
                pplib.PositionGallery(
                    id="arrange_set_customarea"+postfix,
                    label="Benutzerdef. Bereich wählen",
                    on_position_change = bkt.Callback(self.specify_master_customarea),
                    get_item_supertip = bkt.Callback(self.get_item_supertip)
                ),
            ])
        ],
        **kwargs)

    def get_button(self, arrange_id, postfix="", **kwargs):
        return bkt.ribbon.Button(
            id=arrange_id+postfix,
            on_action=bkt.Callback(getattr(self, arrange_id)),
            get_enabled=bkt.Callback(self.enabled),
            image=arrange_id,
            show_label=False,
        **kwargs)


arrange_advaced = ArrangeAdvanced()




arrange_advanced_group = bkt.ribbon.Group(
    id="bkt_arrage_adv_group",
    label=u'Erweitertes Anordnen',
    image='arrange_left_at_left',
    children=[
        arrange_advaced.get_button('arrange_left_at_left',         label="Links an Links",   screentip='Ausrichtung linke Kante an linke Kante',   supertip='Ausrichtung der linken Kante an der linken Kante des Mastershapes.\n(Abstand wird addiert)\n\nMit SHIFT-Taste wird auf Dehnen/Stauchen umgestellt.'),
        arrange_advaced.get_button('arrange_middle_at_left',       label="Mitte an Links",   screentip='Ausrichtung Mitte an linke Kante',         supertip='Ausrichtung der Shapemitte an der linken Kante des Mastershapes.\n(Abstand wird addiert)\n\nMit SHIFT-Taste wird auf Dehnen/Stauchen umgestellt.'),
        arrange_advaced.get_button('arrange_right_at_left',        label="Rechts an Links",  screentip='Ausrichtung rechte Kante an linke Kante',  supertip='Ausrichtung der rechten Kante an der linken Kante des Mastershapes.\n(Abstand wird subtrahiert)\n\nMit SHIFT-Taste wird auf Dehnen/Stauchen umgestellt.'),
        arrange_advaced.get_button('arrange_left_at_middle',       label="Links an Mitte",   screentip='Ausrichtung linke Kante an Shapemitte',    supertip='Ausrichtung der linken Kante an der Shapemitte des Mastershapes.\n(Abstand wird addiert)\n\nMit SHIFT-Taste wird auf Dehnen/Stauchen umgestellt.'),
        arrange_advaced.get_button('arrange_middle_at_middle',     label="Mitte an Mitte",   screentip='Ausrichtung Shapemitte an Shapemitte',     supertip='Ausrichtung der Shapemitte an der Shapemitte des Mastershapes.\n(kein Abstand)\n\nMit SHIFT-Taste wird auf Dehnen/Stauchen umgestellt.'),
        arrange_advaced.get_button('arrange_right_at_middle',      label="Rechts an Mitte",  screentip='Ausrichtung rechte Kante an Shapemitte',   supertip='Ausrichtung der rechten Kante an der Shapemitte des Mastershapes.\n(Abstand wird subtrahiert)\n\nMit SHIFT-Taste wird auf Dehnen/Stauchen umgestellt.'),
        arrange_advaced.get_button('arrange_left_at_right',        label="Links an Rechts",  screentip='Ausrichtung linke Kante an rechte Kante',  supertip='Ausrichtung der linken Kante an der rechten Kante des Mastershapes.\n(Abstand wird addiert)\n\nMit SHIFT-Taste wird auf Dehnen/Stauchen umgestellt.'),
        arrange_advaced.get_button('arrange_middle_at_right',      label="Mitte an Rechts",  screentip='Ausrichtung Shapemitte an rechte Kante',   supertip='Ausrichtung der Shapemitte an der rechten Kante des Mastershapes.\n(Abstand wird subtrahiert)\n\nMit SHIFT-Taste wird auf Dehnen/Stauchen umgestellt.'),
        arrange_advaced.get_button('arrange_right_at_right',       label="Rechts an Rechts", screentip='Ausrichtung rechte Kante an rechte Kante', supertip='Ausrichtung der rechten Kante an der rechten Kante des Mastershapes.\n(Abstand wird subtrahiert)\n\nMit SHIFT-Taste wird auf Dehnen/Stauchen umgestellt.'),
        arrange_advaced.get_button('arrange_top_at_top',           label="Oben an oben",     screentip='Ausrichtung obere Kante an obere Kante',   supertip='Ausrichtung der oberen Kante an der oberen Kante des Mastershapes.\n(Abstand wird addiert)\n\nMit SHIFT-Taste wird auf Dehnen/Stauchen umgestellt.'),
        arrange_advaced.get_button('arrange_middle_at_top',        label="Mitte an Oben",    screentip='Ausrichtung Shapemitte an obere Kante',    supertip='Ausrichtung der Shapemitte an der oberen Kante des Mastershapes.\n(Abstand wird addiert)\n\nMit SHIFT-Taste wird auf Dehnen/Stauchen umgestellt.'),
        arrange_advaced.get_button('arrange_bottom_at_top',        label="Unten an Oben",    screentip='Ausrichtung untere Kante an obere Kante',  supertip='Ausrichtung der unteren Kante an der oberen Kante des Mastershapes.\n(Abstand wird subtrahiert)\n\nMit SHIFT-Taste wird auf Dehnen/Stauchen umgestellt.'),
        arrange_advaced.get_button('arrange_top_at_middle',        label="Oben an Mitte",    screentip='Ausrichtung obere Kante an Shapemitte',    supertip='Ausrichtung der oberen Kante an der Shapemitte des Mastershapes.\n(Abstand wird addiert)\n\nMit SHIFT-Taste wird auf Dehnen/Stauchen umgestellt.'),
        arrange_advaced.get_button('arrange_vmiddle_at_vmiddle',   label="Mitte an Mitte",   screentip='Ausrichtung Shapemitte an Shapemitte',     supertip='Ausrichtung der Shapemitte an der Shapemitte des Mastershapes.\n(kein Abstand)\n\nMit SHIFT-Taste wird auf Dehnen/Stauchen umgestellt.'),
        arrange_advaced.get_button('arrange_bottom_at_middle',     label="Unten an Mitte",   screentip='Ausrichtung untere Kante an Shapemitte',   supertip='Ausrichtung der unteren Kante an der Shapemitte des Mastershapes.\n(Abstand wird subtrahiert)\n\nMit SHIFT-Taste wird auf Dehnen/Stauchen umgestellt.'),
        arrange_advaced.get_button('arrange_top_at_bottom',        label="Oben an Unten",    screentip='Ausrichtung obere Kante an untere Kante',  supertip='Ausrichtung der oberen Kante an der unteren Kante des Mastershapes.\n(Abstand wird addiert)\n\nMit SHIFT-Taste wird auf Dehnen/Stauchen umgestellt.'),
        arrange_advaced.get_button('arrange_middle_at_bottom',     label="Mitte an Unten",   screentip='Ausrichtung Shapemitte an untere Kante',   supertip='Ausrichtung der Shapemitte an der unteren Kante des Mastershapes.\n(Abstand wird subtrahiert)\n\nMit SHIFT-Taste wird auf Dehnen/Stauchen umgestellt.'),
        arrange_advaced.get_button('arrange_bottom_at_bottom',     label="Unten an Unten",   screentip='Ausrichtung untere Kante an untere Kante', supertip='Ausrichtung der unteren Kante an der unteren Kante des Mastershapes.\n(Abstand wird subtrahiert)\n\nMit SHIFT-Taste wird auf Dehnen/Stauchen umgestellt.'),
        
        bkt.ribbon.Separator(),

        bkt.ribbon.RoundingSpinnerBox(
            id='arrange_distance',
            label="Ausrichtungsabstand",
            show_label=False,
            round_cm=True, convert="pt_to_cm",
            image_mso="HorizontalSpacingIncrease", on_change=bkt.Callback(arrange_advaced.set_margin), get_text=bkt.Callback(arrange_advaced.get_margin),
            screentip="Abstand bei Ausrichtung",
            supertip="Eingestellter Abstand wird bei der Ausrichtung von Shapes links/rechts berücksichtigt.\n\nDer Abstand wird addiert: bei Ausrichtung der linken/oberen Kante (des zu verschiebenden Shapes) und bei Ausrichtung der Shape-Mitte an der linken/oberen Kante des Mastershapes.\n\nDer Abstand wird subtrahiert: bei Ausrichtung der rechten/unteren Kante (des zu verschiebenden Shapes) und bei Ausrichtung der Shape-Mitte an der rechten/unteren Kante des Mastershapes."
        ),
        
        #bkt.ribbon.Menu(label="Optionen", children=[
        bkt.ribbon.Box(box_style="horizontal", children=[
            bkt.ribbon.Label(label="Modus: "),
            bkt.ribbon.ToggleButton(id="arrange_move",   label="Bewegen",         show_label=False, supertip="Gewünschte Shape-Anordnung wird durch Positionierung von Shapes erreicht", image_mso="ObjectNudgeRight", on_toggle_action=bkt.Callback(arrange_advaced.set_moving),   get_pressed=bkt.Callback(arrange_advaced.get_moving)),
            bkt.ribbon.ToggleButton(id="arrange_resize", label="Dehnen/Stauchen [SHIFT]", show_label=False, supertip="Gewünschte Shape-Anordnung wird durch Verkleinerung/Vergrößerung von Shapes erreicht", image_mso="ShapeWidth",       on_toggle_action=bkt.Callback(arrange_advaced.set_resizing), get_pressed=bkt.Callback(arrange_advaced.get_resizing))
        ]),
        
        arrange_advaced.get_master_button(),

        bkt.ribbon.Separator(),

        bkt.ribbon.Label(label="Schnellwahl:"),
        bkt.ribbon.Button(
            id="arrange_quick_position",
            label="Position",
            image_mso="ControlAlignToGrid",
            on_action=bkt.Callback(arrange_advaced.arrange_quick_position),
            get_enabled=bkt.Callback(arrange_advaced.enabled),
            screentip="Gleiche Position wie Master",
        ),
        bkt.ribbon.Button(
            id="arrange_quick_size",
            label="Größe",
            image_mso="SizeToControlHeightAndWidth",
            on_action=bkt.Callback(arrange_advaced.arrange_quick_size),
            get_enabled=bkt.Callback(arrange_advaced.enabled),
            screentip="Gleiche Größe wie Master",
        ),
    ]
)



class ArrangeAdvancedEasy(ArrangeAdvanced):
    margin = 0

    def __init__(self, resize_mode):
        self.resize_mode = resize_mode
        # super(ArrangeAdvancedEasy, self).__init__()

    def get_master_from_shapes(self, shapes):
        if len(shapes) == 1:
            ## fallback if only one shape in selection
            return pplib.BoundingFrame(shapes[0].parent, contentarea=True)
        elif ArrangeAdvanced.master == "FIRST":
            return shapes[0]
        elif ArrangeAdvanced.master == "PPTDEFAULT":
            return pplib.BoundingFrame.from_shapes(shapes)
        else:
            # default: master == "LAST"
            return shapes[-1]

    # POSITION (resize=False)
    #arrange_left_at_left
    #arrange_middle_at_middle
    #arrange_right_at_right
    #arrange_top_at_top
    #arrange_vmiddle_at_vmiddle
    #arrange_bottom_at_bottom

    # DOCK (resize=False)
    #arrange_left_at_right
    #arrange_right_at_left
    #arrange_top_at_bottom
    #arrange_bottom_at_top

    # STRETCH (resize=True)
    #arrange_left_at_left
    #arrange_right_at_right
    #arrange_top_at_top
    #arrange_bottom_at_bottom

    # FILL (resize=True)
    #arrange_left_at_right
    #arrange_right_at_left
    #arrange_top_at_bottom
    #arrange_bottom_at_top

    def use_resizing(self):
        return self.resize_mode


arrange_adv_position = ArrangeAdvancedEasy(False)
arrange_adv_size     = ArrangeAdvancedEasy(True)

arrange_adv_easy_group = bkt.ribbon.Group(
    id="bkt_arrage_adv_easy_group",
    label=u'Einfaches Anordnen',
    image='arrange_left_at_left',
    children=[
        #POSITION
        bkt.ribbon.Button(id='arrange_position_left',     on_action=bkt.Callback(arrange_adv_position.arrange_left_at_left),       get_enabled=bkt.Callback(arrange_adv_position.enabled), image='arrange_position_left',        label="Links an Links",   show_label=False, screentip='Ausrichtung linke Kante an linke Kante',   supertip='Ausrichtung der linken Kante an der linken Kante des Mastershapes.'),
        bkt.ribbon.Button(id='arrange_position_right',    on_action=bkt.Callback(arrange_adv_position.arrange_right_at_right),     get_enabled=bkt.Callback(arrange_adv_position.enabled), image='arrange_position_right',       label="Rechts an Rechts", show_label=False, screentip='Ausrichtung rechte Kante an rechte Kante', supertip='Ausrichtung der rechten Kante an der rechten Kante des Mastershapes.'),
        bkt.ribbon.Button(id='arrange_position_middle',   on_action=bkt.Callback(arrange_adv_position.arrange_middle_at_middle),   get_enabled=bkt.Callback(arrange_adv_position.enabled), image='arrange_position_middle',      label="Mitte an Mitte",   show_label=False, screentip='Ausrichtung Shapemitte an Shapemitte',     supertip='Ausrichtung der Shapemitte an der Shapemitte des Mastershapes.'),

        bkt.ribbon.Button(id='arrange_position_top',      on_action=bkt.Callback(arrange_adv_position.arrange_top_at_top),         get_enabled=bkt.Callback(arrange_adv_position.enabled), image='arrange_position_top',        label="Oben an oben",     show_label=False, screentip='Ausrichtung obere Kante an obere Kante',   supertip='Ausrichtung der oberen Kante an der oberen Kante des Mastershapes.'),
        bkt.ribbon.Button(id='arrange_position_bottom',   on_action=bkt.Callback(arrange_adv_position.arrange_bottom_at_bottom),   get_enabled=bkt.Callback(arrange_adv_position.enabled), image='arrange_position_bottom',     label="Unten an Unten",   show_label=False, screentip='Ausrichtung untere Kante an untere Kante', supertip='Ausrichtung der unteren Kante an der unteren Kante des Mastershapes.'),
        bkt.ribbon.Button(id='arrange_position_vmiddle',  on_action=bkt.Callback(arrange_adv_position.arrange_vmiddle_at_vmiddle), get_enabled=bkt.Callback(arrange_adv_position.enabled), image='arrange_position_vmiddle',    label="Mitte an Mitte",   show_label=False, screentip='Ausrichtung Shapemitte an Shapemitte',     supertip='Ausrichtung der Shapemitte an der Shapemitte des Mastershapes.'),
        bkt.ribbon.Separator(),

        #DOCK
        bkt.ribbon.Box(box_style="horizontal", children=[
            bkt.ribbon.Button(id='arrange_dock_left',    on_action=bkt.Callback(arrange_adv_position.arrange_left_at_right),      get_enabled=bkt.Callback(arrange_adv_position.enabled), image='arrange_dock_left',      label="Links an Rechts",  show_label=False, screentip='Ausrichtung linke Kante an rechte Kante',  supertip='Ausrichtung der linken Kante an der rechten Kante des Mastershapes.'),
            bkt.ribbon.Button(id='arrange_dock_right',   on_action=bkt.Callback(arrange_adv_position.arrange_right_at_left),      get_enabled=bkt.Callback(arrange_adv_position.enabled), image='arrange_dock_right',     label="Rechts an Links",  show_label=False, screentip='Ausrichtung rechte Kante an linke Kante',  supertip='Ausrichtung der rechten Kante an der linken Kante des Mastershapes.'),
            bkt.ribbon.Button(id='arrange_dock_bottom',  on_action=bkt.Callback(arrange_adv_position.arrange_bottom_at_top),      get_enabled=bkt.Callback(arrange_adv_position.enabled), image='arrange_dock_top',       label="Unten an Oben",    show_label=False, screentip='Ausrichtung untere Kante an obere Kante',  supertip='Ausrichtung der unteren Kante an der oberen Kante des Mastershapes.'),
            bkt.ribbon.Button(id='arrange_dock_top',     on_action=bkt.Callback(arrange_adv_position.arrange_top_at_bottom),      get_enabled=bkt.Callback(arrange_adv_position.enabled), image='arrange_dock_bottom',    label="Oben an Unten",    show_label=False, screentip='Ausrichtung obere Kante an untere Kante',  supertip='Ausrichtung der oberen Kante an der unteren Kante des Mastershapes.'),
        ]),

        #STRETCH
        bkt.ribbon.Box(box_style="horizontal", children=[
            bkt.ribbon.Button(id='arrange_stretch_left',   on_action=bkt.Callback(arrange_adv_size.arrange_left_at_left),       get_enabled=bkt.Callback(arrange_adv_size.enabled), image='arrange_stretch_left',        label="Links an Links",   show_label=False, screentip='Ausrichtung linke Kante an linke Kante',   supertip='Ausrichtung der linken Kante an der linken Kante des Mastershapes.'),
            bkt.ribbon.Button(id='arrange_stretch_right',  on_action=bkt.Callback(arrange_adv_size.arrange_right_at_right),     get_enabled=bkt.Callback(arrange_adv_size.enabled), image='arrange_stretch_right',       label="Rechts an Rechts", show_label=False, screentip='Ausrichtung rechte Kante an rechte Kante', supertip='Ausrichtung der rechten Kante an der rechten Kante des Mastershapes.'),

            bkt.ribbon.Button(id='arrange_stretch_top',    on_action=bkt.Callback(arrange_adv_size.arrange_top_at_top),         get_enabled=bkt.Callback(arrange_adv_size.enabled), image='arrange_stretch_top',         label="Oben an oben",     show_label=False, screentip='Ausrichtung obere Kante an obere Kante',   supertip='Ausrichtung der oberen Kante an der oberen Kante des Mastershapes.'),
            bkt.ribbon.Button(id='arrange_stretch_bottom', on_action=bkt.Callback(arrange_adv_size.arrange_bottom_at_bottom),   get_enabled=bkt.Callback(arrange_adv_size.enabled), image='arrange_stretch_bottom',      label="Unten an Unten",   show_label=False, screentip='Ausrichtung untere Kante an untere Kante', supertip='Ausrichtung der unteren Kante an der unteren Kante des Mastershapes.'),
        ]),

        #FILL
        bkt.ribbon.Box(box_style="horizontal", children=[
            bkt.ribbon.Button(id='arrange_fill_left',       on_action=bkt.Callback(arrange_adv_size.arrange_left_at_right),      get_enabled=bkt.Callback(arrange_adv_size.enabled), image='arrange_fill_left',          label="Links an Rechts",  show_label=False, screentip='Ausrichtung linke Kante an rechte Kante',  supertip='Ausrichtung der linken Kante an der rechten Kante des Mastershapes.'),
            bkt.ribbon.Button(id='arrange_fill_right',      on_action=bkt.Callback(arrange_adv_size.arrange_right_at_left),      get_enabled=bkt.Callback(arrange_adv_size.enabled), image='arrange_fill_right',         label="Rechts an Links",  show_label=False, screentip='Ausrichtung rechte Kante an linke Kante',  supertip='Ausrichtung der rechten Kante an der linken Kante des Mastershapes.'),
            bkt.ribbon.Button(id='arrange_fill_bottom',     on_action=bkt.Callback(arrange_adv_size.arrange_bottom_at_top),      get_enabled=bkt.Callback(arrange_adv_size.enabled), image='arrange_fill_bottom',        label="Unten an Oben",    show_label=False, screentip='Ausrichtung untere Kante an obere Kante',  supertip='Ausrichtung der unteren Kante an der oberen Kante des Mastershapes.'),
            bkt.ribbon.Button(id='arrange_fill_top',        on_action=bkt.Callback(arrange_adv_size.arrange_top_at_bottom),      get_enabled=bkt.Callback(arrange_adv_size.enabled), image='arrange_fill_top',           label="Oben an Unten",    show_label=False, screentip='Ausrichtung obere Kante an untere Kante',  supertip='Ausrichtung der oberen Kante an der unteren Kante des Mastershapes.'),
        ]),

    ]
)




class TableArrange(object):
    ARRANGE_HAUTO = -1 #auto: for pars=none, for table=center
    ARRANGE_HNONE = 0
    ARRANGE_LEFT = 1
    ARRANGE_HCENTER = 2
    ARRANGE_RIGHT = 3

    ARRANGE_VAUTO = -1 #auto: for pars=line-center, for table=center
    ARRANGE_VNONE = 0
    ARRANGE_TOP = 1
    ARRANGE_VCENTER = 2
    ARRANGE_BOTTOM = 3
    ARRANGE_LCENTER = 4 #center on first line
    
    horizontal_arrangement = ARRANGE_HAUTO
    vertical_arrangement = ARRANGE_VAUTO
    
    
    @classmethod
    def arrange_shapes_on_table(cls, shapes, table):
        # all width/height of table cols/rows
        widths = [col.width for col in table.table.columns]
        heights = [row.height for row in table.table.rows]
        
        # left/top of cols/rows
        agg_widths = [ sum(widths[0:i])  for i in range(len(widths)+1)]
        agg_heights = [ sum(heights[0:i])  for i in range(len(heights)+1)]
        col_lefts = [table.x +w  for w in agg_widths]
        row_tops = [table.y +h for h in agg_heights]
        
        for shape in shapes:
            shape_midpoint = [ shape.center_x, shape.center_y ]
            
            try:
                # determine target-cell and target-rect
                try:
                    col = [c - shape_midpoint[0] >=0 for c in col_lefts].index(True)
                    row = [r - shape_midpoint[1] >=0 for r in row_tops].index(True)
                except ValueError:
                    #no col/row found, shape outside bottom-right boundaries of table
                    continue
                if col == 0 or row == 0:
                    #shape outside top-left boundaries of table
                    continue
                target_rect = [ col_lefts[col-1], row_tops[row-1], widths[col-1], heights[row-1]   ]

                # determine target-midpoint from arrangement-setting
                target_midpoint = [0,0]
                if cls.horizontal_arrangement == cls.ARRANGE_HNONE:
                    target_midpoint[0] = shape_midpoint[0] #no change in position
                elif cls.horizontal_arrangement == cls.ARRANGE_LEFT:
                    target_midpoint[0] = target_rect[0] + shape.width/2
                elif cls.horizontal_arrangement == cls.ARRANGE_RIGHT:
                    target_midpoint[0] = target_rect[0] + target_rect[2] - shape.width/2
                else: # ARRANGE_HCENTER or ARRANGE_HAUTO
                    target_midpoint[0] = target_rect[0] + target_rect[2]/2

                if cls.vertical_arrangement == cls.ARRANGE_VNONE:
                    target_midpoint[1] = shape_midpoint[1] #no change in position
                elif cls.vertical_arrangement == cls.ARRANGE_TOP:
                    target_midpoint[1] = target_rect[1] + shape.height/2
                elif cls.vertical_arrangement == cls.ARRANGE_BOTTOM:
                    target_midpoint[1] = target_rect[1] + target_rect[3] - shape.height/2
                elif cls.vertical_arrangement == cls.ARRANGE_LCENTER:
                    textframe = table.table.cell(row, col).shape.textframe2
                    if textframe.HasText == -1: #cell has text
                        first_line_height = textframe.textrange.lines[1].boundheight
                    else: #cell has no text
                        first_line_height = textframe.textrange.boundheight
                    target_midpoint[1] = target_rect[1] + first_line_height/2
                else: # ARRANGE_VCENTER or ARRANGE_VAUTO
                    target_midpoint[1] = target_rect[1] + target_rect[3]/2
                
                # move shape
                shape.x += target_midpoint[0] - shape_midpoint[0]
                shape.y  += target_midpoint[1] - shape_midpoint[1]
            
            except:
                bkt.helpers.exception_as_message()
    
    @classmethod
    def arrange_table_shapes(cls, shapes):
        shapes = pplib.wrap_shapes(shapes)
        # determine table in shapes-list
        # tables = [s for s in shapes if s.Type == pplib.MsoShapeType['msoTable']]
        tables = [s for s in shapes if s.HasTable == -1]
        shapes = [s for s in shapes if s not in tables]
        # for each table call arrange_shapes_on_table width shapes
        for table in tables:
            cls.arrange_shapes_on_table(shapes, table)
    
    
    @classmethod
    def arrange_shapes_on_paragraph(cls, shapes, textshape):
        try:
            paragraphs = [textshape.TextFrame2.TextRange.Paragraphs(idx) for idx in range(1,textshape.TextFrame2.TextRange.Paragraphs().Count+1)]
            paragraph_v_bounds = [ p.boundtop for p in paragraphs]
            # bound below paragraphs
            paragraph_v_bounds.append( paragraphs[-1].boundtop+paragraphs[-1].boundheight  )

            # FIXME: find solution that works with columns within textframe
            # if textshape.TextFrame2.Column.Number > 1:
            #     colstart = textshape.left+textshape.TextFrame2.MarginLeft
            #     colsize  = (textshape.width - textshape.TextFrame2.MarginLeft - textshape.TextFrame2.MarginRight) / textshape.TextFrame2.Column.Number
            #     cols = [ colstart + x*colsize for x in range(1,textshape.TextFrame2.Column.Number)]
            
            #list of paragraphs already used to assign a shape
            par_idx_used = []
            
            for shape in shapes:
                shape_midpoint = [ shape.center_x, shape.center_y ]
                # par_idx in [1...len(paragraphs)]
                try:
                    par_idx = [ v_bound - shape_midpoint[1]  >=0 for v_bound in paragraph_v_bounds].index(True)
                except ValueError:
                    # shape is below every paragraph, use last
                    par_idx = len(paragraphs)
                if par_idx ==0:
                    # shape is over every paragraph, use first
                    par_idx = 1

                #prevent multiple shapes on the same paragraph, e.g. on top of each other
                if par_idx in par_idx_used:
                    #FIXME: better behavior to increase par_idx until "empty slot" is found?
                    continue
                else:
                    par_idx_used.append(par_idx)
                    
                #FIXME: SpaceBefore/After give lines (not points) when LineRuleBefore/After, but should be a special case as this cannot be defined via GUI

                # target_box: left , top , width, height
                target_box = [ paragraphs[par_idx-1].boundleft, paragraphs[par_idx-1].boundtop + paragraphs[par_idx-1].paragraphFormat.SpaceBefore, 
                    paragraphs[par_idx-1].boundwidth, paragraphs[par_idx-1].boundheight - paragraphs[par_idx-1].paragraphFormat.SpaceBefore - paragraphs[par_idx-1].paragraphFormat.SpaceAfter   ]

                #alignment on first line
                align_lcenter = False
                lines = paragraphs[par_idx-1].Lines
                if cls.vertical_arrangement in [cls.ARRANGE_LCENTER, cls.ARRANGE_VAUTO] and lines().Count > 0:
                    align_lcenter = True
                    target_box[3] = lines[1].boundheight - paragraphs[par_idx-1].paragraphFormat.SpaceBefore - paragraphs[par_idx-1].paragraphFormat.SpaceAfter
                    
                    if paragraphs[par_idx-1].paragraphFormat.LineRuleWithin == -1: #msoTrue
                        linefactor = paragraphs[par_idx-1].paragraphFormat.SpaceWithin
                    else:
                        linefactor = paragraphs[par_idx-1].paragraphFormat.SpaceWithin / 1.2 / lines[1].Font.Size #1.2 seems to be an arbitrary constant of ppt


                if par_idx == 1:
                    # this is the first paragraph
                    # here boundtop and boundheight do not consider SpaceBefore
                    target_box[1] = target_box[1] - paragraphs[0].paragraphFormat.SpaceBefore
                    target_box[3] = target_box[3] + paragraphs[0].paragraphFormat.SpaceBefore


                #the following line doesnt always works (if last paragraph is single line and has space after):
                #is_last_par = textshape.TextFrame2.TextRange.Boundtop + textshape.TextFrame2.TextRange.Boundheight == paragraphs[par_idx-1].boundtop + paragraphs[par_idx-1].boundheight
                is_last_par = par_idx == len(paragraphs) #TESTME: really this simple?
                if is_last_par or (align_lcenter and lines().Count > 1):
                    # this is the last paragraph containing text OR a paragraph with more than one line and center to line is selected
                    # here boundheight does not include SpaceAfter / boundheight must be removed from height by adding it again
                    target_box[3] = target_box[3] + paragraphs[par_idx-1].paragraphFormat.SpaceAfter
                    
                    if lines().Count == 1:
                        # this is the last paragraph containing exactly one line
                        # in this case the boundwidth is a little bit smaller as new line character is not counted
                        target_box[2] = target_box[2] + 5 #5 is randomly chosen here as it is not possible to determine width of new line character
                    
                        # if align_lcenter:
                        #     # this is the last paragraph containing exactly one line and center to line is selected
                        #     # SpaceWithin under the text is not considered in some rare cases (e.g. when font size changed after line spacing is defined)
                        #     target_box[3] = target_box[3] * 1.2


                # determine target-midpoint from arrangement-setting
                target_midpoint = [0,0]
                if cls.horizontal_arrangement in [cls.ARRANGE_HNONE, cls.ARRANGE_HAUTO]:
                    target_midpoint[0] = shape_midpoint[0] #no change in position
                elif cls.horizontal_arrangement == cls.ARRANGE_LEFT:
                    target_midpoint[0] = target_box[0] - shape.width - 5 + shape.width/2 #5 is randomly chosen to give a little bit of spacing
                elif cls.horizontal_arrangement == cls.ARRANGE_RIGHT:
                    target_midpoint[0] = target_box[0] + target_box[2] + shape.width - shape.width/2
                else: # ARRANGE_HCENTER
                    target_midpoint[0] = target_box[0] + target_box[2]/2

                if cls.vertical_arrangement == cls.ARRANGE_VNONE:
                    target_midpoint[1] = shape_midpoint[1] #no change in position
                elif cls.vertical_arrangement == cls.ARRANGE_TOP:
                    target_midpoint[1] = target_box[1] + shape.height/2
                elif cls.vertical_arrangement == cls.ARRANGE_BOTTOM:
                    target_midpoint[1] = target_box[1] + target_box[3] - shape.height/2
                elif align_lcenter: #includes ARRANGE_VAUTO
                    linefactor = 1 + (linefactor-1)/10 if linefactor < 2 else 1 + linefactor/10
                    target_midpoint[1] = target_box[1] + target_box[3]/linefactor/2 + (target_box[3] - target_box[3]/linefactor)
                else: # ARRANGE_VCENTER
                    target_midpoint[1] = target_box[1] + target_box[3]/2
            
                shape.x += target_midpoint[0] - shape_midpoint[0]
                shape.y += target_midpoint[1] - shape_midpoint[1]
        except:
            bkt.helpers.exception_as_message()
            
    
    @classmethod
    def arrange_paragraph_shapes(cls, shapes):
        shapes = pplib.wrap_shapes(shapes)
        for master in shapes:
            # master has no paragraphs -> skip master
            if master.HasTextFrame != -1 or master.TextFrame2.TextRange.Paragraphs().Count == 0:
                continue

            # shapes in selection, complete in master
            #inner_shapes = [ s  for s in shapes if (s!=master and s.left>=master.left and s.top>=master.top and s.top+s.height<=master.top+master.height and s.left+s.width<=master.left+master.width )]
            # shapes in selection, being smaller and midpoint in master
            inner_shapes = [
                s  for s in shapes
                if s!=master and cls._is_shape_within(master, s)
                ]
            
            if len(inner_shapes) > 0:
                cls.arrange_shapes_on_paragraph(inner_shapes, master)
    
    @classmethod
    def arrange_shapes_on_shapes(cls, shapes, background_shapes):
        for shape in shapes:
            if cls.horizontal_arrangement == cls.ARRANGE_LEFT:
                shape.x = max(s.x for s in background_shapes)
            elif cls.horizontal_arrangement == cls.ARRANGE_RIGHT:
                shape.x1 = min(s.x1 for s in background_shapes)
            # elif cls.horizontal_arrangement == cls.ARRANGE_HCENTER or (cls.horizontal_arrangement == cls.ARRANGE_HAUTO and background_shape.height >= background_shape.width):
            elif cls.horizontal_arrangement in [cls.ARRANGE_HCENTER, cls.ARRANGE_HAUTO]:
                if cls.horizontal_arrangement != cls.ARRANGE_HAUTO or len(background_shapes) > 1 or background_shapes[0].height >= background_shapes[0].width:
                    shape.center_x = min(background_shapes, key=lambda s: s.width).center_x

            if cls.vertical_arrangement == cls.ARRANGE_TOP:
                shape.y = max(s.y for s in background_shapes)
            elif cls.vertical_arrangement == cls.ARRANGE_BOTTOM:
                shape.y1 = min(s.y1 for s in background_shapes)
            # elif cls.vertical_arrangement in [cls.ARRANGE_VCENTER,cls.ARRANGE_LCENTER] or (cls.vertical_arrangement == cls.ARRANGE_VAUTO and background_shape.width >= background_shape.height):
            elif cls.vertical_arrangement in [cls.ARRANGE_VCENTER, cls.ARRANGE_LCENTER, cls.ARRANGE_VAUTO]:
                if cls.vertical_arrangement != cls.ARRANGE_VAUTO or len(background_shapes) > 1 or background_shapes[0].width >= background_shapes[0].height:
                    shape.center_y = min(background_shapes, key=lambda s: s.height).center_y


    @classmethod
    def arrange_shapes_shapes(cls, shapes):
        shapes = pplib.wrap_shapes(shapes)
        for child in shapes:
            background_shapes = [
                s for s in shapes 
                if s!=child and s.ZOrderPosition < child.ZOrderPosition and
                cls._is_shape_within(s, child)
                ]

            if len(background_shapes) > 0:
                cls.arrange_shapes_on_shapes([child], background_shapes)
        # for master in shapes:
        #     inner_shapes = [
        #         s for s in shapes 
        #         if s!=master and s.ZOrderPosition > master.ZOrderPosition and
        #         cls._is_shape_within(master, s)
        #         ]

        #     if len(inner_shapes) > 0:
        #         cls.arrange_shapes_on_shapes(inner_shapes, [master])


    @classmethod
    def _is_shape_within(cls, outer_s, inner_s):
        #test if center point of inner_s is within bounds of outer_s
        return inner_s.width<=outer_s.width and inner_s.height<=outer_s.height and outer_s.x <= inner_s.center_x <= outer_s.x1 and outer_s.y <= inner_s.center_y <= outer_s.y1


    @classmethod
    def arrange_overlay_shapes(cls, shapes):
        from itertools import chain #chain allows to concatenate lists and return a generator

        shapes = pplib.wrap_shapes(shapes) #all functions support/require wrapped shapes

        table_shapes = []
        table_childs = []
        par_shapes = []
        par_childs = []
        remaining_shapes = []
        shape_shapes = []

        #step 1: seperate tables, paragraph shapes, and all the remaining shapes
        for s in shapes:
            #test if shape is a table
            if s.HasTable == -1:
                table_shapes.append(s)
            
            #test if shape has a paragraph
            elif s.HasTextFrame == -1 and s.TextFrame2.TextRange.Paragraphs().Count > 1:
                par_shapes.append(s)

            else:
                remaining_shapes.append(s)
        
        #step 2: from remaining shapes, find shapes within tables and paragraph shapes from step 1
        for s in remaining_shapes:
            #get all shapes within tables
            if any(cls._is_shape_within(o, s) for o in table_shapes):
                table_childs.append(s)
            
            #get all shapes within paragraphs
            elif any(cls._is_shape_within(o, s) for o in par_shapes):
                par_childs.append(s)
            
            else:
                shape_shapes.append(s)

        #arrange on table
        cls.arrange_table_shapes(chain(table_shapes, table_childs))
        #arrange on paragraph
        cls.arrange_paragraph_shapes(chain(par_shapes, par_childs))
        #arrange shapes on shapes
        cls.arrange_shapes_shapes(shape_shapes)


class TALocPin(pplib.LocPin):
    def __init__(self):
        super(TALocPin, self).__init__()
        self.locpins = [
            (TableArrange.ARRANGE_VAUTO,  TableArrange.ARRANGE_HAUTO), (TableArrange.ARRANGE_VNONE,  TableArrange.ARRANGE_LEFT), (TableArrange.ARRANGE_VNONE,  TableArrange.ARRANGE_HCENTER), (TableArrange.ARRANGE_VNONE,  TableArrange.ARRANGE_RIGHT),
            (TableArrange.ARRANGE_TOP,    TableArrange.ARRANGE_HNONE), (TableArrange.ARRANGE_TOP,    TableArrange.ARRANGE_LEFT), (TableArrange.ARRANGE_TOP,    TableArrange.ARRANGE_HCENTER), (TableArrange.ARRANGE_TOP,    TableArrange.ARRANGE_RIGHT),
            (TableArrange.ARRANGE_LCENTER,TableArrange.ARRANGE_HNONE), (TableArrange.ARRANGE_LCENTER,TableArrange.ARRANGE_LEFT), (TableArrange.ARRANGE_LCENTER,TableArrange.ARRANGE_HCENTER), (TableArrange.ARRANGE_LCENTER,TableArrange.ARRANGE_RIGHT),
            (TableArrange.ARRANGE_VCENTER,TableArrange.ARRANGE_HNONE), (TableArrange.ARRANGE_VCENTER,TableArrange.ARRANGE_LEFT), (TableArrange.ARRANGE_VCENTER,TableArrange.ARRANGE_HCENTER), (TableArrange.ARRANGE_VCENTER,TableArrange.ARRANGE_RIGHT),
            (TableArrange.ARRANGE_BOTTOM, TableArrange.ARRANGE_HNONE), (TableArrange.ARRANGE_BOTTOM, TableArrange.ARRANGE_LEFT), (TableArrange.ARRANGE_BOTTOM, TableArrange.ARRANGE_HCENTER), (TableArrange.ARRANGE_BOTTOM, TableArrange.ARRANGE_RIGHT),
        ]

class TAGallery(bkt.ribbon.Gallery):
    def __init__(self, **kwargs):
        self.locpin = TALocPin()
        self.items = [
            ("fix_locpin_auto", "Auto", "Shapes werden bei Shape-Anordnung in:\n\n- Tabellenzellen horizontal und vertikal zentriert angeordnet,\n- Shapes horizontal zentriert sofern Shape höher ist als breit und vertikal zentriert sofern Shape breiter ist als hoch,\n- Textabsätzen vertikal zentriert von der ersten Zeile und horizontal nicht verändert."),
            ("fix_locpin_l", "Links", "Shapes werden links angeordnet und vertikal nicht verändert."),
            ("fix_locpin_c", "Zentriert", "Shapes werden zentriert angeordnet und vertikal nicht verändert."),
            ("fix_locpin_r", "Rechts", "Shapes werden rechts angeordnet und vertikal nicht verändert."),

            ("fix_locpin_t", "Oben", "Shapes werden oben angeordnet und horizontal nicht verändert."),
            ("fix_locpin_tl", "Oben-links", "Shapes werden oben links angeordnet."),
            ("fix_locpin_tm", "Oben-mitte", "Shapes werden oben zentriert angeordnet."),
            ("fix_locpin_tr", "Oben-rechts", "Shapes werden oben rechts angeordnet."),

            ("fix_locpin_m_line", "Mitte 1. Zeile", "Shapes werden vertikal zentriert von dem Text in der ersten Zeile angeordnet und horizontal nicht verändert."),
            ("fix_locpin_ml_line", "Mitte 1. Zeile-links", "Shapes werden vertikal zentriert von dem Text in der ersten Zeile und horizontal links angeordnet."),
            ("fix_locpin_mm_line", "Mitte 1. Zeile-mitte", "Shapes werden vertikal zentriert von dem Text in der ersten Zeile und horizontal zentriert angeordnet."),
            ("fix_locpin_mr_line", "Mitte 1. Zeile-rechts", "Shapes werden vertikal zentriert von dem Text in der ersten Zeile und horizontal rechts angeordnet."),

            ("fix_locpin_m", "Mitte", "Shapes werden vertikal zentriert angeordnet und horizontal nicht verändert."),
            ("fix_locpin_ml", "Mitte-links", "Shapes werden mittig links angeordnet."),
            ("fix_locpin_mm", "Mitte-mitte", "Shapes werden mittig zentriert angeordnet."),
            ("fix_locpin_mr", "Mitte-rechts", "Shapes werden mittig rechts angeordnet."),

            ("fix_locpin_b", "Unten", "Shapes werden unten angeordnet und horizontal nicht verändert."),
            ("fix_locpin_bl", "Unten-links", "Shapes werden unten links angeordnet."),
            ("fix_locpin_bm", "Unten-mitte", "Shapes werden unten zentriert angeordnet."),
            ("fix_locpin_br", "Unten-rechts", "Shapes werden unten rechts angeordnet."), 
        ]
        
        my_kwargs = dict(
            columns="4",
            item_height="32",
            item_width="32",
            on_action_indexed  = bkt.Callback(self.locpin_on_action_indexed),
            get_selected_item_index = bkt.Callback(lambda: self.locpin.index),
            get_image = bkt.Callback(self.locpin_get_image, context=True),
            get_item_count = bkt.Callback(lambda: len(self.items)),
            # get_item_label = bkt.Callback(lambda index: self.items[index][1]),
            get_item_image = bkt.Callback(self.locpin_get_image, context=True),
            get_item_screentip = bkt.Callback(lambda index: self.items[index][1]),
            get_item_supertip = bkt.Callback(lambda index: self.items[index][2]),
            # children = [
            #     bkt.ribbon.Item(image=gal_item[0], screentip=gal_item[1], supertip=gal_item[2])
            #     for gal_item in self.items
            # ]
        )
        my_kwargs.update(kwargs)
        
        super(TAGallery, self).__init__(**my_kwargs)

    def locpin_on_action_indexed(self, selected_item, index):
        self.locpin.index = index
        TableArrange.vertical_arrangement, TableArrange.horizontal_arrangement = self.locpin.fixation
    
    def locpin_get_image(self, context, index=None):
        if index is None:
            return context.python_addin.load_image(self.items[self.locpin.index][0])
        else:
            return context.python_addin.load_image(self.items[index][0])



tablearrange_button = bkt.ribbon.SplitButton(
    show_label=False,
    get_enabled = bkt.apps.ppt_shapes_min2_selected,
    children=[
        bkt.ribbon.Button(
            id = 'table-shapes',
            label="Auf Tabelle/Absatz/Shapes anordnen",
            image='arrange_shape_table',
            screentip="Shape-Objekte in Tabellen/Absätzen/Shapes anordnen",
            supertip="Bei Markierung von Tabellen und Shapes:\nOrdne jedes Shape, das über einer (ebenfalls selektierten) Tabelle liegt, innerhalb der nächst-liegenden Tabellenzelle an. Die Zelle wird anhand des Shape-Mittelpunkts bestimmt.\n\nBei Markierung von Shapes:\nOrdne jedes Shape, das vollständig über einem anderen (ebenfalls selektierten) Shapes (=Mastershape) liegt, in dem nächstliegenden Textabsatz im Mastershape an. Sind weniger as 2 Textabsätze vorhanden, wird innerhalb des gesamten Shapes angeordnet.",
            on_action=bkt.Callback(TableArrange.arrange_overlay_shapes, shapes=True, shapes_min=2),
            # get_enabled = bkt.apps.ppt_shapes_min2_selected,
        ),
        bkt.ribbon.Menu(label="Einstellungen für Anordnen auf Tabelle/Absatz", item_size="large", children=[
            bkt.ribbon.MenuSeparator(title="Shapes anordnen"),
            bkt.ribbon.Button(
                id = 'table-shapes2',
                label="Automatisch anordnen",
                image='arrange_shape_table_auto',
                screentip="Shape-Objekte in Tabellenzellen/Absätzen/Shapes anordnen",
                description="Automatische Auswahl der Anordnungsfunktionen (Tabellenzellen/Absätzen/Shapes)",
                on_action=bkt.Callback(TableArrange.arrange_overlay_shapes, shapes=True, shapes_min=2),
                # get_enabled = bkt.apps.ppt_shapes_min2_selected,
            ),
            TAGallery(
                id="arrange_shape_table_mode",
                label="Anordnung anpassen",
                screentip="Shape-Position beim Anordnen festlegen.",
                description="Auswahl der horizontalen und vertikalen Anordnung."
            ),
            bkt.ribbon.MenuSeparator(title="Manuelle Auswahl"),
            bkt.ribbon.Button(
                id="arrange_on_table",
                image="arrange_on_table",
                label="Auf Tabellenzellen anordnen",
                screentip="Shape-Objekte in Tabellenzellen anordnen",
                description="Der Mittelpunkt der Shapes muss innerhalb der Tabellenzelle liegen für die automatische Anordnung.",
                on_action=bkt.Callback(TableArrange.arrange_table_shapes, shapes=True, shapes_min=2),
                # get_enabled = bkt.apps.ppt_shapes_min2_selected,
            ),
            bkt.ribbon.Button(
                id="arrange_on_paragraph",
                image="arrange_on_paragraph",
                label="Auf Text-Absätzen anordnen",
                screentip="Shape-Objekte in Textabsätzen anordnen",
                description="Der Mittelpunkt der Shapes muss innerhalb des Text-Shapes und des jeweiligen Absatzes liegen für die automatische Anordnung.",
                on_action=bkt.Callback(TableArrange.arrange_paragraph_shapes, shapes=True, shapes_min=2),
                # get_enabled = bkt.apps.ppt_shapes_min2_selected,
            ),
            bkt.ribbon.Button(
                id="arrange_on_shapes",
                image="arrange_on_shapes",
                label="Auf Hintergrund-Shapes anordnen",
                screentip="Shape-Objekte in untergeordneten Shapes anordnen",
                description="Der Mittelpunkt der Shapes muss innerhalb der darunterliegenden Shapes liegen für die automatische Anordnung.",
                on_action=bkt.Callback(TableArrange.arrange_shapes_shapes, shapes=True, shapes_min=2),
                # get_enabled = bkt.apps.ppt_shapes_min2_selected,
            ),

            # bkt.ribbon.MenuSeparator(title="Horizontale Anordnung"),
            # bkt.ribbon.CheckBox(
            #     id="arrange_shape_table_hauto",
            #     label="Automatisch",
            #     screentip="Shape-Position automatisch",
            #     supertip="Shapes werden bei Shape-Anordnung in:\n\n- Tabellenzellen horizontal zentriert angeordnet,\n- Shapes horizontal zentriert sofern Shape höher ist als breit,\n- Textabsätzen nicht angeordnet.",
            #     on_toggle_action=bkt.Callback(lambda pressed: setattr(TableArrange, 'horizontal_arrangement', TableArrange.ARRANGE_HAUTO) ),
            #     get_pressed=bkt.Callback(lambda : TableArrange.horizontal_arrangement == TableArrange.ARRANGE_HAUTO)
            # ),
            # bkt.ribbon.CheckBox(
            #     id="arrange_shape_table_hnone",
            #     label="Keine Änderung",
            #     screentip="Shape-Position nicht ändern",
            #     supertip="Shapes werden bei Shape-Anordnung in Tabellenzellen/Textabsätzen/Shapes horizontal nicht angeordnet.",
            #     on_toggle_action=bkt.Callback(lambda pressed: setattr(TableArrange, 'horizontal_arrangement', TableArrange.ARRANGE_HNONE) ),
            #     get_pressed=bkt.Callback(lambda : TableArrange.horizontal_arrangement == TableArrange.ARRANGE_HNONE)
            # ),
            # bkt.ribbon.MenuSeparator(),
            # bkt.ribbon.CheckBox(
            #     id="arrange_shape_table_left",
            #     label="Links",
            #     screentip="Shapes links anordnen",
            #     supertip="Shapes werden bei Shape-Anordnung in Tabellenzellen/Textabsätzen/Shapes links angeordnet.",
            #     on_toggle_action=bkt.Callback(lambda pressed: setattr(TableArrange, 'horizontal_arrangement', TableArrange.ARRANGE_LEFT) ),
            #     get_pressed=bkt.Callback(lambda : TableArrange.horizontal_arrangement == TableArrange.ARRANGE_LEFT)
            # ),
            # bkt.ribbon.CheckBox(
            #     id="arrange_shape_table_center",
            #     label="Zentriert",
            #     screentip="Shapes zentriert anordnen",
            #     supertip="Shapes werden bei Shape-Anordnung in Tabellenzellen/Textabsätzen/Shapes horizontal zentriert angeordnet.",
            #     on_toggle_action=bkt.Callback(lambda pressed: setattr(TableArrange, 'horizontal_arrangement', TableArrange.ARRANGE_HCENTER) ),
            #     get_pressed=bkt.Callback(lambda : TableArrange.horizontal_arrangement == TableArrange.ARRANGE_HCENTER)
            # ),
            # bkt.ribbon.CheckBox(
            #     id="arrange_shape_table_right",
            #     label="Rechts",
            #     screentip="Shapes rechts anordnen",
            #     supertip="Shapes werden bei Shape-Anordnung in Tabellenzellen/Textabsätzen/Shapes rechts angeordnet.",
            #     on_toggle_action=bkt.Callback(lambda pressed: setattr(TableArrange, 'horizontal_arrangement', TableArrange.ARRANGE_RIGHT) ),
            #     get_pressed=bkt.Callback(lambda : TableArrange.horizontal_arrangement == TableArrange.ARRANGE_RIGHT)
            # ),
            # bkt.ribbon.MenuSeparator(title="Vertikale Anordnung"),
            # bkt.ribbon.CheckBox(
            #     id="arrange_shape_table_vauto",
            #     label="Automatisch",
            #     screentip="Shape-Position automatisch",
            #     supertip="Shapes werden bei Shape-Anordnung in:\n\n- Tabellenzellen vertikal zentriert angeordnet,\n- Shapes vertikal zentriert sofern Shape breiter ist als hoch,\n- Textabsätzen vertikal zentriert von erster Zeile angeordnet.",
            #     on_toggle_action=bkt.Callback(lambda pressed: setattr(TableArrange, 'vertical_arrangement', TableArrange.ARRANGE_VAUTO) ),
            #     get_pressed=bkt.Callback(lambda : TableArrange.vertical_arrangement == TableArrange.ARRANGE_VAUTO)
            # ),
            # bkt.ribbon.CheckBox(
            #     id="arrange_shape_table_vnone",
            #     label="Keine Änderung",
            #     screentip="Shape-Position nicht ändern",
            #     supertip="Shapes werden bei Shape-Anordnung in Tabellenzellen/Textabsätzen/Shapes vertikal nicht angeordnet.",
            #     on_toggle_action=bkt.Callback(lambda pressed: setattr(TableArrange, 'vertical_arrangement', TableArrange.ARRANGE_VNONE) ),
            #     get_pressed=bkt.Callback(lambda : TableArrange.vertical_arrangement == TableArrange.ARRANGE_VNONE)
            # ),
            # bkt.ribbon.MenuSeparator(),
            # bkt.ribbon.CheckBox(
            #     id="arrange_shape_table_top",
            #     label="Oben",
            #     screentip="Shapes oben anordnen",
            #     supertip="Shapes werden bei Shape-Anordnung in Tabellenzellen/Textabsätzen/Shapes oben angeordnet.",
            #     on_toggle_action=bkt.Callback(lambda pressed: setattr(TableArrange, 'vertical_arrangement', TableArrange.ARRANGE_TOP) ),
            #     get_pressed=bkt.Callback(lambda : TableArrange.vertical_arrangement == TableArrange.ARRANGE_TOP)
            # ),
            # bkt.ribbon.CheckBox(
            #     id="arrange_shape_table_vcenter",
            #     label="Mitte",
            #     screentip="Shapes mittig anordnen",
            #     supertip="Shapes werden bei Shape-Anordnung in Tabellenzellen/Textabsätzen/Shapes vertikal zentriert angeordnet.",
            #     on_toggle_action=bkt.Callback(lambda pressed: setattr(TableArrange, 'vertical_arrangement', TableArrange.ARRANGE_VCENTER) ),
            #     get_pressed=bkt.Callback(lambda : TableArrange.vertical_arrangement == TableArrange.ARRANGE_VCENTER)
            # ),
            # bkt.ribbon.CheckBox(
            #     id="arrange_shape_table_lcenter",
            #     label="Mitte von 1. Zeile",
            #     screentip="Shapes mittig von erster Zeile anordnen",
            #     supertip="Shapes werden bei Shape-Anordnung in Tabellenzellen/Textabsätzen/Shapes vertikal zentriert von dem Text in der ersten Zeile angeordnet.",
            #     on_toggle_action=bkt.Callback(lambda pressed: setattr(TableArrange, 'vertical_arrangement', TableArrange.ARRANGE_LCENTER) ),
            #     get_pressed=bkt.Callback(lambda : TableArrange.vertical_arrangement == TableArrange.ARRANGE_LCENTER)
            # ),
            # bkt.ribbon.CheckBox(
            #     id="arrange_shape_table_bottom",
            #     label="Unten",
            #     screentip="Shapes unten anordnen",
            #     supertip="Shapes werden bei Shape-Anordnung in Tabellenzellen/Textabsätzen/Shapes unten angeordnet.",
            #     on_toggle_action=bkt.Callback(lambda pressed: setattr(TableArrange, 'vertical_arrangement', TableArrange.ARRANGE_BOTTOM) ),
            #     get_pressed=bkt.Callback(lambda : TableArrange.vertical_arrangement == TableArrange.ARRANGE_BOTTOM)
            # )
        ])
    ]
)





class RepositionGallery(pplib.PositionGallery):
    
    def __init__(self, **kwargs):
        super(RepositionGallery, self).__init__(
            on_position_change = bkt.Callback(self.on_position_change),
            **kwargs
        )
    
    def get_item_supertip(self, index):
        return 'Positioniere die ausgewählten Shapes an der angezeigten Position/Größe.\nNur Position ändern [STRG],\nNur Größe ändern [SHIFT]'
    
    def on_position_change(self, target_frame, selection, shapes):
        if len(shapes) > 1:
            shape = selection.ShapeRange.Group()
            self.change_shape_position(shape, target_frame)
            shape.Ungroup().Select()
        else:
            self.change_shape_position(shapes[0], target_frame)
    
    def change_shape_position(self, shape, target_frame):
        # position shape
        # CTRL = position only
        # SHIFT = size only
        
        if not bkt.library.system.get_key_state(bkt.library.system.key_code.SHIFT):
            shape.left   = target_frame.left
            shape.top    = target_frame.top
        
        if not bkt.library.system.get_key_state(bkt.library.system.key_code.CTRL):
            shape.width  = target_frame.width
            shape.height = target_frame.height

 

reposition_gallery = RepositionGallery(id="positions", label="Shapes positionieren")




class ChartShapes(object):
    chart_dimensions = [None, None] #height, width
    plotarea_dimensions = [None, None, None, None] #top, left, height, width

    @classmethod
    def is_chart_shape(cls, shape):
        try:
            # HasChart throws NotImplementedError for SmartArts
            return shape.HasChart == -1
        except NotImplementedError:
            return False
        # return shape.Type == pplib.MsoShapeType['msoChart'] or shape.Type == pplib.MsoShapeType['msoDiagram']

    @classmethod
    def is_paste_enabled(cls, shape):
        return cls.is_chart_shape(shape) and cls.chart_dimensions[0] is not None

    @classmethod
    def copy_dimensions(cls, shape):
        plotarea = shape.Chart.PlotArea
        cls.chart_dimensions = [shape.Height, shape.Width]
        cls.plotarea_dimensions = [plotarea.Top, plotarea.Left, plotarea.Height, plotarea.Width]

    @classmethod
    def paste_dimensions(cls, shape):
        plotarea = shape.Chart.PlotArea
        shape.Height, shape.Width = cls.chart_dimensions
        plotarea.Top, plotarea.Left, plotarea.Height, plotarea.Width = cls.plotarea_dimensions




class PictureFormat(object):
    shape_dimensions = [None, None] #ShapeHeight, ShapeWidth #, ShapeTop, ShapeLeft
    pic_dimensions   = [None, None, None, None] #PictureHeight, PictureWidth, PictureOffsetX, PictureOffsetY

    @classmethod
    def is_pic_shape(cls, shape):
        try:
            return shape.Type == pplib.MsoShapeType["msoPicture"]
        except:
            return False

    @classmethod
    def is_paste_enabled(cls, shape):
        return cls.is_pic_shape(shape) and cls.shape_dimensions[0] is not None

    @classmethod
    def copy_dimensions(cls, shape):
        croparea = shape.PictureFormat.crop
        cls.shape_dimensions = [croparea.ShapeHeight, croparea.ShapeWidth] #, croparea.ShapeTop, croparea.ShapeLeft
        cls.pic_dimensions   = [croparea.PictureHeight, croparea.PictureWidth, croparea.PictureOffsetX, croparea.PictureOffsetY]

    @classmethod
    def paste_dimensions(cls, shape):
        croparea = shape.PictureFormat.crop
        croparea.ShapeHeight, croparea.ShapeWidth = cls.shape_dimensions
        croparea.PictureHeight, croparea.PictureWidth, croparea.PictureOffsetX, croparea.PictureOffsetY = cls.pic_dimensions


class TableFormat(object):
    col_widths = []
    row_heights = []

    @classmethod
    def is_table_shape(cls, shape):
        try:
            # return shape.Type == pplib.MsoShapeType["msoTable"]
            return shape.HasTable == -1 #also covers tables in placeholders
        except:
            return False

    @classmethod
    def is_paste_enabled(cls, shape):
        return cls.is_table_shape(shape) and len(cls.col_widths) > 0

    @classmethod
    def copy_dimensions(cls, shape):
        cls.col_widths = [col.width for col in shape.table.columns]
        cls.row_heights = [row.height for row in shape.table.rows]

    @classmethod
    def paste_dimensions(cls, shape):
        for i, col_width in enumerate(cls.col_widths):
            if shape.table.columns.count-1 < i:
                break
            shape.table.columns(i+1).width = col_width
        
        for i, row_height in enumerate(cls.row_heights):
            if shape.table.rows.count-1 < i:
                break
            shape.table.rows(i+1).height = row_height


class EdgeAutoFixer(object):
    threshold  = bkt.settings.get("toolbox.autofixer_threshold", 0.3)
    groupitems = bkt.settings.get("toolbox.autofixer_groupitems", True)
    order_key  = bkt.settings.get("toolbox.autofixer_order_key", "diagonal-down")

    @classmethod
    def settings_setter(cls, name, value):
        bkt.settings["toolbox.autofixer_"+name] = value
        setattr(cls, name, value)

    @classmethod
    def _iterate_all_shapes(cls, shapes, groupitems=True):
        for shape in shapes:
            #shapes that are rotated other than 0, 90, 180 or 270 degree are excluded
            if shape.rotation % 90 != 0:
                continue
            #connected connectors should not be moved
            if shape.Connector and (shape.ConnectorFormat.BeginConnected or shape.ConnectorFormat.EndConnected):
                continue
            
            if groupitems and shape.Type == 6: #pplib.MsoShapeType['msoGroup']
                for gShape in shape.GroupItems:
                    yield gShape
            else:
                yield shape
    
    @classmethod
    def get_image(cls, context):
        if cls.order_key == "diagonal-down":
            return context.python_addin.load_image("autofixer_dd")
        elif cls.order_key == "top-down":
            return context.python_addin.load_image("autofixer_td")
        else:
            return context.python_addin.load_image("autofixer_lr")

    @classmethod
    def autofix_edges_diagonal_down(cls, shapes):
        cls.settings_setter("order_key", "diagonal-down")
        cls.autofix_edges(shapes)
    
    @classmethod
    def autofix_edges_left_right(cls, shapes):
        cls.settings_setter("order_key", "left-right")
        cls.autofix_edges(shapes)
    
    @classmethod
    def autofix_edges_top_down(cls, shapes):
        cls.settings_setter("order_key", "top-down")
        cls.autofix_edges(shapes)

    @classmethod
    def autofix_edges(cls, shapes):
        cls._autofix_edges(shapes, pplib.cm_to_pt(cls.threshold), cls.groupitems, cls.order_key)
    
    @classmethod
    def _autofix_edges(cls, shapes, threshold=None, groupitems=True, order_key="diagonal-down"):
        #TODO: how to handle locked aspect-ratio and autosize? rotated shapes? ojects with 0 height/width? exclude placeholders?

        threshold = threshold or pplib.cm_to_pt(0.3)

        shapes = pplib.wrap_shapes(cls._iterate_all_shapes(shapes, groupitems))

        # shapes.sort(key=lambda shape: (shape.left, shape.top))
        order_keys = {
            "diagonal-down": [lambda shape: shape.visual_x+shape.visual_y, False],
            "diagonal-up":   [lambda shape: shape.visual_x+shape.visual_y, True],
            "left-right": [lambda shape: (shape.visual_x,shape.visual_y), False],
            "top-down":   [lambda shape: (shape.visual_y,shape.visual_x), False],
            "right-left": [lambda shape: (shape.visual_x,shape.visual_y), True],
            "bottom-up":  [lambda shape: (shape.visual_y,shape.visual_x), True],
        }
        shapes.sort(key=order_keys[order_key][0], reverse=order_keys[order_key][1])

        # logging.debug("Autofix: top-left")
        child_shapes = shapes[:]
        for master_shape in shapes:
            child_shapes.remove(master_shape)
            
            for shape in child_shapes:
                # logging.debug("Autofix1: {} x {}".format(master_shape.name, shape.name))

                #save values before moving shape
                # visual_x1, visual_y1 = shape.visual_x1, shape.visual_y1

                if 1e-4 < abs(shape.visual_x-master_shape.visual_x) < threshold:
                    #resize to left edge
                    delta = master_shape.visual_x - shape.visual_x
                    shape.visual_x += delta
                    shape.visual_width -= delta

                if 1e-4 < abs(shape.visual_y-master_shape.visual_y) < threshold:
                    #resize to top edge
                    delta = master_shape.visual_y - shape.visual_y
                    shape.visual_y += delta
                    shape.visual_height -= delta

                if 1e-4 < abs(shape.visual_x1-master_shape.visual_x1) < threshold:
                    #resize to right edge
                    shape.visual_width = master_shape.visual_x1-shape.visual_x

                if 1e-4 < abs(shape.visual_y1-master_shape.visual_y1) < threshold:
                    #resize to bottom edge
                    shape.visual_height = master_shape.visual_y1-shape.visual_y



class GroupsMore(object):
    @staticmethod
    def add_into_group(shapes):
        if shapes[0].Type == pplib.MsoShapeType['msoGroup']:
            master = pplib.GroupManager(shapes.pop(0))
            master.add_child_items(shapes).select()
        elif shapes[-1].Type == pplib.MsoShapeType['msoGroup']:
            master = pplib.GroupManager(shapes.pop(-1))
            master.add_child_items(shapes).select()
        else:
            pplib.shapes_to_range(shapes).group().select()
    
    @classmethod
    def visible_add_into_group(cls, shapes):
        return len(shapes) > 1 and cls.contains_group(shapes)
    
    @staticmethod
    def contains_group(shapes):
        return any(shp.Type == pplib.MsoShapeType['msoGroup'] for shp in shapes)

    @staticmethod
    def recursive_ungroup(shapes):
        for shape in shapes:
            if shape.Type == pplib.MsoShapeType['msoGroup']:
                grp = pplib.GroupManager(shape)
                grp.recursive_ungroup().select(False)





arrange_group = bkt.ribbon.Group(
    id="bkt_align_group",
    label='Anordnen',
    image_mso='ObjectsArrangeMenu',
    children = [
        bkt.ribbon.Box(box_style="vertical",
            children = [
                equal_height_button,
                equal_width_button,
                swap_button,
            ]
        ),

        bkt.mso.control.ObjectsAlignLeftSmart,
        bkt.mso.control.ObjectsAlignRightSmart,
        bkt.mso.control.ObjectsAlignCenterHorizontalSmart,
        bkt.mso.control.ObjectsAlignTopSmart,
        bkt.mso.control.ObjectsAlignBottomSmart,
        bkt.mso.control.ObjectsAlignMiddleVerticalSmart,
        bkt.mso.control.AlignDistributeHorizontally,
        bkt.mso.control.AlignDistributeVertically,
        
        #bkt.mso.control.ObjectRotateGallery,
        bkt.ribbon.Menu(
            label='Mehr',
            show_label=False,
            image_mso='TableDesign',
            screentip="Weitere Funktionen",
            supertip="Funktionen wie Positionierung, Verlinkte Shapes, ...",
            children=[
                # bkt.ribbon.MenuSeparator(title="Rotation"),
                # bkt.mso.button.ObjectRotateRight90,
                # bkt.mso.button.ObjectRotateLeft90,
                # bkt.mso.button.ObjectFlipHorizontal,
                # bkt.mso.button.ObjectFlipVertical,
                # bkt.ribbon.MenuSeparator(),
                # bkt.mso.button.ObjectRotationOptionsDialog,
                bkt.ribbon.MenuSeparator(title="Positionieren"),
                bkt.mso.control.ObjectRotateGallery,
                reposition_gallery,
                bkt.ribbon.Menu(
                    label='Dimensionen/Größen übertragen',
                    image_mso='PasteWithColumnWidths',
                    children=[
                        bkt.ribbon.Button(
                            id = 'chart_dimensions_copy',
                            label="Diagramm-Dimensionen kopieren",
                            image_mso="ChartPlotArea",
                            screentip="Größe und Position vom Diagrammbereich kopieren",
                            supertip="Kopiert Höhe und Breite des Diagramms sowie Größe und Position der Zeichnungsfläche, um ein anderes Diagramm anzugleichen.",
                            on_action=bkt.Callback(ChartShapes.copy_dimensions, shape=True),
                            get_enabled = bkt.Callback(ChartShapes.is_chart_shape, shape=True),
                        ),
                        bkt.ribbon.Button(
                            id = 'chart_dimensions_paste',
                            label="Diagramm-Dimensionen einfügen",
                            image_mso="PasteWithColumnWidths",
                            screentip="Größe und Position vom Diagrammbereich einfügen",
                            supertip="Überträgt die kopierte Größe und Position des Diagramms bzw. der Zeichnungsfläche auf das ausgewählte Diagramm.",
                            on_action=bkt.Callback(ChartShapes.paste_dimensions, shape=True),
                            get_enabled = bkt.Callback(ChartShapes.is_paste_enabled, shape=True),
                        ),
                        bkt.ribbon.MenuSeparator(),
                        bkt.ribbon.Button(
                            id = 'pic_crop_copy',
                            label="Bild-Zuschnitt kopieren",
                            image_mso="PictureCrop",
                            screentip="Größe und Position des Bildausschnitts kopieren",
                            supertip="Kopiert Höhe und Breite des Ausschnitts bei einem zugeschnittenen Bild, um den Ausschnitt mit einem anderen Bild anzugleichen.",
                            on_action=bkt.Callback(PictureFormat.copy_dimensions, shape=True),
                            get_enabled = bkt.Callback(PictureFormat.is_pic_shape, shape=True),
                        ),
                        bkt.ribbon.Button(
                            id = 'pic_crop_paste',
                            label="Bild-Zuschnitt einfügen",
                            image_mso="PasteWithColumnWidths",
                            screentip="Größe und Position des Bildausschnitts einfügen",
                            supertip="Überträgt die kopierte Größe und Position des Bilde-Ausschnitts auf das ausgewählte Bild.",
                            on_action=bkt.Callback(PictureFormat.paste_dimensions, shape=True),
                            get_enabled = bkt.Callback(PictureFormat.is_paste_enabled, shape=True),
                        ),
                        bkt.ribbon.MenuSeparator(),
                        bkt.ribbon.Button(
                            id = 'table_dimensions_copy',
                            label="Tabellengrößen kopieren",
                            image_mso="TableColumnsDistribute",
                            screentip="Breite/Höhe der Tabellenspalten/-zeilen kopieren",
                            supertip="Kopiert Höhe und Breite der Tabellenzeilen bzw. Tabellenspalten, um diese mit einer anderen Tabelle anzugleichen.",
                            on_action=bkt.Callback(TableFormat.copy_dimensions, shape=True),
                            get_enabled = bkt.Callback(TableFormat.is_table_shape, shape=True),
                        ),
                        bkt.ribbon.Button(
                            id = 'table_dimensions_paste',
                            label="Tabellengrößen einfügen",
                            image_mso="PasteWithColumnWidths",
                            screentip="Breite/Höhe der Tabellenspalten/-zeilen einfügen",
                            supertip="Überträgt die kopierten Tabellen-Dimensionen auf die ausgewählte Tabelle.",
                            on_action=bkt.Callback(TableFormat.paste_dimensions, shape=True),
                            get_enabled = bkt.Callback(TableFormat.is_paste_enabled, shape=True),
                        ),
                    ]
                ),
                bkt.ribbon.SplitButton(
                    id = 'edge_autofixer_splitbutton',
                    children=[
                        bkt.ribbon.Button(
                            id = 'edge_autofixer',
                            label="Kanten-Autofixer",
                            # image_mso='GridSettings',
                            get_image=bkt.Callback(EdgeAutoFixer.get_image, context=True),
                            supertip="Gleicht minimale Verschiebungen der Kanten der gewählten Shapes aus.",
                            on_action=bkt.Callback(EdgeAutoFixer.autofix_edges, shapes=True),
                            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                        ),
                        bkt.ribbon.Menu(
                            label="Kanten-Autofixer Menü",
                            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                            children=[
                                bkt.ribbon.Button(
                                    id = 'edge_autofixer-dd',
                                    label="Kanten-Autofixer diagonal von links-oben",
                                    image='autofixer_dd',
                                    supertip="Gleicht minimale Verschiebungen der Kanten der gewählten Shapes aus durch Vergrößerung auf Shapes links-oberhalb der anzupassenden Shapes.",
                                    on_action=bkt.Callback(EdgeAutoFixer.autofix_edges_diagonal_down, shapes=True),
                                    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                                ),
                                bkt.ribbon.Button(
                                    id = 'edge_autofixer-td',
                                    label="Kanten-Autofixer von oben nach unten",
                                    image='autofixer_td',
                                    supertip="Gleicht minimale Verschiebungen der Kanten der gewählten Shapes aus durch Vergrößerung auf Shapes links der anzupassenden Shapes.",
                                    on_action=bkt.Callback(EdgeAutoFixer.autofix_edges_top_down, shapes=True),
                                    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                                ),
                                bkt.ribbon.Button(
                                    id = 'edge_autofixer-lr',
                                    label="Kanten-Autofixer von links nach rechts",
                                    image='autofixer_lr',
                                    supertip="Gleicht minimale Verschiebungen der Kanten der gewählten Shapes aus durch Vergrößerung auf Shapes oberhalb der anzupassenden Shapes.",
                                    on_action=bkt.Callback(EdgeAutoFixer.autofix_edges_left_right, shapes=True),
                                    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                                ),
                                bkt.ribbon.MenuSeparator(),
                                bkt.ribbon.ToggleButton(
                                    label="Gruppen-Elemente einzeln anpassen",
                                    supertip="Gibt an ob Elemente einer Gruppe einzeln betrachtet werden, oder die gesamte Gruppe als Ganzes.",
                                    get_pressed=bkt.Callback(lambda: EdgeAutoFixer.groupitems is True),
                                    on_toggle_action=bkt.Callback(lambda pressed: EdgeAutoFixer.settings_setter("groupitems", pressed)),
                                ),
                                bkt.ribbon.Menu(
                                    label="Toleranz ändern",
                                    screentip="Kanten-Autofixer Toleranz",
                                    supertip="Schwellwert für Kanten-Autofixer anpassen.",
                                    children=[
                                        bkt.ribbon.ToggleButton(
                                            label="Klein 0,1cm",
                                            screentip="Toleranz klein 0,1cm",
                                            get_pressed=bkt.Callback(lambda: EdgeAutoFixer.threshold == 0.1),
                                            on_toggle_action=bkt.Callback(lambda pressed: EdgeAutoFixer.settings_setter("threshold", 0.1)),
                                        ),
                                        bkt.ribbon.ToggleButton(
                                            label="Mittel 0,3cm",
                                            screentip="Toleranz mittel 0,3cm",
                                            get_pressed=bkt.Callback(lambda: EdgeAutoFixer.threshold == 0.3),
                                            on_toggle_action=bkt.Callback(lambda pressed: EdgeAutoFixer.settings_setter("threshold", 0.3)),
                                        ),
                                        bkt.ribbon.ToggleButton(
                                            label="Groß 1 cm",
                                            screentip="Toleranz groß 1 cm",
                                            get_pressed=bkt.Callback(lambda: EdgeAutoFixer.threshold == 1.0),
                                            on_toggle_action=bkt.Callback(lambda pressed: EdgeAutoFixer.settings_setter("threshold", 1.0)),
                                        ),
                                    ]
                                ),
                            ]
                        ),
                    ]
                ),
                bkt.ribbon.MenuSeparator(),
                bkt.ribbon.Button(
                    id = 'add_into_group',
                    label="In Gruppe einfügen",
                    image_mso="ObjectsRegroup",
                    screentip="Shapes in Gruppe einfügen",
                    supertip="Sofern das zuerst oder zuletzt markierte Shape eine Gruppe ist, werden alle anderen Shapes in diese Gruppe eingefügt. Anderenfalls werden alle Shapes gruppiert.",
                    on_action=bkt.Callback(GroupsMore.add_into_group, shapes=True),
                    get_enabled = bkt.apps.ppt_shapes_min2_selected,
                ),
                bkt.ribbon.Button(
                    id = 'recursive_ungroup',
                    label="Rekursives Gruppe aufheben",
                    image_mso="ObjectsUngroup",
                    screentip="Gruppe aufheben rekursiv ausführen",
                    supertip="Wendet Gruppe aufheben solange an, bis alle verschachtelten Gruppen aufgelöst sind.",
                    on_action=bkt.Callback(GroupsMore.recursive_ungroup, shapes=True),
                    get_enabled = bkt.Callback(GroupsMore.contains_group, shapes=True),
                ),
                bkt.ribbon.MenuSeparator(title="Verknüpfte Shapes"),
                bkt.ribbon.Button(
                    id = 'shape_copy_to_all',
                    label="Shape auf Folgefolien kopieren und verknüpfen…",
                    image_mso="ShapesDuplicate",
                    screentip="Shape auf Folgefolien duplizieren",
                    supertip="Dupliziert das aktuelle Shapes auf alle Folien hinter der aktuellen Folie und verknüpft diese für zukünftige Operationen.",
                    on_action=bkt.Callback(LinkedShapes.copy_to_all),
                    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                ),
                bkt.ribbon.Button(
                    id = 'shape_find_similar_and_link',
                    label="Shape auf Folgefolien suchen und verknüpfen…",
                    image_mso="FindTag",
                    screentip="Gleiches Shape auf Folgefolien suchen und verknüpfen",
                    supertip="Sucht das aktuelle Shape auf allen Folien hinter der aktuellen Folie anhand Position und Größe und verknüpft diese miteinander.",
                    on_action=bkt.Callback(LinkedShapes.find_similar_and_link),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.MenuSeparator(),
                bkt.ribbon.Button(
                    id = 'link_shapes',
                    label="Ausgewählte Shapes miteinander verknüpfen",
                    image_mso="HyperlinkCreate",
                    screentip="Alle ausgewählte Shapes miteinander verknüpfen",
                    supertip="Die ausgewählten Shapes für zukünftige Operationen verknüpfen. Die Verknüpfung bleibt beim Kopieren der Shapes erhalten.",
                    on_action=bkt.Callback(LinkedShapes.link_shapes, shapes=True),
                    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                ),
                bkt.ribbon.Button(
                    id = 'extend_link_shapes',
                    label="Bestehende Shape-Verknüpfung erweitern",
                    # image_mso="HyperlinkCreate",
                    screentip="Bestehende Shape-Verknüpfung erweitern",
                    supertip="Um die bestehende Shape-Verknüpfung zu erweitern, wird die interne ID zwischengespeichert. Über 'Ausgewählte Shapes zur Verknüpfung hinzufügen' können dann weitere Shapes zur Verknüpfung hinzugefügt werden.",
                    on_action=bkt.Callback(LinkedShapes.extend_link_shapes, shape=True),
                    get_enabled = bkt.Callback(LinkedShapes.is_linked_shape, shape=True),
                ),
                bkt.ribbon.Button(
                    id = 'add_to_link_shapes',
                    label="Ausgewählte Shapes zur Verknüpfung hinzufügen",
                    # image_mso="HyperlinkCreate",
                    screentip="Ausgewählte Shapes zur Verknüpfung hinzufügen",
                    supertip="Ausgewählte Shapes zur zwischengespeicherten ID hinzufügen. Vorher muss eine neue Verknüpfung angelegt oder eine bestehende erweitert werden.",
                    on_action=bkt.Callback(LinkedShapes.add_to_link_shapes, shapes=True),
                    get_enabled = bkt.Callback(LinkedShapes.enabled_add_linked_shapes),
                ),
                bkt.ribbon.MenuSeparator(),
                bkt.ribbon.Button(
                    id = 'unlink_shape',
                    label="Einzelne Shape-Verknüpfung entfernen",
                    image_mso="HyperlinkRemove",
                    screentip="Verknüpfung des ausgewählten Shapes entfernen",
                    supertip="Entfernt die ID zur Verknüpfung vom aktuellen Shape. Alle anderen Shapes mit der gleichen ID bleiben verknüpft.",
                    on_action=bkt.Callback(LinkedShapes.unlink_shape, shape=True),
                    get_enabled = bkt.Callback(LinkedShapes.is_linked_shape),
                ),
                bkt.ribbon.Button(
                    id = 'unlink_all_shapes',
                    label="Gesamte Shape-Verknüpfung auflösen",
                    # image_mso="HyperlinkRemove",
                    screentip="Alle Shape-Verknüpfungen entfernen",
                    supertip="Entfernt die ID zur Verknüpfung vom aktuellen Shape sowie allen verknüpften Shapes mit der gleichen ID.",
                    on_action=bkt.Callback(LinkedShapes.unlink_all_shapes, shape=True, context=True),
                    get_enabled = bkt.Callback(LinkedShapes.is_linked_shape),
                ),

                # bkt.ribbon.MenuSeparator(),
                # bkt.ribbon.Menu(
                #     label='Verknüpfte Shapes',
                #     image_mso='ControlAlignToGrid',
                #     screentip="Operationen auf verknüpfte Shapes",
                #     supertip="Alle verknüpften Shapes löschen oder ausrichten. Optionen stehen auch im Kontextmenü von verknüpften Shapes zur Verfügung.",
                #     get_enabled = bkt.Callback(LinkedShapes.is_linked_shape, shape=True),
                #     children=[
                #         bkt.ribbon.Button(
                #             id = 'linked_shapes_count',
                #             label="Anzahl verknüpfter Shapes",
                #             image_mso="FindDialog",
                #             screentip="Verknüpfte Shapes zählen",
                #             supertip="Zählt die Anzahl der verknüpften Shapes auf allen Folien.",
                #             on_action=bkt.Callback(LinkedShapes.count_link_shapes, shape=True, context=True),
                #             # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                #         ),
                #         bkt.ribbon.Button(
                #             id = 'linked_shapes_next',
                #             label="Nächstes verknüpfte Shape finden",
                #             image_mso="FindNext",
                #             screentip="Zum nächsten verknüpften Shape gehen",
                #             supertip="Sucht nach dem nächste verknüpften Shape. Sollte auf den Folgefolien kein Shape mehr kommen, wird das erste verknüpfte Shape der Präsentation gesucht.",
                #             on_action=bkt.Callback(LinkedShapes.goto_linked_shape, shape=True, context=True),
                #             # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                #         ),
                #         bkt.ribbon.MenuSeparator(),
                #         bkt.ribbon.Button(
                #             id = 'linked_shapes_delete',
                #             label="Andere löschen",
                #             image_mso="HyperlinkRemove",
                #             screentip="Verknüpfte Shapes löschen",
                #             supertip="Alle verknüpften Shapes auf allen Folien löschen.",
                #             on_action=bkt.Callback(LinkedShapes.delete_linked_shapes, shape=True, context=True),
                #             # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                #         ),
                #         bkt.ribbon.Button(
                #             id = 'linked_shapes_replace',
                #             label="Andere mit diesem ersetzen",
                #             image_mso="HyperlinkCreate",
                #             screentip="Verknüpfte Shapes ersetzen",
                #             supertip="Alle verknüpften Shapes auf allen Folien mit ausgewähltem Shape ersetzen.",
                #             on_action=bkt.Callback(LinkedShapes.replace_with_this, shape=True, context=True),
                #             # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                #         ),
                #         bkt.ribbon.MenuSeparator(),
                #         bkt.ribbon.Button(
                #             id = 'linked_shapes_align',
                #             label="Position angleichen",
                #             image_mso="ControlAlignToGrid",
                #             screentip="Position verknüpfter Shapes angleichen",
                #             supertip="Position und Rotation aller verknüpfter Shapes auf Position wie ausgewähltes Shape setzen.",
                #             on_action=bkt.Callback(LinkedShapes.align_linked_shapes, shape=True, context=True),
                #             # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                #         ),
                #         bkt.ribbon.Button(
                #             id = 'linked_shapes_size',
                #             label="Größe angleichen",
                #             image_mso="SizeToControlHeightAndWidth",
                #             screentip="Größe verknüpfter Shapes angleichen",
                #             supertip="Größe aller verknüpfter Shapes auf Größe wie ausgewähltes Shape setzen.",
                #             on_action=bkt.Callback(LinkedShapes.size_linked_shapes, shape=True, context=True),
                #             # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                #         ),
                #         bkt.ribbon.Button(
                #             id = 'linked_shapes_format',
                #             label="Formatierung angleichen",
                #             image_mso="FormatPainter",
                #             screentip="Formatierung verknüpfter Shapes angleichen",
                #             supertip="Formatierung aller verknüpfter Shapes auf Größe wie ausgewähltes Shape setzen.",
                #             on_action=bkt.Callback(LinkedShapes.format_linked_shapes, shape=True, context=True),
                #             # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                #         ),
                #         bkt.ribbon.Button(
                #             id = 'linked_shapes_text',
                #             label="Text angleichen",
                #             image_mso="TextBoxInsert",
                #             screentip="Text verknüpfter Shapes angleichen",
                #             supertip="Text aller verknüpfter Shapes auf Größe wie ausgewähltes Shape setzen.",
                #             on_action=bkt.Callback(LinkedShapes.text_linked_shapes, shape=True, context=True),
                #             # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                #         ),
                #         bkt.ribbon.MenuSeparator(),
                #         bkt.ribbon.Button(
                #             id = 'linked_shapes_all',
                #             label="Alles angleichen",
                #             image_mso="GroupUpdate",
                #             screentip="Alle Eigenschaften verknüpfter Shapes angleichen",
                #             supertip="Alle Eigenschaften aller verknüpfter Shapes wie ausgewähltes Shape setzen.",
                #             on_action=bkt.Callback(LinkedShapes.equalize_linked_shapes, shape=True, context=True),
                #             # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                #         ),
                #     ]
                # )
            ]
        ),
        
        bkt.mso.control.ObjectsGroup,
        bkt.mso.control.ObjectsUngroup,
        bkt.mso.control.ObjectsRegroup,
        # bkt.mso.control.ObjectBringToFrontMenu,
        bkt.ribbon.SplitButton(
            id="bkt_bringtofront_menu",
            show_label=False,
            children=[
                bkt.mso.button.ObjectBringToFront,
                bkt.ribbon.Menu(label="Vordergrund-Menü", children=[
                    bkt.mso.control.ObjectBringToFront,
                    bkt.mso.control.ObjectBringForward,
                    bkt.ribbon.Button(
                        label="Hintere nach vorne",
                        supertip="Bringt alle hinteren Shapes genau vor das vorderste Shape",
                        image="zorder_back_to_front",
                        get_enabled=bkt.apps.ppt_shapes_min2_selected,
                        on_action=bkt.Callback(PositionSize.back_to_front, shapes=True),
                    ),
                ])
            ]
        ),
        # bkt.mso.control.ObjectSendToBackMenu,
        bkt.ribbon.SplitButton(
            id="bkt_sendtoback_menu",
            show_label=False,
            children=[
                bkt.mso.button.ObjectSendToBack,
                bkt.ribbon.Menu(label="Hintergrund-Menü", children=[
                    bkt.mso.control.ObjectSendToBack,
                    bkt.mso.control.ObjectSendBackward,
                    bkt.ribbon.Button(
                        label="Vordere nach hinten",
                        supertip="Bringt alle vordere Shapes genau hinter das hinterste Shape",
                        image="zorder_front_to_back",
                        get_enabled=bkt.apps.ppt_shapes_min2_selected,
                        on_action=bkt.Callback(PositionSize.front_to_back, shapes=True),
                    ),
                ])
            ]
        ),

        tablearrange_button

    ]
)




