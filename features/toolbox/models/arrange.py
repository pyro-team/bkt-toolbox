# -*- coding: utf-8 -*-
'''
Created on 19.01.2023

'''




import math

import bkt

import bkt.library.algorithms as algos
import bkt.library.powerpoint as pplib
pt_to_cm = pplib.pt_to_cm
cm_to_pt = pplib.cm_to_pt

from ..arrange import LocPinCollection, GlobalMasterShape


class Swap(object):
    locpin = LocPinCollection.swap

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
        if bkt.get_key_state(bkt.KeyCodes.SHIFT):
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


    @classmethod
    def replace_keep_size(cls, shapes):
        shapes = pplib.wrap_shapes(shapes, cls.locpin)
        master = shapes.pop(0) #first selected is master shape
        first = True
        for ref in shapes:
            if first:
                new = master
                first = False
            else:
                new = master.Duplicate()
            new.rotation = ref.rotation
            new.width    = ref.width
            if new.LockAspectRatio == 0 or new.height > ref.height:
                new.height   = ref.height
            new.top      = ref.top
            new.left     = ref.left
            pplib.set_shape_zorder(new, value=ref.ZOrderPosition)
            ref.Delete()
            new.Select(False)


class EqualSize(object):
    funcs = {
        "min": min,
        "max": max,
        "mean": bkt.library.algorithms.mean,
        "median": bkt.library.algorithms.median
    }

    @classmethod
    def _get_func(cls):
        if bkt.get_key_state(bkt.KeyCodes.SHIFT):
            return min
        elif bkt.get_key_state(bkt.KeyCodes.CTRL):
            if GlobalMasterShape.master == "FIRST":
                return lambda l: l.pop(0)
            else:
                return lambda l: l.pop()
        else:
            return max


    @classmethod
    def equal_height_master(cls, shapes):
        if GlobalMasterShape.master == "FIRST":
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
    def equal_height_control(cls, shapes, current_control):
        func = cls.funcs.get(current_control["tag"], cls._get_func())
        cls.equal_height(shapes, func)


    @classmethod
    def equal_width_master(cls, shapes):
        if GlobalMasterShape.master == "FIRST":
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

    @classmethod
    def equal_width_control(cls, shapes, current_control):
        func = cls.funcs.get(current_control["tag"], cls._get_func())
        cls.equal_width(shapes, func)


