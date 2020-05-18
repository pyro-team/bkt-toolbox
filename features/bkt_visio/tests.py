# -*- coding: utf-8 -*-
'''
Created on 2016-04-27
@author: Florian Stallmann
'''

from __future__ import absolute_import

import math
# import logging

import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.IO')
clr.AddReference("Microsoft.Office.Interop.Visio")

import Microsoft.Office.Interop.Visio as Visio
import System.Windows.Forms as F
import System.IO as IO

import bkt
# from bkt.library import visio

class TestVisio(object):
    @staticmethod
    def bounding_box(page, shape):
        rect = shape.bounding_box

        #bkt.message("Bounding Box: " + str(rect.x) + "/" + str(rect.y) + " & " + str(rect.width) + "/" + str(rect.height))
        #bkt.message("Real Values: " + str(shape._x) + "/" + str(shape._y) + " & " + str(shape.width) + "/" + str(shape.height))

        shp_bb = page.drawRect(rect.x, rect.y, rect.width, rect.height)
        shp_bb.fillpattern = 0

        #bkt.message("farbe: " + str(shp_bb.fillpattern))

        #shp_bb.CellsSRC(visSectionObject, visRowFill, visFillPattern).FormulaU = "0"

        # out = [clr.Reference[float]() for _ in range(4)]
        # shape.shape.BoundingBox(8192+4, *out) #visBBoxDrawingCoords + visBBoxExtents
        # dblLeft = visio.inch2mm(out[0].Value)
        # dblBottom = visio.inch2mm(out[1].Value)
        # dblRight = visio.inch2mm(out[2].Value)
        # dblTop = visio.inch2mm(out[3].Value)
        # bkt.message("Bounding Box: " + str(dblLeft) + "/" + str(dblBottom) + "/" + str(dblRight) + "/" + str(dblTop))


    @staticmethod
    def locpin_dis(shape):
        cos_angle = math.cos(shape.angle)
        sin_angle = math.sin(shape.angle)

        # cos_x = cos_angle * shape.localpinx
        # cos_y = cos_angle * shape.localpiny
        # sin_x = sin_angle * shape.localpinx
        # sin_y = sin_angle * shape.localpiny

        # hypotenuse = shape.localpinx ** 2 + shape.localpiny ** 2
        # hypotenuse = sqrt(hypotenuse)
        hypotenuse = math.hypot(shape.localpinx, shape.localpiny)
        #hypotenuse = shape.localpinx
        #hypotenuse = hypot(shape.width, shape.height)
        #bkt.message("hyp: " + str(round(hypotenuse,2)))

        #bkt.message("cos: " + str(round(cos(shape.angle),2)))
        #bkt.message("sin: " + str(round(sin(shape.angle),2)))

        ankathete1 = hypotenuse * cos_angle
        ankathete2 = hypotenuse * sin_angle
        #ankathete4 = hypotenuse / cos_angle
        #ankathete5 = hypotenuse / sin_angle
        #ankathete7 = cos(radians(degrees(shape.angle))) / hypotenuse
        #ankathete8 = sin(radians(degrees(shape.angle))) / hypotenuse

        bkt.message("ank1: " + str(round(shape.x - ankathete1,2)))
        bkt.message("ank2: " + str(round(shape.x - ankathete2,2)))
        #bkt.message("ank4: " + str(round(shape.x - ankathete4,2)))
        #bkt.message("ank5: " + str(round(shape.x - ankathete5,2)))
        #bkt.message("ank7: " + str(round(shape.x - ankathete7,2)))
        #bkt.message("ank8: " + str(round(shape.x - ankathete8,2)))

        # bkt.message("bb_x: " + str(round(shape.x-ankathete,2)))
        #bb_x = shape.x - dis
        #bkt.message("dis: " + str(round(dis,2)))
        #bkt.message("bb_x: " + str(round(bb_x,2)))

    @staticmethod
    def clipboard_data(application):
        formats = F.Clipboard.GetDataObject().GetFormats(True)
        bkt.message("Formats: " + ", ".join(formats))

        #TestVisio._show_clipboard_data("Visio 15.0 Shapes")
        #TestVisio._show_clipboard_data("Preferred DropEffect")
        #TestVisio._show_clipboard_data("Embed Source")
        #TestVisio._show_clipboard_data("Link Source")
        #TestVisio._show_clipboard_data("Link Source Descriptor")
        #TestVisio._show_clipboard_data("Object Descriptor")

    @staticmethod
    def _show_clipboard_data(name):
        try:
            memstream = F.Clipboard.GetData(name)
            data = IO.StreamReader(memstream)
            data_string = data.ReadToEnd()

            print "Clipboard data for " + name + ": " + data_string
            bkt.message("Clipboard data for " + name + ": " + data_string)
        except:
            print "could not read " + name
    
    @staticmethod
    def shape_data(shape):
        import bkt.console
        msg = '''--- ALL SHAPE DATA ---

    X:    {}
    _X:   {}
    PinX: {}
    Y:    {}
    _Y:   {}
    PinY: {}
    BeginX: {}
    BeginY: {}
    EndX: {}
    EndY: {}
    1D: {}
    '''.format(shape.x, shape._x, shape.pinx, shape.y, shape._y, shape.piny, shape.beginx, shape.beginy, shape.endx, shape.endy, shape.shape.OneD)
        bkt.console.show_message(bkt.ui.endings_to_windows(msg))


class Adjustments(object):
    properties = [
        ("X", "x"),
        ("Begin X", "beginx"),
        ("End X", "endx"),

        ("Y", "y"),
        ("Begin Y", "beginy"),
        ("End Y", "endy"),

        ("Locpin X", "localpinx"),
        ("Left", "_left"),
        ("_X", "_x"),
        
        ("Locpin Y", "localpiny"),
        ("Bottom", "_bottom"),
        ("_Y", "_y"),

        ("Width", "width"),
        ("Height", "height"),
        ("Angle", "angle"),
    ]

    @classmethod
    def adjustment_edit_box(cls, prop):
        editbox= bkt.ribbon.EditBox(
            label = ' ' + prop[0],
            sizeString = '######',
            
            on_change   = bkt.Callback(
                lambda shapes, value: map( lambda shape: cls.set_adjustment(shape, prop[1], value), shapes),
                shapes=True),
            
            get_text    = bkt.Callback(
                lambda shape       : cls.get_adjustment(shape, prop[1]),
                shape=True),
            
            get_enabled = bkt.Callback(
                lambda shapes : len(shapes) > 0,
                shapes=True)

        )
        return editbox


    @classmethod
    def set_adjustment(cls, shape, attr, value):
        try:
            setattr(shape, attr, value)
        except:
            bkt.helpers.exception_as_message()

    @classmethod
    def get_adjustment(cls, shape, attr):
        try:
            return round(getattr(shape, attr), 2)
        except:
            # bkt.helpers.exception_as_message()
            return None



group_adjustments = bkt.ribbon.Group(
    label = "Properties",
    
    children=[
        Adjustments.adjustment_edit_box(prop)
        
        for prop in Adjustments.properties
    ]
)


test_gruppe = bkt.ribbon.Group(
    label="Experimental",
    children=[
        bkt.ribbon.Button(
            id = 'bounding_box',
            label="BoundingBox",
            show_label=True,
            image_mso='HappyFace',
            screentip="Draw the bounding box",
            on_action=bkt.Callback(TestVisio.bounding_box, page=True, shape=True),
            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name
        ),
        bkt.ribbon.Button(
            id = 'locpin_dis',
            label="Locpin Distance",
            show_label=True,
            image_mso='HappyFace',
            screentip="Calculate LocPin distance for rotated shapes",
            on_action=bkt.Callback(TestVisio.locpin_dis, shape=True),
            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name
        ),
        bkt.ribbon.Button(
            id = 'clipboard_data',
            label="Clipboard Analysis",
            show_label=True,
            image_mso='HappyFace',
            screentip="Get info about clipboard contents",
            on_action=bkt.Callback(TestVisio.clipboard_data, application=True),
            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name
        ),
        bkt.ribbon.Button(
            id = 'shape_data',
            label="Shape-Daten",
            show_label=True,
            image_mso='HappyFace',
            screentip="Get all shape data",
            on_action=bkt.Callback(TestVisio.shape_data, shape=True),
            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name
        )
    ]
)

bkt.visio.add_tab(
    bkt.ribbon.Tab(
        id="bkt_visio_toolbox_tests",
        #id_q="nsBKT:visio_toolbox_advanced",
        label=u"Toolbox TEST",
        insert_before_mso="TabHome",
        get_visible=bkt.Callback(lambda: True),
        children = [
            test_gruppe,
            group_adjustments
        ]
    )
)