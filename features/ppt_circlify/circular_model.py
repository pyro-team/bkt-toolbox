# -*- coding: utf-8 -*-
'''

@author: rdebeerst
'''



import math
# import logging

# import bkt
import bkt.library.algorithms as algorithms

from bkt.library.powerpoint import cm_to_pt



# Initialisierung
# nehme alle Mittelpunkte
# nehme alle Ellipsen bestimmt durch vier Mittelpukte
# (wenn nur drei Objekte gegeben, dann kreis durch drei Punkte)
# nehme Ellipse mit kleinster Fläche
# schiebe Ellipse so, dass Top-most-Shape oben in der Mitte liegt
# nehme das als Initial-Ellipse

# gewichteten Mittelpunkt
# Radius aus Abstand zum ersten Punkt
# h-Squeeze mitteln

class CircularArrangement(object):
    DEBUG = False

    midpoint = [0,0]
    
    rotated = False
    fixed_radius = False
    centerpoint = False

    width = cm_to_pt(10.0)
    height = cm_to_pt(8.5)
    segment_start = 0


    @classmethod
    def _draw_debug_point(cls, shape, point, text=None):
        if not cls.DEBUG:
            return
        dot = shape.parent.shapes.addshape(
                9, #msoShapeOval
                point[0]-5,point[1]-5,
                10,10
            )
        dot.TextFrame2.TextRange.Font.Size=8
        if text:
            dot.TextFrame2.TextRange.Text = text

    @classmethod
    def _draw_debug_circle(cls, shape, point, width, height):
        if not cls.DEBUG:
            return
        dot = shape.parent.shapes.addshape(
                9, #msoShapeOval
                point[0]-width/2,point[1]-height/2,
                width,height
            )
        dot.fill.visible=0
        dot.line.ForeColor.RGB=255


    @classmethod
    def arrange_circular(cls, shapes):
        midpoint, width, height, segment_start = cls.get_ellipse_params(shapes)
        cls.arrange_circular_wargs(shapes, midpoint, width, height, segment_start)

    @classmethod
    def get_ellipse_params(cls, shapes):
        # compute weightend midpoint
        # shape_midpoints = [ [s.left+s.width/2.0, s.top+s.height/2.0] for s in shapes]
        # cls.midpoint = algorithms.mid_point(shape_midpoints)
        cls.midpoint = algorithms.mid_point_shapes(shapes)
        
        if cls.fixed_radius:
            cls.height = cls.width

        return (cls.midpoint, cls.width, cls.height, cls.segment_start)

    @classmethod
    def determine_ellipse_params(cls, shapes):
        # compute weightend midpoint
        shape_midpoints = [ [s.left+s.width/2, s.top+s.height/2] for s in shapes]
        cls.midpoint = algorithms.mid_point(shape_midpoints)

        # compute if centerpoint exists
        if algorithms.is_close(shape_midpoints[0][0], cls.midpoint[0], 0.1) and algorithms.is_close(shape_midpoints[0][1], cls.midpoint[1], 0.1):
            cls.centerpoint = True
            #exclude centerpoint from further calculations
            del shape_midpoints[0]
            shapes = shapes[1:]
        else:
            cls.centerpoint = False

        # compute all vectors from midpoints to shapes
        vectors = [[sm[0]-cls.midpoint[0], sm[1]-cls.midpoint[1]] for sm in shape_midpoints]

        # interpolate segment start (angle to frist vector)
        cls.segment_start = math.degrees( math.atan2(vectors[0][1], vectors[0][0]) + math.pi/2) #add pi/2 as 90° is subtracted in determine_points
        if cls.segment_start <= 0:
            cls.segment_start = (360+cls.segment_start) % 360 #ensure positive value
        
        # interpolate radius as max for each vector part
        cls.width = max([v[0] for v in vectors]) *2
        cls.height = max([v[1] for v in vectors]) *2

        # # determine points
        # points = cls.determine_points(shapes, cls.midpoint, 2, 2, 90)

        # # compute radius
        # # x-stretch faktor for every point
        # factors = [ 1.0* (shape_midpoints[i][0]-cls.midpoint[0])/(points[i][0]-cls.midpoint[0])  for i in range(0, len(shapes)) if (points[i][0]-cls.midpoint[0]) != 0]
        # # middle strech factor
        # radius_x = sum(factors)/len(factors)
        # cls.width = 2*radius_x
        
        # # y-stretch faktor for every point
        # factors = [ 1.0* (shape_midpoints[i][1]-cls.midpoint[1])/(points[i][1]-cls.midpoint[1])  for i in range(0, len(shapes)) if (points[i][1]-cls.midpoint[1]) != 0]
        # # middle strech factor
        # radius_y = sum(factors)/len(factors)
        # cls.height = 2*radius_y

        # compute options
        cls.fixed_radius = algorithms.is_close(cls.height, cls.width, 0.1)
        cls.rotated = any(shapes[0].rotation != s.rotation for s in shapes)

        #debug drawings
        cls._draw_debug_circle(shapes[0], cls.midpoint, cls.width, cls.height)
        cls._draw_debug_point(shapes[0], cls.midpoint, "C")
        cls._draw_debug_point(shapes[0], shape_midpoints[0], "1")
        
        return (cls.midpoint, cls.width, cls.height, cls.segment_start)
    
    @classmethod
    def set_circ_width(cls, shapes, value):
        value = float(max(0,value))
        cls.width = value
        if cls.fixed_radius:
            cls.height = value
        cls.arrange_circular(shapes)
    
    @classmethod
    def get_circ_width(cls, shapes):
        return cls.width
    
    @classmethod
    def set_circ_height(cls, shapes, value):
        value = float(max(0,value))
        cls.height = value
        if cls.fixed_radius:
            cls.width = value
        cls.arrange_circular(shapes)
    
    @classmethod
    def get_circ_height(cls, shapes):
        return cls.height

    @classmethod
    def get_segment_start(cls, shapes):
        return round(cls.segment_start,1)
    
    @classmethod
    def set_segment_start(cls, shapes, value):
        #ensure that value is positive and between 0 and 359
        cls.segment_start = value if 0 <= value < 360 else (360+value)%360
        cls.arrange_circular(shapes)
    
    @classmethod
    def determine_points(cls, shapes, midpoint, width, height, segment_start):
        # # Nullpunkt als Mittelpunkt verwenden, mit Radius 1
        # segments = bezier.kreisSegmente(len(shapes), 1, [0,0])
        # points = [ [s[0][0][0], s[0][0][1]] for s in segments]
        
        # # Segments starten rechts im Kreis
        # # Alle Punkte um 90 Grad nach links drehen, damit erstes Objekt oben steht
        # cls._draw_debug_point(shapes[0], midpoint, "C")
        # cls._draw_debug_point(shapes[0], points[0], "A")
        # points = [ algorithms.rotate_point(p[0], p[1], 0, 0, segment_start)  for p in points]
        # cls._draw_debug_point(shapes[0], points[0], "B")
        
        # # Punkte skalieren (Höhe/Breite)
        # points = [ [ width/2 * p[0], height/2 * p[1] ]   for p in points]
        
        # # Punkte verschieben anhand midpoint
        # points = [ [ p[0] + midpoint[0], p[1] + midpoint[1] ]   for p in points]

        # if centerpoint-shape is active, lead first shape out for remaining calculation
        if cls.centerpoint:
            shapes = shapes[1:]
        
        points = algorithms.get_ellipse_points(len(shapes), width/2.0, height/2.0, segment_start-90, midpoint)
        return points
    
    @classmethod
    def arrange_circular_wargs(cls, shapes, midpoint, width, height, segment_start):
        points = cls.determine_points(shapes, midpoint, width, height, segment_start)

        if cls.centerpoint:
            center_shape = shapes.pop(0)
            center_shape.left = midpoint[0] - center_shape.width /2
            center_shape.top  = midpoint[1] - center_shape.height /2
        
        for i in range(0, len(shapes)):
            shapes[i].left = points[i][0] - shapes[i].width /2
            shapes[i].top  = points[i][1] - shapes[i].height /2
            
            if cls.rotated:
                shapes[i].rotation = (360/len(shapes)*i +segment_start)%360
            else:
                shapes[i].rotation = shapes[0].rotation
        
    
    
    @classmethod
    def arrange_circular_rotated(cls, pressed):
        cls.rotated = pressed

    @classmethod
    def arrange_circular_rotated_pressed(cls):
        return cls.rotated

    @classmethod
    def arrange_circular_fixed(cls, pressed):
        cls.fixed_radius = pressed

    @classmethod
    def arrange_circular_fixed_pressed(cls):
        return cls.fixed_radius

    @classmethod
    def arrange_circular_centerpoint(cls, pressed):
        cls.centerpoint = pressed

    @classmethod
    def arrange_circular_centerpoint_pressed(cls):
        return cls.centerpoint