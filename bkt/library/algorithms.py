# -*- coding: utf-8 -*-
'''
Created on 11.09.2013

@author: cschmitt
'''

from __future__ import division #always force float-division, for int divison use //

import math

def median(values):
    v = sorted(values)
    if not v:
        raise ValueError
    n = len(v)
    if n % 2 == 0:
        return (v[n//2-1]+v[n//2])*0.5
    else:
        return v[(n-1)//2]

def mean(values):
    return sum(values)/len(values)

def mid_point(points):
    sum_x = 0
    sum_y = 0
    
    for p in points:
        sum_x +=p[0]
        sum_y +=p[1]
    
    return [sum_x/len(points), sum_y/len(points)]

def mid_point_shapes(shapes):
    sum_x = 0
    sum_y = 0
    
    for s in shapes:
        sum_x +=s.left+s.width/2
        sum_y +=s.top+s.height/2
    
    return [sum_x/len(shapes), sum_y/len(shapes)]

def is_close(a, b, tolerence=1e-9):
    # refer to https://github.com/PythonCHB/close_pep/blob/master/is_close.py
    if a == b:
        return True
    diff = abs(a-b)
    return (diff <= abs(tolerence * b)) or (diff <= abs(tolerence * a))

def get_bounds(points):
    x = [p[0] for p in points]
    y = [p[1] for p in points]
    
    left = min(x)
    top = min(y)
    width = max(x)-left
    height = max(y)-top
    
    return left,top,width,height

def get_bounds_shapes(shapes):
    def iter_points():
        for cell in shapes:
            x0 = cell.Left
            y0 = cell.Top
            yield (x0,y0)
            x1 = x0 + cell.Width
            y1 = y0 + cell.Height
            yield (x1,y1)
    
    points = list(iter_points())
    return get_bounds(points)

def rotate_point(x, y, x0, y0, deg):
    ''' rotate (x,y) arround (x0, y0) by degree '''
    # theta = deg*2*math.pi/360
    theta = math.radians(deg)
    return x0+(x-x0)*math.cos(theta)+(y-y0)*math.sin(theta), y0-(x-x0)*math.sin(theta)+(y-y0)*math.cos(theta)

def rotate_point_by_shape_rotation(x, y, shape):
    ''' rotate (x,y) arround (0,0) by shape rotation '''
    return rotate_point(x, y, 0, 0, 360-shape.rotation)

def get_bounding_nodes(shape):
    ''' compute bounding nodes (surrounding-square) for rotated shapes '''
    points = [ [ shape.left, shape.top ], [shape.left, shape.top+shape.height], [shape.left+shape.width, shape.top+shape.height], [shape.left+shape.width, shape.top] ]

    x0 = shape.left + shape.width/2
    y0 = shape.top + shape.height/2

    rotated_points = [
        list(rotate_point(p[0], p[1], x0, y0, 360-shape.rotation)) #rotation in ppt is inverted
        for p in points
    ]
    return rotated_points

def get_ellipse_points(n, r1, r2, start_deg=0, midpoint=(0,0)):
    ''' compute n points distributed among an ellipse with radius r1 and r2. first point starts at start_deg degree. optional midpoint can be provided. '''
    #Note: for an ellipse the points are actually not equally distributed, see https://stackoverflow.com/questions/6972331/how-can-i-generate-a-set-of-points-evenly-distributed-along-the-perimeter-of-an
    return [
            (r1 * math.cos(theta) + midpoint[0], r2 * math.sin(theta) + midpoint[1])
            for theta in (math.radians(start_deg) + math.pi*2 * i/n for i in range(n))
            ]

def get_rgb_from_ole(ole):
    ''' get rgb values from ole color value '''
    # return ole%256, ole//256%256, ole//256//256%256
    ole = int(ole)
    return (ole & 255 << 0) >> 0, (ole & 255 << 8) >> 8, (ole & 255 << 16) >> 16

def get_ole_from_rgb(r,g,b):
    ''' get rgb values from ole color value '''
    # return round(r + g*256 + b*256*256)
    return int(r) << 0 | int(g) << 8 | int(b) << 16

def get_brightness_from_rgb(r,g,b):
    ''' get brightness/lightness from rgb values '''
    #code from colorsys.rgb_to_hls
    maxc = max(r, g, b)
    minc = min(r, g, b)
    l = (minc+maxc)/2.0
    return l