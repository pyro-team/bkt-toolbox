# -*- coding: utf-8 -*-
'''
Created on 02.11.2017

@author: fstallmann
'''



# import math

import bkt.library.algorithms as algorithms


class ShapeWrapper(object):

    def __init__(self, shape, locpin=None):
        self.shape = shape
        self.locpin = locpin or self._get_global_locpin()
        self.locpin_nodes = None
        self.bounding_nodes = None

    def __getattr__(self, name):
        # provides easy access to shape properties
        return getattr(self.shape, name)
    
    def _get_global_locpin(self):
        from .helpers import GlobalLocPin
        return GlobalLocPin
    
    @property
    def left(self):
        ''' get left position considering locpin setting '''
        return round(self.shape.left + self.locpin.get_fractions()[1]*self.shape.width, 4) #max precision for position in ppt is 3 decimal places
    @left.setter
    def left(self, value):
        ''' set left position considering locpin setting '''
        # self.shape.left = value - self.locpin.get_fractions()[1]*self.shape.width
        self.shape.incrementLeft(value-self.left)
    
    @property
    def top(self):
        ''' get top position considering locpin setting '''
        return round(self.shape.top + self.locpin.get_fractions()[0]*self.shape.height, 4) #max precision for position in ppt is 3 decimal places
    @top.setter
    def top(self, value):
        ''' set top position considering locpin setting '''
        # self.shape.top = value - self.locpin.get_fractions()[0]*self.shape.height
        self.shape.incrementTop(value-self.top)

    @property
    def width(self):
        ''' get width considering locpin setting '''
        return self.shape.width
    @width.setter
    def width(self, value):
        ''' set width considering locpin setting '''
        # msoScaleFromTopLeft =0
        # msoScaleFromMiddle =1
        # msoScaleFromBottomRight =2
        fix_height, fix_width = self.locpin.fixation
        orig_lar = self.shape.LockAspectRatio
        self.shape.LockAspectRatio = 0
        if self.shape.width > 0:
            factor = value/self.shape.width
            self.shape.scaleWidth(factor, 0, fix_width-1)
            self.shape.width = value #ensure that width is set to new width. this is important for tables as scaleWidth only moves the table but does not change the size!
            if orig_lar == -1:
                new_height = factor * self.shape.height
                self.shape.scaleHeight(factor, 0, fix_height-1)
                self.shape.height = new_height #ensure that height is set to new height. this is important for tables as scaleHeight only moves the table but does not change the size!
        else:
            #workaround for div by zero
            self.shape.width = 1
            self.shape.scaleWidth(value, 0, fix_width-1)
            self.shape.width = value
        self.shape.LockAspectRatio = orig_lar

    @property
    def height(self):
        ''' get height considering locpin setting '''
        return self.shape.height
    @height.setter
    def height(self, value):
        ''' set height considering locpin setting '''
        # msoScaleFromTopLeft =0
        # msoScaleFromMiddle =1
        # msoScaleFromBottomRight =2
        fix_height, fix_width = self.locpin.fixation
        orig_lar = self.shape.LockAspectRatio
        self.shape.LockAspectRatio = 0
        if self.shape.height > 0:
            factor = value/self.shape.height
            self.shape.scaleHeight(factor, 0, fix_height-1)
            self.shape.height = value #ensure that height is set to new height. this is important for tables as scaleHeight only moves the table but does not change the size!
            if orig_lar == -1:
                new_width = factor * self.shape.width
                self.shape.scaleWidth(factor, 0, fix_width-1)
                self.shape.width = new_width #ensure that width is set to new width. this is important for tables as scaleWidth only moves the table but does not change the size!
        else:
            #workaround for div by zero
            self.shape.height = 1
            self.shape.scaleHeight(value, 0, fix_height-1)
            self.shape.height = value
        self.shape.LockAspectRatio = orig_lar


    @property
    def x(self):
        ''' get left position '''
        return self.shape.left
    @x.setter
    def x(self, value):
        ''' set left position '''
        self.shape.incrementLeft(value-self.x) #using IncrementLeft() has advantage that connected connectors are not moved; setting left directly has strange effect on connectors

    @property
    def y(self):
        ''' get top position '''
        return self.shape.top
    @y.setter
    def y(self, value):
        ''' set top position '''
        self.shape.incrementTop(value-self.y) #using IncrementTop() has advantage that connected connectors are not moved; setting top directly has strange effect on connectors

    @property
    def x1(self):
        ''' get right position '''
        return self.shape.left+self.shape.width
    @x1.setter
    def x1(self, value):
        ''' set right position '''
        self.shape.incrementLeft(value-self.x1)

    @property
    def y1(self):
        ''' get bottom position '''
        return self.shape.top+self.shape.height
    @y1.setter
    def y1(self, value):
        ''' set bottom position '''
        self.shape.incrementTop(value-self.y1)


    def resize_to_x(self, value):
        ''' resize shape to given left edge (x-value) '''
        self.shape.width += self.x-value
        self.x = value

    def resize_to_y(self, value):
        ''' resize shape to given top edge (y-value) '''
        self.shape.height += self.y-value
        self.y = value

    def resize_to_x1(self, value):
        ''' resize shape to given right edge (x1-value) '''
        self.shape.width = value-self.x

    def resize_to_y1(self, value):
        ''' resize shape to given bottom edge (y1-value) '''
        self.shape.height = value-self.y


    @property
    def lock_aspect_ratio(self):
        return self.shape.LockAspectRatio == -1
    @lock_aspect_ratio.setter
    def lock_aspect_ratio(self, value):
        self.shape.LockAspectRatio = -1 if value else 0


    def transpose(self):
        ''' switch shape height and width '''
        orig_lar = self.shape.LockAspectRatio
        self.shape.LockAspectRatio = 0
        self.width,self.height = self.height,self.width
        self.shape.LockAspectRatio = orig_lar
    
    def force_aspect_ratio(self, ratio, landscape=True):
        ''' force specific aspect ratio by settings  '''
        orig_lar = self.shape.LockAspectRatio
        self.shape.LockAspectRatio = 0
        if landscape or ratio < 1:
            self.width = self.height * ratio
        else:
            self.height = self.width * ratio
        self.shape.LockAspectRatio = orig_lar
    
    def square(self, w2h=True):
        ''' square shape by setting width to height (if w2h=True) or height to width '''
        self.force_aspect_ratio(1, w2h)


    @property
    def center_x(self):
        ''' get center x position '''
        return self.shape.left + self.shape.width/2
    @center_x.setter
    def center_x(self, value):
        ''' set center x position '''
        # self.shape.left = value - self.shape.width/2
        self.shape.incrementLeft(value-self.center_x)
    
    @property
    def center_y(self):
        ''' get center y position '''
        return self.shape.top + self.shape.height/2
    @center_y.setter
    def center_y(self, value):
        ''' set center y position '''
        # self.shape.top = value - self.shape.height/2
        self.shape.incrementTop(value-self.center_y)


    @property
    def visual_left(self):
        ''' get visual left position considering locpin setting '''
        return round(self.visual_x + self.locpin.get_fractions()[1]*self.visual_width, 4)
    @visual_left.setter
    def visual_left(self, value):
        ''' set visual left position considering locpin setting '''
        self.shape.incrementLeft(value-self.visual_left)

    @property
    def visual_top(self):
        ''' get visual top position considering locpin setting '''
        return round(self.visual_y + self.locpin.get_fractions()[0]*self.visual_height, 4)
    @visual_top.setter
    def visual_top(self, value):
        ''' set visual top position considering locpin setting '''
        self.shape.incrementTop(value-self.visual_top)


    @property
    def visual_x(self):
        ''' get visual x (=left) position considering rotation '''
        if self.shape.rotation == 0 or self.shape.rotation == 180:
            return self.x
        elif self.shape.rotation == 90 or self.shape.rotation == 270:
            return self.center_x-self.shape.height/2
        return min( p[0] for p in self.get_bounding_nodes() )
    @visual_x.setter
    def visual_x(self, value):
        ''' set visual x (=left) position considering rotation '''
        # delta = self.shape.left - self.visual_x
        # self.shape.left = value + delta
        self.shape.incrementLeft(value-self.visual_x)
        # force recalculation of bounding notes
        self.reset_caches()

    @property
    def visual_y(self):
        ''' get visual y (=top) position considering rotation '''
        if self.shape.rotation == 0 or self.shape.rotation == 180:
            return self.y
        elif self.shape.rotation == 90 or self.shape.rotation == 270:
            return self.center_y-self.shape.width/2
        return min( p[1] for p in self.get_bounding_nodes() )
    @visual_y.setter
    def visual_y(self, value):
        ''' set visual y (=top) position considering rotation '''
        # delta = self.shape.top - self.visual_y
        # self.shape.top = value + delta
        self.shape.incrementTop(value-self.visual_y)
        # force recalculation of bounding notes
        self.reset_caches()
    
    @property
    def visual_x1(self):
        ''' get visual x1 (=right) position considering rotation '''
        if self.shape.rotation == 0 or self.shape.rotation == 180:
            return self.x1
        elif self.shape.rotation == 90 or self.shape.rotation == 270:
            return self.center_x+self.shape.height/2
        return max( p[0] for p in self.get_bounding_nodes() )
    @visual_x1.setter
    def visual_x1(self, value):
        ''' set visual x1 (=right) position considering rotation '''
        # delta = self.shape.left - self.visual_x1
        # self.shape.left = value + delta
        self.shape.incrementLeft(value-self.visual_x1)
        # force recalculation of bounding notes
        self.reset_caches()

    @property
    def visual_y1(self):
        ''' get visual y1 (=bottom) position considering rotation '''
        if self.shape.rotation == 0 or self.shape.rotation == 180:
            return self.y1
        elif self.shape.rotation == 90 or self.shape.rotation == 270:
            return self.center_y+self.shape.width/2
        return max( p[1] for p in self.get_bounding_nodes() )
    @visual_y1.setter
    def visual_y1(self, value):
        ''' set visual y1 (=bottom) position considering rotation '''
        # delta = self.shape.top - self.visual_y1
        # self.shape.top = value + delta
        self.shape.incrementTop(value-self.visual_y1)
        # force recalculation of bounding notes
        self.reset_caches()

    @property
    def visual_width(self):
        ''' get visual width considering rotation '''
        if self.shape.rotation == 0 or self.shape.rotation == 180:
            return self.shape.width
        elif self.shape.rotation == 90 or self.shape.rotation == 270:
            return self.shape.height
        #else:
        points = self.get_bounding_nodes()
        return max( p[0] for p in points ) - min( p[0] for p in points )
    @visual_width.setter
    def visual_width(self, value):
        ''' set visual width considering rotation '''
        if self.shape.rotation == 0 or self.shape.rotation == 180:
            self.width = value
        elif self.shape.rotation == 90 or self.shape.rotation == 270:
            cur_x = self.visual_left #save current x
            self.shape.height = value #might change left edge of shape
            self.visual_left = cur_x #move back to x
        else:
            delta = value - self.visual_width
            # delta_vector (delta-width, 0) um shape-rotation drehen
            delta_vector = algorithms.rotate_point_by_shape_rotation(delta, 0, self.shape)
            # vorzeichen beibehalten (entweder vergrößern oder verkleinern - nicht beides)
            vorzeichen = 1 if delta > 0 else -1
            delta_vector = [vorzeichen * abs(delta_vector[0]), vorzeichen * abs(delta_vector[1]) ]
            # aktuelle position speichern
            cur_x, cur_y = self.visual_left, self.visual_top
            # shape anpassen
            self.shape.width += delta_vector[0]
            self.shape.height += delta_vector[1]
            # force recalculation of bounding notes
            self.reset_caches()
            # vorherige position wiederherstellen
            self.visual_left, self.visual_top = cur_x, cur_y

    @property
    def visual_height(self):
        ''' get visual height considering rotation '''
        if self.shape.rotation == 0 or self.shape.rotation == 180:
            return self.shape.height
        elif self.shape.rotation == 90 or self.shape.rotation == 270:
            return self.shape.width
        #else:
        points = self.get_bounding_nodes()
        return max( p[1] for p in points ) - min( p[1] for p in points )
    @visual_height.setter
    def visual_height(self, value):
        ''' set visual height considering rotation '''
        if self.shape.rotation == 0 or self.shape.rotation == 180:
            self.height = value
        elif self.shape.rotation == 90 or self.shape.rotation == 270:
            cur_y = self.visual_top #save current y
            self.shape.width = value #might change top edge of shape
            self.visual_top = cur_y #move back to y
        else:
            delta = value - self.visual_height
            # delta_vector (delta-width, 0) um shape-rotation drehen
            delta_vector = algorithms.rotate_point_by_shape_rotation(0, delta, self.shape)
            # vorzeichen beibehalten (entweder vergrößern oder verkleinern - nicht beides)
            vorzeichen = 1 if delta > 0 else -1
            delta_vector = [vorzeichen * abs(delta_vector[0]), vorzeichen * abs(delta_vector[1]) ]
            # aktuelle position speichern
            cur_x, cur_y = self.visual_left, self.visual_top
            # shape anpassen
            self.shape.width += delta_vector[0]
            self.shape.height += delta_vector[1]
            # force recalculation of bounding notes
            self.reset_caches()
            # vorherige position wiederherstellen
            self.visual_left, self.visual_top = cur_x, cur_y
    
    @property
    def locpin_x(self):
        ''' get x-coordinates of locpin '''
        points = self.get_locpin_nodes()
        return points[self.locpin.index][0]
    @locpin_x.setter
    def locpin_x(self, value):
        ''' set x-coordinates of locpin '''
        self.shape.incrementLeft(value-self.locpin_x)
        # force recalculation of bounding notes
        self.reset_caches()
    
    @property
    def locpin_y(self):
        ''' get y-coordinates of locpin '''
        points = self.get_locpin_nodes()
        return points[self.locpin.index][1]
    @locpin_y.setter
    def locpin_y(self, value):
        ''' set y-coordinates of locpin '''
        self.shape.incrementTop(value-self.locpin_y)
        # force recalculation of bounding notes
        self.reset_caches()

    @property
    def text(self):
        ''' get text if shape has textframe, otherwise None '''
        if self.shape.HasTextFrame == 0 or self.shape.TextFrame.HasText == 0:
            return None
        return self.shape.TextFrame.TextRange.Text
    @text.setter
    def text(self, value):
        ''' set text if shape has textframe; delete text if None is provided '''
        if self.shape.HasTextFrame == 0:
            raise AttributeError("Shape has no textframe")
        if value is None:
            self.shape.TextFrame.TextRange.Delete()
        else:
            self.shape.TextFrame.TextRange.Text = value

    def reset_caches(self):
        self.bounding_nodes = None
        self.locpin_nodes = None

    def get_bounding_nodes(self, force_update=False):
        ''' get and cache bounding points '''
        if force_update or not self.bounding_nodes:
            self.bounding_nodes = algorithms.get_bounding_nodes(self.shape)
        return self.bounding_nodes
    
    def get_locpin_nodes(self, force_update=False):
        ''' get and cache loc pin points '''
        points = self.get_bounding_nodes(force_update) #left-top, left-bottom, right-bottom, right-top
        if force_update or not self.locpin_nodes:
            self.locpin_nodes = [
                points[0], algorithms.mid_point([points[0], points[3]]), points[3],
                algorithms.mid_point([points[0], points[1]]), algorithms.mid_point(points), algorithms.mid_point([points[3], points[2]]),
                points[1], algorithms.mid_point([points[1], points[2]]), points[2],
            ]
        return self.locpin_nodes

    @property
    def rotation(self):
        ''' get shape rotation '''
        return self.shape.rotation
    @rotation.setter
    def rotation(self, value):
        ''' set shape rotation '''
        #save current locpin position
        top, left = self.locpin_y, self.locpin_x
        #rotate
        self.shape.incrementRotation(value-self.rotation)
        #reset caches to force recalculation of locpins
        self.reset_caches()
        #restore locpin position
        self.locpin_y = top
        self.locpin_x = left



def wrap_shape(shape, locpin=None):
    if isinstance(shape, ShapeWrapper):
        return shape
    return ShapeWrapper(shape, locpin)

def wrap_shapes(shapes, locpin=None):
    return [wrap_shape(shape, locpin) for shape in shapes]