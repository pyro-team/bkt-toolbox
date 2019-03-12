# -*- coding: utf-8 -*-

from helpers import GlobalLocPin
import bkt.library.algorithms as algos
import math

class ShapeWrapper(object):

    def __init__(self, shape, locpin=None):
        self.shape = shape
        self.locpin = locpin or GlobalLocPin
        self.bounding_nodes = None

    def __getattr__(self, name):
        # provides easy access to shape properties
        return getattr(self.shape, name)
    
    @property
    def left(self):
        ''' get left position considering locpin setting '''
        return self.shape.left + self.locpin.get_fractions()[1]*self.shape.width
    @left.setter
    def left(self, value):
        ''' set left position considering locpin setting '''
        self.shape.left = value - self.locpin.get_fractions()[1]*self.shape.width
    
    @property
    def top(self):
        ''' get top position considering locpin setting '''
        return self.shape.top + self.locpin.get_fractions()[0]*self.shape.height
    @top.setter
    def top(self, value):
        ''' set top position considering locpin setting '''
        self.shape.top = value - self.locpin.get_fractions()[0]*self.shape.height

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
        self.shape.left = value

    @property
    def y(self):
        ''' get top position '''
        return self.shape.top
    @y.setter
    def y(self, value):
        ''' set top position '''
        self.shape.top = value

    @property
    def x1(self):
        ''' get right position '''
        return self.shape.left+self.shape.width
    @x1.setter
    def x1(self, value):
        ''' set right position '''
        self.shape.left = value-self.shape.width

    @property
    def y1(self):
        ''' get bottom position '''
        return self.shape.top+self.shape.height
    @y1.setter
    def y1(self, value):
        ''' set bottom position '''
        self.shape.top = value-self.shape.height


    @property
    def center_x(self):
        ''' get center x position '''
        return self.shape.left + self.shape.width/2
    @center_x.setter
    def center_x(self, value):
        ''' set center x position '''
        self.shape.left = value - self.shape.width/2
    
    @property
    def center_y(self):
        ''' get center y position '''
        return self.shape.top + self.shape.height/2
    @center_y.setter
    def center_y(self, value):
        ''' set center y position '''
        self.shape.top = value - self.shape.height/2


    @property
    def visual_x(self):
        ''' get visual x (=left) position considering rotation '''
        return min( p[0] for p in self.get_bounding_nodes() )
    @visual_x.setter
    def visual_x(self, value):
        ''' set visual x (=left) position considering rotation '''
        delta = self.shape.left - self.visual_x
        self.shape.left = value + delta

    @property
    def visual_y(self):
        ''' get visual y (=top) position considering rotation '''
        return min( p[1] for p in self.get_bounding_nodes() )
    @visual_y.setter
    def visual_y(self, value):
        ''' set visual y (=top) position considering rotation '''
        delta = self.shape.top - self.visual_y
        self.shape.top = value + delta
    
    @property
    def visual_x1(self):
        ''' get visual x1 (=right) position considering rotation '''
        return max( p[0] for p in self.get_bounding_nodes() )
    @visual_x1.setter
    def visual_x1(self, value):
        ''' set visual x1 (=right) position considering rotation '''
        delta = self.shape.left - self.visual_x1
        self.shape.left = value + delta

    @property
    def visual_y1(self):
        ''' get visual y1 (=bottom) position considering rotation '''
        return max( p[1] for p in self.get_bounding_nodes() )
    @visual_y1.setter
    def visual_y1(self, value):
        ''' set visual y1 (=bottom) position considering rotation '''
        delta = self.shape.top - self.visual_y1
        self.shape.top = value + delta

    @property
    def visual_width(self):
        ''' get visual width considering rotation '''
        points = self.get_bounding_nodes()
        return max( p[0] for p in points ) - min( p[0] for p in points )
    @visual_width.setter
    def visual_width(self, value):
        ''' set visual width considering rotation '''
        if self.shape.rotation == 0 or self.shape.rotation == 180:
            self.shape.width = value
        elif self.shape.rotation == 90 or self.shape.rotation == 270:
            self.shape.height = value
        else:
            delta = value - self.visual_width
            # delta_vector (delta-width, 0) um shape-rotation drehen
            delta_vector = algos.rotate_point(delta, 0, 0, 0, self.shape.rotation)
            # vorzeichen beibehalten (entweder vergrößern oder verkleinern - nicht beides)
            vorzeichen = 1 if delta > 0 else -1
            delta_vector = [vorzeichen * abs(delta_vector[0]), vorzeichen * abs(delta_vector[1]) ]
            # shape anpassen
            self.shape.width += delta_vector[0]
            self.shape.height += delta_vector[1]

    @property
    def visual_height(self):
        ''' get visual height considering rotation '''
        points = self.get_bounding_nodes()
        return max( p[1] for p in points ) - min( p[1] for p in points )
    @visual_height.setter
    def visual_height(self, value):
        ''' set visual height considering rotation '''
        if self.shape.rotation == 0 or self.shape.rotation == 180:
            self.shape.height = value
        elif self.shape.rotation == 90 or self.shape.rotation == 270:
            self.shape.width = value
        else:
            delta = value - self.visual_height
            # delta_vector (delta-width, 0) um shape-rotation drehen
            delta_vector = algos.rotate_point(0, delta, 0, 0, self.shape.rotation)
            # vorzeichen beibehalten (entweder vergrößern oder verkleinern - nicht beides)
            vorzeichen = 1 if delta > 0 else -1
            delta_vector = [vorzeichen * abs(delta_vector[0]), vorzeichen * abs(delta_vector[1]) ]
            # shape anpassen
            self.shape.width += delta_vector[0]
            self.shape.height += delta_vector[1]
    
    @property
    def text(self):
        if self.shape.HasTextFrame == 0 or self.shape.TextFrame.HasText == 0:
            return None
        return self.shape.TextFrame.TextRange.Text
    @text.setter
    def text(self, value):
        if self.shape.HasTextFrame == 0:
            raise AttributeError("Shape has no textframe")
        if value is None:
            self.shape.TextFrame.TextRange.Delete()
        else:
            self.shape.TextFrame.TextRange.Text = value

    
    def get_bounding_nodes(self, force_update=False):
        ''' get and cache bounding points '''
        if force_update or not self.bounding_nodes:
            self.bounding_nodes = algos.get_bounding_nodes(self.shape)
        return self.bounding_nodes

    # @property
    # def rotation(self):
    #     return self.shape.rotation
    # @rotation.setter
    # def rotation(self, value):

    #     points = self.get_bounding_nodes()
    #     pivotX, pivotY = points[0]
    #     midX, midY = algos.mid_point(points)

    #     print("midx %s, midy %s" % (midX, midY))

    #     delta = value-self.shape.rotation
    #     # theta = delta*2*math.pi/360

    #     # dx = self.center_x - pivotX #+ (self.shape.left-visual_x)
    #     # dy = self.center_y - pivotY
        
    #     # newx = dx*math.cos(theta) - dy*math.sin(theta)
    #     # newy = dx*math.sin(theta) + dy*math.cos(theta)

    #     newx, newy = algos.rotate_point(self.center_x, self.center_y, pivotX, pivotY, delta)

    #     self.center_x = newx
    #     self.center_y = newy

    #     self.shape.rotation = value



def wrap_shape(shape, locpin=None):
    if isinstance(shape, ShapeWrapper):
        return shape
    return ShapeWrapper(shape, locpin)

def wrap_shapes(shapes, locpin=None):
    return [wrap_shape(shape, locpin) for shape in shapes]