# -*- coding: utf-8 -*-

from __future__ import absolute_import, division

import logging
import math

import bkt.ui
notify_property = bkt.ui.notify_property

import bkt.library.bezier as bezier
import bkt.library.powerpoint as pplib
pt_to_cm = pplib.pt_to_cm
cm_to_pt = pplib.cm_to_pt


# ===================================================================
# = functionality to create segmented circles or circular processes =
# ===================================================================

class SegmentedCircle(object):

    @staticmethod
    def is_segmented_circle(shape):
        def test_shape(s):
            return s.type == pplib.MsoShapeType["msoGraphic"] \
                and s.autoshapetype == pplib.MsoAutoShapeType["msoShapeNotPrimitive"] \
                and s.nodes.count >= 8
        try:
            return (shape.type == pplib.MsoShapeType["msoGroup"] and all(test_shape(s) for s in shape.GroupItems)) \
                or test_shape(shape)
        except:
            return False

    @staticmethod
    def determine_from_shape(shape):
        try:
            num_segments = shape.GroupItems.Count
            shp0 = shape.GroupItems[1]
        except:
            num_segments = 1
            shp0 = shape
        num_nodes = shp0.nodes.count

        #arrow vs segment
        use_arrow_shape = num_nodes in [28,10,16] #28 if 1 segment, 10 or 16 if >1 segment

        #width
        x1 = shp0.nodes[1].points[0,0]
        y1 = shp0.nodes[1].points[0,1]
        if use_arrow_shape:
            x2 = shp0.nodes[num_nodes-1].points[0,0]
            y2 = shp0.nodes[num_nodes-1].points[0,1]
        else:
            x2 = shp0.nodes[num_nodes].points[0,0]
            y2 = shp0.nodes[num_nodes].points[0,1]

        width = math.hypot(x1-x2,y1-y2)*2
        width_percentage = min(100,max(1,round(width/shape.width *100)))

        size_outer = shape.width/2

        return num_segments, width_percentage, size_outer, use_arrow_shape

#TODO: update_segmented_circle, consider flip_and_rotation_correction from processshapes

    @staticmethod
    def create_segmented_circle(slide, num_segments, width_percentage, size_outer=100, use_arrow_shape=False):
        # aussenKurve = bezier.bezierKreisNRM(2,100,[100,100])
        # innenKurve = bezier.bezierKreisNRM(2,75,[100,100])
        size_inner = size_outer * (1-width_percentage/100.)
        aussenKurve = bezier.kreisSegmente(num_segments, size_outer, [200,200])
        innenKurve  = bezier.kreisSegmente(num_segments, size_inner, [200,200])

        # shapeCount = slide.shapes.count

        def addNodesForCurves(ffb, kurven):
            for k in kurven:
                ffb.AddNodes(1, 1, k[1][0], k[1][1], k[2][0], k[2][1], k[3][0], k[3][1])

        for i in range(0,len(aussenKurve)):
            # Außen- und Innenkurve jeweils als Liste von Bezierkurven
            aK = aussenKurve[i]
            iK = innenKurve[i]; iK.reverse()
            for k in iK:
                k.reverse()
            # Hinweg auf Außenkurve
            ffb = slide.Shapes.BuildFreeform(1, aK[0][0][0], aK[0][0][1])
            addNodesForCurves(ffb, aK)
            # Uebergang zur Innenkurve
            # mit kleiner Spitze
            A=aK[-1][3]; B=iK[0][0]; M=[A[0]+(B[0]-A[0])/2, A[1]+(B[1]-A[1])/2.]; h = [B[1]-A[1], A[0]-B[0]]
            if use_arrow_shape:
                ffb.AddNodes(0, 0, M[0] + h[0]/3., M[1] + h[1]/3.)
            ffb.AddNodes(0, 0, B[0], B[1])
            # Rueckweg auf Innenkurve
            addNodesForCurves(ffb, iK)
            # Uebergang zur Außenkurve
            # mit kleiner Spitze (Mittelpunkt M, Orthogale h)
            A=iK[-1][3]; B=aK[0][0]; M=[A[0]+(B[0]-A[0])/2, A[1]+(B[1]-A[1])/2.]; h = [B[1]-A[1], A[0]-B[0]]
            if use_arrow_shape:
                ffb.AddNodes(0, 0, M[0] - h[0]/3., M[1] - h[1]/3.)
            ffb.AddNodes(0, 0, B[0], B[1])
            shp = ffb.ConvertToShape()
            
            # einen Shape-Punkt bewegen --> textframe wird auf shape zentriert
            shp.nodes.setposition( 1, aK[0][0][0] + 1, aK[0][0][1])
            shp.nodes.setposition( 1, aK[0][0][0],     aK[0][0][1])

        #slide.Shapes.Range(array('l', [i + shapeCount for i in range(0,len(aussenKurve))])).Group.select
        # slide.Shapes.Range(Array[int]([i+shapeCount+1 for i in range(0,len(aussenKurve))])).group().select()
        grp = pplib.last_n_shapes_on_slide(slide, num_segments)
        if num_segments > 1:
            grp = grp.group()
        grp.Name = "[BKT] Segmented Circle %s" % grp.id
        grp.select()


# =======================
# = UI MODEL AND WINDOW =
# =======================


class SegmentedCircleViewModel(bkt.ui.ViewModelSingleton):
    
    def __init__(self):
        super(SegmentedCircleViewModel, self).__init__()
        
        self._num_segments = 3
        self._radius = 4.0
        self._width = 25
        self._use_arrow_shape = True

    def set_values_based_on_shape(self, shape):
        try:
            if not SegmentedCircle.is_segmented_circle(shape):
                return
            self.num_segments, self.width, self._radius, arrow_shape = SegmentedCircle.determine_from_shape(shape)
            self.radius = pt_to_cm(self._radius)
            if arrow_shape:
                self.use_arrow_shape = True
            else:
                self.use_segment_shape = True
        except:
            logging.exception("failed to determine segmented circle parameters")
    
    
    @notify_property
    def num_segments(self):
        return self._num_segments

    @num_segments.setter
    def num_segments(self, value):
        self._num_segments = value
    
    @notify_property
    def radius(self):
        return self._radius

    @radius.setter
    def radius(self, value):
        self._radius = value
    
    @notify_property
    def width(self):
        return self._width

    @width.setter
    def width(self, value):
        self._width = value
    
    ## getters/setters for radio buttons
    
    @notify_property
    def use_arrow_shape(self):
        return self._use_arrow_shape

    @use_arrow_shape.setter
    def use_arrow_shape(self, value):
        self._use_arrow_shape = True
    
    @notify_property
    def use_segment_shape(self):
        return not self._use_arrow_shape

    @use_segment_shape.setter
    def use_segment_shape(self, value):
        self._use_arrow_shape = False



class SegmentedCircleWindow(bkt.ui.WpfWindowAbstract):
    _xamlname = 'circular_segments'
    _vm_class = SegmentedCircleViewModel
    
    def __init__(self, context, slide):
        super(SegmentedCircleWindow, self).__init__(context)
        
        self.ref_slide = slide

        try:
            self._vm.set_values_based_on_shape(context.shape)
        except:
            pass #e.g. nothing selected

    def cancel(self, sender, event):
        self.Close()
    
    def create_segments(self, sender, event):
        try:
            SegmentedCircle.create_segmented_circle(self.ref_slide, self._vm.num_segments, self._vm.width, cm_to_pt(self._vm.radius), use_arrow_shape=self._vm.use_arrow_shape)
        except:
            logging.error("Dialog action failed")
        finally:
            self.Close()