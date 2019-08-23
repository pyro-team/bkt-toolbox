# -*- coding: utf-8 -*-

import os.path
import bkt.ui
notify_property = bkt.ui.notify_property

import logging

import bkt.library.bezier as bezier
import bkt.library.algorithms as algorithms
import bkt.library.powerpoint as pplib


# ===================================================================
# = functionality to create segmented circles or circular processes =
# ===================================================================

class SegmentedCircle(object):
    
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
        pplib.last_n_shapes_on_slide(slide, num_segments).group().select()


# =======================
# = UI MODEL AND WINDOW =
# =======================


class SegmentedCircleViewModel(bkt.ui.ViewModelSingleton):
    
    def __init__(self):
        super(SegmentedCircleViewModel, self).__init__()
        
        self.num_segments = 3
        self.width = 25
        self.use_arrow_shape = True
    
    
    ## getters/setters for radio buttons
    
    @notify_property
    def use_segment_shape(self):
        return not self.use_arrow_shape

    @use_segment_shape.setter
    def use_segment_shape(self, value):
        self.use_arrow_shape = not value
        self.OnPropertyChanged('use_arrow_shape')



class SegmentedCircleWindow(bkt.ui.WpfWindowAbstract):
    _filename = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'circular_segments.xaml')
    _vm_class = SegmentedCircleViewModel
    
    def __init__(self, slide):
        self.ref_slide = slide
        super(SegmentedCircleWindow, self).__init__()

    def cancel(self, sender, event):
        self.Close()
    
    def create_segments(self, sender, event):
        try:
            SegmentedCircle.create_segmented_circle(self.ref_slide, self._vm.num_segments, self._vm.width, use_arrow_shape=self._vm.use_arrow_shape)
        except:
            logging.error("Dialog action failed")
        finally:
            self.Close()