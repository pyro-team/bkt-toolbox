# -*- coding: utf-8 -*-
'''
Created on 01.08.2022

@author: fstallmann
'''



from contextlib import contextmanager #for flip and rotation correction

import bkt.library.bezier as bezier
import bkt.library.powerpoint as pplib
pt_to_cm = pplib.pt_to_cm
cm_to_pt = pplib.cm_to_pt

# ===================================================================
# = functionality to create segmented circles or circular processes =
# ===================================================================

@contextmanager
def flip_and_rotation_correction(shape):
    try:
        st_rot, st_wdt, st_hgh, st_lar = shape.rotation, shape.width, shape.height, shape.lockaspectratio
        st_fliph, st_flipv = shape.horizontalflip, shape.verticalflip
        if st_fliph:
            shape.Flip(0) #msoFlipHorizontal
        if st_flipv:
            shape.Flip(1) #msoFlipVertical
        shape.rotation = 0
        shape.lockaspectratio = 0
        shape.width = shape.height = max(shape.width, shape.height)
        yield shape
    finally:
        shape.rotation, shape.width, shape.height = st_rot, st_wdt, st_hgh
        shape.lockaspectratio = st_lar
        if st_fliph:
            shape.Flip(0) #msoFlipHorizontal
        if st_flipv:
            shape.Flip(1) #msoFlipVertical


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
        from math import hypot

        with flip_and_rotation_correction(shape):
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

            width = hypot(x1-x2,y1-y2)*2
            width_percentage = min(100,max(1,round(width/shape.width *100)))

            #spacing
            x3 = shape.left+shape.width
            y3 = shape.top+shape.height/2
            spacing = round(hypot(x1-x3,y1-y3) / shape.width * 100)

            size_outer = shape.width/2

        return num_segments, width_percentage, size_outer, use_arrow_shape, spacing

    @staticmethod
    def updated_segmented_circle(shape, num_segments, width_percentage, size_outer=100, use_arrow_shape=False, spacing=0):
        assert num_segments > 1

        size_inner = size_outer * (1-width_percentage/100.)
        M = [shape.left+shape.width/2, shape.top+shape.height/2]
        aussenKurve = bezier.kreisSegmente(num_segments, size_outer, M, spacing)
        innenKurve  = bezier.kreisSegmente(num_segments, size_inner, M, spacing)

        def addNodesForCurves(ffb, kurven):
            for k in kurven:
                #Insert (Index, SegmentType, EditingType, X1, Y1, X2, Y2, X3, Y3)
                ffb.nodes.insert(ffb.nodes.count, 1, 1, k[1][0], k[1][1], k[2][0], k[2][1], k[3][0], k[3][1])

        group = pplib.GroupManager(shape)
        group.prepare_ungroup()

        childs = list(group.child_items)

        for i in range(0,len(aussenKurve)):
            # Außen- und Innenkurve jeweils als Liste von Bezierkurven
            aK = aussenKurve[i]
            iK = innenKurve[i]; iK.reverse()
            for k in iK:
                k.reverse()
            
            # Nächstes existierendes Shape nehmen oder letztes duplizieren
            try:
                ffb = childs.pop(0)
            except IndexError:
                ffb = group.child_items[-1].duplicate()

            # Alle Nodes löschen (2 bleiben immer über)
            for _ in range(ffb.nodes.count+1):
                ffb.nodes.seteditingtype(2,1)
                # ffb.nodes.setsegmenttype(2,1)
                ffb.nodes.delete(2)
            
            assert ffb.nodes.count == 2
            
            # Erste Node richtig setzen
            ffb.nodes.setposition( 1, aK[0][0][0],     aK[0][0][1])
            # Hinweg auf Außenkurve
            addNodesForCurves(ffb, aK)
            # Übrig gebliebene Node löschen
            ffb.nodes.delete(2)
            # Uebergang zur Innenkurve
            # mit kleiner Spitze
            A=aK[-1][3]; B=iK[0][0]; M=[A[0]+(B[0]-A[0])/2, A[1]+(B[1]-A[1])/2.]; h = [B[1]-A[1], A[0]-B[0]]
            if use_arrow_shape:
                ffb.nodes.insert(ffb.nodes.count, 0, 0, M[0] + h[0]/3., M[1] + h[1]/3.)
            ffb.nodes.insert(ffb.nodes.count, 0, 0, B[0], B[1])
            # Rueckweg auf Innenkurve
            addNodesForCurves(ffb, iK)
            # Uebergang zur Außenkurve
            # mit kleiner Spitze (Mittelpunkt M, Orthogale h)
            A=iK[-1][3]; B=aK[0][0]; M=[A[0]+(B[0]-A[0])/2, A[1]+(B[1]-A[1])/2.]; h = [B[1]-A[1], A[0]-B[0]]
            if use_arrow_shape:
                ffb.nodes.insert(ffb.nodes.count, 0, 0, M[0] - h[0]/3., M[1] - h[1]/3.)
            # ffb.nodes.insert(ffb.nodes.count, 0, 0, B[0], B[1])
            
            # einen Shape-Punkt bewegen --> textframe wird auf shape zentriert
            # ffb.nodes.setposition( 1, aK[0][0][0] + 1, aK[0][0][1])
            # ffb.nodes.setposition( 1, aK[0][0][0],     aK[0][0][1])
        
        # Bei Reduzierung der Segmente übrig gebliebene Shapes löschen
        for child in childs:
            child.delete()
        
        # Gruppe wiederherstellen und auswählen
        group.refresh()
        group.select()


    @staticmethod
    def create_segmented_circle(slide, num_segments, width_percentage, size_outer=100, use_arrow_shape=False, spacing=0):
        # aussenKurve = bezier.bezierKreisNRM(2,100,[100,100])
        # innenKurve = bezier.bezierKreisNRM(2,75,[100,100])
        size_inner = size_outer * (1-width_percentage/100.)
        M = [slide.master.width/2, slide.master.height/2]
        aussenKurve = bezier.kreisSegmente(num_segments, size_outer, M, spacing)
        innenKurve  = bezier.kreisSegmente(num_segments, size_inner, M, spacing)

        def addNodesForCurves(ffb, kurven):
            for k in kurven:
                #AddNodes(SegmentType, EditingType, X1, Y1, X2, Y2, X3, Y3)
                ffb.AddNodes(1, 1, k[1][0], k[1][1], k[2][0], k[2][1], k[3][0], k[3][1])

        for i in range(0,len(aussenKurve)):
            # Außen- und Innenkurve jeweils als Liste von Bezierkurven
            aK = aussenKurve[i]
            iK = innenKurve[i]; iK.reverse()
            for k in iK:
                k.reverse()
            # Hinweg auf Außenkurve
            #BuildFreeForm(EditingType, X1, Y1)
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