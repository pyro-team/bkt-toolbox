# -*- coding: utf-8 -*-
'''
Created on 12.09.2014

@author: cschmitt
'''

import bkt

from bkt.library.visio import unwrap, VisioShape
from System.Runtime.InteropServices import COMException

'''
Zur Aktivierung untenstehende Zeilen zur config.txt hinzufügen.
Danach Visio neu starten oder BKT neu laden (Entwicklertools).
Die eigentliche Shape-Identifikation findet in der Methode
ShapeIdentifier.handle_shape() statt. Dort könnte man auch eine 
komplexe Mappinglogik implementieren, die die Shapes Objekten aus
einem Datenmodell zuordnet (z.B. zuvor aus Excel-Tabelle geladen).  

pythonpath = <path-to-svn-root>\development\python\trunk\samples\viso-shape-ident
module = identify_shapes
'''

class ShapeIdentifier(object):
    def __init__(self, shapes, recurse=True):
        self.shapes = shapes
        self.recurse = recurse
        self.no_id = []
        self.duplicates = []
        self.by_id = {}
        
    def run(self):
        if self.recurse:
            collection = self.traverse()
        else:
            collection = self.shapes
    
        for shape in collection:
            self.handle_shape(shape)
        
    def expand(self, shape):
        try:
            return [VisioShape(c) for c in unwrap(shape).Shapes]
        except COMException:
            return []
        
    def handle_shape(self, shape):
        try:
            shape_id = shape.shape_data.ShapeID
            # Zugriff auch dictionary-like möglich (dann müsste gemäß Python-Konvetion ine KeyError gefangen werden)
            # shape_id = shape.shape_data["ShapeID"]
        except AttributeError:
            self.no_id.append(shape)
            return
            
        if shape_id in ('', None):
            self.no_id.append(shape)
        elif shape_id in self.by_id:
            self.duplicates.append(shape)
        else:
            self.by_id[shape_id] = shape
        
    def traverse(self):
        stack = list(self.shapes)
        while stack:
            current = stack.pop()
            yield current
            stack.extend(self.expand(current))
            

class IDGroup(bkt.FeatureContainer):
    @bkt.image_mso('HappyFace')
    @bkt.arg_page_shapes
    @bkt.large_button('Show Shape IDs')
    def show_ids(self, page_shapes):
        sidfer = ShapeIdentifier(page_shapes)
        sidfer.run()
        
        def lines():
            yield 'found shape IDs:'
            for sid in sidfer.by_id:
                yield sid
            yield 'duplicate shapes IDs:'
            for shape in sidfer.duplicates:
                yield shape.shape_data.ShapeID
            yield '%d shapes without id' % len(sidfer.no_id)

        
        bkt.console.show_message('\r\n'.join(lines()))

    @bkt.image_mso('HappyFace')
    @bkt.arg_page_shapes
    @bkt.arg_context
    @bkt.large_button('Mark Duplicates')
    def mark_duplicates(self, context, page_shapes):
        sidfer = ShapeIdentifier(page_shapes)
        sidfer.run()

        w = context.app.ActiveWindow
        w.DeselectAll()

        def mark(shape):
            shape.linepattern = 1 
            shape.linecolor = (255,0,0)
            shape.stroke = 5
        
        for shape in sidfer.duplicates:
            master = sidfer.by_id[shape.shape_data.ShapeID]
            
            mark(shape)
            mark(master)
            
            try:
                w.Select(unwrap(shape),2)
                w.Select(unwrap(master),2)
            except COMException:
                continue

@bkt.visio
class IDTab(bkt.Tab):
    label = "Shape ID Sample"
    groups = [IDGroup]        
