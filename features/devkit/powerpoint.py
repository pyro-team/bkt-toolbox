# -*- coding: utf-8 -*-
'''
Created on 26.02.2020

@author: fstallmann
'''

from __future__ import absolute_import

import bkt
import bkt.library.powerpoint as pplib
import bkt.library.algorithms as algos


class Shapetags(object):
    @classmethod
    def presentation_tags(cls, presentation):
        cls.show_message(
            cls.get_tags_for_message(presentation)
        )

    @classmethod
    def slide_tags(cls, slides):
        cls.show_message(
            "\r\n\r\n".join(
                cls.get_tags_for_message(slide)
                for slide in slides
            )
        )
        
    @classmethod
    def shape_tags(cls, shapes):
        cls.show_message(
            "\r\n\r\n".join(
                cls.get_tags_for_message(shape)
                for shape in shapes
            )
        )

    @staticmethod
    def get_tags_for_message(obj):
        if(hasattr(obj, 'SlideId')):
            name = "Slide No %s (ID %s)" % (obj.SlideNumber, obj.SlideId)
        elif(hasattr(obj, 'Id')):
            name = "Shape %s (ID %s)" % (obj.Name, obj.Id)
        else:
            name = "Object"

        if obj.Tags.Count > 0:
            msg = "Found {} tag(s) for {}".format(obj.Tags.Count, name)
            for idx in range(1, obj.Tags.Count+1):
                msg += "\r\n{:30}: {}".format(obj.Tags.Name(idx), obj.Tags.Value(idx))
        else:
            msg = "No tags for " + name + " found!"

        return msg

    @staticmethod
    def show_message(msg):
        import bkt.console
        bkt.console.show_message(msg)

    @staticmethod
    def remove_all_tags(obj):
        for idx in range(obj.Tags.Count,0,-1):
            obj.tags.delete(obj.Tags.Name(idx))
        
    @classmethod
    def remove_presentation_tags(cls, presentation):
        cls.remove_all_tags(presentation)

    @classmethod
    def remove_slide_tags(cls, slides):
        for slide in slides:
            cls.remove_all_tags(slide)
        
    @classmethod
    def remove_shape_tags(cls, shapes):
        for shape in shapes:
            cls.remove_all_tags(shape)


tags_gruppe = bkt.ribbon.Group(
    id="bkt_pptdev_tags_group",
    label='Tags',
    image_mso='NameManager',
    children = [
        bkt.ribbon.Button(
            id = 'dev_presentation_tags',
            label="Presentation-Tags",
            show_label=True,
            image_mso='NameDefine',
            on_action=bkt.Callback(Shapetags.presentation_tags, presentation=True),
            get_enabled = bkt.get_enabled_auto,
        ),
        bkt.ribbon.Button(
            id = 'dev_slide_tags',
            label="Slide-Tags",
            show_label=True,
            image_mso='NameDefine',
            on_action=bkt.Callback(Shapetags.slide_tags, slides=True),
            get_enabled = bkt.get_enabled_auto,
        ),
        bkt.ribbon.Button(
            id = 'dev_shape_tags',
            label="Shape-Tags",
            show_label=True,
            image_mso='NameDefine',
            on_action=bkt.Callback(Shapetags.shape_tags, shapes=True),
            get_enabled = bkt.get_enabled_auto,
        ),

        bkt.ribbon.Button(
            id = 'dev_presentation_tags-remove',
            label="Remove all presentation tags",
            show_label=False,
            image_mso='Delete',
            on_action=bkt.Callback(Shapetags.remove_presentation_tags, presentation=True),
            get_enabled = bkt.get_enabled_auto,
        ),
        bkt.ribbon.Button(
            id = 'dev_slide_tags-remove',
            label="Remove all slide tags",
            show_label=False,
            image_mso='Delete',
            on_action=bkt.Callback(Shapetags.remove_slide_tags, slides=True),
            get_enabled = bkt.get_enabled_auto,
        ),
        bkt.ribbon.Button(
            id = 'dev_shape_tags-remove',
            label="Remove all shape tags",
            show_label=False,
            image_mso='Delete',
            on_action=bkt.Callback(Shapetags.remove_shape_tags, shapes=True),
            get_enabled = bkt.get_enabled_auto,
        ),
    ]
)


class ShapeNodes(object):
    @staticmethod
    def draw_nodes(slide, nodes, tag="SHAPE"):
        size=10
        for i,node in enumerate(nodes, start=1):
            s=slide.shapes.AddShape(
                9, #msoShapeOval
                node[0]-size/2, node[1]-size/2,
                size,size
            )
            s.textframe.AutoSize = 0
            s.textframe.WordWrap = 0
            s.textframe.textrange.text = str(i)
            s.textframe.textrange.font.size=8
            s.tags.add("BKT_DEVKIT_NODE", tag)

    @classmethod
    def draw_shape_nodes(cls, shape, slide):
        #convert into freeform by adding and deleting in order to acces points
        dummy = shape.duplicate()
        dummy.left, dummy.top = shape.left, shape.top
        dummy.nodes.insert(1,0,0,0,0)
        dummy.nodes.delete(2)
        shape_nodes = [(node.points[0,0], node.points[0,1]) for node in dummy.nodes]
        dummy.delete()
        cls.draw_nodes(slide, shape_nodes, "SHAPE")

        # mid = algos.mid_point(shape_nodes)
        # cls.draw_nodes(slide, [mid], "SHAPE")
    
    @classmethod
    def draw_bounding_nodes(cls, shape, slide):
        shape_nodes = algos.get_bounding_nodes(shape)
        cls.draw_nodes(slide, shape_nodes, "BOUNDING")
    
    @classmethod
    def draw_locpin_nodes(cls, shape, slide):
        shape = pplib.wrap_shape(shape)
        all_nodes = shape.get_locpin_nodes()
        cls.draw_nodes(slide, all_nodes, "LOCPIN")

    @staticmethod
    def remove_nodes(slide, tag="SHAPE"):
        for shape in list(iter(slide.shapes)):
            try:
                if shape.tags("BKT_DEVKIT_NODE") == tag:
                    shape.delete()
            except:
                continue
    
    @classmethod
    def remove_shape_nodes(cls, slide):
        cls.remove_nodes(slide, "SHAPE")
    
    @classmethod
    def remove_bounding_nodes(cls, slide):
        cls.remove_nodes(slide, "BOUNDING")
    
    @classmethod
    def remove_locpin_nodes(cls, slide):
        cls.remove_nodes(slide, "LOCPIN")

nodes_gruppe = bkt.ribbon.Group(
    id="bkt_pptdev_nodes_group",
    label='Nodes',
    image_mso='RecursiveSection',
    children = [
        bkt.ribbon.Button(
            label="Shape Nodes",
            show_label=True,
            image_mso='RecursiveSection',
            on_action=bkt.Callback(ShapeNodes.draw_shape_nodes, shape=True, slide=True),
            get_enabled = bkt.get_enabled_auto,
        ),
        bkt.ribbon.Button(
            label="Bounding Nodes",
            show_label=True,
            image_mso='RecursiveSection',
            on_action=bkt.Callback(ShapeNodes.draw_bounding_nodes, shape=True, slide=True),
            get_enabled = bkt.get_enabled_auto,
        ),
        bkt.ribbon.Button(
            label="Locpin Nodes",
            show_label=True,
            image_mso='RecursiveSection',
            on_action=bkt.Callback(ShapeNodes.draw_locpin_nodes, shape=True, slide=True),
            get_enabled = bkt.get_enabled_auto,
        ),

        bkt.ribbon.Button(
            label="Remove all shape nodes",
            show_label=False,
            image_mso='Delete',
            on_action=bkt.Callback(ShapeNodes.remove_shape_nodes, slide=True),
            get_enabled = bkt.get_enabled_auto,
        ),

        bkt.ribbon.Button(
            label="Remove all bounding nodes",
            show_label=False,
            image_mso='Delete',
            on_action=bkt.Callback(ShapeNodes.remove_bounding_nodes, slide=True),
            get_enabled = bkt.get_enabled_auto,
        ),

        bkt.ribbon.Button(
            label="Remove all locpin nodes",
            show_label=False,
            image_mso='Delete',
            on_action=bkt.Callback(ShapeNodes.remove_locpin_nodes, slide=True),
            get_enabled = bkt.get_enabled_auto,
        ),
    ]
)


class ShapePoints(object):
    form = None

    @staticmethod
    def _get_edittype(node):
        try:
            return str(node.editingType)
        except:
            return "-2"

    @classmethod
    def _get_pointlist(cls, shape):
        if not shape.Type == pplib.MsoShapeType['msoFreeform']:
            #convert shape into freeform by adding and deleting node (not sure if this is required)
            shape.Nodes.Insert(1, 0, 0, 0, 0) #msoSegmentLine, msoEditingAuto, x, y
            shape.Nodes.Delete(2)
            # shape.Nodes.SetPosition(1, shape.Left, shape.Top)
        
        pointlist = "["
        for i,node in enumerate(shape.nodes, start=1):
            if i > 1:
                pointlist += ","
            pointlist += "\r\n"
            pointlist += ' {"i":' + str(i)
            pointlist += ', "x":' + str(node.points[0,0])
            pointlist += ', "y":' + str(node.points[0,1])
            pointlist += ', "segmentType": ' + str(node.segmentType)
            pointlist += ', "editingType": ' + cls._get_edittype(node)
            pointlist += '}'
        pointlist += "\r\n]"

        return pointlist

    @classmethod
    def display_points(cls, shape):
        
        def json_callback(json_points):
            cls.change_points(shape, json_points=json_points)
        
        cls.form = bkt.console.show_input(cls._get_pointlist(shape), json_callback)

    @classmethod
    def change_points(cls, shape, json_points=None):
        import json

        points = json.loads(json_points)
        
        # richtige Anzahl Punkte
        nodes_count = shape.nodes.count
        while len(points) > nodes_count:
            shape.nodes.insert(nodes_count, 0,0,  0.0, 0.0)
        while len(points) < nodes_count:
            shape.nodes.delete(nodes_count)
        
        for p in reversed(points):
            index = p['i']

            if p['segmentType'] in (0,1) and shape.nodes[index].segmentType != p['segmentType']:
                shape.nodes.setSegmentType(index, p['segmentType'])
            
            if index > 1 and index < nodes_count and p['editingType'] in (0,1,2,3) and cls._get_edittype(shape.nodes[index]) != p['editingType']:
                shape.nodes.setEditingType(index, p['editingType'])
            
            shape.nodes.setPosition(index, p['x'], p['y'])

        # update list (as changing editingType may result in new or deleted nodes)
        cls.form.input.Text = cls._get_pointlist(shape)

shape_points_gruppe = bkt.ribbon.Group(
    label="Shape Points",
    children=[
        bkt.ribbon.Button(
            label='Shape Points',
            size="large",
            imageMso='ObjectEditPoints',
            on_action=bkt.Callback(ShapePoints.display_points)
        )
    ]
)

powerpoint_groups = [tags_gruppe, nodes_gruppe, shape_points_gruppe]