# -*- coding: utf-8 -*-
'''

@author: rdebeerst
'''


import bkt
import bkt.library.bezier as bezier
import bkt.library.algorithms as algorithms
import math

import logging
import traceback




# class SegmentedCircle(object):
    
#     @staticmethod
#     def createSegmentedCircleDialog(slide):
#         logging.debug("createSegmentedCircleDialog")
#         from circular_segments import SegmentedCircleWindow
#         SegmentedCircleWindow.create_and_show_dialog(slide)


# group_segmented_circle = bkt.ribbon.Group(
#     label=u"Kreissegmente",
#     image='segmented circle',
#     children=[
#         bkt.ribbon.Button(
#             label="Kreissegmente erstellen",
#             image='segmented circle',
#             size="large",
#             on_action=bkt.Callback(SegmentedCircle.createSegmentedCircleDialog)
#         ),
#         # bkt.ribbon.Button(
#         #     label="Kreissegmente erstellen",
#         #     image='segmented circle',
#         #     size="large",
#         #     on_action=bkt.Callback(segmented_circle.createSegmentedCircle)
#         # ),
#         # bkt.ribbon.Label(label='Segmente'),
#         # bkt.ribbon.Label(label='Breite'),
#         # bkt.ribbon.Label(label='Form'),
#         # bkt.ribbon.RoundingSpinnerBox(
#         #     label='Segmente', size_string='###',
#         #     show_label=False,
#         #     big_step = 1,
#         #     small_step = 1,
#         #     on_change=bkt.Callback(segmented_circle.set_segments_number),
#         #     get_text=bkt.Callback(segmented_circle.get_segments_number),
#         # ),
#         # bkt.ribbon.RoundingSpinnerBox(
#         #     label='Breite', size_string='###',
#         #     show_label=False,
#         #     round_int=True,
#         #     on_change=bkt.Callback(segmented_circle.set_width_percentage),
#         #     get_text=bkt.Callback(segmented_circle.get_width_percentage),
#         # ),
#         # bkt.ribbon.Box(box_style="horizontal", children=[
#         #     bkt.ribbon.ToggleButton(
#         #         label='Segmente', image='segmented circle segments', show_label=False,
#         #         on_toggle_action=bkt.Callback(segmented_circle.set_segment_type),
#         #         get_pressed=bkt.Callback(segmented_circle.get_segment_type)
#         #     ),
#         #     bkt.ribbon.ToggleButton(
#         #         label='Pfeile', image='segmented circle arrows', show_label=False,
#         #         on_toggle_action=bkt.Callback(segmented_circle.set_arrow_type),
#         #         get_pressed=bkt.Callback(segmented_circle.get_arrow_type)
#         #     )
#         # ])
#     ]
# )










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
    
    rotated = False
    fixed_radius = False

    midpoint = [0,0]
    radius = 1
    ptToCmFactor = 2.54 / 72;

    width = 10/ptToCmFactor
    height = 8.5/ptToCmFactor
    
    # DiagramCycleInsertClassic  DiagramRadialInsertClassic
    @classmethod
    def arrange_circular(cls, shapes):
        midpoint, width, height = cls.get_ellipse_params(shapes)
        cls.arrange_circular_wargs(shapes, midpoint, width, height)
    
    @classmethod
    def get_ellipse_params(cls, shapes):
        # compute weightend midpoint
        shape_midpoints = [ [s.left+s.width/2, s.top+s.height/2] for s in shapes]
        cls.midpoint = algorithms.mid_point(shape_midpoints)
        
        if cls.fixed_radius:
            cls.height = cls.width

        return cls.midpoint, cls.width, cls.height
        
        # # take x-value from top-most shape
        # #sorted_shapes = list(shapes)
        # #sorted_shapes.sort(key=lambda shape: shape.Top)
        # #cls.midpoint = [ sorted_shapes[0].left + sorted_shapes[0].width/2, cls.midpoint[1]]
        #
        # # round cm-value to 10^(-1)
        # #round_factor = cls.ptToCmFactor * 10
        # #cls.midpoint = [ round(cls.midpoint[0]* round_factor)/round_factor, round(cls.midpoint[1]* round_factor)/round_factor  ]
        #
        # # radius = distance midpoint to shape's midpoint
        # #        = difference, assumed that first shape is positioned at 12 o'clock
        # #radius = abs(cls.midpoint[1] - shape_midpoints[0][1])
        #
        # # determine points
        # points = cls.determine_points(shapes, cls.midpoint, 2, 2)
        # # FIXME: rearrange shape_midpoints with lowest distance from points
        #
        # # x-radius for elipses
        # # x-stretch faktor for every point
        # factors = [ 1.0* (shape_midpoints[i][0]-cls.midpoint[0])/(points[i][0]-cls.midpoint[0])  for i in range(0, len(shapes)) if (points[i][0]-cls.midpoint[0]) != 0]
        # # middle strech factor
        # radius_x = sum(factors)/len(factors)
        #
        # # y-stretch faktor for every point
        # factors = [ 1.0* (shape_midpoints[i][1]-cls.midpoint[1])/(points[i][1]-cls.midpoint[1])  for i in range(0, len(shapes)) if (points[i][1]-cls.midpoint[1]) != 0]
        # # middle strech factor
        # radius_y = sum(factors)/len(factors)
        #
        # return cls.midpoint, round(2*radius_x,2), round(2*radius_y,2)
        
    @classmethod
    def determine_ellipse_params(cls, shapes):
        # compute weightend midpoint
        shape_midpoints = [ [s.left+s.width/2, s.top+s.height/2] for s in shapes]
        cls.midpoint = algorithms.mid_point(shape_midpoints)

        # determine points
        points = cls.determine_points(shapes, cls.midpoint, 2, 2)

        # compute radius
        # x-stretch faktor for every point
        factors = [ 1.0* (shape_midpoints[i][0]-cls.midpoint[0])/(points[i][0]-cls.midpoint[0])  for i in range(0, len(shapes)) if (points[i][0]-cls.midpoint[0]) != 0]
        # middle strech factor
        radius_x = sum(factors)/len(factors)
        cls.width = round(2*radius_x,2)
        
        # y-stretch faktor for every point
        factors = [ 1.0* (shape_midpoints[i][1]-cls.midpoint[1])/(points[i][1]-cls.midpoint[1])  for i in range(0, len(shapes)) if (points[i][1]-cls.midpoint[1]) != 0]
        # middle strech factor
        radius_y = sum(factors)/len(factors)
        cls.height = round(2*radius_y,2)
        
        return cls.midpoint, cls.width, cls.height
    
    @classmethod
    def set_circ_width(cls, shapes, value):
        value = float(value)/cls.ptToCmFactor
        cls.width = value
        if cls.fixed_radius:
            cls.height = value
        cls.arrange_circular(shapes)
        # midpoint, width, height = cls.get_ellipse_params(shapes)
        # cls.arrange_circular_wargs(shapes, midpoint, float(value), height)
    
    @classmethod
    def get_circ_width(cls, shapes):
        # midpoint, width, height = cls.get_ellipse_params(shapes)
        # return round(width*cls.ptToCmFactor, 2)
        return round(cls.width*cls.ptToCmFactor, 2)
    
    @classmethod
    def set_circ_height(cls, shapes, value):
        value = float(value)/cls.ptToCmFactor
        cls.height = value
        if cls.fixed_radius:
            cls.width = value
        cls.arrange_circular(shapes)
        # midpoint, width, height = cls.get_ellipse_params(shapes)
        # # move midpoint, so y-coordinate of top-shape is not moved
        # # midpoint[1] += (float(value) - height)/2
        # cls.arrange_circular_wargs(shapes, midpoint, width, float(value))
    
    @classmethod
    def get_circ_height(cls, shapes):
        # midpoint, width, height = cls.get_ellipse_params(shapes)
        # return round(height*cls.ptToCmFactor, 2)
        return round(cls.height*cls.ptToCmFactor, 2)
    
    @classmethod
    def determine_points(cls, shapes, midpoint, width, height):
        # Nullpunkt als Mittelpunkt verwenden, mit Radius 1
        segments = bezier.kreisSegmente(len(shapes), 1, [0,0])
        points = [ [s[0][0][0], s[0][0][1]] for s in segments]
        
        # Segments starten rechts im Kreis
        # Alle Punkte um 90 Grad nach links drehen, damit erstes Objekt oben steht
        points = [ algorithms.rotate_point(p[0], p[1], 0, 0, 90)  for p in points]
        
        # Punkte skalieren (Höhe/Breite)
        points = [ [ width/2 * p[0], height/2 * p[1] ]   for p in points]
        
        # # Punkte  verschieben auf Mitte der Folie
        # slide = shapes[0].Parent
        # points = [ [ p[0] + slide.Parent.PageSetup.SlideWidth/2, p[1] + slide.Parent.PageSetup.SlideHeight/2 ]   for p in points]
        
        # Punkte verschieben anhand midpoint
        points = [ [ p[0] + midpoint[0], p[1] + midpoint[1] ]   for p in points]
        
        return points
    
    @classmethod
    def arrange_circular_wargs(cls, shapes, midpoint, width, height):
        points = cls.determine_points(shapes, midpoint, width, height)
        # print "new points"
        # print points
        
        for i in range(0, len(shapes)):
            shapes[i].left = points[i][0] - shapes[i].width /2
            shapes[i].top  = points[i][1] - shapes[i].height /2
            
            if cls.rotated == True:
                shapes[i].rotation = 360/len(shapes)*i
            else:
                shapes[i].rotation = shapes[0].rotation
        
    
    
    @classmethod
    def arrange_circular_rotated(cls, pressed):
        cls.rotated = (pressed == True)

    @classmethod
    def arrange_circular_rotated_pressed(cls):
        return cls.rotated == True

    @classmethod
    def arrange_circular_fixed(cls, pressed):
        cls.fixed_radius = (pressed == True)

    @classmethod
    def arrange_circular_fixed_pressed(cls):
        return cls.fixed_radius == True



group_circlify = bkt.ribbon.Group(
	label=u"Kreisanordnung",
    image="circlify",
	children=[
		bkt.ribbon.Button(label="Kreisförmig anordnen", image="circlify", #image_mso="DiagramRadialInsertClassic",
            size='large',
            supertip="Ausgewählte Shapes werden Kreis-förmig angeordnet, entsprechend der eingestellten Breite/Höhe.\nDie Reihenfolge der Shapes ist abhängig von der Selektionsreihenfolge: das zuerst selektierte Shape wird auf 12 Uhr positioniert, die weiteren Shapes folgen im Urzeigersinn.",
			on_action=bkt.Callback(CircularArrangement.arrange_circular, shapes=True, shapes_min=3),
			get_enabled="PythonGetEnabled"
		),
		bkt.ribbon.RoundingSpinnerBox(label="Breite", round_cm=True, image_mso="ShapeWidth",
            show_label=False,
            supertip="Breite der Ellipse (Diagonale) für die Kreisanordnung",
			on_change=bkt.Callback(CircularArrangement.set_circ_width, shapes=True, shapes_min=3),
			get_enabled="PythonGetEnabled",
			get_text=bkt.Callback(CircularArrangement.get_circ_width, shapes=True, shapes_min=3),
		),
		bkt.ribbon.RoundingSpinnerBox(label="Höhe", round_cm=True, image_mso="ShapeHeight",
            show_label=False,
            supertip="Höhe der Ellipse (Diagonale) für die Kreisanordnung",
			on_change=bkt.Callback(CircularArrangement.set_circ_height, shapes=True, shapes_min=3),
			get_enabled="PythonGetEnabled",
			get_text=bkt.Callback(CircularArrangement.get_circ_height, shapes=True, shapes_min=3),
		),
        bkt.ribbon.Menu(
            label="Optionen",
            item_size="large",
            children=[
                bkt.ribbon.MenuSeparator(title="Shapes drehen:"),
                bkt.ribbon.ToggleButton(
                    label="Rotation an/aus",
                    image_mso="ObjectRotateFree",
                    description="Objekte in der Kreisanordnung entsprechend ihrer Position im Kreis rotieren",
                    on_toggle_action=bkt.Callback(CircularArrangement.arrange_circular_rotated),
                    get_pressed=bkt.Callback(CircularArrangement.arrange_circular_rotated_pressed)
                ),
                bkt.ribbon.MenuSeparator(title="Radius gleichstellen:"),
                bkt.ribbon.ToggleButton(
                    label="Breite = Höhe",
                    description="Bei Veränderung der Höhe wird auch die Breite geändert und umgekehrt",
                    image_mso="ShapeDonut",
                    on_toggle_action=bkt.Callback(CircularArrangement.arrange_circular_fixed),
                    get_pressed=bkt.Callback(CircularArrangement.arrange_circular_fixed_pressed)
                ),
                bkt.ribbon.MenuSeparator(title="Radius setzen:"),
                bkt.ribbon.Button(
                    label="Aktuellen Radius bestimmen",
                    description="Es wird versucht den aktuellen Radius der ausgewählten Shapes zu bestimmen",
                    image_mso="DiagramRadialInsertClassic",
                    on_action=bkt.Callback(CircularArrangement.determine_ellipse_params, shapes=True, shapes_min=3),
			        get_enabled="PythonGetEnabled",
                ),
            ]
        ),
	]	
)



bkt.powerpoint.add_tab(bkt.ribbon.Tab(
    id="bkt_powerpoint_toolbox_extensions",
    insert_before_mso="TabHome",
    label=u'Toolbox 3/3',
    # get_visible defaults to False during async-startup
    get_visible=bkt.Callback(lambda:True),
    children = [
        group_circlify,
        # group_segmented_circle
    ]
), extend=True)

