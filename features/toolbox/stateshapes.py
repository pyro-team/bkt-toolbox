# -*- coding: utf-8 -*-
'''
Created on 21.12.2017

@author: fstallmann
'''

import bkt
import bkt.library.powerpoint as pplib

from System import Array

from bkt import dotnet
Drawing = dotnet.import_drawing()


class StateShape(object):
    
    @classmethod
    def is_state_shape(cls, shape):
        return shape.Type == pplib.MsoShapeType['msoGroup']
    
    @classmethod
    def are_state_shapes(cls, shapes):
        return all(cls.is_state_shape(s) for s in shapes)
    
    @classmethod
    def switch_state(cls, shape, delta=0, pos=None):
        # ungroup shape, to get list of groups inside grouped items
        ungrouped_shapes = shape.Ungroup()
        shapes = list(iter(ungrouped_shapes))
        # shapes.sort(key=lambda s: s.ZOrderPosition)
        # pos = min(max(pos, len(shapes)-1), -(len(shapes)-1)) #pos between -/+ number of shapes in group
        for i, s in enumerate(shapes):
            if pos is None and s.visible == -1:
                pos = i
            s.visible = False
        # for s in shapes[:pos]:
        #     s.ZOrder(0) #0=msoBringToFront, 1=msoSendToBack

        pos = (pos + delta) % len(shapes)
        shapes[pos].visible = True
        grp = ungrouped_shapes.Group()
        try:
            #sometimes throws "Invalid request.  To select a shape, its view must be active.", e.g. right after duplicating the shape
            grp.Select(replace=False)
        except:
            grp.Select()


    @classmethod
    def reset_state(cls, shapes):
        for shape in shapes:
            cls.switch_state(shape, pos=0)

    @classmethod
    def next_state(cls, shapes):
        for shape in shapes:
            cls.switch_state(shape, delta=1)

    @classmethod
    def previous_state(cls, shapes):
        for shape in shapes:
            cls.switch_state(shape, delta=-1)

    @classmethod
    def set_state(cls, shapes, value):
        value = int(value)
        for shape in shapes:
            cls.switch_state(shape, pos=value)

    @classmethod
    def get_show_all(cls, shape):
        return cls.is_state_shape(shape) and shape.GroupItems.Range(None).Visible == -1

    @classmethod
    def toggle_show_all(cls, shape, pressed):
        if not pressed:
            cls.switch_state(shape, pos=0)
        else:
            ungrouped_shapes = shape.Ungroup()
            for s in list(iter(ungrouped_shapes)):
                s.visible = True
            ungrouped_shapes.Group().Select()

    @classmethod
    def set_color_fill_rgb(cls, shapes, color):
        for shape in shapes:
            for s in list(iter(shape.GroupItems)):
                if s.Fill.visible == -1 and s.Fill.ForeColor.RGB != 16777215: #white
                    s.Fill.ForeColor.RGB = color

    @classmethod
    def set_color_fill_theme(cls, shapes, color_index, brightness):
        for shape in shapes:
            for s in list(iter(shape.GroupItems)):
                if s.Fill.visible == -1 and s.Fill.ForeColor.RGB != 16777215: #white
                    s.Fill.ForeColor.ObjectThemeColor = color_index
                    s.Fill.ForeColor.Brightness = brightness

    @classmethod
    def set_color_line_rgb(cls, shapes, color):
        for shape in shapes:
            for s in list(iter(shape.GroupItems)):
                if s.Line.visible == -1 and s.Line.ForeColor.RGB != 16777215: #white
                    s.Line.ForeColor.RGB = color

    @classmethod
    def set_color_line_theme(cls, shapes, color_index, brightness):
        for shape in shapes:
            for s in list(iter(shape.GroupItems)):
                if s.Line.visible == -1 and s.Line.ForeColor.RGB != 16777215: #white
                    s.Line.ForeColor.ObjectThemeColor = color_index
                    s.Line.ForeColor.Brightness = brightness

    @staticmethod
    def show_help():
        bkt.helpers.message("TODO: show help file, image, or something...")


class LikertScale(object):
    spacing = 5
    size = 20
    color_line = 0
    color_filled = 14540253
    color_empty = 16777215

    likert_sizes = [3,4,5]
    likert_columns = len(likert_sizes)
    likert_shapes = {1: "Quadratisch", 9: "Kreisförmig", 92: "Sternförmig"} #rectangle, oval, star
    likert_buttons = [
        [n,m]
        for n in likert_shapes.keys()
        for m in likert_sizes
    ]
    
    @classmethod
    def get_item_count(cls):
        return len(cls.likert_buttons)

    @classmethod
    def get_item_screentip(cls, index):
        return "%sen %ser-Likert-Scale einfügen" % (cls.likert_shapes[cls.likert_buttons[index][0]], cls.likert_buttons[index][1])

    @classmethod
    def get_item_image(cls, index):
        return LikertScale.get_likert_image( count=cls.likert_buttons[index][1], shape=cls.likert_buttons[index][0] )

    @classmethod
    def on_action_indexed(cls, selected_item, index, slide):
        cls._create_stateshape_scale(slide, cls.likert_buttons[index][0], cls.likert_buttons[index][1]),

    @staticmethod
    def get_likert_image(size=16, count=3, shape=1):
        img = Drawing.Bitmap(5*16, size)
        color_black = Drawing.ColorTranslator.FromOle(0)
        color_grey  = Drawing.ColorTranslator.FromOle(14540253)
        color_white = Drawing.ColorTranslator.FromOle(16777215)
        g = Drawing.Graphics.FromImage(img)
        
        #Draw smooth rectangle/ellipse
        g.SmoothingMode = Drawing.Drawing2D.SmoothingMode.AntiAlias

        pen    = Drawing.Pen(color_black,1)
        brush1 = Drawing.SolidBrush(color_grey)
        brush2 = Drawing.SolidBrush(color_white)
        star_points = [(0,6),(5,8),(5,12),(8,9),(12,10),(10,6),(12,2),(8,3),(5,0),(5,4)]
        
        left = 2
        for i in range(count):
            brush = brush2 if i>0 else brush1
            if shape == 92: #star
                points = Array[Drawing.Point]([Drawing.Point(left+l,t) for t,l in star_points])
                g.FillPolygon(brush, points)
                g.DrawPolygon(pen, points)
            elif shape == 9: #oval
                g.FillEllipse(brush, left, 2, 12, 12) #left, top, width, height
                g.DrawEllipse(pen, left, 2, 12, 12) #left, top, width, height
            else: #fallback shape=1 rectangle
                g.FillRectangle(brush, left, 2, 12, 12) #left, top, width, height
                g.DrawRectangle(pen, left, 2, 12, 12) #left, top, width, height
            left += 16
        return img

    @classmethod
    def _create_single_scale(cls, slide, shape_type=1, state=0, total=3, visible=0):
        shapecount = slide.Shapes.Count
        left = 90
        for i in range(total):
            left += cls.size + cls.spacing
            s = slide.Shapes.AddShape( shape_type, left, 100, cls.size, cls.size )
            s.Line.ForeColor.RGB = cls.color_line
            if i < state:
                s.Fill.ForeColor.RGB = cls.color_filled
            else:
                s.Fill.ForeColor.RGB = cls.color_empty

        grp = slide.Shapes.Range(Array[int](range(shapecount+1, shapecount+1+total))).group()
        grp.Visible = visible
        return grp

    @classmethod
    def _create_stateshape_scale(cls, slide, shape_type, total):
        shapecount = slide.Shapes.Count
        for i in range(total+1):
            cls._create_single_scale(slide, shape_type, i, total)
        
        slide.Shapes.Range(shapecount+1).visible = -1 #make first visible
        grp = slide.Shapes.Range(Array[int](range(shapecount+1, shapecount+1+total+1))).group()
        grp.select()



likert_button = bkt.ribbon.Gallery(
        label = 'Likert-Scale',
        image = 'likert',
        screentip="Likert-Scale als Wechselshape einfügen",
        supertip="Eine Likert-Scale als Wechselshape einfügen. Über die Wechselshape-Funktionen kann der Füllstand, sowie die Farben verändert werden.",
        columns=str(LikertScale.likert_columns),
        on_action_indexed = bkt.Callback(LikertScale.on_action_indexed, slide=True),
        get_item_count    = bkt.Callback(LikertScale.get_item_count),
        get_item_screentip = bkt.Callback(LikertScale.get_item_screentip),
        # get_item_supertip = bkt.Callback(lambda index: "Passe den Füllstand eines Harvey-Balls entsprechend der Auswahl an."),
        get_item_image    = bkt.Callback(LikertScale.get_item_image),
    )


stateshape_gruppe = bkt.ribbon.Group(
    id="bkt_stateshape_group",
    label='Wechsel-Shapes',
    image_mso='GroupSmartArtQuickStyles',
    children = [
        bkt.ribbon.ToggleButton(
            id="stateshape_show_all",
            label=u"Alle anzeigen",
            screentip="Alle Shapes sichtbar machen",
            supertip="Bei gruppierten Shapes (Wechsel-Shapes) kann zwischen den Shapes innerhalb der Gruppe gewechselt werden, d.h. es ist immer nur ein Shape der Gruppe sichtbar. Dies ist bspw. nützlich für Ampeln, Skalen, etc.\n\nMit diesem Button können die Shapes innerhalb der Gruppe ein- und ausgeblendet werden.",
            size="large",
            image_mso='GroupSmartArtQuickStyles',
            get_pressed=bkt.Callback(StateShape.get_show_all),
            on_toggle_action=bkt.Callback(StateShape.toggle_show_all),
            get_enabled=bkt.Callback(StateShape.is_state_shape),
        ),
        bkt.ribbon.Separator(),
        # bkt.ribbon.LabelControl(label="Wechsel: "),
        bkt.ribbon.Box(box_style="horizontal", children=[
            bkt.ribbon.Button(
                id="stateshape_reset",
                image_mso="Undo",
                label=u"Zurücksetzen",
                show_label=False,
                screentip="Auf erstes Shape zurücksetzen",
                supertip="Setzt alle Wechsel-Shapes auf den ersten Status, d.h. das erste Shape der Gruppe zurück.",
                on_action=bkt.Callback(StateShape.reset_state),
                get_enabled=bkt.Callback(StateShape.are_state_shapes),
            ),
            bkt.ribbon.Button(
                id="stateshape_prev",
                image_mso="PreviousResource",
                label=u'Vorheriges',
                show_label=False,
                screentip="Vorheriges Shape",
                supertip="Wechselt zum vorherigen Status (d.h. Shape in der Gruppe) des Wechsel-Shapes.",
                on_action=bkt.Callback(StateShape.previous_state),
                get_enabled=bkt.Callback(StateShape.are_state_shapes),
            ),
            # bkt.ribbon.EditBox(
            #     id="stateshape_set",
            #     label="Position",
            #     show_label=False,
            #     size_string="#",
            #     on_change=bkt.Callback(StateShape.set_state),
            #     get_enabled=bkt.Callback(StateShape.are_state_shapes),
            #     get_text=bkt.Callback(lambda: None),
            # ),
            bkt.ribbon.Button(
                id="stateshape_next",
                image_mso="NextResource",
                label=u"Nächstes",
                # show_label=False,
                screentip="Nächstes Shape",
                supertip="Wechselt zum nächsten Status (d.h. Shape in der Gruppe) des Wechsel-Shapes.",
                on_action=bkt.Callback(StateShape.next_state),
                get_enabled=bkt.Callback(StateShape.are_state_shapes),
            )
        ]),
        bkt.ribbon.Menu(
            id="stateshape_color_menu",
            label="Farbe ändern",
            image_mso="RecolorColorPicker",
            children=[
                bkt.ribbon.ColorGallery(
                    id="stateshape_color_fill",
                    label = 'Hintergrund ändern',
                    image_mso = 'ShapeFillColorPicker',
                    screentip="Farbe eines Wechsel-Shapes ändern",
                    supertip="Passt die Hintergrundfarbe aller Shapes im Wechsel-Shape an. Dabei werden nicht gefüllte und weiß gefüllte Shapes nicht verändert.",
                    on_rgb_color_change   = bkt.Callback(StateShape.set_color_fill_rgb, shapes=True),
                    on_theme_color_change = bkt.Callback(StateShape.set_color_fill_theme, shapes=True),
                    # get_selected_color    = bkt.Callback(StateShape.get_selected_color, shapes=True),
                    get_enabled           = bkt.Callback(StateShape.are_state_shapes),
                ),
                bkt.ribbon.ColorGallery(
                    id="stateshape_color_line",
                    label = 'Linie ändern',
                    image_mso = 'ShapeOutlineColorPicker',
                    screentip="Linie eines Wechsel-Shapes ändern",
                    supertip="Passt die Linienfarbe aller Shapes im Wechsel-Shape an. Dabei werden Shape ohne Linie oder mit weißer Linie nicht verändert.",
                    on_rgb_color_change   = bkt.Callback(StateShape.set_color_line_rgb, shapes=True),
                    on_theme_color_change = bkt.Callback(StateShape.set_color_line_theme, shapes=True),
                    # get_selected_color    = bkt.Callback(StateShape.get_selected_color, shapes=True),
                    get_enabled           = bkt.Callback(StateShape.are_state_shapes),
                ),
            ]
        ),
        # bkt.ribbon.Button(
        #     id="stateshape_help",
        #     image_mso="Help",
        #     label=u"Anleitung",
        #     on_action=bkt.Callback(StateShape.show_help),
        #     # get_enabled=bkt.Callback(StateShape.are_state_shapes),
        # ),
        # likert_button,
    ]
)