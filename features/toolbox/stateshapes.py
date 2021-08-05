# -*- coding: utf-8 -*-
'''
Created on 21.12.2017

@author: fstallmann
'''

from __future__ import absolute_import

import logging

from System import Array

import bkt
import bkt.library.powerpoint as pplib

from bkt import dotnet
Drawing = dotnet.import_drawing()



class AbstractStateShape(object):
    def switch_state(self, shape, delta=0, pos=None):
        return shape
    def set_color_fill1(self, shape, color=None):
        pass
    def set_color_fill2(self, shape, color=None):
        pass
    def set_color_line(self, shape, color=None):
        pass

class DefaultStateShape(AbstractStateShape):
    def switch_state(self, shape, delta=0, pos=None):
        # ungroup shape, to get list of groups inside grouped items
        # ungrouped_shapes = shape.Ungroup()
        # shapes = list(iter(ungrouped_shapes))
        group = pplib.GroupManager(shape).ungroup()
        shapes = group.child_items
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
        group.regroup()
        # group.select(replace=False) #NOTE: this causes trouble if shape is corrupt, better to return and select all at the end
        return group.shape
        # grp = ungrouped_shapes.Group()
        # grp.Tags.Add(bkt.contextdialogs.BKT_CONTEXTDIALOG_TAGKEY, cls.BKT_DIALOG_TAG)
        # try:
        #     #sometimes throws "Invalid request.  To select a shape, its view must be active.", e.g. right after duplicating the shape
        #     grp.Select(replace=False)
        # except:
        #     grp.Select()

    def set_color_fill1(self, shape, color=None):
        ref_shape = shape.GroupItems[1]
        ref_color = ref_shape.Fill.ForeColor.RGB
        ref_visible = ref_shape.Fill.Visible == -1
        for s in shape.GroupItems:
            if ref_visible and s.Fill.ForeColor.RGB == ref_color or not ref_visible and s.Fill.Visible == 0:
                if color:
                    s.Fill.Visible = -1
                    color.apply_to_color_obj(s.Fill.ForeColor)
                else:
                    s.Fill.Visible = 0

    def set_color_fill2(self, shape, color=None):
        ref_shape = shape.GroupItems[1]
        ref_color = ref_shape.Fill.ForeColor.RGB if ref_shape.Fill.Visible == -1 else None
        for s in shape.GroupItems:
            if s.Fill.Visible == -1 and s.Fill.ForeColor.RGB != ref_color:
                if color:
                    s.Fill.Visible = -1
                    color.apply_to_color_obj(s.Fill.ForeColor)
                else:
                    s.Fill.Visible = 0

    def set_color_line(self, shape, color=None):
        ref_color = None
        for s in shape.GroupItems:
            if ref_color is None:
                if s.Line.visible == -1:
                    ref_color = s.Line.ForeColor.RGB
                else:
                    continue
            if s.Line.visible == -1 and s.Line.ForeColor.RGB == ref_color:
                if color:
                    s.Line.Visible = -1
                    color.apply_to_color_obj(s.Line.ForeColor)
                else:
                    s.Line.Visible = 0


class StateShape(object):
    BKT_DIALOG_TAG = 'BKT_DIALOG_STATESHAPE'

    _default_stateshape = DefaultStateShape()
    _special_stateshapes = dict()
    
    @classmethod
    def register_special_stateshape(cls, name, special_object):
        cls._special_stateshapes[name] = special_object

    @classmethod
    def is_convertable_to_state_shape(cls, shapes):
        try:
            if len(shapes) > 1:
                return not any(cls.is_state_shape(s) for s in shapes)
            else:
                shape = shapes[0]
                return shape.Type == pplib.MsoShapeType['msoGroup'] and not cls.is_state_shape(shape)
        except:
            return False

    @classmethod
    def convert_to_state_shape(cls, shapes, special_name=None):
        if len(shapes) > 1:
            shape = pplib.shapes_to_range(shapes).group()
        else:
            shape = shapes[0]
        try:
            shape.Tags.Add(bkt.contextdialogs.BKT_CONTEXTDIALOG_TAGKEY, cls.BKT_DIALOG_TAG)
            if special_name and special_name in cls._special_stateshapes:
                shape.Tags.Add(cls.BKT_DIALOG_TAG, special_name)
        except:
            logging.exception("Error converting to state stape")
        
        cls.switch_states([shape], pos=0)

    @classmethod
    def is_state_shape(cls, shape):
        return pplib.TagHelper.has_tag(shape, bkt.contextdialogs.BKT_CONTEXTDIALOG_TAGKEY, cls.BKT_DIALOG_TAG)
        # return shape.Type == pplib.MsoShapeType['msoGroup']
    
    @classmethod
    def are_state_shapes(cls, shapes):
        return all(cls.is_state_shape(s) for s in shapes)

    @classmethod
    def get_stateshape_object(cls, shape):
        special_name = pplib.TagHelper.get_tag(shape, cls.BKT_DIALOG_TAG)
        try:
            return cls._special_stateshapes[special_name]
        except KeyError:
            return cls._default_stateshape

    
    @classmethod
    def switch_state(cls, shape, delta=0, pos=None):
        if not cls.is_state_shape(shape):
            raise ValueError("Shape is not a state shape")
    
        stateshape_obj = cls.get_stateshape_object(shape)
        return stateshape_obj.switch_state(shape, delta, pos)

    
    @classmethod
    def switch_states(cls, shapes, **kwargs):
        resulting_shapes = []
        for shape in shapes:
            try:
                resulting_shapes.append( cls.switch_state(shape, **kwargs) )
            except:
                logging.exception("Statehape error swichting state")
                continue
        if resulting_shapes:
            pplib.shapes_to_range(resulting_shapes).select()


    @classmethod
    def reset_state(cls, shapes):
        cls.switch_states(shapes, pos=0)

    @classmethod
    def next_state(cls, shapes):
        cls.switch_states(shapes, delta=1)

    @classmethod
    def previous_state(cls, shapes):
        cls.switch_states(shapes, delta=-1)

    @classmethod
    def set_state(cls, shapes, value):
        cls.switch_states(shapes, pos=int(value))

    # @classmethod
    # def get_show_all(cls, shape):
    #     return cls.is_state_shape(shape) and shape.GroupItems.Range(None).Visible == -1

    # @classmethod
    # def toggle_show_all(cls, shape, pressed):
    #     if not pressed:
    #         cls.switch_state(shape, pos=0)
    #     else:
    #         cls.show_all(shape)

    @classmethod
    def show_all(cls, shape):
        if shape.Type == pplib.MsoShapeType['msoGroup']:
            ungrouped_shapes = shape.Ungroup()
            for s in list(iter(ungrouped_shapes)):
                s.visible = True
            ungrouped_shapes.Group().Select()


    @classmethod
    def set_color_fill1(cls, shapes, color=None):
        for shape in shapes:
            stateshape_obj = cls.get_stateshape_object(shape)
            stateshape_obj.set_color_fill1(shape, color)

    @classmethod
    def set_color_fill2(cls, shapes, color=None):
        for shape in shapes:
            stateshape_obj = cls.get_stateshape_object(shape)
            stateshape_obj.set_color_fill2(shape, color)

    @classmethod
    def set_color_line(cls, shapes, color=None):
        for shape in shapes:
            stateshape_obj = cls.get_stateshape_object(shape)
            stateshape_obj.set_color_line(shape, color)


    @classmethod
    def set_color_fill_rgb1(cls, shapes, color):
        cls.set_color_fill1(shapes, pplib.PPTColor(pplib.PPTColor.COLOR_RGB, color_rgb=color))
    @classmethod
    def set_color_fill_theme1(cls, shapes, color_index, brightness):
        cls.set_color_fill1(shapes, pplib.PPTColor(pplib.PPTColor.COLOR_THEME, color_index=color_index, brightness=brightness))
    @classmethod
    def set_color_fill_none1(cls, shapes):
        cls.set_color_fill1(shapes, pplib.PPTColor(pplib.PPTColor.COLOR_NONE))


    @classmethod
    def set_color_fill_rgb2(cls, shapes, color):
        cls.set_color_fill2(shapes, pplib.PPTColor(pplib.PPTColor.COLOR_RGB, color_rgb=color))
    @classmethod
    def set_color_fill_theme2(cls, shapes, color_index, brightness):
        cls.set_color_fill2(shapes, pplib.PPTColor(pplib.PPTColor.COLOR_THEME, color_index=color_index, brightness=brightness))


    @classmethod
    def set_color_line_rgb(cls, shapes, color):
        cls.set_color_line(shapes, pplib.PPTColor(pplib.PPTColor.COLOR_RGB, color_rgb=color))
    @classmethod
    def set_color_line_theme(cls, shapes, color_index, brightness):
        cls.set_color_line(shapes, pplib.PPTColor(pplib.PPTColor.COLOR_THEME, color_index=color_index, brightness=brightness))
    @classmethod
    def set_color_line_none(cls, shapes):
        cls.set_color_line(shapes, pplib.PPTColor(pplib.PPTColor.COLOR_NONE))




class LikertScaleTags(pplib.BKTTag):
    TAG_NAME = "BKT_LIKERT_CONFIG"

class LikertScale(AbstractStateShape):
    SHAPE_TAG = "BKT_LIKERT"

    # def _determine_current_pos(self, shape):
    #     scales = list(iter(shape.GroupItems))
    #     colors = sum(1 for s in scales if s.Fill.ForeColor.RGB == scales[0].Fill.ForeColor.RGB)
    #     return colors
    
    def _update_color_fills(self, shape):
        with LikertScaleTags(shape.Tags) as tags:
            pos = tags.get("position", 0)
            try:
                col1 = pplib.PPTColor.from_color_tuple( tags["color_fill1"] )
            except KeyError:
                col1 = pplib.PPTColor.new_theme(14)
            try:
                col2 = pplib.PPTColor.from_color_tuple( tags["color_fill2"] )
            except KeyError:
                col2 = pplib.PPTColor.new_theme(16)

        for i,s in enumerate(iter(shape.GroupItems)):
            color = col2 if i < pos else col1
            if color:
                s.Fill.Visible = -1
                color.apply_to_color_obj(s.Fill.ForeColor)
            else:
                s.Fill.Visible = 0

    def switch_state(self, shape, delta=0, pos=None):
        scales = shape.GroupItems.Count+1
        with LikertScaleTags(shape.Tags) as tags:
            if pos is None:
                pos = tags.get("position", 0)
            pos = (pos + delta) % scales
            tags["position"] = pos

        self._update_color_fills(shape)
        return shape

    def set_color_fill1(self, shape, color=None):
        with LikertScaleTags(shape.Tags) as tags:
            tags["color_fill1"] = color.get_color_tuple()
        self._update_color_fills(shape)

    def set_color_fill2(self, shape, color=None):
        with LikertScaleTags(shape.Tags) as tags:
            tags["color_fill2"] = color.get_color_tuple()
        self._update_color_fills(shape)

    def set_color_line(self, shape, color=None):
        for shape in shape.GroupItems:
            if color:
                shape.Line.Visible = -1
                color.apply_to_color_obj(shape.Line.ForeColor)
            else:
                shape.Line.Visible = 0


class LikertScaleGallery(bkt.ribbon.Gallery):
    #settings in powerpoint
    spacing = 5
    size = 20

    #settings for  gallery
    likert_sizes = [3,4,5]
    likert_columns = len(likert_sizes)
    likert_shapes = {1: "Quadratisch", 9: "Kreisförmig", 92: "Sternförmig"} #rectangle, oval, star
    likert_buttons = []
    
    def __init__(self, **kwargs):
        self.likert_buttons = [
            [n,m]
            for n in self.likert_shapes.keys()
            for m in self.likert_sizes
        ]

        # parent_id = kwargs.get('id') or ""
        my_kwargs = dict(
            label = 'Likert-Scale',
            image = 'likert',
            columns = self.likert_columns,
            screentip="Likert-Scale als Wechselshape einfügen",
            supertip="Eine Likert-Scale als Wechselshape einfügen. Über die Wechselshape-Funktionen kann der Füllstand, sowie die Farben verändert werden.",
            item_width=max(self.likert_sizes)*16,
            item_height=16,
        )
        my_kwargs.update(kwargs)

        super(LikertScaleGallery, self).__init__(**my_kwargs)

    def get_item_count(self):
        return len(self.likert_buttons)

    def get_item_screentip(self, index):
        return "%sen %ser-Likert-Scale einfügen" % (self.likert_shapes[self.likert_buttons[index][0]], self.likert_buttons[index][1])

    def get_item_image(self, index):
        return self.get_likert_image( count=self.likert_buttons[index][1], shape=self.likert_buttons[index][0] )

    def on_action_indexed(self, selected_item, index, slide):
        self._create_stateshape_scale(slide, self.likert_buttons[index][0], self.likert_buttons[index][1]),

    def get_likert_image(self, size_factor=2, count=3, shape=1):
        img = Drawing.Bitmap(max(self.likert_sizes)*16*size_factor, 16*size_factor)
        color_black = Drawing.Color.Black
        color_grey  = Drawing.Color.Gray
        color_white = Drawing.Color.White
        g = Drawing.Graphics.FromImage(img)
        
        #Draw smooth rectangle/ellipse
        g.SmoothingMode = Drawing.Drawing2D.SmoothingMode.AntiAlias

        pen    = Drawing.Pen(color_black,1*size_factor)
        brush1 = Drawing.SolidBrush(color_grey)
        brush2 = Drawing.SolidBrush(color_white)
        star_points = [(0,6),(5,8),(5,12),(8,9),(12,10),(10,6),(12,2),(8,3),(5,0),(5,4)]
        
        left = 2
        for i in range(count):
            brush = brush2 if i>0 else brush1
            if shape == 92: #star
                points = Array[Drawing.Point]([Drawing.Point(left+l*size_factor,t*size_factor) for t,l in star_points])
                g.FillPolygon(brush, points)
                g.DrawPolygon(pen, points)
            elif shape == 9: #oval
                g.FillEllipse(brush, left, 2, 12*size_factor, 12*size_factor) #left, top, width, height
                g.DrawEllipse(pen, left, 2, 12*size_factor, 12*size_factor) #left, top, width, height
            else: #fallback shape=1 rectangle
                g.FillRectangle(brush, left, 2, 12*size_factor, 12*size_factor) #left, top, width, height
                g.DrawRectangle(pen, left, 2, 12*size_factor, 12*size_factor) #left, top, width, height
            left += 16*size_factor
        return img

    # def _create_single_scale(self, slide, shape_type=1, state=0, total=3):
    #     # shapecount = slide.Shapes.Count
    #     left = 90
    #     for i in range(total):
    #         left += self.size + self.spacing
    #         s = slide.Shapes.AddShape( shape_type, left, 100, self.size, self.size )
    #         # s.Line.Weight = 0.75
    #         # s.Line.ForeColor.RGB = self.color_line
    #         s.Line.Visible = -1
    #         s.Line.ForeColor.ObjectThemeColor = 13 #msoThemeColorText1
    #         s.Fill.Visible = -1
    #         if i < state:
    #             # s.Fill.ForeColor.RGB = self.color_filled
    #             s.Fill.ForeColor.ObjectThemeColor = 16 #msoThemeColorBackground2
    #         else:
    #             # s.Fill.ForeColor.RGB = self.color_empty
    #             s.Fill.ForeColor.ObjectThemeColor = 14 #msoThemeColorBackground1

    #     # grp = slide.Shapes.Range(Array[int](range(shapecount+1, shapecount+1+total))).group()
    #     grp = pplib.last_n_shapes_on_slide(slide, total).group()
    #     return grp

    def _create_stateshape_scale(self, slide, shape_type, total):
        left = 90
        for _ in range(total):
            # self._create_single_scale(slide, shape_type, i, total)
            left += self.size + self.spacing
            s = slide.Shapes.AddShape( shape_type, left, 100, self.size, self.size )
            s.Line.Visible = -1
            s.Line.ForeColor.ObjectThemeColor = 13 #msoThemeColorText1
        
        grp = pplib.last_n_shapes_on_slide(slide, total).group()
        grp.LockAspectRatio = -1
        StateShape.convert_to_state_shape([grp], LikertScale.SHAPE_TAG)


#Slider-Element
#Progressbar


class CheckBox(AbstractStateShape):
    SHAPE_TAG = "BKT_CHECKBOX"

    checkmark_groups = [
        dict(font="Wingdings", chars=[u'\xfc', u'\xfb', u'\x6c', u'\x6e', '']),
        dict(font="Arial Unicode", chars=[u'\u2713', u'\u2717', u'\u2715', '']),
    ]

    def switch_state(self, shape, delta=0, pos=None):
        textrange = shape.TextFrame2.TextRange
        font = textrange.Font.Name
        text = textrange.Text
        for d in self.checkmark_groups:
            if d["font"] == font:
                chars = d["chars"]
                break
        else:
            bkt.message.error("Fehler! Unbekannter Checkbox-Font.")
            return
        if pos is None:
            try:
                pos = chars.index(text)
            except ValueError:
                pos = 0
        pos = (pos + delta) % len(chars)
        textrange.Text = chars[pos]
        # shape.select(False)
        return shape

    def set_color_fill1(self, shape, color=None):
        if color:
            shape.Fill.Visible = -1
            color.apply_to_color_obj(shape.Fill.ForeColor)
        else:
            shape.Fill.Visible = 0

    def set_color_fill2(self, shape, color=None):
        font_fill = shape.textframe2.textrange.font.fill
        if color:
            font_fill.Visible = -1
            color.apply_to_color_obj(font_fill.ForeColor)
        else:
            font_fill.Visible = 0

    def set_color_line(self, shape, color=None):
        if color:
            shape.Line.Visible = -1
            color.apply_to_color_obj(shape.Line.ForeColor)
        else:
            shape.Line.Visible = 0


class CheckBoxGallery(bkt.ribbon.Gallery):
    #settings in powerpoint
    size = 16

    #settings for gallery
    shape_types = {1:"Quadratisch", 9:"Kreisförmig"} #rectangle, oval
    checkmark_columns = 4
    checkmark_maxchars = 3
    item_size = 24
    items = []
    
    def __init__(self, **kwargs):
        self.items = [ dict(shape_type=t, style=s, font_group=c) for c in CheckBox.checkmark_groups for t in self.shape_types.keys() for s in ['light', 'dark'] ]

        # parent_id = kwargs.get('id') or ""
        my_kwargs = dict(
            label="Checkbox",
            image_mso="FormControlCheckBox",
            screentip='Checkbox erstellen',
            supertip="Füge ein Checkbox-Symbol als Wechselshape ein, welches interaktiv geändert werden kann.",
            columns=self.checkmark_columns,
            item_width=self.checkmark_maxchars*self.item_size,
            item_height=self.item_size,
        )
        my_kwargs.update(kwargs)

        super(CheckBoxGallery, self).__init__(**my_kwargs)

    def get_item_count(self):
        return len(self.items)

    def get_item_screentip(self, index):
        return "%se Checkbox mit %s einfügen" % (self.shape_types[self.items[index]["shape_type"]], self.items[index]["font_group"]["font"])

    def get_item_image(self, index):
        item = self.items[index]
        return self.get_checkmark_image( item["shape_type"], item["style"], item["font_group"] )

    def on_action_indexed(self, selected_item, index, slide):
        item = self.items[index]
        self._insert_checkmark(slide, item["shape_type"], item["style"], item["font_group"])

    def _insert_checkmark(self, slide, shape_type, style, font_group):
        # for char in font_group["chars"]:
        #     self._insert_single_box(slide, shape_type, style, (font_group["font"], char))
        # #add empty box
        # self._insert_single_box(slide, shape_type, style)
        # #group and make stateshape
        # grp = pplib.last_n_shapes_on_slide(slide, 1+len(font_group["chars"])).group()

        grp = self._insert_single_box(slide, shape_type, style, (font_group["font"], font_group["chars"][0]))
        grp.LockAspectRatio = -1
        grp.Name = "[BKT] Checkbox %s" % grp.id
        StateShape.convert_to_state_shape([grp], CheckBox.SHAPE_TAG)
    
    def _insert_single_box(self, slide, shape_type=1, style='light', font_char=None):
        s = slide.Shapes.AddShape( shape_type, 100, 100, self.size, self.size )

        if style == "dark":
            col_background = 13 #msoThemeColorText1
            col_foreground = 14 #msoThemeColorBackground1
        else:
            col_background = 14 #msoThemeColorBackground1
            col_foreground = 13 #msoThemeColorText1

        s.Line.Visible = -1
        s.Line.ForeColor.ObjectThemeColor = col_foreground
        s.Fill.Visible = -1
        s.Fill.ForeColor.ObjectThemeColor = col_background
        s.LockAspectRatio = -1

        textframe = s.TextFrame2
        textframe.AutoSize     = 0
        textframe.MarginBottom = 0
        textframe.MarginLeft   = 0
        textframe.MarginRight  = 0
        textframe.MarginTop    = 0
        # textframe.HorizontalAnchor = 2
        textframe.VerticalAnchor   = 3

        textrange = textframe.TextRange
        textrange.ParagraphFormat.Bullet.Visible = False
        textrange.ParagraphFormat.Alignment = 2 #ppAlignCenter
        textrange.ParagraphFormat.SpaceBefore = 0
        textrange.ParagraphFormat.SpaceAfter  = 0
        textrange.ParagraphFormat.SpaceWithin = 1
        textrange.ParagraphFormat.BaselineAlignment = 4 #default is 5=TextBottom, but 4=TextTop looks more centered
        textrange.Font.Bold   = 0
        textrange.Font.Italic = 0
        textrange.Font.Size   = self.size
        textrange.Font.Fill.ForeColor.ObjectThemeColor = col_foreground

        if font_char:
            textrange.Font.Name = font_char[0]
            textrange.InsertSymbol(font_char[0], ord(font_char[1]), -1) #symbol: FontName, CharNumber (decimal), Unicode=True
        # s.Visible = visible
        return s

    def get_checkmark_image(self, shape_type, style, font_group):
        size_factor = 2
        base_size = self.item_size*size_factor

        img = Drawing.Bitmap(self.checkmark_maxchars*base_size, base_size)
        g = Drawing.Graphics.FromImage(img)
        
        #Draw smooth rectangle/ellipse
        g.SmoothingMode = Drawing.Drawing2D.SmoothingMode.AntiAlias

        if style == "dark":
            pen_border = Drawing.Pen(Drawing.Color.White,1*size_factor)
            brush_fill = Drawing.Brushes.Black
            text_brush = Drawing.Brushes.White
        else:
            pen_border = Drawing.Pen(Drawing.Color.Black,1*size_factor)
            brush_fill = Drawing.Brushes.White
            text_brush = Drawing.Brushes.Black

        font = Drawing.Font(font_group["font"], base_size-6*size_factor, Drawing.GraphicsUnit.Pixel)

        left = 2
        for char in font_group["chars"]:

            if shape_type == 9: #oval
                g.FillEllipse(brush_fill, left, 2, base_size-4, base_size-4) #left, top, width, height
                g.DrawEllipse(pen_border, left, 2, base_size-4, base_size-4) #left, top, width, height
            else: #fallback shape=1 rectangle
                g.FillRectangle(brush_fill, left, 2, base_size-4, base_size-4) #left, top, width, height
                g.DrawRectangle(pen_border, left, 2, base_size-4, base_size-4) #left, top, width, height
        
            # set string format
            strFormat = Drawing.StringFormat()
            strFormat.Alignment = Drawing.StringAlignment.Center
            strFormat.LineAlignment = Drawing.StringAlignment.Center
            
            g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAlias
            g.DrawString(char,
                        font, text_brush,
                        Drawing.RectangleF(left-1, 4, base_size, base_size), 
                        strFormat)
            
            left += base_size

        return img



likert_button = LikertScaleGallery(id="likert_insert")
checkbox_button = CheckBoxGallery(id="checkbox_insert")

StateShape.register_special_stateshape(CheckBox.SHAPE_TAG, CheckBox())
StateShape.register_special_stateshape(LikertScale.SHAPE_TAG, LikertScale())


def stateshape_fill1_gallery(**kwargs):
    return bkt.ribbon.ColorGallery(
                    label = 'Farbe 1 (Hintergrund) ändern',
                    image_mso = 'ShapeFillColorPicker',
                    screentip="Hintergrundfarbe eines Wechsel-Shapes ändern",
                    supertip="Passt die Hintergrundfarbe aller Shapes im Wechsel-Shape an. Die Hintergrundfarbe ist die Farbe des zuerst gefundenen Shapes.",
                    on_rgb_color_change   = bkt.Callback(StateShape.set_color_fill_rgb1, shapes=True),
                    on_theme_color_change = bkt.Callback(StateShape.set_color_fill_theme1, shapes=True),
                    # get_selected_color    = bkt.Callback(StateShape.get_selected_color1, shapes=True),
                    get_enabled           = bkt.Callback(StateShape.are_state_shapes),
                    children=[
                        bkt.ribbon.Button(
                            label="Kein Hintergrund",
                            supertip="Wechsel-Shape Hintergrundfarbe auf transparent setzen",
                            on_action=bkt.Callback(StateShape.set_color_fill_none1, shapes=True),
                        ),
                    ],
                    **kwargs
                )

def stateshape_fill2_gallery(**kwargs):
    return bkt.ribbon.ColorGallery(
                    label = 'Farbe 2 (Vordergrund) ändern',
                    image_mso = 'ShapeFillColorPicker',
                    screentip="Vordergrundfarbe eines Wechsel-Shapes ändern",
                    supertip="Passt die Vordergrundfarbe aller Shapes im Wechsel-Shape an. Die Vordergrundfarbe ist jede Farbe ungleich der Hintergrundfarbe.",
                    on_rgb_color_change   = bkt.Callback(StateShape.set_color_fill_rgb2, shapes=True),
                    on_theme_color_change = bkt.Callback(StateShape.set_color_fill_theme2, shapes=True),
                    # get_selected_color    = bkt.Callback(StateShape.get_selected_color2, shapes=True),
                    get_enabled           = bkt.Callback(StateShape.are_state_shapes),
                    **kwargs
                )

def stateshape_line_gallery(**kwargs):
    return bkt.ribbon.ColorGallery(
                    label = 'Linie ändern',
                    image_mso = 'ShapeOutlineColorPicker',
                    screentip="Linie eines Wechsel-Shapes ändern",
                    supertip="Passt die Linienfarbe aller Shapes im Wechsel-Shape an, die der ersten gefundenen Linienfarbe entsprechen.",
                    on_rgb_color_change   = bkt.Callback(StateShape.set_color_line_rgb, shapes=True),
                    on_theme_color_change = bkt.Callback(StateShape.set_color_line_theme, shapes=True),
                    # get_selected_color    = bkt.Callback(StateShape.get_selected_line, shapes=True),
                    get_enabled           = bkt.Callback(StateShape.are_state_shapes),
                    children=[
                        bkt.ribbon.Button(
                            label="Keine Linie",
                            supertip="Wechsel-Shape Linie auf transparent setzen",
                            on_action=bkt.Callback(StateShape.set_color_line_none, shapes=True),
                        ),
                    ],
                    **kwargs
                )



stateshape_gruppe = bkt.ribbon.Group(
    id="bkt_stateshape_group",
    label='Wechsel-Shapes',
    image_mso='GroupSmartArtQuickStyles',
    children = [
        bkt.ribbon.SplitButton(
            id="stateshape_convert_splitbutton",
            # size="large",
            children=[
                bkt.ribbon.Button(
                    id="stateshape_convert",
                    label=u"Konvertieren",
                    image_mso='GroupSmartArtQuickStyles',
                    screentip="Gruppierte Shapes in ein Wechselshape konvertieren",
                    supertip="Bei gruppierten Shapes (Wechsel-Shapes) kann zwischen den Shapes innerhalb der Gruppe gewechselt werden, d.h. es ist immer nur ein Shape der Gruppe sichtbar. Dies ist bspw. nützlich für Ampeln, Skalen, etc.",
                    on_action=bkt.Callback(StateShape.convert_to_state_shape),
                    get_enabled=bkt.Callback(StateShape.is_convertable_to_state_shape),
                ),
                bkt.ribbon.Menu(
                    label="Wechselshapes-Menü",
                    supertip="In Wechselshapes konvertieren oder wieder alle Shapes sichtbar machen",
                    children=[
                        bkt.ribbon.Button(
                            id="stateshape_convert2",
                            label=u"In Wechselshape konvertieren",
                            image_mso='GroupSmartArtQuickStyles',
                            screentip="Gruppierte Shapes in ein Wechselshape konvertieren",
                            supertip="Bei gruppierten Shapes (Wechsel-Shapes) kann zwischen den Shapes innerhalb der Gruppe gewechselt werden, d.h. es ist immer nur ein Shape der Gruppe sichtbar. Dies ist bspw. nützlich für Ampeln, Skalen, etc.",
                            on_action=bkt.Callback(StateShape.convert_to_state_shape),
                            get_enabled=bkt.Callback(StateShape.is_convertable_to_state_shape),
                        ),
                        bkt.ribbon.MenuSeparator(),
                        # bkt.ribbon.ToggleButton(
                        bkt.ribbon.Button(
                            id="stateshape_show_all",
                            label=u"Alle Shapes wieder anzeigen",
                            screentip="Alle Shapes sichtbar machen",
                            supertip="Mit diesem Button können die Shapes innerhalb der Wechselshape-Gruppe eingeblendet werden.",
                            # image_mso='GroupSmartArtQuickStyles',
                            # get_pressed=bkt.Callback(StateShape.get_show_all),
                            # on_toggle_action=bkt.Callback(StateShape.toggle_show_all),
                            on_action=bkt.Callback(StateShape.show_all),
                            get_enabled=bkt.Callback(StateShape.is_state_shape),
                        ),
                    ]
                )
            ]
        ),
        # bkt.ribbon.Separator(),
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
            supertip="Die Farben von Wechselshapes anpassen",
            image_mso="RecolorColorPicker",
            children=[
                stateshape_fill1_gallery(id="stateshape_color_fill1"),
                stateshape_fill2_gallery(id="stateshape_color_fill2"),
                stateshape_line_gallery(id="stateshape_color_line"),
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
