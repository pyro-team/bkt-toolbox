# -*- coding: utf-8 -*-
'''
Created on 06.07.2016

@author: rdebeerst
'''

import bkt
import bkt.library.powerpoint as powerpoint
import bkt.library.graphics as glib

from bkt import dotnet
Drawing = dotnet.import_drawing()

from ..harvey import HarveyBallsUi


class HarveyBalls(HarveyBallsUi):

    # _line_weight = 0.5
    # _line_color  = 0
    # _fill_color  = 16777215
    
    # =========================
    # = Harvey Ball erstellen =
    # =========================
    
    def _add_tags(self, shape, denominator=None):
        shape.Tags.Add(self.BKT_HARVEY_DIALOG_TAG, self.BKT_HARVEY_VERSION)
        shape.Tags.Add(bkt.contextdialogs.BKT_CONTEXTDIALOG_TAGKEY, self.BKT_HARVEY_DIALOG_TAG)

        if denominator is not None:
            shape.Tags.Add(self.BKT_HARVEY_DENOM_TAG, denominator)
    
    def create_harvey_ball(self, context, slide, fill=0.25):
        # shapeCount = slide.shapes.count

        circ = slide.shapes.addshape( powerpoint.MsoAutoShapeType['msoShapeOval'] , 100, 100, 30,30)
        # circ.line.weight = type(self)._line_weight
        # circ.line.ForeColor.RGB = type(self)._line_color
        # circ.fill.ForeColor.RGB = type(self)._fill_color
        circ.line.visible = -1 #msoTrue, important if default shape style does not have line
        circ.line.forecolor.ObjectThemeColor = 13 #msoThemeColorText1
        circ.fill.visible = -1 #msoTrue, important if default shape style does not have fill
        circ.fill.forecolor.ObjectThemeColor = 14 #msoThemeColorBackground1
        circ.LockAspectRatio = -1 #msoTrue

        pie = slide.shapes.addshape( powerpoint.MsoAutoShapeType['msoShapePie'], 100, 100, 30,30)
        # pie.line.weight = type(self)._line_weight
        # pie.line.ForeColor.RGB = type(self)._line_color
        # pie.fill.ForeColor.RGB = type(self)._line_color
        pie.line.visible = -1 #msoTrue, important if default shape style does not have line
        pie.line.forecolor.ObjectThemeColor = 13 #msoThemeColorText1
        pie.fill.visible = -1 #msoTrue, important if default shape style does not have fill
        pie.fill.forecolor.ObjectThemeColor = 13 #msoThemeColorText1
        pie.LockAspectRatio = -1 #msoTrue

        # gruppieren
        # grp = slide.Shapes.Range(System.Array[int]([shapeCount+1, shapeCount+2])).group()
        # grp = powerpoint.shape_indices_on_slide(slide, [shapeCount+1, shapeCount+2]).group()
        grp = powerpoint.last_n_shapes_on_slide(slide, 2).group()
        grp.LockAspectRatio = -1 #msoTrue

        # Fuellstand einstellen
        self.set_harvey(grp, fill, 1)

        # Tag erstellen
        self._add_tags(grp, int(1./fill))

        # Name
        grp.Name = "[BKT] Harvey Ball %s" % grp.id

        # selektieren und contextual tab aktivieren
        grp.select()
        context.ribbon.ActivateTab('bkt_context_tab_harvey')
    
    
    # ========================
    # = Fuellstand-% aendern =
    # ========================
    
    # big_inc_value = 5
    def harvey_percent_setter(self, shapes, value):
        if str(value) == "":
            value = 0
        # self.set_harveys(shapes, min(100, max(0, float(value))), 100)
        bkt.apply_delta_on_ALT_key(
            lambda shape, value: self.set_harveys([shape], float(value), 100), 
            lambda shape: self.get_harvey_percent([shape]), 
            shapes, value)

    def set_harveys(self, shapes, num, max_num):
        if num > max_num or num < 0:
            num = num % max_num
        for shape in shapes:
            self.set_harvey(shape, num, max_num)

    def set_harvey(self, shape, num, max_num):
        pie, _ = self.get_pie_circ(shape)
        if num == 0:
            pie.visible = 0
        else:
            pie.visible = -1
        if pie.HorizontalFlip:
            pie.adjustments.item[1] = -90 - (num*1./max_num*360.)
            pie.adjustments.item[2] = -90
        else:
            pie.adjustments.item[1] = -90
            pie.adjustments.item[2] = -90 + (num*1./max_num*360.)
        
        # Set tags if max_num is a denominator
        if max_num in self.harvey_denominators:
            self._add_tags(shape, max_num)
    
    def get_harvey_percent(self, shapes):
        shape = shapes[0]
        pie, _ = self.get_pie_circ(shape)
        if pie == None:
            return None
        #return pie.adjustments.item[2]
        if pie.HorizontalFlip:
            num = round((-(pie.adjustments.item[1] + 90) % 360) /360. * 100,1)
        else:
            num = round(( (pie.adjustments.item[2] + 90) % 360) /360. * 100,1)
        if num == 0:
            return 0 if pie.visible == 0 else 100
        else:
            return num
        

    def harvey_percent_inc(self, shapes):
        if bkt.get_key_state(bkt.KeyCodes.CTRL):
            step = 1
        elif bkt.get_key_state(bkt.KeyCodes.SHIFT):
            step = 25
        else:
            step = 5
        # step = 1 if bkt.get_key_state(bkt.KeyCodes.CTRL) else 5
        value = round(self.get_harvey_percent(shapes),3) + step
        self.harvey_percent_setter(shapes, value)

    def harvey_percent_dec(self, shapes):
        if bkt.get_key_state(bkt.KeyCodes.CTRL):
            step = 1
        elif bkt.get_key_state(bkt.KeyCodes.SHIFT):
            step = 25
        else:
            step = 5
        # step = 1 if bkt.get_key_state(bkt.KeyCodes.CTRL) else 5
        value = round(self.get_harvey_percent(shapes),3) - step
        self.harvey_percent_setter(shapes, value)
    

    # ====================
    # = Popup Funktionen =
    # ====================
    
    def harvey_percent_setter_popup(self, shapes, inc=True):
        for shape in shapes:
            old_value = self.get_harvey_percent([shape])
            if old_value == 0 and not inc:
                new_value = 100
            elif old_value == 100 and inc:
                new_value = 0
            else:
                step = 100./powerpoint.TagHelper.get_tag(shape, self.BKT_HARVEY_DENOM_TAG, 4, int)
                delta = step if inc else -step
                new_value = old_value+delta
                new_value = step * round(new_value/step) #round to multiple of step
            self.set_harveys([shape], new_value, 100),

    # =================
    # = Farbe aendern =
    # =================
    
    def color_gallery_action(self, shapes, color):
        for shape in shapes:
            self.set_harvey_color_rgb(shape, color)
        #type(self).set_harvey_colors_rgb(shapes, color)
    
    def set_harvey_color_rgb(self, shape, color):
        pie, circ = self.get_pie_circ(shape)
        if pie == None:
            return
        
        # For backwards compatibility: if line color hasn't changed, also update line
        pie_fill = pie.Fill.ForeColor
        circ_line = circ.Line.ForeColor
        if [pie_fill.ObjectThemeColor, pie_fill.Brightness, pie_fill.RGB] == [circ_line.ObjectThemeColor, circ_line.Brightness, circ_line.RGB]:
            circ_line.rgb = color
            pie.Line.ForeColor.rgb  = color

        pie_fill.rgb  = color
    
    def color_gallery_theme_color_change(self, shapes, color_index, brightness):
        for shape in shapes:
            self.set_harvey_color_theme(shape, color_index, brightness)
        #type(self).set_harvey_colors_theme(shapes, color_index, brightness)

    def set_harvey_color_theme(self, shape, color_index, brightness=0):
        pie, circ = self.get_pie_circ(shape)
        if pie == None:
            return
        
        # For backwards compatibility: if line color hasn't changed, also update line
        pie_fill = pie.Fill.ForeColor
        circ_line = circ.Line.ForeColor
        if [pie_fill.ObjectThemeColor, pie_fill.Brightness, pie_fill.RGB] == [circ_line.ObjectThemeColor, circ_line.Brightness, circ_line.RGB]:
            circ_line.ObjectThemeColor = color_index
            circ_line.Brightness = brightness
            pie.Line.ForeColor.Brightness  = brightness
            pie.Line.ForeColor.ObjectThemeColor  = color_index

        pie_fill.ObjectThemeColor  = color_index
        pie_fill.Brightness  = brightness
    
    def get_selected_color(self, shapes):
        # _,circ = self.get_pie_circ(shapes[0])
        # if circ != None:
        #     return [circ.Line.ForeColor.ObjectThemeColor, circ.Line.ForeColor.Brightness, circ.Line.ForeColor.RGB]
        # else:
        #     return None
        pie,_ = self.get_pie_circ(shapes[0])
        if pie != None:
            return [pie.Fill.ForeColor.ObjectThemeColor, pie.Fill.ForeColor.Brightness, pie.Fill.ForeColor.RGB]
        else:
            return None


    # =======================
    # = Hintergrund aendern =
    # =======================
    
    def background_gallery_action(self, shapes, color):
        for shape in shapes:
            self.set_harvey_background_rgb(shape, color)
        #type(self).set_harvey_colors_rgb(shapes, color)
    
    def set_harvey_background_rgb(self, shape, color):
        pie, circ = self.get_pie_circ(shape)
        if pie == None:
            return
        circ.Fill.Visible = -1
        circ.Fill.Transparency = 0
        circ.Fill.ForeColor.rgb  = color
    
    def background_gallery_theme_color_change(self, shapes, color_index, brightness):
        for shape in shapes:
            self.set_harvey_background_theme(shape, color_index, brightness)
        #type(self).set_harvey_colors_theme(shapes, color_index, brightness)

    def set_harvey_background_theme(self, shape, color_index, brightness=0):
        pie, circ = self.get_pie_circ(shape)
        if pie == None:
            return
        circ.Fill.Visible = -1
        circ.Fill.Transparency = 0
        circ.Fill.ForeColor.ObjectThemeColor  = color_index
        circ.Fill.ForeColor.Brightness  = brightness
    
    def get_selected_background(self, shapes):
        _,circ = self.get_pie_circ(shapes[0])
        if circ != None and circ.Fill.Visible and circ.Fill.Transparency == 0:
            return [circ.Fill.ForeColor.ObjectThemeColor, circ.Fill.ForeColor.Brightness, circ.Fill.ForeColor.RGB]
        else:
            return None
    
    def harvey_background_off(self, shapes):
        for shape in shapes:
            _, circ = self.get_pie_circ(shape)
            # circ.Fill.Visible = 0
            circ.Fill.Transparency = 1 #transparency=1 is preferred as background is still selectable then
    
    # def toggle_harvey_background(self, shapes, pressed):
    #     for shape in shapes:
    #         pie, circ = self.get_pie_circ(shape)
    #         circ.fill.visible = -1 if pressed else 0

    # def get_pressed_background(self, shapes):
    #     pie,circ = self.get_pie_circ(shapes[0])
    #     return circ.fill.visible == -1

    # =======================
    # = Linienfarbe aendern =
    # =======================
    
    def line_gallery_action(self, shapes, color, upgrade=False):
        for shape in shapes:
            self.set_harvey_line_rgb(shape, color)
    
    def set_harvey_line_rgb(self, shape, color):
        pie, circ = self.get_pie_circ(shape)
        if pie == None:
            return
        circ.Line.Visible = -1
        circ.Line.ForeColor.rgb = color
        pie.Line.Visible = -1
        pie.Line.ForeColor.rgb = color
    
    def line_gallery_theme_color_change(self, shapes, color_index, brightness, upgrade=False):
        for shape in shapes:
            self.set_harvey_line_theme(shape, color_index, brightness)

    def set_harvey_line_theme(self, shape, color_index, brightness=0):
        pie, circ = self.get_pie_circ(shape)
        if pie == None:
            return
        circ.Line.Visible = -1
        circ.Line.ForeColor.ObjectThemeColor = color_index
        circ.Line.ForeColor.Brightness = brightness

        #determine if style is modern, otherwise enable line
        if pie.top <= circ.top:
            pie.Line.Visible = -1
        pie.Line.ForeColor.ObjectThemeColor = color_index
        pie.Line.ForeColor.Brightness = brightness
    
    def get_selected_line(self, shapes):
        _,circ = self.get_pie_circ(shapes[0])
        if circ != None and circ.Line.Visible and circ.Line.Transparency == 0:
            return [circ.Line.ForeColor.ObjectThemeColor, circ.Line.ForeColor.Brightness, circ.Line.ForeColor.RGB]
        else:
            return None
    
    def harvey_line_off(self, shapes):
        for shape in shapes:
            pie, circ = self.get_pie_circ(shape)
            pie.Line.Visible = 0
            circ.Line.Visible = 0
    
    def harvey_line_outside_only(self, shapes):
        #determine new shape type for all
        pie, _ = self.get_pie_circ(shapes[0])
        if pie.AutoShapeType != powerpoint.MsoAutoShapeType['msoShapePie']:
            new_type = powerpoint.MsoAutoShapeType['msoShapePie']
        else:
            new_type = powerpoint.MsoAutoShapeType['msoShapeArc']
        
        #adjust all shape types without changing the value
        for shape in shapes:
            pie, _ = self.get_pie_circ(shape)
            cur1 = pie.adjustments.item[1]
            cur2 = pie.adjustments.item[2]
            pie.AutoShapeType = new_type
            pie.adjustments.item[1] = cur1
            pie.adjustments.item[2] = cur2


    # ================
    # = Stil aendern =
    # ================

    def harvey_change_style_classic(self, shapes):
        self.harvey_change_style(shapes, "classic")
    def harvey_change_style_modern(self, shapes):
        self.harvey_change_style(shapes, "modern")
    def harvey_change_style_chart(self, shapes):
        self.harvey_change_style(shapes, "chart")

    def harvey_change_style(self, shapes, style="classic"):
        for shape in shapes:
            pie, circ = self.get_pie_circ(shape)
            if pie == None:
                continue

            # Store adjustment values and make arc a full circle; otherwise, incorrect width/height values are given for arc-type (no problem with pie)
            cur1 = pie.adjustments.item[1]
            cur2 = pie.adjustments.item[2]
            pie.adjustments.item[1] = -90
            pie.adjustments.item[2] = -90

            # always start from classic
            circ.line.DashStyle = 1 #straight
            pie.left, pie.top = circ.left, circ.top
            pie.width, pie.height = circ.width, circ.height
            pie.line.visible = -1
            pie.LockAspectRatio = -1
            
            if style == "modern":
                pie.line.visible = 0
                pie.scaleHeight(0.8, 0, 1)
            
            elif style == "chart":
                circ.line.DashStyle = 10 #dashed
                pie.scaleHeight(1.1, 0, 1)

            # Restore adjustment values
            pie.adjustments.item[1] = cur1
            pie.adjustments.item[2] = cur2

    def harvey_fliph_pressed(self, shapes):
        return shapes[0].HorizontalFlip == -1

    def harvey_fliph(self, shapes, pressed):
        pressed = -1 if pressed else 0
        for shape in shapes:
            if shape.HorizontalFlip != pressed:
                shape.Flip(0) #msoFlipHorizontal


    # 2022-12-16:
    # Developed this code to upgrade harvey version but decided to take a different approach.
    # Maybe this code can be reused later, so I did not delete it.

    # def is_legacy(self, shape):
    #     # return not powerpoint.TagHelper.has_tag(shape, self.BKT_HARVEY_DIALOG_TAG, self.BKT_HARVEY_VERSION)
    #     version = powerpoint.TagHelper.get_tag(shape, self.BKT_HARVEY_DIALOG_TAG)
    #     return not version or version in self.BKT_HARVEY_LEGACY_VERSION

    # def is_legacy_any(self, shapes):
    #     return any(self.is_legacy(shape) for shape in shapes)

    # def upgrade(self, shape):
    #     pie, _ = self.get_pie_circ(shape)
    #     cur1 = pie.adjustments.item[1]
    #     cur2 = pie.adjustments.item[2]
    #     pie.AutoShapeType = powerpoint.MsoAutoShapeType['msoShapeArc']
    #     pie.adjustments.item[1] = cur1
    #     pie.adjustments.item[2] = cur2
    #     self._add_tags(shape)

    # def upgrade_all(self, shapes, pressed=True):
    #     if bkt.message.confirmation("Nach der Aktualisierung ist die Optik der Harvey Balls minimal verändert (angeglichen an ThinkCell) und die Linienfarbe kann geändert werden, jedoch ist eine Anpassung in älteren BKT-Versionen ist nicht mehr möglich. Vorgang fortsetzen?", "BKT: Harvey Balls"):
    #         for shape in shapes:
    #             self.upgrade(shape)
    
    
    # ========================
    # = Groesse genau setzen =
    # ========================
    
    harvey_denominators = [3,4,5,6,8]
    # harvey_columns = max(harvey_denominators) +1
    harvey_buttons = [
        [n, n_max]
        for n_max in harvey_denominators
        for n in range(0,9) #9=harvey_columns, stopped working in ipy3 and I don't know why
    ]
    harvey_labels = [
        '%s/%s' % (n,n_max) if n<=n_max else ' '
        for [n, n_max] in harvey_buttons
    ]
    
    
    def change_harvey(self, selected_item, index, shapes):
        self.set_harveys(shapes, self.harvey_buttons[index][0], self.harvey_buttons[index][1]),
    
    def get_item_count(self):
        return len(self.harvey_buttons)

    def get_item_label(self, index):
        return self.harvey_labels[index]

    def get_item_screentip(self, index):
        label = self.harvey_labels[index]
        return "Füllstand eines Harvey-Balls ändern auf %s" % label if label else ""
    
    def get_harvey_item_image(self, index):
        return self.get_harvey_image(self.harvey_buttons[index][0]*1. / self.harvey_buttons[index][1] )

    def get_harvey_image(self, percent, size=32, color=Drawing.Brushes.Gray):
        if percent < 0 or percent > 1:
            return glib.empty_image(size, size)

        img = Drawing.Bitmap(size, size)
        # color = Drawing.ColorTranslator.FromOle(0)
        g = Drawing.Graphics.FromImage(img)
        
        #Draw smooth rectangle/ellipse
        g.SmoothingMode = Drawing.Drawing2D.SmoothingMode.AntiAlias

        g.DrawEllipse(Drawing.Pen(color,2), 2,2, size-5, size-5)
        # g.FillPie(Drawing.SolidBrush(color), Drawing.Rectangle(1,1,size-3,size-3), -90, percent*360. )
        g.FillPie(color, Drawing.Rectangle(1,1,size-3,size-3), -90, percent*360. )
        return img
    
    
    # =========================
    # = Grosse Buttons im Tab =
    # =========================

    def get_harvey_image_by_control(self, current_control):
        percent = float(current_control["tag"])/100
        return self.get_harvey_image(percent, 64)

    def set_harvey_by_control(self, shapes, current_control):
        args = {
            "0": (0,1),
            "25": (1,4),
            "33.3": (1,3),
            "50": (2,4),
            "66.6": (2,3),
            "75": (3,4),
            "100": (1,1),
        }
        self.set_harveys(shapes, *args.get(current_control["tag"]))


harvey_balls = HarveyBalls()