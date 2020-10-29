# -*- coding: utf-8 -*-
'''
Created on 06.07.2016

@author: rdebeerst
'''

from __future__ import absolute_import

import bkt
import bkt.library.powerpoint as powerpoint
import bkt.library.graphics as glib

from bkt import dotnet
Drawing = dotnet.import_drawing()


class HarveyBalls(object):
    BKT_HARVEY_DIALOG_TAG = "BKT_SHAPE_HARVEY"
    BKT_HARVEY_DENOM_TAG = "BKT_HARVEY_DENOM_TAG"
    BKT_HARVEY_VERSION = "BKT_HARVEY_V1"

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
        pie.Fill.ForeColor.rgb  = color
        pie.Line.ForeColor.rgb  = color
        circ.Line.ForeColor.rgb = color
    
    def color_gallery_theme_color_change(self, shapes, color_index, brightness):
        for shape in shapes:
            self.set_harvey_color_theme(shape, color_index, brightness)
        #type(self).set_harvey_colors_theme(shapes, color_index, brightness)

    def set_harvey_color_theme(self, shape, color_index, brightness=0):
        pie, circ = self.get_pie_circ(shape)
        if pie == None:
            return
        pie.Fill.ForeColor.ObjectThemeColor  = color_index
        pie.Line.ForeColor.ObjectThemeColor  = color_index
        circ.Line.ForeColor.ObjectThemeColor = color_index
        pie.Fill.ForeColor.Brightness  = brightness
        pie.Line.ForeColor.Brightness  = brightness
        circ.Line.ForeColor.Brightness = brightness
    
    def get_selected_color(self, shapes):
        _,circ = self.get_pie_circ(shapes[0])
        if circ != None:
            return [circ.Line.ForeColor.ObjectThemeColor, circ.Line.ForeColor.Brightness, circ.Line.ForeColor.RGB]
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

    # ================
    # = Stil aendern =
    # ================

    def harvey_change_style_classic(self, shapes):
        self.harvey_change_style(shapes, "classic")
    def harvey_change_style_modern(self, shapes):
        self.harvey_change_style(shapes, "modern")

    def harvey_change_style(self, shapes, style="classic"):
        for shape in shapes:
            pie, circ = self.get_pie_circ(shape)
            if pie == None:
                continue

            # always start from classic
            pie.left, pie.top = circ.left, circ.top
            pie.width, pie.height = circ.width, circ.height
            pie.line.visible = -1
            pie.LockAspectRatio = -1
            
            if style == "modern":
                pie.line.visible = 0
                pie.scaleHeight(0.8, 0, 1)


    # =====================================
    # = Feature-Logik und Hilfsfunktionen =
    # =====================================
    
    def is_harvey_group(self, shape):
        # "new" method via tags
        if powerpoint.TagHelper.has_tag(shape, self.BKT_HARVEY_DIALOG_TAG):
            return True
        # "old" method via shape types
        pie, _ = self.get_pie_circ(shape)
        return pie != None

    def get_pie_circ(self, shape):
        if not shape.Type == powerpoint.MsoShapeType['msoGroup']:
            return None, None
        if not shape.GroupItems.Count == 2:
            return None, None

        if shape.GroupItems(1).AutoShapeType == powerpoint.MsoAutoShapeType['msoShapePie']:
            return shape.GroupItems(1), shape.GroupItems(2)
        elif shape.GroupItems(2).AutoShapeType == powerpoint.MsoAutoShapeType['msoShapePie']:
            return shape.GroupItems(2), shape.GroupItems(1)
        else:
            return None, None

    
    
    # ========================
    # = Groesse genau setzen =
    # ========================
    
    harvey_denominators = [3,4,5,6,8]
    harvey_columns = max(harvey_denominators) +1
    harvey_buttons = [
        [n, n_max]
        for n_max in harvey_denominators
        for n in range(0,harvey_columns)
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

    def change_harvey_enabled(self, shapes):
        return self.is_harvey_group(shapes[0])
    
    def get_harvey_item_image(self, index):
        return self.get_harvey_image(self.harvey_buttons[index][0]*1. / self.harvey_buttons[index][1] )

    def get_harvey_image(self, percent, size=32):
        if percent < 0 or percent > 1:
            return glib.empty_image(size, size)

        img = Drawing.Bitmap(size, size)
        color = Drawing.ColorTranslator.FromOle(0)
        g = Drawing.Graphics.FromImage(img)
        
        #Draw smooth rectangle/ellipse
        g.SmoothingMode = Drawing.Drawing2D.SmoothingMode.AntiAlias

        g.DrawEllipse(Drawing.Pen(color,2), 2,2, size-5, size-5)
        g.FillPie(Drawing.SolidBrush(color), Drawing.Rectangle(1,1,size-3,size-3), -90, percent*360. )
        return img
    



harvey_balls = HarveyBalls()

def harvey_color_gallery(**kwargs):
    return bkt.ribbon.ColorGallery(
        label = 'Farbe ändern',
        #image_mso = 'RecolorColorPicker',
        image='harvey ball color',
        screentip="Farbe eines Harvey-Balls ändern",
        supertip="Passe die Farbe eines Harvey-Balls entsprechend der Auswahl an.\n\nEin Harvey-Ball-Shape ist eine Gruppe aus Kreis- und Torten-Shape.",
        on_rgb_color_change   = bkt.Callback(harvey_balls.color_gallery_action, shapes=True),
        on_theme_color_change = bkt.Callback(harvey_balls.color_gallery_theme_color_change, shapes=True),
        get_selected_color    = bkt.Callback(harvey_balls.get_selected_color, shapes=True),
        get_enabled           = bkt.Callback(harvey_balls.change_harvey_enabled, shapes=True),
        item_width=16, item_height=16,
        **kwargs
    )

def harvey_background_gallery(**kwargs):
    return bkt.ribbon.ColorGallery(
        label = 'Hintergrund ändern',
        #image_mso = 'RecolorColorPicker',
        image='harvey ball background',
        screentip="Hintergrund eines Harvey-Balls ändern",
        supertip="Passe die Hintergrund-Farbe eines Harvey-Balls entsprechend der Auswahl an.\n\nEin Harvey-Ball-Shape ist eine Gruppe aus Kreis- und Torten-Shape.",
        on_rgb_color_change   = bkt.Callback(harvey_balls.background_gallery_action, shapes=True),
        on_theme_color_change = bkt.Callback(harvey_balls.background_gallery_theme_color_change, shapes=True),
        get_selected_color    = bkt.Callback(harvey_balls.get_selected_background, shapes=True),
        get_enabled           = bkt.Callback(harvey_balls.change_harvey_enabled, shapes=True),
        children=[
            bkt.ribbon.Button(
                label='Hintergrund aus',
                supertip="Harvey-Ball Hintergrund auf transparent setzen",
                #get_image=bkt.Callback(lambda: harvey_balls.get_harvey_image(0.6, 64)),
                image='harvey ball background',
                on_action=bkt.Callback(harvey_balls.harvey_background_off),
            ),
        ],
        item_width=16, item_height=16,
        **kwargs
    )

def harvey_size_gallery(**kwargs):
    return bkt.ribbon.Gallery(
        label = 'Füllstand ändern',
        image = 'harvey ball size',
        #get_image=bkt.Callback(lambda: harvey_balls.get_harvey_image(0.6, 64)),
        screentip="Füllstand eines Harvey-Balls ändern",
        supertip="Passe den Füllstand eines Harvey-Balls entsprechend der Auswahl an.\n\nEin Harvey-Ball-Shape ist eine Gruppe aus Kreis- und Torten-Shape.",
        columns=str(harvey_balls.harvey_columns),
        on_action_indexed = bkt.Callback(harvey_balls.change_harvey, shapes=True),
        get_item_count    = bkt.Callback(harvey_balls.get_item_count),
        get_item_label    = bkt.Callback(harvey_balls.get_item_label),
        get_item_screentip = bkt.Callback(harvey_balls.get_item_screentip),
        get_item_supertip = bkt.Callback(lambda index: "Passe den Füllstand eines Harvey-Balls entsprechend der Auswahl an."),
        get_enabled       = bkt.Callback(harvey_balls.change_harvey_enabled, shapes=True),
        get_item_image    = bkt.Callback(harvey_balls.get_harvey_item_image),
        item_width=16, item_height=16,
        **kwargs
    )


harvey_create_button = bkt.ribbon.Button(
    id='create_harvey_ball',
    label='Harvey Ball',
    screentip='Harvey Ball erstellen',
    image='harvey ball',
    on_action=bkt.Callback(harvey_balls.create_harvey_ball),
    supertip="Füge ein Harvey-Ball ein, welcher sich bzgl. Farbe/Füllstand konfigurieren lässt.\n\nFarbe und Füllstand lassen sich über Kontext-Menü und Kontext-Tab konfigurieren, im Tab auch Prozent-Angaben möglich.\n\nEin Harvey-Ball-Shape ist eine Gruppe aus Kreis- und Torten-Shape."
)


harvey_ball_group = bkt.ribbon.Group(
    id="bkt_harvey_group",
    label = "Harvey Balls",
    children = [
        bkt.ribbon.Button(
            id='harvey_ball_create',
            size='large',
            label='Neuer Harvey Ball',
            screentip='Harvey Ball erstellen',
            image='harvey ball',
            on_action=bkt.Callback(harvey_balls.create_harvey_ball),
            supertip="Füge ein Harvey-Ball ein, welcher sich bzgl. Farbe/Füllstand konfigurieren lässt.\n\nFarbe und Füllstand lassen sich über Kontext-Menü und Kontext-Tab konfigurieren, im Tab auch Prozent-Angaben möglich.\n\nEin Harvey-Ball-Shape ist eine Gruppe aus Kreis- und Torten-Shape."
        ),
        bkt.ribbon.Button(
            id='harvey_ball_duplicate',
            size='large',
            label='Harvey Ball duplizieren',
            screentip='Harvey Ball duplizieren',
            image='harvey ball duplicate',
            on_action=bkt.Callback(lambda selection: selection.ShapeRange.Duplicate().Select()),
            supertip="Dupliziert den aktuell ausgewählten Harvey-Ball."
        ),
        bkt.ribbon.Separator(),

        #bkt.ribbon.SplitButton(show_label=False, children=[
            # bkt.ribbon.Button(
            #     id='create_harvey_ball',
            #     label='Harvey Ball erstellen',
            #     screentip='Harvey Ball erstellen',
            #     image='harvey ball',
            #     on_action=bkt.Callback(harvey_balls.create_harvey_ball)
            # ),
            # bkt.ribbon.Menu(label='menu',
            #     children = [
        harvey_size_gallery(id='harvey_ball_size_gallery', size="large"),
        harvey_color_gallery(id='harvey_ball_color_gallery', size="large"),
        #         ]
        #     )
        # ]),

        harvey_background_gallery(id='harvey_ball_background', size="large"),

        bkt.ribbon.Separator(),

        bkt.ribbon.Button(
            id='harvey_ball_style_classic',
            size='large',
            label='Style klassisch',
            supertip="Harvey-Ball im klassischen Style ohne zusätzlichem Rand darstellen.",
            image='harvey ball classic',
            on_action=bkt.Callback(harvey_balls.harvey_change_style_classic, shapes=True),
        ),

        bkt.ribbon.Button(
            id='harvey_ball_style_modern',
            size='large',
            label='Style modern',
            supertip="Harvey-Ball im modernen Style mit weißem Rand darstellen.",
            image='harvey ball modern',
            on_action=bkt.Callback(harvey_balls.harvey_change_style_modern, shapes=True),
        ),

        bkt.ribbon.Separator(),
        #bkt.ribbon.LabelControl(label="Füllstand:"),
        
        bkt.ribbon.SpinnerBox(label='Füllstand in %', size_string='33,33',
            id = 'harvey_spinner',
            screentip="Füllstand eines Harvey-Balls ändern",
            supertip="Passe den Füllstand eines Harvey-Balls entsprechend der Auswahl an.\n\nEin Harvey-Ball-Shape ist eine Gruppe aus Kreis- und Torten-Shape.",
            on_change = bkt.Callback(harvey_balls.harvey_percent_setter, shapes=True),
            get_text  = bkt.Callback(harvey_balls.get_harvey_percent, shapes=True),
            increment = bkt.Callback(harvey_balls.harvey_percent_inc, shapes=True),
            decrement = bkt.Callback(harvey_balls.harvey_percent_dec, shapes=True)
        ),
        bkt.ribbon.LabelControl(label="   mit SHIFT: Schrittweite +/-25"),
        bkt.ribbon.LabelControl(label="   mit ALT: Delta je Harvey Ball"),

        bkt.ribbon.Separator(),

        bkt.ribbon.Button(
            id='harvey_ball_0',
            size='large',
            label='0%',
            screentip="Harvey-Ball auf 0%",
            supertip="Setzt alle gewählten Harvey-Balls auf 0%",
            get_image=bkt.Callback(lambda: harvey_balls.get_harvey_image(0, 64)),
            on_action=bkt.Callback(lambda shapes: harvey_balls.set_harveys(shapes, 0, 1)),
        ),
        bkt.ribbon.Button(
            id='harvey_ball_25',
            size='large',
            label='25%',
            screentip="Harvey-Ball auf 25%",
            supertip="Setzt alle gewählten Harvey-Balls auf 25%",
            get_image=bkt.Callback(lambda: harvey_balls.get_harvey_image(0.25, 64)),
            on_action=bkt.Callback(lambda shapes: harvey_balls.set_harveys(shapes, 1, 4)),
        ),
        bkt.ribbon.Button(
            id='harvey_ball_33',
            size='large',
            label='33%',
            screentip="Harvey-Ball auf 33%",
            supertip="Setzt alle gewählten Harvey-Balls auf 33%",
            get_image=bkt.Callback(lambda: harvey_balls.get_harvey_image(0.333, 64)),
            on_action=bkt.Callback(lambda shapes: harvey_balls.set_harveys(shapes, 1, 3)),
        ),
        bkt.ribbon.Button(
            id='harvey_ball_50',
            size='large',
            label='50%',
            screentip="Harvey-Ball auf 50%",
            supertip="Setzt alle gewählten Harvey-Balls auf 50%",
            get_image=bkt.Callback(lambda: harvey_balls.get_harvey_image(0.5, 64)),
            on_action=bkt.Callback(lambda shapes: harvey_balls.set_harveys(shapes, 2, 4)),
        ),
        bkt.ribbon.Button(
            id='harvey_ball_66',
            size='large',
            label='66%',
            screentip="Harvey-Ball auf 66%",
            supertip="Setzt alle gewählten Harvey-Balls auf 66%",
            get_image=bkt.Callback(lambda: harvey_balls.get_harvey_image(0.667, 64)),
            on_action=bkt.Callback(lambda shapes: harvey_balls.set_harveys(shapes, 2, 3)),
        ),
        bkt.ribbon.Button(
            id='harvey_ball_75',
            size='large',
            label='75%',
            screentip="Harvey-Ball auf 75%",
            supertip="Setzt alle gewählten Harvey-Balls auf 75%",
            get_image=bkt.Callback(lambda: harvey_balls.get_harvey_image(0.75, 64)),
            on_action=bkt.Callback(lambda shapes: harvey_balls.set_harveys(shapes, 3, 4)),
        ),
        bkt.ribbon.Button(
            id='harvey_ball_100',
            size='large',
            label='100%',
            screentip="Harvey-Ball auf 100%",
            supertip="Setzt alle gewählten Harvey-Balls auf 100%",
            get_image=bkt.Callback(lambda: harvey_balls.get_harvey_image(1, 64)),
            on_action=bkt.Callback(lambda shapes: harvey_balls.set_harveys(shapes, 1, 1)),
        ),
    ]
)

harvey_ball_tab = bkt.ribbon.Tab(
    id = "bkt_context_tab_harvey",
    label = "[BKT] Harvey Balls",
    get_visible=bkt.Callback(harvey_balls.change_harvey_enabled, shapes=True),
    children = [
        # Harvey Balls
        harvey_ball_group
    ]
)