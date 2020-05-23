# -*- coding: utf-8 -*-
'''
Created on 2018-01-10
@author: Florian Stallmann
'''

from __future__ import absolute_import

import os.path
import io
import json

import logging

import bkt
import bkt.library.powerpoint as pplib

D = bkt.dotnet.import_drawing()


COLOR_THEME = 1
COLOR_RGB = 2

BUTTON_THEME  = 4
BUTTON_RECENT = 8
BUTTON_USERDEFINED = 16


class PPTColor(object):
    '''
    This class represents a single color similar to the powerpoint color object.
    Helper methods provided to pickup or apply color from powerpoint color object, 
    to export color as tuple and to export color as html-code for WPF.
    '''

    @classmethod
    def from_color_obj(cls, color_obj):
        return cls().pickup_from_color_obj(color_obj)

    @classmethod
    def new_rgb(cls, color_rgb):
        return cls(COLOR_RGB, color_rgb=color_rgb)
    
    @classmethod
    def new_theme(cls, color_index, brightness, color_rgb=None, shade_index=None):
        return cls(COLOR_THEME, color_index, brightness, color_rgb, shade_index)


    def __init__(self, color_type=COLOR_RGB, color_index=None, brightness=None, color_rgb=None, shade_index=None):
        self.color_type  = color_type
        self.color_index = color_index
        self.brightness  = brightness
        self.color_rgb   = color_rgb
        self.shade_index = shade_index #-1=not yet defined, set index on next update_from_context; None=ignore, only use brightness
    
    def __eq__(self, other):
        if isinstance(other, PPTColor):
            return self.get_color_tuple() == other.get_color_tuple()
        return False
    
    def __ne__(self, other):
        return not self.__eq__(other)


    def update_rgb(self, rgb):
        self.color_rgb = rgb

    def update_rgb_from_context(self, context):
        if self.color_type == COLOR_RGB:
            raise TypeError("Cannot update RGB value from context")

        if self.shade_index == -1:
            color = pplib.ColorHelper.get_theme_color(context, self.color_index, brightness=self.brightness)
            self.shade_index = color.shade_index
            logging.debug("QuickEdit: shade index changed to %s", self.shade_index)
        elif self.shade_index is not None:
            color = pplib.ColorHelper.get_theme_color(context, self.color_index, shade_index=self.shade_index)
            self.brightness = color.brightness
            logging.debug("QuickEdit: brightness changed to %s", self.brightness)
        else:
            color = pplib.ColorHelper.get_theme_color(context, self.color_index, brightness=self.brightness)
        self.color_rgb = color.rgb


    def pickup_from_color_obj(self, color_obj):
        if color_obj.Type == pplib.MsoColorType['msoColorTypeScheme']:
            self.color_type  = COLOR_THEME
            self.color_index = color_obj.ObjectThemeColor
            # self.color_index = color_obj.SchemeColor
            self.brightness  = color_obj.Brightness
            self.color_rgb   = color_obj.RGB
            self.shade_index = -1 #-1: shade index will be set on first update
        else:
            self.color_type  = COLOR_RGB
            self.color_index = None
            self.brightness  = None
            self.color_rgb   = color_obj.RGB
            self.shade_index = None
        return self
    
    def apply_to_color_obj(self, color_obj):
        if self.color_type == COLOR_THEME:
            color_obj.ObjectThemeColor = self.color_index
            # color_obj.SchemeColor = self.color_index
            color_obj.Brightness = self.brightness
        else:
            color_obj.RGB = self.color_rgb


    def get_color_tuple(self):
        if self.color_type == COLOR_THEME:
            return (COLOR_THEME, self.color_index, int(100*self.brightness)) #convert brightness to int to avoid floating point comparison problems
        else:
            return (COLOR_RGB, self.color_rgb)
    
    def get_color_html(self):
        if self.color_rgb is None:
            raise ValueError("RGB for color not defined")
        color = D.ColorTranslator.FromOle(self.color_rgb)
        return D.ColorTranslator.ToHtml(color)


class QEColorButtons(object):
    '''
    This class manages all color buttons, e.g. update color buttons based on current context
    and set the active colors to define which buttons are checked.
    '''

    all_buttons = {} #store all color buttons in a dict with index as key
    active_colors = set() #set of colors that are currently selected, so need to be checked

    @classmethod
    def update_colors(cls, context):
        for color in cls.all_buttons.values():
            color.set_rgb_from_context(context)

    @classmethod
    def set_active_colors(cls, active_colors=None):
        cls.active_colors = active_colors or set()
        for color in cls.all_buttons.values():
            color.OnPropertyChanged("Checked")
    
    @classmethod
    def get(cls, identifier):
        return cls.all_buttons[identifier]
    
    @classmethod
    def add(cls, qebutton):
        cls.all_buttons[qebutton.identifier] = qebutton
    
    @classmethod
    def next_identifier(cls):
        return len(cls.all_buttons)


class QEColorButton(bkt.ui.NotifyPropertyChangedBase):
    '''
    This class represents a single color button in the panel window. It support notify properties
    for WPF. Each button can be assigned to a color object, otherwise fallback color is used.
    '''
    # fallback_color = 7829367 #fallback color/ not defined
    fallback_color = PPTColor.new_rgb(7829367) #fallback color/ not defined

    def __init__(self, label, button_type, index=None):
        self.identifier = QEColorButtons.next_identifier()
        # self.identifier = str(button_type*10 + len(QEColorButton.all_buttons))
        
        self._index = index
        if button_type == BUTTON_THEME:
            self._color = PPTColor.new_theme(index, 0)
        else:
            self._color = None

        self.label = label
        self.button_type = button_type
        self.image = None
        QEColorButtons.add(self)

        super(QEColorButton, self).__init__()


    ### WPF PROPERTIES ###

    @property
    def Tag(self):
        return self.identifier

    @property
    def Label(self):
        if not self.is_defined:
            return "%s (%s)" % (self.label, "Undefined")
        elif self.is_userdefined:
            if self.is_theme_color:
                return "%s (%s)" % (self.label, "Theme")
            else:
                return "%s (%s)" % (self.label, "RGB")
        else:
            return self.label
    
    @property
    def Color(self):
        try:
            return self._color.get_color_html()
        except:
            return self.fallback_color.get_color_html()
    
    @property
    def Checked(self):
        return self.get_checked()
    @Checked.setter
    def Checked(self, value):
        # enforce onPropertyChange to ensure correct checked state
        self.OnPropertyChanged("Checked")

    ### END WPF PROPERTIES ###


    def set_rgb_from_context(self, context):
        if self.is_theme_color:
            self._color.update_rgb_from_context(context)
            self.OnPropertyChanged("Color")
            self.OnPropertyChanged("Label")
            # self.image = None
        elif self.is_recent:
            count_recent = context.app.ActivePresentation.ExtraColors.Count
            if count_recent >= self._index:
                self.set_color( PPTColor.new_rgb(context.app.ActivePresentation.ExtraColors(min(count_recent,10)-self._index+1)) )
            else:
                self.set_color(None)

    def get_color(self):
        return self._color

    def set_color(self, color):
        self._color = color
        self.OnPropertyChanged("Color")
        self.OnPropertyChanged("Label")
        # self.image = None

    def set_userdefined_rgb(self, color_rgb):
        # self.color_type = COLOR_RGB | COLOR_USERDEFINED
        self.set_color( PPTColor.new_rgb(color_rgb) )

    def set_userdefined_theme(self, color_index, brightness=0, color_rgb=None, shade_index=None):
        # self.color_type = COLOR_THEME | COLOR_USERDEFINED
        self.set_color( PPTColor.new_theme(color_index, brightness, color_rgb, shade_index) )


    def get_checked(self):
        if not self.is_defined:
            return False
        if self._color.get_color_tuple() in QEColorButtons.active_colors:
            return True
        # if self.is_theme_color:
        #     if (COLOR_THEME, self.color_index, self.brightness) in QEColor.active_colors:
        #         return True
        # else:
        #     if (COLOR_RGB, self.color_rgb) in QEColor.active_colors:
        #         return True
        return False


    @property
    def is_theme_color(self):
        return self._color is not None and self._color.color_type == COLOR_THEME
        # return self.color_type & COLOR_THEME == COLOR_THEME

    @property
    def is_rgb_color(self):
        return self._color is not None and self._color.color_type == COLOR_RGB
        # return self.color_type & COLOR_RGB == COLOR_RGB


    @property
    def is_theme(self):
        return self.button_type == BUTTON_THEME

    @property
    def is_recent(self):
        return self.button_type == BUTTON_RECENT

    @property   
    def is_userdefined(self):
        return self.button_type == BUTTON_USERDEFINED
        # return self.color_type & COLOR_USERDEFINED == COLOR_USERDEFINED


    @property
    def is_defined(self):
        return self._color is not None and self._color.color_rgb is not None


    # def get_image(self, size=16):
    #     if self.image is not None:
    #         return self.image
    #     else:
    #         if self.is_defined:
    #             color = D.ColorTranslator.FromOle(self._color.color_rgb)
    #         else:
    #             color = D.ColorTranslator.FromOle(self.fallback_color.color_rgb)
    #         self.image = D.Bitmap(size, size)
    #         #method 1: pixel by pixel
    #         # for x in range(0, self.image.Height):
    #         #     for y in range(0, self.image.Width):
    #         #         self.image.SetPixel(x, y, color)
    #         #method 2: using graphics
    #         g = D.Graphics.FromImage(self.image)
    #         g.Clear(color)
    #         return self.image
    

    def to_json(self):
        if self.is_rgb_color:
            return (self._color.color_rgb, None, None, None)
        else:
            return (self._color.color_rgb, self._color.color_index, float(self._color.brightness), self._color.shade_index) #explicitly convert to float for json!

    def from_json(self, value):
        if value[1] is None:
            self.set_userdefined_rgb(value[0])
        elif len(value) == 3:
            self.set_userdefined_theme(value[1], value[2], value[0])
        else:
            self.set_userdefined_theme(value[1], value[2], value[0], value[3])


class QECatalog(bkt.ui.NotifyPropertyChangedBase):
    ''' Catalog class for config files, name and whether they are checked '''

    def __init__(self, file, name, checked=False):
        self.File = file
        self.Name = name
        self._checked = checked
        super(QECatalog, self).__init__()

    @property
    def Checked(self):
        return self._checked
    @Checked.setter
    def Checked(self, value):
        self._checked = value
        # enforce onPropertyChange to ensure correct checked state
        self.OnPropertyChanged("Checked")
    
    # namedtuple("QECatalog", ["file", "name", "checked"])


class QuickEdit(object):

    #### Theme colors ####
    _colors = [
        # QEColorButton('Background 1', BUTTON_THEME, 14),
        # QEColorButton('Text 1',       BUTTON_THEME, 13),
        # QEColorButton('Background 2', BUTTON_THEME, 16),
        # QEColorButton('Text 2',       BUTTON_THEME, 15),
        # QEColorButton('Accent 1',     BUTTON_THEME, 5),
        # QEColorButton('Accent 2',     BUTTON_THEME, 6),
        # QEColorButton('Accent 3',     BUTTON_THEME, 7),
        # QEColorButton('Accent 4',     BUTTON_THEME, 8),
        # QEColorButton('Accent 5',     BUTTON_THEME, 9),
        # QEColorButton('Accent 6',     BUTTON_THEME, 10),
        QEColorButton(name,     BUTTON_THEME, index)
        for index,name in pplib.ColorHelper.get_theme_colors()
    ]

    #### Extra/Recent Colors ####
    _recent = [
        QEColorButton('Recent 1',     BUTTON_RECENT, 1),
        QEColorButton('Recent 2',     BUTTON_RECENT, 2),
        QEColorButton('Recent 3',     BUTTON_RECENT, 3),
        QEColorButton('Recent 4',     BUTTON_RECENT, 4),
        QEColorButton('Recent 5',     BUTTON_RECENT, 5),
        QEColorButton('Recent 6',     BUTTON_RECENT, 6),
        QEColorButton('Recent 7',     BUTTON_RECENT, 7),
        QEColorButton('Recent 8',     BUTTON_RECENT, 8),
        QEColorButton('Recent 9',     BUTTON_RECENT, 9),
        QEColorButton('Recent 10',    BUTTON_RECENT, 10),
    ]

    #### Userdefined Colors ####
    _userdefined = [
        QEColorButton('Own 1',     BUTTON_USERDEFINED, 1),
        QEColorButton('Own 2',     BUTTON_USERDEFINED, 2),
        QEColorButton('Own 3',     BUTTON_USERDEFINED, 3),
        QEColorButton('Own 4',     BUTTON_USERDEFINED, 4),
        QEColorButton('Own 5',     BUTTON_USERDEFINED, 5),
        QEColorButton('Own 6',     BUTTON_USERDEFINED, 6),
        QEColorButton('Own 7',     BUTTON_USERDEFINED, 7),
        QEColorButton('Own 8',     BUTTON_USERDEFINED, 8),
        QEColorButton('Own 9',     BUTTON_USERDEFINED, 9),
        QEColorButton('Own 10',    BUTTON_USERDEFINED, 10),
    ]

    _catalogs = [
        QECatalog("default.json", "Katalog 1", True),
        QECatalog("default2.json", "Katalog 2", False),
        QECatalog("default3.json", "Katalog 3", False)
    ]

    config_folder = bkt.helpers.get_fav_folder("quickedit")
    current_file = "default.json"

    @classmethod
    def initialize(cls):
        cls.current_file = bkt.settings.get("quickedit.default_file", "default.json")
        cls.read_from_config(cls.current_file)


    @classmethod
    def save_to_config(cls):
        import os
        # bkt.console.show_message("%r" % cls._usercolors)
        # bkt.console.show_message(json.dumps(cls._usercolors))
        file = os.path.join(cls.config_folder, cls.current_file)
        if not os.path.exists(cls.config_folder):
            os.makedirs(cls.config_folder)
        elif os.path.isfile(file):
            #make backup
            backup_file = file[:-5] + "-backup.json"
            try:
                os.remove(backup_file)
            except OSError:
                pass
            os.rename(file, backup_file)

        values = [v.to_json() for v in cls._userdefined if v.is_defined]
        with io.open(file, 'w', encoding='utf-8') as json_file:
            json.dump(values, json_file, ensure_ascii=False)

    @classmethod
    def read_from_config(cls, filename="default.json"):
        #store settings
        cls.current_file = filename
        bkt.settings["quickedit.default_file"] = filename
        file = os.path.join(cls.config_folder, filename)

        #update catalog list
        for cat in cls._catalogs:
            cat.Checked = cat.File == filename
        
        #create default file if not exists
        if not os.path.isfile(file):
            cls.reset_own_colors()
            return

        #load file if exists
        with io.open(file, 'r', encoding='utf-8') as json_file:
            values = json.load(json_file)
        for i,v in enumerate(values):
            cls._userdefined[i].from_json(v)
            # data = json.load(json_file)
            # bkt.console.show_message("%r" % data)


    @classmethod
    def _get_shape_range(cls, selection):
        if selection.Type not in [2,3]:
            raise TypeError("nothing is selected")
        if selection.HasChildShapeRange:
            return selection.ChildShapeRange
        else:
            return selection.ShapeRange

    @classmethod
    def _get_color_from_selection(cls, selection):
        if selection.Type not in [2,3]:
            return
        
        shaperange = cls._get_shape_range(selection)
        shpcol = shaperange[1].Fill.ForeColor
        return PPTColor.from_color_obj(shpcol)
    
    @classmethod
    def _set_forecolor(cls, obj, color):
        obj.Visible = -1
        color.apply_to_color_obj(obj.ForeColor)
        # cls._set_color(obj.ForeColor, color)

    @classmethod
    def _compare_color(cls, obj, color):
        if obj.Visible == 0:
            return False

        if color == PPTColor.from_color_obj( obj.ForeColor ):
            return True

        # if color.is_theme_color:
        #     if obj_color.Type == pplib.MsoColorType['msoColorTypeScheme'] and obj_color.ObjectThemeColor == color.get_color_index() and obj_color.Brightness == color.get_brightness():
        #         return True
        # else:
        #     if obj_color.RGB == color.get_color_rgb():
        #         return True
        return False

    @classmethod
    def _color_key(cls, obj):
        if obj.Visible == 0:
            return None
        return PPTColor.from_color_obj(obj.ForeColor).get_color_tuple()
        # if obj.ForeColor.Type == pplib.MsoColorType['msoColorTypeScheme']:
        #     return (COLOR_THEME, obj.ForeColor.ObjectThemeColor, obj.ForeColor.Brightness)
        # else:
        #     return (COLOR_RGB, obj.ForeColor.RGB)


    @classmethod
    def pickup_color(cls, context, selected_color=None):
        shift = bkt.get_key_state(bkt.KeyCodes.SHIFT)
        selection = context.app.ActiveWindow.Selection

        color_from_selection = cls._get_color_from_selection(selection)
        selected_color = selected_color or color_from_selection

        if shift or selection.Type not in [2,3]:
            if selected_color:
                color_rgb = cls._show_color_dialog(selected_color.color_rgb)
            else:
                color_rgb = cls._show_color_dialog()
            if color_rgb is None:
                return
            else:
                return PPTColor.new_rgb(color_rgb)
        else:
            return color_from_selection

    @classmethod
    def pickup_own_color(cls, qebutton, context):
        new_color = cls.pickup_color(context, qebutton.get_color())
        if new_color:
            # cls._add_own_color(context, color)
            # color.set_userdefined_rgb(new_color)
            qebutton.set_color(new_color)
            cls.update_colors(context) #update colors here in order to define shade_index
            cls.update_pressed(context)
            cls.save_to_config()

    @classmethod
    def pickup_recent_color(cls, context):
        color = cls.pickup_color(context)
        if color:
            cls._add_recent_color(context, color.color_rgb)
            # cls.update_colors(context)
            cls.update_pressed(context)
    
    @classmethod
    def _show_color_dialog(cls, color_rgb=None):
        F = bkt.dotnet.import_forms()
        cd = F.ColorDialog()
        if color_rgb is not None:
            cd.Color = D.ColorTranslator.FromOle(color_rgb)
        cd.FullOpen = True
        if cd.ShowDialog() == F.DialogResult.OK:
            color = D.ColorTranslator.ToOle(cd.Color)
            return color
        else:
            return None

    @classmethod
    def _add_recent_color(cls, context, color_rgb):
        context.app.ActivePresentation.ExtraColors.Add(color_rgb)
        cls.update_colors(context)


    @classmethod
    def update_pressed(cls, context):
        try:
            shaperange = cls._get_shape_range(context.selection)
            if shaperange.HasTable:
                QEColorButtons.set_active_colors( ) #nothing selected
                # QEColorButtons.set_active_colors( [cls._color_key(shaperange.Fill)] ) #NOTE: this causes a strange behavior when selecting text in a table cell, the whole text gets selected (sometimes)
            else:
                QEColorButtons.set_active_colors( set([cls._color_key(shaperange.Fill), cls._color_key(shaperange.Line)]) )
        except:
            QEColorButtons.set_active_colors( ) #nothing selected
            # bkt.helpers.exception_as_message()

    @classmethod
    def update_colors(cls, context):
        QEColorButtons.update_colors(context)

    @classmethod
    def reset_own_colors(cls):
        default = [192, 255, 49407, 65535, 5296274, 5287936, 15773696, 12611584, 6299648, 10498160]
        for i, color in enumerate(default):
            # cls._usercolors[i] = color
            cls._userdefined[i].set_userdefined_rgb(color)
        cls.save_to_config()


    # @classmethod
    # def _set_color(cls, obj, color):
    #     if color.is_theme_color:
    #         obj.ObjectThemeColor = color.get_color_index()
    #         obj.Brightness = color.get_brightness()
    #     else:
    #         obj.RGB = color.get_color_rgb()

    # @classmethod
    # def _get_index(cls, current_control):
    #     return current_control['tag']

    # @classmethod
    # def get_image_by_control(cls, current_control, context, size=16):
    #     #if no slide selected, get color from slidemaster
    #     try:
    #         color = cls._get_color_rgb(context, cls._get_index(current_control))
    #     except:
    #         color = 0
    #     return cls.get_image(size, color)

    # @classmethod
    # def get_pressed_by_control(cls, context, current_control):
    #     color_id = cls._get_index(current_control)
    #     color = QEColor.get_color(color_id)
    #     return color.get_checked()
    #     # return cls.get_pressed(context, color)

    # @classmethod
    # def get_pressed(cls, context, color):
    #     try:
    #         shaperange = cls._get_shape_range(context.selection)
    #         if cls._compare_color(shaperange.Fill, color) or cls._compare_color(shaperange.Line, color):
    #             return True
    #     except:
    #         return False
    #     return False

    # @classmethod
    # def get_pressed(cls, current_control, context, shapes):
    #     color_index = cls._get_index(current_control)

    #     if color_index in cls._buttons2 and (color_index%20) > context.app.ActivePresentation.ExtraColors.Count-1:
    #         #undefined recent color
    #         return False

    #     shaperange = cls._get_shape_range(context.selection)
    #     color_rgb = cls._get_color_rgb(context, color_index)

    #     try:
    #         if cls._compare_color(context, shaperange.Fill, color_index, color_rgb) or cls._compare_color(context, shaperange.Line, color_index, color_rgb):
    #             return True
    #     except:
    #         return False

    #     # This version is too performance expensive:
    #     # for shape in shapes:
    #     #     try:
    #     #         if cls._compare_color(context, shape.Fill, color_index, color_rgb) or cls._compare_color(context, shape.Line, color_index, color_rgb):
    #     #             return True
    #     #         # if color_index in cls._buttons1:
    #     #         #     if (shape.Fill.Visible == -1 and shape.Fill.ForeColor.ObjectThemeColor == color_index and shape.Fill.ForeColor.Brightness == 0) or (shape.Line.Visible == -1 and shape.Line.ForeColor.ObjectThemeColor == color_index):
    #     #         #         return True
    #     #         # else:
    #     #         #     if (shape.Fill.Visible == -1 and shape.Fill.ForeColor.RGB == color_rgb) or (shape.Line.Visible == -1 and shape.Line.ForeColor.RGB == color_rgb):
    #     #         #         return True
    #     #     except:
    #     #         continue
    #     return False

    # @classmethod
    # def get_pressed_no_fill(cls, context, shapes):
    #     for shape in shapes:
    #         try:
    #             if shape.Fill.Visible == 0 or shape.Line.Visible == 0:
    #                 return True
    #         except:
    #             continue
    #     return False

    # @classmethod
    # def get_enabled(cls, current_control, context):
    #     color_index = cls._get_index(current_control)
    #     if color_index in cls._buttons1:
    #         return True
    #     else:
    #         return context.app.ActivePresentation.ExtraColors.Count > (color_index%20)

    # @classmethod
    # def action_by_control(cls, current_control, context, pressed=False):
    #     color_id = cls._get_index(current_control)
    #     color = QEColor.get_color(color_id)
    #     cls.action(color, context, pressed)

    @classmethod
    def action(cls, qebutton, context, pressed=False):
        shift = bkt.get_key_state(bkt.KeyCodes.SHIFT)
        ctrl  = bkt.get_key_state(bkt.KeyCodes.CTRL)
        alt   = bkt.get_key_state(bkt.KeyCodes.ALT)
        # color_index = cls._get_index(current_control)
        
        selection = context.app.ActiveWindow.Selection
        color = qebutton.get_color()

        # if color_index in cls._buttons2 and (color_index%20) > context.app.ActivePresentation.ExtraColors.Count-1:
        if not qebutton.is_defined:
            if qebutton.is_recent:
                #define recent colors
                cls.pickup_recent_color(context)
            else:
                #define own colors
                cls.pickup_own_color(qebutton, context)
        elif shift or selection.Type not in [2,3]:
            #select shapes by color
            if selection.HasChildShapeRange:
                shapes = selection.ShapeRange[1].GroupItems
            else:
                shapes = selection.SlideRange[1].Shapes
            # color_rgb = cls._get_color_rgb(context, color_index)
            for shape in shapes:
                try:
                    if alt and ctrl:
                        # if shape.HasTextFrame == -1 and shape.TextFrame2.TextRange.Font.Line.Visible == -1 and shape.TextFrame2.TextRange.Font.Line.ForeColor.RGB == color_rgb:
                        if shape.HasTextFrame == -1 and cls._compare_color(shape.TextFrame2.TextRange.Font.Line, color):
                            shape.Select(replace=False)
                    elif alt:
                        # if shape.HasTextFrame == -1 and shape.TextFrame2.TextRange.Font.Fill.Visible == -1 and shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB == color_rgb:
                        if shape.HasTextFrame == -1 and cls._compare_color(shape.TextFrame2.TextRange.Font.Fill, color):
                            shape.Select(replace=False)
                    elif ctrl:
                        # if shape.Line.Visible == -1 and shape.Line.ForeColor.RGB == color_rgb:
                        if cls._compare_color(shape.Line, color):
                            shape.Select(replace=False)
                    else:
                        # if shape.Fill.Visible == -1 and shape.Fill.ForeColor.RGB == color_rgb:
                        if cls._compare_color(shape.Fill, color):
                            shape.Select(replace=False)
                except:
                    continue
        else:
            #shapes or text selected, apply color
            shaperange = cls._get_shape_range(selection)
            if alt and ctrl:
                try:
                    if selection.TextRange2.Count == 0:
                        raise TypeError("no text selected, fallback to shape")
                    cls._set_forecolor( selection.TextRange2.Font.Line, color)
                except: #TypeError, COMException (e.g. if table/table cells are selected)
                    for textframe in pplib.iterate_shape_textframes(shaperange):
                        try:
                            cls._set_forecolor( textframe.TextRange.Font.Line, color)
                        except:
                            continue
            elif alt:
                try:
                    if selection.TextRange2.Count == 0:
                        raise TypeError("no text selected, fallback to shape")
                    cls._set_forecolor( selection.TextRange2.Font.Fill, color)
                except: #TypeError, COMException (e.g. if table/table cells are selected)
                    for textframe in pplib.iterate_shape_textframes(shaperange):
                        try:
                            cls._set_forecolor( textframe.TextRange.Font.Fill, color)
                        except:
                            continue
            elif ctrl or shaperange.Connector == -1: #only connectors selected
                # if shaperange.HasTable == 0:
                #     cls._set_forecolor(context, shaperange.Line, color_index)
                # else:
                # NOTE: Line property of shape range does not change all lines depending on selected shapes,
                #       e.g. when shapes and connectores are selected, only connector color is changed
                for shape in pplib.iterate_shape_subshapes(shaperange):
                    try:
                        cls._set_forecolor(shape.Line, color)
                    except:
                        #table cell shapes do not have line property #FIXME: set borders instead
                        continue
            else:
                # if shaperange.HasTable == 0:
                #     shaperange.Fill.Solid() #switch to solid background
                #     cls._set_forecolor(context, shaperange.Fill, color_index)
                # else:
                # NOTE: Better to iterate shapes, see also line property
                for shape in pplib.iterate_shape_subshapes(shaperange):
                    try:
                        shape.Fill.Solid() #switch to solid background
                        cls._set_forecolor(shape.Fill, color)
                    except:
                        continue
            cls.update_pressed(context)

    @classmethod
    def action_no_fill(cls, context):
        shift = bkt.get_key_state(bkt.KeyCodes.SHIFT)
        ctrl  = bkt.get_key_state(bkt.KeyCodes.CTRL)
        alt   = bkt.get_key_state(bkt.KeyCodes.ALT)
        
        selection = context.app.ActiveWindow.Selection

        if shift or selection.Type not in [2,3]:
            #select shapes by color
            if selection.HasChildShapeRange:
                shapes = selection.ShapeRange[1].GroupItems
            else:
                shapes = selection.SlideRange[1].Shapes
            for shape in shapes:
                try:
                    if alt and ctrl:
                        if shape.HasTextFrame == -1 and shape.TextFrame2.TextRange.Font.Line.Visible == 0:
                            shape.Select(replace=False)
                    elif alt:
                        if shape.HasTextFrame == -1 and shape.TextFrame2.TextRange.Font.Fill.Visible == 0:
                            shape.Select(replace=False)
                    elif ctrl:
                        if shape.Line.Visible == 0:
                            shape.Select(replace=False)
                    else:
                        if shape.Fill.Visible == 0:
                            shape.Select(replace=False)
                except:
                    continue
        else:
            #shapes or text selected, apply color
            shaperange = cls._get_shape_range(selection)
            if alt and ctrl:
                # NOTE: deactivation text outline is not implemented in ppt object model, no workaround available
                try:
                    if selection.TextRange2.Count == 0:
                        raise Exception("no text selected, fallback to shape")
                    selection.TextRange2.Font.Line.Visible = 0
                except:
                    for textframe in pplib.iterate_shape_textframes(shaperange):
                        try:
                            textframe.TextRange.Font.Line.Visible = 0
                        except:
                            continue
                bkt.message("Sorry, aber Microsoft hat die Funktion zur Deaktivierung der Textkontur nicht in die Schnittstelle implementiert.")
            elif alt:
                try:
                    if selection.TextRange2.Count == 0:
                        raise Exception("no text selected, fallback to shape")
                    selection.TextRange2.Font.Fill.Visible = 0
                except:
                    for textframe in pplib.iterate_shape_textframes(shaperange):
                        try:
                            textframe.TextRange.Font.Fill.Visible = 0
                        except:
                            continue
            elif ctrl:
                # if shaperange.HasTable == 0:
                #     shaperange.Line.Visible = 0
                # else:
                for shape in pplib.iterate_shape_subshapes(shaperange):
                    try:
                        shape.Line.Visible = 0
                    except:
                        continue
            else:
                # if shaperange.HasTable == 0:
                #     shaperange.Fill.Visible = 0
                # else:
                for shape in pplib.iterate_shape_subshapes(shaperange):
                    try:
                        shape.Fill.Visible = 0
                    except:
                        continue
            cls.update_pressed(context)

    @classmethod
    def action_transparency(cls, context, delta):
        ctrl  = bkt.get_key_state(bkt.KeyCodes.CTRL)

        selection = context.app.ActiveWindow.Selection

        if selection.Type in [2,3]:
            #select shapes by color
            shaperange = cls._get_shape_range(selection)
            for shape in pplib.iterate_shape_subshapes(shaperange):
                try:
                    if ctrl:
                        if shape.Line.Visible == -1:
                            shape.Line.Transparency = max(0,min(1,shape.Line.Transparency+delta))
                    else:
                        if shape.Fill.Visible == -1:
                            shape.Fill.Transparency = max(0,min(1,shape.Fill.Transparency+delta))
                except:
                    continue

    @staticmethod
    def show_help():
        helpfile = bkt.helpers.file_base_path_join(__file__, "resources", "QuickEdit Help.pdf")
        try:
            from os import startfile
            startfile(helpfile)
        except:
            logging.error("QuickEdit: Error opening the help file.")
            bkt.message("Fehler beim Öffnen der PDF-Hilfedatei. Bitte Datei manuell öffnen: {}".format(helpfile))

#         help_msg = '''
# 1. Reihe: Farben des Design-Farbschemas der aktuellen Folie.
# 2. Reihe: Zuletzt verwendete Farben (außerhalb des Farbschemas).
#           Sind (noch) keine zuletzt verwendeten Farben definiert, sind die Buttons grau. Über den "Pickup"-
#           Button kann eine Farbe hinzugefügt, entweder die Hintergrundfarbe des markierten Shapes, oder
#           mittels eines Farbmischers (wenn nichts markiert ist oder bei gedrückter SHIFT-Taste).

# Folgende Funktionen stehen zur Verfügung:

# [Ohne Tasten]:    Setzt Hintergrundfarbe der selektierten Shapes auf gewählte Farbe.
#                   Ist kein Shape selektiert, werden alle Shapes mit der gewählten Hintergrundfarbe markiert.
# [STRG]:           Setzt Liniefarbe der selektierten Shapes auf gewählte Farbe.
#                   Ist kein Shape selektiert, werden alle Shapes mit der gewählten Linienfarbe markiert.
# [ALT]:            Setzt Textfarbe der selektierten Shapes auf gewählte Farbe.
#                   Ist kein Shape selektiert, werden alle Shapes mit der gewählten Textfarbe markiert.
# [ALT+STRG]:       Setzt Textkontur der selektierten Shapes auf gewählte Farbe.
#                   Ist kein Shape selektiert, werden alle Shapes mit der gewählten Textkontur markiert.

# [SHIFT]:          Selektiert alle Shapes mit entsprechender Hintergrundfarbe.
# [SHIFT+STRG]:     Selektiert alle Shapes mit entsprechender Linienfarbe.
# [SHIFT+ALT]:      Selektiert alle Shapes mit entsprechender Textfarbe.
# [SHIFT+ALT+STRG]: Selektiert alle Shapes mit entsprechender Textkontur.

# Hinweis für Mac-Nutzer: Je nach Einstellung fängt Parallels einige Tastenkombinationen ab.
# Hinweis für Experten: Die 2. Reihe weist ausschließlich RGB-Werte zu und nicht Farben des Farbschemas.
# '''
#         # bkt.message(help_msg)
#         import bkt.console
#         bkt.console.show_message(bkt.ui.endings_to_windows(help_msg))


QuickEdit.initialize()
# bkt.AppEvents.bkt_load += bkt.Callback(QuickEdit.read_from_config)

bkt.AppEvents.selection_changed       += bkt.Callback(QuickEdit.update_pressed, context=True)
bkt.AppEvents.slide_selection_changed += bkt.Callback(QuickEdit.update_colors, context=True)