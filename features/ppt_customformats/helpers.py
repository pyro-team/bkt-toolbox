# -*- coding: utf-8 -*-
'''
Created on 2018-05-29
@author: Florian Stallmann
'''



import logging

from collections import OrderedDict
from functools import wraps

import bkt.library.powerpoint as pplib


def textframe_group_check(func):
    @wraps(func)
    def wrapper(cls, textframe_obj, *args, **kwargs):
        try:
            shape = textframe_obj.Parent
            if shape.Type == pplib.MsoShapeType["msoGroup"]:
                logging.debug("customformats: found group")
                for shp in shape.GroupItems:
                    func(cls, shp.TextFrame2, *args, **kwargs)
            else:
                logging.debug("customformats: found non-group")
                func(cls, textframe_obj, *args, **kwargs)
        except:
            logging.exception("customformats: group check failed")
            func(cls, textframe_obj, *args, **kwargs)
    return wrapper


class ShapeFormats(object):
    always_keep_theme_color = True
    always_consider_indentlevels = True

    @classmethod
    def mult_setattr(cls, obj, name, value):
        logging.debug("mult_setattr: setting %s = %s", name, value)
        attrs = name.split(".")
        for name in attrs[:-1]:
            try:
                obj = getattr(obj, name)
            except:
                raise AttributeError("Cannot find attribute %s" % name)
        try:
            if attrs[-1] == "BKTColor":
                obj.ObjectThemeColor = value[0]
                obj.Brightness = value[1]
                # If theme color is different from saved RGB color, use RGB color instead
                if not cls.always_keep_theme_color and obj.RGB != value[2]:
                    obj.RGB = value[2]
            else:
                setattr(obj, attrs[-1], value)
        except ValueError:
            # bkt.helpers.exception_as_message("Cannot set %s = %s" % (attrs[-1], value))
            raise ValueError("Cannot set %s = %s" % (attrs[-1], value))
            
    @classmethod
    def _write_color_to_array(cls, dict_ref, color_obj, arr_key):
        if color_obj.Type == pplib.MsoColorType['msoColorTypeScheme'] and color_obj.ObjectThemeColor > 0:
            dict_ref['%s.BKTColor' % arr_key] = [color_obj.ObjectThemeColor, float(color_obj.Brightness), color_obj.RGB]
            # dict_ref['%s.ObjectThemeColor' % arr_key] = color_obj.ObjectThemeColor
            # dict_ref['%s.Brightness' % arr_key] = float(color_obj.Brightness)
        else:
            dict_ref['%s.RGB' % arr_key] = color_obj.RGB
        # dict_ref['%s.RGB' % arr_key] = color_obj.RGB #always use RGB for now as it can be used cross-presentations
        dict_ref['%s.TintAndShade' % arr_key] = float(color_obj.TintAndShade)

    @classmethod
    def _get_indentlevels(cls, textframe_object, what): #paragraph or font
        indent_levels = OrderedDict()
        if textframe_object.TextRange.Paragraphs().Count > 0:
            # at least one paragraph
            indent_level_range = list(range(0,6)) #indent levels 0 to 5, whereas 0 is used as internal fallback format!
            for par in textframe_object.TextRange.Paragraphs():
                indent_level = par.ParagraphFormat.IndentLevel
                if indent_level == 1 and par.ParagraphFormat.Bullet.Visible == 0:
                    indent_level = 0 #fallback indent level

                if indent_level in indent_level_range:
                    indent_level_range.remove(indent_level)
                    indent_levels[str(indent_level)] = cls._get_indentlevel_formats(par, what)
            if 0 in indent_level_range:
                #fallback not yet defined
                indent_levels["0"] = cls._get_indentlevel_formats(textframe_object.TextRange.Paragraphs(1,1), what)
        else:
            indent_levels["0"] = cls._get_indentlevel_formats(textframe_object.TextRange, what)
        return indent_levels

    @classmethod
    def _get_indentlevel_formats(cls, textrange_object, what):
        if what == "paragraph":
            return cls._get_paragraphformat(textrange_object.ParagraphFormat)
        else:
            return cls._get_font(textrange_object.Font)
    
    @classmethod
    @textframe_group_check
    def _set_indentlevels(cls, textframe_object, what, indentlevels_dict):
        if cls.always_consider_indentlevels and textframe_object.TextRange.Paragraphs().Count > 0:
            for par in textframe_object.TextRange.Paragraphs():
                indent_level = str(par.ParagraphFormat.IndentLevel)
                if indent_level not in indentlevels_dict or (indent_level == "1" and par.ParagraphFormat.Bullet.Visible == 0):
                    indent_level = "0"
                cls._set_indentlevel_formats(par, what, indentlevels_dict[indent_level])
        else:
            cls._set_indentlevel_formats(textframe_object.TextRange, what, indentlevels_dict["0"])

    @classmethod
    def _set_indentlevel_formats(cls, textrange_object, what, what_dict):
        if what == "paragraph":
            cls._set_paragraphformat(textrange_object.ParagraphFormat, what_dict)
        else:
            cls._set_font(textrange_object.Font, what_dict)

    @classmethod
    def _get_type(cls, shape):
        tmp = OrderedDict()
        if shape.Connector == -1:
            tmp['ConnectorFormat.Type'] = shape.ConnectorFormat.Type
        else:
            #for connectors, autoshapetype is -2 and throws error setting this value
            tmp['AutoShapeType']  = shape.AutoShapeType
        tmp['VerticalFlip']   = shape.VerticalFlip #method
        tmp['HorizontalFlip'] = shape.HorizontalFlip #method
        tmp['Adjustments'] = [
                                float(shape.adjustments.item[i])
                                for i in range(1,shape.adjustments.count+1)
                                ]
        return tmp
    @classmethod
    def _set_type(cls, shape, type_dict):
        logging.debug("customformats: set type")

        if shape.Connector == -1 and "ConnectorFormat.Type" in type_dict:
            shape.ConnectorFormat.Type = type_dict["ConnectorFormat.Type"]
        elif "AutoShapeType" in type_dict:
            shape.AutoShapeType = type_dict["AutoShapeType"]

        if shape.VerticalFlip != type_dict["VerticalFlip"]:
            shape.Flip(1) #msoFlipVertical

        if shape.HorizontalFlip != type_dict["HorizontalFlip"]:
            shape.Flip(0) #msoFlipHorizontal

        for i in range(1,shape.adjustments.count+1):
            try:
                shape.adjustments.item[i] = type_dict["Adjustments"][i-1]
            except:
                continue

    @classmethod
    def _get_fill(cls, fill_object):
        tmp = OrderedDict()
        if fill_object.Visible == -1:
            tmp['Visible'] = -1

            if fill_object.Type == pplib.MsoFillType['msoFillBackground']:
                tmp['Background'] = True #method!

            elif fill_object.Type == pplib.MsoFillType['msoFillPatterned']:
                tmp['Pattern'] = fill_object.Pattern #read-only attribute!
                cls._write_color_to_array(tmp, fill_object.ForeColor, 'ForeColor')
                cls._write_color_to_array(tmp, fill_object.BackColor, 'BackColor')

            elif fill_object.Type == pplib.MsoFillType['msoFillGradient']:
                save_color_stops = True
                if fill_object.GradientColorType == 1: #msoGradientOneColor
                    tmp['GradientOneColor'] = [
                                                                            fill_object.GradientStyle,
                                                                            fill_object.GradientVariant,
                                                                            float(fill_object.GradientDegree),
                                                                        ]
                elif fill_object.GradientColorType == 2: #msoGradientTwoColors
                    tmp['GradientTwoColor'] = [
                                                                            fill_object.GradientStyle,
                                                                            fill_object.GradientVariant,
                                                                        ]
                elif fill_object.GradientColorType == 3: #msoGradientPresetColors
                    tmp['GradientPresetColor'] = [
                                                                            fill_object.GradientStyle,
                                                                            fill_object.GradientVariant,
                                                                            fill_object.PresetGradientType,
                                                                        ]
                    save_color_stops = False #no need to save color stops for preset gradients
                elif fill_object.GradientColorType == 4: #msoGradientMultiColor
                    tmp['GradientMultiColor'] = [
                                                                            fill_object.GradientStyle,
                                                                            fill_object.GradientVariant,
                                                                        ]
                else:
                    raise ValueError('unkown gradient type')

                #NOTE: If angle is changed (for linear gradients), style can be -2 and variant 0 which are invalid values! This is handled is the setter function.

                if save_color_stops:
                    tmp['GradientStops'] = []
                    for stop in fill_object.GradientStops:
                        stop_dict = OrderedDict()
                        stop_dict["Position"] = float(stop.Position)
                        cls._write_color_to_array(stop_dict, stop.Color, 'Color')
                        stop_dict["Transparency"] = float(stop.Transparency) #IMPORTANT: Set Transparency after color, because color resets transparency
                        tmp['GradientStops'].append(stop_dict)
                                    #     (stop.color.rgb,
                                    #     float(stop.Position),
                                    #     float(stop.Transparency),
                                    #     i+1,
                                    #     float(stop.color.brightness))
                                    #     for i,stop in enumerate(fill_object.GradientStops)
                
                tmp['RotateWithObject'] = fill_object.RotateWithObject
                try:
                    #angle is only accessible for certain gradient types/styles/variants...
                    tmp['GradientAngle'] = float(fill_object.GradientAngle)
                except:
                    pass

            #elif fill_object.Type == pplib.MsoFillType['msoFillTextured']:
            # Textures in VBA is broken, property PresetTexture always returns -2
            
            else:
                tmp['Solid'] = True #method!
                cls._write_color_to_array(tmp, fill_object.ForeColor, 'ForeColor')
                tmp['Transparency'] = float(fill_object.Transparency)
            
            # tmp['Type'] = fill_object.Type #read-only attribute!
        else:
            tmp['Visible'] = 0
        return tmp
    @classmethod
    def _set_fill(cls, fill_object, fill_dict):
        logging.debug("customformats: set fill")
        for key, value in fill_dict.items():
            if key == "Pattern":
                    fill_object.Patterned(value)
            elif key == "Background":
                    fill_object.Background()
            elif key == "Solid":
                    fill_object.Solid()
            elif key == "GradientOneColor":
                # fill_object.OneColorGradient(*value) #style, variant, degree
                fill_object.OneColorGradient(max(1,value[0]), max(1,value[1]), value[2]) #style, variant, degree
            elif key == "GradientTwoColor":
                fill_object.TwoColorGradient(max(1,value[0]), max(1,value[1])) #style, variant
            elif key == "GradientPresetColor":
                fill_object.PresetGradient(max(1,value[0]), max(1,value[1]), value[2]) #style, variant, preset-gradient-type
            elif key == "GradientMultiColor":
                fill_object.TwoColorGradient(max(1,value[0]), max(1,value[1])) #style, variant
            elif key == "GradientStops":
                cur_stops = fill_object.GradientStops.Count
                for i in range(max(cur_stops, len(value))):
                    if i > len(value):
                        fill_object.GradientStops.Delete(i+1)
                        continue
                    elif i < cur_stops:
                        pass
                        # fill_object.GradientStops[i+1].color.rgb        = value[i][0]
                        # fill_object.GradientStops[i+1].Position         = value[i]["Position"]
                        # fill_object.GradientStops[i+1].Transparency     = value[i]["Transparency"]
                        # fill_object.GradientStops[i+1].color.brightness = value[i][4]
                    else:
                        # fill_object.GradientStops.Insert2(*value[i])
                        fill_object.GradientStops.Insert(1, 1.0) #rgb, position
                    
                    stop_object = fill_object.GradientStops[i+1]
                    for k, v in value[i].items():
                        # logging.debug("Setting %s = %s", k, v)
                        cls.mult_setattr(stop_object, k, v)
            else:
                cls.mult_setattr(fill_object, key, value)

    @classmethod
    def _get_line(cls, line_object):
        tmp = OrderedDict()
        if line_object.Visible == -1:
            #NOTE: Line gradient not supported via VBA
            tmp['Visible'] = -1
            cls._write_color_to_array(tmp, line_object.ForeColor, 'ForeColor')
            cls._write_color_to_array(tmp, line_object.BackColor, 'BackColor')
            tmp['Style'] = line_object.Style
            tmp['DashStyle'] = line_object.DashStyle
            tmp['Weight'] = float(line_object.Weight)
            tmp['Transparency'] = max(0, float(line_object.Transparency)) #NOTE: transparency can be -2.14748e+09 if line gradient is active
            # tmp['InsetPen'] = line_object.InsetPen #NOTE: this property is not accessible via UI as it was default until PPT97
            #the following properties are relevant for connectors and special shapes, e.g. freeform-line. other shapes will throw ValueError
            tmp['BeginArrowheadLength'] = line_object.BeginArrowheadLength
            tmp['BeginArrowheadStyle']  = line_object.BeginArrowheadStyle
            tmp['BeginArrowheadWidth']  = line_object.BeginArrowheadWidth
            tmp['EndArrowheadLength'] = line_object.EndArrowheadLength
            tmp['EndArrowheadStyle']  = line_object.EndArrowheadStyle
            tmp['EndArrowheadWidth']  = line_object.EndArrowheadWidth
        else:
            tmp['Visible'] = 0
        return tmp
    @classmethod
    def _set_line(cls, line_object, line_dict):
        logging.debug("customformats: set line")
        for key, value in line_dict.items():
            cls.mult_setattr(line_object, key, value)

    @classmethod
    def _get_shadow(cls, shadow_object):
        tmp = OrderedDict()
        if shadow_object.Visible == -1:
            tmp['Visible'] = -1
            if shadow_object.Type != -2: #msoShadowMixed
                tmp['Type'] = shadow_object.Type
                cls._write_color_to_array(tmp, shadow_object.ForeColor, 'ForeColor')
            else:
                tmp['Style'] = shadow_object.Style
                cls._write_color_to_array(tmp, shadow_object.ForeColor, 'ForeColor')
                tmp['Size'] = float(shadow_object.Size)
                tmp['Blur'] = float(shadow_object.Blur)
                tmp['OffsetX'] = float(shadow_object.OffsetX)
                tmp['OffsetY'] = float(shadow_object.OffsetY)
                tmp['Transparency'] = float(shadow_object.Transparency)
        else:
            tmp['Visible'] = 0
        return tmp
    @classmethod
    def _set_shadow(cls, shadow_object, shadow_dict):
        logging.debug("customformats: set shadow")
        for key, value in shadow_dict.items():
            cls.mult_setattr(shadow_object, key, value)

    @classmethod
    def _get_glow(cls, glow_object):
        tmp = OrderedDict()
        if glow_object.Radius > 0:
            cls._write_color_to_array(tmp, glow_object.Color, 'Color')
            tmp['Radius'] = float(glow_object.Radius)
            tmp['Transparency'] = float(glow_object.Transparency)
        else:
            tmp['Radius'] = 0.0
        return tmp
    @classmethod
    def _set_glow(cls, glow_object, glow_dict):
        logging.debug("customformats: set glow")
        for key, value in glow_dict.items():
            cls.mult_setattr(glow_object, key, value)

    @classmethod
    def _get_softedge(cls, softedge_object):
        tmp = OrderedDict()
        if softedge_object.Radius > 0:
            tmp['Type'] = softedge_object.Type
            tmp['Radius'] = float(softedge_object.Radius)
        else:
            tmp['Radius'] = 0.0
        return tmp
    @classmethod
    def _set_softedge(cls, softedge_object, softedge_dict):
        logging.debug("customformats: set softedge")
        for key, value in softedge_dict.items():
            cls.mult_setattr(softedge_object, key, value)

    @classmethod
    def _get_reflection(cls, reflection_object):
        tmp = OrderedDict()
        if reflection_object.Type > 0:
            tmp['Type'] = reflection_object.Type
            tmp['Blur'] = float(reflection_object.Blur)
            tmp['Offset'] = float(reflection_object.Offset)
            tmp['Size'] = float(reflection_object.Size)
            tmp['Transparency'] = float(reflection_object.Transparency)
        else:
            tmp['Type'] = 0
        return tmp
    @classmethod
    def _set_reflection(cls, reflection_object, reflection_dict):
        logging.debug("customformats: set reflection")
        for key, value in reflection_dict.items():
            cls.mult_setattr(reflection_object, key, value)

    @classmethod
    def _get_textframe(cls, textframe_object):
        tmp = OrderedDict()
        tmp['HorizontalAnchor'] = textframe_object.HorizontalAnchor
        tmp['VerticalAnchor'] = textframe_object.VerticalAnchor
        tmp['Orientation'] = textframe_object.Orientation
        tmp['AutoSize'] = textframe_object.AutoSize
        tmp['WordWrap'] = textframe_object.WordWrap
        tmp['MarginBottom'] = float(textframe_object.MarginBottom)
        tmp['MarginLeft'] = float(textframe_object.MarginLeft)
        tmp['MarginRight'] = float(textframe_object.MarginRight)
        tmp['MarginTop'] = float(textframe_object.MarginTop)
        tmp['Column.Number'] = textframe_object.Column.Number
        tmp['Column.Spacing'] = float(textframe_object.Column.Spacing)
        return tmp
    @classmethod
    def _set_textframe(cls, textframe_object, textframe_dict):
        logging.debug("customformats: set textframe")
        for key, value in textframe_dict.items():
            cls.mult_setattr(textframe_object, key, value)

    @classmethod
    def _get_font(cls, font_object):
        tmp = OrderedDict()
        # #Font color
        # if font_object.Fill.Visible == -1:
        #     tmp['Fill.Visible'] = -1
        #     cls._write_color_to_array(tmp, font_object.Fill.ForeColor, 'Fill.ForeColor')
        # else:
        #     tmp['Fill.Visible'] = 0
        # #Font line color
        # if font_object.Line.Visible == -1:
        #     tmp['Line.Visible'] = -1
        #     cls._write_color_to_array(tmp, font_object.Line.ForeColor, 'Line.ForeColor')
        # else:
        #     tmp['Line.Visible'] = 0 #NOTE: this is not working in VBA
        
        #Font setting
        tmp['Name'] = font_object.Name
        tmp['Size'] = float(font_object.Size)
        tmp['Bold'] = font_object.Bold
        tmp['Italic'] = font_object.Italic
        tmp['UnderlineStyle'] = font_object.UnderlineStyle
        if font_object.UnderlineColor.Type > 0:
            cls._write_color_to_array(tmp, font_object.UnderlineColor, 'UnderlineColor')
        tmp['Caps'] = font_object.Caps
        tmp['Strike'] = font_object.Strike
        tmp['Kerning'] = float(font_object.Kerning)
        tmp['Spacing'] = float(font_object.Spacing)
        
        #Fill, line and all effects objects
        tmp['Fill']       = cls._get_fill(font_object.Fill)
        tmp['Line']       = cls._get_line(font_object.Line)
        tmp['Glow']       = cls._get_glow(font_object.Glow)
        tmp['Reflection'] = cls._get_reflection(font_object.Reflection)
        tmp['Shadow']     = cls._get_shadow(font_object.Shadow)
        #NOTE: Highlight property is not accessible via UI and cannot be disabled via VBA, so we don't use it
        return tmp
    @classmethod
    def _set_font(cls, font_object, font_dict):
        logging.debug("customformats: set font")
        for key, value in font_dict.items():
            if key == "Fill":
                try:
                    cls._set_fill(font_object.Fill, value)
                except:
                    logging.error("customformats: error in setting font fill")
            elif key == "Line":
                try:
                    cls._set_line(font_object.Line, value)
                except:
                    logging.error("customformats: error in setting font line")
            elif key == "Shadow":
                try:
                    cls._set_shadow(font_object.Shadow, value)
                except:
                    logging.error("customformats: error in setting font shadow")
            elif key == "Glow":
                try:
                    cls._set_glow(font_object.Glow, value)
                except:
                    logging.error("customformats: error in setting font glow")
            elif key == "Reflection":
                try:
                    cls._set_reflection(font_object.Reflection, value)
                except:
                    logging.error("customformats: error in setting font reflection")
            else:
                cls.mult_setattr(font_object, key, value)
    
    @classmethod
    def _get_paragraphformat(cls, parfor_object):
        tmp = OrderedDict()
        tmp['Alignment'] = parfor_object.Alignment
        tmp['BaselineAlignment'] = parfor_object.BaselineAlignment

        tmp['LineRuleAfter'] = parfor_object.LineRuleAfter
        tmp['SpaceAfter'] = float(parfor_object.SpaceAfter)
        tmp['LineRuleBefore'] = parfor_object.LineRuleBefore
        tmp['SpaceBefore'] = float(parfor_object.SpaceBefore)
        tmp['LineRuleWithin'] = parfor_object.LineRuleWithin
        tmp['SpaceWithin'] = float(parfor_object.SpaceWithin)

        #Bullet points
        if parfor_object.Bullet.Visible == -1:
            tmp['Bullet.Visible'] = -1
            btype = parfor_object.Bullet.Type
            tmp['Bullet.Type'] = btype
            if btype == 1: #ppBulletUnnumbered
                tmp['Bullet.Character'] = parfor_object.Bullet.Character
            elif btype == 2: #ppBulletNumbered
                tmp['Bullet.Style'] = parfor_object.Bullet.Style
                tmp['Bullet.StartValue'] = parfor_object.Bullet.StartValue
            tmp['Bullet.RelativeSize'] = float(parfor_object.Bullet.RelativeSize)
            if parfor_object.Bullet.UseTextFont == -1:
                tmp['Bullet.UseTextFont'] = -1
            else:
                tmp['Bullet.Font.Name'] = parfor_object.Bullet.Font.Name
            if parfor_object.Bullet.UseTextColor == -1:
                tmp['Bullet.UseTextColor'] = -1
            else:
                cls._write_color_to_array(tmp, parfor_object.Bullet.Font.Fill.ForeColor, 'Bullet.Font.Fill.ForeColor')
        else:
            tmp['Bullet.Type'] = 0
            tmp['Bullet.Visible'] = 0
        
        tmp['FirstLineIndent'] = float(parfor_object.FirstLineIndent)
        tmp['LeftIndent'] = float(parfor_object.LeftIndent)
        tmp['RightIndent'] = float(parfor_object.RightIndent)
        tmp['HangingPunctuation'] = parfor_object.HangingPunctuation

        tmp['TabStops.DefaultSpacing'] = float(parfor_object.TabStops.DefaultSpacing)
        tmp['TabStops.Items'] = [(ts.type, float(ts.position)) for ts in parfor_object.TabStops]
        return tmp
    @classmethod
    def _set_paragraphformat(cls, parfor_object, parfor_dict):
        logging.debug("customformats: set parformat")
        for key, value in parfor_dict.items():
            if key == 'TabStops.Items':
                #remove all tabstops
                for t in list(iter(parfor_object.TabStops)):
                    t.Clear()
                #add tabstops
                for ts_type, ts_pos in value:
                    parfor_object.TabStops.Add(ts_type, ts_pos)
            else:
                cls.mult_setattr(parfor_object, key, value)

    @classmethod
    def _get_size(cls, shape):
        tmp = OrderedDict()
        tmp['Width'] = float(shape.Width)
        tmp['Height'] = float(shape.Height)
        tmp['LockAspectRatio'] = shape.LockAspectRatio
        return tmp
    @classmethod
    def _set_size(cls, shape, size_dict):
        logging.debug("customformats: set size")
        shape.LockAspectRatio = 0
        shape.Width = size_dict["Width"]
        shape.Height = size_dict["Height"]
        shape.LockAspectRatio = size_dict["LockAspectRatio"]

    @classmethod
    def _get_position(cls, shape):
        tmp = OrderedDict()
        tmp['Left'] = float(shape.Left)
        tmp['Top'] = float(shape.Top)
        tmp['Rotation'] = float(shape.Rotation)
        return tmp
    @classmethod
    def _set_position(cls, shape, position_dict):
        logging.debug("customformats: set position")
        shape.Left = position_dict["Left"]
        shape.Top = position_dict["Top"]
        shape.Rotation = position_dict["Rotation"]