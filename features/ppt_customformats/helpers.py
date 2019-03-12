# -*- coding: utf-8 -*-
'''
Created on 2018-05-29
@author: Florian Stallmann
'''


import bkt.library.powerpoint as pplib
from collections import OrderedDict


class ShapeFormats(object):
    always_keep_theme_color = True

    @classmethod
    def mult_setattr(cls, obj, name, value):
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
    def _get_indentlevel_formats(cls, textrange_object):
        tmp = OrderedDict()
        tmp["ParagraphFormat"] = cls._get_paragraphformat(textrange_object.ParagraphFormat)
        tmp["Font"] = cls._get_font(textrange_object.Font)
        return tmp

    @classmethod
    def _get_type(cls, shape):
        tmp = OrderedDict()
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
                elif fill_object.GradientColorType == 4: #msoGradientMultiColor
                    tmp['GradientMultiColor'] = [
                                                                            fill_object.GradientStyle,
                                                                            fill_object.GradientVariant,
                                                                        ]
                else:
                    raise ValueError('unkown gradient type')
                tmp['GradientStops'] = [
                                                                    (stop.color.rgb,
                                                                    float(stop.Position),
                                                                    float(stop.Transparency),
                                                                    i+1,
                                                                    float(stop.color.brightness))
                                                                    for i,stop in enumerate(fill_object.GradientStops)
                                                                ]
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
        for key, value in fill_dict.items():
            if key == "Pattern":
                    fill_object.Patterned(value)
            elif key == "Background":
                    fill_object.Background()
            elif key == "Solid":
                    fill_object.Solid()
            elif key == "GradientOneColor":
                fill_object.OneColorGradient(*value)
            elif key == "GradientTwoColor":
                fill_object.TwoColorGradient(*value)
            elif key == "GradientPresetColor":
                fill_object.PresetGradient(*value)
            elif key == "GradientMultiColor":
                fill_object.TwoColorGradient(*value)
            elif key == "GradientStops":
                cur_stops = fill_object.GradientStops.Count
                for i in range(max(cur_stops, len(value))):
                    if i > len(value):
                        fill_object.GradientStops.Delete(i+1)
                    elif i < cur_stops:
                        fill_object.GradientStops[i+1].color.rgb        = value[i][0]
                        fill_object.GradientStops[i+1].Position         = value[i][1]
                        fill_object.GradientStops[i+1].Transparency     = value[i][2]
                        fill_object.GradientStops[i+1].color.brightness = value[i][4]
                    else:
                        fill_object.GradientStops.Insert2(*value[i])
            else:
                cls.mult_setattr(fill_object, key, value)

    @classmethod
    def _get_line(cls, line_object):
        tmp = OrderedDict()
        if line_object.Visible == -1:
            #FIXME: Add support for line gradient
            tmp['Visible'] = -1
            cls._write_color_to_array(tmp, line_object.ForeColor, 'ForeColor')
            cls._write_color_to_array(tmp, line_object.BackColor, 'BackColor')
            tmp['Style'] = line_object.Style
            tmp['DashStyle'] = line_object.DashStyle
            tmp['Weight'] = float(line_object.Weight)
            tmp['Transparency'] = float(line_object.Transparency)
        else:
            tmp['Visible'] = 0
        return tmp
    @classmethod
    def _set_line(cls, line_object, line_dict):
        for key, value in line_dict.items():
            cls.mult_setattr(line_object, key, value)

    @classmethod
    def _get_shadow(cls, shadow_object):
        tmp = OrderedDict()
        if shadow_object.Visible == -1:
            tmp['Visible'] = -1
            if shadow_object.Type > 0:
                tmp['Type'] = shadow_object.Type
            cls._write_color_to_array(tmp, shadow_object.ForeColor, 'ForeColor')
            tmp['Size'] = float(shadow_object.Size)
            tmp['Style'] = shadow_object.Style
            tmp['Blur'] = float(shadow_object.Blur)
            tmp['OffsetX'] = float(shadow_object.OffsetX)
            tmp['OffsetY'] = float(shadow_object.OffsetY)
            tmp['Transparency'] = float(shadow_object.Transparency)
        else:
            tmp['Visible'] = 0
        return tmp
    @classmethod
    def _set_shadow(cls, shadow_object, shadow_dict):
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
            tmp['Radius'] = 0
        return tmp
    @classmethod
    def _set_glow(cls, glow_object, glow_dict):
        for key, value in glow_dict.items():
            cls.mult_setattr(glow_object, key, value)

    @classmethod
    def _get_softedge(cls, softedge_object):
        tmp = OrderedDict()
        if softedge_object.Type > 0:
            tmp['Type'] = softedge_object.Type
            tmp['Radius'] = float(softedge_object.Radius)
        else:
            tmp['Type'] = 0
        return tmp
    @classmethod
    def _set_softedge(cls, softedge_object, softedge_dict):
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
        for key, value in textframe_dict.items():
            cls.mult_setattr(textframe_object, key, value)

    @classmethod
    def _get_font(cls, font_object):
        tmp = OrderedDict()
        #Font color
        if font_object.Fill.Visible == -1:
            tmp['Fill.Visible'] = -1
            cls._write_color_to_array(tmp, font_object.Fill.ForeColor, 'Fill.ForeColor')
        else:
            tmp['Fill.Visible'] = 0
        #Font line color
        if font_object.Line.Visible == -1:
            tmp['Line.Visible'] = -1
            cls._write_color_to_array(tmp, font_object.Line.ForeColor, 'Line.ForeColor')
        else:
            tmp['Line.Visible'] = 0
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
        #FIXME: Add Glow, Highlight, Shadow, Reflection...
        return tmp
    @classmethod
    def _set_font(cls, font_object, font_dict):
        for key, value in font_dict.items():
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
            tmp['Bullet.Type'] = parfor_object.Bullet.Type
            tmp['Bullet.Style'] = parfor_object.Bullet.Style
            tmp['Bullet.StartValue'] = parfor_object.Bullet.StartValue
            tmp['Bullet.RelativeSize'] = float(parfor_object.Bullet.RelativeSize)
            tmp['Bullet.Character'] = parfor_object.Bullet.Character
            if parfor_object.Bullet.UseTextFont == -1:
                tmp['Bullet.UseTextFont'] = -1
            else:
                tmp['Bullet.Font.Name'] = parfor_object.Bullet.Font.Name
            if parfor_object.Bullet.UseTextColor == -1:
                tmp['Bullet.UseTextColor'] = -1
            else:
                cls._write_color_to_array(tmp, parfor_object.Bullet.Font.Fill.ForeColor, 'Bullet.Font.Fill.ForeColor')
        else:
            tmp['Bullet.Visible'] = 0
        
        tmp['FirstLineIndent'] = float(parfor_object.FirstLineIndent)
        tmp['LeftIndent'] = float(parfor_object.LeftIndent)
        tmp['RightIndent'] = float(parfor_object.RightIndent)
        tmp['HangingPunctuation'] = parfor_object.HangingPunctuation
        #FIXME: value -2 indicates different values per paragraph, so get values from first paragraph
        return tmp
    @classmethod
    def _set_paragraphformat(cls, parfor_object, parfor_dict):
        for key, value in parfor_dict.items():
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
        shape.Left = position_dict["Left"]
        shape.Top = position_dict["Top"]
        shape.Rotation = position_dict["Rotation"]