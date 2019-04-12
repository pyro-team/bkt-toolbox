# -*- coding: utf-8 -*-
'''
Created on 2018-05-29
@author: Florian Stallmann
'''

import bkt
import bkt.library.powerpoint as pplib

import logging

import os.path
import io
import json

from collections import OrderedDict

D = bkt.dotnet.import_drawing()
# wpf = bkt.dotnet.import_wpf()
# notify_property = bkt.ui.notify_property

# import System
# Window = System.Windows.Window

import helpers


CF_VERSION = "20180824"

class CustomQuickEdit(object):
    custom_styles = 5*[None]
    style_settings = {
        'type':             True,
        'fill':             True,
        'line':             True,
        'textframe2':       True,
        'paragraphformat':  True,
        'font':             True,
        'shadow':           True,
        'size':             True,
        'position':         True,
    }
    always_keep_theme_color = True #set to true to remain theme color even if RGB value differs due to different color scheme
    # use_real_thumbnails = False #set to true to use thumbnails of the actual shape and not generated images for the buttons

    config_folder = os.path.join(bkt.helpers.get_fav_folder(), "custom_formats")
    cache_images = {}

    current_file = "styles.json"
    initialized = False

    #TODO: option to save into presentation (instead of json file)
    #TODO: import/export of style definitions

    @classmethod
    def _initialize(cls):
        if cls.initialized:
            return
        
        cls.always_keep_theme_color = bkt.settings.get("customformats.theme_colors", True)
        cls.default_file = bkt.settings.get("customformats.default_file", "styles.json")
        cls.current_file = cls.default_file
        
        cls.read_from_config(cls.current_file)
        cls.initialized = True

    @classmethod
    def save_to_config(cls):
        # bkt.console.show_message("%r" % cls.custom_styles)
        # bkt.console.show_message(json.dumps(cls.custom_styles))
        file   = os.path.join(cls.config_folder, cls.current_file)
        if not os.path.exists(cls.config_folder):
            os.makedirs(cls.config_folder)
        with io.open(file, 'w') as json_file:
            json.dump(cls.custom_styles, json_file)

    @classmethod
    def read_from_config(cls, filename="styles.json"):
        file   = os.path.join(cls.config_folder, filename)
        if not os.path.isfile(file):
            return
        with io.open(file, 'r') as json_file:
            cls.custom_styles = json.load(json_file, object_pairs_hook=OrderedDict)
            # data = json.load(json_file, object_pairs_hook=OrderedDict)
            # bkt.console.show_message("%r" % data)
        cls.current_file = filename
        cls.clear_cache()

    @classmethod
    def set_default_style(cls, filename=None):
        bkt.settings["customformats.default_file"] = filename or cls.current_file
    
    @classmethod
    def is_default_style(cls, filename=None):
        return cls.default_file == (filename or cls.current_file)

    @classmethod
    def create_new_style(cls, filename=None):
        if not filename:
            filename = bkt.ui.show_user_input("Bitte Dateiname für neuen Style-Katalog eingeben", "Dateiname eingeben", "styles-neu")
            if filename is None:
                return
        if not filename.endswith(".json"):
            filename += ".json"
        file = os.path.join(cls.config_folder, filename)
        if os.path.exists(file):
            bkt.helpers.message("Dateiname existiert bereits")
            return

        cls.custom_styles = 5*[None]
        cls.current_file = filename
        cls.clear_cache()
        cls.save_to_config()

    @classmethod
    def get_styles(cls):
        def style_button(file):
            return bkt.ribbon.ToggleButton(
                label= file,
                # image_mso='DeleteThisFolder',
                get_pressed=bkt.Callback(lambda: file == cls.current_file),
                on_toggle_action=bkt.Callback(lambda pressed: cls.read_from_config(file))
            )

        return bkt.ribbon.Menu(
            xmlns="http://schemas.microsoft.com/office/2009/07/customui",
            id=None,
            children=[
                style_button(file)
                for file in os.listdir(cls.config_folder) if file.endswith(".json")
            ]
        )

    @classmethod
    def get_image_filename_by_index(cls, index, usage="button"):
        file, ext = os.path.splitext(cls.current_file)
        return os.path.join(cls.config_folder, "{}_{}_{}.png".format(file, usage, chr(65+index)) )

    @classmethod
    def get_image_by_index(cls, index, size=16, real_thumb=False):
        cls._initialize()
        cache_key = "{}-{}".format(index, size)
        try:
            return cls.cache_images[cache_key]
        except:
            if real_thumb:
                ### OPTION A: thumbnail of original shape
                file = cls.get_image_filename_by_index(index, "thumb")
            else:
                ## OPTION B: generated thumbnail
                file = cls.get_image_filename_by_index(index)

            if os.path.exists(file):
                #version that should not lock the file, which prevents updating of thumbnails:
                with D.Bitmap.FromFile(file) as img:
                    cls.cache_images[cache_key] = D.Bitmap(img)
                    img.Dispose()
            else:
                # black image
                settings = [0, None, None, "X"]
                cls.cache_images[cache_key] = cls.generate_image(size, *settings)
            
            return cls.cache_images[cache_key]

    @classmethod
    def clear_cache(cls):
        cls.cache_images = {}

    @classmethod
    def generate_button_image(cls, index, shape, size=16):
        file = cls.get_image_filename_by_index(index)
        img = cls.generate_image(size, *cls.custom_styles[index]['button_setting'])
        try:
            img.Save(file)
        except:
            logging.error('Creation of button image failed: %s' % file)
            logging.debug(traceback.format_exc())
        finally:
            img.Dispose()

    @classmethod
    def generate_thumbnail(cls, index, shape, size=64):
        file = cls.get_image_filename_by_index(index, "thumb")
        shape.Export(file, 2) #2=ppShapeFormatPNG, width, height, export-mode: 1=ppRelativeToSlide, 2=ppClipRelativeToSlide, 3=ppScaleToFit, 4=ppScaleXY

        # resize thumbnail image to square
        if os.path.exists(file):
            try:
                # init croped image
                width = size
                height = size
                image = D.Bitmap(file)
                bmp = D.Bitmap(width, height)
                graph = D.Graphics.FromImage(bmp)
                # compute scale
                scale = min(float(width) / image.Width, float(height) / image.Height)
                scaleWidth = int(image.Width * scale)
                scaleHeight = int(image.Height * scale)
                # set quality
                graph.InterpolationMode  = D.Drawing2D.InterpolationMode.High
                graph.CompositingQuality = D.Drawing2D.CompositingQuality.HighQuality
                graph.SmoothingMode      = D.Drawing2D.SmoothingMode.AntiAlias
                # redraw and save
                # logging.debug('crop image from %sx%s to %sx%s. rect %s.%s-%sx%s' % (image.Width, image.Height, width, height, int((width - scaleWidth)/2), int((height - scaleHeight)/2), scaleWidth, scaleHeight))
                graph.DrawImage(image, D.Rectangle(int((width - scaleWidth)/2), int((height - scaleHeight)/2), scaleWidth, scaleHeight))

                # close and save files
                image.Dispose()
                bmp.Save(file)
                bmp.Dispose()
            except:
                logging.error('Creation of croped thumbnail image failed: %s' % file)
                logging.debug(traceback.format_exc())
            finally:
                if image:
                    image.Dispose()
                if bmp:
                    bmp.Dispose()

    @classmethod
    def generate_image(cls, size=16, background=None, line=None, font=None, letter="A"):
        img = D.Bitmap(size, size)
        g = D.Graphics.FromImage(img)
        
        #Draw smooth rectangle/ellipse
        # g.SmoothingMode = D.Drawing2D.SmoothingMode.AntiAlias

        if background is not None:
            background_color = D.ColorTranslator.FromOle(background)
            brush = D.SolidBrush(background_color)
            g.FillRectangle(brush, 0,0,size,size)
        
        if line is not None:
            line_color = D.ColorTranslator.FromOle(line)
            pen = D.Pen(line_color,2)
            g.DrawRectangle(pen, 1,1,size-2,size-2)

        if font is not None:
            font_color = D.ColorTranslator.FromOle(font)
            text_brush = D.SolidBrush(font_color)

            # set string format
            strFormat = D.StringFormat()
            strFormat.Alignment = D.StringAlignment.Center
            strFormat.LineAlignment = D.StringAlignment.Center
            
            # draw string
            # g.TextRenderingHint = D.Text.TextRenderingHint.AntiAliasGridFit
            # g.TextRenderingHint = D.Text.TextRenderingHint.AntiAlias
            g.DrawString(letter,
                         D.Font("Arial", int(size/3*2), D.FontStyle.Bold, D.GraphicsUnit.Pixel), text_brush, 
                         D.RectangleF(0,0,size,size), 
                         strFormat)
            
        return img

    @classmethod
    def show_pickup_window(cls, shape, buttonindex=None):
        from pickup_style import PickupWindow
        PickupWindow.create_and_show_dialog(cls, shape, buttonindex)

    @classmethod
    def pickup_custom_style(cls, index, shape):
        # shift = bkt.library.system.get_key_state(bkt.library.system.key_code.SHIFT)
        # ctrl  = bkt.library.system.get_key_state(bkt.library.system.key_code.CTRL)
        # alt   = bkt.library.system.get_key_state(bkt.library.system.key_code.ALT)

        cls.custom_styles[index] = OrderedDict()
        cls.custom_styles[index]["version"] = CF_VERSION #adding version number in case data structure will change in the future
        cls.custom_styles[index]["button_setting"] = [None, None, None, chr(65+index)] #background (rgb), line (rgb), font (rgb), letter
        cls.custom_styles[index]["style_settings"] = cls.style_settings.copy() #make a copy of style settings (no reference)

        ### BUTTON SETTINGS
        if shape.Fill.Visible == -1:
            cls.custom_styles[index]['button_setting'][0] = shape.Fill.ForeColor.RGB
        if shape.Line.Visible == -1:
            cls.custom_styles[index]['button_setting'][1] = shape.Line.ForeColor.RGB
        if shape.HasTextFrame == -1:
            textrange = shape.TextFrame2.TextRange
            try:
                font_fill = textrange.Characters(1).Font.Fill
            except:
                font_fill = textrange.Font.Fill
            if font_fill.Visible == -1:
                cls.custom_styles[index]['button_setting'][2] = textrange.Font.Fill.ForeColor.RGB

        ### TYPE
        cls.custom_styles[index]['Type'] = helpers.ShapeFormats._get_type(shape)
        
        ### BACKGROUND
        cls.custom_styles[index]['Fill'] = helpers.ShapeFormats._get_fill(shape.Fill)

        ### LINE
        cls.custom_styles[index]['Line'] = helpers.ShapeFormats._get_line(shape.Line)

        ### TEXTFRAME
        if shape.HasTextFrame == -1:
            cls.custom_styles[index]['TextFrame2'] = helpers.ShapeFormats._get_textframe(shape.TextFrame2)

        ### INDENT LEVEL SPECIFIC FORMATS (PARAGRAPH, FONT)
        if shape.HasTextFrame == -1:
            cls.custom_styles[index]["IndentLevels"] = OrderedDict()
            if shape.TextFrame2.TextRange.Paragraphs().Count > 0:
                # at least one paragraph
                indent_levels = range(0,6) #indent levels 0 to 5, whereas 0 is used as internal fallback format!
                for par in shape.TextFrame2.TextRange.Paragraphs():
                    indent_level = par.ParagraphFormat.IndentLevel
                    if indent_level == 1 and par.ParagraphFormat.Bullet.Visible == 0:
                        indent_level = 0 #fallback indent level

                    if indent_level in indent_levels:
                        indent_levels.remove(indent_level)
                        cls.custom_styles[index]["IndentLevels"][str(indent_level)] = helpers.ShapeFormats._get_indentlevel_formats(par)
                if 0 in indent_levels:
                    #fallback not yet defined
                    cls.custom_styles[index]["IndentLevels"]["0"] = helpers.ShapeFormats._get_indentlevel_formats(shape.TextFrame2.TextRange.Paragraphs(1,1))
            else:
                cls.custom_styles[index]["IndentLevels"]["0"] = helpers.ShapeFormats._get_indentlevel_formats(shape.TextFrame2.TextRange)


        ### SHADOW
        cls.custom_styles[index]['Shadow'] = helpers.ShapeFormats._get_shadow(shape.Shadow)
        cls.custom_styles[index]['Glow'] = helpers.ShapeFormats._get_glow(shape.Glow)
        cls.custom_styles[index]['SoftEdge'] = helpers.ShapeFormats._get_softedge(shape.SoftEdge)
        cls.custom_styles[index]['Reflection'] = helpers.ShapeFormats._get_reflection(shape.Reflection)

        #FIXME: Add: ThreeD, AnimationSettings

        ### SIZE
        cls.custom_styles[index]['Size'] = helpers.ShapeFormats._get_size(shape)
        
        ### POSITION
        cls.custom_styles[index]['Position'] = helpers.ShapeFormats._get_position(shape)

        # save to file
        cls.save_to_config()
        cls.clear_cache()

        #generate thumbnails
        cls.generate_thumbnail(index, shape)
        cls.generate_button_image(index, shape)


    @classmethod
    def _convert_style_version(cls, index):
        if cls.custom_styles[index]["version"] == CF_VERSION:
            return
        else:
            raise ValueError("Unable to convert style")

    @classmethod
    def _create_shape(cls, index, context):
        style = cls.custom_styles[index]
        left = (context.presentation.PageSetup.SlideWidth-50)*0.5
        top  = (context.presentation.PageSetup.SlideHeight-50)*0.5
        shp  = context.slides[0].Shapes.AddShape(1, left, top, 50, 50)
        shp.select()

        #for new shapes always consider type, size and position
        # cls.apply_custom_style(index, context, {'type': True, 'size': True, 'position': True}, [shp])
        cls._apply_custom_style_on_shape(shp, cls.custom_styles[index], {'type': True, 'size': True, 'position': True})

        return shp

    @classmethod
    def new_shape_custom_style(cls, index, context):
        if cls.custom_styles[index] is None:
            bkt.helpers.message("Style nicht definiert!")
            return
        
        try:
            cls._convert_style_version(index)
        except ValueError:
            bkt.helpers.message("Veraltete Style-Definition! Konvertierung fehlgeschlagen. Bitte Style neu anlegen.")
            return
        
        cls._create_shape(index, context)
        cls.apply_custom_style(index, context)

    @classmethod
    def edit_custom_style(cls, index, context):
        if cls.custom_styles[index] is None:
            bkt.helpers.message("Style nicht definiert!")
            return
        
        try:
            cls._convert_style_version(index)
        except ValueError:
            bkt.helpers.message("Veraltete Style-Definition! Konvertierung fehlgeschlagen. Bitte Style neu anlegen.")
            return

        from apply_style import ApplyWindow
        ApplyWindow.create_and_show_dialog(cls, index, context)

    @classmethod
    def apply_custom_style(cls, index, context, style_settings=None):
        if cls.custom_styles[index] is None:
            bkt.helpers.message("Style nicht definiert!")
            return
        
        try:
            cls._convert_style_version(index)
        except ValueError:
            bkt.helpers.message("Veraltete Style-Definition! Konvertierung fehlgeschlagen. Bitte Style neu anlegen.")
            return

        settings = style_settings or cls.custom_styles[index]["style_settings"]
        
        shift = bkt.library.system.get_key_state(bkt.library.system.key_code.SHIFT)
        # ctrl  = bkt.library.system.get_key_state(bkt.library.system.key_code.CTRL)
        alt   = bkt.library.system.get_key_state(bkt.library.system.key_code.ALT)

        if alt:
            cls.edit_custom_style(index, context)

        # elif shift:
        #     #select shapes with style
        #     shapes = list(iter(context.slides[0].Shapes))
        #     # settings = cls.custom_styles[index]['button_setting']
        #     style = cls.custom_styles[index]
        #     for shape in shapes:
        #         try:
        #             select = [False,False,True]
        #             if not settings["background"] or (shape.Fill.Visible == style["Fill"]["Visible"] and (style["Fill"]["Visible"] == 0 or shape.Fill.ForeColor.RGB == style['button_setting'][0])):
        #                 select[0] = True
        #             if not settings["line"] or (shape.Line.Visible == style["Line"]["Visible"] and (style["Line"]["Visible"] == 0 or shape.Line.ForeColor.RGB == style['button_setting'][1])):
        #                 select[1] = True
        #             # if not settings["font"] or shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB == settings[2]:
        #             #     select[2] = True
        #             if all(select):
        #                 shape.Select(replace=False)
        #         except:
        #             # bkt.helpers.exception_as_message()
        #             continue
        
        else:
            #appyly style

            if shift or context.selection.Type not in [2,3]:
                #create new shape with this style
                shp = cls._create_shape(index, context)
                shapes = [shp]
            else:
                # shapes = pplib.get_shapes_from_selection(selection)
                shapes = context.shapes

            helpers.ShapeFormats.always_keep_theme_color = cls.always_keep_theme_color
            style = cls.custom_styles[index]

            errors = False
            for shape in shapes:
                errors = errors or cls._apply_custom_style_on_shape(shape, style, settings)

            if errors:
                bkt.helpers.message("Einige Eigenschaften konnten nicht auf die ausgewählten Shapes übertragen werden. Dies kann beispielweise bei Tabellen passieren, da diese nicht alle Eigenschaften unterstützen.")

    @classmethod
    def _apply_custom_style_on_shape(cls, shape, style, settings):
        errors = False
        
        try:
            if settings.get("type", False) and "Type" in style:
                helpers.ShapeFormats._set_type(shape, style["Type"])
        except Exception as e:
            errors = True
            logging.error("Custom formats: Error in setting shape type with error: {}".format(e))

        try:
            if settings.get("fill", False) and "Fill" in style:
                helpers.ShapeFormats._set_fill(shape.fill, style["Fill"])
        except Exception as e:
            errors = True
            logging.error("Custom formats: Error in setting fill with error: {}".format(e))

        try:
            if settings.get("line", False) and "Line" in style:
                helpers.ShapeFormats._set_line(shape.line, style["Line"])
        except Exception as e:
            errors = True
            logging.error("Custom formats: Error in setting line with error: {}".format(e))

        try:
            if settings.get("textframe2", False) and "TextFrame2" in style:
                helpers.ShapeFormats._set_textframe(shape.textframe2, style["TextFrame2"])
        except Exception as e:
            errors = True
            logging.error("Custom formats: Error in setting textframe with error: {}".format(e))

        try:
            if settings.get("shadow", False):
                # order is important here. shadow must be last as setting glow, reflection or softedge will re-enable shadow
                if "Glow" in style:
                    helpers.ShapeFormats._set_glow(shape.glow, style["Glow"])
                if "Reflection" in style:
                    helpers.ShapeFormats._set_reflection(shape.reflection, style["Reflection"])
                if "SoftEdge" in style:
                    helpers.ShapeFormats._set_softedge(shape.softedge, style["SoftEdge"])
                if "Shadow" in style:
                    helpers.ShapeFormats._set_shadow(shape.shadow, style["Shadow"])
        except Exception as e:
            errors = True
            logging.error("Custom formats: Error in setting shadow with error: {}".format(e))

        try:
            if settings.get("size", False) and "Size" in style:
                helpers.ShapeFormats._set_size(shape, style["Size"])
        except Exception as e:
            errors = True
            logging.error("Custom formats: Error in setting shape size with error: {}".format(e))

        try:
            if settings.get("position", False) and "Position" in style:
                helpers.ShapeFormats._set_position(shape, style["Position"])
        except Exception as e:
            errors = True
            logging.error("Custom formats: Error in setting shape position with error: {}".format(e))

        try:
            if (settings.get("paragraphformat", False) or settings.get("font", False)) and "IndentLevels" in style:
                if shape.TextFrame2.TextRange.Paragraphs().Count > 0:
                    for par in shape.TextFrame2.TextRange.Paragraphs():
                        indent_level = str(par.ParagraphFormat.IndentLevel)
                        if indent_level not in style["IndentLevels"] or (indent_level == "1" and par.ParagraphFormat.Bullet.Visible == 0):
                            indent_level = "0"
                        if settings.get("paragraphformat", False):
                            helpers.ShapeFormats._set_paragraphformat(par.ParagraphFormat, style["IndentLevels"][indent_level]["ParagraphFormat"])
                        if settings.get("font", False):
                            helpers.ShapeFormats._set_font(par.Font, style["IndentLevels"][indent_level]["Font"])
                else:
                    if settings.get("paragraphformat", False):
                        helpers.ShapeFormats._set_paragraphformat(shape.TextFrame2.TextRange.ParagraphFormat, style["IndentLevels"]["0"]["ParagraphFormat"])
                    if settings.get("font", False):
                        helpers.ShapeFormats._set_font(shape.TextFrame2.TextRange.Font, style["IndentLevels"]["0"]["Font"])
        except Exception as e:
            errors = True
            logging.error("Custom formats: Error in setting paragraph format with error: {}".format(e))
        
        return errors



    @classmethod
    def get_supertip(cls, index):
        cls._initialize()

        default = "Style auf aktuelle Auswahl anwenden.{}\n\nMit SHIFT-Taste: Neues Shape im gewählten Format anlegen."
        if cls.custom_styles[index] is None or "style_settings" not in cls.custom_styles[index]:
            return default.format("")
        
        styles = "\n" + "\n".join( ["{}: {}".format(k.capitalize(), "ja" if v else "nein") for k,v in cls.custom_styles[index]["style_settings"].iteritems()] )
        return default.format(styles)


def qe_button(i):
    return bkt.ribbon.SplitButton(
                id="quickedit_custom_apply_%s" % (i+1),
                show_label=False,
                children=[
                    bkt.ribbon.Button(
                        label="Format/Style %s" % (i+1),
                        # supertip="Style auf aktuelle Auswahl anwenden.\n\nMit SHIFT-Taste: Shapes mit Style auswählen.",
                        get_supertip=bkt.Callback(lambda: CustomQuickEdit.get_supertip(i)),
                        get_image=bkt.Callback(lambda: CustomQuickEdit.get_image_by_index(i)),
                        on_action=bkt.Callback(lambda context: CustomQuickEdit.apply_custom_style(i, context), context=True),
                    ),
                    bkt.ribbon.Menu(
                        label="Format/Style %s" % (i+1),
                        children=[
                            bkt.ribbon.Button(
                                label="Style auf Shape(s) anwenden",
                                screentip="Style auf ausgewählte(s) Shape(s) anwenden",
                                get_image=bkt.Callback(lambda: CustomQuickEdit.get_image_by_index(i)),
                                on_action=bkt.Callback(lambda context: CustomQuickEdit.apply_custom_style(i, context), context=True),
                            ),
                            bkt.ribbon.Button(
                                label="Neues Shape anlegen [SHIFT]",
                                screentip="Neues Shape mit Style anlegen",
                                on_action=bkt.Callback(lambda context: CustomQuickEdit.new_shape_custom_style(i, context), context=True),
                            ),
                            bkt.ribbon.MenuSeparator(),
                            bkt.ribbon.Button(
                                label="Nur Hintergrund anwenden",
                                on_action=bkt.Callback(lambda context: CustomQuickEdit.apply_custom_style(i, context, {'fill': True}), context=True),
                            ),
                            bkt.ribbon.Button(
                                label="Nur Rahmen anwenden",
                                on_action=bkt.Callback(lambda context: CustomQuickEdit.apply_custom_style(i, context, {'line': True}), context=True),
                            ),
                            bkt.ribbon.MenuSeparator(),
                            bkt.ribbon.Button(
                                label="Style-Formate auswählen [ALT]",
                                image_mso = 'ShowCustomPropertiesPage',
                                screentip="Auswählen der anzuwendenden Style-Formate",
                                on_action=bkt.Callback(lambda context: CustomQuickEdit.edit_custom_style(i, context), context=True),
                            ),
                            # bkt.ribbon.Button(
                            #     label="Style überschreiben",
                            #     image_mso="PickUpStyle",
                            #     screentip="Aktuelles Shape als Style aufnehmen",
                            #     on_action=bkt.Callback(lambda shape: CustomQuickEdit.show_pickup_window(shape, i), shape=True),
                            #     get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                            # ),
                        ]
                    )
                ]
            )


class FormatLibGallery(bkt.ribbon.Gallery):
    
    def __init__(self, **kwargs):
        parent_id = kwargs.get('id') or ""
        my_kwargs = dict(
            label = 'Styles',
            columns = 5,
            # image = 'shapetable',
            image_mso = 'SlidesPerPage4Slides',
            show_item_label=False,
            screentip="Custom-Styles Gallerie",
            # supertip="Füge eine Tabelle aus Standard-Shapes ein",
            children=[
                bkt.ribbon.Button(id=parent_id + "_pickup", label="Neuen Style aufnehmen", image_mso="PickUpStyle", on_action=bkt.Callback(CustomQuickEdit.show_pickup_window, shape=True), get_enabled = bkt.apps.ppt_shapes_exactly1_selected,),
            ]
        )
        my_kwargs.update(kwargs)

        super(FormatLibGallery, self).__init__(**my_kwargs)

    def on_action_indexed(self, selected_item, index, context):
        CustomQuickEdit.apply_custom_style(index, context)
    
    def get_item_count(self):
        return len(CustomQuickEdit.custom_styles)
        
    def get_item_label(self, index):
        return "Style {}".format(index+1)
    
    def get_item_screentip(self, index):
        return "Style {} anwenden".format(index+1)
        
    def get_item_supertip(self, index):
        return CustomQuickEdit.get_supertip(index)
    
    def get_item_image(self, index):
        return CustomQuickEdit.get_image_by_index(index, size=32, real_thumb=True)


customformats_group = bkt.ribbon.Group(
    id="bkt_customformats_group",
    label='Styles',
    image_mso='SmartArtChangeColorsGallery',
    children = [
            qe_button(i)
            for i in range(0, len(CustomQuickEdit.custom_styles))
    ] + [
        bkt.ribbon.Menu(
            id="quickedit_config_menu",
            label="Custom Styles Konfiguration",
            show_label=False,
            image_mso="PickUpStyle",
            children=[
                bkt.ribbon.Button(
                    id="quickedit_custom_define",
                    label="Aktuelles Shape als Style aufnehmen",
                    # show_label=False,
                    image_mso="PickUpStyle",
                    supertip="Style (Hintergrund, Linie, Text, Schatten) des ausgewählten Shapes speichern.",
                    on_action=bkt.Callback(CustomQuickEdit.show_pickup_window, shape=True),
                    get_enabled = bkt.apps.ppt_shapes_exactly1_selected,
                ),
                bkt.ribbon.MenuSeparator(title="Style-Kataloge verwalten"),
                bkt.ribbon.DynamicMenu(
                    label='Style-Katalog ändern',
                    # image_mso='ModuleInsert',
                    get_content = bkt.Callback(CustomQuickEdit.get_styles)
                ),
                bkt.ribbon.Button(
                    label='Neuen Style-Katalog anlegen',
                    # image_mso='ModuleInsert',
                    on_action=bkt.Callback(CustomQuickEdit.create_new_style)
                ),
                bkt.ribbon.Button(
                    label='Aktuellen Katalog als Standard',
                    # image_mso='ModuleInsert',
                    on_action=bkt.Callback(CustomQuickEdit.set_default_style),
                    get_enabled=bkt.Callback(lambda: not CustomQuickEdit.is_default_style())
                ),
            ]
        ),
        FormatLibGallery(id="customformats_gallery", size="large")
    ]
)


bkt.powerpoint.add_tab(bkt.ribbon.Tab(
    id="bkt_powerpoint_toolbox_extensions",
    #id_q="nsBKT:powerpoint_toolbox_extensions",
    #insert_after_q="nsBKT:powerpoint_toolbox_advanced",
    insert_before_mso="TabHome",
    label=u'Toolbox 3/3',
    # get_visible defaults to False during async-startup
    get_visible=bkt.Callback(lambda:True),
    children = [
        customformats_group,
    ]
), extend=True)
