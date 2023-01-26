# -*- coding: utf-8 -*-
'''
Created on 06.07.2016

@author: rdebeerst
'''


from threading import Thread

import bkt
import bkt.library.powerpoint as pplib


# class Calendar(object):
#     @classmethod
#     def get_calendar(cls):
#         pass
        # import calendar
        # from itertools import count, cycle

        # daynames_iter   = cycle(calendar.day_name)
        # monthnames_iter = cycle(calendar.month_name[1:])

        # dayabbrs_iter   = cycle(calendar.day_abbr)
        # monthabbrs_iter = cycle(calendar.month_abbr[1:])

        # "day_names":    daynames_iter.next(),
        # "month_names":  monthnames_iter.next(),
        # "day_abbrs":    dayabbrs_iter.next(),
        # "month_abbrs":  monthabbrs_iter.next(),

        # with pplib.override_locale(presentation.DefaultLanguageID):
        # with pplib.override_locale(first_textframe.TextRange.LanguageID):

    #create 1 month view like outlook
    #create week table
    #create 1 month mini calendar
    #create month-> calendar week rows
    #datumsfeld? (WPF Calendar, Datepicker)


class TextPlaceholder(object):
    recent_placeholder = bkt.settings.get("toolbox.recent_placeholder", "…")
    #labels for counter, but max 0..25
    # label_a = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z']
    # label_A = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
    # label_I = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI', 'XII', 'XIII', 'XIV', 'XV', 'XVI', 'XVII', 'XVIII', 'XIX', 'XX', 'XXI', 'XXII', 'XXIII', 'XXIV', 'XXV', 'XXVI']

    @staticmethod
    def set_text_for_shape(textframe, text=None): #None=delete text
        if text is not None:
            if isinstance(text, list):
                if text[0] != '':
                    textframe.TextRange.InsertBefore( text[0] )
                if text[1] != '':
                    textframe.TextRange.InsertAfter( text[1] )
            else:
                textframe.TextRange.Text = text
        else:
            textframe.TextRange.Delete()

    @classmethod
    def text_truncate(cls, shapes):
        for textframe in pplib.iterate_shape_textframes(shapes):
            cls.set_text_for_shape(textframe, None)

    @classmethod
    def text_replace(cls, shapes, presentation):
        from string import Template
        from formatter import AbstractFormatter, DumbWriter

        input_text = bkt.ui.show_user_input('''Text eingeben, der auf alle Shapes angewendet werden soll. Es stehen folgende Platzhalter zur Verfügung:

${text}:\t\tSetzt bestehenden Shape-Text ein
$text_len:\tAnzahl Zeichen im bestehenden Text
$counter:\tNummerierung (Fortsetzung wenn erstes Shape eine Zahl enthält)
$counter_a/A:\tNummerierung mit Klein- bzw. Großbuchstaben
$counter_i/I:\tNummerierung mit römischen Ziffern (bspw. xi bzw. XI)''', "Text ersetzen", cls.recent_placeholder, True)

        if input_text is None:
            return
        cls.recent_placeholder = bkt.settings["toolbox.recent_placeholder"] = input_text

        #get very first textframe to define language and counter start
        first_textframe = next(iter(pplib.iterate_shape_textframes(shapes)))

        #if first shape has a number, use it as counter start
        try:
            count_start = int(first_textframe.TextRange.Text)-1
        except (ValueError, TypeError):
            count_start = 0

        template = Template(input_text)
        count_formatter = AbstractFormatter(DumbWriter())

        for i,textframe in enumerate(pplib.iterate_shape_textframes(shapes), start=1):
            placeholders = {
                "counter":   count_formatter.format_counter('1', i+count_start),
                "counter_a": count_formatter.format_counter('a', i),
                "counter_A": count_formatter.format_counter('A', i),
                "counter_i": count_formatter.format_counter('i', i),
                "counter_I": count_formatter.format_counter('I', i),

                "text":      "[$text]",
                "text_len":  len(textframe.TextRange.Text),
            }

            # run template engine
            new_text = template.safe_substitute(placeholders)
            
            # check if $text placeholder was present
            placeholder_count = new_text.count("[$text]")
            if placeholder_count > 1:
                #replace placeholder with text, might loose existing formatting
                new_text = new_text.replace("[$text]", textframe.TextRange.Text)
            elif placeholder_count == 1:
                #only one occurence of text-placeholder, make use of InsertBefore/After to keep formatting
                new_text = new_text.split("[$text]", 1)
            
            cls.set_text_for_shape(textframe, new_text)

    @classmethod
    def text_tbd(cls, shapes):
        for textframe in pplib.iterate_shape_textframes(shapes):
            cls.set_text_for_shape(textframe, "tbd")

    @classmethod
    def text_counter(cls, shapes):
        for counter,textframe in enumerate(pplib.iterate_shape_textframes(shapes), start=1):
            cls.set_text_for_shape(textframe, str(counter))

    @classmethod
    def text_lorem(cls, shapes):
        lorem_text = '''Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat, sed diam voluptua.
At vero eos et accusam et justo duo dolores et ea rebum. Stet clita kasd gubergren, no sea takimata sanctus est Lorem ipsum dolor sit amet.
Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat, sed diam voluptua.
At vero eos et accusam et justo duo dolores et ea rebum. Stet clita kasd gubergren, no sea takimata sanctus est Lorem ipsum dolor sit amet.'''
        lorem_text = bkt.helpers.endings_to_windows(lorem_text)

        for textframe in pplib.iterate_shape_textframes(shapes):
            cls.set_text_for_shape(textframe, lorem_text)

    @staticmethod
    def remove_placeholders(slides):
        for sld in slides:
            for plh in list(iter(sld.Shapes.Placeholders)):
                if plh.HasTextFrame == -1 and plh.TextFrame2.HasText == 0:
                    #placeholder is a text placeholder and has no text. note: placeholder can also be a picture, table or diagram without text!
                    plh.Delete()


class BulletStyle(object):
    
    @staticmethod
    def set_bullet_color_rgb(selection, shapes, color):
        if selection.Type == 3:
            # text selected
            selection.TextRange2.ParagraphFormat.Bullet.Font.Fill.ForeColor.RGB = color

        else:
            for textframe in pplib.iterate_shape_textframes(shapes):
                textframe.TextRange.ParagraphFormat.Bullet.Font.Fill.ForeColor.RGB = color

        # for shape in shapes:
        #     shape.TextFrame2.TextRange.ParagraphFormat.Bullet.Font.Fill.ForeColor.RGB = color

    @staticmethod
    def set_bullet_theme_color(selection, shapes, color_index, brightness):
        if selection.Type == 3:
            # text selected
            selection.TextRange2.ParagraphFormat.Bullet.Font.Fill.ForeColor.ObjectThemeColor = color_index
            selection.TextRange2.ParagraphFormat.Bullet.Font.Fill.ForeColor.Brightness = brightness

        else:
            for textframe in pplib.iterate_shape_textframes(shapes):
                textframe.TextRange.ParagraphFormat.Bullet.Font.Fill.ForeColor.ObjectThemeColor = color_index
                textframe.TextRange.ParagraphFormat.Bullet.Font.Fill.ForeColor.Brightness = brightness

        # for shape in shapes:
        #     shape.TextFrame2.TextRange.ParagraphFormat.Bullet.Font.Fill.ForeColor.ObjectThemeColor = color_index
        #     shape.TextFrame2.TextRange.ParagraphFormat.Bullet.Font.Fill.ForeColor.Brightness = brightness

    @staticmethod
    def set_bullet_color_auto(selection, shapes):
        if selection.Type == 3:
            # text selected
            selection.TextRange2.ParagraphFormat.Bullet.UseTextColor = -1

        else:
            for textframe in pplib.iterate_shape_textframes(shapes):
                textframe.TextRange.ParagraphFormat.Bullet.UseTextColor = -1

    @staticmethod
    def set_bullet_symbol(selection, shapes, symbol):
        def _set_bullet(par):
            par.Bullet.Character = ord(symbol[1])
            if symbol[0]:
                par.Bullet.Font.Name = symbol[0]
            else:
                par.Bullet.UseTextFont = -1


        if selection.Type == 3:
            # text selected
            _set_bullet(selection.TextRange2.ParagraphFormat)

        else:
            for textframe in pplib.iterate_shape_textframes(shapes):
                _set_bullet(textframe.TextRange.ParagraphFormat)




    @classmethod
    def get_bullet_color_rgb(cls, selection, shapes):
        return cls._get_from_par_format(selection, shapes, cls._get_bullet_color_from_par)

    @classmethod
    def get_bullet_symbol(cls, selection, shapes):
        return cls._get_from_par_format(selection, shapes, cls._get_bullet_symbol_from_par)


    @classmethod
    def _get_from_par_format(cls, selection, shapes, getter_method):
        if selection.Type == 3:
            # text selected
            try:
                # produces error if no text is selected
                return getter_method(selection.TextRange2.Paragraphs(1,1).ParagraphFormat)
            except ValueError: #ValueError: Der Index in der angegebenen Sammlung ist außerhalb des zulässigen Bereichs.
                return getter_method(selection.TextRange2.ParagraphFormat)
        
        else:
            # shapes selected
            for textframe in pplib.iterate_shape_textframes(shapes):
                return getter_method(textframe.TextRange.ParagraphFormat)


    @classmethod
    def _get_bullet_color_from_par(cls, par_format):
        return [par_format.Bullet.Font.Fill.ForeColor.ObjectThemeColor, par_format.Bullet.Font.Fill.ForeColor.Brightness, par_format.Bullet.Font.Fill.ForeColor.RGB]

    @classmethod
    def _get_bullet_symbol_from_par(cls, par_format):
        if par_format.Bullet.Visible:
            return chr(par_format.Bullet.Character)
        return None

    
    @classmethod
    def fix_bullet_style(cls, shapes):
        shape = shapes[0]
        slide = shape.Parent
        placeholders = [shape for shape in slide.Master.Shapes if shape.Type == 14 and shape.PlaceholderFormat.Type == 2]
        ref_shape = placeholders[0]
        cls.fix_bullet_style_by_reference(shapes, ref_shape)
    
    
    @staticmethod
    def fix_bullet_style_by_reference(shapes, ref_shape):
        # shape = shapes[0]
        # slide = shape.Parent
        # placeholders = [shape for shape in slide.Master.Shapes if shape.Type == 14 and shape.PlaceholderFormat.Type == 2]
        # textph = placeholders[0]
        ph_paragraphs = [p for p in ref_shape.TextFrame2.TextRange.Paragraphs() ]
        ph_paragraphs = [[p for p in ph_paragraphs if p.ParagraphFormat.IndentLevel == indent_level] for indent_level in range(1,6) ] #IndentLevel is between 1 and 5
        ph_paragraphs = [ None if len(ph_list) == 0 else ph_list[0] for ph_list in ph_paragraphs ]
        
        ph_formats = [None if p==None else p.ParagraphFormat for p in ph_paragraphs ]
        
        # for shape in shapes:
            # for par in shape.TextFrame2.TextRange.Paragraphs() :
        for textframe in pplib.iterate_shape_textframes(shapes):
            for par in textframe.TextRange.Paragraphs() :
                indent_level = par.ParagraphFormat.IndentLevel
            
                if ph_paragraphs[indent_level] is None:
                    continue

                par.ParagraphFormat.Bullet.Style = ph_formats[indent_level].Bullet.Style
                par.ParagraphFormat.Bullet.Type = ph_formats[indent_level].Bullet.Type
                par.ParagraphFormat.Bullet.RelativeSize = ph_formats[indent_level].Bullet.RelativeSize
                par.ParagraphFormat.Bullet.Character = ph_formats[indent_level].Bullet.Character

                if ph_formats[indent_level].Bullet.UseTextFont == -1:
                    par.ParagraphFormat.Bullet.UseTextFont = -1
                else:
                    par.ParagraphFormat.Bullet.Font.Name = ph_formats[indent_level].Bullet.Font.Name
                
                if ph_formats[indent_level].Bullet.UseTextColor == -1:
                    par.ParagraphFormat.Bullet.UseTextColor = -1
                else:
                    if ph_formats[indent_level].Bullet.Font.Fill.ForeColor.ObjectThemeColor == 0:
                        par.ParagraphFormat.Bullet.Font.Fill.ForeColor.RGB = ph_formats[indent_level].Bullet.Font.Fill.ForeColor.RGB
                    else:
                        par.ParagraphFormat.Bullet.Font.Fill.ForeColor.ObjectThemeColor = ph_formats[indent_level].Bullet.Font.Fill.ForeColor.ObjectThemeColor
                        par.ParagraphFormat.Bullet.Font.Fill.ForeColor.Brightness = ph_formats[indent_level].Bullet.Font.Fill.ForeColor.Brightness
            
                # par.ParagraphFormat.Bullet.UseTextColor = ph_formats[indent_level].Bullet.UseTextColor
                # par.ParagraphFormat.Bullet.UseTextFont = ph_formats[indent_level].Bullet.UseTextFont

                par.ParagraphFormat.Bullet.Visible = ph_formats[indent_level].Bullet.Visible
            
                par.ParagraphFormat.LeftIndent = ph_formats[indent_level].LeftIndent
                par.ParagraphFormat.RightIndent = ph_formats[indent_level].RightIndent
                par.ParagraphFormat.FirstLineIndent = ph_formats[indent_level].FirstLineIndent
                par.ParagraphFormat.HangingPunctuation = ph_formats[indent_level].HangingPunctuation
            
                par.ParagraphFormat.BaselineAlignment = ph_formats[indent_level].BaselineAlignment
                par.ParagraphFormat.LineRuleBefore    = ph_formats[indent_level].LineRuleBefore
                par.ParagraphFormat.SpaceBefore       = ph_formats[indent_level].SpaceBefore
                par.ParagraphFormat.LineRuleAfter     = ph_formats[indent_level].LineRuleAfter
                par.ParagraphFormat.SpaceAfter        = ph_formats[indent_level].SpaceAfter
                par.ParagraphFormat.LineRuleWithin    = ph_formats[indent_level].LineRuleWithin
                par.ParagraphFormat.SpaceWithin       = ph_formats[indent_level].SpaceWithin
            
                # Text format
                par.Font.Name   = ph_paragraphs[indent_level].Font.Name
                par.Font.Size   = ph_paragraphs[indent_level].Font.Size
                par.Font.Bold   = ph_paragraphs[indent_level].Font.Bold
                par.Font.Italic = ph_paragraphs[indent_level].Font.Italic
                par.Font.Caps   = ph_paragraphs[indent_level].Font.Caps
                if ph_paragraphs[indent_level].Font.Fill.ForeColor.ObjectThemeColor == 0:
                    par.Font.Fill.ForeColor.RGB = ph_paragraphs[indent_level].Font.Fill.ForeColor.RGB
                else:
                    par.Font.Fill.ForeColor.ObjectThemeColor = ph_paragraphs[indent_level].Font.Fill.ForeColor.ObjectThemeColor
                    par.Font.Fill.ForeColor.Brightness = ph_paragraphs[indent_level].Font.Fill.ForeColor.Brightness


class TextOnShape(object):

    @classmethod
    def find_shape_on_shape(cls, master_shape, shapes):
        if master_shape.HasTextFrame == 0:
            return None
        for s in shapes:
            #shape on top of master shape and shape midpoint within master shape
            if s != master_shape and s.HasTextFrame == -1 and s.ZOrderPosition > master_shape.ZOrderPosition and s.left+s.width/2 >= master_shape.left and s.left+s.width/2 <= master_shape.left+master_shape.width and s.top+s.height/2 >= master_shape.top and s.top+s.height/2 <= master_shape.top+master_shape.height:
                return s
        return None
    
    @classmethod
    def merge_shapes(cls, master_shape, text_shape):
        # Kein TextFrame, bspw. bei Linien
        if master_shape.HasTextFrame == 0 or text_shape.HasTextFrame == 0:
            return

        # Text kopieren
        # text_shape.TextFrame2.TextRange.Copy()
        # master_shape.TextFrame2.TextRange.Paste()
        pplib.transfer_textrange(text_shape.TextFrame2.TextRange, master_shape.TextFrame2.TextRange)
        # Textbox loeschen
        text_shape.Delete()

    @classmethod
    def textIntoShape(cls, shapes):
        def loop(worker=None):
            if len(shapes) == 2 and shapes[0].Type == pplib.MsoShapeType['msoTextBox']:
                cls.merge_shapes(shapes[1], shapes[0])
            elif len(shapes) == 2 and shapes[1].Type == pplib.MsoShapeType['msoTextBox']:
                cls.merge_shapes(shapes[0], shapes[1])
            else:
                sorted_shapes = sorted(shapes, key=lambda s: s.ZOrderPosition) #important due to removal of items in for loops
                for shape in sorted_shapes:
                    inner_shp = cls.find_shape_on_shape(shape, sorted_shapes)
                    if inner_shp is not None:
                        cls.merge_shapes(shape, inner_shp)
                        sorted_shapes.remove(inner_shp)

        # IMPORTANT: many copy-paste operations lead to EnvironmentError. Putting all in a thread solves this issue.
        t = Thread(target=loop)
        t.start()
        t.join()

        # Alternative way:
        # bkt.ui.execute_with_progress_bar(loop, indeterminate=True)

    @staticmethod
    def create_txt_shape_onto_shape(shp, slide):
        shpTxt = slide.shapes.AddTextbox(
            1, #msoTextOrientationHorizontal
            shp.Left, shp.Top, shp.Width, shp.Height)
        # WordWrap / AutoSize
        # shpTxt.TextFrame2.WordWrap = -1 #msoTrue
        shpTxt.TextFrame2.WordWrap = shp.TextFrame2.WordWrap
        shpTxt.TextFrame2.AutoSize = 0 #ppAutoSizeNone
        shpTxt.Height   = shp.Height
        shpTxt.Rotation = shp.Rotation
        shpTxt.Name     = shp.Name + " Text"
        # Seitenraender
        shpTxt.TextFrame2.MarginBottom = shp.TextFrame2.MarginBottom
        shpTxt.TextFrame2.MarginTop    = shp.TextFrame2.MarginTop
        shpTxt.TextFrame2.MarginLeft   = shp.TextFrame2.MarginLeft
        shpTxt.TextFrame2.MarginRight  = shp.TextFrame2.MarginRight
        # Ausrichtung
        shpTxt.TextFrame2.Orientation      = shp.TextFrame2.Orientation
        shpTxt.TextFrame2.HorizontalAnchor = shp.TextFrame2.HorizontalAnchor
        shpTxt.TextFrame2.VerticalAnchor   = shp.TextFrame2.VerticalAnchor
        # Text kopieren
        pplib.transfer_textrange(shp.TextFrame2.TextRange, shpTxt.TextFrame2.TextRange)
        # shp.TextFrame2.TextRange.Copy()
        # shpTxt.TextFrame2.TextRange.Paste()
        shp.TextFrame2.DeleteText()
        # Größe wiederherstellen
        shp.Top = shpTxt.Top
        shp.Height = shpTxt.Height
        shp.Width = shpTxt.Width
        # Textfeld selektieren
        shpTxt.Select(0)

    @classmethod
    def textOutOfShape(cls, shapes, slide):
        def loop(worker=None):
            errors = 0
            # shapes_len = len(shapes)
            # i = 1.
            for shp in shapes:
                # worker.ReportProgress(i/shapes_len*100)
                # i += 1

                #if shp.TextFrame.TextRange.text != "":
                if not shp.HasTextFrame or not shp.TextFrame.HasText:
                    continue
                try:
                    # bkt.Clipboard.clear()

                    cls.create_txt_shape_onto_shape(shp, slide)
                    # cls.duplicate_txt_shape_onto_shape(shp, slide)
                except EnvironmentError:
                    errors += 1
            if errors > 0:
                bkt.message.error("Kopierfehler bei %s Shape(s)." % errors)

        # IMPORTANT: many copy-paste operations lead to EnvironmentError. Putting all in a thread solves this issue.
        t = Thread(target=loop)
        t.start()
        t.join()

        # Alternative way:
        # bkt.ui.execute_with_progress_bar(loop, context, indeterminate=True)
    
    ### context menu callbacks ###

    @staticmethod
    def is_outable(shape):
        return shape.HasTextFrame == -1 and shape.TextFrame.HasText == -1 and shape.Type not in [pplib.MsoShapeType['msoTextBox'], pplib.MsoShapeType['msoPlaceholder']]
    
    @staticmethod
    def is_mergable(shapes):
        return len(shapes) == 2 and (shapes[0].Type == pplib.MsoShapeType['msoTextBox'] or shapes[1].Type == pplib.MsoShapeType['msoTextBox'])


class SplitTextShapes(object):


    ### This method was required in Office 2007 were BoundXXX methods were not available
    # @classmethod
    # def paragraph_height(cls, par, with_par_spaces=True):
    #     parHeight = par.Lines().Count * cls.line_height(par) * 1.0
    #     if with_par_spaces:
    #         parHeight = parHeight + max(0, par.ParagraphFormat.SpaceBefore) + max(0, par.ParagraphFormat.SpaceAfter)
    #     return parHeight
    

    ### This method was required in Office 2007 were BoundXXX methods were not available
    # @staticmethod
    # def line_height(par):
    #     if par.ParagraphFormat.LineRuleWithin == -1:
    #         # spacing = number of lines
    #         # Annahme zur Korrektur der Abstände: Abstand zwischen zwei Zeilen ist 0.2pt
    #         return par.Font.Size * (max(0, par.ParagraphFormat.SpaceWithin) + 0.2)
    #     else:
    #         # spacing = number of pt
    #         # Annahme zur Korrektur der Abstände: Abstand zwischen zwei Zeilen ist 0.2pt
    #         return par.ParagraphFormat.SpaceWithin #+ 0.1 * .Font.Size
    
    
    @staticmethod
    def trim_newline_character(par):
        if par.Characters(par.Length, 1).Text == "\r":
            par.Characters(par.Length, 1).Delete()
    
    
    @classmethod
    def splitShapesByParagraphs(cls, shapes, context):
        for shp in shapes:
            # if shp.TextFrame2.TextRange.Text != "":
            if shp.HasTextFrame == -1 and shp.TextFrame.HasText == -1 and shp.TextFrame.TextRange.Paragraphs().Count > 1:
                #Shape exklusiv markieren (alle anderen deselektieren)
                # shp.Select(-1) # msoTrue

                par_count = shp.TextFrame2.TextRange.Paragraphs().Count
                for par_index in range(2, par_count+1):
                    par = shp.TextFrame2.TextRange.Paragraphs(par_index)
                    # Leere Paragraphen überspringen
                    if par.text in ["", "\r"]:
                        continue
                    # Shape dublizieren
                    shpCopy = shp.Duplicate()
                    shpCopy.Select(0) # msoFalse
                    shpCopy.Top  = par.BoundTop - shp.TextFrame2.MarginTop + par.ParagraphFormat.SpaceBefore
                    shpCopy.Left = shp.Left
                    
                    # Absaetze 1..i-1 entfernen und Shape entsprechend verschieben
                    shpCopy.TextFrame2.TextRange.Paragraphs(1, par_index-1).Delete()
                    # for index in range(1, par_index):
                    #     # Textbox Position entsprechend Absatzhoehe anpassen
                    #     # shpCopy.Top = shpCopy.Top + cls.paragraph_height(shpCopy.TextFrame2.TextRange.Paragraphs(1))
                    #     # Absatz entfernen
                    #     shpCopy.TextFrame2.TextRange.Paragraphs(1).Delete()
                    
                    # Absaetze i+1..n entfernen
                    shpCopy.TextFrame2.TextRange.Paragraphs(2, shpCopy.TextFrame2.TextRange.Paragraphs().Count-1).Delete()
                    # for index in range(par_index + 1, shp.TextFrame2.TextRange.Paragraphs().Count + 1):
                    #     shpCopy.TextFrame2.TextRange.Paragraphs(2).Delete()

                    # Letztes CR-Zeichen loeschen
                    cls.trim_newline_character(shpCopy.TextFrame2.TextRange)

                    # Shape Hoehe abhaengig von Absaetzhoehe
                    # shpCopy.Height = cls.paragraph_height(shpCopy.TextFrame.TextRange.Paragraphs(1)) + shpCopy.TextFrame2.MarginTop + shpCopy.TextFrame2.MarginBottom
                    shpCopy.Height = par.BoundHeight + shp.TextFrame2.MarginTop + shp.TextFrame2.MarginBottom - par.ParagraphFormat.SpaceBefore - par.ParagraphFormat.SpaceAfter
                    if par_index == par_count:
                        #last paragraph does not have spaceafter
                        shpCopy.Height += par.ParagraphFormat.SpaceAfter

                    # --> ein Absatz bleibt übrig

                # letzten Shape nach unten schieben
                # shpCopy.Top = max(shpCopy.Top, shp.Top + shp.Height - shpCopy.Height)

                # Absaetze 2..n im Original-Shape entfernen
                shp.TextFrame2.TextRange.Paragraphs(2, par_count-1).Delete()

                # Letztes CR-Zeichen loeschen
                cls.trim_newline_character(shp.TextFrame2.TextRange)
                # Textbox Hoehe an Absatzhoehe anpassen
                # shp.Height = cls.paragraph_height(shp.TextFrame2.TextRange.Paragraphs(1)) + shp.TextFrame2.MarginTop + shp.TextFrame2.MarginBottom
                shp.Height = shp.TextFrame2.TextRange.Paragraphs(1).BoundHeight + shp.TextFrame2.MarginTop + shp.TextFrame2.MarginBottom
                
                #Verteilung bei 2 Shapes führt zu Fehler
                # if context.app.ActiveWindow.Selection.ShapeRange.Count > 2:
                #     # Objekte vertikal verteilen
                #     context.app.ActiveWindow.Selection.ShapeRange.Distribute(
                #         1, #msoDistributeVertically
                #         0) #msoFalse)

    @classmethod
    def joinShapesWithText(cls, shapes):
        # Shapes nach top sortieren
        shapes = sorted(shapes, key=lambda shape: shape.Top)
        # Anapssung Größe des ersten Shapes (Master-Shape) mit TextFrame
        for i in range(len(shapes)):
            shpMaster = shapes.pop(i) #shapes[0]
            if shpMaster.HasTextFrame:
                break
        else:
            # no shape with textframe found
            return
        
        shpMaster.Height = max(shpMaster.Height, shapes[-1].Top + shapes[-1].Height - shpMaster.Top)

        for shp in shapes: #[1:]:
            # Text aus Shape kopieren
            shp.TextFrame2.TextRange.Copy()
            # neuen Absatz in Master-Shape erstellen
            # Bug in PowerPoint: machmal muss InsertAfter mehrmals aufgerufen werden
            parCount = shpMaster.TextFrame2.TextRange.Paragraphs().Count
            for i in range(10):
                txtRange = shpMaster.TextFrame2.TextRange.Paragraphs().InsertAfter("\r")
                if parCount < shpMaster.TextFrame2.TextRange.Paragraphs().Count:
                    break
            # Text in Master-Shape einfuegen
            txtRange.Paste()
            # Letztes CR-Zeichen loeschen
            cls.trim_newline_character(txtRange)
            # Shape loeschen
            shp.Delete()
    
    ### context menu callbacks ###

    @staticmethod
    def is_splitable(shape):
        return shape.HasTextFrame == -1 and shape.TextFrame2.TextRange.Paragraphs().Count>1
    
    @staticmethod
    def is_joinable(shapes):
        return any(shp.HasTextFrame == -1 and shp.TextFrame2.HasText == -1 for shp in shapes)


class TextShapes(object):
    sticker_alignment = bkt.settings.get("toolbox.sticker_alignment", "right")
    sticker_fontsize = bkt.settings.get("toolbox.sticker_fontsize", 14)
    sticker_custom = bkt.settings.get("toolbox.sticker_custom", None)

    @classmethod
    def settings_setter(cls, name, value):
        setattr(cls, name, value)
        bkt.settings["toolbox."+name] = value
    
    @staticmethod
    def addUnderlinedTextbox(slide, presentation):
        # Textbox erstellen, damit Standardformatierung der Textbox genommen wird
        shp = slide.shapes.AddTextbox( 1 #msoTextOrientationHorizontal
            , 100, 100, 200, 50)
        # Shape-Typ ist links-rechts-Pfeil, weil es die passenden Connector-Ecken hat
        shp.AutoShapeType = pplib.MsoAutoShapeType['msoShapeLeftRightArrow']
        # Shape-Anpassung, so dass es wie ein Rechteck aussieht
        shp.Adjustments.item[1] = 1
        shp.Adjustments.item[2] = 0
        # Text
        shp.TextFrame.TextRange.text = "Lorem ipsum"

        # Mitting ausrichten
        shp.Top = (presentation.PageSetup.SlideHeight - shp.height) /2
        shp.Left = (presentation.PageSetup.SlideWidth - shp.width) /2

        # Connectoren erstellen und mit Connector-Ecken des Shapes verbinden
        connector = slide.shapes.AddConnector(Type=1 #msoConnectorStraight
            , BeginX=0, BeginY=0, EndX=100, EndY=100)
        connector.ConnectorFormat.BeginConnect(ConnectedShape=shp, ConnectionSite=5)
        connector.ConnectorFormat.EndConnect(ConnectedShape=shp, ConnectionSite=7)
        
        # Default Formatierung
        color = shp.TextFrame.TextRange.Font.Color
        if color.Type == pplib.MsoColorType["msoColorTypeScheme"]:
            connector.Line.ForeColor.ObjectThemeColor = color.ObjectThemeColor
            connector.Line.ForeColor.Brightness = color.Brightness
        else:
            connector.Line.ForeColor.RGB = color.RGB
        connector.Line.Weight = 1.5
        # shp.TextFrame.MarginBottom = 0
        # shp.TextFrame.MarginTop    = 0
        # shp.TextFrame.MarginLeft   = 0
        # shp.TextFrame.MarginRight  = 0

        # Text auswählen
        shp.TextFrame2.TextRange.Select()
    
    
    @classmethod
    def addSticker(cls, slide, presentation, sticker_text="DRAFT", select_text=True):
        # Textbox erstellen, damit Standardformatierung der Textbox genommen wird
        shp = slide.shapes.AddTextbox( 1 #msoTextOrientationHorizontal
            , 0, 60, 100, 20)
        # Shape-Typ ist links-rechts-Pfeil, weil es die passenden Connector-Ecken hat
        shp.AutoShapeType = pplib.MsoAutoShapeType['msoShapeLeftRightArrow']
        # Shape-Anpassung, so dass es wie ein Rechteck aussieht
        shp.Adjustments.item[1] = 1
        shp.Adjustments.item[2] = 0
        # Shape-Stil
        # shp.Line.Weight = 0.75
        shp.Fill.Visible = 0 #msoFalse
        shp.Line.Visible = 0 #msoFalse
        # Text-Stil
        # shp.TextFrame.TextRange.Font.Color.RGB = 0
        shp.TextFrame.TextRange.Font.Size = cls.sticker_fontsize
        shp.TextFrame.TextRange.ParagraphFormat.Bullet.Visible = False
        # Autosize / Text nicht umbrechen
        shp.TextFrame.WordWrap = 0 #msoFalse
        shp.TextFrame.AutoSize = 1 #ppAutoSizeShapeToFitText
        # Innenabstand
        shp.TextFrame.MarginBottom = 0
        shp.TextFrame.MarginTop    = 0
        shp.TextFrame.MarginLeft   = 0
        shp.TextFrame.MarginRight  = 0
        # Text
        shp.TextFrame.TextRange.text = sticker_text
        # Top-Position
        shp.Top = 15

        # Text alignment + Left-Position
        if cls.sticker_alignment == "left":
            shp.TextFrame.TextRange.ParagraphFormat.Alignment = 1 #ppAlignLeft
            shp.Left = 15
        elif cls.sticker_alignment == "center":
            shp.TextFrame.TextRange.ParagraphFormat.Alignment = 2 #ppAlignCenter
            shp.Left = presentation.PageSetup.SlideWidth/2 - shp.width/2
        else: #right
            shp.TextFrame.TextRange.ParagraphFormat.Alignment = 3 #ppAlignRight
            shp.Left = presentation.PageSetup.SlideWidth - shp.width - 15

        # Connectoren erstellen und mit Connector-Ecken des Shapes verbinden
        connector1 = slide.shapes.AddConnector(Type=1 #msoConnectorStraight
            , BeginX=0, BeginY=0, EndX=100, EndY=100)
        connector1.ConnectorFormat.BeginConnect(ConnectedShape=shp, ConnectionSite=1)
        connector1.ConnectorFormat.EndConnect(ConnectedShape=shp, ConnectionSite=3)
        connector1.Line.ForeColor.RGB = 0
        connector1.Line.Weight = 0.75

        connector2 = slide.shapes.AddConnector(Type=1 #msoConnectorStraight
            , BeginX=0, BeginY=0, EndX=100, EndY=100)
        connector2.ConnectorFormat.BeginConnect(ConnectedShape=shp, ConnectionSite=5)
        connector2.ConnectorFormat.EndConnect(ConnectedShape=shp, ConnectionSite=7)
        connector2.Line.ForeColor.RGB = 0
        connector2.Line.Weight = 0.75

        color = shp.TextFrame.TextRange.Font.Color
        if color.Type == pplib.MsoColorType["msoColorTypeScheme"]:
            connector1.Line.ForeColor.ObjectThemeColor = color.ObjectThemeColor
            connector1.Line.ForeColor.Brightness = color.Brightness
            connector2.Line.ForeColor.ObjectThemeColor = color.ObjectThemeColor
            connector2.Line.ForeColor.Brightness = color.Brightness
        else:
            connector1.Line.ForeColor.RGB = color.RGB
            connector2.Line.ForeColor.RGB = color.RGB

        if select_text:
            # Text auswählen
            shp.Select()
            shp.TextFrame2.TextRange.Select()
    
    
    @classmethod
    def add_sticker_to_slides(cls, slides, presentation, sticker_text="DRAFT", current_control=None):
        if current_control:
            sticker_text = current_control["tag"]
        select_text = len(slides) == 1
        for slide in slides:
            cls.addSticker(slide, presentation, sticker_text, select_text)
    
    @classmethod
    def own_sticker_enabled(cls):
        return cls.sticker_custom is not None
    
    @classmethod
    def own_sticker_label(cls):
        if cls.sticker_custom:
            return cls.sticker_custom + "-Sticker"
        else:
            return "Noch nicht definiert"
    
    @classmethod
    def own_sticker_insert(cls, slides, presentation):
        if cls.sticker_custom and not bkt.get_key_state(bkt.KeyCodes.SHIFT):
            cls.add_sticker_to_slides(slides, presentation, cls.sticker_custom)
        else:
            res = cls.own_sticker_edit(slides, presentation)
    
    @classmethod
    def own_sticker_edit(cls, slides, presentation):
        res = bkt.ui.show_user_input("Selbst definierten Sticker-Text eingeben:", "Sticker bearbeiten", cls.sticker_custom)
        if res:
            cls.sticker_custom = bkt.settings["toolbox.sticker_custom"] = res
            cls.add_sticker_to_slides(slides, presentation, res)


sticker_menu = lambda: bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None, 
                children=[
                    bkt.ribbon.Button(
                        id="sticker_draft",
                        label = "DRAFT-Sticker",
                        image = "Sticker",
                        screentip="DRAFT-Sticker einfügen",
                        supertip="Füge ein Sticker-Shape oben rechts auf dem aktuellen Slide mit Text DRAFT ein.",
                        on_action=bkt.Callback(TextShapes.add_sticker_to_slides, slides=True, presentation=True)
                    ),
                    bkt.ribbon.Button(
                        id="sticker_backup",
                        label = "BACKUP-Sticker",
                        screentip="BACKUP-Sticker einfügen",
                        supertip="Füge ein Sticker-Shape oben rechts auf dem aktuellen Slide mit Text BACKUP ein.",
                        on_action=bkt.Callback(TextShapes.add_sticker_to_slides, slides=True, presentation=True, current_control=True),
                        tag="BACKUP"
                    ),
                    bkt.ribbon.Button(
                        id="sticker_discussion",
                        label = "FOR DISCUSSION-Sticker",
                        screentip="FOR DISCUSSION-Sticker einfügen",
                        supertip="Füge ein Sticker-Shape oben rechts auf dem aktuellen Slide mit Text FOR DISCUSSION ein.",
                        on_action=bkt.Callback(TextShapes.add_sticker_to_slides, slides=True, presentation=True, current_control=True),
                        tag="FOR DISCUSSION"
                    ),
                    bkt.ribbon.Button(
                        id="sticker_illustrative",
                        label = "ILLUSTRATIVE-Sticker",
                        screentip="ILLUSTRATIVE-Sticker einfügen",
                        supertip="Füge ein Sticker-Shape oben rechts auf dem aktuellen Slide mit Text ILLUSTRATIVE ein.",
                        on_action=bkt.Callback(TextShapes.add_sticker_to_slides, slides=True, presentation=True, current_control=True),
                        tag="ILLUSTRATIVE"
                    ),
                    bkt.ribbon.Button(
                        id="sticker_confidential",
                        label = "CONFIDENTIAL-Sticker",
                        screentip="CONFIDENTIAL-Sticker einfügen",
                        supertip="Füge ein Sticker-Shape oben rechts auf dem aktuellen Slide mit Text CONFIDENTIAL ein.",
                        on_action=bkt.Callback(TextShapes.add_sticker_to_slides, slides=True, presentation=True, current_control=True),
                        tag="CONFIDENTIAL"
                    ),
                    bkt.ribbon.MenuSeparator(),
                    bkt.ribbon.Button(
                        id="sticker_own",
                        get_label=bkt.Callback(TextShapes.own_sticker_label),
                        screentip="Selbst definierten Sticker einfügen",
                        supertip="Füge ein Sticker-Shape oben rechts auf dem aktuellen Slide mit selbst definiertem ein.",
                        on_action=bkt.Callback(TextShapes.own_sticker_insert, slides=True, presentation=True),
                        get_enabled=bkt.Callback(TextShapes.own_sticker_enabled)
                    ),
                    bkt.ribbon.Button(
                        id="sticker_own_edit",
                        label = "Sticker-Text ändern",
                        screentip="Selbst definierten Sticker bearbeiten",
                        supertip="Ändere den Text des selbst definierten Stickers und füge diesen sofort ein.",
                        on_action=bkt.Callback(TextShapes.own_sticker_edit, slides=True, presentation=True)
                    ),
                    bkt.ribbon.MenuSeparator(),
                    bkt.ribbon.Menu(
                        label="Ausrichtung",
                        supertip="Ausrichtungsoptionen für Sticker einstellen",
                        children=[
                            bkt.ribbon.ToggleButton(
                                label="Links",
                                screentip="Sticker-Ausrichtung Links",
                                supertip="Setzt die Position beim Einfügen der Sticker auf links.",
                                get_pressed=bkt.Callback(lambda: TextShapes.sticker_alignment == "left"),
                                on_toggle_action=bkt.Callback(lambda pressed: TextShapes.settings_setter("sticker_alignment", "left")),
                            ),
                            bkt.ribbon.ToggleButton(
                                label="Mitte",
                                screentip="Sticker-Ausrichtung Mitte",
                                supertip="Setzt die Position beim Einfügen der Sticker auf mittig.",
                                get_pressed=bkt.Callback(lambda: TextShapes.sticker_alignment == "center"),
                                on_toggle_action=bkt.Callback(lambda pressed: TextShapes.settings_setter("sticker_alignment", "center")),
                            ),
                            bkt.ribbon.ToggleButton(
                                label="Rechts",
                                screentip="Sticker-Ausrichtung Rechts",
                                supertip="Setzt die Position beim Einfügen der Sticker auf rechts.",
                                get_pressed=bkt.Callback(lambda: TextShapes.sticker_alignment == "right"),
                                on_toggle_action=bkt.Callback(lambda pressed: TextShapes.settings_setter("sticker_alignment", "right")),
                            ),
                        ]
                    ),
                    bkt.ribbon.Menu(
                        label="Schriftgröße",
                        supertip="Schriftgrößenoptionen für Sticker einstellen",
                        children=[
                            bkt.ribbon.ToggleButton(
                                label="10",
                                screentip="Sticker-Schriftgröße 10",
                                supertip="Setzt die Schriftgröße beim Einfügen der Sticker auf 10",
                                get_pressed=bkt.Callback(lambda: TextShapes.sticker_fontsize == 10),
                                on_toggle_action=bkt.Callback(lambda pressed: TextShapes.settings_setter("sticker_fontsize", 10)),
                            ),
                            bkt.ribbon.ToggleButton(
                                label="11",
                                screentip="Sticker-Schriftgröße 11",
                                supertip="Setzt die Schriftgröße beim Einfügen der Sticker auf 11",
                                get_pressed=bkt.Callback(lambda: TextShapes.sticker_fontsize == 11),
                                on_toggle_action=bkt.Callback(lambda pressed: TextShapes.settings_setter("sticker_fontsize", 11)),
                            ),
                            bkt.ribbon.ToggleButton(
                                label="12",
                                screentip="Sticker-Schriftgröße 12",
                                supertip="Setzt die Schriftgröße beim Einfügen der Sticker auf 12",
                                get_pressed=bkt.Callback(lambda: TextShapes.sticker_fontsize == 12),
                                on_toggle_action=bkt.Callback(lambda pressed: TextShapes.settings_setter("sticker_fontsize", 12)),
                            ),
                            bkt.ribbon.ToggleButton(
                                label="14",
                                screentip="Sticker-Schriftgröße 14",
                                supertip="Setzt die Schriftgröße beim Einfügen der Sticker auf 14",
                                get_pressed=bkt.Callback(lambda: TextShapes.sticker_fontsize == 14),
                                on_toggle_action=bkt.Callback(lambda pressed: TextShapes.settings_setter("sticker_fontsize", 14)),
                            ),
                        ]
                    ),
                ])