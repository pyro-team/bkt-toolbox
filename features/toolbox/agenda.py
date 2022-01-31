# -*- coding: utf-8 -*-
'''
Created on 18.05.2016

@author: rdebeerst
'''

from __future__ import absolute_import

import logging
import json
import uuid

import bkt
import bkt.library.powerpoint as pplib

import bkt.dotnet as dotnet
ppt = dotnet.import_powerpoint()
office = dotnet.import_officecore()


TOOLBOX_AGENDA = "TOOLBOX-AGENDA"
TOOLBOX_AGENDA_SLIDENO  = "TOOLBOX-AGENDA-SLIDENO"
TOOLBOX_AGENDA_SELECTOR = "TOOLBOX-AGENDA-SELECTOR"
TOOLBOX_AGENDA_TEXTBOX  = "TOOLBOX-AGENDA-TEXTBOX"
TOOLBOX_AGENDA_SETTINGS = "TOOLBOX-AGENDA-SETTINGS"

TOOLBOX_AGENDA_POPUP = "TOOLBOX-AGENDA-POPUP"

#internal settings
SETTING_POSITION = "position"
SETTING_TEXT = "text"
SETTING_AGENDA_ID = "id"
SETTING_INDENT_LEVEL = "indent-level"
#user configurable
SETTING_HIDE_SUBITEMS = "hide-sub-items"
SETTING_CREATE_SECTIONS = "sections-create"
SETTING_CREATE_LINKS = "hyperlinks-create"
SETTING_SLIDES_FOR_SUBITEMS = "slides-for-sub-items"
SETTING_SELECTOR_MARGIN = "selector-margin"
#old selector settings (for backwards compatibility):
SETTING_SELECTOR_FILL_COLOR = "selector-fill-color"
SETTING_SELECTOR_LINE_COLOR = "selector-line-color"
#new selector settings:
SETTING_SELECTOR_STYLE_FILL = "selector-style-fill"
SETTING_SELECTOR_STYLE_LINE = "selector-style-line"
SETTING_SELECTOR_STYLE_TEXT = "selector-style-text"


# Kind of mutable named tuple for agenda items:
class AgendaEntry(object):
    __slots__ = ("position", "text", "indentlevel", "slide")

    def __init__(self, position, text, indentlevel, slide=None):
        self.position    = position
        self.text        = text
        self.indentlevel = indentlevel
        self.slide       = slide

    def __getitem__(self, item):
        return getattr(self, self.__slots__[item])

    def __setitem__(self, item, value):
        return setattr(self, self.__slots__[item], value)

    def __len__(self):
        return len(self.__slots__)

    def __str__(self):
        clsname = self.__class__.__name__
        values = ', '.join('%s=%r' % (k, getattr(self, k))
                           for k in self.__slots__)
        return '%s(%s)' % (clsname, values)

    __repr__ = __str__


class ToolboxAgenda(object):
    '''
    Class to manage agendas on PowerPoint slides
    
    Agendas are typically a bullet-type list, but
    any textbox can be used for an agenda.
    
    Agenda slides are tagged and recognized by "Toolbox-Agenda"
    
    Updating Agenda:
    When updating an agenda, existing agenda-slides should be updated.
    Existing agenda slides are recognised by a specific tag added to the slide.
    Update is done by replacing the current agenda slide with newly
    created agenda slides. Personal customizations on agenda-slides are lost.
    
    '''
    # selectorFillColor = 12566463 # 193 193 193   ##   193+193*255+193*255*255
    # selectorLineColor = 8355711 # 127 127 127    ##   ((long(127)*255)+127)*255+127

    #color: [theme, brightness, rgb]
    default_selectorFillColor = {
        'color': [16, 0, 0],
        'visibility': -1,
        }
    default_selectorLineColor = {
        'color': [13, 0, 0],
        'visibility': 0,
        'weight': 0.75,
        'style': 1,
        'dashstyle': 1,
        }
    default_selectorTextColor = {
        'color': [13, 0, 0],
        'bold': True,
        'italic': False,
        'underline': False,
        }

    selectorFillColor = default_selectorFillColor.copy()
    selectorLineColor = default_selectorLineColor.copy()
    selectorTextColor = default_selectorTextColor.copy()

    #indicate if selector style has been picked up (reset on each agenda update)
    selector_style_pickup = False
    
    default_settings = {
        SETTING_AGENDA_ID: None,
        SETTING_HIDE_SUBITEMS: False,
        SETTING_CREATE_SECTIONS: False,
        SETTING_CREATE_LINKS: False,
        SETTING_SLIDES_FOR_SUBITEMS: True,
        SETTING_SELECTOR_MARGIN: 0.2
    }
    
    settings = None

    
    
    # =================
    # = create agenda =
    # =================
    
    @classmethod
    def create_agenda_textbox_on_slide(cls, slide, context=None):
        '''
        Create a new agenda textbox on current slide with default formatting and agenda-tag.
        Textbox is prefilled with section if they exist. With context agenda-tab is activated.
        '''
        
        try:
            #slide.shapes(1).TextFrame.TextRange.text = "Agenda"
            # Shape rechts oben auf slide erstellen
            shp = slide.shapes.AddShape(office.MsoAutoShapeType.msoShapeRectangle.value__, 160, 240, 400, 100)
        
            # Shape-Typ ist links-rechts-Pfeil, weil es die passenden Connector-Ecken hat
            shp.AutoShapeType = office.MsoAutoShapeType.msoShapeLeftRightArrow.value__
            # Shape-Anpassung, so dass es wie ein Rechteck aussieht
            shp.Adjustments.item[1] = 1
            shp.Adjustments.item[2] = 0
            # Shape-Stil
            shp.Line.Weight = 0.75
            shp.Fill.Visible = False
            shp.Line.Visible = False
            # Text-Stil
            # shp.TextFrame.TextRange.Font.Color.RGB = 0
            shp.TextFrame.TextRange.Font.Color.ObjectThemeColor = 13 #msoThemeColorText1
            shp.TextFrame.TextRange.Font.Size = 14
            shp.TextFrame.TextRange.ParagraphFormat.Alignment = ppt.PpParagraphAlignment.ppAlignLeft.value__
            # Autosize / Textumbruch
            shp.TextFrame.WordWrap = True
            shp.TextFrame.AutoSize = ppt.PpAutoSize.ppAutoSizeShapeToFitText.value__
            shp.width = 400
            # Text (prefill with sections)
            sections = slide.Parent.SectionProperties
            if sections.count > 1:
                text = ""
                for i in range(sections.count):
                    text += sections.Name(i+1) + "\r"
                shp.TextFrame.TextRange.text = text.strip()
            else:
                shp.TextFrame.TextRange.text = "Abschnitt 1\rAbschnitt 2\rAbschnitt 3"
            
            # Einrückung
            shp.TextFrame.VerticalAnchor = office.MsoVerticalAnchor.msoAnchorMiddle.value__
            shp.TextFrame.Ruler.Levels.item(1).FirstMargin = 0
            shp.TextFrame.Ruler.Levels.item(1).LeftMargin = 14
            # Tab-Stop für rechte Spalte (bspw. für Zeit)
            shp.TextFrame.Ruler.TabStops.Add(ppt.PpTabStopType.ppTabStopRight.value__, shp.width)
            # Innenabstand
            shp.TextFrame.MarginBottom = 12
            shp.TextFrame.MarginTop = 12
            shp.TextFrame.MarginLeft = 6
            shp.TextFrame.MarginRight = 6
            # Bullet Style
            shp.TextFrame.TextRange.ParagraphFormat.Bullet.Type = ppt.PpBulletType.ppBulletUnnumbered.value__
            shp.TextFrame.TextRange.ParagraphFormat.Bullet.Character = 167
            shp.TextFrame.TextRange.ParagraphFormat.Bullet.Font.Name = "Wingdings"
            # Absatzabstand und Zeilenabstand
            shp.TextFrame.TextRange.ParagraphFormat.LineRuleWithin = -1
            shp.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 1
            shp.TextFrame.TextRange.ParagraphFormat.LineRuleBefore = 0
            shp.TextFrame.TextRange.ParagraphFormat.SpaceBefore = 0
            shp.TextFrame.TextRange.ParagraphFormat.LineRuleAfter = 0
            shp.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 18
            # Mittig anordnen
            shp.left = (slide.Parent.PageSetup.SlideWidth-shp.Width)/2
            
            # Connectoren erstellen und mit Connector-Ecken des Shapes verbinden
            connector = slide.shapes.AddConnector(Type=office.MsoConnectorType.msoConnectorStraight.value__, BeginX=0,BeginY=0, EndX=100, EndY=100)
            connector.ConnectorFormat.BeginConnect(ConnectedShape=shp, ConnectionSite=1)
            connector.ConnectorFormat.EndConnect(ConnectedShape=shp, ConnectionSite=3)
            connector.Line.Visible = -1
            connector.Line.ForeColor.ObjectThemeColor = 13 #msoThemeColorText1
            # connector.Line.ForeColor.RGB = 0
            connector.Line.Weight = 0.75
            connector.Line.DashStyle = 1 #straight line

            connector = slide.shapes.AddConnector(Type=office.MsoConnectorType.msoConnectorStraight.value__, BeginX=0,BeginY=0, EndX=100, EndY=100)
            connector.ConnectorFormat.BeginConnect(ConnectedShape=shp, ConnectionSite=5)
            connector.ConnectorFormat.EndConnect(ConnectedShape=shp, ConnectionSite=7)
            connector.Line.Visible = -1
            connector.Line.ForeColor.ObjectThemeColor = 13 #msoThemeColorText1
            # connector.Line.ForeColor.RGB = 0
            connector.Line.Weight = 0.75
            connector.Line.DashStyle = 1 #straight line
            
            cls.set_tags_for_slide(slide)
            cls.set_tags_for_textbox(shp)

            # select shape (to show popup)
            shp.Name = "[BKT] Agenda-Textbox %s" % shp.id
            shp.select()

            if context:
                context.ribbon.ActivateTab('bkt_context_tab_agenda')
        
        except:
            logging.exception("Agenda: agenda textbox creation failed")
            bkt.message.error("Fehler beim Anlegen der Agenda-Textbox", title="Toolbox: Agenda")
            # bkt.helpers.exception_as_message()
    
    
    @classmethod
    def create_agenda_from_textbox(cls, master_textbox, context=None):
        '''
        Initially create agenda slides based on agenda textbox created before.
        With context agenda-tab is activated.
        '''
        master_slide = master_textbox.parent
        
        # set tags for master slide
        # default settings with new id
        settings = dict(cls.default_settings)
        settings.update({
            SETTING_AGENDA_ID: str(uuid.uuid4()),
            SETTING_POSITION: 0,
            SETTING_TEXT: None,
            SETTING_INDENT_LEVEL: None
        })
        cls.write_agenda_settings_to_slide(master_slide, settings)
        cls.set_tags_for_slide(master_slide, 0)
        #cls.set_tags_for_textbox(master_textbox)

        # retreive agenda settings from textbox
        agenda_entries = cls.agenda_entries_from_textbox(master_textbox, slides_for_subitems=settings.get(SETTING_SLIDES_FOR_SUBITEMS) or False)
        #bkt.console.show_message("entries: %s" % agenda_entries)
        
        new_slide_count = 0
        
        # iterate through paragraphs
        for agena_item in agenda_entries:
            if agena_item.text != "":
                # create slide 
                new_slide_count = new_slide_count + 1
                slide = master_slide.Duplicate(1)
                slide.SlideShowTransition.Hidden = 0
                slide.MoveTo(master_slide.SlideIndex + new_slide_count)
                agena_item.slide = slide
                
                # update agenda
                textbox = cls.get_agenda_textbox_on_slide(slide)
                cls.update_agenda_on_slide_new(slide, textbox, agena_item.position, settings)

                # update sections
                cls.update_agenda_sections_for_slide(slide, agena_item.text, settings)

        # go to agenda tab
        if context:
            context.ribbon.ActivateTab('bkt_context_tab_agenda')
    
    @classmethod
    def create_agenda_from_slide(cls, slide, context):
        '''
        Initially create agenda slides based on agenda textbox created before on current slide.
        '''
        
        master_textbox = cls.get_agenda_textbox_on_slide(slide)
        if master_textbox is None:
            bkt.message.warning("Keine Agenda-Textbox auf der Folie vorhanden.", title="Toolbox: Agenda")
            return
        
        cls.create_agenda_from_textbox(master_textbox, context)
    
    
    @classmethod
    def get_or_create_selector_on_slide(cls, sld):
        '''
        Finds selector on slide or creates a new selector-shape.
        If existing selector is found, also picks-up selector style.
        '''
        shp = cls.get_shape_with_tag_item(sld, TOOLBOX_AGENDA_SELECTOR)
        if not shp is None:
            # pickup selector format (this allows to use gradients and other fancy stuff)
            cls.selector_style_pickup = True
            shp.PickUp()
            return shp

        # Neues Selector-Shape erstellen
        shp = sld.shapes.AddShape(office.MsoAutoShapeType.msoShapeRectangle.value__, 0, 0, 100, 20)
        cls.set_tags_for_selector(shp)
        shp.Name = "[BKT] Agenda-Selektor %s" % shp.id
        try:
            #try to set selector right behind textbox (better for fancy agenda formatting)
            textbox = cls.get_agenda_textbox_on_slide(sld)
            pplib.set_shape_zorder(shp, textbox.ZOrderPosition)
        except:
            shp.ZOrder(office.MsoZOrderCmd.msoSendToBack.value__)
        # Grauer Hintergrund/Rand
        cls.set_selector_fill(shp.Fill, cls.selectorFillColor)
        cls.set_selector_line(shp.Line, cls.selectorLineColor)
        
        return shp
    
    
    @classmethod
    def update_or_create_agenda_from_slide(cls, slide, context):
        if cls.is_agenda_slide(slide):
            cls.update_agenda_slides_by_slide(slide)
        elif cls.can_create_agenda_from_slide(slide):
            cls.create_agenda_from_slide(slide, context)
        else:
            bkt.message.warning("Agenda nicht gefunden!", title="Toolbox: Agenda")


    
    # ===============
    # = find agenda =
    # ===============
    
    @classmethod
    def find_agenda_items_by_slide(cls, slide):
        '''
        returns list of agenda entries [position, text, indentlevel, slide-reference] 
        considering agenda slides according to the agenda-id of the given slide or
        (if no agenda is contained on the given slide) according to the first agenda
        in the presentation 
        '''
        
        settings = cls.get_agenda_settings_from_slide(slide)
        if settings == {}:
            # no agenda settings found on slide
            if cls.is_agenda_slide(slide):
                # find all agenda-slides in presentation
                # bkt.message(slide.Tags.Item(TOOLBOX_AGENDA))
                bkt.message.warning("Keine Agenda-Einstellungen auf aktueller Folie. Durchsuche alle Agenda-Folien.", title="Toolbox: Agenda")
                return cls.find_all_agenda_slides(slide.parent)
            else:
                # slide is not an agenda slide
                bkt.message.warning("Aktuelle Folie ist keine Agenda-Folie!", title="Toolbox: Agenda")
                return []
        
        return cls.find_agenda_items_by_id(slide.parent, settings[SETTING_AGENDA_ID])
    
    
    @classmethod
    def find_agenda_items_by_id(cls, presentation, id):
        '''
        returns list of agenda entries [position, text, indentlevel, slide-reference] 
        considering all agenda slides with given id in the presentation
        '''
        agenda_slides = []
        
        for slide in presentation.slides:
            settings = cls.get_agenda_settings_from_slide(slide)
            if settings.get(SETTING_AGENDA_ID, None) == id:
                agenda_slides.append( AgendaEntry( settings.get(SETTING_POSITION), settings.get(SETTING_TEXT), settings.get(SETTING_INDENT_LEVEL), slide ) )
        
        return agenda_slides
    
    
    @classmethod
    def find_all_agenda_slides(cls, presentation):
        '''
        returns list of agenda entries [position, text=None, indentlevel=None, slide-reference] 
        considering all agenda slides in the presentation
        '''
        agenda_slides = []
        
        for slide in presentation.slides:
            if cls.is_agenda_slide(slide):
                agenda_slides.append( AgendaEntry( int(slide.Tags.Item(TOOLBOX_AGENDA_SLIDENO)), None, None, slide ) )
        
        return agenda_slides
    
    @classmethod
    def presentation_has_agenda(cls, presentation):
        '''
        check if any slide has agenda-tag on slide
        '''
        for slide in presentation.slides:
            if cls.is_agenda_slide(slide):
                return True
        return False
    
    @classmethod
    def get_agenda_textbox_on_slide(cls, sld):
        '''
        return agenda-textbox on given slide
        agenda-textbox is recognised by the tag TOOLBOX_AGENDA_TEXTBOX
        '''
        return cls.get_shape_with_tag_item(sld, TOOLBOX_AGENDA_TEXTBOX)
    
    
    @staticmethod
    def agenda_entries_from_textbox(textbox, slides_for_subitems=True):
        '''
        returns list of agenda entries [position(=par_index), text, indentlevel, slide-reference]
        '''
        agenda_entries = []
        for idx in range(1, textbox.TextFrame.TextRange.Paragraphs().Count+1):
            cur_paragraph = textbox.TextFrame.TextRange.Paragraphs(idx)
            if cur_paragraph.IndentLevel <=1 or slides_for_subitems:
                agenda_entries.append( AgendaEntry(idx, cur_paragraph.Text.strip(), cur_paragraph.IndentLevel, None) )
        
        return agenda_entries
    
    
    
    
    
    # =================
    # = update agenda =
    # =================
    
    @classmethod 
    def update_agenda_slides_by_slide(cls, slide):
        '''
        Find agenda master slide and old settings based on current slide, try to find and 
        save selector formatting, then trigger update of all agenda slide.
        '''
        agenda_slides = cls.find_agenda_items_by_slide(slide)
        if len(agenda_slides) == 0:
            bkt.message.warning("Keine Agenda gefunden!", title="Toolbox: Agenda")
            return
        
        # search master slide
        master_slide = None
        for item in agenda_slides:
            if item.position == 0:
                # item with index=0 is master slide
                master_slide = item.slide
                break
        
        # fallback if master slide is not found
        if master_slide == None:
            # no master slide in presentation --> use first agenda slide
            master_slide = cls.restore_master_slide_by_agenda_slide(agenda_slides[0].slide)
            if not master_slide:
                return
        
        # get old settings
        old_settings = cls.get_agenda_settings_from_slide(master_slide)
        selector_fill_color = cls.get_selector_fillcolor_from_settings(old_settings)
        selector_line_color = cls.get_selector_linecolor_from_settings(old_settings)
        selector_text_color = cls.get_selector_textcolor_from_settings(old_settings)
        
        # reset selector pickup before update
        cls.selector_style_pickup = False

        # find first selector and remember color
        for item in agenda_slides:
            if item.position != 0:
                # first item with index!=0 will set selector color
                shp = cls.get_or_create_selector_on_slide(item.slide)
                selector_fill_color["color"]      = [shp.Fill.ForeColor.ObjectThemeColor,float(shp.Fill.ForeColor.Brightness),shp.Fill.ForeColor.RGB]
                selector_fill_color['visibility'] = shp.Fill.Visible

                selector_line_color["color"]      = [shp.Line.ForeColor.ObjectThemeColor,float(shp.Line.ForeColor.Brightness),shp.Line.ForeColor.RGB]
                selector_line_color['visibility'] = shp.Line.Visible
                if shp.Line.Visible:
                    selector_line_color['weight']     = float(shp.Line.Weight)
                    selector_line_color['style']      = shp.Line.Style
                    selector_line_color['dashstyle']  = shp.Line.DashStyle
                
                # active paragraph defines the selector text style
                agenda = cls.get_agenda_textbox_on_slide(item.slide)
                try:
                    active_paragraph = agenda.TextFrame.TextRange.Paragraphs(item.position)
                    selector_text_color['color']      = [active_paragraph.Font.Color.ObjectThemeColor,float(active_paragraph.Font.Color.Brightness),active_paragraph.Font.Color.RGB]
                    selector_text_color['bold']       = active_paragraph.Font.Bold
                    selector_text_color['italic']     = active_paragraph.Font.Italic
                    selector_text_color['underline']  = active_paragraph.Font.Underline
                except:
                    # bkt.helpers.exception_as_message() #only for debug
                    pass
                break

        # == ensure settings on master slide
        # default settings with new id
        new_settings = dict(cls.default_settings) #copy default settings
        new_settings[SETTING_AGENDA_ID] = str(uuid.uuid4())
        new_settings[SETTING_POSITION] = 0
        # overwrite existing settings
        new_settings.update(old_settings)
        new_settings[SETTING_SELECTOR_STYLE_FILL] = selector_fill_color
        new_settings[SETTING_SELECTOR_STYLE_LINE] = selector_line_color
        new_settings[SETTING_SELECTOR_STYLE_TEXT] = selector_text_color
        # old selector definitions for backwards compatibility
        new_settings[SETTING_SELECTOR_FILL_COLOR] = selector_fill_color["color"][2]
        new_settings[SETTING_SELECTOR_LINE_COLOR] = selector_line_color["color"][2]
        # write settings
        cls.write_agenda_settings_to_slide(master_slide, new_settings)
            
        master_textbox = cls.get_agenda_textbox_on_slide(master_slide)
        cls.set_tags_for_textbox(master_textbox) #create tags, i.e. add contextdialog-tag for old agenda textboxes
        cls.update_agenda_slides(master_textbox, agenda_slides=agenda_slides)
        
    

    @classmethod
    def restore_master_slide_by_agenda_slide(cls, agenda_slide):
        '''
        Restore master slide based on agenda slide
        '''
        # show warning-message and ask to continue, if SETTING_HIDE_SUBITEMS=True
        settings = cls.get_agenda_settings(agenda_slide)
        if settings.get(SETTING_HIDE_SUBITEMS, False) == True:
            return_value = bkt.message.confirmation(
                "Agenda-Hauptseite nicht gefunden!\nAgenda kann aus der ersten Agendafolie wiederhergestellt werden, aber versteckte Unterpunkte gehen dabei verloren.\n\nAgenda-Aktualisierung forsetzen?",
                "Toolbox: Agenda", 
                bkt.MessageBox.MB_YESNO,
                bkt.MessageBox.WARNING)
        else:
            return_value = True
        
        if return_value:
            #duplicate first agenda_slide as new master slide
            master_slide = agenda_slide.Duplicate(1)
            master_slide.MoveTo(agenda_slide.SlideIndex)
            master_slide.SlideShowTransition.Hidden = -1
            
            #remove selector from master slide
            shp = cls.get_shape_with_tag_item(master_slide, TOOLBOX_AGENDA_SELECTOR)
            if not shp is None:
                shp.Delete()
            
            #remove text formatting of active paragraph by setting format of last paragraph
            master_textbox = cls.get_agenda_textbox_on_slide(master_slide)
            textrange = master_textbox.TextFrame.TextRange
            last_paragraph = textrange.Paragraphs(textrange.Paragraphs().Count)
            if last_paragraph.Font.Color.ObjectThemeColor == 0:
                textrange.Font.Color.RGB = last_paragraph.Font.Color.RGB
            else:
                textrange.Font.Color.ObjectThemeColor = last_paragraph.Font.Color.ObjectThemeColor
                textrange.Font.Color.Brightness = last_paragraph.Font.Color.Brightness
            textrange.Font.Bold = last_paragraph.Font.Bold
            textrange.Font.Italic = last_paragraph.Font.Italic
            textrange.Font.Underline = last_paragraph.Font.Underline

            return master_slide
        
        else:
            return None



    
    
    @classmethod
    def update_agenda_slides(cls, master_textbox, agenda_slides=None):
        '''
        Update agenda slides based on given agenda master textbox. If agenda slide are not
        provided, find agenda slides by id.
        '''
        master_slide = master_textbox.parent
        settings = cls.get_agenda_settings(master_slide)
        
        # get agenda entries from textbox
        agenda_entries = cls.agenda_entries_from_textbox(master_textbox, slides_for_subitems=settings.get(SETTING_SLIDES_FOR_SUBITEMS) or False)
        # agenda entries from agenda_slides
        if not agenda_slides:
            agenda_slides = cls.find_agenda_items_by_id(master_slide.parent, settings[SETTING_AGENDA_ID])

        # find slides for agenda-entries by text
        for item in agenda_entries:
            for idx in range(len(agenda_slides)):
                slide = agenda_slides[idx].slide
                slide_settings = cls.get_agenda_settings_from_slide(slide)
                if item.text == slide_settings.get(SETTING_TEXT):
                    item.slide = slide
                    del agenda_slides[idx]
                    break
        
        # find slides for agenda-entries by slide-order
        # - For every agenda-entry without corresponding slide,
        #   pick the first agenda-slide, which was not yet assigned to an agenda-entry
        # - This will update agenda-slides for renamed entries - if no entries were reordered or deleted
        for item in agenda_entries:
            if item.slide == None and len(agenda_slides) > 1:
                item.slide = agenda_slides[1].slide #0 is master slide
                del agenda_slides[1]
        
        
        # udpate 1 (first run creates slides)
        last_agenda_index = master_slide.SlideIndex
        for item in agenda_entries:
            par_index, text, _, old_slide = item
            
            if old_slide == master_slide:
                old_slide = None
                new_slide = master_slide
            else:
                # duplicate master_slide
                new_slide = master_slide.Duplicate(1)
                # FIXME: there is a bug in ppt 2016 that "disconnects" the connectors on slide.duplicate even though they appear connected
                # move duplicated slide
                if old_slide:
                    #FIXME: reordering menu items does not reorder agenda slides
                    new_slide.MoveTo(old_slide.SlideIndex)
                    new_slide.SlideShowTransition.Hidden = old_slide.SlideShowTransition.Hidden
                else:
                    new_slide.MoveTo(last_agenda_index+1)
                    new_slide.SlideShowTransition.Hidden = 0
            
            # update agenda
            textbox = cls.get_agenda_textbox_on_slide(new_slide)
            cls.update_agenda_on_slide_new(new_slide, textbox, par_index, settings)
            
            # delete old slide
            if old_slide:
                old_slide.delete()

            # update sections (important: first delete old slide)
            cls.update_agenda_sections_for_slide(new_slide, text, settings)

            # remember slide reference for second run
            item.slide = new_slide

            # remember index for repositioning
            last_agenda_index = new_slide.SlideIndex
        

        # update 2 (second run creates hyperlinks)
        cls.update_hyperlinks_on_slide(master_slide, settings, agenda_entries)
        for item in agenda_entries:
            cls.update_hyperlinks_on_slide(item.slide, settings, agenda_entries)


        # andere loeschen
        for item in agenda_slides:
            if item.slide != master_slide:
                cls.delete_agenda_sections_for_slide(item.slide, settings)
                item.slide.delete()
    
    
    
    @classmethod
    def update_agenda_on_slide_new(cls, slide, textbox, selected_paragraph_index, settings):
        '''
        Update agenda textbox on given slide and set selector style.
        '''
        # slide_settings = dict(cls.default_settings)
        # slide_settings.update(settings)
        # slide_settings[text] = 
        # slide_settings[par_no] = 
        try:
            # Handle hidden subitems in agenda textbox
            if settings.get(SETTING_HIDE_SUBITEMS, False):
                paragraphs = [textbox.TextFrame.TextRange.Paragraphs(idx) for idx in range(1,textbox.TextFrame.TextRange.Paragraphs().Count+1)]
                indent_levels = [p.IndentLevel for p in paragraphs]
                
                # take indent-levels between paragraph and curent paragraph
                # if a paragraph with indent-level 1 is between paragraph and current paragraph
                # and indent-level is >1 then hide the paragraph
                indent_level_1_inbetween = [ 1 in indent_levels[idx:selected_paragraph_index] if idx<selected_paragraph_index-1 else 1 in indent_levels[selected_paragraph_index:idx+1]  for idx in range(len(indent_levels)) ]
                hide_paragraph = [ indent_levels[idx] > 1 and indent_level_1_inbetween[idx] for idx in range(len(indent_levels)) ]
                par_idx_to_hide = [ idx  for idx in range(len(indent_levels)) if hide_paragraph[idx]]
                
                # hide paragraphs
                par_idx_to_hide.reverse()
                for idx in par_idx_to_hide:
                    textbox.TextFrame.TextRange.Paragraphs(idx+1).Delete()
                    if idx == textbox.TextFrame.TextRange.Paragraphs().Count:
                        # deleted last paragraph
                        # textbox.Textframe.TextRange.paragraphs(idx).characters(textbox.Textframe.TextRange.paragraphs(idx).characters().count  ).Delete()
                        textbox.Textframe.TextRange.paragraphs(idx).Delete()
                
                # paragraphs hidden before current
                hidden_before_current = sum(1 for hide in hide_paragraph[0:selected_paragraph_index-1] if hide)
                selected_paragraph_index = selected_paragraph_index - hidden_before_current
            
            # currently selected paragraph
            par = textbox.TextFrame.TextRange.Paragraphs(selected_paragraph_index)
            
            # Tag auf Slide setzen
            slide_settings = dict(settings) #copy settings
            slide_settings[SETTING_POSITION] = selected_paragraph_index
            slide_settings[SETTING_TEXT] = par.text.strip()
            slide_settings[SETTING_INDENT_LEVEL] = par.IndentLevel
            
            
            # # Rand ober- und unterhalb / um wieviel ist Markierung größer als Absatz?
            # selectorMargin = textbox.TextFrame.TextRange.Font.Size * 0.15
            # # Position der Markierung für ersten Eintrag der Agenda
            # selectorTop = textbox.Top + textbox.TextFrame.MarginTop - selectorMargin
            
            # # Position der Markierung bestimmen
            # # Absatz-Höhe und -Abstände pro Absatz addieren
            # for idx in range(1, selected_paragraph_index): # To selected_paragraph_index - 1
            #     # Absaetzhoehe addieren
            #     selectorTop = selectorTop + paragraph_height(textbox.TextFrame.TextRange.Paragraphs(idx), False)
            #     # Absatzabsatand danach
            #     selectorTop = selectorTop + textbox.TextFrame.TextRange.Paragraphs(idx).ParagraphFormat.SpaceAfter
            #     # Absatzabsatand davor vom naechsten Absatz
            #     selectorTop = selectorTop + textbox.TextFrame.TextRange.Paragraphs(idx + 1).ParagraphFormat.SpaceBefore
        
            # selectorHeight = paragraph_height(par, False) + 2 * selectorMargin

            # Selector Größe abhängig von Margin
            selector_rect = cls.get_selector_dimensions_for_margin(settings.get(SETTING_SELECTOR_MARGIN, 0.2), slide, textbox, selected_paragraph_index)

            # Selector (Markierung) aktualisieren
            selector = cls.get_or_create_selector_on_slide(slide)
            selector.Top = selector_rect.top
            selector.Left = selector_rect.left
            selector.Height = selector_rect.height
            selector.width = selector_rect.width

            cls.set_selector_fill(selector.Fill, cls.get_selector_fillcolor_from_settings(settings))
            cls.set_selector_line(selector.Line, cls.get_selector_linecolor_from_settings(settings))
            # selector.Fill.ForeColor.RGB = settings.get(SETTING_SELECTOR_FILL_COLOR) or cls.selectorFillColor
            # selector.Line.ForeColor.RGB = settings.get(SETTING_SELECTOR_LINE_COLOR) or cls.selectorLineColor
            
            # apply previously picked up styles, e.g. gradient or other fancy stuff
            if cls.selector_style_pickup:
                try:
                    selector.Apply()
                except:
                    pass

            # Selector-Font Einstellungen
            selector_paragraph = textbox.TextFrame.TextRange.Paragraphs(selected_paragraph_index)
            cls.set_selector_text(selector_paragraph, cls.get_selector_textcolor_from_settings(settings))
        
            # write settings and set tag
            cls.write_agenda_settings_to_slide(slide, slide_settings)
            cls.set_tags_for_slide(slide, selected_paragraph_index)
            
        except:
            logging.exception("Agenda: agenda update failed")
            bkt.message.error("Fehler beim Aktualisieren der Agenda", title="Toolbox: Agenda")
            # bkt.helpers.exception_as_message()


    @classmethod
    def get_selector_dimensions_for_margin(cls, margin, slide, textbox, selected_paragraph_index):
        par = textbox.TextFrame.TextRange.Paragraphs(selected_paragraph_index)

        # Rand ober- und unterhalb / um wieviel ist Markierung größer als Absatz?
        selectorMargin = textbox.TextFrame.TextRange.Font.Size * margin

        # INFO: previous code (see above) was compatibel with PPT<2010 as Bound* attributes were not available
        selectorTop = par.BoundTop + par.ParagraphFormat.SpaceBefore - selectorMargin
        selectorHeight = par.BoundHeight - par.ParagraphFormat.SpaceBefore - par.ParagraphFormat.SpaceAfter + 2*selectorMargin
        if selected_paragraph_index == 1:
            # first paragraph does not contain space before
            selectorTop -= par.ParagraphFormat.SpaceBefore
            selectorHeight += par.ParagraphFormat.SpaceBefore
        if selected_paragraph_index == textbox.TextFrame.TextRange.Paragraphs().Count:
            # last paragraph does not contain space after
            selectorHeight += par.ParagraphFormat.SpaceAfter
        
        return pplib.BoundingFrame.from_rect(top=selectorTop, height=selectorHeight, left=textbox.Left, width=textbox.Width)


    @classmethod
    def update_agenda_sections_for_slide(cls, slide, text, settings):
        '''
        Update (create or rename) sections for each agenda item
        '''
        try:
            if settings.get(SETTING_CREATE_SECTIONS, False) == False:
                return
            sections = slide.parent.SectionProperties
            section_title = text[:24].strip() #FIXME: test for not allowed characters
            if sections.Count > 0 and sections.FirstSlide(slide.SectionIndex) == slide.SlideIndex: #agenda slide is first slide of section, so rename section, otherwise create new
                sections.Rename(slide.SectionIndex, section_title)
            else:
                sections.AddBeforeSlide(slide.SlideIndex, section_title)
        except:
            logging.exception("Agenda: agenda sections update failed")
            bkt.message.error("Fehler beim Aktualisieren der Agenda-Abschnitte", title="Toolbox: Agenda")
            # bkt.helpers.exception_as_message()


    @classmethod
    def delete_agenda_sections_for_slide(cls, slide, settings):
        '''
        Delete section for old agenda slide if slide is first slide within section
        '''
        try:
            if settings.get(SETTING_CREATE_SECTIONS, False) == False:
                return
            sections = slide.parent.SectionProperties
            if sections.Count > 0 and sections.FirstSlide(slide.SectionIndex) == slide.SlideIndex:
                sections.Delete(slide.SectionIndex, False)
        except:
            logging.exception("Agenda: agenda sections delete failed")
            bkt.message.error("Fehler beim Löschen der Agenda-Abschnitte", title="Toolbox: Agenda")
            # bkt.helpers.exception_as_message()
        
    

    @classmethod
    def update_hyperlinks_on_slide(cls, slide, settings, agenda_entries):
        try:
            textbox = cls.get_agenda_textbox_on_slide(slide)

            # remove all hyperlinks from agenda textbox
            textbox.textframe.textrange.ActionSettings[1].Hyperlink.delete()

            # check settings
            if settings.get(SETTING_CREATE_LINKS, False) == False:
                return

            # link each paragraph
            for paragraph in textbox.textframe.textrange.paragraphs():
                # find correct agenda_entry
                for item in agenda_entries:
                    if item.indentlevel == paragraph.IndentLevel and item.text == paragraph.text.strip():
                        ref_slide = item.slide
                        break
                else:
                    # no entry found, delete any existing hyperlink
                    paragraph.ActionSettings(1).Hyperlink.Delete()
                    continue

                # reference slide found, create hyperlink
                #ActionSettings(1)=ppMouseClick
                paragraph.ActionSettings(1).Hyperlink.SubAddress = "{},{},{}".format(ref_slide.SlideId,ref_slide.SlideIndex,ref_slide.Name)
        
        except:
            logging.exception("Agenda: agenda hyperlink update failed")
            bkt.message.error("Fehler beim Aktualisieren der Agenda-Hyperlinks", title="Toolbox: Agenda")
            # bkt.helpers.exception_as_message()




    
    # =================
    # = ALTE METHODEN =
    # =================
    
    # @staticmethod
    # def get_agenda_slides(context):
    #     '''
    #     return list of all agenda-slides
    #     '''
    #     agenda_slides = {}
        
    #     for sld in context.app.ActiveWindow.Presentation.Slides:
    #         if sld.Tags.Item(TOOLBOX_AGENDA) != "":
    #             agenda_slides["slide-" + str(int(sld.Tags.Item(TOOLBOX_AGENDA_SLIDENO)))] = sld
        
    #     return agenda_slides

    # @classmethod
    # def can_update_agenda(cls, context):
    #     # Aktualisieren oder Löschen nur möglich, wenn es Slides mit Agenda-Meta-Informationen gibt
    #     can_update = len(cls.get_agenda_slides(context)) > 0
    #     return can_update
    
    # @classmethod
    # def create_or_update_agenda(cls, context):
    #     if cls.can_update_agenda(context):
    #         # vorhandene Agenda aktualisieren
    #         cls.update_agenda(context)
    #     else:
    #         cls.create_from_current_slide(context)
    
    # @classmethod
    # def create_from_current_slide(cls, context):
    #     #Dim shp As Shape
    #     #Dim sld As Slide
    #     #Dim answer As Integer
    #
    #     try:
    #         # In normaler Ansicht
    #         if context.app.ActiveWindow.View.Type != ppt.PpViewType.ppViewNormal.value__:
    #             bkt.message("In Slide-View wechseln!")
    #             return
    #
    #         # Aktuelles Slide
    #         slide = context.app.ActiveWindow.View.Slide
    #
    #         # Markiertes Shape oder erstes Shape mit Bullet-Point-Liste auswaehlen
    #         if context.app.ActiveWindow.selection.Type == ppt.PpSelectionType.ppSelectionShapes.value__ or context.app.ActiveWindow.selection.Type == ppt.PpSelectionType.ppSelectionText.value__:
    #             shp = context.app.ActiveWindow.selection.ShapeRange.item(1)
    #         else:
    #             shp = None
    #             for shp in slide.shapes:
    #                 if shp.TextFrame.HasText:
    #                     if shp.TextFrame.TextRange.ParagraphFormat.Bullet.Type == ppt.PpBulletType.ppBulletUnnumbered.value__:
    #                         break
    #
    #             if shp is None:
    #                 bkt.message("Textbox mit Agenda-Einträgen auswählen oder als Bullet-Liste formatieren.")
    #
    #         if not shp is None:
    #             if slide.Tags.Item(TOOLBOX_AGENDA) != "":
    #                 # FIXME: Da create_from_current_slide nur aufgerufen wird, wenn bei create_or_update_agenda kein Update
    #                 # moeglich war, wird dieser Code nie aufgerufen
    #                 cls.remove_agenda_b(context, deleteSlides=False)
    #
    #             cls.create_agenda(context, slide, shp)
    #     except:
    #         bkt.helpers.exception_as_message()
    #
    #
    # @classmethod
    # def create_agenda(cls, context, sld, textbox):
    #     cls.set_tags_for_slide(sld, 0)
    #     cls.set_tags_for_textbox(textbox)
    #     cls.update_agenda_b(context, recreateFromMaster=True)
    
    
    
    # @classmethod
    # def update_agenda(cls, context):
    #     if not cls.can_update_agenda(context):
    #         return
    #     # Neue Kopien der ersten Agenda-Folie erstellen
    #     cls.determine_selector_settings(context)
    #     cls.update_agenda_b(context, recreateFromMaster=True)
    
    # @classmethod
    # def determine_selector_settings(cls, context):
    #     #raise ValueError("some text in determine_selector_settings")
    #     slides = cls.get_agenda_slides(context).values()
    #     slides.sort(key=lambda slide: slide.SlideIndex)
    #
    #     for sld in slides:
    #         shp = cls.get_shape_with_tag_item(sld, TOOLBOX_AGENDA_SELECTOR)
    #         if not shp is None:
    #             cls.selectorFillColor = shp.Fill.ForeColor.RGB
    #             cls.selectorLineColor = shp.Line.ForeColor.RGB
    #             return
    
    # @classmethod
    # def update_agenda_b(cls, context, recreateFromMaster):
    #
    #     try:
    #         # Textbox auf erstem Agenda-Slide holen
    #         newSldCount = 0
    #         slides = cls.get_agenda_slides(context)
    #         if len(slides) == 0:
    #             return
    #
    #         sldMaster = slides.values()[0]
    #         textbox = cls.get_agenda_textbox_on_slide(sldMaster)
    #         if textbox is None:
    #             bkt.message("Update nicht möglich! Agenda-Textbox fehlt auf erstem Agenda-Slide.")
    #
    #         # Pro Absatz das zugehörige Slide aktualisieren (bzw. neu erstellen)
    #         for idx in range(1, textbox.TextFrame.TextRange.Paragraphs().Count+1):
    #             sld = None
    #             try:
    #                 sld = slides["slide-" + str(idx)]
    #             except:
    #                 pass
    #
    #             if textbox.TextFrame.TextRange.Paragraphs(idx).text.strip() != "":
    #                 if sld is None:
    #                     # Slide nicht vorhanden, neu erstellen und nach Master-Agenda-Slide einfuegen
    #                     newSldCount = newSldCount + 1
    #                     sld = sldMaster.Duplicate(1)
    #                     sld.MoveTo(context.app.ActiveWindow.View.Slide.SlideIndex + newSldCount)
    #                     cls.set_tags_for_slide( sld, idx)
    #                 else:
    #                     if recreateFromMaster:
    #                         if sld.SlideID != sldMaster.SlideID:
    #                             # Slide neu erstellen erzwungen, alte Position wird beibehalten
    #                             sld = sldMaster.Duplicate(1)
    #                             sld.MoveTo(slides["slide-" + str(idx)].SlideIndex)
    #                             slides["slide-" + str(idx)].Delete()
    #                             cls.set_tags_for_slide(sld, idx)
    #                     else:
    #                         bkt.message("Not implemented yet!")
    #                     slides.pop("slide-" + str(idx))
    #
    #                 # Agenda aktualisieren
    #                 cls.update_agenda_on_slide(sld, textbox, idx)
    #
    #         try:
    #             slides.pop("slide-0")
    #         except:
    #             pass
    #
    #         for sld in slides.values():
    #             sld.Delete()
    #
    #     except:
    #         bkt.helpers.exception_as_message()
    
    
    # @classmethod
    # def update_agenda_on_slide(cls, sld, textbox, parIdx):
    #     try:
    #         # Tag auf Slide setzen
    #         cls.set_tags_for_slide(sld, parIdx)
    #         # Rand ober- und unterhalb / um wieviel ist Markierung größer als Absatz?
    #         selectorMargin = textbox.TextFrame.TextRange.Font.Size * 0.15
    #         # Position der Markierung für ersten Eintrag der Agenda
    #         selectorTop = textbox.Top + textbox.TextFrame.MarginTop - selectorMargin
    #
    #         # Position der Markierung bestimmen
    #         # Absatz-Höhe und -Abstände pro Absatz addieren
    #         for idx in range(1, parIdx): # To parIdx - 1
    #             # Absaetzhoehe addieren
    #             selectorTop = selectorTop + paragraph_height(textbox.TextFrame.TextRange.Paragraphs(idx), False)
    #             # Absatzabsatand danach
    #             selectorTop = selectorTop + textbox.TextFrame.TextRange.Paragraphs(idx).ParagraphFormat.SpaceAfter
    #             # Absatzabsatand davor vom naechsten Absatz
    #             selectorTop = selectorTop + textbox.TextFrame.TextRange.Paragraphs(idx + 1).ParagraphFormat.SpaceBefore
    #
    #         selectorHeight = paragraph_height(textbox.TextFrame.TextRange.Paragraphs(parIdx), False) + 2 * selectorMargin
    #
    #         # Selector (Markierung) aktualisieren
    #         selector = cls.get_or_create_selector_on_slide(sld)
    #         selector.Top = selectorTop
    #         selector.Left = textbox.Left
    #         selector.Height = selectorHeight
    #         selector.width = textbox.width
    #
    #         # Text und Textbox-Position aktualisieren
    #         oTextBox = cls.get_agenda_textbox_on_slide(sld)
    #
    #         #oTextBox.TextFrame.TextRange.text = textbox.TextFrame.TextRange.text
    #         textbox.TextFrame.TextRange.Copy()
    #         oTextBox.TextFrame.TextRange.Paste()
    #         oTextBox.TextFrame.TextRange.Font.Bold = False
    #         oTextBox.TextFrame.TextRange.Paragraphs(parIdx).Font.Bold = True
    #         oTextBox.Top    = textbox.Top
    #         oTextBox.Left   = textbox.Left
    #         oTextBox.Width  = textbox.Width
    #         oTextBox.Height = textbox.Height
    #
    #     except:
    #         bkt.helpers.exception_as_message()



    # ====================
    # = Agenda entfernen =
    # ====================

    @classmethod
    def remove_agenda(cls, slide, presentation, delete_master_slide=False):
        '''
        remove all agenda-slides of current agenda from the presentation, removing all slides
        '''
        try:
            agenda_slides = cls.find_agenda_items_by_slide(slide)
            if len(agenda_slides) == 0:
                if bkt.message.confirmation("Keine zugehörigen Agenda-Folien gefunden!\nStattdessen alle Agenda-Folie der Präsentation löschen?", title="Toolbox: Agenda"):
                    cls.remove_agendas_from_presentation(presentation, True)
                return

            for item in agenda_slides:
                if item.position == 0 and not delete_master_slide:
                    # Tags von Master-Slide loeschen
                    slide.Tags.Delete(TOOLBOX_AGENDA)
                    slide.Tags.Delete(TOOLBOX_AGENDA_SLIDENO)
                    slide.Tags.Delete(TOOLBOX_AGENDA_SETTINGS)
                    continue
                if item.slide:
                    item.slide.Delete()
        except:
            logging.exception("Agenda: agenda deletion failed")
            bkt.message.error("Fehler beim Löschen der Agenda", title="Toolbox: Agenda")
            # bkt.helpers.exception_as_message()
    
    
    @classmethod
    def remove_agendas_from_presentation(cls, presentation, delete_slides=True):
        '''
        removes all agenda-slides from the presentation, removing all meta-information
        '''
        agenda_slides = []
        
        for slide in presentation.slides:
            # if slide.Tags.Item(TOOLBOX_AGENDA) != "":
            if cls.is_agenda_slide(slide):
                # Tags von Slide loeschen
                slide.Tags.Delete(TOOLBOX_AGENDA)
                slide.Tags.Delete(TOOLBOX_AGENDA_SLIDENO)
                slide.Tags.Delete(TOOLBOX_AGENDA_SETTINGS)
                
                # Tags von Shapes auf Slide loeschen
                shape = cls.get_agenda_textbox_on_slide(slide)
                if not shape is None:
                    shape.Tags.Delete(TOOLBOX_AGENDA_TEXTBOX)
                shape = cls.get_agenda_textbox_on_slide(slide)
                if not shape is None:
                    shape.Tags.Delete(TOOLBOX_AGENDA_SELECTOR)
            
                if delete_slides:
                    agenda_slides.append(slide)

        if delete_slides:
            for slide in agenda_slides:
                slide.Delete()




    # ===================
    # = agenda settings =
    # ===================

    @classmethod
    def get_agenda_settings(cls, slide):
        ''' load agenda settings from current slide or presentation '''
        settings = cls.get_agenda_settings_from_slide(slide)
        
        if settings == {}:
            for slide in slide.parent.slides:
                settings = cls.get_agenda_settings_from_slide(slide)
                if settings != None:
                    cls.settings = settings
                    return settings
            return None
            
        else:
            cls.settings = settings
            return settings
    
    
    @classmethod
    def get_agenda_settings_from_slide(cls, slide):
        ''' load agenda settings from given slide '''
        value = cls.get_tag_value(slide, TOOLBOX_AGENDA_SETTINGS, "{}" )
        settings = json.loads( value ) or {}

        #convert old settings with rgb color to new color format (but remain old settings)
        if SETTING_SELECTOR_STYLE_FILL not in settings and type(settings.get(SETTING_SELECTOR_FILL_COLOR)) == int:
            settings[SETTING_SELECTOR_STYLE_FILL] = cls.default_selectorFillColor.copy()
            settings[SETTING_SELECTOR_STYLE_FILL]['color'] = [0, 0, settings.get(SETTING_SELECTOR_FILL_COLOR)]
        
        if SETTING_SELECTOR_STYLE_LINE not in settings and type(settings.get(SETTING_SELECTOR_LINE_COLOR)) == int:
            settings[SETTING_SELECTOR_STYLE_LINE] = cls.default_selectorLineColor.copy()
            settings[SETTING_SELECTOR_STYLE_LINE]['color'] = [0, 0, settings.get(SETTING_SELECTOR_LINE_COLOR)]

        return settings
    
    
    @staticmethod
    def write_agenda_settings_to_slide(slide, settings):
        ''' write agenda settings to slide tags '''
        slide.Tags.Add(TOOLBOX_AGENDA_SETTINGS, json.dumps(settings, ensure_ascii=False) )
    
    
    @classmethod
    def update_agenda_settings_on_slide(cls, slide, settings):
        ''' update agenda settings of current slide to slide tags '''
        current_settings = cls.get_agenda_settings_from_slide(slide)
        current_settings.update(settings)
        cls.write_agenda_settings_to_slide(slide, current_settings)
    
    
    @classmethod
    def get_hide_subitems(cls, slide):
        ''' reads setting hide_subitems from current slide or presentation'''
        settings = cls.get_agenda_settings_from_slide(slide)
        return settings.get(SETTING_HIDE_SUBITEMS) or False
    
    
    @classmethod
    def get_create_sections(cls, slide):
        ''' reads setting sections-create from current slide or presentation'''
        settings = cls.get_agenda_settings_from_slide(slide)
        return settings.get(SETTING_CREATE_SECTIONS) or False
    
    
    @classmethod
    def get_create_links(cls, slide):
        ''' reads setting link-create from current slide or presentation'''
        settings = cls.get_agenda_settings_from_slide(slide)
        return settings.get(SETTING_CREATE_LINKS) or False
    

    @classmethod
    def get_slides_for_subitems(cls, slide):
        ''' reads setting hide_subitems from current slide or presentation'''
        settings = cls.get_agenda_settings_from_slide(slide)
        return settings.get(SETTING_SLIDES_FOR_SUBITEMS) or False
    

    @classmethod
    def get_selector_margin(cls, slide):
        settings = cls.get_agenda_settings_from_slide(slide)
        return settings.get(SETTING_SELECTOR_MARGIN, 0.2)
    

    @classmethod
    def get_selector_fillcolor_from_settings(cls, settings):
        cls.selectorFillColor.update(settings.get(SETTING_SELECTOR_STYLE_FILL, {}))
        return cls.selectorFillColor
    

    @classmethod
    def get_selector_linecolor_from_settings(cls, settings):
        cls.selectorLineColor.update(settings.get(SETTING_SELECTOR_STYLE_LINE, {}))
        return cls.selectorLineColor
    

    @classmethod
    def get_selector_textcolor_from_settings(cls, settings):
        cls.selectorTextColor.update(settings.get(SETTING_SELECTOR_STYLE_TEXT, {}))
        return cls.selectorTextColor
    
    
    # ========================
    # = save/load tag values =
    # ========================
    
    @staticmethod
    def set_tags_for_slide(sld, slideNo=None):
        ''' Meta-Informationen für Slide einstellen '''
        sld.Tags.Add(TOOLBOX_AGENDA, "1")
        if slideNo is not None:
            sld.Tags.Add(TOOLBOX_AGENDA_SLIDENO, str(slideNo))

    @staticmethod
    def set_tags_for_textbox(textbox):
        ''' Meta-Informationen für Textbox einstellen '''
        textbox.Tags.Add(TOOLBOX_AGENDA_TEXTBOX, "1")
        textbox.Tags.Add(bkt.contextdialogs.BKT_CONTEXTDIALOG_TAGKEY, TOOLBOX_AGENDA_POPUP)

    @staticmethod
    def set_tags_for_selector(selector):
        ''' Meta-Informationen für Selector-Shape einstellen '''
        selector.Tags.Add(TOOLBOX_AGENDA_SELECTOR, "1")

    @staticmethod
    def get_shape_with_tag_item(sld, tagKey):
        ''' Shape auf Slide finden, das einen bestimmten TagKey enthaelt '''
        for shp in sld.shapes:
            if shp.Tags.Item(tagKey) != "":
                return shp 
        return None

    @staticmethod
    def get_tag_value(obj, tagname, default=''):
        ''' load tag-value from object or return default '''
        for idx in range(1,obj.tags.count+1):
            if obj.tags.name(idx) == tagname:
                return obj.tags.value(idx)
        return default
    
    
    
    # =============================
    # = callback-specific methods =
    # =============================
    
    @classmethod
    def is_agenda_slide(cls, slide):
        ''' check if current slide is agenda-slide '''
        try:
            return slide.Tags.Item(TOOLBOX_AGENDA_SLIDENO) != ""
        #     if slide.Tags.Item(TOOLBOX_AGENDA_SLIDENO) != "":
        #         textbox = cls.get_agenda_textbox_on_slide(slide)
        #         return textbox != None
        #     else:
        #         return False
        except: #AttributeError
            return False
        # settings = cls.get_agenda_settings_from_slide(slide)
        # return settings != {}
    
    @classmethod
    def can_create_agenda_from_slide(cls, slide):
        ''' check if agenda textbox is on slide in order to create agenda from textbox '''
        try:
            if slide.Tags.Item(TOOLBOX_AGENDA) != "":
                textbox = cls.get_agenda_textbox_on_slide(slide)
                return textbox != None
            else:
                return False
        except: #AttributeError
            return False
    
    @classmethod
    def set_hide_subitems(cls, slide, pressed):
        ''' callback to write hide_subitems setting to all agenda slides '''
        cls._update_setting_value(slide, SETTING_HIDE_SUBITEMS, pressed)
    
    @classmethod
    def set_create_sections(cls, slide, pressed):
        ''' callback to write sections-create setting to all agenda slides '''
        cls._update_setting_value(slide, SETTING_CREATE_SECTIONS, pressed)
    
    @classmethod
    def set_create_links(cls, slide, pressed):
        ''' callback to write links-create setting to all agenda slides '''
        cls._update_setting_value(slide, SETTING_CREATE_LINKS, pressed)

    @classmethod
    def set_slides_for_subitems(cls, slide, pressed):
        ''' callback to write slides_for_subitmes setting to all agenda slides '''
        cls._update_setting_value(slide, SETTING_SLIDES_FOR_SUBITEMS, pressed)
    
    @classmethod
    def _update_setting_value(cls, slide, key, value):
        ''' callback to write settings value to all agenda slides '''
        agenda_items = cls.find_agenda_items_by_slide(slide)
        # settings = cls.get_agenda_settings_from_slide(agenda_items[0].slide)
        # settings[key] = value
        for agenda_item in agenda_items:
            cls.update_agenda_settings_on_slide(agenda_item.slide, {key: value})
        
        if bkt.message.confirmation("Agenda jetzt aktualisieren?", title="Toolbox: Agenda"):
            cls.update_agenda_slides_by_slide(slide)
    
    
    
    # =========================
    # = color gallery methods =
    # =========================
    
    @classmethod
    def get_selector_fillcolor(cls, slide, key="color"):
        try:
            settings = cls.get_agenda_settings_from_slide(slide)
            return settings.get(SETTING_SELECTOR_STYLE_FILL)[key]
        except:
            return cls.selectorFillColor[key]
    
    @classmethod
    def get_selector_linecolor(cls, slide, key="color"):
        try:
            settings = cls.get_agenda_settings_from_slide(slide)
            return settings.get(SETTING_SELECTOR_STYLE_LINE)[key]
        except:
            return cls.selectorLineColor[key]
    
    @classmethod
    def get_selector_textcolor(cls, slide, key="color"):
        try:
            settings = cls.get_agenda_settings_from_slide(slide)
            return settings.get(SETTING_SELECTOR_STYLE_TEXT)[key]
        except:
            return cls.selectorTextColor[key]
    
    @classmethod
    def set_selector_fillcolor_rgb(cls, color, slide, visibility=-1):
        color_dict = {
            'color': [0,0,color],
            'visibility': visibility,
        }
        cls.set_selector_fillcolor(color_dict, slide)
    
    @classmethod
    def set_selector_linecolor_rgb(cls, color, slide, visibility=-1):
        color_dict = {
            'color': [0,0,color],
            'visibility': visibility,
        }
        cls.set_selector_linecolor(color_dict, slide)
    
    @classmethod
    def set_selector_textcolor_rgb(cls, color, slide):
        color_dict = {
            'color': [0,0,color],
        }
        cls.set_selector_textcolor(color_dict, slide)
    
    @classmethod
    def set_selector_fillcolor_theme(cls, color_index, brightness, slide, visibility=-1):
        color_dict = {
            'color': [color_index,brightness,0],
            'visibility': visibility,
        }
        cls.set_selector_fillcolor(color_dict, slide)
    
    @classmethod
    def set_selector_linecolor_theme(cls, color_index, brightness, slide, visibility=-1):
        color_dict = {
            'color': [color_index,brightness,0],
            'visibility': visibility,
        }
        cls.set_selector_linecolor(color_dict, slide)
    
    @classmethod
    def set_selector_textcolor_theme(cls, color_index, brightness, slide):
        color_dict = {
            'color': [color_index,brightness,0],
        }
        cls.set_selector_textcolor(color_dict, slide)
    
    @classmethod
    def reset_selector_fillcolor(cls, slide):
        # cls.set_selector_fillcolor_rgb(12566463, slide)
        cls.set_selector_fillcolor(cls.default_selectorFillColor, slide)
    
    @classmethod
    def reset_selector_linecolor(cls, slide):
        # cls.set_selector_linecolor_rgb(8355711, slide)
        cls.set_selector_linecolor(cls.default_selectorLineColor, slide)
    
    @classmethod
    def reset_selector_textcolor(cls, slide):
        cls.set_selector_textcolor(cls.default_selectorTextColor, slide)
    
    @classmethod
    def hide_selector_fill(cls, slide):
        cls.set_selector_fillcolor({'visibility': 0}, slide)
    
    @classmethod
    def hide_selector_line(cls, slide):
        cls.set_selector_linecolor({'visibility': 0}, slide)
    
    @classmethod
    def toggle_selector_text_style(cls, slide, style="bold"):
        cls.set_selector_textcolor({style: not cls.get_selector_textcolor(slide, style)}, slide)
    
    # ===============================
    # = selector adjustment methods =
    # ===============================
    
    @classmethod
    def set_selector_margin(cls, margin, slide):
        # cls._update_setting_value(slide, SETTING_SELECTOR_MARGIN, margin)
        # if bkt.message.confirmation("Agenda jetzt aktualisieren?", title="Toolbox: Agenda"):
        #     cls.update_agenda_slides_by_slide(slide)
        
        agenda_items = cls.find_agenda_items_by_slide(slide)
        # update slides
        for agenda_item in agenda_items:
            cls.update_agenda_settings_on_slide(agenda_item.slide, {SETTING_SELECTOR_MARGIN: margin})
            #recalculate selector dimensions
            try:
                agenda = cls.get_agenda_textbox_on_slide(agenda_item.slide)
                selector_rect = cls.get_selector_dimensions_for_margin(margin, agenda_item.slide, agenda, agenda_item.position)

                shp = cls.get_shape_with_tag_item(agenda_item.slide, TOOLBOX_AGENDA_SELECTOR)
                shp.top = selector_rect.top
                shp.left = selector_rect.left
                shp.height = selector_rect.height
                shp.width = selector_rect.width
            except:
                continue

    @classmethod
    def set_selector_fillcolor(cls, color_dict, slide):
        agenda_items = cls.find_agenda_items_by_slide(slide)
        #get and update settings
        settings = cls.get_agenda_settings_from_slide(agenda_items[0].slide)
        #extract only fill style from settings from master-agenda-slide
        settings = { SETTING_SELECTOR_STYLE_FILL: cls.get_selector_fillcolor_from_settings(settings) }
        settings[SETTING_SELECTOR_STYLE_FILL].update(color_dict)
        # old selector definitions for backwards compatibility
        if 'color' in color_dict:
            settings[SETTING_SELECTOR_FILL_COLOR] = color_dict['color'][2]
        # update slides
        for agenda_item in agenda_items:
            # cls.write_agenda_settings_to_slide(agenda_item.slide, settings)
            cls.update_agenda_settings_on_slide(agenda_item.slide, settings)
            #recolor each selector right away
            shp = cls.get_shape_with_tag_item(agenda_item.slide, TOOLBOX_AGENDA_SELECTOR)
            if not shp is None:
                cls.set_selector_fill(shp.Fill, settings[SETTING_SELECTOR_STYLE_FILL])
    
    @classmethod
    def set_selector_linecolor(cls, color_dict, slide):
        agenda_items = cls.find_agenda_items_by_slide(slide)
        #get and update settings
        settings = cls.get_agenda_settings_from_slide(agenda_items[0].slide)
        #extract only fill style from settings from master-agenda-slide
        settings = { SETTING_SELECTOR_STYLE_LINE: cls.get_selector_linecolor_from_settings(settings) }
        settings[SETTING_SELECTOR_STYLE_LINE].update(color_dict)
        # old selector definitions for backwards compatibility
        if 'color' in color_dict:
            settings[SETTING_SELECTOR_LINE_COLOR] = color_dict['color'][2]
        # update slides
        for agenda_item in agenda_items:
            # cls.write_agenda_settings_to_slide(agenda_item.slide, settings)
            cls.update_agenda_settings_on_slide(agenda_item.slide, settings)
            #recolor each selector right away
            shp = cls.get_shape_with_tag_item(agenda_item.slide, TOOLBOX_AGENDA_SELECTOR)
            if not shp is None:
                cls.set_selector_line(shp.Line, settings[SETTING_SELECTOR_STYLE_LINE])
    
    @classmethod
    def set_selector_textcolor(cls, color_dict, slide):
        agenda_items = cls.find_agenda_items_by_slide(slide)
        #get and update settings
        settings = cls.get_agenda_settings_from_slide(agenda_items[0].slide)
        #extract only fill style from settings from master-agenda-slide
        settings = { SETTING_SELECTOR_STYLE_TEXT: cls.get_selector_textcolor_from_settings(settings) }
        settings[SETTING_SELECTOR_STYLE_TEXT].update(color_dict)
        # update slides
        for agenda_item in agenda_items:
            # cls.write_agenda_settings_to_slide(agenda_item.slide, settings)
            cls.update_agenda_settings_on_slide(agenda_item.slide, settings)
            #check for master agenda textbox
            if agenda_item.position == 0:
                continue
            #recolor each selector right away
            agenda = cls.get_agenda_textbox_on_slide(agenda_item.slide)
            if not agenda is None:
                try:
                    active_paragraph = agenda.TextFrame.TextRange.Paragraphs(agenda_item.position)
                    cls.set_selector_text(active_paragraph, settings[SETTING_SELECTOR_STYLE_TEXT])
                except:
                    continue
    
    @classmethod
    def set_selector_fill(cls, shape_fill_obj, color_dict):
        shape_fill_obj.Solid() #default shape might have gradient or other non-solid background
        if 'visibility' in color_dict:
            shape_fill_obj.Visible = color_dict["visibility"]
        #only if visible:
        if shape_fill_obj.Visible and 'color' in color_dict:
            color_list = color_dict["color"]
            if color_list[0] == 0:
                shape_fill_obj.ForeColor.RGB = color_list[2]
            else:
                shape_fill_obj.ForeColor.ObjectThemeColor = color_list[0]
                shape_fill_obj.ForeColor.Brightness = color_list[1]
    
    @classmethod
    def set_selector_line(cls, shape_line_obj, color_dict):
        if 'visibility' in color_dict:
            shape_line_obj.Visible = color_dict["visibility"]
        #only if visible:
        if shape_line_obj.Visible:
            if 'color' in color_dict:
                color_list = color_dict["color"]
                if color_list[0] == 0:
                    shape_line_obj.ForeColor.RGB = color_list[2]
                else:
                    shape_line_obj.ForeColor.ObjectThemeColor = color_list[0]
                    shape_line_obj.ForeColor.Brightness = color_list[1]
            if 'weight' in color_dict:
                shape_line_obj.Weight     = color_dict["weight"]
            if 'style' in color_dict:
                shape_line_obj.Style      = color_dict["style"]
            if 'dashstyle' in color_dict:
                shape_line_obj.DashStyle  = color_dict["dashstyle"]
    
    @classmethod
    def set_selector_text(cls, shape_textrange_obj, color_dict):
        if 'color' in color_dict:
            color_list = color_dict["color"]
            if color_list[0] == 0:
                shape_textrange_obj.Font.Color.RGB = color_list[2]
            else:
                shape_textrange_obj.Font.Color.ObjectThemeColor = color_list[0]
                shape_textrange_obj.Font.Color.Brightness = color_list[1]
        if 'bold' in color_dict:
            shape_textrange_obj.Font.Bold      = color_dict["bold"]
        if 'italic' in color_dict:
            shape_textrange_obj.Font.Italic    = color_dict["italic"]
        if 'underline' in color_dict:
            shape_textrange_obj.Font.Underline = color_dict["underline"]





# def paragraph_height(par, with_par_spaces=True):
#     # Absatzhoehe bestimmen
#     parHeight = par.Lines().Count * line_height(par)
#     if with_par_spaces:
#         parHeight = parHeight + max(0, par.ParagraphFormat.SpaceBefore) + max(0, par.ParagraphFormat.SpaceAfter)
    
#     return parHeight

# def line_height(par):
#     if par.ParagraphFormat.LineRuleWithin:
#         # spacing = number of lines
#         # Annahme zur Korrektur der Abstände: Abstand zwischen zwei Zeilen ist 0.2pt
#         return par.Font.Size * (max(0, par.ParagraphFormat.SpaceWithin) + 0.2)
#     else:
#         # spacing = number of pt
#         # Annahme zur Korrektur der Abstände: Abstand zwischen zwei Zeilen ist 0.2pt
#         return par.ParagraphFormat.SpaceWithin #+ 0.1 * .Font.Size





# Menu
#   Neue Agenda erstellen --> macht immer neue Id
#   Agenda aktualisieren --> nur verfügbar, wenn Id gefunden
#   Remove Agenda-Information from slide
#   Remove Agenda-Information from presentation
#   Show/Hide Agenda-Overview
#
# Einstellungen
# {
#   id
#   hide-agenda-overview
#   hide other sub-agenda
#   tracker-color
#   ...
# }
# 
# Vorhandene Agenda-Slides finden --> [ [Nr, Text, Slide], ...]
# Neue Agenda-Slides bestimmen --> [ [Nr, Text, Indent, Slide], ...]
# 
# Mapping neu auf vorhanden
#   for a in new_agenda:
#       take first that matches text
#       mapped_b = [b for b in current_agenda if b[1] == a[1]][1]
#       remove mapped_b from current_agenda
#

agendamenu = bkt.ribbon.Menu(
    label="Agenda",
    children=[
        bkt.ribbon.Button(
            id='add-agenda-textbox',
            label="Agenda-Textbox einfügen",
            supertip="Standard Agenda-Textbox einfügen, um daraus eine aktualisierbare Agenda zu generieren.",
            imageMso="TextBoxInsert",
            on_action=bkt.Callback(ToolboxAgenda.create_agenda_textbox_on_slide)
        ),
        bkt.ribbon.Button(
            id='agenda-new-create',
            label="Agenda neu erstellen",
            supertip="Neue Agenda auf Basis der aktuellen Folie erstellen. Aktuelle Folien wird Hauptfolie der Agenda.",
            imageMso="TableOfContentsAddTextGallery",
            on_action=bkt.Callback(ToolboxAgenda.create_agenda_from_slide),
            get_enabled=bkt.Callback(ToolboxAgenda.can_create_agenda_from_slide)
        ),
        bkt.ribbon.MenuSeparator(),
        bkt.ribbon.Button(
            id='agenda-new-update',
            label="Agenda aktualisieren",
            supertip="Agenda aktualisieren und durch Agenda auf dem Agenda-Hauptfolie ersetzen; Folien werden dabei neu erstellt.",
            imageMso="SaveSelectionToTableOfContentsGallery",
            on_action=bkt.Callback(ToolboxAgenda.update_agenda_slides_by_slide),
            get_enabled=bkt.Callback(ToolboxAgenda.is_agenda_slide)
        ),
        bkt.ribbon.Menu(
            id='agenda-options-menu',
            label="Agenda-Einstellungen",
            children=[
                bkt.ribbon.ToggleButton(
                    id='agenda-slide-for-subitems',
                    label="Agenda-Slides für Unterpunkte",
                    supertip="Für Unterpunkte eines Agendapunkts (Indent-Level>1) werden Agenda-Slides erstellt",
                    on_toggle_action=bkt.Callback(ToolboxAgenda.set_slides_for_subitems),
                    get_pressed=bkt.Callback(ToolboxAgenda.get_slides_for_subitems),
                    get_enabled=bkt.Callback(ToolboxAgenda.is_agenda_slide)
                ),
                bkt.ribbon.ToggleButton(
                    id='agenda-hide-subitems',
                    label="Andere Agenda-Unterpunkte ausblenden",
                    supertip="Unterpunkte eines Agendapunkts (Indent-Level>1) werden in den anderen Abschnitten ausgeblendet",
                    on_toggle_action=bkt.Callback(ToolboxAgenda.set_hide_subitems),
                    get_pressed=bkt.Callback(ToolboxAgenda.get_hide_subitems),
                    get_enabled=bkt.Callback(ToolboxAgenda.is_agenda_slide)
                ),
                bkt.ribbon.ToggleButton(
                    id='agenda-create-sections',
                    label="Abschnitte für Agenda-Punkte erstellen",
                    supertip="Einen neuen Abschnitt je Agenda-Folie beginnen.",
                    on_toggle_action=bkt.Callback(ToolboxAgenda.set_create_sections),
                    get_pressed=bkt.Callback(ToolboxAgenda.get_create_sections),
                    get_enabled=bkt.Callback(ToolboxAgenda.is_agenda_slide)
                ),
                bkt.ribbon.ToggleButton(
                    id='agenda-create-links',
                    label="Hyperlinks für Agenda-Punkte erstellen",
                    supertip="Jeden Agenda-Punkt mit der zugehörigen Agenda-Folie verlinken.",
                    on_toggle_action=bkt.Callback(ToolboxAgenda.set_create_links),
                    get_pressed=bkt.Callback(ToolboxAgenda.get_create_links),
                    get_enabled=bkt.Callback(ToolboxAgenda.is_agenda_slide)
                ),
            ]
        ),
        bkt.ribbon.MenuSeparator(),
        bkt.ribbon.Button(
            id='agenda-remove',
            label="Agenda-Folie entfernen",
            supertip="Entfernt Agenda-Folien der gewählten Agenda, alle Meta-Informationen werden gelöscht.",
            imageMso="TableOfContentsRemove",
            on_action=bkt.Callback(ToolboxAgenda.remove_agenda),
            get_enabled=bkt.Callback(ToolboxAgenda.is_agenda_slide)
        ),
        bkt.ribbon.Button(
            id='agenda-remove-all',
            label="Alle Agenden aus Präsentation entfernen",
            supertip="Entfernt alle Agenda-Folien in der ganzen Präsentation, alle Meta-Informationen werden gelöscht.",
            imageMso="TableOfContentsRemove",
            on_action=bkt.Callback(ToolboxAgenda.remove_agendas_from_presentation),
            get_enabled=bkt.Callback(ToolboxAgenda.presentation_has_agenda)
        ),
        # bkt.ribbon.Menu(
        #     label="Old",
        #     children=[
        #         bkt.ribbon.Button(
        #             id='agenda-update',
        #             label="Agenda aktualisieren",
        #             screentip="Bestehende Agenda aktualisieren (auf Basis erstes Agenda-Slide) oder Agenda-Slides aus aktueller Folie erstellen.",
        #             imageMso="GroupAddInsMenuCommands",
        #             on_action=bkt.Callback(ToolboxAgenda.create_or_update_agenda)
        #         ),
        #     ]
        # )
        
    ]
)

agenda_tab = bkt.ribbon.Tab(
    id = "bkt_context_tab_agenda",
    label = "[BKT] Agenda",
    # get_visible=bkt.Callback(ToolboxAgenda.is_agenda_slide, slide=True),
    get_visible=bkt.Callback(ToolboxAgenda.can_create_agenda_from_slide, slide=True),
    children = [
        bkt.ribbon.Group(
            id="bkt_agenda_manual",
            label = "Anleitung",
            children = [
                bkt.ribbon.Label(label='Schritt 1: Textbox mit Agenda füllen und "Agenda neu erstellen"'),
                bkt.ribbon.Label(label='Schritt 2: Nach jeder weiteren Änderung "Agenda aktualisieren"'),
                bkt.ribbon.Label(label='Hinweis: Agenda-Hauptfolie sollte nicht gelöscht werden!'),
            ]
        ),
        bkt.ribbon.Group(
            id="bkt_agenda_group",
            label = "Agenda",
            children = [
                bkt.ribbon.Button(
                    id='agenda_new_create',
                    label="Agenda neu erstellen",
                    size="large",
                    supertip="Neue Agenda auf Basis des aktuellen Slides erstellen. Aktuelles Slide wird Hauptfolie der Agenda.",
                    imageMso="TableOfContentsAddTextGallery",
                    on_action=bkt.Callback(ToolboxAgenda.create_agenda_from_slide),
                    get_enabled=bkt.Callback(ToolboxAgenda.can_create_agenda_from_slide)
                ),
                bkt.ribbon.Button(
                    id='agenda_new_update',
                    label="Agenda aktualisieren",
                    size="large",
                    supertip="Agenda aktualisieren und durch Agenda auf dem Agenda-Hauptfolie ersetzen; Folien werden dabei neu erstellt.",
                    imageMso="SaveSelectionToTableOfContentsGallery",
                    on_action=bkt.Callback(ToolboxAgenda.update_agenda_slides_by_slide),
                    get_enabled=bkt.Callback(ToolboxAgenda.is_agenda_slide)
                ),
                bkt.ribbon.Separator(),
                bkt.ribbon.Menu(
                    label="Optionen",
                    screentip="Agenda-Optionen",
                    supertip="Verschiedene Agenda-Optionen ändern",
                    imageMso="TableProperties",
                    size="large",
                    get_enabled=bkt.Callback(ToolboxAgenda.is_agenda_slide),
                    children=[
                        bkt.ribbon.ToggleButton(
                            id='agenda_slide_for_subitems',
                            label="Agenda-Slides für Unterpunkte",
                            supertip="Für Unterpunkte eines Agendapunkts (Indent-Level>1) werden Agenda-Slides erstellt",
                            on_toggle_action=bkt.Callback(ToolboxAgenda.set_slides_for_subitems),
                            get_pressed=bkt.Callback(ToolboxAgenda.get_slides_for_subitems),
                            # get_enabled=bkt.Callback(ToolboxAgenda.is_agenda_slide)
                        ),
                        bkt.ribbon.ToggleButton(
                            id='agenda_hide_subitems',
                            label="Andere Agenda-Unterpunkte ausblenden",
                            supertip="Unterpunkte eines Agendapunkts (Indent-Level>1) werden in den anderen Abschnitten ausgeblendet",
                            on_toggle_action=bkt.Callback(ToolboxAgenda.set_hide_subitems),
                            get_pressed=bkt.Callback(ToolboxAgenda.get_hide_subitems),
                            # get_enabled=bkt.Callback(ToolboxAgenda.is_agenda_slide)
                        ),
                        bkt.ribbon.ToggleButton(
                            id='agenda_create_sections',
                            label="Abschnitte für Agenda-Punkte erstellen",
                            supertip="Einen neuen Abschnitt je Agenda-Folie beginnen.",
                            on_toggle_action=bkt.Callback(ToolboxAgenda.set_create_sections),
                            get_pressed=bkt.Callback(ToolboxAgenda.get_create_sections),
                            # get_enabled=bkt.Callback(ToolboxAgenda.is_agenda_slide)
                        ),
                        bkt.ribbon.ToggleButton(
                            id='agenda_create_links',
                            label="Hyperlinks für Agenda-Punkte erstellen",
                            supertip="Jeden Agenda-Punkt mit der zugehörigen Agenda-Folie verlinken.",
                            on_toggle_action=bkt.Callback(ToolboxAgenda.set_create_links),
                            get_pressed=bkt.Callback(ToolboxAgenda.get_create_links),
                            # get_enabled=bkt.Callback(ToolboxAgenda.is_agenda_slide)
                        ),
                    ]
                ),
                bkt.ribbon.Separator(),
                bkt.ribbon.Button(
                    id='agenda_remove',
                    label="Agenda-Folien entfernen",
                    size="large",
                    screentip="Alle zugehörigen Agenda-Folien entfernen",
                    supertip="Entfernt alle Agenda-Folien, die zur aktuellen Agenda gehören, außer der Hauptfolie. Alle Meta-Informationen werden gelöscht.",
                    imageMso="TableOfContentsRemove",
                    on_action=bkt.Callback(ToolboxAgenda.remove_agenda),
                    get_enabled=bkt.Callback(ToolboxAgenda.is_agenda_slide)
                ),
            ]
        ),
        bkt.ribbon.Group(
            id="bkt_agenda_selector_group",
            label = "Agenda Selektor",
            children = [
                bkt.ribbon.ColorGallery(
                    label = 'Hintergrund ändern',
                    size="large",
                    image_mso = 'ShapeFillColorPicker',
                    screentip="Hintergrundfarbe für Selektor",
                    supertip="Passe die Hintergrundfarbe für den Selektor, der den aktiven Agendapunkt hervorhebt, an.",
                    on_rgb_color_change   = bkt.Callback(ToolboxAgenda.set_selector_fillcolor_rgb),
                    on_theme_color_change = bkt.Callback(ToolboxAgenda.set_selector_fillcolor_theme),
                    get_selected_color    = bkt.Callback(ToolboxAgenda.get_selector_fillcolor),
                    get_enabled=bkt.Callback(ToolboxAgenda.is_agenda_slide),
                    children=[
                        bkt.ribbon.Button(
                            label="Keine Füllung",
                            supertip="Selektor-Hintergrund auf transparent ändern",
                            on_action=bkt.Callback(ToolboxAgenda.hide_selector_fill),
                            get_image=bkt.Callback(lambda slide: bkt.ribbon.Gallery.get_check_image(ToolboxAgenda.get_selector_fillcolor(slide, "visibility") == 0)),
                        ),
                        bkt.ribbon.Button(
                            label="Zurücksetzen",
                            screentip="Selektor-Hintergrund zurücksetzen",
                            supertip="Selektor-Hintergrund auf Standard zurücksetzen",
                            on_action=bkt.Callback(ToolboxAgenda.reset_selector_fillcolor)
                        ),
                    ]
                ),
                bkt.ribbon.ColorGallery(
                    label = 'Rahmen ändern',
                    size="large",
                    image_mso = 'ShapeOutlineColorPicker',
                    screentip="Linienfarbe für Selektor",
                    supertip="Passe die Linienfarbe für den Selektor, der den aktiven Agendapunkt hervorhebt, an.",
                    on_rgb_color_change   = bkt.Callback(ToolboxAgenda.set_selector_linecolor_rgb),
                    on_theme_color_change = bkt.Callback(ToolboxAgenda.set_selector_linecolor_theme),
                    get_selected_color    = bkt.Callback(ToolboxAgenda.get_selector_linecolor),
                    get_enabled=bkt.Callback(ToolboxAgenda.is_agenda_slide),
                    children=[
                        bkt.ribbon.Button(
                            label="Kein Rahmen",
                            supertip="Selektor-Rahmen auf transparent ändern",
                            on_action=bkt.Callback(ToolboxAgenda.hide_selector_line),
                            get_image=bkt.Callback(lambda slide: bkt.ribbon.Gallery.get_check_image(ToolboxAgenda.get_selector_linecolor(slide, "visibility") == 0)),
                        ),
                        bkt.ribbon.Button(
                            label="Zurücksetzen",
                            screentip="Selektor-Rahmen zurücksetzen",
                            supertip="Selektor-Rahmen auf Standard zurücksetzen",
                            on_action=bkt.Callback(ToolboxAgenda.reset_selector_linecolor)
                        ),
                    ]
                ),
                bkt.ribbon.ColorGallery(
                    label = 'Text ändern',
                    size="large",
                    image_mso = 'TextFillColorPicker',
                    screentip="Textfarbe für Selektor",
                    supertip="Passe die Textfarbe für den Selektor, der den aktiven Agendapunkt hervorhebt, an.",
                    on_rgb_color_change   = bkt.Callback(ToolboxAgenda.set_selector_textcolor_rgb),
                    on_theme_color_change = bkt.Callback(ToolboxAgenda.set_selector_textcolor_theme),
                    get_selected_color    = bkt.Callback(ToolboxAgenda.get_selector_textcolor),
                    get_enabled=bkt.Callback(ToolboxAgenda.is_agenda_slide),
                    children=[
                        bkt.ribbon.Button(
                            label="Fett",
                            screentip="Selektor-Text fett",
                            supertip="Selektor-Text fett darstellen ein/aus",
                            on_action=bkt.Callback(lambda slide: ToolboxAgenda.toggle_selector_text_style(slide, 'bold')),
                            get_image=bkt.Callback(lambda slide: bkt.ribbon.Gallery.get_check_image(ToolboxAgenda.get_selector_textcolor(slide, "bold"))),
                        ),
                        bkt.ribbon.Button(
                            label="Kursiv",
                            screentip="Selektor-Text kursiv",
                            supertip="Selektor-Text kursiv darstellen ein/aus",
                            on_action=bkt.Callback(lambda slide: ToolboxAgenda.toggle_selector_text_style(slide, 'italic')),
                            get_image=bkt.Callback(lambda slide: bkt.ribbon.Gallery.get_check_image(ToolboxAgenda.get_selector_textcolor(slide, "italic"))),
                        ),
                        bkt.ribbon.Button(
                            label="Unterstrichen",
                            screentip="Selektor-Text unterstrichen",
                            supertip="Selektor-Text unterstrichen darstellen ein/aus",
                            on_action=bkt.Callback(lambda slide: ToolboxAgenda.toggle_selector_text_style(slide, 'underline')),
                            get_image=bkt.Callback(lambda slide: bkt.ribbon.Gallery.get_check_image(ToolboxAgenda.get_selector_textcolor(slide, "underline"))),
                        ),
                        bkt.ribbon.Button(
                            label="Zurücksetzen",
                            screentip="Selektor-Text zurücksetzen",
                            supertip="Selektor-Text auf Standard zurücksetzen",
                            on_action=bkt.Callback(ToolboxAgenda.reset_selector_textcolor)
                        ),
                    ]
                ),
                bkt.ribbon.Menu(
                    label = "Höhe ändern",
                    size="large",
                    image_mso = 'GroupInkEdit',
                    screentip="Höhe für Selektor",
                    supertip="Passt die Höhe des Selektors relativ zur Schriftgröße an.",
                    get_enabled=bkt.Callback(ToolboxAgenda.is_agenda_slide),
                    children=[
                        bkt.ribbon.ToggleButton(
                            label="20% (Standard)",
                            screentip="Selektor-Höhe 20%",
                            supertip="Selektor-Überhang entspricht 20% der Schriftgröße",
                            get_pressed=bkt.Callback(lambda slide: ToolboxAgenda.get_selector_margin(slide) == 0.2),
                            on_toggle_action=bkt.Callback(lambda slide,pressed: ToolboxAgenda.set_selector_margin(0.2, slide)),
                        ),
                        bkt.ribbon.ToggleButton(
                            label="40%",
                            screentip="Selektor-Höhe 40%",
                            supertip="Selektor-Überhang entspricht 40% der Schriftgröße",
                            get_pressed=bkt.Callback(lambda slide: ToolboxAgenda.get_selector_margin(slide) == 0.4),
                            on_toggle_action=bkt.Callback(lambda slide,pressed: ToolboxAgenda.set_selector_margin(0.4, slide)),
                        ),
                        bkt.ribbon.ToggleButton(
                            label="60%",
                            screentip="Selektor-Höhe 60%",
                            supertip="Selektor-Überhang entspricht 60% der Schriftgröße",
                            get_pressed=bkt.Callback(lambda slide: ToolboxAgenda.get_selector_margin(slide) == 0.6),
                            on_toggle_action=bkt.Callback(lambda slide,pressed: ToolboxAgenda.set_selector_margin(0.6, slide)),
                        ),
                        bkt.ribbon.ToggleButton(
                            label="80% (sehr groß)",
                            screentip="Selektor-Höhe 80%",
                            supertip="Selektor-Überhang entspricht 80% der Schriftgröße",
                            get_pressed=bkt.Callback(lambda slide: ToolboxAgenda.get_selector_margin(slide) == 0.8),
                            on_toggle_action=bkt.Callback(lambda slide,pressed: ToolboxAgenda.set_selector_margin(0.8, slide)),
                        ),
                    ]
                ),
            ]
        )
    ]
)
