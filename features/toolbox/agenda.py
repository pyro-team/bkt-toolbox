# -*- coding: utf-8 -*-
'''
Created on 18.05.2016

@author: rdebeerst
'''

import bkt
import bkt.library.powerpoint as pplib

import json
import uuid

#import traceback

import bkt.dotnet as dotnet
ppt = dotnet.import_powerpoint()
office = dotnet.import_officecore()


TOOLBOX_AGENDA = "TOOLBOX-AGENDA"
TOOLBOX_AGENDA_SLIDENO  = "TOOLBOX-AGENDA-SLIDENO"
TOOLBOX_AGENDA_SELECTOR = "TOOLBOX-AGENDA-SELECTOR"
TOOLBOX_AGENDA_TEXTBOX  = "TOOLBOX-AGENDA-TEXTBOX"
TOOLBOX_AGENDA_SETTINGS = "TOOLBOX-AGENDA-SETTINGS"

SETTING_POSITION = "position"
SETTING_TEXT = "text"
SETTING_AGENDA_ID = "id"
SETTING_INDENT_LEVEL = "indent-level"
SETTING_HIDE_SUBITEMS = "hide-sub-items"
SETTING_CREATE_SECTIONS = "sections-create"
SETTING_SLIDES_FOR_SUBITEMS = "slides-for-sub-items"
SETTING_SELECTOR_FILL_COLOR = "selector-fill-color"
SETTING_SELECTOR_LINE_COLOR = "selector-line-color"



    


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
    #theme, brightness, rgb
    # selectorFillColor = [16, 0, 12566463] # 193 193 193   ##   193+193*255+193*255*255
    # selectorLineColor = [13, 0, 8355711] # 127 127 127    ##   ((long(127)*255)+127)*255+127

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

    selectorFillColor = default_selectorFillColor.copy()
    selectorLineColor = default_selectorLineColor.copy()
    
    default_settings = {
        SETTING_AGENDA_ID: None,
        SETTING_HIDE_SUBITEMS: False,
        SETTING_CREATE_SECTIONS: False,
        SETTING_SLIDES_FOR_SUBITEMS: True
    }
    
    settings = None
    
    
    # =================
    # = create agenda =
    # =================
    
    @classmethod
    def create_agenda_textbox_on_slide(cls, slide, context=None):
        
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
                    text += sections.Name(i+1) + "\r\n"
                shp.TextFrame.TextRange.text = text.strip()
            else:
                shp.TextFrame.TextRange.text = "Abschnitt 1\r\nAbschnitt 2\r\nAbschnitt 3"
            
            shp.TextFrame.VerticalAnchor = office.MsoVerticalAnchor.msoAnchorMiddle.value__
            shp.TextFrame.Ruler.Levels.item(1).FirstMargin = 0
            shp.TextFrame.Ruler.Levels.item(1).LeftMargin = 14
            shp.TextFrame.Ruler.TabStops.Add(ppt.PpTabStopType.ppTabStopRight.value__, shp.width)
            # Innenabstand
            shp.TextFrame.MarginBottom = 12
            shp.TextFrame.MarginTop = 12
            shp.TextFrame.MarginLeft = 6
            shp.TextFrame.MarginRight = 6
            shp.TextFrame.TextRange.ParagraphFormat.Bullet.Type = ppt.PpBulletType.ppBulletUnnumbered.value__
            shp.TextFrame.TextRange.ParagraphFormat.Bullet.Character = 167
            shp.TextFrame.TextRange.ParagraphFormat.Bullet.Font.Name = "Wingdings"
            shp.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 18
            # mittig anordnen
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
            
            #cls.set_tags_for_slide(slide, 0)
            cls.set_tags_for_textbox(shp)

            if context:
                context.ribbon.ActivateTab('bkt_context_tab_agenda')
        
        except:
            bkt.helpers.exception_as_message()
    
    
    @classmethod
    def create_agenda_from_textbox(cls, master_textbox, context=None):
        '''
        TODO
        '''
        master_slide = master_textbox.parent
        
        # set tags for master slide
        id = str(uuid.uuid4())
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
        for idx in range(1, master_textbox.TextFrame.TextRange.Paragraphs().Count+1):
            
            text = master_textbox.TextFrame.TextRange.Paragraphs(idx).text.strip()
            if text != "":
                # create slide 
                new_slide_count = new_slide_count + 1
                slide = master_slide.Duplicate(1)
                slide.SlideShowTransition.Hidden = 0
                slide.MoveTo(master_slide.SlideIndex + new_slide_count)
                
                # update agenda
                textbox = cls.get_agenda_textbox_on_slide(slide)
                cls.update_agenda_on_slide_new(slide, textbox, idx, settings)

                # update sections
                cls.update_agenda_sections_for_slide(slide, text, settings)

        if context:
            context.ribbon.ActivateTab('bkt_context_tab_agenda')
    
    @classmethod
    def create_agenda_from_slide(cls, slide, context):
        '''
        TODO
        '''
        
        master_textbox = cls.get_agenda_textbox_on_slide(slide)
        if master_textbox is None:
            bkt.helpers.message("Keine Agenda-Textbox auf dem Slide vorhanden.")
            return
        
        cls.create_agenda_from_textbox(master_textbox, context)
    
    
    @classmethod
    def get_or_create_selector_on_slide(cls, sld):
        '''
        finds selector on slide or creates selector-shape
        '''
        shp = cls.get_shape_with_tag_item(sld, TOOLBOX_AGENDA_SELECTOR)
        if not shp is None:
            return shp

        # Neues Selector-Shape erstellen
        shp = sld.shapes.AddShape(office.MsoAutoShapeType.msoShapeRectangle.value__, 0, 0, 100, 20)
        cls.set_tags_for_selector(shp)
        shp.ZOrder(office.MsoZOrderCmd.msoSendToBack.value__)
        # Grauer Hintergrund/Rand
        cls.set_selector_fill(shp.Fill, cls.selectorFillColor)
        cls.set_selector_line(shp.Line, cls.selectorLineColor)
        
        return shp
    
    
    
    # ===============
    # = find agenda =
    # ===============
    
    @classmethod
    def find_agenda_items_by_slide(cls, slide):
        '''
        returns list of agenda entries [position, text, indent-level, slide-reference] 
        considering agenda slides according to the agenda-id of the given slide or
        (if no agenda is contained on the given slide) according to the first agenda
        in the presentation 
        '''
        
        settings = cls.get_agenda_settings_from_slide(slide)
        if settings == {}:
            # no agenda settings found on slide
            if slide.Tags.Item(TOOLBOX_AGENDA) != "":
                # find all agenda-slides in presentation
                # print "fallback: find all agenda slides in presentation"
                bkt.helpers.message(slide.Tags.Item(TOOLBOX_AGENDA))
                bkt.helpers.message("No Agenda settings on current slide. Using all agenda slides", title="Toolbox: Agenda")
                return cls.find_all_agenda_slides(slide.parent)
            else:
                # slide is not an agenda slide
                bkt.helpers.message("Slide is no agenda slide!", title="Toolbox: Agenda")
                return []
        
        return cls.find_agenda_items_by_id(slide.parent, settings[SETTING_AGENDA_ID])
    
    
    @classmethod
    def find_agenda_items_by_id(cls, presentation, id):
        '''
        returns list of agenda entries [position, text, indent-level, slide-reference] 
        considering all agenda slides with given id in the presentation
        '''
        agenda_slides = []
        
        for slide in presentation.slides:
            settings = cls.get_agenda_settings_from_slide(slide)
            if settings.get(SETTING_AGENDA_ID, None) == id:
                agenda_slides.append( [ settings.get(SETTING_POSITION), settings.get(SETTING_TEXT), settings.get(SETTING_INDENT_LEVEL), slide ] )
        
        return agenda_slides
    
    
    @staticmethod
    def find_all_agenda_slides(presentation):
        '''
        returns list of agenda entries [position, text, indent-level, slide-reference] 
        considering all agenda slides in the presentation
        '''
        agenda_slides = []
        
        for slide in presentation.slides:
            if slide.Tags.Item(TOOLBOX_AGENDA) != "":
                agenda_slides.append( [ int(slide.Tags.Item(TOOLBOX_AGENDA_SLIDENO)), None, None, slide ] )
        
        return agenda_slides
    
    @staticmethod
    def presentation_has_agenda(presentation):
        for slide in presentation.slides:
            if slide.Tags.Item(TOOLBOX_AGENDA) != "":
                return True
        return False
    
    
    @staticmethod
    def get_agenda_slides(context):
        '''
        return list of all agenda-slides
        '''
        agendaSlides = {}
        
        for sld in context.app.ActiveWindow.Presentation.Slides:
            if sld.Tags.Item(TOOLBOX_AGENDA) != "":
                agendaSlides["slide-" + str(int(sld.Tags.Item(TOOLBOX_AGENDA_SLIDENO)))] = sld
        
        return agendaSlides
    

    @classmethod
    def get_agenda_textbox_on_slide(cls, sld):
        '''
        return agenda-textbox on given slide
        agenda-textbox is recognised by the tag TOOLBOX_AGENDA_TEXTBOX
        '''
        return cls.get_shape_with_tag_item(sld, TOOLBOX_AGENDA_TEXTBOX)
    
    
    @staticmethod
    def agenda_entries_from_textbox(textbox, slides_for_subitems=True):
        ''' returns list of agenda entries [par_index, text, indent-level, slide-reference] '''
        agenda_entries = []
        for idx in range(1, textbox.TextFrame.TextRange.Paragraphs().Count+1):
            if textbox.TextFrame.TextRange.Paragraphs(idx).IndentLevel <=1 or slides_for_subitems:
                agenda_entries.append( [idx, textbox.TextFrame.TextRange.Paragraphs(idx).Text.strip(), textbox.TextFrame.TextRange.Paragraphs(idx).IndentLevel, None] )
        
        return agenda_entries
    
    
    
    
    
    # =================
    # = update agenda =
    # =================
    
    @classmethod 
    def update_agenda_slides_by_slide(cls, slide):
        agenda_slides = cls.find_agenda_items_by_slide(slide)
        if len(agenda_slides) == 0:
            bkt.helpers.message("No Agenda found!", title="Toolbox: Update Agenda")
            return
        
        master_slide = None
        for item in agenda_slides:
            if item[0] == 0:
                # item with index=0 is master slide
                master_slide = item[3]
                break
        
        
        if master_slide == None:
            # no master slide in presentation --> use first agenda slide
            # show warning-message and ask to continue, if SETTING_HIDE_SUBITEMS=True
            settings = cls.get_agenda_settings(slide)
            if settings.get(SETTING_HIDE_SUBITEMS, False) == True:
                return_value = bkt.helpers.Forms.MessageBox.Show(
                    "Master-agenda-slide was not found!\nAgenda can be recreated from first agenda slide, but hidden sub-items will be lost.\n\nContinue agenda update?", 
                    "Toolbox: Update Agenda", 
                    bkt.helpers.Forms.MessageBoxButtons.YesNo,
                    bkt.helpers.Forms.MessageBoxIcon.Warning)
            else:
                return_value = bkt.helpers.Forms.DialogResult.Yes
            
            if return_value == bkt.helpers.Forms.DialogResult.Yes:
                master_slide = agenda_slides[0][3]
                # FIXME: master_slide hat selector, sieht aus wie erstes slide, erstes slide wird trotzdem noch erstellt
            else:
                return
        
        # get old settings
        old_settings = cls.get_agenda_settings_from_slide(master_slide)
        selector_fill_color = cls.get_selector_fillcolor_from_settings(old_settings)
        selector_line_color = cls.get_selector_linecolor_from_settings(old_settings)
        
        # find first selector and remember color
        for item in agenda_slides:
            if item[0] != 0:
                # first item with index!=0 will set selector color
                slide = item[3]
                shp = cls.get_or_create_selector_on_slide(slide)
                selector_fill_color["color"]      = [shp.Fill.ForeColor.ObjectThemeColor,float(shp.Fill.ForeColor.Brightness),shp.Fill.ForeColor.RGB]
                selector_fill_color['visibility'] = shp.Fill.Visible

                selector_line_color["color"]      = [shp.Line.ForeColor.ObjectThemeColor,float(shp.Line.ForeColor.Brightness),shp.Line.ForeColor.RGB]
                selector_line_color['visibility'] = shp.Line.Visible
                selector_line_color['weight']     = float(shp.Line.Weight)
                selector_line_color['style']      = shp.Line.Style
                selector_line_color['dashstyle']  = shp.Line.DashStyle
                break

        # == ensure settings on master slide
        # default settings with new id
        new_settings = dict(cls.default_settings)
        new_settings[SETTING_AGENDA_ID] = str(uuid.uuid4())
        new_settings[SETTING_POSITION] = 0
        # overwrite existing settings
        new_settings.update(old_settings)
        new_settings[SETTING_SELECTOR_FILL_COLOR] = selector_fill_color
        new_settings[SETTING_SELECTOR_LINE_COLOR] = selector_line_color
        # write settings
        cls.write_agenda_settings_to_slide(master_slide, new_settings)
            
        master_textbox = cls.get_agenda_textbox_on_slide(master_slide)
        cls.update_agenda_slides(master_textbox, agenda_slides=agenda_slides)
        
        
    
    
    @classmethod
    def update_agenda_slides(cls, master_textbox, agenda_slides=None):
        master_slide = master_textbox.parent
        settings = cls.get_agenda_settings(master_slide)
        
        # get agenda entries from textbox
        agenda_entries = cls.agenda_entries_from_textbox(master_textbox, slides_for_subitems=settings.get(SETTING_SLIDES_FOR_SUBITEMS) or False)
        # agenda entries from agenda_slides
        if not agenda_slides:
            agenda_slides = cls.find_agenda_items_by_id(master_slide.parent, settings[SETTING_AGENDA_ID])

        # find slides for agenda-entries by text
        for item in agenda_entries:
            par_index, text, indent_level, slide = item
            for idx in range(len(agenda_slides)):
                slide = agenda_slides[idx][3]
                slide_settings = cls.get_agenda_settings_from_slide(slide)
                if text == slide_settings.get(SETTING_TEXT):
                    item[3] = slide
                    del agenda_slides[idx]
                    break
        
        # find slides for agenda-entries by slide-order
        # - For every agenda-entry without corresponding slide,
        #   pick the first agenda-slide, which was not yet assigned to an agenda-entry
        # - This will update agenda-slides for renamed entries - if no entries were reordered or deleted
        for item in agenda_entries:
            par_index, text, indent_level, slide = item
            if slide == None and len(agenda_slides) > 1:
                item[3] = agenda_slides[1][3] #0 is master slide
                del agenda_slides[1]
        
        
        # udpate
        last_agenda_index = master_slide.slideindex
        for item in agenda_entries:
            par_index, text, indent_level, old_slide = item
            
            if old_slide == master_slide:
                old_slide = None
                new_slide = master_slide
            else:
                # duplicate master_slide
                new_slide = master_slide.Duplicate(1)
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

            # remember index for repositioning
            last_agenda_index = new_slide.slideindex
        
        
        # andere loeschen
        for item in agenda_slides:
            if item[3] != master_slide:
                cls.delete_agenda_sections_for_slide(item[3], settings)
                item[3].delete()
    
    
    
    @classmethod
    def update_agenda_on_slide_new(cls, slide, textbox, selected_paragraph_index, settings):
        '''
        TODO
        '''
        # slide_settings = dict(cls.default_settings)
        # slide_settings.update(settings)
        # slide_settings[text] = 
        # slide_settings[par_no] = 
        try:
            par = textbox.TextFrame.TextRange.Paragraphs(selected_paragraph_index)

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
                        textbox.Textframe.TextRange.paragraphs(idx).characters(textbox.Textframe.TextRange.paragraphs(idx).characters().count  ).Delete()
                
                # paragraphs hidden before current
                hidden_before_current = sum([1 for hide in hide_paragraph[0:selected_paragraph_index-1] if hide ])
                selected_paragraph_index = selected_paragraph_index - hidden_before_current
            
            
            # Tag auf Slide setzen
            slide_settings = dict(settings)
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
    
            # Rand ober- und unterhalb / um wieviel ist Markierung größer als Absatz?
            selectorMargin = textbox.TextFrame.TextRange.Font.Size * 0.2

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

            # Selector (Markierung) aktualisieren
            selector = cls.get_or_create_selector_on_slide(slide)
            selector.Top = selectorTop
            selector.Left = textbox.Left
            selector.Height = selectorHeight
            selector.width = textbox.width

            cls.set_selector_fill(selector.Fill, cls.get_selector_fillcolor_from_settings(settings))
            cls.set_selector_line(selector.Line, cls.get_selector_linecolor_from_settings(settings))
            # selector.Fill.ForeColor.RGB = settings.get(SETTING_SELECTOR_FILL_COLOR) or cls.selectorFillColor
            # selector.Line.ForeColor.RGB = settings.get(SETTING_SELECTOR_LINE_COLOR) or cls.selectorLineColor
    
            # Text und Textbox-Position aktualisieren
            oTextBox = cls.get_agenda_textbox_on_slide(slide)
    
            #oTextBox.TextFrame.TextRange.text = textbox.TextFrame.TextRange.text
            textbox.TextFrame.TextRange.Copy()
            oTextBox.TextFrame.TextRange.Paste()
            oTextBox.TextFrame.TextRange.Font.Bold = False
            oTextBox.TextFrame.TextRange.Paragraphs(selected_paragraph_index).Font.Bold = True
            oTextBox.Top    = textbox.Top
            oTextBox.Left   = textbox.Left
            oTextBox.Width  = textbox.Width
            oTextBox.Height = textbox.Height
        
            # write settings and set tag
            cls.write_agenda_settings_to_slide(slide, slide_settings)
            cls.set_tags_for_slide(slide, selected_paragraph_index)
            
        except:
            bkt.helpers.exception_as_message()


    @classmethod
    def update_agenda_sections_for_slide(cls, slide, text, settings):
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
            bkt.helpers.exception_as_message()


    @classmethod
    def delete_agenda_sections_for_slide(cls, slide, settings):
        try:
            if settings.get(SETTING_CREATE_SECTIONS, False) == False:
                return
            sections = slide.parent.SectionProperties
            if sections.Count > 0 and sections.FirstSlide(slide.SectionIndex) == slide.SlideIndex:
                sections.Delete(slide.SectionIndex, False)
        except:
            bkt.helpers.exception_as_message()
        
    
    
    
    # =================
    # = ALTE METHODEN =
    # =================
    
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
    #             bkt.helpers.message("In Slide-View wechseln!")
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
    #                 bkt.helpers.message("Textbox mit Agenda-Einträgen auswählen oder als Bullet-Liste formatieren.")
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
    #             bkt.helpers.message("Update nicht möglich! Agenda-Textbox fehlt auf erstem Agenda-Slide.")
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
    #                         bkt.helpers.message("Not implemented yet!")
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
    def remove_agenda(cls, presentation):
        try:
            cls.remove_agenda_from_presentation(presentation, delete_slides=True)
        except:
            bkt.helpers.exception_as_message()
    
    
    @classmethod
    def remove_agenda_from_presentation(cls, presentation, delete_slides=False):
        '''
        removes all agenda-slides from the presentation, removing all meta-information
        '''
        agenda_slides = []
        
        for slide in presentation.slides:
            if slide.Tags.Item(TOOLBOX_AGENDA) != "":
                if slide.Tags.Item(TOOLBOX_AGENDA) != "":
                    # Tags von Slide loeschen
                    slide.Tags.Add(TOOLBOX_AGENDA, "")
                    slide.Tags.Add(TOOLBOX_AGENDA_SLIDENO, "")
                    slide.Tags.Add(TOOLBOX_AGENDA_SETTINGS, "")
                    
                    # Tags von Shapes auf Slide loeschen
                    shape = cls.get_agenda_textbox_on_slide(slide)
                    if not shape is None:
                        shape.Tags.Add(TOOLBOX_AGENDA_TEXTBOX, "")
                    shape = cls.get_agenda_textbox_on_slide(slide)
                    if not shape is None:
                        shape.Tags.Add(TOOLBOX_AGENDA_SELECTOR, "")
                
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
        ''' load agenda settings form current slide or presentation '''
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
        ''' load agenda settings form given slide '''
        value = cls.get_tag_value(slide, TOOLBOX_AGENDA_SETTINGS, "{}" )
        settings = json.loads( value ) or {}

        #convert old settings with rgb color to new color format
        if type(settings.get(SETTING_SELECTOR_FILL_COLOR)) == int:
            settings[SETTING_SELECTOR_FILL_COLOR] = cls.default_selectorFillColor.copy()
            settings[SETTING_SELECTOR_FILL_COLOR]['color'] = [0, 0, settings.get(SETTING_SELECTOR_FILL_COLOR)]
        
        if type(settings.get(SETTING_SELECTOR_LINE_COLOR)) == int:
            settings[SETTING_SELECTOR_LINE_COLOR] = cls.default_selectorLineColor.copy()
            settings[SETTING_SELECTOR_LINE_COLOR]['color'] = [0, 0, settings.get(SETTING_SELECTOR_LINE_COLOR)]

        return settings
    
    
    @staticmethod
    def write_agenda_settings_to_slide(slide, settings):
        ''' write agenda settings to slide tags '''
        slide.Tags.Add(TOOLBOX_AGENDA_SETTINGS, json.dumps(settings, ensure_ascii=False) )
    
    
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
    def get_slides_for_subitems(cls, slide):
        ''' reads setting hide_subitems from current slide or presentation'''
        settings = cls.get_agenda_settings_from_slide(slide)
        return settings.get(SETTING_SLIDES_FOR_SUBITEMS) or False
    

    @classmethod
    def get_selector_fillcolor_from_settings(cls, settings):
        # selector_fill_color = cls.selectorFillColor.copy()
        # selector_fill_color.update(settings.get(SETTING_SELECTOR_FILL_COLOR, {}))
        # return selector_fill_color
        cls.selectorFillColor.update(settings.get(SETTING_SELECTOR_FILL_COLOR, {}))
        return cls.selectorFillColor
    

    @classmethod
    def get_selector_linecolor_from_settings(cls, settings):
        # selector_line_color = cls.selectorLineColor.copy()
        # selector_line_color.update(settings.get(SETTING_SELECTOR_LINE_COLOR, {}))
        # return selector_line_color
        cls.selectorLineColor.update(settings.get(SETTING_SELECTOR_LINE_COLOR, {}))
        return cls.selectorLineColor

    
    # ========================
    # = save/load tag values =
    # ========================
    
    @staticmethod
    def set_tags_for_slide(sld, slideNo):
        ''' Meta-Informationen für Slide einstellen '''
        sld.Tags.Add(TOOLBOX_AGENDA, "1")
        sld.Tags.Add(TOOLBOX_AGENDA_SLIDENO, str(slideNo))

    @staticmethod
    def set_tags_for_textbox(textbox):
        ''' Meta-Informationen für Textbox einstellen '''
        textbox.Tags.Add(TOOLBOX_AGENDA_TEXTBOX, "1")

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
        settings = cls.get_agenda_settings_from_slide(slide)
        return settings != {}
    
    @classmethod
    def can_create_agenda_from_slide(cls, slide):
        textbox = cls.get_agenda_textbox_on_slide(slide)
        return textbox != None
    
    @classmethod
    def set_hide_subitems(cls, slide, pressed):
        ''' callback to write hide_subitems setting to all agenda slides '''
        agenda_items = cls.find_agenda_items_by_slide(slide)
        slide = agenda_items[0][3]
        settings = cls.get_agenda_settings_from_slide(slide)
        settings[SETTING_HIDE_SUBITEMS] = (pressed==True)
        for agenda_item in agenda_items:
            slide = agenda_item[3]
            cls.write_agenda_settings_to_slide(slide, settings)
    
    @classmethod
    def set_create_sections(cls, slide, pressed):
        ''' callback to write sections-create setting to all agenda slides '''
        agenda_items = cls.find_agenda_items_by_slide(slide)
        slide = agenda_items[0][3]
        settings = cls.get_agenda_settings_from_slide(slide)
        settings[SETTING_CREATE_SECTIONS] = (pressed==True)
        for agenda_item in agenda_items:
            slide = agenda_item[3]
            cls.write_agenda_settings_to_slide(slide, settings)

    @classmethod
    def set_slides_for_subitems(cls, slide, pressed):
        ''' callback to write slides_for_subitmes setting to all agenda slides '''
        agenda_items = cls.find_agenda_items_by_slide(slide)
        slide = agenda_items[0][3]
        settings = cls.get_agenda_settings_from_slide(slide)
        settings[SETTING_SLIDES_FOR_SUBITEMS] = (pressed==True)
        for agenda_item in agenda_items:
            slide = agenda_item[3]
            cls.write_agenda_settings_to_slide(slide, settings)
    
    # @classmethod
    # def _update_setting_value(cls, slide, key, value):
    #     ''' callback to write settings value to all agenda slides '''
    #     agenda_items = cls.find_agenda_items_by_slide(slide)
    #     slide = agenda_items[0][3]
    #     settings = cls.get_agenda_settings_from_slide(slide)
    #     settings[key] = value
    #     for agenda_item in agenda_items:
    #         slide = agenda_item[3]
    #         cls.write_agenda_settings_to_slide(slide, settings)
    
    
    
    # =========================
    # = color gallery methods =
    # =========================
    
    @classmethod
    def get_selector_fillcolor(cls, slide):
        try:
            settings = cls.get_agenda_settings_from_slide(slide)
            return settings.get(SETTING_SELECTOR_FILL_COLOR)["color"]
        except:
            return cls.selectorFillColor["color"]
        # return [0, 0, cls.selectorFillColor]
    
    @classmethod
    def get_selector_linecolor(cls, slide):
        try:
            settings = cls.get_agenda_settings_from_slide(slide)
            return settings.get(SETTING_SELECTOR_LINE_COLOR)["color"]
        except:
            return cls.selectorLineColor["color"]
        # return [0, 0, cls.selectorLineColor]
    
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
    def reset_selector_fillcolor(cls, slide):
        # cls.set_selector_fillcolor_rgb(12566463, slide)
        cls.set_selector_fillcolor(cls.default_selectorFillColor, slide)
    
    @classmethod
    def reset_selector_linecolor(cls, slide):
        # cls.set_selector_linecolor_rgb(8355711, slide)
        cls.set_selector_linecolor(cls.default_selectorLineColor, slide)
    
    @classmethod
    def hide_selector_fill(cls, slide):
        cls.set_selector_fillcolor({'visibility': 0}, slide)
    
    @classmethod
    def hide_selector_line(cls, slide):
        cls.set_selector_linecolor({'visibility': 0}, slide)
    
    
    # ===============================
    # = selector adjustment methods =
    # ===============================

    @classmethod
    def set_selector_fillcolor(cls, color_dict, slide):
        agenda_items = cls.find_agenda_items_by_slide(slide)
        slide = agenda_items[0][3]

        settings = cls.get_agenda_settings_from_slide(slide)
        settings[SETTING_SELECTOR_FILL_COLOR] = cls.get_selector_fillcolor_from_settings(settings)
        settings[SETTING_SELECTOR_FILL_COLOR].update(color_dict)
        for agenda_item in agenda_items:
            slide = agenda_item[3]
            cls.write_agenda_settings_to_slide(slide, settings)
            #recolor each selector right away
            shp = cls.get_shape_with_tag_item(slide, TOOLBOX_AGENDA_SELECTOR)
            if not shp is None:
                cls.set_selector_fill(shp.Fill, settings[SETTING_SELECTOR_FILL_COLOR])
    
    @classmethod
    def set_selector_linecolor(cls, color_dict, slide):
        agenda_items = cls.find_agenda_items_by_slide(slide)
        slide = agenda_items[0][3]

        settings = cls.get_agenda_settings_from_slide(slide)
        settings[SETTING_SELECTOR_LINE_COLOR] = cls.get_selector_linecolor_from_settings(settings)
        settings[SETTING_SELECTOR_LINE_COLOR].update(color_dict)
        for agenda_item in agenda_items:
            slide = agenda_item[3]
            cls.write_agenda_settings_to_slide(slide, settings)
            #recolor each selector right away
            shp = cls.get_shape_with_tag_item(slide, TOOLBOX_AGENDA_SELECTOR)
            if not shp is None:
                cls.set_selector_line(shp.Line, settings[SETTING_SELECTOR_LINE_COLOR])
    
    @classmethod
    def set_selector_fill(cls, shape_fill_obj, color_dict):
        shape_fill_obj.Solid() #default shape might have gradient or other non-solid background
        if 'visibility' in color_dict:
            shape_fill_obj.Visible = color_dict["visibility"]
        if 'color' in color_dict:
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
            screentip="Standard Agenda-Textbox einfügen.",
            imageMso="TextBoxInsert",
            on_action=bkt.Callback(ToolboxAgenda.create_agenda_textbox_on_slide)
        ),
        bkt.ribbon.Button(
            id='agenda-new-create',
            label="Agenda neu erstellen",
            screentip="Neue Agenda auf Basis des aktuellen Slides erstellen. Aktuelles Slide wird Master-Slide der Agenda.",
            imageMso="TableOfContentsAddTextGallery",
            on_action=bkt.Callback(ToolboxAgenda.create_agenda_from_slide),
            get_enabled=bkt.Callback(ToolboxAgenda.can_create_agenda_from_slide)
        ),
        bkt.ribbon.MenuSeparator(),
        bkt.ribbon.Button(
            id='agenda-new-update',
            label="Agenda aktualisieren",
            screentip="Agenda aktualisieren und durch Agenda auf dem Agenda-Master-Slide ersetzen; Folien werden dabei neu erstellt.",
            imageMso="SaveSelectionToTableOfContentsGallery",
            on_action=bkt.Callback(ToolboxAgenda.update_agenda_slides_by_slide),
            get_enabled=bkt.Callback(ToolboxAgenda.is_agenda_slide)
        ),
        bkt.ribbon.ToggleButton(
            id='agenda-slide-for-subitems',
            label="Agenda-Slides für Unterpunkte",
            screentip="Für Unterpunkte eines Agendapunkts (Indent-Level>1) werden Agenda-Slides erstellt",
            on_toggle_action=bkt.Callback(ToolboxAgenda.set_slides_for_subitems),
            get_pressed=bkt.Callback(ToolboxAgenda.get_slides_for_subitems),
            get_enabled=bkt.Callback(ToolboxAgenda.is_agenda_slide)
        ),
        bkt.ribbon.ToggleButton(
            id='agenda-hide-subitems',
            label="Andere Agenda-Unterpunkte ausblenden",
            screentip="Unterpunkte eines Agendapunkts (Indent-Level>1) werden in den anderen Abschnitten ausgeblendet",
            on_toggle_action=bkt.Callback(ToolboxAgenda.set_hide_subitems),
            get_pressed=bkt.Callback(ToolboxAgenda.get_hide_subitems),
            get_enabled=bkt.Callback(ToolboxAgenda.is_agenda_slide)
        ),
        bkt.ribbon.ToggleButton(
            id='agenda-create-sections',
            label="Abschnitte für Agenda-Punkte erstellen",
            screentip="Einen neuen Abschnitt je Agenda-Folie beginnen.",
            on_toggle_action=bkt.Callback(ToolboxAgenda.set_create_sections),
            get_pressed=bkt.Callback(ToolboxAgenda.get_create_sections),
            get_enabled=bkt.Callback(ToolboxAgenda.is_agenda_slide)
        ),
        bkt.ribbon.MenuSeparator(),
        bkt.ribbon.Button(
            id='agenda-remove',
            label="Alle Agenda-Slides entfernen",
            screentip="Entfernt alle Agenda-Slides, alle Meta-Informationen werden gelöscht.",
            imageMso="TableOfContentsRemove",
            on_action=bkt.Callback(ToolboxAgenda.remove_agenda),
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
            label = "Anleitung",
            children = [
                bkt.ribbon.Label(label='Schritt 1: Textbox mit Agenda füllen und "Agenda neu erstellen"'),
                bkt.ribbon.Label(label='Schritt 2: Nach jeder weiteren Änderung "Agenda aktualisieren"'),
                bkt.ribbon.Label(label='Hinweis: Agenda-Masterfolie sollte nicht gelöscht werden!'),
            ]
        ),
        bkt.ribbon.Group(
            label = "Agenda",
            children = [
                bkt.ribbon.Button(
                    id='agenda_new_create',
                    label="Agenda neu erstellen",
                    size="large",
                    screentip="Agenda neu erstellen",
                    supertip="Neue Agenda auf Basis des aktuellen Slides erstellen. Aktuelles Slide wird Master-Slide der Agenda.",
                    imageMso="TableOfContentsAddTextGallery",
                    on_action=bkt.Callback(ToolboxAgenda.create_agenda_from_slide),
                    get_enabled=bkt.Callback(ToolboxAgenda.can_create_agenda_from_slide)
                ),
                bkt.ribbon.Button(
                    id='agenda_new_update',
                    label="Agenda aktualisieren",
                    size="large",
                    screentip="Agenda aktualisieren",
                    supertip="Agenda aktualisieren und durch Agenda auf dem Agenda-Master-Slide ersetzen; Folien werden dabei neu erstellt.",
                    imageMso="SaveSelectionToTableOfContentsGallery",
                    on_action=bkt.Callback(ToolboxAgenda.update_agenda_slides_by_slide),
                    get_enabled=bkt.Callback(ToolboxAgenda.is_agenda_slide)
                ),
                bkt.ribbon.Separator(),
                bkt.ribbon.Menu(
                    label="Optionen",
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
                    ]
                ),
                bkt.ribbon.Separator(),
                bkt.ribbon.Button(
                    id='agenda_remove',
                    label="Alle Agenda-Slides entfernen",
                    size="large",
                    screentip="Alle Agenda-Slides entfernen",
                    supertip="Entfernt alle Agenda-Slides, alle Meta-Informationen werden gelöscht.",
                    imageMso="TableOfContentsRemove",
                    on_action=bkt.Callback(ToolboxAgenda.remove_agenda),
                    get_enabled=bkt.Callback(ToolboxAgenda.presentation_has_agenda)
                ),
            ]
        ),
        bkt.ribbon.Group(
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
                            label="Farbe zurücksetzen",
                            on_action=bkt.Callback(ToolboxAgenda.reset_selector_fillcolor)
                        ),
                        bkt.ribbon.Button(
                            label="Keine Füllung",
                            on_action=bkt.Callback(ToolboxAgenda.hide_selector_fill)
                        )
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
                            label="Farbe zurücksetzen",
                            on_action=bkt.Callback(ToolboxAgenda.reset_selector_linecolor)
                        ),
                        bkt.ribbon.Button(
                            label="Kein Rahmen",
                            on_action=bkt.Callback(ToolboxAgenda.hide_selector_line)
                        )
                    ]
                ),
            ]
        )
    ]
)