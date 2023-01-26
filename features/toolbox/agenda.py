# -*- coding: utf-8 -*-
'''
Created on 18.05.2016

@author: rdebeerst
'''


import bkt


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

TOOLBOX_AGENDA = "TOOLBOX-AGENDA"
TOOLBOX_AGENDA_SLIDENO  = "TOOLBOX-AGENDA-SLIDENO"
TOOLBOX_AGENDA_SELECTOR = "TOOLBOX-AGENDA-SELECTOR"
TOOLBOX_AGENDA_TEXTBOX  = "TOOLBOX-AGENDA-TEXTBOX"
# TOOLBOX_AGENDA_SETTINGS = "TOOLBOX-AGENDA-SETTINGS"

class ToolboxAgendaUi(object):
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
    def is_agenda_slide(cls, slide):
        ''' check if current slide is agenda-slide '''
        try:
            return slide.Tags.Item(TOOLBOX_AGENDA_SLIDENO) != ""
        except: #AttributeError
            return False

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
    def get_shape_with_tag_item(sld, tagKey):
        ''' Shape auf Slide finden, das einen bestimmten TagKey enthaelt '''
        for shp in sld.shapes:
            if shp.Tags.Item(tagKey) != "":
                return shp 
        return None


agendamenu = bkt.ribbon.Menu(
    label="Agenda",
    children=[
        bkt.ribbon.Button(
            id='add-agenda-textbox',
            label="Agenda-Textbox einfügen",
            supertip="Standard Agenda-Textbox einfügen, um daraus eine aktualisierbare Agenda zu generieren.",
            imageMso="TextBoxInsert",
            on_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda" ,"create_agenda_textbox_on_slide", slide=True, context=True)
        ),
        bkt.ribbon.Button(
            id='agenda-new-create',
            label="Agenda neu erstellen",
            supertip="Neue Agenda auf Basis der aktuellen Folie erstellen. Aktuelle Folien wird Hauptfolie der Agenda.",
            imageMso="TableOfContentsAddTextGallery",
            on_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "create_agenda_from_slide", slide=True, context=True),
            get_enabled=bkt.Callback(ToolboxAgendaUi.can_create_agenda_from_slide, slide=True)
        ),
        bkt.ribbon.MenuSeparator(),
        bkt.ribbon.Button(
            id='agenda-new-update',
            label="Agenda aktualisieren",
            supertip="Agenda aktualisieren und durch Agenda auf dem Agenda-Hauptfolie ersetzen; Folien werden dabei neu erstellt.",
            imageMso="SaveSelectionToTableOfContentsGallery",
            on_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "update_agenda_slides_by_slide", slide=True),
            get_enabled=bkt.Callback(ToolboxAgendaUi.is_agenda_slide, slide=True)
        ),
        bkt.ribbon.DynamicMenu(
            id='agenda-options-menu',
            label="Agenda-Einstellungen",
            get_content=bkt.CallbackLazy("toolbox.models.agenda", "agenda_options_menu"),
        ),
        bkt.ribbon.MenuSeparator(),
        bkt.ribbon.Button(
            id='agenda-remove',
            label="Agenda-Folie entfernen",
            supertip="Entfernt Agenda-Folien der gewählten Agenda, alle Meta-Informationen werden gelöscht.",
            imageMso="TableOfContentsRemove",
            on_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "remove_agenda", slide=True, presentation=True),
            get_enabled=bkt.Callback(ToolboxAgendaUi.is_agenda_slide, slide=True)
        ),
        bkt.ribbon.Button(
            id='agenda-remove-all',
            label="Alle Agenden aus Präsentation entfernen",
            supertip="Entfernt alle Agenda-Folien in der ganzen Präsentation, alle Meta-Informationen werden gelöscht.",
            imageMso="TableOfContentsRemove",
            on_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "remove_agendas_from_presentation", presentation=True),
            get_enabled=bkt.Callback(ToolboxAgendaUi.presentation_has_agenda, presentation=True)
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
    get_visible=bkt.Callback(ToolboxAgendaUi.can_create_agenda_from_slide, slide=True),
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
                    on_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "create_agenda_from_slide", slide=True, context=True),
                    get_enabled=bkt.Callback(ToolboxAgendaUi.can_create_agenda_from_slide, slide=True)
                ),
                bkt.ribbon.Button(
                    id='agenda_new_update',
                    label="Agenda aktualisieren",
                    size="large",
                    supertip="Agenda aktualisieren und durch Agenda auf dem Agenda-Hauptfolie ersetzen; Folien werden dabei neu erstellt.",
                    imageMso="SaveSelectionToTableOfContentsGallery",
                    on_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "update_agenda_slides_by_slide", slide=True),
                    get_enabled=bkt.Callback(ToolboxAgendaUi.is_agenda_slide, slide=True)
                ),
                bkt.ribbon.Separator(),
                bkt.ribbon.DynamicMenu(
                    label="Optionen",
                    screentip="Agenda-Optionen",
                    supertip="Verschiedene Agenda-Optionen ändern",
                    imageMso="TableProperties",
                    size="large",
                    get_enabled=bkt.Callback(ToolboxAgendaUi.is_agenda_slide, slide=True),
                    get_content=bkt.CallbackLazy("toolbox.models.agenda", "agenda_options_menu"),
                ),
                bkt.ribbon.Separator(),
                bkt.ribbon.Button(
                    id='agenda_remove',
                    label="Agenda-Folien entfernen",
                    size="large",
                    screentip="Alle zugehörigen Agenda-Folien entfernen",
                    supertip="Entfernt alle Agenda-Folien, die zur aktuellen Agenda gehören, außer der Hauptfolie. Alle Meta-Informationen werden gelöscht.",
                    imageMso="TableOfContentsRemove",
                    on_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "remove_agenda", slide=True, presentation=True),
                    get_enabled=bkt.Callback(ToolboxAgendaUi.is_agenda_slide, slide=True)
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
                    on_rgb_color_change   = bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "set_selector_fillcolor_rgb", slide=True),
                    on_theme_color_change = bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "set_selector_fillcolor_theme", slide=True),
                    get_selected_color    = bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "get_selector_fillcolor", slide=True),
                    get_enabled=bkt.Callback(ToolboxAgendaUi.is_agenda_slide, slide=True),
                    children=[
                        bkt.ribbon.Button(
                            label="Keine Füllung",
                            supertip="Selektor-Hintergrund auf transparent ändern",
                            on_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "hide_selector_fill", slide=True),
                            get_image=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "get_check_fillcolor", slide=True),
                        ),
                        bkt.ribbon.Button(
                            label="Zurücksetzen",
                            screentip="Selektor-Hintergrund zurücksetzen",
                            supertip="Selektor-Hintergrund auf Standard zurücksetzen",
                            on_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "reset_selector_fillcolor", slide=True)
                        ),
                    ]
                ),
                bkt.ribbon.ColorGallery(
                    label = 'Rahmen ändern',
                    size="large",
                    image_mso = 'ShapeOutlineColorPicker',
                    screentip="Linienfarbe für Selektor",
                    supertip="Passe die Linienfarbe für den Selektor, der den aktiven Agendapunkt hervorhebt, an.",
                    on_rgb_color_change   = bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "set_selector_linecolor_rgb", slide=True),
                    on_theme_color_change = bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "set_selector_linecolor_theme", slide=True),
                    get_selected_color    = bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "get_selector_linecolor", slide=True),
                    get_enabled=bkt.Callback(ToolboxAgendaUi.is_agenda_slide, slide=True),
                    children=[
                        bkt.ribbon.Button(
                            label="Kein Rahmen",
                            supertip="Selektor-Rahmen auf transparent ändern",
                            on_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "hide_selector_line", slide=True),
                            get_image=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "get_check_linecolor", slide=True),
                        ),
                        bkt.ribbon.Button(
                            label="Zurücksetzen",
                            screentip="Selektor-Rahmen zurücksetzen",
                            supertip="Selektor-Rahmen auf Standard zurücksetzen",
                            on_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "reset_selector_linecolor", slide=True)
                        ),
                    ]
                ),
                bkt.ribbon.ColorGallery(
                    label = 'Text ändern',
                    size="large",
                    image_mso = 'TextFillColorPicker',
                    screentip="Textfarbe für Selektor",
                    supertip="Passe die Textfarbe für den Selektor, der den aktiven Agendapunkt hervorhebt, an.",
                    on_rgb_color_change   = bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "set_selector_textcolor_rgb", slide=True),
                    on_theme_color_change = bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "set_selector_textcolor_theme", slide=True),
                    get_selected_color    = bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "get_selector_textcolor", slide=True),
                    get_enabled=bkt.Callback(ToolboxAgendaUi.is_agenda_slide, slide=True),
                    children=[
                        bkt.ribbon.Button(
                            label="Fett",
                            screentip="Selektor-Text fett",
                            supertip="Selektor-Text fett darstellen ein/aus",
                            on_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "toggle_selector_text_style", slide=True, current_control=True),
                            get_image=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "get_check_textcolor", slide=True, current_control=True),
                            tag="bold",
                        ),
                        bkt.ribbon.Button(
                            label="Kursiv",
                            screentip="Selektor-Text kursiv",
                            supertip="Selektor-Text kursiv darstellen ein/aus",
                            on_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "toggle_selector_text_style", slide=True, current_control=True),
                            get_image=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "get_check_textcolor", slide=True, current_control=True),
                            tag="italic",
                        ),
                        bkt.ribbon.Button(
                            label="Unterstrichen",
                            screentip="Selektor-Text unterstrichen",
                            supertip="Selektor-Text unterstrichen darstellen ein/aus",
                            on_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "toggle_selector_text_style", slide=True, current_control=True),
                            get_image=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "get_check_textcolor", slide=True, current_control=True),
                            tag="underline",
                        ),
                        bkt.ribbon.Button(
                            label="Zurücksetzen",
                            screentip="Selektor-Text zurücksetzen",
                            supertip="Selektor-Text auf Standard zurücksetzen",
                            on_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "reset_selector_textcolor", slide=True)
                        ),
                    ]
                ),
                bkt.ribbon.Menu(
                    label = "Höhe ändern",
                    size="large",
                    image_mso = 'GroupInkEdit',
                    screentip="Höhe für Selektor",
                    supertip="Passt die Höhe des Selektors relativ zur Schriftgröße an.",
                    get_enabled=bkt.Callback(ToolboxAgendaUi.is_agenda_slide),
                    children=[
                        bkt.ribbon.ToggleButton(
                            label="20% (Standard)",
                            screentip="Selektor-Höhe 20%",
                            supertip="Selektor-Überhang entspricht 20% der Schriftgröße",
                            get_pressed=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "get_pressed_selector_margin", slide=True, current_control=True),
                            on_toggle_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "set_selector_margin", slide=True, current_control=True),
                            tag="0.2",
                        ),
                        bkt.ribbon.ToggleButton(
                            label="40%",
                            screentip="Selektor-Höhe 40%",
                            supertip="Selektor-Überhang entspricht 40% der Schriftgröße",
                            get_pressed=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "get_pressed_selector_margin", slide=True, current_control=True),
                            on_toggle_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "set_selector_margin", slide=True, current_control=True),
                            tag="0.4",
                        ),
                        bkt.ribbon.ToggleButton(
                            label="60%",
                            screentip="Selektor-Höhe 60%",
                            supertip="Selektor-Überhang entspricht 60% der Schriftgröße",
                            get_pressed=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "get_pressed_selector_margin", slide=True, current_control=True),
                            on_toggle_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "set_selector_margin", slide=True, current_control=True),
                            tag="0.6",
                        ),
                        bkt.ribbon.ToggleButton(
                            label="80% (sehr groß)",
                            screentip="Selektor-Höhe 80%",
                            supertip="Selektor-Überhang entspricht 80% der Schriftgröße",
                            get_pressed=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "get_pressed_selector_margin", slide=True, current_control=True),
                            on_toggle_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "set_selector_margin", slide=True, current_control=True),
                            tag="0.8",
                        ),
                    ]
                ),
            ]
        )
    ]
)
