# -*- coding: utf-8 -*-
'''
Created on 29.03.2017

@author: tweuffel
'''

from __future__ import absolute_import

import bkt

TOOLBOX_NOTE = "TOOLBOX-NOTE"


class EditModeShapes(object):
    color_rgb = bkt.settings.get("ppt_notes.color_rgb", 16777062)
    color_theme = bkt.settings.get("ppt_notes.color_theme", 0)
    color_brightness = bkt.settings.get("ppt_notes.color_brightness", 0)
    
    @classmethod
    def addNote(cls, slide, context):
        from datetime import datetime
        # from System import Environment #used for Environment.UserName
        from getpass import getuser

        # Positionsanpassung ermitteln (unter existierenden Shape)
        yPosition = 0
        for cShp in slide.shapes:
            if cShp.Tags.Item(TOOLBOX_NOTE) != "":
                yPosition = cShp.top + cShp.height + 2
        # Shape rechts oben auf slide erstellen
        shp = slide.shapes.AddShape( 1 #msoShapeRectangle
            , 0, yPosition, 300, 20)
        shp.Left = context.app.ActivePresentation.PageSetup.SlideWidth - shp.width
        shp.Tags.Add(TOOLBOX_NOTE, "1")
        # Shape-Stil
        shp.Line.Weight = 0
        shp.Line.Visible = 0 #msoFalse
        shp.Fill.Visible = 1 #msoTrue
        if cls.color_theme > 0:
            shp.Fill.ForeColor.ObjectThemeColor = cls.color_theme
            shp.Fill.ForeColor.Brightness = cls.color_brightness
        else:
            shp.Fill.ForeColor.RGB = cls.color_rgb
        # Text-Stil
        shp.TextFrame.TextRange.Font.Color.RGB = 0
        shp.TextFrame.TextRange.Font.Size = 12
        shp.TextFrame.TextRange.Font.Bold = True
        shp.TextFrame.TextRange.ParagraphFormat.Alignment = 1 #ppAlignLeft
        shp.TextFrame.TextRange.ParagraphFormat.Bullet.Visible = False
        shp.TextFrame.VerticalAnchor = 1 #msoAnchorTop
        # Autosize / Text nicht umbrechen
        shp.TextFrame.WordWrap = 1 #msoTrue
        shp.TextFrame.AutoSize = 1 #ppAutoSizeShapeToFitText
        # Innenabstand
        shp.TextFrame.MarginBottom = 3
        shp.TextFrame.MarginTop    = 3
        shp.TextFrame.MarginLeft   = 5
        shp.TextFrame.MarginRight  = 5
        # Text
        dt = datetime.now()
        new_text = dt.strftime("%d.%m.%y %H:%M") + " (" + getuser() + "): EDIT"
        shp.TextFrame.TextRange.Text = new_text
        shp.Select() #first select shape, then text in shape. otherwise test is not selected in some cases.
        shp.TextFrame.TextRange.Characters(len(new_text)-3, 4).Select()
    
    
    @staticmethod
    def toogleNotesOnSlide(slide, context):
        visible = None
        for cShp in slide.shapes:
            if cShp.Tags.Item(TOOLBOX_NOTE) != "":
                if visible == None:
                    visible = 1 if cShp.Visible == 0 else 0
                cShp.Visible = visible
    
    
    @staticmethod
    def toggleNotesOnAllSlides(slide, context):
        visible = None
        for sld in slide.parent.slides:            
            for cShp in sld.shapes:
                if cShp.Tags.Item(TOOLBOX_NOTE) != "":
                    if visible == None:
                        visible = 1 if cShp.Visible == 0 else 0
                    cShp.Visible = visible
    
    
    @staticmethod
    def removeNotesOnSlide(slide, context):
        shapesToRemove = []
        
        for cShp in slide.shapes:
            if cShp.Tags.Item(TOOLBOX_NOTE) != "":
                shapesToRemove.append(cShp)
        
        for cShp in shapesToRemove:
            cShp.Delete()
    
    
    @staticmethod
    def removeNotesOnAllSlides(slide, context):
        for sld in slide.parent.slides:
            shapesToRemove = []
            
            for cShp in sld.shapes:
                if cShp.Tags.Item(TOOLBOX_NOTE) != "":
                    shapesToRemove.append(cShp)
        
            for cShp in shapesToRemove:
                cShp.Delete()

    @classmethod
    def set_color_default(cls):
        cls.color_rgb = 16777062
        cls.color_theme = 0
        cls.color_brightness = 0
        cls._save_color()

    @classmethod
    def set_color_rgb(cls, color):
        cls.color_rgb = color
        cls.color_theme = 0
        cls.color_brightness = 0
        cls._save_color()

    @classmethod
    def set_color_theme(cls, color_index, brightness):
        cls.color_rgb = 0
        cls.color_theme = color_index
        cls.color_brightness = brightness
        cls._save_color()
    
    @classmethod
    def _save_color(cls):
        bkt.settings["ppt_notes.color_rgb"] = cls.color_rgb
        bkt.settings["ppt_notes.color_theme"] = cls.color_theme
        bkt.settings["ppt_notes.color_brightness"] = cls.color_brightness

    @classmethod
    def get_color(cls):
        return [cls.color_theme, cls.color_brightness, cls.color_rgb]


notes_gruppe = bkt.ribbon.Group(
    id="bkt_notes_group",
    label='Notes',
    supertip="Ermöglicht das Einfügen von Bearbeitungsnotizen auf Folien. Das Feature `ppt_notes` muss installiert sein.",
    image='noteAdd',
    children = [
        bkt.ribbon.Button(
            label='Notizen (+)', screentip='Notiz hinzufügen',
            supertip="Fügt eine Bearbeitungsnotiz oben rechts auf der Folie ein inkl. Autor und Datum.",
            image='noteAdd',
            on_action=bkt.Callback(EditModeShapes.addNote)
        ),
        bkt.ribbon.Button(
            label='Notizen (I/O)', screentip='Notizen auf Folie ein-/ausblenden',
            supertip="Alle Notizen der aktuellen Folie temporär ausblenden und wieder einblenden.",
            image='noteToggle',
            on_action=bkt.Callback(EditModeShapes.toogleNotesOnSlide)
        ),
        bkt.ribbon.Button(
            label='Notizen (-)', screentip='Notizen auf Folie löschen',
            supertip="Alle Notizen der aktuellen Folie entfernen.",
            image='noteRemove',
            on_action=bkt.Callback(EditModeShapes.removeNotesOnSlide)
        ),
        bkt.ribbon.Button(
            label='Alle Notizen (I/O)', screentip='Alle Notizen ein-/ausblenden',
            supertip="Alle Notizen auf allen Folien temporär ausblenden und wieder einblenden.",
            image='noteToggleAll',
            on_action=bkt.Callback(EditModeShapes.toggleNotesOnAllSlides)
        ),
        bkt.ribbon.Button(
            label='Alle Notizen (-)', screentip='Alle Notizen löschen',
            supertip="Alle Notizen auf allen Folien entfernen.",
            image='noteRemoveAll',
            on_action=bkt.Callback(EditModeShapes.removeNotesOnAllSlides)
        ),
        bkt.ribbon.ColorGallery(
            id = 'notes_color',
            label=u'Farbe ändern',
            supertip="Hintergrundfarbe für neue Bearbeitungsnotizen ändern.",
            on_rgb_color_change = bkt.Callback(EditModeShapes.set_color_rgb),
            on_theme_color_change = bkt.Callback(EditModeShapes.set_color_theme),
            get_selected_color = bkt.Callback(EditModeShapes.get_color),
            children=[
                bkt.ribbon.Button(
                    id="notes_color_default",
                    label="Standardfarbe",
                    supertip="Hintergrundfarbe für Bearbeitungsnotizen auf Standard zurücksetzen.",
                    on_action=bkt.Callback(EditModeShapes.set_color_default),
                    image_mso="ColorTeal",
                )
            ]
            # get_enabled = bkt.apps.ppt_shapes_or_text_selected,
        ),
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
        notes_gruppe,
    ]
), extend=True)


