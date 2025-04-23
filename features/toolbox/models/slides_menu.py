# -*- coding: utf-8 -*-
'''
Created on 19.01.2023

'''

import logging
import os.path

from contextlib import contextmanager

import bkt
import bkt.ui
import bkt.library.powerpoint as pplib

# for ui composition
from .. import agenda
from .. import language



class SendOrSaveSlides(object):
    @classmethod
    def _slide_range_identifier(cls, slides):
        indices = sorted([slide.SlideNumber for slide in slides])
        ind_str = str(indices[0])
        last_index = indices[0]
        range_from = last_index
        in_range = False
        for ind in indices[1:]:
            if ind == last_index +1:
                in_range = True
            else:
                if in_range:
                    if last_index - range_from ==1:
                        ind_str += ',' + str(last_index)
                    else:
                        ind_str += '-' + str(last_index)
                in_range = False
                range_from = ind
                ind_str +=',' + str(ind)
            last_index = ind
        if in_range:
            if last_index - range_from ==1:
                ind_str += ',' + str(last_index)
            else:
                ind_str += '-' + str(last_index)
        return ind_str

    @classmethod
    def initial_file_name(cls, presentation, slides=None):
        import time
        
        # bisheriger Dateiname
        fileName, fileExt = os.path.splitext(presentation.Name)
        if fileExt == '':
            fileExt = ".pptx"

        # Foliennummern
        if slides is not None:
            if len(slides) == 1:
                fileName = fileName + "_Folie_" + str(slides[0].SlideNumber)
            else:
                fileName = fileName + "_Folien_" + cls._slide_range_identifier(slides)

        # Zeitstempel
        fileName = fileName + "_" + time.strftime("%Y%m%d%H%M")
        # Dateiendung
        fileName = fileName + fileExt

        return fileName

    @classmethod
    def _delete_unselected_slides(cls, slides, newPres):
        from System import Array, Int32

        # get slide indicies of slides to keep
        slideIndices = [slide.SlideIndex for slide in slides]
        # list of all indices - the one to keep = all indices to remove
        removeIndices = list(set(range(1,newPres.Slides.Count+1)) - set(slideIndices))
        # if there are any slides to remove, delete all at once
        if len(removeIndices) > 0:
            newPres.Slides.Range(Array[Int32](removeIndices)).Delete()
        # Folien rueckwaerts durchgehen und schauen, ob Index in Range dabei ist
        # removeIndices.sort()
        # removeIndices.reverse()
        # for slideId in removeIndices:
        #     newPres.Slides(slideId).Delete()
    
    @classmethod
    @contextmanager
    def _create_temporary_copy(cls, application, filename):
        import tempfile, os
        #avoid referenced before assignment error in finally clause
        temporary_ppt_file = None
        newPres = None
        try:
            presentation = application.ActiveWindow.Presentation
            # Richtige Dateiendung prüfen und temporären Pfad erstellen
            pres_ext = os.path.splitext(presentation.Name)[1]
            if not pres_ext:
                pres_ext = '.pptx'
            filename = os.path.splitext(filename)[0] + pres_ext
            temporary_ppt_file = os.path.join(tempfile.gettempdir(), filename)
            logging.debug("save temporary copy to: %s", temporary_ppt_file)
            # Temporäre Kopie erstellen und öffnen
            presentation.SaveCopyAs(temporary_ppt_file)
            newPres = application.Presentations.Open(temporary_ppt_file, False, False, False) #readonly, untitled, withwindow
            yield newPres
        finally:
            # Präsentation schließen
            if newPres:
                newPres.Saved = True
                newPres.Close()
            # Temporäre Datei löschen
            if temporary_ppt_file and os.path.exists(temporary_ppt_file):
                os.remove(temporary_ppt_file)

    @classmethod
    def save_slides(cls, application, slides, filepath, fileformat="ppt"):
        import os

        if fileformat == "png":
            filename, ext = os.path.splitext(filepath)
            for i, slide in enumerate(slides, start=1):
                slide.Export("{}_{}{}".format(filename, i, ext), "PNG", 2000)
            #open folder
            os.startfile(os.path.dirname(filepath))
        elif fileformat == "pdf":
            filename = os.path.basename(filepath)
            filename = os.path.splitext(filename)[0]
            with cls._create_temporary_copy(application, filename) as newPres:
                # Folien entfernen, die nicht ausgewählt waren
                cls._delete_unselected_slides(slides, newPres)
                # PDF erstellen und öffnen
                newPres.SaveCopyAs(filepath, 32) #ppSaveAsPDF
                os.startfile(filepath)
        else:
            # Kopie speichern und öffnen
            application.ActiveWindow.Presentation.SaveCopyAs(filepath)
            newPres = application.Presentations.Open(filepath, False, False, False) #readonly, untitled, withwindow

            # Folien entfernen, die nicht ausgewählt waren
            cls._delete_unselected_slides(slides, newPres)

            # Speichern und anzeigen
            newPres.Save()
            newPres.NewWindow()

    @classmethod
    def send_slides(cls, application, slides, filename, fileformat="ppt", remove_empty_sections=True, remove_all_sections=False, remove_author=False, remove_designs=False, remove_hidden=False):
        from bkt import dotnet
        Outlook = dotnet.import_outlook()

        try:
            oApp = Outlook.ApplicationClass()
            oMail = oApp.CreateItem(Outlook.OlItemType.olMailItem)
        except:
            bkt.message.error("Fehler beim Erstellen der E-Mail in Outlook!")
            return

        # Betreff
        oMail.Subject = filename

        # Kopie speichern und öffnen
        with cls._create_temporary_copy(application, filename) as newPres:

            if slides is not None:
                # Folien entfernen, die nicht ausgewählt waren
                cls._delete_unselected_slides(slides, newPres)
                newPres.Save()

            if remove_all_sections:
                # Alle Abschnitte entfernen
                sections = newPres.SectionProperties
                for i in reversed(range(sections.count)):
                    sections.Delete(i+1, 0) #index, deleteSlides=False
                newPres.Save()
            elif remove_empty_sections:
                # Leere Abschnitte entfernen
                sections = newPres.SectionProperties
                for i in reversed(range(sections.count)):
                    if sections.SlidesCount(i+1) == 0:
                        sections.Delete(i+1, 0) #index, deleteSlides=False
                newPres.Save()
            
            if remove_author:
                newPres.BuiltInDocumentProperties.item["author"].value = ''
                newPres.Save()
            
            if remove_designs:
                for design in list(iter(newPres.Designs)):
                    for cl in list(iter(design.SlideMaster.CustomLayouts)): #list(iter()) required as delete function will not work on all elements otherwise!
                        try:
                            cl.Delete()
                        except: #deletion fails if layout in use
                            continue
                    if design.SlideMaster.CustomLayouts.Count == 0:
                        try:
                            design.Delete()
                        except:
                            continue
                newPres.Save()
            
            if remove_hidden:
                for slide in list(iter(newPres.Slides)):
                    if slide.SlideShowTransition.Hidden == -1:
                        slide.Delete()
                newPres.Save()

            if fileformat != "pdf":
                # PPT anhängen
                oMail.Attachments.Add(newPres.FullName, Outlook.OlAttachmentType.olByValue)

            if fileformat != "ppt":
                # PDF exportieren und anhängen
                pdfFileName = os.path.splitext(newPres.FullName)[0] + ".pdf"
                #newPres.ExportAsFixedFormat(pdfFileNameRef, 2) #ppFixedFormatTypePDF #ValueError: Could not convert argument 0 for call to ExportAsFixedFormat.
                newPres.SaveCopyAs(pdfFileName, 32) #ppSaveAsPDF
                oMail.Attachments.Add(pdfFileName, Outlook.OlAttachmentType.olByValue)

            # Email anzeigen
            oMail.Display()


class SlideMenu(object):

    @classmethod
    def send_slides_dialog(cls, context):
        from ..dialogs.slides_send import SendWindow
        SendWindow.create_and_show_dialog(SendOrSaveSlides, context)

    @classmethod
    def sync_slides_dialog(cls, context):
        from ..dialogs.slides_sync import SlideSyncWindow
        SlideSyncWindow.create_and_show_dialog(context)


    @classmethod
    def save_slides_dialog(cls, context):
        from bkt import dotnet
        Forms = dotnet.import_forms()

        presentation = context.presentation
        slides = context.slides
        fileformats = [
            "PowerPoint (*.pptx;*.pptm;*.ppt)|*.pptx;*.pptm;*.ppt",
            "PDF (*.pdf)|*.pdf",
            "PNG (*.png)|*.png",
            "Alle Dateien (*.*)|*.*"
        ]

        fileDialog = Forms.SaveFileDialog()
        fileDialog.Filter = "|".join(fileformats)
        if presentation.Path:
            fileDialog.InitialDirectory = presentation.Path + '\\'
        fileDialog.FileName = SendOrSaveSlides.initial_file_name(presentation, slides)
        if len(slides) == 1:
            fileDialog.Title = "Ausgewählte Folie speichern unter"
        else:
            fileDialog.Title = str(len(slides)) + " ausgewählte Folien speichern unter"

        # Bei Abbruch ist Rückgabewert leer
        if not fileDialog.ShowDialog() == Forms.DialogResult.OK:
            return

        fileformat = ["ppt", "pdf", "png", "ppt"][fileDialog.FilterIndex-1]
        SendOrSaveSlides.save_slides(context.app, slides, fileDialog.FileName, fileformat)


    SLIDENUMBERING = 'Toolbox-SlideNumbering'

    @classmethod
    def add_slide_numbering(cls, slides, context):
        # Alle Slides durchlaufen
        for sld in slides:
            # msoTextOrientationHorizontal = 1
            shp = sld.shapes.AddTextbox(1 # msoTextOrientationHorizontal
                , 0, 0, 100, 100)
            shp.TextFrame.TextRange.Font.Size = 32
            shp.TextFrame.TextRange.Font.Bold = -1 # msoTrue
            shp.TextFrame.TextRange.Font.Color = 192 + 0 * 256 + 0 * 256**2
            shp.TextFrame.TextRange.ParagraphFormat.Alignment = 3 #ppAlignRight
            shp.TextFrame.TextRange.text = sld.SlideNumber #SlideIndex always starts with 1, but in PageSetup beginning can be changed so using SlideNumber
            shp.TextFrame.MarginBottom = 0
            shp.TextFrame.MarginTop = 0
            shp.TextFrame.MarginRight = 0
            shp.TextFrame.MarginLeft = 0
            shp.Left = context.app.ActivePresentation.PageSetup.SlideWidth - shp.width - 15
            shp.Top = 15
            shp.Tags.Add(cls.SLIDENUMBERING, cls.SLIDENUMBERING)


    @classmethod
    def remove_slide_numbering(cls, slides):
        for slide in slides:
            for shp in slide.shapes:
                # Shape mit SlideNumberTag loeschen
                if shp.Tags.Item(cls.SLIDENUMBERING) == cls.SLIDENUMBERING:
                    shp.Delete()
                    break

    @classmethod
    def toggle_slide_numbering(cls, context):
        has_numbering = False

        slides = context.app.ActivePresentation.Slides
        # Alle Shapes in allen Slides durchlaufen
        for sld in slides:
            for shp in sld.shapes:
                # Shape mit SlideNumberTag gefunden
                if shp.Tags.Item(cls.SLIDENUMBERING) == cls.SLIDENUMBERING:
                    has_numbering = True
                    break
            if has_numbering:
                break

        if has_numbering:
            cls.remove_slide_numbering(slides)
        else:
            cls.add_slide_numbering(slides, context)


    @classmethod
    def select_and_apply_theme(cls, context):
        from bkt import dotnet
        F = dotnet.import_forms()
        
        fileDialog = F.OpenFileDialog()
        fileDialog.Filter = "PowerPoint (*.pptx;*.ppt;*.pot;*.potx)|*.pptx;*.ppt;*.pot;*.potx|Alle Dateien (*.*)|*.*"
        if context.presentation.Path:
            fileDialog.InitialDirectory = context.presentation.Path + '\\'
        fileDialog.Title = "PowerPoint-Datei auswählen"

        if not fileDialog.ShowDialog() == F.DialogResult.OK:
            return

        filename = fileDialog.FileName
        try:
            context.presentation.ApplyTemplate(filename)
        except:
            logging.exception("error appyling theme %s", filename)
    
    
    @classmethod
    def remove_all(cls, context):
        from ..dialogs.slides_clean import SlideCleanWindow
        SlideCleanWindow.create_and_show_dialog(cls, context)

        
    @classmethod
    def _iterate_all_shapes(cls, context, groupitems=False):
        slides = context.app.ActivePresentation.Slides
        for slide in slides:
            for shape in slide.shapes:
                if groupitems and shape.Type == 6: #pplib.MsoShapeType['msoGroup']
                    for gShape in shape.GroupItems:
                        yield gShape
                else:
                    yield shape


    @classmethod
    def remove_transitions(cls, context):
        slides = context.app.ActivePresentation.Slides
        for slide in slides:
            slide.SlideShowTransition.EntryEffect = 0
    
    @classmethod
    def remove_animations(cls, context):
        for shape in cls._iterate_all_shapes(context):
            shape.AnimationSettings.Animate = 0

    @classmethod
    def remove_hidden_slides(cls, context):
        slides = context.app.ActivePresentation.Slides
        for slide in list(iter(slides)): #list(iter()) required as delete function will not work on all elements otherwise!
            if slide.SlideShowTransition.Hidden == -1:
                slide.Delete()

    @classmethod
    def remove_slide_notes(cls, context):
        slides = context.app.ActivePresentation.Slides
        for slide in slides:
            for shape in slide.NotesPage.Shapes:
                try:
                    if shape.PlaceholderFormat.type == 2: 
                        # ppt.PpPlaceholderType.ppPlaceholderBody.value__
                        shape.TextFrame.TextRange.Text = ""
                except SystemError:
                    # SystemError: System.Runtime.InteropServices.COMException (0x80048240): PlaceholderFormat.Type : Invalid request.  Shape is not a placeholder.
                    pass

    @classmethod
    def remove_slide_comments(cls, context):
        slides = context.app.ActivePresentation.Slides
        for slide in slides:
            for comment in list(iter(slide.Comments)): #list(iter()) required as delete function will not work on all elements otherwise!
                comment.Delete()

    @classmethod
    def remove_doublespaces(cls, context):
        slides = context.app.ActivePresentation.Slides
        for slide in slides:
            for textframe in pplib.iterate_shape_textframes(slide.shapes):
        # for shape in cls._iterate_all_shapes(context, groupitems=True):
        #     if shape.HasTextFrame == -1:
                found = True
                while found is not None:
                    found = textframe.TextRange.Replace("  ", " ")
    
    @classmethod
    def remove_empty_placeholders(cls, context):
        slides = context.app.ActivePresentation.Slides
        for sld in slides:
            for plh in list(iter(sld.Shapes.Placeholders)): #list(iter()) required as delete function will not work on all elements otherwise!
                if plh.HasTextFrame == -1 and plh.TextFrame2.HasText == 0:
                    #placeholder is a text placeholder and has no text. note: placeholder can also be a picture, table or diagram without text!
                    plh.Delete()

    @classmethod
    def blackwhite_gray_scale(cls, context):
        for shape in cls._iterate_all_shapes(context, groupitems=True):
            if shape.BlackWhiteMode == 1:
                shape.BlackWhiteMode = 2

    @classmethod
    def remove_author(cls, context):
        context.presentation.BuiltInDocumentProperties.item["author"].value = ''

    @classmethod
    def remove_unused_masters(cls, context, silent=True):
        deleted_layouts = 0
        unused_designs = []
        for design in context.presentation.Designs:
            for cl in list(iter(design.SlideMaster.CustomLayouts)): #list(iter()) required as delete function will not work on all elements otherwise!
                try:
                    name = cl.name
                    cl.Delete()
                    deleted_layouts += 1
                    logging.info("deleted custom layout %s", name)
                except SystemError: #deletion fails if layout in use
                    #SystemError: System.Runtime.InteropServices.COMException (0x80048240): Slide (unknown member) : Invalid request.  Can't delete master.
                    continue
            if design.SlideMaster.CustomLayouts.Count == 0:
                unused_designs.append(design)

        if silent:
            return
        
        unused_designs_len = len(unused_designs)
        if unused_designs_len > 0:
            if bkt.message.confirmation("Es wurden {} Folienlayouts gelöscht und {} Folienmaster sind nun ohne Layout. Sollen diese gelöscht werden?".format(deleted_layouts, unused_designs_len)):
                for design in unused_designs:
                    try:
                        name = design.name
                        design.Delete()
                        logging.info("deleted design %s", name)
                    except:
                        logging.exception("error deleting design")
            bkt.message("Leere Folienmaster wurden gelöscht!")
        else:
            bkt.message("Es wurden {} Folienlayouts gelöscht!".format(deleted_layouts))
    
    @classmethod
    def remove_unused_designs(cls, context, silent=True):
        deleted_designs = 0
        designs = context.presentation.designs
        #list incides of all designs
        unused_designs = list(range(1,designs.count+1))
        for slide in context.presentation.slides:
            try:
                #remove indices of used designs
                unused_designs.remove(slide.design.index)
            except ValueError: #index already removed
                pass
        
        #remove all remaining indices
        for i in reversed(unused_designs):
            try:
                name = designs[i].name
                designs[i].delete()
                deleted_designs += 1
                logging.info("deleted design %s", name)
            except:
                logging.exception("error deleting design")
        
        if not silent:
            bkt.message("Es wurden {} Folienmaster gelöscht!".format(deleted_designs))

    @classmethod
    def break_links(cls, context):
        for shape in cls._iterate_all_shapes(context, groupitems=True):
            try:
                pst = pplib.MsoShapeType
                if shape.Type in (pst["msoLinkedGraphic"], pst["msoLinkedOLEObject"], pst["msoLinkedPicture"], pst["msoLinked3DModel"]):
                    shape.LinkFormat.BreakLink()
            except:
                logging.exception("error breaking link")
    
    @staticmethod
    def _iter_all_layouts(context):
        for design in context.presentation.designs:
            if design.HasTitleMaster:
                yield design.TitleMaster
            yield design.SlideMaster
            for layout in design.SlideMaster.CustomLayouts:
                yield layout
    
    @classmethod
    def _iter_master_shapes(cls, context):
        for layout in cls._iter_all_layouts(context):
            for shape in layout.Shapes:
                if pplib.TagHelper.has_tag(shape, "THINKCELLSHAPEDONOTDELETE"):
                    continue
                yield shape
    
    @classmethod
    def toggle_hide_master_shapes(cls, context):
        visibility = None
        for s in cls._iter_master_shapes(context):
            if visibility is None:
                visibility = 0 if s.visible else -1
            s.visible = visibility
    
    @staticmethod
    def open_in_explorer(context):
        from os import startfile
        path = context.presentation.path
        if not path:
            return bkt.message.error("Präsentation ist nicht gespeichert!")
        if os.path.isdir(path):
            startfile(path)


class SlideShow(object):
    @classmethod
    def windowed_slideshow(cls, context):
        cls._slideshow(context, 2) #ppShowTypeWindow

    @classmethod
    def fullscreen_slideshow(cls, context):
        cls._slideshow(context, 1) #ppShowTypeSpeaker

    @classmethod
    def _slideshow(cls, context, show_type):
        #get slide (as later activewindow is not present anymore)
        slide = context.slide
        #use with-notation to avoid comrelease error
        with context.presentation.SlideShowSettings as sld_settings:
            #save current setting
            prev = sld_settings.ShowType
            #define type (windowed or fullscreen)
            sld_settings.ShowType = show_type
            #run slideshow
            sld_window = sld_settings.Run()
            #go to selected slide
            if slide:
                sld_window.View.GoToSlide(slide.SlideIndex)
                sld_window.view.LaserPointerEnabled = True
            #restore setting
            sld_settings.ShowType = prev


slides_menu = lambda: bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None, 
                children=[
                    bkt.ribbon.MenuSeparator(title="Layout"),
                    bkt.mso.control.SlideLayoutGallery,
                    bkt.mso.control.SlideReset,
                    bkt.mso.control.SectionMenu,
                    bkt.ribbon.Button(
                        id="slide_apply_theme",
                        label="Folienmaster aus Datei ersetzen…",
                        image_mso="SlideMasterMasterLayout",
                        supertip="Ersetzt den aktuellen Folienmaster (Templates und Design) in der Präsentation durch den Folienmaster aus der gewählten Datei.",
                        on_action=bkt.Callback(SlideMenu.select_and_apply_theme)
                    ),
                    bkt.ribbon.MenuSeparator(title="Agenda")
                    ] + agenda.agendamenu.children + [
                    bkt.ribbon.MenuSeparator(title="Funktionen"),
                    bkt.ribbon.Button(
                        id = 'slide_no',
                        label='Foliennummerierung ein-/ausblenden',
                        image_mso='NumberInsert',
                        #screentip="Foliennummerierung ein-/ausblenden",
                        supertip="Füge Foliennummerierungen ein, welche sich bei Umsortierung der Folien nicht ändern.\n\nHilfreich bei der Erfassung von Anmerkungen, wenn man während einer Diskussion des Foliensatzes Umsortierungen durchführt.",
                        on_action=bkt.Callback(SlideMenu.toggle_slide_numbering)
                    ),
                    bkt.ribbon.Button(
                        id = 'save_slides',
                        label='Ausgewählte Folien speichern',
                        image_mso='SaveSelectionToTextBoxGallery',
                        supertip="Speichert die ausgewählten Folien in einer neuen Präsentation, als PDF-Datei oder in PNG-Bilder.",
                        on_action=bkt.Callback(SlideMenu.save_slides_dialog)
                    ),
                    bkt.ribbon.Button(
                        id = 'send_slides',
                        label='Ausgewählte Folien senden',
                        image_mso='FileSendAsAttachment',
                        supertip="Sendet die ausgewählten Folien als Email-Anhang, wahlweise auch als PDF-Datei.",
                        on_action=bkt.Callback(SlideMenu.send_slides_dialog)
                    ),
                    bkt.ribbon.Button(
                        id = 'sync_slides',
                        label='Ausgewählte Folien angleichen',
                        image_mso='SlideShowShowPresentationOnGallery',
                        supertip="Gleicht alle Shapes auf den Folien der ersten selektierten Folie an anhand des Shape-Namens. Die Funktion ist nützlich wenn man ein Template erstellt, die mehrfach dupliziert und Änderungen danach synchronisieren möchte",
                        on_action=bkt.Callback(SlideMenu.sync_slides_dialog)
                    ),
                    bkt.ribbon.SplitButton(children=[
                        bkt.ribbon.Button(
                            id = 'slide_remove_all',
                            label='Slidedeck aufräumen…',
                            image_mso='SlideShowFromCurrent', #AcceptTask, SlideShowFromCurrent, FilePublishSlides
                            supertip="Zeigt Dialog zur Auswahl der anzuwendenden Funktionen.",
                            on_action=bkt.Callback(SlideMenu.remove_all)
                        ),
                        bkt.ribbon.Menu(label="Slidedeck aufräumen", supertip="Funktionen zum Aufräumen aller Folien der Präsentation", image_mso='SlideShowFromCurrent', children=[
                            bkt.ribbon.Button(
                                id = 'slide_remove_all2',
                                label='Funktionen auswählen…',
                                image_mso='SlideShowFromCurrent', #AcceptTask, SlideShowFromCurrent, FilePublishSlides
                                supertip="Zeigt Dialog zur Auswahl der anzuwendenden Funktionen.",
                                on_action=bkt.Callback(SlideMenu.remove_all)
                            ),
                            bkt.ribbon.MenuSeparator(title="Inhalte"),
                            bkt.ribbon.Button(
                                id = 'slide_remove_hidden_slides',
                                label='Ausgeblendete Slides entfernen',
                                image_mso='SlideHide',
                                supertip="Lösche alle ausgeblendeten Slides im gesamten Foliensatz.",
                                on_action=bkt.Callback(SlideMenu.remove_hidden_slides)
                            ),
                            bkt.ribbon.Button(
                                id = 'slide_remove_notes',
                                label='Notizen entfernen',
                                image_mso='SpeakerNotes',
                                supertip="Lösche alle Notizen im gesamten Foliensatz.",
                                on_action=bkt.Callback(SlideMenu.remove_slide_notes)
                            ),
                            bkt.ribbon.Button(
                                id = 'slide_remove_comments',
                                label='Kommentare entfernen',
                                image_mso='ReviewDeleteComment',
                                supertip="Lösche alle Kommentare im gesamten Foliensatz.",
                                on_action=bkt.Callback(SlideMenu.remove_slide_comments)
                            ),
                            bkt.ribbon.Button(
                                id = 'presentation_remove_author',
                                label='Autor entfernen',
                                image_mso='ContactPictureMenu',
                                supertip="Autor aus den Dokumenteneigenschaften entfernen.",
                                on_action=bkt.Callback(SlideMenu.remove_author)
                            ),
                            bkt.ribbon.Button(
                                id = 'presentation_break_links',
                                label='Externe Verknüpfungen entfernen',
                                image_mso='HyperlinkRemove',
                                supertip="Hebt den Link von verknüpften Objekten (bspw. Bilder und OLE-Objekten) auf.",
                                on_action=bkt.Callback(SlideMenu.break_links)
                            ),
                            bkt.ribbon.MenuSeparator(title="Animationen"),
                            bkt.ribbon.Button(
                                id = 'slide_remove_transitions',
                                label='Folienübergänge entfernen',
                                image_mso='AnimationTransitionGallery',
                                supertip="Lösche alle Übergänge zwischen Folien.",
                                on_action=bkt.Callback(SlideMenu.remove_transitions)
                            ),
                            bkt.ribbon.Button(
                                id = 'slide_remove_animation',
                                label='Shapeanimationen entfernen',
                                image_mso='AnimationGallery',
                                supertip="Lösche alle Shape-Animationen im gesamten Foliensatz.",
                                on_action=bkt.Callback(SlideMenu.remove_animations)
                            ),
                            bkt.ribbon.MenuSeparator(title="Format bereinigen"),
                            bkt.ribbon.Button(
                                id = 'slide_grayscale',
                                label='Automatischen Schwarz-/Weiß-Modus deaktivieren',
                                image_mso='BlackAndWhiteGrayscale',
                                supertip="Ersetze den Schwarz-/Weiß-Modus 'Automatisch' durch 'Graustufen'.",
                                on_action=bkt.Callback(SlideMenu.blackwhite_gray_scale)
                            ),
                            bkt.ribbon.Button(
                                id = 'slide_remove_doublespaces',
                                label='Doppelte Leerzeichen entfernen',
                                image_mso='ParagraphMarks',
                                supertip="Lösche alle doppelten Leerzeichen im gesamten Foliensatz.",
                                on_action=bkt.Callback(SlideMenu.remove_doublespaces)
                            ),
                            bkt.ribbon.Button(
                                id = 'slide_remove_empty_placeholders',
                                label='Leere Platzhalter entfernen',
                                image_mso='HeaderFooterRemoveHeaderWord',
                                supertip="Lösche leere Platzhalter-Textboxen im gesamten Foliensatz.",
                                on_action=bkt.Callback(SlideMenu.remove_empty_placeholders)
                            ),
                            bkt.ribbon.MenuSeparator(title="Folienmaster"),
                            bkt.ribbon.Button(
                                id = 'slide_remove_unused_masters',
                                label='Nicht genutzte Folienlayouts entfernen',
                                image_mso='SlideDelete',
                                supertip="Lösche alle nicht verwendeten Folienmaster-Layouts sowie leere Folienmaster (Designs).",
                                on_action=bkt.Callback(lambda context: SlideMenu.remove_unused_masters(context, False))
                            ),
                            bkt.ribbon.Button(
                                id = 'slide_remove_unused_designs',
                                label='Nicht genutzte Folienmaster entfernen',
                                image_mso='SlideDelete',
                                supertip="Lösche alle nicht verwendeten Folienmaster (Designs).",
                                on_action=bkt.Callback(lambda context: SlideMenu.remove_unused_designs(context, False))
                            ),
                        ]),
                    ]),
                    language.sprachen_menu,
                    bkt.ribbon.MenuSeparator(title="Ansicht"),
                    bkt.ribbon.Menu(label="Masteransichten", supertip="Umschalten auf verschiedene Master-Ansichten", image_mso='GroupPresentationViews', children=[
                        bkt.mso.control.ViewSlideMasterView(show_label=True),
                        bkt.mso.control.ViewHandoutMasterView(show_label=True),
                        bkt.mso.control.ViewNotesMasterView(show_label=True),
                    ]),
                    bkt.ribbon.Menu(label="Farbe/Graustufen", supertip="Umschalten auf verschiedene Farbmodi", image_mso='ColorGrayscaleMenu', children=[
                        bkt.mso.control.ViewDisplayInColor(show_label=True),
                        bkt.mso.control.ViewDisplayInGrayscale(show_label=True),
                        bkt.mso.control.ViewDisplayInPureBlackAndWhite(show_label=True),
                    ]),
                    bkt.ribbon.MenuSeparator(),
                    bkt.mso.control.GuidesShowHide(show_label=True),
                    bkt.ribbon.Button(
                        label="Template-Shapes ein-/ausblenden",
                        image_mso='SlideHide',
                        supertip="Alle Shapes im Folienmaster ein- und ausblenden, um ungestört und vertraulich an Folien arbeiten zu können.",
                        on_action=bkt.Callback(SlideMenu.toggle_hide_master_shapes)
                    ),
                    bkt.ribbon.Button(
                        label="Im Explorer öffnen",
                        image_mso='OpenFolder',
                        supertip="Ordner der aktuellen Präsentation im Explorer öffnen.",
                        on_action=bkt.Callback(SlideMenu.open_in_explorer)
                    ),
                    bkt.ribbon.MenuSeparator(title="Bildschirmpräsentation"),
                    bkt.ribbon.Button(
                        id="slide_windowed_slideshow",
                        image_mso="SlideShowInAWindow",
                        label="Im Fenster starten",
                        supertip="Startet eine Bilschirmpräsentation im Fenster beginnend mit der aktuellen Folie und aktiviertem Laserpointer.",
                        on_action=bkt.Callback(SlideShow.windowed_slideshow),
                    ),
                    bkt.ribbon.Button(
                        id="slide_fullscreen_slideshow",
                        image_mso="SlideShowFromCurrent",
                        label="Im Vollbild starten",
                        supertip="Startet eine Bilschirmpräsentation im Vollbild beginnend mit der aktuellen Folie und aktiviertem Laserpointer.",
                        on_action=bkt.Callback(SlideShow.fullscreen_slideshow),
                    ),
                ]
            )