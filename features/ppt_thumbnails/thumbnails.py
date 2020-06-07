# -*- coding: utf-8 -*-
'''
Created on 2017-07-24
@author: Florian Stallmann
'''

from __future__ import absolute_import

import os #required for relative paths and tempfile removal
import tempfile
import logging

from contextlib import contextmanager
from System import Array #required to create ShapeRanges
from System.Windows import Visibility

import bkt
import bkt.library.powerpoint as pplib

import bkt.dotnet as dotnet
Forms = dotnet.import_forms() #required to read clipboard


PASTE_DATATYPE_BTM = 1 #ppPasteBitmap
PASTE_DATATYPE_EMF = 2 #ppPasteEnhancedMetafile
PASTE_DATATYPE_PNG = 6 #ppPastePNG

#mapping of paste types (old method) to export functions (slide.export requires filter-name, shape.export requires ppShapeFormat)
DATATYPE_MAPPING = {
    PASTE_DATATYPE_BTM: (3, "BMP"), #ppShapeFormatBMP, filter-name
    PASTE_DATATYPE_EMF: (5, "EMF"), #ppShapeFormatEMF, filter-name
    PASTE_DATATYPE_PNG: (2, "PNG"), #ppShapeFormatPNG, filter-name
}

BKT_THUMBNAIL = "BKT_THUMBNAIL"

USE_RELATIVE_PATHS = True

class ThumbnailerTags(pplib.BKTTag):
    TAG_NAME = BKT_THUMBNAIL

    @property
    def is_thumbnail(self):
        return "slide_id" in self.data

    def set_thumbnail(self, slide_id, slide_path, data_type=PASTE_DATATYPE_PNG, content_only=False, shape_id=None):
        self.data["slide_id"] = slide_id
        self.data["slide_path"] = slide_path
        self.data["data_type"] = data_type
        self.data["content_only"] = content_only #=exclude placeholder shapes
        if shape_id is not None:
            self.data["shape_id"] = shape_id


class Thumbnailer(object):
    #copied_slide_id = None
    #copied_slide_path = None

    @classmethod
    def slides_copy(cls, presentation, slides):
        #cls.copied_slide_id = slide.SlideId
        #cls.copied_slide_path = presentation.FullName
        #slide.Copy()
        cls.set_clipboard_data([slide.SlideId for slide in slides], presentation.FullName)

    @classmethod
    def shape_copy(cls, presentation, slide, shape):
        cls.set_clipboard_data([(slide.SlideId, shape.Id)], presentation.FullName)

    @classmethod
    def _get_presentation(cls, application, path, silent=True):
        logging.debug("Thumbnails: get presentation for path %s", path)
        if path == "CURRENT" or path == application.ActivePresentation.FullName:
            pres = application.ActivePresentation
            close_afterwards = False
            logging.debug("Thumbnails: return current presentation")
        else:
            #convert relative to absolute paths
            if not os.path.isabs(path):
                path = os.path.normpath(os.path.join(application.ActivePresentation.Path, path))
                logging.debug("Thumbnails: relative path converted to %s", path)
            try:
                #app.presentations can be used using a full path, but it fails if the path contains special characters, so fallback to filename
                try:
                    pres = application.Presentations[path]
                except:
                    basename = os.path.basename(path)
                    pres = application.Presentations[basename]
                    #different open files might have the same filename
                    if pres.FullName != path:
                        raise IndexError("deviating path. fallback to open presentation.")
                close_afterwards = False
                logging.debug("Thumbnails: return already open presentation")
            except:
                if silent:
                    pres = application.Presentations.Open(path, True, False, False) #Readonly, Untitled, WithWindow
                else:
                    pres = application.Presentations.Open(path)
                close_afterwards = True
                logging.debug("Thumbnails: open and return presentation")

        return pres, close_afterwards

    @classmethod
    @contextmanager
    def find_and_export_object(cls, application, slide_id, slide_path, content_only=False, shape_id=None, data_type=None):
        #avoid referenced before assignment error in finally clause
        close = None
        tmpfile = None
        try:
            pres, close = cls._get_presentation(application, slide_path)

            slide = pres.Slides.FindBySlideId(slide_id)
            filetype = DATATYPE_MAPPING.get(data_type, PASTE_DATATYPE_PNG)
            tmpfile = os.path.join(tempfile.gettempdir(), "bkt-thumbnail-tempfile."+filetype[1])

            if shape_id is None and not content_only:
                if data_type == PASTE_DATATYPE_PNG:
                    slide.Export(tmpfile, filetype[1], 2000)
                else:
                    slide.Export(tmpfile, filetype[1])
                # slide.Copy()
            elif content_only:
                shpr = cls._find_content_shapes(slide)
                if shpr.Count == 0:
                    raise ValueError("empty slide")
                shpr.Export(tmpfile, filetype[0]) 
                # shpr.Copy()
            else:
                shp = cls._find_by_shape_id(slide, shape_id)
                shp.Export(tmpfile, filetype[0]) 
                # shp.Copy()

            yield tmpfile

        except EnvironmentError:
            logging.exception("slide id not found")
            raise IndexError("slide id not found")
        except IndexError:
            logging.exception("shape id not found")
            raise IndexError("shape id not found")
        
        finally:
            if close:
                pres.Close()
            if tmpfile and os.path.exists(tmpfile):
                os.remove(tmpfile)

    @classmethod
    def _find_by_shape_id(cls, slide, shape_id):
        for shp in slide.Shapes:
            if shape_id == shp.Id:
                return shp
        raise IndexError("shape not found")

    @classmethod
    def _find_content_shapes(cls, slide):
        shape_indices = []
        for shape_index, shape in enumerate(slide.Shapes, start=1):
            if shape.type != 14 and shape.visible == -1: # shape is not a placeholder and visible
                shape_indices.append(shape_index)
        return pplib.shape_indices_on_slide(slide, shape_indices)
        # return slide.Shapes.Range(Array[int](shape_indices))

    @classmethod
    def remain_position_and_zorder(cls, orig_shp, new_shp):
        new_shp.LockAspectRatio = 0 #msoFalse
        new_shp.Top, new_shp.Left = orig_shp.Top, orig_shp.Left
        new_shp.Rotation = orig_shp.Rotation
        new_shp.Width, new_shp.Height = orig_shp.Width, orig_shp.Height
        new_shp.LockAspectRatio = orig_shp.LockAspectRatio
        while new_shp.ZOrderPosition > orig_shp.ZOrderPosition:
            new_shp.ZOrder(3) #msoSendBackward

    @classmethod
    def reset_aspect_ratio(cls, shape):
        height = shape.Height
        shape.ScaleHeight(1, True)
        shape.ScaleWidth(1, True)
        shape.Height = height

    @classmethod
    def slide_paste(cls, application, data_type=PASTE_DATATYPE_PNG, content_only=False):
        if not cls.has_clipboard_data():
            return

        data = cls.get_clipboard_data(application)
        # cur_slide = application.ActiveWindow.View.Slide
        cur_slide = application.ActiveWindow.Selection.SlideRange[1]
        # cur_shapes = cur_slide.Shapes.Count
        pasted_shapes = 0
        for slide_id in data["slide_ids"]:
            if type(slide_id) == tuple:
                shape_id = slide_id[1]
                slide_id = slide_id[0]
            else:
                shape_id = None
            try:
                try:
                    #Copy
                    with cls.find_and_export_object(application, slide_id, data["slide_path"], content_only, shape_id, data_type) as filename:
                        lefttop = 200+pasted_shapes*20
                        shape = cur_slide.Shapes.AddPicture(filename, 0, -1, lefttop, lefttop)
                        pasted_shapes += 1

                except Exception as e:
                    # bkt.helpers.exception_as_message()
                    bkt.message.error("Fehler! Referenz nicht gefunden.\n\n{}".format(e), "BKT: Thumbnails")
                    logging.exception("Thumbnails: Error finding slide reference!")
                    continue
                #Paste
                # shape = cur_slide.Shapes.PasteSpecial(Datatype=data_type)
                # pasted_shapes += 1
                #Save tags
                # shape = application.ActiveWindow.Selection.ShapeRange(1)
                with ThumbnailerTags(shape.Tags) as tags:
                    tags.set_thumbnail(slide_id, data["slide_path"], data_type, content_only, shape_id)
                shape.Tags.Add(bkt.contextdialogs.BKT_CONTEXTDIALOG_TAGKEY, BKT_THUMBNAIL)
            except Exception as e:
                #bkt.helpers.exception_as_message()
                bkt.message.error("Fehler! Thumbnail konnte nicht im gewählten Format eingefügt werden.\n\n{}".format(e), "BKT: Thumbnails")
                logging.exception("Thumbnails: Error pasting slide!")
        
        # select pasted shapes
        if pasted_shapes > 0:
            # cur_slide.Shapes.Range(Array[int](range(cur_shapes+1, cur_slide.Shapes.Count+1))).Select()
            pplib.last_n_shapes_on_slide(cur_slide, pasted_shapes).Select()
        
        #Restore clipboard
        # cls.set_clipboard_data(**data)

    @classmethod
    def replace_ref(cls, shape, application):
        data = cls.get_clipboard_data(application)
        with ThumbnailerTags(shape.Tags) as tags:
            tags["slide_path"] = data["slide_path"]

            if type(data["slide_ids"][0]) == tuple: #tuple of (slide_id, shape_id)
                tags["slide_id"] = data["slide_ids"][0][0]
                tags["shape_id"] = data["slide_ids"][0][1]
            else:
                tags["slide_id"] = data["slide_ids"][0]

        cls.shape_refresh(shape, application)
        
        #Restore clipboard
        # cls.set_clipboard_data(**data)

    @classmethod
    def replace_file_ref(cls, shape, application):
        fileDialog = Forms.OpenFileDialog()
        fileDialog.Filter = "PowerPoint (*.pptx;*.pptm;*.ppt)|*.pptx;*.pptm;*.ppt|Alle Dateien (*.*)|*.*"
        if application.ActiveWindow.Presentation.Path:
            fileDialog.InitialDirectory = application.ActiveWindow.Presentation.Path + '\\'
        fileDialog.Title = "Neue PowerPoint-Datei auswählen"

        # fileDialog = application.FileDialog(1) #msoFileDialogOpen
        # fileDialog.InitialFileName = application.ActiveWindow.Presentation.Path
        # fileDialog.Title = "Neue Datei auswählen"

        # Bei Abbruch ist Rückgabewert leer
        if not fileDialog.ShowDialog() == Forms.DialogResult.OK:
            return

        path = cls._prepare_path(application, fileDialog.FileName)

        with ThumbnailerTags(shape.Tags) as tags:
            tags["slide_path"] = path

        cls.shape_refresh(shape, application)

    @classmethod
    def goto_ref(cls, shape, application):
        with ThumbnailerTags(shape.Tags) as tags:
            slide_id = tags["slide_id"]
            slide_path = tags["slide_path"]
            try:
                pres, _ = cls._get_presentation(application, slide_path, False)
                
                #bring window to front
                if pres.Windows.Count > 0:
                    pres.Windows[1].Activate()
                else:
                    pres.NewWindow()

                try:
                    slide = pres.Slides.FindBySlideId(slide_id)
                    slide.Select()
                    
                    if "shape_id" in tags.data and tags["shape_id"] is not None:
                        try:
                            shp = cls._find_by_shape_id(slide, tags["shape_id"])
                            shp.Select()
                        except:
                            bkt.message.error("Fehler! Shape in der referenzierten Präsentation nicht gefunden.", "BKT: Thumbnails")
                except:
                    bkt.message.error("Fehler! Folie in der referenzierten Präsentation nicht gefunden.", "BKT: Thumbnails")
            except:
                if bkt.message.confirmation("Fehler! Referenzierte Präsentation '%s' nicht gefunden. Neue Datei auswählen?" % slide_path, "BKT: Thumbnails", icon=bkt.MessageBox.WARNING):
                    cls.replace_file_ref(shape, application)

    @classmethod
    def presentation_refresh(cls, application, presentation):
        cls.slide_refresh(application, presentation.slides)

    @classmethod
    def slide_refresh(cls, application, slides):
        thumbs = []
        for sld in slides:
            for shp in sld.shapes:
                if cls.is_thumbnail(shp):
                    thumbs.append(shp)

        if len(thumbs) == 0:
            bkt.message.warning("Keine Folien-Thumbnails gefunden.", "BKT: Thumbnails")
            return

        cls.shapes_refresh(thumbs, application)

    @classmethod
    def shapes_refresh(cls, shapes, application):
        err_counter = 0
        for shp in shapes:
            try:
                cls._shape_refresh(shp, application) #FIXME: currently file is opened for each thumbnail, can be improved for better performance
            except:
                cls._mark_erroneous_shape(shp)
                err_counter += 1
                # bkt.helpers.exception_as_message()
        if err_counter > 0:
            bkt.message.warning("Es wurde/n %r Folien-Thumbnail/s aktualisiert, aber %r Folien-Thumbnail/s konnten wegen eines Fehlers nicht aktualisiert werden. Die fehlerhaften Thumbnails wurden mit dem Text 'BKT THUMB UPDATE FAILED' markiert." % (len(shapes)-err_counter, err_counter), "BKT: Thumbnails")
        else:
            bkt.message("Es wurde/n %r Folien-Thumbnail/s aktualisiert." % len(shapes), "BKT: Thumbnails")

    @classmethod
    def shape_refresh(cls, shape, application):
        try:
            return cls._shape_refresh(shape, application)
        except IndexError:
            bkt.message.error("Fehler! Folien-Referenz nicht gefunden.")
        except ValueError:
            bkt.message.error("Fehler! Folie hat keinen Inhalt.")
        except IOError:
            if bkt.message.confirmation("Fehler! Präsentation aus Folien-Referenz nicht gefunden. Neue Datei auswählen?", "BKT: Thumbnails", icon=bkt.MessageBox.WARNING):
                cls.replace_file_ref(shape, application)
        except Exception as e:
            bkt.message.error("Fehler! Thumbnail konnte nicht aktualisiert werden.\n\n{}".format(e), "BKT: Thumbnails")
            logging.exception("Thumbnails: Error updating thumbnail!")

    @classmethod
    def _shape_refresh(cls, shape, application):
        with ThumbnailerTags(shape.Tags) as tags_old:
            #Copy
            with cls.find_and_export_object(application, **tags_old.data) as filename:
                # cls.find_and_copy_object(application, tags_old["slide_id"], tags_old["slide_path"])
                #Paste (shapes.Parent = slide)
                # new_shp = shape.Parent.Shapes.PasteSpecial(Datatype=tags_old["data_type"]).Item(1)
                new_shp = shape.Parent.Shapes.AddPicture(filename, 0, -1, 200, 200)
            #Duplicate tags
            with ThumbnailerTags(new_shp.Tags) as tags_new:
                tags_new.set_thumbnail(**tags_old.data)

        #handle thumbnail in group and in placeholders
        group = None
        if pplib.shape_is_group_child(shape):
            group = pplib.GroupManager(shape.ParentGroup)
            group.ungroup()

        elif shape.Type == pplib.MsoShapeType["msoPlaceholder"]:
            logging.warning("Thumbnails: Update of placeholder not possible!")
            #FIXME: any way to update within placeholder?

        new_shp.Tags.Add(bkt.contextdialogs.BKT_CONTEXTDIALOG_TAGKEY, BKT_THUMBNAIL)

        new_shp.PictureFormat.crop.ShapeHeight = shape.PictureFormat.crop.ShapeHeight 
        new_shp.PictureFormat.crop.ShapeWidth  = shape.PictureFormat.crop.ShapeWidth  
        new_shp.PictureFormat.crop.ShapeTop    = shape.PictureFormat.crop.ShapeTop    
        new_shp.PictureFormat.crop.ShapeLeft   = shape.PictureFormat.crop.ShapeLeft   
    
        new_shp.PictureFormat.crop.PictureHeight  = shape.PictureFormat.crop.PictureHeight
        new_shp.PictureFormat.crop.PictureWidth   = shape.PictureFormat.crop.PictureWidth
        new_shp.PictureFormat.crop.PictureOffsetX = shape.PictureFormat.crop.PictureOffsetX
        new_shp.PictureFormat.crop.PictureOffsetY = shape.PictureFormat.crop.PictureOffsetY

        cls.remain_position_and_zorder(shape, new_shp)
        shape.PickUp()
        new_shp.Apply()

        shape_name = shape.Name

        #handle thumbnail in group (part 2)
        if group:
            group.select()
            shape.Delete()
            new_shp.Select(False)
            group.regroup(application.ActiveWindow.Selection.ShapeRange)
            new_shp.Select()
        else:
            shape.Delete()
            new_shp.Name = shape_name
            new_shp.Select()

        return new_shp

    @classmethod
    def _mark_erroneous_shape(cls, shape):
        txt = shape.Parent.Shapes.AddTextbox(1 # msoTextOrientationHorizontal
                , shape.Left, shape.Top, shape.Width, shape.Height)
        txt.TextFrame.TextRange.Font.Bold = -1 # msoTrue
        txt.TextFrame.TextRange.Font.Color = 192 + 0 * 256 + 0 * 256**2
        txt.TextFrame.TextRange.Text = "BKT THUMB UPDATE FAILED"
        txt.TextFrame.MarginBottom = 0
        txt.TextFrame.MarginTop = 0
        txt.TextFrame.MarginRight = 0
        txt.TextFrame.MarginLeft = 0

    @classmethod
    def set_clipboard_data(cls, slide_ids, slide_path):
        return Forms.Clipboard.SetData(BKT_THUMBNAIL, (slide_ids, slide_path))

    @classmethod
    def get_clipboard_data(cls, application):
        if Forms.Clipboard.ContainsData(BKT_THUMBNAIL):
            logging.info("Thumbnails: Get thumbnail from BKT_THUMBNAIL clipboard data")
            try:
                data = Forms.Clipboard.GetData(BKT_THUMBNAIL)
                #bruteforce method to convert data into correct type
                data = tuple(data)
                data = (list(data[0]), unicode(data[1]))
            except:
                raise ValueError("Invalid clipboard format")
        
        else:
            logging.info("Thumbnails: Get thumbnail from OLE object in clipboard")
            try:
                shp = application.ActiveWindow.Selection.SlideRange[1].Shapes.PasteSpecial(Datatype=10, Link=True) #ppPasteOLEObject
                try:
                    shp = shp[1] #PasteSpecial might return a shaperange with 2 references to the same shape
                except:
                    pass
                if not shp.OLEFormat.ProgID.startswith("PowerPoint"):
                    raise Exception("Invalid program")
                path,slideid = shp.LinkFormat.SourceFullName.split("!")
                data = ([slideid], path)
            except:
                logging.exception("Thumbnails: Invalid clipboard data!")
                raise ValueError("Invalid clipboard format")
            finally:
                if shp:
                    shp.Delete()
        
        #check consistency of clipboard data
        # if type(data) != tuple or len(data) != 2 or type(data[0]) != list or type(data[1]) != str:
        #     raise ValueError("Invalid clipboard data")
        
        path = cls._prepare_path(application, data[1])
        return {"slide_ids": data[0], "slide_path": path}

    @classmethod
    def _prepare_path(cls, application, path):
        drive1, _ = os.path.splitdrive(path)
        drive2, _ = os.path.splitdrive(application.ActivePresentation.FullName)
        if path == application.ActivePresentation.FullName:
            path = "CURRENT"
        elif USE_RELATIVE_PATHS and drive1 != '' and drive1 == drive2: #same drive -> use relative path
            path = os.path.relpath(path, application.ActivePresentation.Path)
        else:
            path = os.path.normpath(path)
        
        return path


    @classmethod
    def has_clipboard_data(cls):
        return Forms.Clipboard.ContainsData(BKT_THUMBNAIL) or (Forms.Clipboard.ContainsData("PowerPoint 12.0 Internal Slides") and Forms.Clipboard.ContainsData("Link Source")) #"PowerPoint 14.0 Slides Package"
        # return Forms.Clipboard.ContainsData(BKT_THUMBNAIL)

    @classmethod
    def enabled_paste(cls):
        return cls.has_clipboard_data()
        #return Forms.Clipboard.ContainsImage()

    @classmethod
    def enabled_slideref(cls):
        return cls.has_clipboard_data()
        #return cls.copied_slide_id != None

    @classmethod
    def is_thumbnail(cls, shape):
        return pplib.TagHelper.has_tag(shape, BKT_THUMBNAIL)

    @classmethod
    def unset_thumbnail(cls, shape):
        if bkt.message.confirmation("Dies löscht dauerhaft die Folien-Referenz und damit die Möglichkeit der Aktualisierung des Thumbnails.", "BKT: Thumbnails"):
            shape.Tags.Delete(BKT_THUMBNAIL)

    @classmethod
    def get_quality(cls, shape):
        with ThumbnailerTags(shape.Tags) as tags:
            return tags["data_type"]

    @classmethod
    def set_quality(cls, shape, application, quality):
        with ThumbnailerTags(shape.Tags) as tags:
            tags["data_type"] = quality
        cls.shape_refresh(shape, application)

    @classmethod
    def get_content_only(cls, shape):
        with ThumbnailerTags(shape.Tags) as tags:
            if "content_only" in tags.data:
                return tags["content_only"]
            else:
                return False

    @classmethod
    def set_content_only(cls, shape, application, content_only):
        with ThumbnailerTags(shape.Tags) as tags:
            tags["content_only"] = content_only
        new_shp = cls.shape_refresh(shape, application)
        if new_shp:
            cls.reset_aspect_ratio(new_shp)


thumbnail_gruppe = bkt.ribbon.Group(
    id="bkt_slidethumbnails_group",
    label='Folien-Thumbnails',
    supertip="Ermöglicht das Einfügen von aktualisierbaren Folien-Thumbnails. Das Feature `ppt_thumbnails` muss installiert sein.",
    image_mso='PasteAsPicture',
    children = [
        bkt.ribbon.Button(
            id = 'slide_copy',
            label="Folie(n) als Thumbnail kopieren",
            show_label=True,
            image_mso='Copy',
            supertip="Aktuelle Folie zum Erstellen vom aktualisierbaren Folien-Thumbnails kopieren.",
            on_action=bkt.Callback(Thumbnailer.slides_copy, presentation=True, slides=True),
        ),
        # bkt.ribbon.Button(
        #     id = 'shape_copy',
        #     label="Shape als Thumbnail kopieren",
        #     show_label=True,
        #     image_mso='Copy',
        #     supertip="Aktuelle Folie zum Erstellen vom aktualisierbaren Folien-Thumbnails kopieren.",
        #     on_action=bkt.Callback(Thumbnailer.shape_copy, presentation=True, slide=True, shape=True),
        # ),
        bkt.ribbon.SplitButton(
            get_enabled = bkt.Callback(Thumbnailer.enabled_paste),
            children=[
                bkt.ribbon.Button(
                    id = 'slide_paste',
                    label="Folien-Thumbnail einfügen",
                    show_label=True,
                    image_mso='PasteAsPicture',
                    supertip="Kopierte Folie als aktualisierbares Thumbnail mit Referenz auf Originalfolie einfügen.\n\nIst die Originalfolie aus einer anderen Datei im gleichen Verzeichnis, wird nur der Dateiname hinterlegt, anderenfalls wird der absolute Pfad hinterlegt und die Originaldatei darf nicht verschoben werden.",
                    on_action=bkt.Callback(Thumbnailer.slide_paste, application=True),
                    # get_enabled = bkt.Callback(Thumbnailer.enabled_paste),
                ),
                bkt.ribbon.Menu(label="Einfügen-Menü", supertip="Einfüge-Optionen für aktualisierbare Folien-Thumbnails", children=[
                    bkt.ribbon.Button(
                        id = 'slide_paste_png',
                        label="Folien-Thumbnail als PNG einfügen",
                        show_label=True,
                        #image_mso='PasteAsPicture',
                        supertip="Kopierte Folie als aktualisierbares Thumbnail im PNG-Format mit Referenz auf Originalfolie einfügen.",
                        on_action=bkt.Callback(lambda application: Thumbnailer.slide_paste(application, PASTE_DATATYPE_PNG), application=True),
                        # get_enabled = bkt.Callback(Thumbnailer.enabled_paste),
                    ),
                    bkt.ribbon.Button(
                        id = 'slide_paste_btm',
                        label="Folien-Thumbnail als Bitmap einfügen",
                        show_label=True,
                        image_mso='PasteAsPicture',
                        supertip="Kopierte Folie als aktualisierbares Thumbnail im Bitmap-Format mit Referenz auf Originalfolie einfügen.",
                        on_action=bkt.Callback(lambda application: Thumbnailer.slide_paste(application, PASTE_DATATYPE_BTM), application=True),
                        # get_enabled = bkt.Callback(Thumbnailer.enabled_paste),
                    ),
                    bkt.ribbon.Button(
                        id = 'slide_paste__emf',
                        label="Folien-Thumbnail als Vektor (EMF) einfügen",
                        show_label=True,
                        #image_mso='PasteAsPicture',
                        supertip="Kopierte Folie als aktualisierbares Thumbnail im Vektor-Format (Enhanced Metafile) mit Referenz auf Originalfolie einfügen.",
                        on_action=bkt.Callback(lambda application: Thumbnailer.slide_paste(application, PASTE_DATATYPE_EMF), application=True),
                        # get_enabled = bkt.Callback(Thumbnailer.enabled_paste),
                    ),
                ])
            ]
        ),
        bkt.ribbon.SplitButton(
            children=[
                bkt.ribbon.Button(
                    id = 'slide_refresh',
                    label="Alle Thumbnails aktualisieren",
                    show_label=True,
                    image_mso='PictureChange',
                    supertip="Alle Folien-Thumbnails auf den ausgewählten Folien aktualisieren. Das Thumbnail muss vorher mit dieser Funktion eingefügt worden sein. Stammt die Folie aus einer anderen Datei, wird diese automatisch kurzzeitig geöffnet.",
                    on_action=bkt.Callback(Thumbnailer.slide_refresh, application=True, slides=True),
                ),
                bkt.ribbon.Menu(label="Aktualisieren-Menü", supertip="Aktualisierung der Folien-Thumbnails auf dieser Folie oder in der ganzen Präsentation", item_size="large", children=[
                    bkt.ribbon.Button(
                        id = 'slide_refresh2',
                        label="Thumbnails auf Folie/n aktualisieren",
                        description="Alle Thumbnails auf aktueller bzw. ausgewählten Folie/n aktualisieren",
                        # show_label=True,
                        image_mso='PictureChange',
                        supertip="Alle Folien-Thumbnails auf den ausgewählten Folien aktualisieren. Das Thumbnail muss vorher mit dieser Funktion eingefügt worden sein. Stammt die Folie aus einer anderen Datei, wird diese automatisch kurzzeitig geöffnet.",
                        on_action=bkt.Callback(Thumbnailer.slide_refresh, application=True, slides=True),
                    ),
                    bkt.ribbon.MenuSeparator(),
                    bkt.ribbon.Button(
                        id = 'presentation_refresh',
                        label="Thumbnails in Präsentation aktualisieren",
                        description="Alle Thumbnails in der gesamten Präsentation aktualisieren",
                        # show_label=True,
                        #image_mso='PictureChange',
                        supertip="Alle Folien-Thumbnails in der Präsentation. Das Thumbnail muss vorher mit dieser Funktion eingefügt worden sein. Stammt die Folie aus einer anderen Datei, wird diese automatisch kurzzeitig geöffnet.",
                        on_action=bkt.Callback(Thumbnailer.presentation_refresh, application=True, presentation=True),
                    ),
                ])
            ]
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
        thumbnail_gruppe,
    ]
), extend=True)


bkt.powerpoint.add_context_menu(
    bkt.ribbon.ContextMenu(id_mso='ContextMenuPicture', children=[
        bkt.ribbon.Button(
            id='context-thumbnail-refresh',
            label="Thumbnail aktualisieren",
            supertip="Ausgewähltes Folien-Thumbnail aktualisieren",
            insertBeforeMso='Cut',
            image_mso='PictureChange',
            on_action=bkt.Callback(Thumbnailer.shape_refresh, shape=True, application=True),
            get_visible=bkt.Callback(Thumbnailer.is_thumbnail, shape=True),
        ),
        bkt.ribbon.Menu(
            id='context-thumbnail-settings',
            label="Thumbnail-Einstellungen",
            supertip="Einstellungen des gewählten Folien-Thumbnails ändern",
            image_mso='PictureSharpenSoftenGallery',
            insertBeforeMso='Cut',
            get_visible=bkt.Callback(Thumbnailer.is_thumbnail, shape=True),
            children=[
                bkt.ribbon.MenuSeparator(title="Qualität"),
                bkt.ribbon.CheckBox(
                    id='context-thumbnail-quality-png',
                    label="PNG (Standard)",
                    on_toggle_action=bkt.Callback(lambda pressed, shape, application: Thumbnailer.set_quality(shape, application, PASTE_DATATYPE_PNG), shape=True, application=True),
                    get_pressed=bkt.Callback(lambda shape: Thumbnailer.get_quality(shape) == PASTE_DATATYPE_PNG, shape=True),
                ),
                bkt.ribbon.CheckBox(
                    id='context-thumbnail-quality-btm',
                    label="Bitmap",
                    on_toggle_action=bkt.Callback(lambda pressed, shape, application: Thumbnailer.set_quality(shape, application, PASTE_DATATYPE_BTM), shape=True, application=True),
                    get_pressed=bkt.Callback(lambda shape: Thumbnailer.get_quality(shape) == PASTE_DATATYPE_BTM, shape=True),
                ),
                bkt.ribbon.CheckBox(
                    id='context-thumbnail-quality-emf',
                    label="Vektor (EMF)",
                    on_toggle_action=bkt.Callback(lambda pressed, shape, application: Thumbnailer.set_quality(shape, application, PASTE_DATATYPE_EMF), shape=True, application=True),
                    get_pressed=bkt.Callback(lambda shape: Thumbnailer.get_quality(shape) == PASTE_DATATYPE_EMF, shape=True),
                ),
                bkt.ribbon.MenuSeparator(title="Inhalt"),
                bkt.ribbon.CheckBox(
                    id='context-thumbnail-content-all',
                    label="Gesamte Folie",
                    on_toggle_action=bkt.Callback(lambda pressed, shape, application: Thumbnailer.set_content_only(shape, application, False), shape=True, application=True),
                    get_pressed=bkt.Callback(lambda shape: not Thumbnailer.get_content_only(shape), shape=True),
                ),
                bkt.ribbon.CheckBox(
                    id='context-thumbnail-content-only',
                    label="Nur Folieninhalt",
                    on_toggle_action=bkt.Callback(lambda pressed, shape, application: Thumbnailer.set_content_only(shape, application, True), shape=True, application=True),
                    get_pressed=bkt.Callback(lambda shape: Thumbnailer.get_content_only(shape), shape=True),
                ),
                bkt.ribbon.MenuSeparator(title="Größe"),
                bkt.ribbon.Button(
                    id='context-thumbnail-reset-aspect-ratio',
                    label="Seitenverhältnis zurücksetzen",
                    on_action=bkt.Callback(Thumbnailer.reset_aspect_ratio, shape=True),
                ),
            ]
        ),
        bkt.ribbon.Menu(
            id='context-thumbnail-reference',
            label="Folien-Referenz",
            supertip="Referenz des gewählten Folien-Thumbnails öffnen oder ändern",
            image_mso='PictureInsertFromFile',
            insertBeforeMso='Cut',
            get_visible=bkt.Callback(Thumbnailer.is_thumbnail, shape=True),
            children=[
                bkt.ribbon.Button(
                    id='context-thumbnail-gotoref',
                    label="Öffnen",
                    supertip="Referenzierte Datei öffnen und Thumbnail-Folie auswählen.",
                    on_action=bkt.Callback(Thumbnailer.goto_ref, shape=True, application=True),
                ),
                bkt.ribbon.MenuSeparator(),
                bkt.ribbon.Button(
                    id='context-thumbnail-replacefileref',
                    label="Datei ersetzen…",
                    supertip="Öffnet Datei-Auswahldialog um referenzierte Datei zu ersetzen. Die Datei muss die gleiche Folien-ID enthalten.",
                    on_action=bkt.Callback(Thumbnailer.replace_file_ref, shape=True, application=True),
                ),
                bkt.ribbon.Button(
                    id='context-thumbnail-replaceref',
                    label="Überschreiben",
                    supertip="Aktuelle Folien-Referenz ersetzen mit kopierter Folie aus Zwischenablage.",
                    on_action=bkt.Callback(Thumbnailer.replace_ref, shape=True, application=True),
                    get_enabled=bkt.Callback(Thumbnailer.enabled_slideref),
                ),
                bkt.ribbon.MenuSeparator(),
                bkt.ribbon.Button(
                    id='context-thumbnail-deleteref',
                    label="Löschen",
                    supertip="Folien-Referenz löschen und Thumbnail damit in normales Bild umwandeln.",
                    on_action=bkt.Callback(Thumbnailer.unset_thumbnail, shape=True),
                ),
            ]
        ),
        bkt.ribbon.MenuSeparator(insertBeforeMso='Cut')
    ])
)


bkt.powerpoint.add_context_menu(
    bkt.ribbon.ContextMenu(id_mso='ContextMenuThumbnail', children=[
        bkt.ribbon.Button(
            id='context-thumbnail-slide-copy',
            label="Als Folien-Thumbnail kopieren",
            supertip="Gewählte Folie als aktualisierbares Thumbnail kopieren",
            insertAfterMso='Copy',
            image_mso='Copy',
            on_action=bkt.Callback(Thumbnailer.slides_copy, presentation=True, slides=True),
            #get_visible=bkt.Callback(Thumbnailer.is_thumbnail, shape=True),
        ),
    ])
)

bkt.powerpoint.add_context_menu(
    bkt.ribbon.ContextMenu(id_mso='ContextMenuFrame', children=[
        bkt.ribbon.SplitButton(
            insertAfterMso='PasteGalleryMini',
            get_enabled=bkt.Callback(Thumbnailer.enabled_paste),
            children=[
                bkt.ribbon.Button(
                    id='context-thumbnail-slide-paste',
                    label="Als Folien-Thumbnail einfügen",
                    supertip="Als aktualisierbares Folien-Thumbnail im PNG-Format einfügen",
                    image_mso='PasteAsPicture',
                    on_action=bkt.Callback(Thumbnailer.slide_paste, application=True),
                    #get_visible=bkt.Callback(Thumbnailer.is_thumbnail, shape=True),
                    get_enabled=bkt.Callback(Thumbnailer.enabled_paste),
                ),
                bkt.ribbon.Menu(label="Als Folien-Thumbnail einfügen Menü", supertip="Format zum Einfügen des Thumbnails auswählen", children=[
                    bkt.ribbon.Button(
                        label="Als PNG einfügen (Standard)",
                        on_action=bkt.Callback(lambda application: Thumbnailer.slide_paste(application, PASTE_DATATYPE_PNG), application=True),
                    ),
                    bkt.ribbon.Button(
                        label="Als Bitmap einfügen",
                        on_action=bkt.Callback(lambda application: Thumbnailer.slide_paste(application, PASTE_DATATYPE_BTM), application=True),
                    ),
                    bkt.ribbon.Button(
                        label="Als Vektor (EMF) einfügen",
                        on_action=bkt.Callback(lambda application: Thumbnailer.slide_paste(application, PASTE_DATATYPE_EMF), application=True),
                    ),
                ])
            ]
        ),
    ])
)


class ThumbnailPopup(bkt.ui.WpfWindowAbstract):
    # _filename = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'thumbnail.xaml')
    _xamlname = 'thumbnail'
    '''
    class representing a popup-dialog for a thumbnail shape
    '''
    
    def __init__(self, context=None):
        self.IsPopup = True

        super(ThumbnailPopup, self).__init__(context)

        if context.app.activewindow.selection.shaperange.count > 1:
            self.btngoto.Visibility = Visibility.Collapsed

    def btnrefresh(self, sender, event):
        try:
            shapes = self._context.shapes
            if len(shapes) == 1:
                Thumbnailer.shape_refresh(shapes[0], self._context.app)
            else:
                Thumbnailer.shapes_refresh(shapes, self._context.app)
        except:
            bkt.message.error("Thumbnail-Aktualisierung aus unbekannten Gründen fehlgeschlagen.", "BKT: Thumbnails")
            logging.exception("Thumbnails: Error in popup!")

    def btngoto(self, sender, event):
        try:
            Thumbnailer.goto_ref(self._context.shape, self._context.app)
        except:
            bkt.message.error("Fehler beim Öffnen der Folienreferenz.", "BKT: Thumbnails")
            logging.exception("Thumbnails: Error in popup!")


# register dialog
bkt.powerpoint.context_dialogs.register_dialog(
    bkt.contextdialogs.ContextDialog(
        id=BKT_THUMBNAIL,
        module=None,
        window_class=ThumbnailPopup
    )
)