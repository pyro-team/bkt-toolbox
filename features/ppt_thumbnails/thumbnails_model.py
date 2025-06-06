# -*- coding: utf-8 -*-
'''
Created on 2023-01-16
@author: Florian Stallmann
'''

import os #required for relative paths and tempfile removal
import tempfile
import logging

from contextlib import contextmanager
# from System import Array #required to create ShapeRanges
# from System.Windows import Visibility

import bkt
import bkt.library.powerpoint as pplib

import bkt.dotnet as dotnet
Forms = dotnet.import_forms() #required to read clipboard and open file dialogs


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

    def set_thumbnail(self, slide_id, slide_path, data_type=PASTE_DATATYPE_PNG, content_only=False, shape_id=None, **kwargs):
        self.data["slide_id"] = slide_id
        self.data["slide_path"] = slide_path
        self.data["data_type"] = data_type
        self.data["content_only"] = content_only #=exclude placeholder shapes
        if shape_id is not None:
            self.data["shape_id"] = shape_id
        
        # for future compatibility
        self.data.update(kwargs)


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
            if not path.startswith("https://") and not os.path.isabs(path):
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
    def find_and_export_object(cls, application, slide_id, slide_path, content_only=False, shape_id=None, data_type=None, **kwargs):
        #kwargs added for for future compatibility
        #avoid referenced before assignment error in finally clause
        close = None
        tmpfile = None
        try:
            try:
                pres, close = cls._get_presentation(application, slide_path)
            except EnvironmentError:
                logging.exception("presentation not found")
                raise IOError("presentation not found")
        
            try:
                slide = pres.Slides.FindBySlideId(slide_id)
            except SystemError:
                logging.exception("slide id not found")
                raise IndexError("slide id not found")

            filetype = DATATYPE_MAPPING.get(data_type, PASTE_DATATYPE_PNG)
            tmpfile = os.path.join(tempfile.gettempdir(), "bkt-thumbnail-tempfile."+filetype[1])

            try:
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
        #reapply ratio (only required if LockAspectRatio=0)
        ratio = shape.Width/shape.Height
        shape.Height = height
        shape.Width = ratio*height

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
            if isinstance(slide_id, tuple):
                shape_id = slide_id[1]
                slide_id = slide_id[0]
            else:
                shape_id = None

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

            try:
                #Paste
                # shape = cur_slide.Shapes.PasteSpecial(Datatype=data_type)
                # pasted_shapes += 1
                #Save tags
                # shape = application.ActiveWindow.Selection.ShapeRange(1)
                with ThumbnailerTags(shape.Tags) as tags:
                    tags.set_thumbnail(slide_id, data["slide_path"], data_type, content_only, shape_id)
                shape.Tags.Add(bkt.contextdialogs.BKT_CONTEXTDIALOG_TAGKEY, BKT_THUMBNAIL)

                # add hyperlink
                cls._update_hyperlink(shape, application)
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
    def slide_paste_png(cls, application):
        cls.slide_paste(application, PASTE_DATATYPE_PNG)
    @classmethod
    def slide_paste_btm(cls, application):
        cls.slide_paste(application, PASTE_DATATYPE_BTM)
    @classmethod
    def slide_paste_emf(cls, application):
        cls.slide_paste(application, PASTE_DATATYPE_EMF)

    @classmethod
    def replace_ref(cls, shape, application):
        data = cls.get_clipboard_data(application)
        with ThumbnailerTags(shape.Tags) as tags:
            tags["slide_path"] = data["slide_path"]

            if isinstance(data["slide_ids"][0], tuple): #tuple of (slide_id, shape_id)
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

        logging.debug("New path: %s", path)

        with ThumbnailerTags(shape.Tags) as tags:
            tags["slide_path"] = path

        cls.shape_refresh(shape, application)

    @classmethod
    def goto_ref(cls, shape, application):
        with ThumbnailerTags(shape.Tags) as tags:
            slide_id = tags["slide_id"]
            slide_path = tags["slide_path"]
            shape_id = tags.get("shape_id")

        try:
            pres, _ = cls._get_presentation(application, slide_path, False)
            
            #bring window to front
            if pres.Windows.Count > 0:
                pres.Windows[1].Activate()
            else:
                pres.NewWindow()
        except EnvironmentError:
            logging.exception("Thumbnails: Error finding presentation")
            if bkt.message.confirmation("Fehler! Referenzierte Präsentation '%s' nicht gefunden. Neue Datei auswählen?" % slide_path, "BKT: Thumbnails", icon=bkt.MessageBox.WARNING):
                cls.replace_file_ref(shape, application)
            return

        try:
            slide = pres.Slides.FindBySlideId(slide_id)
            slide.Select()
        except SystemError:
            bkt.message.error("Fehler! Folie in der referenzierten Präsentation nicht gefunden.", "BKT: Thumbnails")
            return

        if shape_id is not None:
            try:
                shp = cls._find_by_shape_id(slide, shape_id)
                shp.Select()
            except IndexError:
                bkt.message.error("Fehler! Shape in der referenzierten Präsentation nicht gefunden.", "BKT: Thumbnails")
    
    @classmethod
    def presentation_unset(cls, presentation):
        if bkt.message.confirmation("Dies löscht dauerhaft die Folien-Referenz und damit die Möglichkeit der Aktualisierung aller Thumbnails in der Präsentation.", "BKT: Thumbnails"):
            total = 0
            for sld in presentation.slides:
                for shp in pplib.iterate_shape_subshapes( sld.shapes ):
                    try:
                        if cls.is_thumbnail(shp):
                            shp.Tags.Delete(BKT_THUMBNAIL)
                            shp.Tags.Delete(bkt.contextdialogs.BKT_CONTEXTDIALOG_TAGKEY)
                            total += 1
                    except:
                        logging.exception("Thumbnails: Could not determine if shape is thumbnail")

    @classmethod
    def presentation_refresh(cls, application, presentation):
        cls.slide_refresh(application, presentation.slides)

    @classmethod
    def slide_refresh(cls, application, slides):
        total = 0
        err_counter = 0
        for sld in slides:
            #NOTE: first collect all shapes within a slide as refresh will add and remove shapes which crashes the iteration
            thumbs_on_slide = []
            for shp in pplib.iterate_shape_subshapes( sld.shapes ):
                try:
                    if cls.is_thumbnail(shp):
                        thumbs_on_slide.append(shp)
                        total += 1
                except:
                    logging.exception("Thumbnails: Could not determine if shape is thumbnail")
            for shp in thumbs_on_slide:
                try:
                    cls._shape_refresh(shp, application) #FIXME: currently file is opened for each thumbnail, can be improved for better performance
                except:
                    logging.exception("Thumbnails: Failed to update slide")
                    cls._mark_erroneous_shape(shp)
                    err_counter += 1

        if total == 0:
            bkt.message.warning("Keine Folien-Thumbnails gefunden.", "BKT: Thumbnails")
        elif err_counter > 0:
            bkt.message.warning("Es wurde/n %r Folien-Thumbnail/s aktualisiert, aber %r Folien-Thumbnail/s konnten wegen eines Fehlers nicht aktualisiert werden. Die fehlerhaften Thumbnails wurden mit dem Text 'BKT THUMB UPDATE FAILED' markiert." % (total-err_counter, err_counter), "BKT: Thumbnails")
        else:
            bkt.message("Es wurde/n %r Folien-Thumbnail/s aktualisiert." % total, "BKT: Thumbnails")


    @classmethod
    def shapes_refresh(cls, shapes, application):
        err_counter = 0
        new_shapes = []
        for shp in shapes:
            try:
                new_shapes.append( cls._shape_refresh(shp, application) ) #FIXME: currently file is opened for each thumbnail, can be improved for better performance
            except:
                cls._mark_erroneous_shape(shp)
                err_counter += 1
                # bkt.helpers.exception_as_message()
        pplib.shapes_to_range(new_shapes).select()

        if err_counter > 0:
            bkt.message.warning("Es wurde/n %r Folien-Thumbnail/s aktualisiert, aber %r Folien-Thumbnail/s konnten wegen eines Fehlers nicht aktualisiert werden. Die fehlerhaften Thumbnails wurden mit dem Text 'BKT THUMB UPDATE FAILED' markiert." % (len(shapes)-err_counter, err_counter), "BKT: Thumbnails")
        # else:
        #     bkt.message("Es wurde/n %r Folien-Thumbnail/s aktualisiert." % len(shapes), "BKT: Thumbnails")

    @classmethod
    def shape_refresh(cls, shape, application):
        try:
            shp = cls._shape_refresh(shape, application)
            shp.select()
            return shp
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

        cls._update_hyperlink(new_shp, application)

        cls.remain_position_and_zorder(shape, new_shp)
        shape.PickUp()
        new_shp.Apply()

        shape_name = shape.Name

        #handle thumbnail in group (part 2)
        if group:
            # group.select()
            # shape.Delete()
            # new_shp.Select(False)
            # group.regroup(application.ActiveWindow.Selection.ShapeRange)
            # new_shp.Select()
            group.add_child_items([new_shp])
            group.regroup()
            shape.Delete()
        else:
            shape.Delete()
            new_shp.Name = shape_name
            # new_shp.Select()
        #NOTE: selecting here is not a good idea as view might not be active (e.g. refresh whole presentation)

        return new_shp
    
    @classmethod
    def _update_hyperlink(cls, shape, application):
        with ThumbnailerTags(shape.Tags) as tags:
            slide_id = tags["slide_id"]
            slide_path = tags["slide_path"]

        if slide_path == "CURRENT" or slide_path == application.ActivePresentation.FullName:
            try:
                slide = application.ActivePresentation.Slides.FindBySlideId(slide_id)
                shape.ActionSettings(1).Hyperlink.SubAddress = "{},{},{}".format(slide.SlideId,slide.SlideIndex,slide.Name)
            except SystemError:
                logging.warning("Thumbnails: Update of hyperlink failed!")

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
                data = (list(data[0]), str(data[1]))
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
        #NOTE: if presentation is stored in OneDriver a url is returned, refer to https://stackoverflow.com/questions/33734706/excels-fullname-property-with-onedrive

        if path == application.ActivePresentation.FullName:
            return "CURRENT"
        elif path.startswith("https://"): #OneDrive or SharePoint
            return path
        
        drive1, _ = os.path.splitdrive(path)
        drive2, _ = os.path.splitdrive(application.ActivePresentation.FullName)
        if USE_RELATIVE_PATHS and drive1 != '' and drive1 == drive2: #same drive -> use relative path
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
            shape.Tags.Delete(bkt.contextdialogs.BKT_CONTEXTDIALOG_TAGKEY)

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

    @classmethod
    def toggle_content_only(cls, shape, application):
        cls.set_content_only(shape, application, not cls.get_content_only(shape))


context_settings = lambda: bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None, 
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
                ])


context_reference = lambda: bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None, 
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
                ])