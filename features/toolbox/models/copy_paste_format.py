# -*- coding: utf-8 -*-
'''
Created on 06.02.2018

@author: rdebeerst
'''



import logging

import bkt
import bkt.library.powerpoint as pplib

from bkt.contextdialogs import DialogHelpers
# from .shapes import ShapesMore



class SlidesMore(object):
    @staticmethod
    def paste_to_slides(slides):
        for slide in slides:
            slide.Shapes.Paste()

    @staticmethod
    def paste_as_link(slide):
        try:
            slide.Shapes.PasteSpecial(Link=True)
        except:
            logging.exception("error pasting as link")
            bkt.message.error("Das Element in der Zwischenablage unterstützt diesen Einfügetyp nicht.")
    
    @classmethod
    def paste_and_replace_shapes(cls, slide, shapes, keep_size=True):
        for shape in shapes:
            try:
                cls.paste_and_replace(slide, shape, keep_size)
            except:
                logging.exception("error paste-replacing shape")
    
    @staticmethod
    def paste_and_replace(slide, shape, keep_size=True):
        pasted_shapes = slide.Shapes.Paste()
        if pasted_shapes.count > 1:
            pasted_shapes = pasted_shapes.group()
        
        #restore size
        if keep_size:
            pasted_shapes.width = shape.width
            if pasted_shapes.LockAspectRatio == 0 or pasted_shapes.height > shape.height:
                    pasted_shapes.height = shape.height
            pasted_shapes.LockAspectRatio = shape.LockAspectRatio
        
        #restore position and zorder
        pasted_shapes.top = shape.top
        pasted_shapes.left = shape.left
        pasted_shapes.rotation = shape.rotation
        pplib.set_shape_zorder(pasted_shapes, value=shape.ZOrderPosition)

        if pplib.shape_is_group_child(shape):
            #replace shape in group
            master = pplib.GroupManager(shape.ParentGroup)
            master.add_child_items(pasted_shapes)
            shape.delete()
        else:
            #replace shape
            shape.delete()
        
        pasted_shapes.select(False)

    @staticmethod
    def paste_and_distribute(slide, shapes, sort_shapes=True):
        from itertools import cycle

        def par_iterator(selected_shapes):
            for textframe in pplib.iterate_shape_textframes(selected_shapes, False):
                for idx in range(1, textframe.TextRange.Paragraphs().Count+1):
                    yield textframe.TextRange.Paragraphs(idx)
        
        pasted = slide.shapes.paste()
        par_iter = cycle(par_iterator(pasted))

        if sort_shapes:
            shapes = sorted(shapes, key=lambda s: (s.top, s.left))
        
        try:
            for textframe in pplib.iterate_shape_textframes(shapes):
                textframe.DeleteText()
                pplib.transfer_textrange(next(par_iter), textframe.textrange)
        except:
            logging.exception("error pasting texts")
        finally:
            pasted.Copy() #restore clipboard
            pasted.Delete() #remove
        
        pplib.shapes_to_range(shapes).select()

    @staticmethod
    def copy_texts(shapes):
        from bkt import dotnet
        Forms = dotnet.import_forms()

        txts = [textframe.TextRange.Text for textframe in pplib.iterate_shape_textframes(shapes) if textframe.HasText]
        if txts:
            Forms.Clipboard.SetText("\r".join(txts))

    @staticmethod
    def copy_in_highquality(slide):
        import tempfile, os
        from System import IO

        from bkt import dotnet
        Drawing = dotnet.import_drawing()
        Forms = dotnet.import_forms()

        tmpfile = os.path.join(tempfile.gettempdir(), "bkt-slidecopy.png")
        slide.export(tmpfile, "PNG", 2000)
        # logging.debug("high quality slide export at: %s"%tmpfile)

        if not os.path.exists(tmpfile):
            bkt.message.error("Folien-Export in hoher Qualität ist fehlgeschlagen!")
            return

        data = Forms.DataObject()
        png_stream = IO.MemoryStream()
        
        with Drawing.Image.FromFile(tmpfile) as img:
            #bitmap
            data.SetImage(img)
            #png
            img.Save(png_stream, Drawing.Imaging.ImageFormat.Png)
            data.SetData("PNG", False, png_stream)
            # Forms.Clipboard.SetImage(img)
            Forms.Clipboard.SetDataObject(data, True)
            img.Dispose()
        
        os.remove(tmpfile)



class FormatPainter(object):
    # @staticmethod
    # def fp_visible(context):
    #     try:
    #         return len(context.shapes) < 2
    #     except:
    #         return True

    @staticmethod
    def _get_shape_below_cursor(context):
        return DialogHelpers.last_coordinates_within_shape(context)

    @staticmethod
    def _sync_shapes(master, shapes):
        try:
            master.PickUp()
        except ValueError:
            return bkt.message.error("Funktion für ausgewähltes Shape nicht verfügar!")
        for shape in shapes:
            try:
                shape.Apply()
            except:
                logging.exception("failed to apply format")
    
    @classmethod
    def cm_sync_shapes(cls, shapes, context):
        master = cls._get_shape_below_cursor(context) or shapes[0]
        cls._sync_shapes(master, shapes)
    
    @classmethod
    def sync_shapes(cls, shapes):
        cls._sync_shapes(shapes[0], shapes)