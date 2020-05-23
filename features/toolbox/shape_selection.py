# -*- coding: utf-8 -*-
'''
Created on 06.02.2018

@author: rdebeerst
'''

from __future__ import absolute_import

import logging
from collections import OrderedDict

import bkt
import bkt.library.powerpoint as pplib

# from .shapes import ShapesMore


class ShapeSelector(object):
    key_functions = OrderedDict()
    key_functions["shape_type"] =     lambda shp: (shp.Type, shp.AutoShapeType)
    key_functions["shape_width"] =    lambda shp: shp.Width
    key_functions["shape_height"] =   lambda shp: shp.Height
    
    key_functions["pos_left"] =       lambda shp: shp.Left
    key_functions["pos_top"]  =       lambda shp: shp.Top
    key_functions["pos_rotation"] =   lambda shp: shp.Rotation
    
    key_functions["fill_type"] =      lambda shp: (shp.Fill.Visible, shp.Fill.Type)
    key_functions["fill_color"] =     lambda shp: (shp.Fill.Visible, shp.Fill.BackColor.RGB, shp.Fill.ForeColor.RGB)
    key_functions["fill_transp"] =    lambda shp: (shp.Fill.Visible, shp.Fill.Transparency)
    
    key_functions["line_weight"] =    lambda shp: (shp.Line.Visible, shp.Line.Weight)
    key_functions["line_style"] =     lambda shp: (shp.Line.Visible, shp.Line.Style, shp.Line.DashStyle)
    key_functions["line_color"] =     lambda shp: (shp.Line.Visible, shp.Line.BackColor.RGB, shp.Line.ForeColor.RGB)
    key_functions["line_begin"] =     lambda shp: (shp.Line.BeginArrowheadLength, shp.Line.BeginArrowheadStyle, shp.Line.BeginArrowheadWidth)
    key_functions["line_end"] =       lambda shp: (shp.Line.EndArrowheadLength, shp.Line.EndArrowheadStyle, shp.Line.EndArrowheadWidth)
    
    key_functions["font_name"] =      lambda shp: shp.TextFrame.TextRange.Font.Name
    key_functions["font_color"] =     lambda shp: shp.TextFrame.TextRange.Font.Color.RGB
    key_functions["font_size"] =      lambda shp: shp.TextFrame.TextRange.Font.Size
    key_functions["font_style"] =     lambda shp: (shp.TextFrame.TextRange.Font.Bold, shp.TextFrame.TextRange.Font.Underline, shp.TextFrame.TextRange.Font.Italic)
    
    key_functions["content_text"] =   lambda shp: shp.TextFrame.TextRange.Text
    key_functions["content_len"] =    lambda shp: len(shp.TextFrame.TextRange.Text)

    @staticmethod
    def _selectByKeys(master_shapes, all_shapes, keys):
        logging.debug("ShapeSelector._selectByKeys")
        cmp_funcs = [ShapeSelector.key_functions[key] for key in keys]

        master_styles= set()
        for shpMaster in master_shapes:
            master_styles.add( tuple(func(shpMaster) for func in cmp_funcs) )

        all_shapes = set(all_shapes) - set(master_shapes)
        logging.debug("ShapeSelector._selectByKeys: set ready, do select")
        for shp in all_shapes:
            try:
                # if all(func(shpMaster) == func(shp) for func in cmp_funcs):
                if tuple(func(shp) for func in cmp_funcs) in master_styles:
                    shp.Select(replace=False)
            except:
                pass
        logging.debug("ShapeSelector._selectByKeys: select done")

    @staticmethod
    def _get_all_shapes(context):
        if context.selection.HasChildShapeRange:
            return context.selection.ShapeRange[1].GroupItems
        else:
            return context.slide.Shapes

    @staticmethod
    def selectionForm(context):
        from .dialogs.shape_select import SelectWindow
        wnd = SelectWindow(ShapeSelector, context)
        wnd.show_dialog(modal=True)

        # keys = []
        # values = []
        # for k,v in ShapeSelector.key_functions.items():
        #     keys.append(k)
        #     values.append( (v[0], False) )
        
        # user_form = bkt.ui.UserInputBox("Eigenschaften für Selektion auswählen:", "Shapes selektieren")
        # clb = user_form._add_checked_listbox("comparisons", values, clb_return="CheckedIndices")
        # clb.Height = 275
        # form_return = user_form.show()
        # if len(form_return) == 0 or len(form_return["comparisons"]) == 0:
        #     return

        # ShapeSelector.selectByKeys(context, [keys[sel] for sel in form_return["comparisons"]])

    @classmethod
    def selectShapes(cls, context, shapes):
        context.selection.Unselect()
        for shp in shapes:
            shp.Select(replace=False)

    @classmethod
    def selectByKeys(cls, context, keys, master_shapes=None, unselect=False):
        logging.debug("ShapeSelector.selectByKeys")
        master_shapes = master_shapes or context.shapes
        if unselect:
            cls.selectShapes(context, master_shapes)
            logging.debug("ShapeSelector.selectByKeys: unselect done")
        
        cls._selectByKeys(master_shapes, cls._get_all_shapes(context), keys)

    
    @classmethod
    def invert_selection(cls, context):
        selection = context.selection
        if selection.Type == 2 or selection.Type == 3:
            # shapes or text selected
            if selection.HasChildShapeRange:
                selected_shapes = list(iter(selection.childshaperange))
            else:
                selected_shapes = list(iter(selection.shaperange))
        else:
            # slide selected
            selected_shapes = []
        
        all_shapes = cls._get_all_shapes(context)
        
        new_shape_selection = [shape for shape in all_shapes if not shape in selected_shapes]
        if len(new_shape_selection) == 0:
            selection.Unselect()
        else:
            pplib.shapes_to_range(new_shape_selection).Select()
            # new_shape_selection[0].Select(replace=True)
            # for shape in new_shape_selection:
            #     shape.Select(replace=False)

    @classmethod
    def _is_within(cls, outer, inner):
        return (outer.visual_x < inner.visual_x and outer.visual_y < inner.visual_y and
                outer.visual_x+outer.visual_width > inner.visual_x+inner.visual_width and outer.visual_y+outer.visual_height > inner.visual_y+inner.visual_height)
        # return (outer.Left < inner.Left and outer.Top < inner.Top and
        #         outer.Left+outer.Width > inner.Left+inner.Width and outer.Top+outer.Height > inner.Top+inner.Height)
    
    @classmethod
    def _is_ontop(cls, lower, upper):
         return (lower.ZOrderPosition < upper.ZOrderPosition)

    @classmethod
    def _has_overlap(cls, shp1, shp2):
        return (shp1.visual_x < shp2.visual_x+shp2.visual_width  and shp1.visual_x+shp1.visual_width > shp2.visual_x and
                shp1.visual_y < shp2.visual_y+shp2.visual_height and shp1.visual_y+shp1.visual_height > shp2.visual_y)
        # return (shp1.Left < shp2.Left+shp2.Width and shp1.Left+shp1.Width > shp2.Left and
        #         shp1.Top < shp2.Top+shp2.Height and shp1.Top+shp1.Height > shp2.Top)
    
    @classmethod
    def select_overlapping(cls, context):
        all_shapes = pplib.wrap_shapes(cls._get_all_shapes(context))
        for shpMaster in pplib.wrap_shapes(context.shapes):
            for shp in all_shapes:
                if cls._has_overlap(shpMaster, shp):
                    shp.Select(replace=False)
    
    @classmethod
    def select_within(cls, context):
        all_shapes = pplib.wrap_shapes(cls._get_all_shapes(context))
        for shpMaster in pplib.wrap_shapes(context.shapes):
            for shp in all_shapes:
                if cls._is_within(shpMaster, shp):
                    shp.Select(replace=False)
    
    @classmethod
    def select_containing(cls, context):
        all_shapes = pplib.wrap_shapes(cls._get_all_shapes(context))
        for shpMaster in pplib.wrap_shapes(context.shapes):
            for shp in all_shapes:
                if cls._is_ontop(shpMaster, shp) and cls._is_within(shpMaster, shp):
                    shp.Select(replace=False)
    
    @classmethod
    def select_behind(cls, context):
        all_shapes = pplib.wrap_shapes(cls._get_all_shapes(context))
        for shpMaster in pplib.wrap_shapes(context.shapes):
            for shp in all_shapes:
                if not cls._is_ontop(shpMaster, shp) and cls._is_within(shpMaster, shp):
                    shp.Select(replace=False)


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
            bkt.message.error("Das Element in der Zwischenablage unterstützt diesen Einfügetyp nicht.")
    
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
        
        pasted_shapes.select()

    @staticmethod
    def copy_in_highquality(slide):
        import tempfile, os
        from System import IO

        from bkt import dotnet
        Drawing = dotnet.import_drawing()
        Forms = dotnet.import_forms()

        tmpfile = os.path.join(tempfile.gettempdir(), "bkt-slidecopy.png")
        slide.export(tmpfile, "PNG", 2000)
        logging.debug("high quality slide export at: %s", tmpfile)

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



selection_menu = bkt.ribbon.Menu(
    label='Auswahl',
    screentip='Auswahl von Shapes',
    supertip='Auswahl von Shapes, die dem aktuellem Shape bzgl. Typ/Hintergrund/Rahmen ähneln',
    show_label=False,
    image_mso='ObjectsMultiSelect',
    children = [
        bkt.ribbon.MenuSeparator(title="Auswahl von Shapes…"),
        bkt.ribbon.Button(
            id = 'shapes_form',
            image_mso = 'GroupSmartArtQuickStyles',
            label='…mit gleicher Form',
            #show_label=False,
            on_action=bkt.Callback(lambda context: ShapeSelector.selectByKeys(context, ['shape_type']), context=True),
            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
            screentip="Shape-Objekte mit gleicher Form markieren",
            supertip="Selektiere alle Shapes auf dem aktuellen Slide, die die gleiche Form haben wie eine der selektierten Shapes",
        ),

        bkt.ribbon.Button(
            id = 'shapes_bg',
            image_mso = 'AppointmentColor1',
            label='…mit gleichem Hintergrund',
            #show_label=False,
            on_action=bkt.Callback(lambda context: ShapeSelector.selectByKeys(context, ['fill_type', 'fill_color']), context=True),
            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
            screentip="Shape-Objekte mit gleichem Hintergrund markieren",
            supertip="Selektiere alle Shapes auf dem aktuellen Slide, die den gleichen Hintergrund (Farbe) haben wie eine der selektierten Shapes",
        ),

        bkt.ribbon.Button(
            id = 'shapes_border',
            image_mso = 'BlackAndWhiteBlackWithWhiteFill',
            label='…mit gleichem Rahmen',
            #show_label=False,
            on_action=bkt.Callback(lambda context: ShapeSelector.selectByKeys(context, ['line_style', 'line_color']), context=True),
            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
            screentip="Shape-Objekte mit gleichem Rahmen markieren",
            supertip="Selektiere alle Shapes auf dem aktuellen Slide, die den gleichen Rahmen (Farbe, Strichtyp) haben wie eine der selektierten Shapes",
        ),

        bkt.ribbon.Button(
            id = 'shapes_font',
            image_mso = 'FontColorPicker',
            label='…mit gleicher Schriftfarbe',
            #show_label=False,
            on_action=bkt.Callback(lambda context: ShapeSelector.selectByKeys(context, ['font_color']), context=True),
            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
            screentip="Shape-Objekte mit gleicher Schritfarbe markieren",
            supertip="Selektiere alle Shapes auf dem aktuellen Slide, die die gleiche Schriftfarbe haben wie eine der selektierten Shapes",
        ),

        bkt.ribbon.Button(
            id = 'shapes_text',
            image_mso = 'TextBoxInsert',
            label='…mit gleichem Text',
            #show_label=False,
            on_action=bkt.Callback(lambda context: ShapeSelector.selectByKeys(context, ['content_text']), context=True),
            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
            screentip="Shape-Objekte mit gleichem Text markieren",
            supertip="Selektiere alle Shapes auf dem aktuellen Slide, die den gleichen Text haben wie eine der selektierten Shapes",
        ),

        bkt.ribbon.Button(
            id = 'shapes_size',
            image_mso = 'ShowEmptyContainers',
            label='…mit gleicher Größe',
            #show_label=False,
            on_action=bkt.Callback(lambda context: ShapeSelector.selectByKeys(context, ['shape_width', 'shape_height']), context=True),
            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
            screentip="Shape-Objekte mit gleicher Größe markieren",
            supertip="Selektiere alle Shapes auf dem aktuellen Slide, die die gleiche Größe haben wie eine der selektierten Shapes",
        ),
        bkt.ribbon.MenuSeparator(),
        bkt.ribbon.Button(
            id = 'shapes_select_custom',
            image_mso = 'ShowCustomPropertiesPage',
            label='Benutzerdefinierte Auswahl…',
            #show_label=False,
            on_action=bkt.Callback(ShapeSelector.selectionForm, context=True),
            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
            screentip="Shape-Objekte nach Benutzerdefinierter Auswahl markieren",
            supertip="Öffne einen Dialog zur Auswahl der Shape-Eigenschaften, nach welcher die Shapes auf der aktuellen Folie markiert werden sollen.",
        ),

        bkt.ribbon.MenuSeparator(),
        bkt.ribbon.Button(
            id = 'shapes_select_overlapping',
            image_mso = 'SlideShowResolutionGallery',
            label='Überlappend',
            #show_label=False,
            on_action=bkt.Callback(ShapeSelector.select_overlapping, context=True),
            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
            screentip="Shape-Objekte überlappend mit gewählten Shapes markieren",
            supertip="Selektiere alle Shapes auf dem aktuellen Slide, die sich mit einem der selektierten Shapes überlappen.",
        ),
        bkt.ribbon.Button(
            id = 'shapes_select_within',
            image_mso = 'SlideShowResolutionGallery',
            label='Innerhalb',
            #show_label=False,
            on_action=bkt.Callback(ShapeSelector.select_within, context=True),
            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
            screentip="Shape-Objekte innerhalb der gewählten Shapes markieren",
            supertip="Selektiere alle Shapes auf dem aktuellen Slide, die sich vollständig innerhalb eines der selektierten Shapes befinden.",
        ),
        bkt.ribbon.Button(
            id = 'shapes_select_containing',
            image_mso = 'SlideShowResolutionGallery',
            label='Inner- & oberhalb',
            #show_label=False,
            on_action=bkt.Callback(ShapeSelector.select_containing, context=True),
            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
            screentip="Shape-Objekte innerhalb und oberhalb der gewählten Shapes markieren",
            supertip="Selektiere alle Shapes auf dem aktuellen Slide, die sich vollständig innerhalb und oberhalb (d.h. Z-Order ist größer) eines der selektierten Shapes befinden.",
        ),
        bkt.ribbon.Button(
            id = 'shapes_select_behind',
            image_mso = 'SlideShowResolutionGallery',
            label='Inner- & unterhalb',
            #show_label=False,
            on_action=bkt.Callback(ShapeSelector.select_behind, context=True),
            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
            screentip="Shape-Objekte innerhalb und unterhalb der gewählten Shapes markieren",
            supertip="Selektiere alle Shapes auf dem aktuellen Slide, die sich vollständig innerhalb und unterhalb (d.h. Z-Order ist kleiner) eines der selektierten Shapes befinden.",
        ),

        bkt.ribbon.MenuSeparator(title="Markieren"),
        bkt.mso.control.SelectionPane,
        bkt.ribbon.Button(
            id = 'shapes_select_invert',
            image_mso = 'ObjectsMultiSelect',
            label='Auswahl invertieren',
            on_action=bkt.Callback(ShapeSelector.invert_selection, context=True),
            # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            supertip="Invertiert die aktuelle Auswahl. Es werden alle Shapes (auch Platzhalter) markiert, die vorher nicht markiert waren.",
        ),
    ]
)



clipboard_group = bkt.ribbon.Group(
    id="bkt_clipboard_group",
    label='Ablage',
    image_mso='ObjectsMultiSelect',
    children=[
        bkt.ribbon.SplitButton(
            show_label=False,
            get_enabled=bkt.Callback(lambda context: context.app.commandbars.GetEnabledMso("Paste"), context=True),
            children=[
                bkt.mso.button.PasteSpecialDialog,
                bkt.ribbon.Menu(
                    label="Einfügen-Menü",
                    supertip="Menü mit verschiedenen Einfüge-Operationen",
                    children=[
                        bkt.mso.button.PasteSpecialDialog,
                        bkt.ribbon.MenuSeparator(),
                        bkt.ribbon.Button(
                            id='paste_to_slides',
                            label="Auf ausgewählte Folien einfügen",
                            supertip="Zwischenablage auf allen ausgewählten Folien gleichzeitig einfügen.",
                            image_mso='PasteDuplicate',
                            on_action=bkt.Callback(SlidesMore.paste_to_slides, slides=True),
                        ),
                        bkt.ribbon.Button(
                            id='paste_as_link',
                            label="Als Verknüpfung einfügen",
                            supertip="Zwischenablage als verknüpftes Element (bspw. Bild, OLE-Objekt) einfügen.",
                            image_mso='PasteLink',
                            on_action=bkt.Callback(SlidesMore.paste_as_link, slide=True),
                        ),
                        bkt.ribbon.Button(
                            id='paste_and_replace',
                            label="Mit Zwischenablage ersetzen",
                            supertip="Markiertes Shape mit dem Inhalt der Zwischenablage ersetzen und dabei Größe und Position erhalten.",
                            image_mso='PasteSingleCellExcelTableDestinationFormatting',
                            on_action=bkt.Callback(SlidesMore.paste_and_replace, slide=True, shape=True),
                            get_enabled=bkt.apps.ppt_shapes_exactly1_selected,
                        ),
                        bkt.ribbon.MenuSeparator(),
                        bkt.mso.button.ShowClipboard,
                    ]
                )
            ]
        ),
        bkt.ribbon.SplitButton(
            show_label=False,
            get_enabled=bkt.Callback(lambda context: context.app.commandbars.GetEnabledMso("Copy"), context=True),
            children=[
                bkt.mso.button.Copy,
                bkt.ribbon.Menu(
                    label="Kopieren-Menü",
                    supertip="Menü mit verschiedenen Kopier-Operationen",
                    children=[
                        bkt.mso.button.Copy,
                        bkt.mso.button.PasteDuplicate,
                        bkt.ribbon.Button(
                            id="copy_slide_hq",
                            label="Folie als HQ-Bild kopieren",
                            supertip="Kopiert die aktuelle Folie in hoher Qualität in die Zwischenablage.",
                            image_mso='CopyPicture',
                            on_action=bkt.Callback(SlidesMore.copy_in_highquality, slide=True),
                            get_enabled=bkt.get_enabled_auto
                        ),
                    ]
                )
            ]
        ),
        #bkt.mso.control.PasteSpecialDialog,
        #bkt.mso.control.Cut,
        # bkt.mso.control.CopySplitButton,
        
        selection_menu,
        
        bkt.mso.control.PasteApplyStyle,
        bkt.mso.control.PickUpStyle,
        bkt.mso.control.FormatPainter
    ]
)


