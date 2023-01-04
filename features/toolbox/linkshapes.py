# -*- coding: utf-8 -*-
'''
Created on 06.09.2018

@author: fstallmann
'''



import uuid

import bkt
import bkt.library.powerpoint as pplib


BKT_LINK_UUID = "BKT_LINK_UUID"

class LinkedShapes(object):
    current_link_guid = None
    master = "current"
    status_overlay = False
    status_colors = dict(green=0 + 192 * 256 + 0 * 256**2, yellow=192 + 192 * 256 + 0 * 256**2, red=192 + 0 * 256 + 0 * 256**2)

    ### Helpers ###

    @staticmethod
    def _add_tags(shape, link_guid):
        shape.Tags.Add(BKT_LINK_UUID, link_guid)
        shape.Tags.Add(bkt.contextdialogs.BKT_CONTEXTDIALOG_TAGKEY, BKT_LINK_UUID)

    @classmethod
    def get_master_shape(cls, shape, context):
        if cls.master == "first":
            return cls.get_linked_shape(shape, context, 0, False)
        elif cls.master == "last":
            return cls.get_linked_shape(shape, context, -1, False)
        else: #current
            return shape

    @classmethod
    def _iterate_linked_shapes(cls, shape, context):
        link_guid = shape.Tags.Item(BKT_LINK_UUID)
        
        if link_guid == "":
            raise IndexError("Shape has no link uuid")

        # active_slide_index = shape.Parent.SlideIndex
        for slide in context.app.ActivePresentation.Slides:
            # if slide.SlideIndex == active_slide_index:
            #     continue
            # for cShp in list(iter(slide.Shapes)): #list(iter()) required for deletion of shapes
            for i in range(slide.Shapes.Count, 0, -1): #count backwords to support deletion
                cShp = slide.Shapes[i]
                if cShp == shape:
                    continue
                if cShp.Tags.Item(BKT_LINK_UUID) == link_guid:
                    # apply_func(cShp)
                    yield cShp
    
    @classmethod
    def _create_status_overlay(cls, shape, context, color="green", text="updated successfully"):
        if not cls.status_overlay:
            return
        overlay = context.slide.shapes.addshape(1, shape.left-1, shape.top-1, shape.width+2, shape.height+2)
        overlay.rotation = shape.rotation
        overlay.line.visible = 0
        overlay.fill.forecolor.rgb = cls.status_colors.get(color, 0)
        overlay.fill.transparency = 0.5
        txt = overlay.textframe
        txt.textrange.text = text
        txt.textrange.font.color = 0
        txt.textrange.font.size = 8
        txt.textrange.ParagraphFormat.Bullet.Visible = False
        txt.textrange.ParagraphFormat.Alignment = 2 #ppAlignCenter
        # Autosize / Text nicht umbrechen
        txt.WordWrap = 0
        txt.AutoSize = 0
        txt.MarginBottom = 0
        txt.MarginTop = 0
        txt.MarginRight = 0
        txt.MarginLeft = 0
        return overlay


    ### Enabled/Visible callbacks ###

    @staticmethod
    def is_linked_shape(shape):
        return pplib.TagHelper.has_tag(shape, BKT_LINK_UUID)

    @classmethod
    def not_is_linked_shape(cls, shape):
        return not cls.is_linked_shape(shape)

    @classmethod
    def are_linked_shapes(cls, shapes):
        return all(cls.is_linked_shape(shape) for shape in shapes)

    @classmethod
    def enabled_add_linked_shapes(cls):
        return cls.current_link_guid != None

    ### Linked shape creation/search ###

    @classmethod
    def find_similar_and_link(cls, shape, context):
        #open wpf window
        from .dialogs.linkshapes_find import FindWindow
        FindWindow.create_and_show_dialog(cls, context, shape)

    @classmethod
    def find_similar_shapes_and_link(cls, shape, context, attributes, threshold=0, limit_slides=None, dry_run=False):
        from bkt.library.algorithms import is_close

        shape = pplib.wrap_shape(shape)
        if shape.Tags.Item(BKT_LINK_UUID) != '':
            link_guid = shape.Tags.Item(BKT_LINK_UUID)
        else:
            link_guid = str(uuid.uuid4())

        active_slide_index = shape.Parent.SlideIndex
        shapes_found = 0
        # comparer_values = lambda val1, val2: abs(round(val1,3)-round(val2,3))/round(val1,3)<=threshold
        comparer_values = lambda val1, val2: val1==val2 if isinstance(val1, str) else is_close(val1, val2, threshold)
        
        all_slides = context.app.ActivePresentation.Slides
        num_slides = limit_slides or all_slides.Count

        # type should not be compared using threshold
        try:
            attributes.remove("type")
            compare_type = True
        except ValueError:
            compare_type = False

        for slide in context.app.ActivePresentation.Slides:
            if slide.SlideIndex <= active_slide_index:
                continue
            for sld_shape in slide.Shapes:
                cShp = pplib.wrap_shape(sld_shape)
                if (
                    (not compare_type or cShp.Type == shape.Type and cShp.AutoShapeType == shape.AutoShapeType) and
                    all(comparer_values(getattr(shape, attr), getattr(cShp, attr)) for attr in attributes)
                    # cShp.Left == shape.Left and cShp.Top == shape.Top and
                    # cShp.Width == shape.Width and cShp.Height == shape.Height and
                    # cShp.Rotation == shape.Rotation 
                    ):
                    if not dry_run:
                        cls._add_tags(cShp, link_guid)
                    shapes_found += 1
            num_slides -= 1
            if num_slides == 0:
                break

        if shapes_found == 0:
            bkt.message("Keine vergleichbaren Shapes gefunden.", "BKT: Verknüpfte Shapes")
        elif dry_run:
            bkt.message("Es wurden %s Shapes zum verknüpfen gefunden." % shapes_found, "BKT: Verknüpfte Shapes")
        else:
            cls._add_tags(shape, link_guid)
            bkt.message("Das Shape wurde mit %s Shapes verknüpft." % shapes_found, "BKT: Verknüpfte Shapes")
            context.ribbon.ActivateTab('bkt_context_tab_linkshapes')

    @classmethod
    def copy_to_all(cls, shape, context): #shape as parameter for get_enabled required
        #open wpf window
        from .dialogs.linkshapes_copy import CopyWindow
        wnd = CopyWindow(cls, context, shape)
        wnd.show_dialog(modal=False)
        # CopyWindow.create_and_show_dialog(cls, context)

    @classmethod
    def copy_shapes_to_slides(cls, shapes, context, limit_slides=None):
        for shape in shapes:
            cls.copy_shape_to_slides(shape, context, limit_slides)
        
        try:
            #activate view (note: parent.select does not work)
            context.app.ActiveWindow.View.GotoSlide(shapes[0].Parent.SlideIndex)
            shapes[0].select()
            context.ribbon.ActivateTab('bkt_context_tab_linkshapes')
        except:
            bkt.helpers.exception_as_message()

    @classmethod
    def copy_shape_to_slides(cls, shape, context, limit_slides=None):
        link_guid = str(uuid.uuid4())

        cls._add_tags(shape, link_guid)
        shape.Copy()
        active_slide_index = shape.Parent.SlideIndex

        all_slides = context.app.ActivePresentation.Slides
        num_slides=limit_slides or all_slides.Count
        for slide in all_slides:
            if slide.SlideIndex <= active_slide_index:
                continue
            slide.Shapes.Paste()
            num_slides -= 1
            if num_slides == 0:
                break

    @classmethod
    def link_shapes(cls, shapes):
        cls.current_link_guid = str(uuid.uuid4())
        cls.add_to_link_shapes(shapes)
    
    @classmethod
    def each_link_shapes(cls, shapes):
        for shape in shapes:
            link_guid = str(uuid.uuid4())
            cls._add_tags(shape, link_guid)

    ### Remove linked shapes ###

    @classmethod
    def unlink_shape(cls, shape):
        shape.Tags.Delete(BKT_LINK_UUID)
        shape.Tags.Delete(bkt.contextdialogs.BKT_CONTEXTDIALOG_TAGKEY)

    @classmethod
    def unlink_shapes(cls, shapes):
        for shape in shapes:
            cls.unlink_shape(shape)
    
    @classmethod
    def unlink_all_shapes(cls, shapes, context):
        for shape in shapes:
            cls.unlink_shapes(cls._iterate_linked_shapes(shape, context))
            cls.unlink_shape(shape)

    ### Extend linked shapes ###

    @classmethod
    def extend_link_shapes(cls, shape):
        cls.current_link_guid = shape.Tags.Item(BKT_LINK_UUID)
        bkt.message('Die Link-ID wurde zwischengespeichert. Als nächstes können weitere Shapes ausgewählt und über "Ausgewählte Shapes zur Verknüpfung hinzufügen" mit diesem Shape verknüpft werden.', "BKT: Verknüpfte Shapes")

    @classmethod
    def add_to_link_shapes(cls, shapes):
        for shape in shapes:
            cls._add_tags(shape, cls.current_link_guid)

    ### Statistics & Jump ###

    @classmethod
    def count_link_shapes(cls, shape, context):
        count_shapes = sum(1 for _ in cls._iterate_linked_shapes(shape, context))
        bkt.message("Es wurden %s verknüpfte Shapes gefunden." % count_shapes, "BKT: Verknüpfte Shapes")
    
    @classmethod
    def select_link_shapes_slides(cls, shape, context):
        # from System import Array
        # slides = set(s.Parent.SlideIndex for s in cls._iterate_linked_shapes(shape, context))
        # slides.add(shape.Parent.SlideIndex)
        # bkt.message("%s" % slides)
        # context.presentation.Slides.Range(Array[int](slides)).Select()
        slides = set(s.Parent.SlideNumber for s in cls._iterate_linked_shapes(shape, context))
        slides.add(shape.Parent.SlideNumber)
        bkt.message('Folgende Folie enthalten verknüpfte Shapes: %s' % ", ".join(str(i) for i in sorted(slides)), "BKT: Verknüpfte Shapes")

    @classmethod
    def get_linked_shape(cls, shape, context, goto=1, delta=True): #goto=1 -> next linked shape, delta=False -> goto interpreted as absolute index
        import logging
        link_guid = shape.Tags.Item(BKT_LINK_UUID)
        all_slides = context.presentation.Slides

        #fast track for getting first and last shape
        if not delta and goto in (0,-1):
            logging.info("going the fast track")
            if goto == 0:
                s_range = range(1, all_slides.Count+1)
            else:
                s_range = range(all_slides.Count, 0, -1)
            for i in s_range:
                logging.info("fast track on slide %s", i)
                for cShp in all_slides[i].Shapes:
                    if pplib.TagHelper.has_tag(cShp, BKT_LINK_UUID, link_guid):
                        logging.info("fast track successful")
                        return cShp

        cur_shape_index = None
        list_shapes = []
        for slide in all_slides:
            for cShp in slide.Shapes:
                if pplib.TagHelper.has_tag(cShp, BKT_LINK_UUID, link_guid):
                    list_shapes.append(cShp)
                if cShp == shape:
                    cur_shape_index = len(list_shapes)-1
        
        sel_shape_index = goto if not delta else (cur_shape_index + goto) % len(list_shapes)
        return list_shapes[sel_shape_index]
    
    @classmethod
    def goto_linked_shape(cls, shape, context, goto=1, delta=True): #goto=1 -> next linked shape, delta=False -> goto interpreted as absolute index
        goto_shape = cls.get_linked_shape(shape, context, goto, delta)
        #activate view
        context.app.ActiveWindow.View.GotoSlide(goto_shape.Parent.SlideIndex)
        # goto_shape.parent.select() #select slide
        goto_shape.select() #select shape

    @classmethod
    def goto_first_shape(cls, shape, context):
        cls.goto_linked_shape(shape, context, 0, False)

    @classmethod
    def goto_last_shape(cls, shape, context):
        cls.goto_linked_shape(shape, context, -1, False)

    @classmethod
    def goto_next_shape(cls, shape, context):
        cls.goto_linked_shape(shape, context, 1)

    @classmethod
    def goto_previous_shape(cls, shape, context):
        cls.goto_linked_shape(shape, context, -1)

    ### Alignments/Sync ###

    @classmethod
    def size_linked_shapes(cls, shapes, context):
        for shape in shapes:
            try:
                cls.size_linked_shape(shape, context)
            except:
                cls._create_status_overlay(shape, context, "red", "exception")

    @classmethod
    def size_linked_shape(cls, shape, context):
        master_shape = cls.get_master_shape(shape, context)
        ref_heigth = master_shape.Height
        ref_width = master_shape.Width
        ref_lock_ar = master_shape.LockAspectRatio

        run_once = False
        for cShp in cls._iterate_linked_shapes(master_shape, context):
            cShp.LockAspectRatio = 0 #msoFalse
            cShp.Height, cShp.Width = ref_heigth, ref_width
            cShp.LockAspectRatio = ref_lock_ar
            run_once = True
    
        if run_once:
            cls._create_status_overlay(shape, context)
        else:
            cls._create_status_overlay(shape, context, "yellow", "no linked shapes found")

    @classmethod
    def align_linked_shapes(cls, shapes, context):
        for shape in shapes:
            try:
                cls.align_linked_shape(shape, context)
            except:
                cls._create_status_overlay(shape, context, "red", "exception")

    @classmethod
    def align_linked_shape(cls, shape, context):
        master_shape = cls.get_master_shape(shape, context)
        ref_position_left = master_shape.left
        ref_position_top = master_shape.top
        ref_rotation = master_shape.Rotation

        run_once = False
        for cShp in cls._iterate_linked_shapes(master_shape, context):
            cShp.left, cShp.top = ref_position_left, ref_position_top
            try:
                cShp.Rotation = ref_rotation
            except ValueError:
                #certain shape types do not support rotation, e.g. tables
                pass
            run_once = True
    
        if run_once:
            cls._create_status_overlay(shape, context)
        else:
            cls._create_status_overlay(shape, context, "yellow", "no linked shapes found")

    @classmethod
    def format_linked_shapes(cls, shapes, context):
        for shape in shapes:
            try:
                cls.format_linked_shape(shape, context)
            except:
                cls._create_status_overlay(shape, context, "red", "exception")

    @classmethod
    def format_linked_shape(cls, shape, context):
        master_shape = cls.get_master_shape(shape, context)
        try:
            master_shape.Pickup() #fails for some shapes, e.g. a group
            mode = "simple"
        except:
            if master_shape.Type == pplib.MsoShapeType["msoGroup"]:
                mode = "group"
            else:
                bkt.message.warning("Formatierung angleichen für gewähltes Shape nicht verfügbar.", "BKT: Verknüpfte Shapes")
                return

        run_once = False
        for cShp in cls._iterate_linked_shapes(master_shape, context):
            if mode == "simple":
                try:
                    cShp.Apply()
                except:
                    pass
            elif mode == "group" and cShp.Type == pplib.MsoShapeType["msoGroup"]:
                for index, iShp in enumerate(cShp.GroupItems, start=1):
                    try:
                        master_shape.GroupItems[index].Pickup()
                        iShp.Apply()
                    except:
                        pass
            run_once = True
        
        # Adjustment-Werte angleichen
        try:
            if master_shape.adjustments.count > 0:
                cls.adjustments_linked_shape(shape, context)
        except ValueError: #e.g. groups
            pass
    
        if run_once:
            cls._create_status_overlay(shape, context)
        else:
            cls._create_status_overlay(shape, context, "yellow", "no linked shapes found")

    @classmethod
    def adjustments_linked_shapes(cls, shapes, context):
        for shape in shapes:
            cls.adjustments_linked_shape(shape, context)

    @classmethod
    def adjustments_linked_shape(cls, shape, context):
        from .shape_adjustments import ShapeAdjustments
        shape = cls.get_master_shape(shape, context)
        for cShp in cls._iterate_linked_shapes(shape, context):
            ShapeAdjustments.equalize_adjustments([shape, cShp])

    @classmethod
    def text_linked_shapes(cls, shapes, context, with_formatting=True):
        for shape in shapes:
            try:
                cls.text_linked_shape(shape, context, with_formatting)
            except:
                cls._create_status_overlay(shape, context, "red", "exception")

    @classmethod
    def text_linked_shape(cls, shape, context, with_formatting=True):
        master_shape = cls.get_master_shape(shape, context)
        if master_shape.HasTextFrame == -1: #msoTrue
            if with_formatting:
                # shape.TextFrame2.TextRange.Copy()
                pass #nothing to do here as pplib.transfer_textrange function is used
            else:
                ref_text = master_shape.TextFrame2.TextRange.Text
            mode = "simple"
        elif master_shape.Type == pplib.MsoShapeType["msoGroup"]:
            mode = "group"
        else:
            bkt.message.warning("Text angleichen für gewähltes Shape nicht verfügbar.", "BKT: Verknüpfte Shapes")
            return

        run_once = False
        for cShp in cls._iterate_linked_shapes(master_shape, context):
            if mode == "simple":
                try:
                    if with_formatting:
                        # cShp.TextFrame2.TextRange.Paste()
                        pplib.transfer_textrange(master_shape.TextFrame2.TextRange, cShp.TextFrame2.TextRange)
                    else:
                        cShp.TextFrame2.TextRange.Text = ref_text
                except:
                    pass
            elif mode == "group" and cShp.Type == pplib.MsoShapeType["msoGroup"]:
                for index, iShp in enumerate(cShp.GroupItems, start=1):
                    try:
                        if with_formatting:
                            # shape.GroupItems[index].TextFrame2.TextRange.Copy()
                            # iShp.TextFrame2.TextRange.Paste()
                            pplib.transfer_textrange(master_shape.GroupItems[index].TextFrame2.TextRange, iShp.TextFrame2.TextRange)
                        else:
                            iShp.TextFrame2.TextRange.Text = master_shape.GroupItems[index].TextFrame2.TextRange.Text
                    except:
                        pass
            run_once = True
    
        if run_once:
            cls._create_status_overlay(shape, context)
        else:
            cls._create_status_overlay(shape, context, "yellow", "no linked shapes found")

    @classmethod
    def equalize_linked_shapes(cls, shapes, context):
        for shape in shapes:
            try:
                cls.size_linked_shape(shape, context)
            except:
                pass
            
            try:
                cls.align_linked_shape(shape, context)
            except:
                pass
            
            try:
                cls.format_linked_shape(shape, context)
            except:
                pass
            
            try:
                cls.text_linked_shape(shape, context)
            except:
                pass

    @classmethod
    def delete_linked_shapes(cls, shapes, context):
        for shape in shapes:
            cls.delete_linked_shape(shape, context)

    @classmethod
    def delete_linked_shape(cls, shape, context):
        for cShp in cls._iterate_linked_shapes(shape, context):
            cShp.Delete()

    @classmethod
    def replace_with_this(cls, shapes, context):
        for shape in shapes:
            cls.replace_with_this_shape(shape, context)

    @classmethod
    def replace_with_this_shape(cls, shape, context):
        shape = cls.get_master_shape(shape, context)
        shape.Copy()
        for cShp in cls._iterate_linked_shapes(shape, context):
            slide = cShp.Parent
            ref_zorder = cShp.ZOrderPosition
            cShp.Delete()
            new = slide.Shapes.Paste()
            pplib.set_shape_zorder(new, value=ref_zorder)

    ### ACTIONS ###
    @classmethod
    def linked_shapes_toback(cls, shapes, context):
        for shape in shapes:
            shape.ZOrder(1)
            for cShp in cls._iterate_linked_shapes(shape, context):
                cShp.ZOrder(1) #0=msoBringToFront, 1=msoSendToBack

    @classmethod
    def linked_shapes_tofront(cls, shapes, context):
        for shape in shapes:
            shape.ZOrder(0)
            for cShp in cls._iterate_linked_shapes(shape, context):
                cShp.ZOrder(0) #0=msoBringToFront, 1=msoSendToBack

    @classmethod
    def linked_shapes_flipv(cls, shapes, context):
        for shape in shapes:
            shape.Flip(1) #msoFlipVertical
            for cShp in cls._iterate_linked_shapes(shape, context):
                cShp.Flip(1) #msoFlipVertical

    @classmethod
    def linked_shapes_fliph(cls, shapes, context):
        for shape in shapes:
            shape.Flip(0) #msoFlipHorizontal
            for cShp in cls._iterate_linked_shapes(shape, context):
                cShp.Flip(0) #msoFlipHorizontal

    @classmethod
    def linked_shapes_slidenum(cls, shapes, context):
        for shape in shapes:
            if shape.HasTextFrame:
                shape.TextFrame.TextRange.InsertSlideNumber() #InsertSlideNumber only in TextRange, not TextRange2!
            for cShp in cls._iterate_linked_shapes(shape, context):
                try:
                    cShp.TextFrame.TextRange.InsertSlideNumber()
                except:
                    pass

    @classmethod
    def linked_shapes_changecase(cls, shapes, context, mode=1):
        # MsoTextChangeCase:
        # msoCaseLower	2	Zeigt den Text in Kleinbuchstaben an.
        # msoCaseSentence	1	Der erste Buchstabe im Satz wird großgeschrieben. Für alle anderen Buchstaben gilt die entsprechende Groß-/Kleinschreibung (Substantive, Akronyme usw. werden großgeschrieben).
        # msoCaseTitle	4	Der erste Buchstabe aller Wörter im Titel wird großgeschrieben. Alle anderen Buchstaben werden kleingeschrieben. In bestimmten Fällen werden kurze Artikel, Präpositionen und Konjunktionen nicht großgeschrieben.
        # msoCaseToggle	5	Gibt an, dass kleingeschriebener Text in großgeschriebenen Text und umgekehrt konvertiert werden soll.
        # msoCaseUpper	3	Zeigt den Text in Großbuchstaben an.
        for shape in shapes:
            if shape.HasTextFrame:
                shape.TextFrame2.TextRange.ChangeCase(mode)
            for cShp in cls._iterate_linked_shapes(shape, context):
                try:
                    cShp.TextFrame2.TextRange.ChangeCase(mode)
                except:
                    pass

    ### PROPERTIES ###
    @classmethod
    def linked_shapes_custom(cls, shapes, context, property_name, wrap=True):
        for shape in shapes:
            shape = cls.get_master_shape(shape, context)
            wrap_shape = lambda shp: shp if not wrap else pplib.wrap_shape(shp)
            cur_value = getattr(wrap_shape(shape), property_name)
            for cShp in cls._iterate_linked_shapes(shape, context):
                cShp = wrap_shape(cShp)
                try:
                    setattr(cShp, property_name, cur_value)
                except:
                    #not all properties supported by all shapes (e.g. rotation not supported by tables)
                    pass

    ### CUSTOM ONE-TIME DEV METHODS ###
    # @classmethod
    # def linked_shapes_xyz(cls, shape, context):
    #     def svgconv(cShp):
    #         cShp.parent.select()
    #         cShp.select()
    #         context.app.CommandBars.ExecuteMso("SVGEdit")

    #     for cShp in cls._iterate_linked_shapes(shape, context):
    #         svgconv(cShp)
    #     svgconv(shape)




linkshapes_tab = bkt.ribbon.Tab(
    id = "bkt_context_tab_linkshapes",
    label = "[BKT] Verknüpfte Shapes",
    get_visible=bkt.Callback(LinkedShapes.are_linked_shapes, shapes=True),
    children = [
        bkt.ribbon.Group(
            id="bkt_linkshapes_find_group",
            label = "Verknüpfte Shapes finden",
            get_visible=bkt.apps.ppt_shapes_exactly1_selected,
            children = [
                bkt.ribbon.Box(box_style="horizontal", children=[
                    bkt.ribbon.Button(
                        id = 'linked_shapes_first',
                        label="Erstes verknüpfte Shape finden",
                        show_label=False,
                        image_mso="MailMergeGoToFirstRecord",
                        screentip="Zum ersten verknüpften Shape gehen",
                        supertip="Sucht nach dem ersten verknüpften Shape.",
                        on_action=bkt.Callback(LinkedShapes.goto_first_shape, shape=True, context=True),
                    ),
                    bkt.ribbon.Button(
                        id = 'linked_shapes_previous',
                        label="Vorheriges verknüpfte Shape finden",
                        show_label=False,
                        image_mso="MailMergeGoToPreviousRecord",
                        screentip="Zum vorherigen verknüpften Shape gehen",
                        supertip="Sucht nach dem vorherigen verknüpften Shape. Sollte auf den vorherigen Folien kein Shape mehr kommen, wird das letzte verknüpfte Shape der Präsentation gesucht.",
                        on_action=bkt.Callback(LinkedShapes.goto_previous_shape, shape=True, context=True),
                    ),
                    bkt.ribbon.Label(
                        label="Gehe zu",
                    ),
                    bkt.ribbon.Button(
                        id = 'linked_shapes_next',
                        label="Nächstes verknüpfte Shape finden",
                        show_label=False,
                        image_mso="MailMergeGoToNextRecord",
                        screentip="Zum nächsten verknüpften Shape gehen",
                        supertip="Sucht nach dem nächste verknüpften Shape. Sollte auf den Folgefolien kein Shape mehr kommen, wird das erste verknüpfte Shape der Präsentation gesucht.",
                        on_action=bkt.Callback(LinkedShapes.goto_next_shape, shape=True, context=True),
                    ),
                    bkt.ribbon.Button(
                        id = 'linked_shapes_last',
                        label="Letztes verknüpfte Shape finden",
                        show_label=False,
                        image_mso="MailMergeGotToLastRecord",
                        screentip="Zum letzten verknüpften Shape gehen",
                        supertip="Sucht nach dem letzten verknüpften Shape.",
                        on_action=bkt.Callback(LinkedShapes.goto_last_shape, shape=True, context=True),
                    ),
                ]),
                bkt.ribbon.Button(
                    id = 'linked_shapes_count',
                    label="Shapes zählen",
                    image_mso="FormattingUnique",
                    screentip="Alle verknüpften Shapes zählen",
                    supertip="Zählt die Anzahl der verknüpften Shapes auf allen Folien.",
                    on_action=bkt.Callback(LinkedShapes.count_link_shapes, shape=True, context=True),
                ),
                bkt.ribbon.Button(
                    id = 'linked_shapes_select',
                    label="Folien anzeigen",
                    image_mso="SlideTransitionApplyToAll",
                    screentip="Alle Foliennummern mit verknüpften Shapes anzeigen",
                    supertip="Zeigt alle Foliennummern die zugehörige verknüpfte Shapes enthalten.",
                    on_action=bkt.Callback(LinkedShapes.select_link_shapes_slides, shape=True, context=True),
                ),
            ]
        ),
        bkt.ribbon.Group(
            id="bkt_linkshapes_align_group",
            label = "Verknüpfte Shapes angleichen",
            children = [
                bkt.ribbon.Menu(
                    id = 'linked_shapes_master',
                    label="Referenz wählen",
                    image_mso="CircularReferences",
                    size="large",
                    screentip="Referenzshape auswählen",
                    supertip="Auswählen, ob selektiertes, erstes oder letztes Shape als Referenz für alle Angleichungsfunktionen verwendet werden soll. Standard ist das aktuell ausgewählte Shape.",
                    children=[
                        bkt.ribbon.ToggleButton(
                            label="Ausgewähltes Shapes (Standard)",
                            supertip="Referenz für alle verknüpften Shapes ist das aktuell gewählte Shape",
                            get_pressed=bkt.Callback(lambda: LinkedShapes.master == "current"),
                            on_toggle_action=bkt.Callback(lambda pressed: setattr(LinkedShapes, "master", "current")),
                        ),
                        bkt.ribbon.ToggleButton(
                            label="Erstes Shape im Foliensatz",
                            supertip="Die gesamte Präsentation wird nach verknüpften Shapes gescannt und das erste zugehörige Shape im Foliensatz wird als Referenz gesetzt.",
                            get_pressed=bkt.Callback(lambda: LinkedShapes.master == "first"),
                            on_toggle_action=bkt.Callback(lambda pressed: setattr(LinkedShapes, "master", "first")),
                        ),
                        bkt.ribbon.ToggleButton(
                            label="Letztes Shape im Foliensatz",
                            supertip="Die gesamte Präsentation wird nach verknüpften Shapes gescannt und das letzte zugehörige Shape im Foliensatz wird als Referenz gesetzt.",
                            get_pressed=bkt.Callback(lambda: LinkedShapes.master == "last"),
                            on_toggle_action=bkt.Callback(lambda pressed: setattr(LinkedShapes, "master", "last")),
                        ),
                        bkt.ribbon.MenuSeparator(),
                        bkt.ribbon.ToggleButton(
                            label="Status-Overlays erstellen",
                            supertip="Legt über das markierte Shape ein Ampelstatus-Overlay welches anzeigt, ob die Operation erfolgreich war und Shapes aktualisiert wurden. Funktioniert für Größe, Position, Formatierung und Text angleichen.",
                            get_pressed=bkt.Callback(lambda: LinkedShapes.status_overlay),
                            on_toggle_action=bkt.Callback(lambda pressed: setattr(LinkedShapes, "status_overlay", pressed)),
                        ),
                    ]
                ),
                bkt.ribbon.Button(
                    id = 'linked_shapes_all',
                    label="Alles angleichen",
                    image_mso="GroupUpdate",
                    size="large",
                    screentip="Alle Eigenschaften verknüpfter Shapes angleichen",
                    supertip="Alle Eigenschaften aller verknüpfter Shapes wie ausgewähltes Shape setzen.",
                    on_action=bkt.Callback(LinkedShapes.equalize_linked_shapes, shapes=True, context=True),
                ),
                bkt.ribbon.Separator(),
                bkt.ribbon.Button(
                    id = 'linked_shapes_align',
                    label="Position angleichen",
                    image_mso="ControlAlignToGrid",
                    screentip="Position verknüpfter Shapes angleichen",
                    supertip="Position und Rotation aller verknüpfter Shapes auf Position wie ausgewähltes Shape setzen.",
                    on_action=bkt.Callback(LinkedShapes.align_linked_shapes, shapes=True, context=True),
                ),
                bkt.ribbon.Button(
                    id = 'linked_shapes_size',
                    label="Größe angleichen",
                    image_mso="SizeToControlHeightAndWidth",
                    screentip="Größe verknüpfter Shapes angleichen",
                    supertip="Größe aller verknüpfter Shapes auf Größe wie ausgewähltes Shape setzen.",
                    on_action=bkt.Callback(LinkedShapes.size_linked_shapes, shapes=True, context=True),
                ),
                bkt.ribbon.Button(
                    id = 'linked_shapes_format',
                    label="Formatierung angleichen",
                    image_mso="FormatPainter",
                    screentip="Formatierung verknüpfter Shapes angleichen",
                    supertip="Formatierung aller verknüpfter Shapes auf Größe wie ausgewähltes Shape setzen.",
                    on_action=bkt.Callback(LinkedShapes.format_linked_shapes, shapes=True, context=True),
                ),
                bkt.ribbon.Button(
                    id = 'linked_shapes_text',
                    label="Text angleichen",
                    image_mso="TextBoxInsert",
                    screentip="Text verknüpfter Shapes angleichen",
                    supertip="Text aller verknüpfter Shapes auf Größe wie ausgewähltes Shape setzen.",
                    on_action=bkt.Callback(LinkedShapes.text_linked_shapes, shapes=True, context=True),
                ),
                bkt.ribbon.Menu(
                    id="linked_shapes_actions",
                    label="Aktion ausführen",
                    supertip="Diverse Aktionen auf alle verknüpften Shapes ausführen",
                    image_mso="ObjectBringToFront",
                    children=[
                        bkt.ribbon.Button(
                            id = 'linked_shapes_tofront',
                            label="In den Vordergrund",
                            image_mso="ObjectBringToFront",
                            screentip="Verknüpfte Shapes in den Vordergrund",
                            supertip="Alle verknüpften Shapes in den Vordergrund bringen",
                            on_action=bkt.Callback(LinkedShapes.linked_shapes_tofront, shapes=True, context=True),
                        ),
                        bkt.ribbon.Button(
                            id = 'linked_shapes_toback',
                            label="In den Hintergrund",
                            image_mso="ObjectSendToBack",
                            screentip="Verknüpfte Shapes in den Hintergrund",
                            supertip="Alle verknüpften Shapes in den Hintergrund bringen",
                            on_action=bkt.Callback(LinkedShapes.linked_shapes_toback, shapes=True, context=True),
                        ),
                        bkt.ribbon.MenuSeparator(),
                        bkt.ribbon.Button(
                            id = 'linked_shapes_fliph',
                            label="Horizontal spiegeln",
                            image_mso="ObjectFlipHorizontal",
                            screentip="Verknüpfte Shapes horizontal spiegeln",
                            supertip="Alle verknüpften Shapes horizontal spiegeln",
                            on_action=bkt.Callback(LinkedShapes.linked_shapes_fliph, shapes=True, context=True),
                        ),
                        bkt.ribbon.Button(
                            id = 'linked_shapes_flipv',
                            label="Vertikal spiegeln",
                            image_mso="ObjectFlipVertical",
                            screentip="Verknüpfte Shapes vertikal spiegeln",
                            supertip="Alle verknüpften Shapes vertikal spiegeln",
                            on_action=bkt.Callback(LinkedShapes.linked_shapes_flipv, shapes=True, context=True),
                        ),
                        bkt.ribbon.MenuSeparator(),
                        bkt.ribbon.Button(
                            id = 'linked_shapes_slidenum',
                            label="Foliennummer einfügen",
                            image_mso="NumberInsert",
                            screentip="Verknüpfte Shapes aktualisierbare Foliennummer anstellen",
                            supertip="Fügt allen verknüpften Shapes am Ende vom Text automatisch aktualisierbare Foliennummer an",
                            on_action=bkt.Callback(LinkedShapes.linked_shapes_slidenum, shapes=True, context=True),
                        ),
                        bkt.ribbon.Menu(
                            id='linked_shapes_changecase',
                            label="Groß-/Kleinschreibung ändern",
                            supertip="Groß-/Kleinschreibung für alle verknüpften Shapes anpassen",
                            image_mso="ChangeCaseGallery",
                            children=[
                                bkt.ribbon.Button(
                                    id = 'linked_shapes_changecase-1',
                                    label="Ersten Buchstaben im Satz großschreiben",
                                    supertip="Ersten Buchstaben im Satz aller verknüpften Shapes großschreiben",
                                    on_action=bkt.Callback(lambda shapes, context: LinkedShapes.linked_shapes_changecase(shapes, context, 1), shapes=True, context=True),
                                ),
                                bkt.ribbon.Button(
                                    id = 'linked_shapes_changecase-2',
                                    label="kleinbuchstaben",
                                    supertip="Text aller verknüpften Shapes kleinschreiben",
                                    on_action=bkt.Callback(lambda shapes, context: LinkedShapes.linked_shapes_changecase(shapes, context, 2), shapes=True, context=True),
                                ),
                                bkt.ribbon.Button(
                                    id = 'linked_shapes_changecase-3',
                                    label="GROẞBUCHSTABEN",
                                    supertip="Text aller verknüpften Shapes großschreiben",
                                    on_action=bkt.Callback(lambda shapes, context: LinkedShapes.linked_shapes_changecase(shapes, context, 3), shapes=True, context=True),
                                ),
                                bkt.ribbon.Button(
                                    id = 'linked_shapes_changecase-4',
                                    label="Ersten Buchstaben Im Wort Großschreiben",
                                    supertip="Ersten Buchstaben im Wort aller verknüpften Shapes großschreiben",
                                    on_action=bkt.Callback(lambda shapes, context: LinkedShapes.linked_shapes_changecase(shapes, context, 4), shapes=True, context=True),
                                ),
                                bkt.ribbon.Button(
                                    id = 'linked_shapes_changecase-5',
                                    label="gROẞ-/kLEINSCHREIBUNG umkehren",
                                    supertip="Groß-/Kleinschreibung aller verknüpften Shapes umkehren",
                                    on_action=bkt.Callback(lambda shapes, context: LinkedShapes.linked_shapes_changecase(shapes, context, 5), shapes=True, context=True),
                                ),
                            ]
                        )
                    ]
                ),
                bkt.ribbon.Menu(
                    id="linked_shapes_properties",
                    label="Eigenschaft angleichen",
                    supertip="Eine einzelne Eigenschaft auf alle verknüpften Shapes übertragen",
                    image_mso="ObjectNudgeRight",
                    children=[
                        bkt.ribbon.Button(
                            id = 'linked_shapes_custom-text',
                            label="Text (ohne Formatierung)",
                            screentip="Text (ohne Formatierung) angleichen",
                            # image_mso="TextBoxInsert",
                            supertip="Text ohne Formatierungen für alle verknüpften Shapes angleichen",
                            on_action=bkt.Callback(lambda shapes, context: LinkedShapes.text_linked_shapes(shapes, context, with_formatting=False), shapes=True, context=True),
                        ),
                        bkt.ribbon.MenuSeparator(),
                        bkt.ribbon.Button(
                            id = 'linked_shapes_custom-lar',
                            label="Seitenverhältnis gesperrt",
                            screentip="Seitenverhältnis gesperrt angleichen",
                            # image_mso="ObjectBringToFront",
                            supertip="Seitenverhältnis sperren an/aus für alle verknüpften Shapes angleichen",
                            on_action=bkt.Callback(lambda shapes, context: LinkedShapes.linked_shapes_custom(shapes, context, "LockAspectRatio", False), shapes=True, context=True),
                        ),
                        bkt.ribbon.Button(
                            id = 'linked_shapes_custom-rot',
                            label="Rotation",
                            screentip="Rotation angleichen",
                            # image_mso="ObjectBringToFront",
                            supertip="Rotation für alle verknüpften Shapes angleichen",
                            on_action=bkt.Callback(lambda shapes, context: LinkedShapes.linked_shapes_custom(shapes, context, "Rotation", False), shapes=True, context=True),
                        ),
                        bkt.ribbon.Button(
                            id = 'linked_shapes_custom-bwmode',
                            label="Schwarz-Weiß-Modus",
                            screentip="Schwarz-Weiß-Modus angleichen",
                            # image_mso="ObjectBringToFront",
                            supertip="Schwarz-Weiß-Modus für alle verknüpften Shapes angleichen",
                            on_action=bkt.Callback(lambda shapes, context: LinkedShapes.linked_shapes_custom(shapes, context, "BlackWhiteMode", False), shapes=True, context=True),
                        ),
                        bkt.ribbon.MenuSeparator(),
                        bkt.ribbon.Button(
                            id = 'linked_shapes_custom-left',
                            label="Linke Seite",
                            screentip="Linke Seite angleichen",
                            # image_mso="ObjectBringToFront",
                            supertip="Linke Seite für alle verknüpften Shapes angleichen",
                            on_action=bkt.Callback(lambda shapes, context: LinkedShapes.linked_shapes_custom(shapes, context, "x"), shapes=True, context=True),
                        ),
                        bkt.ribbon.Button(
                            id = 'linked_shapes_custom-right',
                            label="Rechte Seite",
                            screentip="Rechte Seite angleichen",
                            # image_mso="ObjectBringToFront",
                            supertip="Rechte Seite für alle verknüpften Shapes angleichen",
                            on_action=bkt.Callback(lambda shapes, context: LinkedShapes.linked_shapes_custom(shapes, context, "x1"), shapes=True, context=True),
                        ),
                        bkt.ribbon.Button(
                            id = 'linked_shapes_custom-top',
                            label="Obere Seite",
                            screentip="Obere Seite angleichen",
                            # image_mso="ObjectBringToFront",
                            supertip="Obere Seite für alle verknüpften Shapes angleichen",
                            on_action=bkt.Callback(lambda shapes, context: LinkedShapes.linked_shapes_custom(shapes, context, "y"), shapes=True, context=True),
                        ),
                        bkt.ribbon.Button(
                            id = 'linked_shapes_custom-bottom',
                            label="Untere Seite",
                            screentip="Untere Seite angleichen",
                            # image_mso="ObjectBringToFront",
                            supertip="Untere Seite für alle verknüpften Shapes angleichen",
                            on_action=bkt.Callback(lambda shapes, context: LinkedShapes.linked_shapes_custom(shapes, context, "y1"), shapes=True, context=True),
                        ),
                        bkt.ribbon.Button(
                            id = 'linked_shapes_custom-centerx',
                            label="Mittelpunkt links",
                            screentip="Mittelpunkt links angleichen",
                            # image_mso="ObjectBringToFront",
                            supertip="Mittelpunkt links für alle verknüpften Shapes angleichen",
                            on_action=bkt.Callback(lambda shapes, context: LinkedShapes.linked_shapes_custom(shapes, context, "center_x"), shapes=True, context=True),
                        ),
                        bkt.ribbon.Button(
                            id = 'linked_shapes_custom-centery',
                            label="Mittelpunkt oben",
                            screentip="Mittelpunkt oben angleichen",
                            # image_mso="ObjectBringToFront",
                            supertip="Mittelpunkt oben für alle verknüpften Shapes angleichen",
                            on_action=bkt.Callback(lambda shapes, context: LinkedShapes.linked_shapes_custom(shapes, context, "center_y"), shapes=True, context=True),
                        ),
                        bkt.ribbon.MenuSeparator(),
                        bkt.ribbon.Button(
                            id = 'linked_shapes_custom-width',
                            label="Breite",
                            # image_mso="ObjectBringToFront",
                            supertip="Breite für alle verknüpften Shapes angleichen",
                            on_action=bkt.Callback(lambda shapes, context: LinkedShapes.linked_shapes_custom(shapes, context, "width", False), shapes=True, context=True),
                        ),
                        bkt.ribbon.Button(
                            id = 'linked_shapes_custom-height',
                            label="Höhe",
                            # image_mso="ObjectBringToFront",
                            supertip="Höhe für alle verknüpften Shapes angleichen",
                            on_action=bkt.Callback(lambda shapes, context: LinkedShapes.linked_shapes_custom(shapes, context, "height", False), shapes=True, context=True),
                        ),
                    ]
                ),
                bkt.ribbon.Separator(),
                bkt.ribbon.Button(
                    id = 'linked_shapes_delete',
                    label="Andere Shapes löschen",
                    image_mso="HyperlinkRemove",
                    screentip="Verknüpfte Shapes löschen",
                    supertip="Alle verknüpften Shapes auf allen Folien löschen.",
                    on_action=bkt.Callback(LinkedShapes.delete_linked_shapes, shapes=True, context=True),
                ),
                bkt.ribbon.Button(
                    id = 'linked_shapes_replace',
                    label="Mit Referenz ersetzen",
                    image_mso="HyperlinkCreate",
                    screentip="Verknüpfte Shapes ersetzen",
                    supertip="Alle verknüpften Shapes auf allen Folien mit Referenz-Shape (standardmäßig das ausgwählte Shape) ersetzen.",
                    on_action=bkt.Callback(LinkedShapes.replace_with_this, shapes=True, context=True),
                ),
                bkt.ribbon.Button(
                    id = 'linked_shapes_search',
                    label="Weitere Shapes suchen",
                    image_mso="FindTag",
                    screentip="Gleiches Shape auf Folgefolien suchen und verknüpfen",
                    supertip="Erneut nach Shapes anhand Position und Größe suche, um weitere Shapes zu dieser Verknüpfung hinzuzufügen.",
                    get_enabled=bkt.apps.ppt_shapes_exactly1_selected,
                    on_action=bkt.Callback(LinkedShapes.find_similar_and_link, shape=True, context=True),
                ),
                ### Custom action
                # bkt.ribbon.Button(
                #     id = 'linked_shapes_xyz',
                #     label="Custom Action",
                #     image_mso="HappyFace",
                #     on_action=bkt.Callback(LinkedShapes.linked_shapes_xyz, shape=True, context=True),
                #     # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                # ),
            ]
        ),
        bkt.ribbon.Group(
            id="bkt_linkshapes_unlink_group",
            label = "Verknüpfung aufheben",
            children = [
                bkt.ribbon.Button(
                    id = 'linked_shapes_unlink',
                    label="Einzelne Shape-Verknüpfung entfernen",
                    image_mso="HyperlinkRemove",
                    screentip="Verknüpfung des ausgewählten Shapes entfernen",
                    supertip="Entfernt die ID zur Verknüpfung vom aktuellen Shape. Alle anderen Shapes mit der gleichen ID bleiben verknüpft.",
                    on_action=bkt.Callback(LinkedShapes.unlink_shapes, shapes=True),
                ),
                bkt.ribbon.Button(
                    id = 'linked_shapes_unlink_all',
                    label="Gesamte Shape-Verknüpfung auflösen",
                    image_mso="HyperlinkRemove",
                    screentip="Alle Shape-Verknüpfungen entfernen",
                    supertip="Entfernt die ID zur Verknüpfung vom aktuellen Shape sowie allen verknüpften Shapes mit der gleichen ID.",
                    on_action=bkt.Callback(LinkedShapes.unlink_all_shapes, shapes=True, context=True),
                ),
            ]
        ),
    ]
)