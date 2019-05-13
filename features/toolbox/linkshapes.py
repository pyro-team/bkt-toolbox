# -*- coding: utf-8 -*-
'''
Created on 06.09.2018

@author: fstallmann
'''

import bkt
import bkt.library.powerpoint as pplib

import uuid
import os.path


BKT_LINK_UUID = "BKT_LINK_UUID"

class LinkedShapes(object):
    current_link_guid = None

    @staticmethod
    def _add_tags(shape, link_guid):
        shape.Tags.Add(BKT_LINK_UUID, link_guid)
        shape.Tags.Add(bkt.contextdialogs.BKT_CONTEXTDIALOG_TAGKEY, BKT_LINK_UUID)

    @staticmethod
    def is_linked_shape(shape):
        try:
            return shape.Tags.Item(BKT_LINK_UUID) != ''
        except:
            #Shape.Tags throws COMException for SmartArt shapes
            return False

    @classmethod
    def not_is_linked_shape(cls, shape):
        return not cls.is_linked_shape(shape)


    @classmethod
    def enabled_add_linked_shapes(cls):
        return cls.current_link_guid != None

    @classmethod
    def find_similar_and_link(cls, shape, context): #shape as parameter for get_enabled required
        #open wpf window
        from dialogs.linkshapes_find import FindWindow
        FindWindow.create_and_show_dialog(cls, context)

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
        comparer_values = lambda val1, val2: is_close(val1, val2, threshold)
        
        all_slides = context.app.ActivePresentation.Slides
        num_slides = limit_slides or all_slides.Count

        for slide in context.app.ActivePresentation.Slides:
            if slide.SlideIndex <= active_slide_index:
                continue
            for sld_shape in slide.Shapes:
                cShp = pplib.wrap_shape(sld_shape)
                if (
                    cShp.Type == shape.Type and
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
            bkt.helpers.message("Keine vergleichbaren Shapes gefunden.")
        elif dry_run:
            bkt.helpers.message("Es wurden %s Shapes zum verknüpfen gefunden." % shapes_found)
        else:
            cls._add_tags(shape, link_guid)
            bkt.helpers.message("Das Shape wurde mit %s Shapes verknüpft." % shapes_found)
            context.ribbon.ActivateTab('bkt_context_tab_linkshapes')

    @classmethod
    def copy_to_all(cls, shape, context): #shape as parameter for get_enabled required
        #open wpf window
        from dialogs.linkshapes_copy import CopyWindow
        CopyWindow.create_and_show_dialog(cls, context)

    @classmethod
    def copy_shapes_to_slides(cls, shapes, context, limit_slides=None):
        for shape in shapes:
            cls.copy_shape_to_slides(shape, context, limit_slides)
        
        shapes[0].select()
        context.ribbon.ActivateTab('bkt_context_tab_linkshapes')

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
    def unlink_shape(cls, shape):
        shape.Tags.Delete(BKT_LINK_UUID)
    
    @classmethod
    def unlink_all_shapes(cls, shape, context):
        for cShp in cls._iterate_linked_shapes(shape, context):
            cls.unlink_shape(cShp)
        cls.unlink_shape(shape)

    @classmethod
    def extend_link_shapes(cls, shape):
        cls.current_link_guid = shape.Tags.Item(BKT_LINK_UUID)
        bkt.helpers.message('Die Link-ID wurde zwischengespeichert. Als nächstes können weitere Shapes ausgewählt und über "Ausgewählte Shapes zur Verknüpfung hinzufügen" mit diesem Shape verknüpft werden.')

    @classmethod
    def add_to_link_shapes(cls, shapes):
        for shape in shapes:
            cls._add_tags(shape, cls.current_link_guid)

    @classmethod
    def count_link_shapes(cls, shape, context):
        count_shapes = 0
        for cShp in cls._iterate_linked_shapes(shape, context):
            count_shapes += 1
        bkt.helpers.message("Es wurden %s verknüpfte Shapes gefunden." % count_shapes)
    
    @classmethod
    def goto_linked_shape(cls, shape, context, goto=1, delta=True): #goto=1 -> next linked shape, delta=False -> goto interpreted as absolute index
        link_guid = shape.Tags.Item(BKT_LINK_UUID)
        cur_shape_index = None
        list_shapes = []

        for slide in context.app.ActivePresentation.Slides:
            for cShp in slide.Shapes:
                if cShp.Tags.Item(BKT_LINK_UUID) == link_guid:
                    list_shapes.append(cShp)
                if cShp == shape:
                    cur_shape_index = len(list_shapes)-1
        
        sel_shape_index = goto if not delta else (cur_shape_index + goto) % len(list_shapes)
        list_shapes[sel_shape_index].parent.select() #select slide
        list_shapes[sel_shape_index].select() #select shape

    @classmethod
    def goto_first_shape(cls, shape, context):
        cls.goto_linked_shape(shape, context, 0, False)

    @classmethod
    def _iterate_linked_shapes(cls, shape, context):
        link_guid = shape.Tags.Item(BKT_LINK_UUID)
        
        if link_guid == "":
            raise IndexError("Shape has no link uuid")

        # active_slide_index = shape.Parent.SlideIndex
        for slide in context.app.ActivePresentation.Slides:
            # if slide.SlideIndex == active_slide_index:
            #     continue
            for cShp in list(iter(slide.Shapes)): #list(iter()) required for deletion of shapes
                if cShp == shape:
                    continue
                if cShp.Tags.Item(BKT_LINK_UUID) == link_guid:
                    # apply_func(cShp)
                    yield cShp

    @classmethod
    def size_linked_shapes(cls, shape, context):
        ref_heigth = shape.Height
        ref_width = shape.Width
        ref_lock_ar = shape.LockAspectRatio

        for cShp in cls._iterate_linked_shapes(shape, context):
            cShp.LockAspectRatio = 0 #msoFalse
            cShp.Height, cShp.Width = ref_heigth, ref_width
            cShp.LockAspectRatio = ref_lock_ar

    @classmethod
    def align_linked_shapes(cls, shape, context):
        ref_position_left = shape.left
        ref_position_top = shape.top
        ref_rotation = shape.Rotation

        for cShp in cls._iterate_linked_shapes(shape, context):
            cShp.left, cShp.top = ref_position_left, ref_position_top
            cShp.Rotation = ref_rotation

    @classmethod
    def format_linked_shapes(cls, shape, context):
        try:
            shape.Pickup() #fails for some shapes, e.g. a group
            mode = "simple"
        except:
            if shape.Type == pplib.MsoShapeType["msoGroup"]:
                mode = "group"
            else:
                bkt.helpers.message("Formatierung angleichen für gewähltes Shape nicht verfügbar.")
                return

        for cShp in cls._iterate_linked_shapes(shape, context):
            if mode == "simple":
                try:
                    cShp.Apply()
                except:
                    pass
            elif mode == "group" and cShp.Type == pplib.MsoShapeType["msoGroup"]:
                for index, iShp in enumerate(cShp.GroupItems):
                    try:
                        shape.GroupItems[index+1].Pickup()
                        iShp.Apply()
                    except:
                        pass
        # Adjustment-Werte angleichen
        cls.adjustments_linked_shapes(shape, context)

    @classmethod
    def adjustments_linked_shapes(cls, shape, context):
        from shape_adjustments import ShapeAdjustments
        for cShp in cls._iterate_linked_shapes(shape, context):
            ShapeAdjustments.equalize_adjustments([shape, cShp])

    @classmethod
    def text_linked_shapes(cls, shape, context):
        if shape.HasTextFrame == -1: #msoTrue
            ref_text = shape.TextFrame2.TextRange.Text
            mode = "simple"
        elif shape.Type == pplib.MsoShapeType["msoGroup"]:
            mode = "group"
        else:
            bkt.helpers.message("Text angleichen für gewähltes Shape nicht verfügbar.")
            return

        for cShp in cls._iterate_linked_shapes(shape, context):
            if mode == "simple":
                try:
                    cShp.TextFrame2.TextRange.Text = ref_text
                except:
                    pass
            elif mode == "group" and cShp.Type == pplib.MsoShapeType["msoGroup"]:
                for index, iShp in enumerate(cShp.GroupItems):
                    try:
                        iShp.TextFrame2.TextRange.Text = shape.GroupItems[index+1].TextFrame2.TextRange.Text
                    except:
                        pass

    @classmethod
    def equalize_linked_shapes(cls, shape, context):
        try:
            cls.size_linked_shapes(shape, context)
        except:
            pass
        
        try:
            cls.align_linked_shapes(shape, context)
        except:
            pass
        
        try:
            cls.format_linked_shapes(shape, context)
        except:
            pass
        
        try:
            cls.text_linked_shapes(shape, context)
        except:
            pass

    @classmethod
    def delete_linked_shapes(cls, shape, context):
        for cShp in cls._iterate_linked_shapes(shape, context):
            cShp.Delete()

    @classmethod
    def replace_with_this(cls, shape, context):
        shape.Copy()
        for cShp in cls._iterate_linked_shapes(shape, context):
            slide = cShp.Parent
            ref_zorder = cShp.ZOrderPosition
            cShp.Delete()
            new = slide.Shapes.Paste()
            pplib.set_shape_zorder(new, value=ref_zorder)

    @classmethod
    def linked_shapes_toback(cls, shape, context):
        shape.ZOrder(1)
        for cShp in cls._iterate_linked_shapes(shape, context):
            cShp.ZOrder(1) #0=msoBringToFront, 1=msoSendToBack

    @classmethod
    def linked_shapes_tofront(cls, shape, context):
        shape.ZOrder(0)
        for cShp in cls._iterate_linked_shapes(shape, context):
            cShp.ZOrder(0) #0=msoBringToFront, 1=msoSendToBack




def linked_shapes_context_menu(prefix):
    return [
        bkt.ribbon.SplitButton(
            id=prefix+'-linked-shapes',
            # label="Verknüpfte Shapes",
            # image_mso='HyperlinkCreate',
            insertBeforeMso='ObjectsGroupMenu',
            get_visible=bkt.Callback(LinkedShapes.is_linked_shape, shape=True),
            children=[
                bkt.ribbon.Button(
                    id = prefix+'-linked-shapes-all',
                    label="Verknüpfte Shapes angleichen",
                    # image_mso="GroupUpdate",
                    image_mso='HyperlinkCreate',
                    on_action=bkt.Callback(LinkedShapes.equalize_linked_shapes, shape=True, context=True),
                    # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.Menu(children=[
                    bkt.ribbon.Button(
                        id = prefix+'-linked-shapes-all2',
                        label="Alles angleichen",
                        # image_mso="GroupUpdate",
                        image_mso='HyperlinkCreate',
                        on_action=bkt.Callback(LinkedShapes.equalize_linked_shapes, shape=True, context=True),
                        # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.MenuSeparator(),
                    bkt.ribbon.Button(
                        id = prefix+'-linked-shapes-align',
                        label="Position angleichen",
                        image_mso="ControlAlignToGrid",
                        on_action=bkt.Callback(LinkedShapes.align_linked_shapes, shape=True, context=True),
                        # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = prefix+'-linked-shapes-size',
                        label="Größe angleichen",
                        image_mso="SizeToControlHeightAndWidth",
                        on_action=bkt.Callback(LinkedShapes.size_linked_shapes, shape=True, context=True),
                        # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = prefix+'-linked-shapes-format',
                        label="Formatierung angleichen",
                        image_mso="FormatPainter",
                        on_action=bkt.Callback(LinkedShapes.format_linked_shapes, shape=True, context=True),
                        # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = prefix+'-linked-shapes-text',
                        label="Text angleichen",
                        image_mso="TextBoxInsert",
                        on_action=bkt.Callback(LinkedShapes.text_linked_shapes, shape=True, context=True),
                        # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.MenuSeparator(),
                    bkt.ribbon.Button(
                        id = prefix+'-linked-shapes-tofront',
                        label="In den Vordergrund",
                        image_mso="ObjectBringToFront",
                        on_action=bkt.Callback(LinkedShapes.linked_shapes_tofront, shape=True, context=True),
                        # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = prefix+'-linked-shapes-toback',
                        label="In den Hintergrund",
                        image_mso="ObjectSendToBack",
                        on_action=bkt.Callback(LinkedShapes.linked_shapes_toback, shape=True, context=True),
                        # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.MenuSeparator(),
                    bkt.ribbon.Button(
                        id = prefix+'-linked-shapes-delete',
                        label="Andere löschen",
                        image_mso="HyperlinkRemove",
                        on_action=bkt.Callback(LinkedShapes.delete_linked_shapes, shape=True, context=True),
                        # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = prefix+'-linked-shapes-replace',
                        label="Andere mit diesem ersetzen",
                        image_mso="HyperlinkCreate",
                        on_action=bkt.Callback(LinkedShapes.replace_with_this, shape=True, context=True),
                        # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                ])
            ]
        ),
        bkt.ribbon.Menu(
            id=prefix+'-not-linked-shapes',
            label="Verknüpftes Shape anlegen",
            image_mso='HyperlinkCreate',
            insertBeforeMso='ObjectsGroupMenu',
            get_visible=bkt.Callback(LinkedShapes.not_is_linked_shape, shape=True),
            children=[
                bkt.ribbon.Button(
                    id=prefix+"-not-linked-shapes-search",
                    label="Ähnliche Shapes suchen…",
                    on_action=bkt.Callback(LinkedShapes.find_similar_and_link, shape=True, context=True),
                ),
                bkt.ribbon.Button(
                    id=prefix+"-not-linked-shapes-create",
                    label="Dieses Shape kopieren…",
                    on_action=bkt.Callback(LinkedShapes.copy_to_all, shape=True, context=True),
                ),
            ]
        )
    ]




linkshapes_tab = bkt.ribbon.Tab(
    id = "bkt_context_tab_linkshapes",
    label = "[BKT] Verknüpfte Shapes",
    get_visible=bkt.Callback(LinkedShapes.is_linked_shape, shape=True),
    children = [
        bkt.ribbon.Group(
            label = "Verknüpfte Shapes finden",
            children = [
                bkt.ribbon.Button(
                    id = 'linked_shapes_count',
                    label="Verknüpfte Shapes zählen",
                    image_mso="FormattingUnique",
                    screentip="Alle verknüpften Shapes zählen",
                    supertip="Zählt die Anzahl der verknüpften Shapes auf allen Folien.",
                    on_action=bkt.Callback(LinkedShapes.count_link_shapes, shape=True, context=True),
                ),
                bkt.ribbon.Button(
                    id = 'linked_shapes_next',
                    label="Nächstes verknüpfte Shape finden",
                    image_mso="FindNext",
                    screentip="Zum nächsten verknüpften Shape gehen",
                    supertip="Sucht nach dem nächste verknüpften Shape. Sollte auf den Folgefolien kein Shape mehr kommen, wird das erste verknüpfte Shape der Präsentation gesucht.",
                    on_action=bkt.Callback(LinkedShapes.goto_linked_shape, shape=True, context=True),
                ),
                bkt.ribbon.Button(
                    id = 'linked_shapes_first',
                    label="Erstes verknüpfte Shape finden",
                    image_mso="FirstPage",
                    screentip="Zum ersten verknüpften Shape gehen",
                    supertip="Sucht nach dem ersten verknüpften Shape.",
                    on_action=bkt.Callback(LinkedShapes.goto_first_shape, shape=True, context=True),
                ),
            ]
        ),
        bkt.ribbon.Group(
            label = "Verknüpfte Shapes angleichen",
            children = [
                bkt.ribbon.Button(
                    id = 'linked_shapes_all',
                    label="Alles angleichen",
                    image_mso="GroupUpdate",
                    size="large",
                    screentip="Alle Eigenschaften verknüpfter Shapes angleichen",
                    supertip="Alle Eigenschaften aller verknüpfter Shapes wie ausgewähltes Shape setzen.",
                    on_action=bkt.Callback(LinkedShapes.equalize_linked_shapes, shape=True, context=True),
                ),
                bkt.ribbon.Separator(),
                bkt.ribbon.Button(
                    id = 'linked_shapes_align',
                    label="Position angleichen",
                    image_mso="ControlAlignToGrid",
                    screentip="Position verknüpfter Shapes angleichen",
                    supertip="Position und Rotation aller verknüpfter Shapes auf Position wie ausgewähltes Shape setzen.",
                    on_action=bkt.Callback(LinkedShapes.align_linked_shapes, shape=True, context=True),
                ),
                bkt.ribbon.Button(
                    id = 'linked_shapes_size',
                    label="Größe angleichen",
                    image_mso="SizeToControlHeightAndWidth",
                    screentip="Größe verknüpfter Shapes angleichen",
                    supertip="Größe aller verknüpfter Shapes auf Größe wie ausgewähltes Shape setzen.",
                    on_action=bkt.Callback(LinkedShapes.size_linked_shapes, shape=True, context=True),
                ),
                bkt.ribbon.Button(
                    id = 'linked_shapes_format',
                    label="Formatierung angleichen",
                    image_mso="FormatPainter",
                    screentip="Formatierung verknüpfter Shapes angleichen",
                    supertip="Formatierung aller verknüpfter Shapes auf Größe wie ausgewähltes Shape setzen.",
                    on_action=bkt.Callback(LinkedShapes.format_linked_shapes, shape=True, context=True),
                ),
                bkt.ribbon.Button(
                    id = 'linked_shapes_text',
                    label="Text angleichen",
                    image_mso="TextBoxInsert",
                    screentip="Text verknüpfter Shapes angleichen",
                    supertip="Text aller verknüpfter Shapes auf Größe wie ausgewähltes Shape setzen.",
                    on_action=bkt.Callback(LinkedShapes.text_linked_shapes, shape=True, context=True),
                ),
                bkt.ribbon.Button(
                    id = 'linked_shapes_tofront',
                    label="In den Vordergrund",
                    image_mso="ObjectBringToFront",
                    screentip="Alle verknüpften Shapes in den Vordergrund bringen",
                    # supertip="",
                    on_action=bkt.Callback(LinkedShapes.linked_shapes_tofront, shape=True, context=True),
                ),
                bkt.ribbon.Button(
                    id = 'linked_shapes_toback',
                    label="In den Hintergrund",
                    image_mso="ObjectSendToBack",
                    screentip="Alle verknüpften Shapes in den Hintergrund bringen",
                    # supertip="",
                    on_action=bkt.Callback(LinkedShapes.linked_shapes_toback, shape=True, context=True),
                ),
                bkt.ribbon.Separator(),
                bkt.ribbon.Button(
                    id = 'linked_shapes_delete',
                    label="Andere Shapes löschen",
                    image_mso="HyperlinkRemove",
                    screentip="Verknüpfte Shapes löschen",
                    supertip="Alle verknüpften Shapes auf allen Folien löschen.",
                    on_action=bkt.Callback(LinkedShapes.delete_linked_shapes, shape=True, context=True),
                ),
                bkt.ribbon.Button(
                    id = 'linked_shapes_replace',
                    label="Andere mit diesem ersetzen",
                    image_mso="HyperlinkCreate",
                    screentip="Verknüpfte Shapes ersetzen",
                    supertip="Alle verknüpften Shapes auf allen Folien mit ausgewähltem Shape ersetzen.",
                    on_action=bkt.Callback(LinkedShapes.replace_with_this, shape=True, context=True),
                ),
            ]
        ),
        bkt.ribbon.Group(
            label = "Verknüpfung aufheben",
            children = [
                bkt.ribbon.Button(
                    id = 'linked_shapes_unlink',
                    label="Einzelne Shape-Verknüpfung entfernen",
                    image_mso="HyperlinkRemove",
                    screentip="Verknüpfung des ausgewählten Shapes entfernen",
                    supertip="Entfernt die ID zur Verknüpfung vom aktuellen Shape. Alle anderen Shapes mit der gleichen ID bleiben verknüpft.",
                    on_action=bkt.Callback(LinkedShapes.unlink_shape, shape=True),
                ),
                bkt.ribbon.Button(
                    id = 'linked_shapes_unlink_all',
                    label="Gesamte Shape-Verknüpfung auflösen",
                    image_mso="HyperlinkRemove",
                    screentip="Alle Shape-Verknüpfungen entfernen",
                    supertip="Entfernt die ID zur Verknüpfung vom aktuellen Shape sowie allen verknüpften Shapes mit der gleichen ID.",
                    on_action=bkt.Callback(LinkedShapes.unlink_all_shapes, shape=True, context=True),
                ),
            ]
        ),
    ]
)


class LinkedShapePopup(bkt.ui.WpfWindowAbstract):
    _filename = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'popups', 'linkedshape.xaml')
    '''
    class representing a popup-dialog for a linked shape
    '''
    
    def __init__(self, context=None):
        self.IsPopup = True
        self._context = context

        super(LinkedShapePopup, self).__init__()

    def btntab(self, sender, event):
        try:
            self._context.ribbon.ActivateTab('bkt_context_tab_linkshapes')
        except:
            bkt.helpers.message("Tab-Wechsel aus unbekannten Gründen fehlgeschlagen.")

    def btnsync_text(self, sender, event):
        try:
            LinkedShapes.text_linked_shapes(self._context.shapes[-1], self._context)
        except:
            bkt.helpers.message("Aktualisierung aus unbekannten Gründen fehlgeschlagen.")

    def btnsync_possize(self, sender, event):
        try:
            LinkedShapes.align_linked_shapes(self._context.shapes[-1], self._context)
            LinkedShapes.size_linked_shapes(self._context.shapes[-1], self._context)
        except:
            bkt.helpers.message("Aktualisierung aus unbekannten Gründen fehlgeschlagen.")

    def btnsync_format(self, sender, event):
        try:
            LinkedShapes.format_linked_shapes(self._context.shapes[-1], self._context)
        except:
            bkt.helpers.message("Aktualisierung aus unbekannten Gründen fehlgeschlagen.")

    def btnnext(self, sender, event):
        try:
            LinkedShapes.goto_linked_shape(self._context.shapes[-1], self._context)
        except:
            bkt.helpers.message("Funktion aus unbekannten Gründen fehlgeschlagen.")

# register dialog
bkt.powerpoint.context_dialogs.register_dialog(
    bkt.contextdialogs.ContextDialog(
        id=BKT_LINK_UUID,
        module=None,
        window_class=LinkedShapePopup
    )
)