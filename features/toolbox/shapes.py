# -*- coding: utf-8 -*-
'''
Created on 06.07.2016

@author: rdebeerst
'''



import logging
import locale

from System import Array

import bkt
import bkt.library.powerpoint as pplib
pt_to_cm = pplib.pt_to_cm
cm_to_pt = pplib.cm_to_pt
get_ambiguity_tuple = bkt.helpers.get_ambiguity_tuple

from bkt.library.algorithms import get_bounding_nodes, mid_point

from bkt import dotnet
Drawing = dotnet.import_drawing()
office = dotnet.import_officecore()

# other toolbox modules
from .chartlib import shapelib_button
from .agenda import ToolboxAgenda
from . import text
from . import harvey
from . import stateshapes






class PositionSize(object):
    use_visual_pos  = bkt.settings.get("toolbox.possize.use_visual_pos", False)
    use_visual_size = bkt.settings.get("toolbox.possize.use_visual_size", False)

    @classmethod
    def toggle_use_visual_pos(cls):
        cls.use_visual_pos = not cls.use_visual_pos
        bkt.settings["toolbox.possize.use_visual_pos"] = cls.use_visual_pos

    @classmethod
    def get_image_use_visual_pos(cls):
        return bkt.ribbon.Gallery.get_check_image(cls.use_visual_pos)

    @classmethod
    def toggle_use_visual_size(cls):
        cls.use_visual_size = not cls.use_visual_size
        bkt.settings["toolbox.possize.use_visual_size"] = cls.use_visual_size

    @classmethod
    def get_image_use_visual_size(cls):
        return bkt.ribbon.Gallery.get_check_image(cls.use_visual_size)

    @classmethod
    def set_top(cls, shapes, value):
        attr = 'visual_top' if cls.use_visual_pos else 'top'
        bkt.apply_delta_on_ALT_key(
            lambda shape, value: setattr(shape, attr, value), 
            lambda shape: getattr(shape, attr), 
            shapes, value)
    
    @classmethod
    def get_top(cls, shapes):
        if not cls.use_visual_pos:
            return get_ambiguity_tuple(shape.top for shape in shapes) #shapes[0].top
        else:
            return get_ambiguity_tuple(shape.visual_top for shape in shapes)
    
    
    @classmethod
    def set_left(cls, shapes, value):
        attr = 'visual_left' if cls.use_visual_pos else 'left'
        bkt.apply_delta_on_ALT_key(
            lambda shape, value: setattr(shape, attr, value), 
            lambda shape: getattr(shape, attr), 
            shapes, value)

    @classmethod
    def get_left(cls, shapes):
        if not cls.use_visual_pos:
            return get_ambiguity_tuple(shape.left for shape in shapes) #shapes[0].left
        else:
            return get_ambiguity_tuple(shape.visual_left for shape in shapes)


    @classmethod
    def set_height(cls, shapes, value):
        attr = 'visual_height' if cls.use_visual_size else 'height'
        bkt.apply_delta_on_ALT_key(
            lambda shape, value: setattr(shape, attr, value), 
            lambda shape: getattr(shape, attr), 
            shapes, value)

    @classmethod
    def get_height(cls, shapes):
        if not cls.use_visual_size:
            return get_ambiguity_tuple(shape.height for shape in shapes) #shapes[0].height
        else:
            return get_ambiguity_tuple(shape.visual_height for shape in shapes)
    
    
    @classmethod
    def set_width(cls, shapes, value):
        attr = 'visual_width' if cls.use_visual_size else 'width'
        bkt.apply_delta_on_ALT_key(
            lambda shape, value: setattr(shape, attr, value), 
            lambda shape: getattr(shape, attr), 
            shapes, value)

    @classmethod
    def get_width(cls, shapes):
        if not cls.use_visual_size:
            return get_ambiguity_tuple(shape.width for shape in shapes) #shapes[0].width
        else:
            return get_ambiguity_tuple(shape.visual_width for shape in shapes)


    @staticmethod
    def set_zorder(shapes, value):
        delta = int(value) - shapes[0].ZOrderPosition
        shapes = sorted(shapes, key=lambda shape: shape.ZOrderPosition, reverse=True if delta > 0 else False)
        for shape in shapes:
            pplib.set_shape_zorder(shape, delta=delta)
        # Normal behavior too confusing for users:
        # bkt.apply_delta_on_ALT_key(
        #     PositionSize._set_shape_zorder, 
        #     lambda shape: shape.ZOrderPosition, 
        #     shapes, int(value))

    @staticmethod
    def get_zorder(shapes):
        if len(shapes) == 1:
            return shapes[0].ZOrderPosition
        else:
            return (True, shapes[0].ZOrderPosition) #force ambiguous mode

    @staticmethod
    def front_to_back(shapes):
        shapes = sorted(shapes, key=lambda shape: shape.ZOrderPosition, reverse=True)
        target_zorder = shapes.pop(-1).ZOrderPosition
        for shape in shapes:
            pplib.set_shape_zorder(shape, value=target_zorder)

    @staticmethod
    def back_to_front(shapes):
        shapes = sorted(shapes, key=lambda shape: shape.ZOrderPosition, reverse=False)
        target_zorder = shapes.pop(-1).ZOrderPosition
        for shape in shapes:
            pplib.set_shape_zorder(shape, value=target_zorder)

    @staticmethod
    def zorder_top2bottom(shapes, reverse=False):
        shapes = sorted(shapes, key=lambda shape: shape.Top, reverse=reverse)
        start = shapes[0].ZOrderPosition
        for shape in shapes:
            pplib.set_shape_zorder(shape, value=start)
            start += 1
        #update selection
        pplib.shapes_to_range(shapes).select()

    @classmethod
    def zorder_bottom2top(cls, shapes):
        cls.zorder_top2bottom(shapes, True)

    @staticmethod
    def zorder_left2right(shapes, reverse=False):
        shapes = sorted(shapes, key=lambda shape: shape.Left, reverse=reverse)
        start = shapes[0].ZOrderPosition
        for shape in shapes:
            pplib.set_shape_zorder(shape, value=start)
            start += 1
        #update selection
        pplib.shapes_to_range(shapes).select()

    @classmethod
    def zorder_right2left(cls, shapes):
        cls.zorder_left2right(shapes, True)
    
    @staticmethod
    def set_height_to_width(shapes):
        for shape in shapes:
            shape.square(w2h=False)
    
    @staticmethod
    def set_width_to_height(shapes):
        for shape in shapes:
            shape.square(w2h=True)
    
    @staticmethod
    def swap_width_and_height(shapes):
        for shape in shapes:
            shape.transpose()
    
    @staticmethod
    def set_top_to_left(shapes):
        for shape in shapes:
            shape.top = shape.left
    
    @staticmethod
    def set_left_to_top(shapes):
        for shape in shapes:
            shape.left = shape.top
    
    @staticmethod
    def swap_left_and_top(shapes):
        for shape in shapes:
            shape.top, shape.left = shape.left, shape.top


class AspectRatio(object):
    types_scale = (
        pplib.MsoShapeType["msoPicture"],
        pplib.MsoShapeType["msoLinkedPicture"],
        pplib.MsoShapeType["msoFreeform"],
        pplib.MsoShapeType["msoEmbeddedOLEObject"],
        pplib.MsoShapeType["msoLinkedOLEObject"],
        pplib.MsoShapeType["msoMedia"],
    )
    types_in_db = (
        pplib.MsoShapeType["msoAutoShape"],
        pplib.MsoShapeType["msoCallout"],
    )

    aspect_ratios = [
        (1,1),
        (3,2),
        (4,3),
        (13,9),
        (15,10),
        (16,9),
    ]

    @classmethod
    def reset(cls, shapes):
        for shape in shapes:
            try:
                shape_type = shape.Type
                #TODO: placeholder support
                if shape_type in cls.types_scale:
                    height = shape.Height
                    shape.ScaleHeight(1, True)
                    shape.ScaleWidth(1, True)
                    #reapply ratio (only required if LockAspectRatio=0)
                    ratio = shape.Width/shape.Height
                    shape.Height = height
                    shape.Width = ratio*height
                elif shape_type in cls.types_in_db:
                    try:
                        shape_db = pplib.GlobalShapeDb.get_by_shape(shape)
                        ratio = shape_db["ratio"]
                    except:
                        logging.exception("shape not found in db")
                        ratio = 1
                    # landscape = shape.width > shape.height
                    shape.force_aspect_ratio(ratio)
            except:
                continue
    
    @staticmethod
    def swap(shapes):
        for shape in shapes:
            shape.transpose()
    
    @classmethod
    def set_aspect_ratio(cls, shapes, current_control):
        index = int(current_control["tag"])
        r1,r2 = cls.aspect_ratios[index]
        value = r1/r2
        for shape in shapes:
            # landscape = shape.width > shape.height
            shape.force_aspect_ratio(value)
            shape.lock_aspect_ratio = True
    
    @staticmethod
    def get_aspect_ratio(shape):
        return shape.width/shape.height
    
    @classmethod
    def get_aspect_ratio_label(cls, context):
        try:
            return "Aktuelles Seiteverhältnis: {:.4n}".format(cls.get_aspect_ratio(context.shape))
        except:
            return "Aktuelles Seiteverhältnis: -"

    @staticmethod
    def lock_aspect_ratio(shapes, pressed):
        for shape in shapes:
            shape.LockAspectRatio = -1 if pressed else 0
    
    @staticmethod
    def get_aspect_ratio_locked(shapes):
        return shapes[0].LockAspectRatio == -1



spinner_top = bkt.ribbon.RoundingSpinnerBox(
    id="pos_size_spinner_top",
    image_mso='ObjectNudgeDown',
    label="Position von oben",
    show_label=False,
    screentip="Position von oben",
    supertip="Änderung der Position von oben.\n\nBei gedrückter STRG-Taste Veränderung um 0,1 cm statt 0,2 cm.\n\nBei gedrückter ALT-Taste Veränderung relativ je Shape (wenn mehrere Shapes ausgewählt sind).",
    round_cm=True,
    on_change=bkt.Callback(PositionSize.set_top, shapes=True, wrap_shapes=True),
    get_text=bkt.Callback(PositionSize.get_top, shapes=True, wrap_shapes=True),
    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
    convert="pt_to_cm",
    image_element=pplib.LocpinGallery(image_mso='ObjectNudgeDown', children=[
        bkt.ribbon.Button(
            label="Visuelle Position",
            get_image=bkt.Callback(PositionSize.get_image_use_visual_pos),
            supertip="Visuelle Position unter Berücksichtigung der Rotation verwenden",
            on_action=bkt.Callback(PositionSize.toggle_use_visual_pos)
        ),
        bkt.ribbon.Button(
            label="Oben = Links",
            image="possize_t2l",
            screentip="Oben = Links setzen",
            supertip="Setzt die obere Kante gleich der linken Kante unter Berücksichtigung des Fixpunkts",
            on_action=bkt.Callback(PositionSize.set_top_to_left, shapes=True, wrap_shapes=True)
        ),
        bkt.ribbon.Button(
            label="Oben ⇄ Links",
            image="possize_swap_tl",
            screentip="Oben und Links tauschen",
            supertip="Tauscht die obere Kante mit der linken Kante unter Berücksichtigung des Fixpunkts",
            on_action=bkt.Callback(PositionSize.swap_left_and_top, shapes=True, wrap_shapes=True)
        ),
    ])
)

spinner_left = bkt.ribbon.RoundingSpinnerBox(
    id="pos_size_spinner_left",
    image_mso='ObjectNudgeRight',
    label="Position von links",
    show_label=False,
    screentip="Position von links",
    supertip="Änderung der Position von links.\n\nBei gedrückter STRG-Taste Veränderung um 0,1 cm statt 0,2 cm.\n\nBei gedrückter ALT-Taste Veränderung relativ je Shape (wenn mehrere Shapes ausgewählt sind).",
    round_cm=True,
    on_change=bkt.Callback(PositionSize.set_left, shapes=True, wrap_shapes=True),
    get_text=bkt.Callback(PositionSize.get_left, shapes=True, wrap_shapes=True),
    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
    convert="pt_to_cm",
    image_element=pplib.LocpinGallery(image_mso='ObjectNudgeRight', children=[
        bkt.ribbon.Button(
            label="Visuelle Position",
            get_image=bkt.Callback(PositionSize.get_image_use_visual_pos),
            supertip="Visuelle Position unter Berücksichtigung der Rotation verwenden",
            on_action=bkt.Callback(PositionSize.toggle_use_visual_pos)
        ),
        bkt.ribbon.Button(
            label="Links = Oben",
            image="possize_l2t",
            screentip="Links = Oben setzen",
            supertip="Setzt die linke Kante gleich der oberen Kante unter Berücksichtigung des Fixpunkts",
            on_action=bkt.Callback(PositionSize.set_left_to_top, shapes=True, wrap_shapes=True)
        ),
        bkt.ribbon.Button(
            label="Links ⇄ Oben",
            image="possize_swap_tl",
            screentip="Links und Oben tauschen",
            supertip="Tauscht die linke Kante mit der oberen Kante unter Berücksichtigung des Fixpunkts",
            on_action=bkt.Callback(PositionSize.swap_left_and_top, shapes=True, wrap_shapes=True)
        ),
    ])
)

spinner_height = bkt.ribbon.RoundingSpinnerBox(
    id="pos_size_spinner_height",
    image_mso='ShapeHeight',
    label="Höhe",
    show_label=False,
    screentip="Höhe",
    supertip="Änderung der Höhe.\n\nBei gedrückter STRG-Taste Veränderung um 0,1 cm statt 0,2 cm.\n\nBei gedrückter ALT-Taste Veränderung relativ je Shape (wenn mehrere Shapes ausgewählt sind).",
    round_cm=True,
    on_change=bkt.Callback(PositionSize.set_height, shapes=True, wrap_shapes=True),
    get_text=bkt.Callback(PositionSize.get_height, shapes=True, wrap_shapes=True),
    get_enabled=bkt.apps.ppt_shapes_or_text_selected,
    convert="pt_to_cm",
    image_element=pplib.LocpinGallery(image_mso='ShapeHeight', children=[
        bkt.ribbon.Button(
            label="Visuelle Größe",
            get_image=bkt.Callback(PositionSize.get_image_use_visual_size),
            supertip="Visuelle Größe unter Berücksichtigung der Rotation verwenden",
            on_action=bkt.Callback(PositionSize.toggle_use_visual_size)
        ),
        bkt.ribbon.Button(
            label="Höhe = Breite",
            image="possize_h2w",
            screentip="Höhe = Breite setzen",
            supertip="Setzt die Höhe gleich der Breite unter Berücksichtigung des Fixpunkts. Ist das Seitenverhältnis gesperrt, wird dies temporär aufgehoben.",
            on_action=bkt.Callback(PositionSize.set_height_to_width, shapes=True, wrap_shapes=True)
        ),
        bkt.ribbon.Button(
            label="Höhe ⇄ Breite",
            image="possize_swap_hw",
            screentip="Höhe und Breite tauschen",
            supertip="Tauscht die Höhe mit der Breite unter Berücksichtigung des Fixpunkts",
            on_action=bkt.Callback(PositionSize.swap_width_and_height, shapes=True, wrap_shapes=True)
        ),
    ])
)

spinner_width = bkt.ribbon.RoundingSpinnerBox(
    id="pos_size_spinner_width",
    image_mso='ShapeWidth',
    label="Breite",
    show_label=False,
    screentip="Breite",
    supertip="Änderung der Breite.\n\nBei gedrückter STRG-Taste Veränderung um 0,1 cm statt 0,2 cm.\n\nBei gedrückter ALT-Taste Veränderung relativ je Shape (wenn mehrere Shapes ausgewählt sind).",
    round_cm=True,
    on_change=bkt.Callback(PositionSize.set_width, shapes=True, wrap_shapes=True),
    get_text=bkt.Callback(PositionSize.get_width, shapes=True, wrap_shapes=True),
    get_enabled=bkt.apps.ppt_shapes_or_text_selected,
    convert="pt_to_cm",
    image_element=pplib.LocpinGallery(image_mso='ShapeWidth', children=[
        bkt.ribbon.Button(
            label="Visuelle Größe",
            get_image=bkt.Callback(PositionSize.get_image_use_visual_size),
            supertip="Visuelle Größe unter Berücksichtigung der Rotation verwenden",
            on_action=bkt.Callback(PositionSize.toggle_use_visual_size)
        ),
        bkt.ribbon.Button(
            label="Breite = Höhe",
            image="possize_w2h",
            screentip="Breite = Höhe setzen",
            supertip="Setzt die Breite gleich der Höhe unter Berücksichtigung des Fixpunkts. Ist das Seitenverhältnis gesperrt, wird dies temporär aufgehoben.",
            on_action=bkt.Callback(PositionSize.set_width_to_height, shapes=True, wrap_shapes=True)
        ),
        bkt.ribbon.Button(
            label="Breite ⇄ Höhe",
            image="possize_swap_hw",
            screentip="Breite und Höhe tauschen",
            supertip="Tauscht die Breite mit der Höhe unter Berücksichtigung des Fixpunkts",
            on_action=bkt.Callback(PositionSize.swap_width_and_height, shapes=True, wrap_shapes=True)
        ),
    ])
)

spinner_zorder = bkt.ribbon.RoundingSpinnerBox(
    id="pos_size_spinner_zorder",
    image_mso='ObjectBringForward',
    label="Z-Order",
    show_label=False,
    screentip="Z-Order",
    supertip="Änderung der Z-Order, also der Reihenfolge der Shapes auf der Folie.",
    on_change=bkt.Callback(PositionSize.set_zorder, shapes=True),
    get_text=bkt.Callback(PositionSize.get_zorder, shapes=True),
    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
    round_int=True,
    small_step=1,
    big_step=1,
    image_element=bkt.ribbon.Menu(
        children=[
            bkt.mso.control.ObjectBringToFront,
            bkt.mso.control.ObjectSendToBack,
            bkt.ribbon.MenuSeparator(title="Anpassen"),
            bkt.ribbon.Button(
                label="Vordere nach hinten",
                supertip="Bringt alle vordere Shapes genau hinter das hinterste Shape",
                image="zorder_front_to_back",
                get_enabled=bkt.apps.ppt_shapes_min2_selected,
                on_action=bkt.Callback(PositionSize.front_to_back, shapes=True),
            ),
            bkt.ribbon.Button(
                label="Hintere nach vorne",
                supertip="Bringt alle hinteren Shapes genau vor das vorderste Shape",
                image="zorder_back_to_front",
                get_enabled=bkt.apps.ppt_shapes_min2_selected,
                on_action=bkt.Callback(PositionSize.back_to_front, shapes=True),
            ),
            bkt.ribbon.MenuSeparator(title="Sortieren"),
            bkt.ribbon.Button(
                label="Oben nach unten",
                supertip="Sortiert die Z-Order von oben nach unten, sodass das unterste Shape das vorderste wird",
                image="zorder_top_to_bottom",
                get_enabled=bkt.apps.ppt_shapes_min2_selected,
                on_action=bkt.Callback(PositionSize.zorder_top2bottom, shapes=True),
            ),
            bkt.ribbon.Button(
                label="Unten nach oben",
                supertip="Sortiert die Z-Order von unten nach oben, sodass das oberste Shape das vorderste wird",
                image="zorder_bottom_to_top",
                get_enabled=bkt.apps.ppt_shapes_min2_selected,
                on_action=bkt.Callback(PositionSize.zorder_bottom2top, shapes=True),
            ),
            bkt.ribbon.MenuSeparator(),
            bkt.ribbon.Button(
                label="Links nach rechts",
                supertip="Sortiert die Z-Order von links nach rechts, sodass das rechte Shape das vorderste wird",
                image="zorder_left_to_right",
                get_enabled=bkt.apps.ppt_shapes_min2_selected,
                on_action=bkt.Callback(PositionSize.zorder_left2right, shapes=True),
            ),
            bkt.ribbon.Button(
                label="Rechts nach links",
                supertip="Sortiert die Z-Order von rechts nach links, sodass das linke Shape das vorderste wird",
                image="zorder_right_to_left",
                get_enabled=bkt.apps.ppt_shapes_min2_selected,
                on_action=bkt.Callback(PositionSize.zorder_right2left, shapes=True),
            ),
        ],
    ),
)

#button_lock_aspect_ratio = bkt.ribbon.CheckBox(
button_lock_aspect_ratio = dict(
    #id = 'shape_lock_aspect_ratio',
    # label="Seitenverhält.",
    screentip="Seitenverhältnis sperren",
    supertip="Wenn das Kontrollkästchen Seitenverhältnis sperren aktiviert ist, ändern sich die Einstellungen von Höhe und Breite im Verhältnis zueinander.",
    on_toggle_action = bkt.Callback(AspectRatio.lock_aspect_ratio, shapes=True),
    get_pressed = bkt.Callback(AspectRatio.get_aspect_ratio_locked, shapes=True),
    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
)

menu_lock_aspect_ratio = bkt.ribbon.Box(
    box_style="horizontal",
    children=[
        bkt.ribbon.Menu(
            label="Sperren",
            show_label=False,
            image_mso="AutoSizePage",
            children=[
                bkt.ribbon.ToggleButton(id="shape_lock_aspect_ratio3", label="Seitenverhältnis sperren an/aus", image_mso="Lock", **button_lock_aspect_ratio),
                bkt.ribbon.Button(
                    get_label=bkt.Callback(AspectRatio.get_aspect_ratio_label, context=True),
                    enabled=False,
                ),
                bkt.ribbon.MenuSeparator(),
            ] + [
                bkt.ribbon.Button(
                    label="Setzen auf {}:{} ({:.4n})".format(r[0], r[1], r[0]/r[1]),
                    tag=str(i),
                    on_action=bkt.Callback(AspectRatio.set_aspect_ratio, shapes=True, current_control=True, wrap_shapes=True),
                ) for i, r in enumerate(AspectRatio.aspect_ratios)
            ] + [
                bkt.ribbon.MenuSeparator(),
                bkt.ribbon.Button(
                    label="Tauschen",
                    screentip="Seitenverhältnis tauschen",
                    supertip="Vertauscht Breite und Größe und dreht damit das Seitenverhältnis um.",
                    image_mso="PageScaleToFitOptionsDialog",
                    on_action = bkt.Callback(AspectRatio.swap, shapes=True, wrap_shapes=True),
                ),
                bkt.ribbon.Button(
                    label="Zurücksetzen",
                    screentip="Seitenverhältnis zurücksetzen",
                    supertip="Setzt das Seitenverhältnis auf den Ursprungszustand zurück.",
                    image_mso="ResetCurrentView",
                    on_action = bkt.Callback(AspectRatio.reset, shapes=True, wrap_shapes=True),
                ),
                # bkt.mso.control.PictureResetAndSize,
            ]
        ),
        bkt.ribbon.CheckBox(id="shape_lock_aspect_ratio2", label="Seitenv.", **button_lock_aspect_ratio),
        # bkt.ribbon.ToggleButton(
        #     label="Gesperrt",
        #     # show_label=False,
        #     image_mso="Lock",
        # ),
        # bkt.ribbon.ToggleButton(
        #     label="Offen",
        #     show_label=False,
        #     image_mso="Lock",
        # ),
    ]
)

size_group = bkt.ribbon.Group(
    id="bkt_size_group",
    label='Größe',
    image_mso='GroupSizeAndPosition',
    children =[
        #spinner_height,
        #spinner_width,
        bkt.mso.control.ShapeHeight(show_label=False),
        bkt.mso.control.ShapeWidth(show_label=False),
        bkt.ribbon.CheckBox(id="shape_lock_aspect_ratio1", label="Seitenverhält.", **button_lock_aspect_ratio),
        bkt.ribbon.DialogBoxLauncher(idMso='ObjectSizeAndPositionDialog')
    ]
)

# pos_group = bkt.ribbon.Group(
#     label='Position',
#     image_mso='GroupSizeAndPosition',
#     children =[
#         spinner_top,
#         spinner_left,
#         spinner_zorder,
#         bkt.ribbon.DialogBoxLauncher(idMso='ObjectSizeAndPositionDialog')
#     ]
# )

pos_size_group = bkt.ribbon.Group(
    id="bkt_possize_group",
    label='Position/Größe',
    image_mso='GroupSizeAndPosition',
    children =[
        spinner_height,
        spinner_width,
        menu_lock_aspect_ratio,
        # bkt.ribbon.CheckBox(id="shape_lock_aspect_ratio2", **button_lock_aspect_ratio),
        spinner_top,
        spinner_left,
        spinner_zorder,
        bkt.ribbon.DialogBoxLauncher(idMso='ObjectSizeAndPositionDialog')
    ]
)



class TrackerShape(object):

    @classmethod
    def generateTracker(cls, shapes, context):
        import uuid
        from .linkshapes import LinkedShapes

        #shapes to copy formatting
        shapes_count = len(shapes)
        highlight_shape = shapes[0]
        default_shape = shapes[1]
        slide_width = context.app.ActivePresentation.PageSetup.SlideWidth

        #copy and paste shapes (note: shapes can also be part of a group)
        pplib.shapes_to_range(shapes).copy()
        grp = context.slide.shapes.paste().group()

        # format unselected elements
        for shp in grp.GroupItems:
            if shp.HasTextFrame:
                shp.TextFrame.DeleteText()
            default_shape.PickUp()
            shp.Apply()

        # generate unique GUID fpr tracker (items)
        tracker_guid = str(uuid.uuid4())

        # format each selected element and paste tracker as image
        for i in range(1, shapes_count+1):
            highlight_shape.PickUp()
            new_grp = grp.Duplicate()
            
            new_grp.GroupItems(i).Apply()
            new_grp.Copy()

            tracker = context.slide.shapes.PasteSpecial(DataType=6) # ppPastePNG = 6
            tracker.Tags.Add("tracker_id", tracker_guid)

            new_grp.Delete()

            tracker.Height = cm_to_pt(1.5)
            tracker.left = slide_width - cm_to_pt(3.0) - shapes_count*cm_to_pt(1/1.5) + cm_to_pt(i/1.5)
            tracker.top = cm_to_pt(3.0) + cm_to_pt(i/1.5)

        #delete duplicated shapes
        grp.Delete()
        all_trackers = pplib.last_n_shapes_on_slide(context.slide, shapes_count)
        all_trackers_list = list(iter(all_trackers))

        #select all tracker
        all_trackers.select()

        #make trackers linked shapes
        LinkedShapes.link_shapes(all_trackers_list)

        #ask to distribute trackers
        if bkt.message.confirmation("Tracker auf Folgefolien verteilen?"):
            cls.distributeTracker(all_trackers_list, context)
            all_trackers_list[0].select()

    @staticmethod
    def isTracker(shape):
        return pplib.TagHelper.has_tag(shape, "tracker_id")

    @staticmethod
    def alignTracker(shape, context):
        tracker_id = shape.Tags.Item("tracker_id")
        if not tracker_id:
            return
        
        tracker_position_left = shape.left
        tracker_position_top = shape.top
        tracker_rotation = shape.Rotation
        tracker_heigth = shape.Height
        tracker_width = shape.Width
        tracker_lock_ar = shape.LockAspectRatio
        
        for sld in context.app.ActivePresentation.Slides:
            for cShp in sld.shapes:
                if cShp.Tags.Item("tracker_id") == tracker_id:
                    cShp.LockAspectRatio = 0 #msoFalse
                    cShp.left, cShp.top = tracker_position_left, tracker_position_top
                    cShp.Height, cShp.Width = tracker_heigth, tracker_width
                    cShp.Rotation = tracker_rotation
                    cShp.LockAspectRatio = tracker_lock_ar

    @staticmethod
    def removeTracker(shape, context):
        tracker_id = shape.Tags.Item("tracker_id")
        if not tracker_id:
            return
        
        for sld in context.app.ActivePresentation.Slides:
            for cShp in sld.shapes:
                if cShp.Tags.Item("tracker_id") == tracker_id:
                    cShp.Delete()

    @classmethod
    def distributeTracker(cls, shapes, context):
        cur_slide_index = shapes[0].Parent.SlideIndex
        max_index = context.app.ActivePresentation.Slides.Count
        for shape in shapes[1:]:
            cur_slide_index = min(max_index, cur_slide_index+1)
            shape.Cut()
            context.app.ActivePresentation.Slides[cur_slide_index].Shapes.Paste()

        cls.alignTracker(shapes[0], context)



class ShapeConnectorTags(pplib.BKTTag):
    TAG_NAME = "BKT_SHAPE_CONNECTORS"

class ShapeConnectors(object):
    _default_shape_nodes = dict(top=(0,3), right=(3,2), bottom=(1,2), left=(0,1))
    _special_shape_nodes = {
        pplib.MsoAutoShapeType["msoShapeChevron"]: dict(top=(0,1), right=(1,3), bottom=(3,4), left=(4,0)),
        pplib.MsoAutoShapeType["msoShapePentagon"]: dict(top=(0,1), right=(1,3), bottom=(3,4), left=(4,0)),
        pplib.MsoAutoShapeType["msoShapeHexagon"]: dict(top=(1,2), right=(2,4), bottom=(4,5), left=(5,1)),
        pplib.MsoAutoShapeType["msoShapeOval"]: dict(top=(0,6), right=(3,9), bottom=(6,0), left=(9,3)),
    }

    @staticmethod
    def is_connector(shape):
        return pplib.TagHelper.has_tag(shape, ShapeConnectorTags.TAG_NAME)
        # return shape.Tags.Item(ShapeConnectorTags.TAG_NAME) != '' #FIXME: EnvironmentError for fancy smart-shapes

    @staticmethod
    def _find_shape_by_id(slide, shape_id):
        for shp in slide.shapes:
            if shp.id == shape_id:
                return shp
        else:
            raise IndexError("shape id not found on slide")

    @classmethod
    def _get_shape_connector_nodes(cls, shape, side):
        dummy = None
        try:
            special_nodes = cls._special_shape_nodes[shape.AutoShapeType]
            #convert into freeform by adding and deleting in order to acces points
            dummy = shape.duplicate()
            dummy.left, dummy.top = shape.left, shape.top
            dummy.nodes.insert(1,0,0,0,0)
            dummy.nodes.delete(2)
            shape_nodes = [(node.points[0,0], node.points[0,1]) for node in dummy.nodes]
            shape_p1, shape_p2 = special_nodes[side]
            return shape_nodes[shape_p1], shape_nodes[shape_p2]
        except: #KeyError, or any COM Error
            shape_nodes = get_bounding_nodes(shape)
            shape_p1, shape_p2 = cls._default_shape_nodes[side]
            return shape_nodes[shape_p1], shape_nodes[shape_p2]
        finally:
            if dummy:
                dummy.delete()
    
    @classmethod
    def _set_connector_shape_nodes(cls, shape_connector, shape1, shape2, shape1_side="bottom", shape2_side="top"):
        from math import atan2

        shape1_p1, shape1_p2 = cls._get_shape_connector_nodes(shape1, shape1_side)
        shape2_p1, shape2_p2 = cls._get_shape_connector_nodes(shape2, shape2_side)

        connector_nodes = [shape1_p1, shape1_p2, shape2_p1, shape2_p2]
        #correct ordering is the key to set nodes, here clockwise ordering (left-top, right-top, right-bottom, left-bottom)
        mid_p = mid_point(connector_nodes)
        connector_nodes.sort(key=lambda p: atan2(p[1]-mid_p[1], p[0]-mid_p[0]))

        #convert shape into freeform by adding and deleting node (not sure if this is required)
        shape_connector.Nodes.Insert(1, 0, 0, 0, 0) #msoSegmentLine, msoEditingAuto, x, y
        shape_connector.Nodes.Delete(2)
        # set nodes (rectangle has 5 nodes as start and end node are the same)
        shape_connector.Nodes.SetPosition(1, connector_nodes[0][0], connector_nodes[0][1]) #top-left start node
        shape_connector.Nodes.SetPosition(2, connector_nodes[1][0], connector_nodes[1][1]) #top-right node
        shape_connector.Nodes.SetPosition(3, connector_nodes[2][0], connector_nodes[2][1]) #bottom-right node
        shape_connector.Nodes.SetPosition(4, connector_nodes[3][0], connector_nodes[3][1]) #bottom-left node
        shape_connector.Nodes.SetPosition(5, connector_nodes[0][0], connector_nodes[0][1]) #top-left end node

    @classmethod
    def update_connector_shape(cls, context, shape):
        with ShapeConnectorTags(shape.Tags) as tags:
            slide = context.slide
            try:
                shape1 = cls._find_shape_by_id(slide, tags["shape1_id"])
                shape2 = cls._find_shape_by_id(slide, tags["shape2_id"])
            except IndexError:
                bkt.message.error("Fehler: Verbundenes Shape nicht gefunden!")
            else:
                cls._set_connector_shape_nodes(shape, shape1, shape2, tags["shape1_side"], tags["shape2_side"])

    @classmethod
    def add_connector_shape(cls, slide, shape1, shape2, shape1_side="bottom", shape2_side="top"):
        shp_connector = slide.shapes.AddShape(
            1, #msoShapeRectangle
            1,1, #left-top
            10,10 #width-height
        )

        cls._set_connector_shape_nodes(shp_connector, shape1, shape2, shape1_side, shape2_side)

        # shp_connector.Fill.ForeColor.RGB = 12566463 #193
        shp_connector.Fill.ForeColor.ObjectThemeColor = 16 #Background 2
        # shp_connector.Line.ForeColor.RGB = 8355711 # 127 127 127
        shp_connector.Line.ForeColor.ObjectThemeColor = 15 #Text 2
        # shp_connector.Line.Weight = 0.75
        shp_connector.Line.Visible = -1 #msoTrue

        shp_connector.Name = "[BKT] Connector %s" % shp_connector.id

        with ShapeConnectorTags(shp_connector.Tags) as tags:
            tags["shape1_id"]   = shape1.id
            tags["shape1_side"] = shape1_side
            tags["shape2_id"]   = shape2.id
            tags["shape2_side"] = shape2_side

        return shp_connector


    @classmethod
    def addHorizontalConnector(cls, shapes, context):
        shapes = sorted(shapes, key=lambda shape: shape.Left)

        cls.add_connector_shape(context.slide, shapes[0], shapes[1], "right", "left").select()

        # shpLeft  = shapes[0]
        # shpRight = shapes[1]

        # shpConnector = context.app.ActivePresentation.Slides(context.app.ActiveWindow.View.Slide.SlideIndex).shapes.AddShape(
        #     1, #msoShapeRectangle
        #     shpLeft.Left + shpLeft.Width, shpLeft.Top,
        #     shpRight.Left - shpLeft.Left - shpLeft.width, shpLeft.Height)
        # # node 2: top right
        # shpConnector.Nodes.SetPosition(2, shpRight.Left, shpRight.Top)
        # # node 3: bottom right
        # shpConnector.Nodes.SetPosition(3, shpRight.Left, shpRight.Top + shpRight.Height)
        # shpConnector.Fill.ForeColor.RGB = 12566463 #193
        # shpConnector.Line.ForeColor.RGB = 8355711 # 127 127 127
        # shpConnector.Line.Weight = 0.75
        # shpConnector.Select()

    @classmethod
    def addVerticalConnector(cls, shapes, context):
        shapes = sorted(shapes, key=lambda shape: shape.Top)

        cls.add_connector_shape(context.slide, shapes[0], shapes[1], "bottom", "top").select()

        # shpTop = shapes[0]
        # shpBottom = shapes[1]

        # shpConnector = context.app.ActivePresentation.Slides(context.app.ActiveWindow.View.Slide.SlideIndex).shapes.AddShape(
        #     1, #msoShapeRectangle,
        #     shpTop.Left, shpTop.Top + shpTop.Height,
        #     shpTop.Width, shpBottom.Top - shpTop.Top - shpTop.Height)

        # # node 3: bottom right
        # shpConnector.Nodes.SetPosition(3, shpBottom.Left + shpBottom.width, shpBottom.Top)
        # # node 4: bottom left
        # shpConnector.Nodes.SetPosition(4, shpBottom.Left, shpBottom.Top)
        # shpConnector.Fill.ForeColor.RGB = 12566463 # 193
        # shpConnector.Line.ForeColor.RGB = 8355711 # 127 127 127
        # shpConnector.Line.Weight = 0.75
        # shpConnector.Select()


class VisibilityToggleTags(pplib.BKTTag):
    TAG_NAME = "BKT_VISIBILITY_TOGGLE"

class ShapesMore(object):

    @staticmethod
    def show_invisible_shapes(context):
        toggle_shapes = list()
        slide = context.slide
        context.selection.Unselect()
        for shape in slide.shapes:
            if not shape.visible and not pplib.TagHelper.has_tag(shape, "THINKCELLSHAPEDONOTDELETE"):
                shape.visible = True
                shape.Select(replace=False)
                toggle_shapes.append(shape.id)
        return toggle_shapes

    @staticmethod
    def _hide_selected_shapes(context, shapes):
        toggle_shapes = list()
        for shape in shapes:
            shape.visible = False
            toggle_shapes.append(shape.id)
        return toggle_shapes
    
    @staticmethod
    def _hide_saved_shapes(context, shape_ids):
        toggle_shapes = list()
        slide = context.slide
        for shape in slide.shapes:
            if shape.id in shape_ids:
                shape.visible = False
                toggle_shapes.append(shape.id)
        return toggle_shapes

    @classmethod
    def toggle_shapes_visibility(cls, context):
        with VisibilityToggleTags(context.slide.Tags) as tags:

            sel_shapes = context.shapes
            if not sel_shapes:
                shapes_shown = cls.show_invisible_shapes(context)
                if shapes_shown:
                    tags["shape_ids"] = shapes_shown
                else:
                    try:
                        shape_ids = tags["shape_ids"]
                        cls._hide_saved_shapes(context, shape_ids)
                        del tags["shape_ids"]
                    except KeyError:
                        bkt.message.warning("Es sind keine Shapes zum Verstecken ausgewählt, es wurden keine vormals versteckten Shapes gefunden, und es gibt keine versteckten Shapes!")
            
            else:
                cls._hide_selected_shapes(context, sel_shapes)
                try:
                    del tags["shape_ids"]
                except KeyError:
                    pass


    # @staticmethod
    # def hide_shapes(shapes):
    #     for shape in shapes:
    #         shape.visible = False

    # @staticmethod
    # def show_shapes(slide):
    #     slide.Application.ActiveWindow.Selection.Unselect()
    #     for shape in slide.shapes:
    #         if not shape.visible and not pplib.TagHelper.has_tag(shape, "THINKCELLSHAPEDONOTDELETE"):
    #             shape.visible = True
    #             shape.Select(replace=False)
    
    @staticmethod
    def _text_to_shape(shape):
        try:
            return pplib.convert_text_into_shape(shape)
        except:
            logging.exception("Text to shape failed")
    
    @classmethod
    def texts_to_shapes(cls, shapes):
        if pplib.shape_is_group_child(shapes[0]) or any(shape.type == pplib.MsoShapeType["msoGroup"] for shape in shapes):
            bkt.message.error("PowerPoint unterstützt diese Funktion leider nicht für Gruppen.")
            return
        all_shapes = []
        for shape in shapes:
            all_shapes.append( cls._text_to_shape(shape) )
        if len(all_shapes)>0:
            pplib.shapes_to_range(all_shapes).select()





class ShapeDialogs(object):
    
    ### DIALOG WINDOWS ###

    @staticmethod
    def shape_split(context, shapes):
        from .dialogs.shape_split import ShapeSplitWindow
        ShapeSplitWindow.create_and_show_dialog(context, shapes)

    @staticmethod
    def shape_scale(context, shapes):
        from .dialogs.shape_scale import ShapeScaleWindow
        ShapeScaleWindow.create_and_show_dialog(context, shapes)
    
    @staticmethod
    def show_segmented_circle_dialog(context, slide):
        from .dialogs.circular_segments import SegmentedCircleWindow
        SegmentedCircleWindow.create_and_show_dialog(context, slide)

    @staticmethod
    def show_process_chevrons_dialog(context, slide):
        from .dialogs.shape_process import ProcessWindow
        ProcessWindow.create_and_show_dialog(context, slide)

    ### DIRECT CREATE ###

    @staticmethod
    def create_headered_pentagon(slide):
        from .models.processshapes import Pentagon
        Pentagon.create_headered_pentagon(slide)

    @staticmethod
    def create_headered_chevron(slide):
        from .models.processshapes import Pentagon
        Pentagon.create_headered_chevron(slide)
    
    @staticmethod
    def create_traffic_light(slide, style):
        from .popups.traffic_light import Ampel
        Ampel.create(slide, style)




class NumberedShapes(object):
    
    label = "1"                 # 1->1,2,3   a->a,b,c   A->A,B,C   I->I,II,III
    shape_type = "square"       # square, circle
    style = "dark"              # dark, light
    position = "top-left"       # top-left, top-right
    position_offset = True      # True, False
    
    # label_1 = [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26]
    # label_a = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z']
    # label_A = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
    # label_I = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI', 'XII', 'XIII', 'XIV', 'XV', 'XVI', 'XVII', 'XVIII', 'XIX', 'XX', 'XXI', 'XXII', 'XXIII', 'XXIV', 'XXV', 'XXVI']
    
    _count_formatter = None
    @classmethod
    def get_count_formatter(cls):
        if not cls._count_formatter:
            from formatter import AbstractFormatter, DumbWriter
            cls._count_formatter = AbstractFormatter(DumbWriter())
        return cls._count_formatter

    
    @classmethod
    def create_numbers_for_shapes(cls, slide, shapes, **kwargs):
        
        settings = {
            # default settings
            'label': cls.label,
            'shape_type': cls.shape_type,
            'style': cls.style,
            'position': cls.position,
            'position_offset': cls.position_offset
        }
        # default settings are overwritten by key-word-arguments
        settings.update(kwargs)
        
        len_shapes = 0
        for number, shape in enumerate(shapes, start=1):
            cls.create_number_shape(slide, shape, number, **settings)
            len_shapes = number
        
        pplib.last_n_shapes_on_slide(slide, len_shapes).select()
        
        
    
    @classmethod
    def create_number_shape(cls, slide, shape, number, label='1', shape_type='square', style='dark', position='top-left', position_offset=True):
        
        if shape_type == 'square':
            numshape = slide.shapes.addshape( pplib.MsoAutoShapeType['msoShapeRectangle'] , shape.left, shape.top, 14, 14)
        elif shape_type == 'diamond':
            numshape = slide.shapes.addshape( pplib.MsoAutoShapeType['msoShapeDiamond'] , shape.left, shape.top, 14, 14)
        else: #circle
            numshape = slide.shapes.addshape( pplib.MsoAutoShapeType['msoShapeOval'] , shape.left, shape.top, 14, 14)
        
        numshape.LockAspectRatio = -1

        if style == "dark":
            col_background = 13 #msoThemeColorText1
            col_foreground = 14 #msoThemeColorBackground1
        else:
            col_background = 14 #msoThemeColorBackground1
            col_foreground = 13 #msoThemeColorText1

        numshape.Line.Visible = -1
        numshape.Line.ForeColor.ObjectThemeColor = col_foreground
        numshape.Fill.Visible = -1
        numshape.Fill.ForeColor.ObjectThemeColor = col_background

        # if style == "dark":
        #     numshape.line.visible = False
        #     numshape.fill.ForeColor.RGB = 0
        #     numshape.TextFrame.TextRange.Font.Color.rgb = 255 + 255 * 256 + 255 * 256**2
            
        # else: # light
        #     numshape.line.style = 1
        #     numshape.line.weight = 1
        #     numshape.line.ForeColor.RGB = 0
        #     numshape.fill.ForeColor.RGB = 255 + 255 * 256 + 255 * 256**2
        #     numshape.TextFrame.TextRange.Font.Color.rgb = 0
        
        # positions corrections for rounded rectangles and pentagon/chevron-shapes
        pos_correction_l = 0
        pos_correction_r = 0
        if shape.AutoShapeType == pplib.MsoAutoShapeType['msoShapeRoundedRectangle']:
            pos_correction_l = shape.Adjustments.item[1] * min(shape.Height, shape.Width)
            pos_correction_r = pos_correction_r
        if shape.AutoShapeType in [pplib.MsoAutoShapeType['msoShapeChevron'], pplib.MsoAutoShapeType['msoShapePentagon']]:
            pos_correction_r = shape.Adjustments.item[1] * min(shape.Height, shape.Width)
        
        # set position
        if position == "top-right":
            numshape.left = shape.left+shape.width-numshape.width -pos_correction_r
            if position_offset:
                numshape.left += numshape.width/2
                numshape.top -= numshape.height/2
        else: # top-left
            numshape.left += pos_correction_l
            if position_offset:
                numshape.left -= numshape.width/2
                numshape.top -= numshape.height/2
        
        # format shape and text
        # numshape.TextFrame.TextRange.text = getattr(cls, 'label_' + label)[(number-1)%26] #at number 26 start from beginning to avoid IndexError
        textframe = numshape.TextFrame2
        textframe.TextRange.Text = cls.get_count_formatter().format_counter(label, number)
        textframe.TextRange.Font.Size = 12
        textframe.TextRange.Font.Fill.ForeColor.ObjectThemeColor = col_foreground
        textframe.TextRange.ParagraphFormat.Alignment = pplib.PowerPoint.PpParagraphAlignment.ppAlignCenter.value__
        textframe.TextRange.ParagraphFormat.Bullet.Type = 0
        textframe.AutoSize = 0
        textframe.WordWrap = False
        textframe.MarginTop = 0
        textframe.MarginLeft = 0
        textframe.MarginRight = 0
        textframe.MarginBottom = 0
        #textframe.HorizontalAnchor = office.MsoHorizontalAnchor.msoAnchorCenter.value__
        textframe.VerticalAnchor = office.MsoVerticalAnchor.msoAnchorMiddle.value__
        
        return numshape
    
    
    
class NumberShapesGallery(bkt.ribbon.Gallery):
    
    # item-settings for gallery
    items = [ dict(label=l, style=s, shape_type=t) for l in ['1', 'a', 'A', 'i', 'I'] for t in ['circle', 'square', 'diamond'] for s in ['dark', 'light'] ]
    item_cols = 6
    
    position = "top-left"
    position_offset = True
    
    def __init__(self, **kwargs):
        parent_id = kwargs.get('id') or ""
        my_kwargs = dict(
            label = 'Nummerierung',
            columns = self.item_cols,
            screentip="Nummerierungs-Shapes einfügen",
            supertip="Fügt für jedes markierte Shape ein Nummerierungs-Shape ein. Nummerierung und Styling entsprechend der Auswahl. Markierte Shapes werden entsprechend der Selektions-Reihenfolge durchnummeriert.",
            get_image=bkt.Callback(lambda: self.get_item_image(0)),
            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
            item_width=24,
            item_height=24,
            children=[
                bkt.ribbon.Button(id=parent_id + "_pos_left", label="Position links oben", screentip="Nummerierungs-Shapes links-oben",    on_action=bkt.Callback(self.set_pos_top_left), get_image=bkt.Callback(lambda: self.get_toggle_image('pos-top-left')),
                    supertip="Nummerierungs-Shapes links oben auf dem zugehörigen Shape platzieren"),
                bkt.ribbon.Button(id=parent_id + "_pos-right", label="Position rechts oben", screentip="Nummerierungs-Shapes rechts-oben", on_action=bkt.Callback(self.set_pos_top_right), get_image=bkt.Callback(lambda: self.get_toggle_image('pos-top-right')),
                    supertip="Nummerierungs-Shapes rechts oben auf dem zugehörigen Shape platzieren"),
                bkt.ribbon.Button(id=parent_id + "_pos-offset", label="Versetzt positionieren", screentip="Nummerierungs-Shapes versetzt positionieren", on_action=bkt.Callback(self.toggle_pos_offset), get_image=bkt.Callback(lambda: self.get_toggle_image('pos-offset')),
                    supertip="Standardmäßig werden Nummerierungs-Shapes genau am Rand des zugehörigen Shapes ausgerichtet.\n\nIst \"Versetzt positionieren\" aktiviert, werden die Nummerierungs-Shapes etwas weiter außerhalb des zugehörigen Shapes plaziert, so dass der Mittelpunkt des Nummerierungs-Shapes auf der Ecke liegt.")
            ],
        )
        my_kwargs.update(kwargs)

        super(NumberShapesGallery, self).__init__(**my_kwargs)
    
    
    def on_action_indexed(self, selected_item, index, slide, shapes):
        ''' create numberd shape according of settings in clicked element '''
        item = self.items[index]
        NumberedShapes.create_numbers_for_shapes(slide, shapes, label=item['label'], shape_type=item['shape_type'], style=item['style'], position=self.position, position_offset=self.position_offset)

                
    def get_item_count(self):
        return len(self.items)
    
    # def get_item_label(self, index):
    #     item = self.items[index]
    #     return "%s" % getattr(NumberedShapes, 'label_' + item['label'])[index%self.columns]
    
    def get_item_screentip(self, index):
        return "Nummerierungs-Shapes einfügen"
        
    def get_item_supertip(self, index):
        return "Fügt für jedes markierte Shape ein Nummerierungs-Shape ein. Nummerierung und Styling entsprechend der Auswahl. Markierte Shapes werden entsprechend der Selektions-Reihenfolge durchnummeriert."
    
    def get_item_image(self, index):
        ''' creates an item image with numberd shape according to settings in the specified item '''
        # retrieve item-settings
        item = self.items[index]
        
        # create bitmap, define pen/brush
        size = 48
        img = Drawing.Bitmap(size, size)
        g = Drawing.Graphics.FromImage(img)
        
        #Draw smooth rectangle/ellipse
        g.SmoothingMode = Drawing.Drawing2D.SmoothingMode.AntiAlias

        if item['style'] == "dark":
            pen_border = Drawing.Pen(Drawing.Color.White,2)
            brush_fill = Drawing.Brushes.Black
            text_brush = Drawing.Brushes.White
        else:
            pen_border = Drawing.Pen(Drawing.Color.Black,2)
            brush_fill = Drawing.Brushes.White
            text_brush = Drawing.Brushes.Black

        if item['shape_type'] == 'circle':
            g.FillEllipse(brush_fill, 2, 2, size-4, size-4) #left, top, width, height
            g.DrawEllipse(pen_border, 2, 2, size-4, size-4) #left, top, width, height
        elif item['shape_type'] == 'diamond':
            diamond_points = [(0,1),(1,2),(2,1),(1,0)]
            size_factor = size/2
            points = Array[Drawing.Point]([Drawing.Point(round(l*size_factor),round(t*size_factor)) for t,l in diamond_points])
            g.FillPolygon(brush_fill, points)
            g.DrawPolygon(pen_border, points)
        else: #fallback shape=1 rectangle
            g.FillRectangle(brush_fill, 2, 2, size-4, size-4) #left, top, width, height
            g.DrawRectangle(pen_border, 2, 2, size-4, size-4) #left, top, width, height
        
        # color_black = Drawing.Color.Black
        # if item['style'] == 'dark':
        #     # create black circle/rectangle
        #     brush = Drawing.SolidBrush(color_black)
        #     text_brush = Drawing.Brushes.White

        #     if item['shape_type'] == 'circle':
        #         g.FillEllipse(brush, 2,2, size-5, size-5)
        #     else: #square
        #         g.FillRectangle(brush, Drawing.Rectangle(2,2, size-5, size-5))

        # else: # light
        #     # create white circle/rectangle
        #     text_brush = Drawing.Brushes.Black
        #     pen = Drawing.Pen(color_black,2)

        #     if item['shape_type'] == 'circle':
        #         g.DrawEllipse(pen, 2,1, size-4, size-4)
        #     else: #square
        #         g.DrawRectangle(pen, Drawing.Rectangle(2,2, size-4, size-4))

        # set string format
        strFormat = Drawing.StringFormat()
        strFormat.Alignment = Drawing.StringAlignment.Center
        strFormat.LineAlignment = Drawing.StringAlignment.Center
        
        # draw string
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAliasGridFit
        # g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAlias
        # g.DrawString(str(getattr(NumberedShapes, 'label_' + item['label'])[index%int(self.columns)]),
        g.DrawString(NumberedShapes.get_count_formatter().format_counter(item['label'], index%int(self.item_cols)+1),
                     Drawing.Font("Arial", 32, Drawing.FontStyle.Bold, Drawing.GraphicsUnit.Pixel), text_brush, 
                     # Drawing.Font("Arial", 7, Drawing.FontStyle.Bold), text_brush, 
                     Drawing.RectangleF(1, 2, size, size-1), 
                     strFormat)
        
        return img
    
    def set_pos_top_left(self):
        self.position = 'top-left'
    
    def set_pos_top_right(self):
        self.position = 'top-right'
    
    def toggle_pos_offset(self):
        self.position_offset = not self.position_offset
        
    
    def get_toggle_image(self, key):
        if key == 'pos-top-left':
            pressed = (self.position == 'top-left')
        elif key == 'pos-top-right':
            pressed = (self.position == 'top-right')
        elif key == 'pos-offset':
            pressed = self.position_offset
        else:
            pressed = False

        if pressed:
            return self.get_check_image()
        else:
            return self.get_check_image(checked=False)



class SplitShapes(object):
    default_row_sep = cm_to_pt(0.2)
    default_col_sep = cm_to_pt(0.2)
    default_rows = 6
    default_cols = 6
    
    @classmethod
    def split_shapes(cls, shapes, rows, cols, row_sep, col_sep):
        for shape in shapes:
            cls.split_shape(shape, rows, cols, row_sep, col_sep)
    
    @classmethod
    def split_shape(cls, shape, rows, cols, row_sep, col_sep):
        shape_width = (shape.width - (cols-1)*col_sep)/cols
        shape_height = (shape.height - (rows-1)*row_sep)/rows
        
        #shape.width = shape_width
        #shape.height = shape_height
        
        for row_idx in range(rows):
            for col_idx in range(cols):
                if row_idx == 0 and col_idx == 0:
                    shape_copy = shape
                else:
                    shape_copy = shape.duplicate()
                shape_copy.left = shape.left + col_idx*(shape_width+col_sep)
                shape_copy.top = shape.top + row_idx*(shape_height+row_sep)
                shape_copy.width = shape_width
                shape_copy.height = shape_height
                shape_copy.select(False)
        #shape.Delete()



#FIXME: no dependency to circular wanted here
#from circular import CircularArrangement


class MultiplyShapes(object):
    
    @classmethod
    def multiply_shapes(cls, shapes, rows, cols, row_sep, col_sep):
        for shape in shapes:
            cls.multiply_shape(shape, rows, cols, row_sep, col_sep)
    
    @classmethod
    def multiply_shape(cls, shape, rows, cols, row_sep, col_sep):
        shape_width = shape.width
        shape_height = shape.height
        
        for row_idx in range(rows):
            for col_idx in range(cols):
                if row_idx == 0 and col_idx == 0:
                    continue
                shape_copy = shape.duplicate()
                shape_copy.left = shape.left + col_idx*(shape_width+col_sep)
                shape_copy.top = shape.top + row_idx*(shape_height+row_sep)
                shape_copy.width = shape_width
                shape_copy.height = shape_height
                shape_copy.select(False)
    
    # @classmethod
    # def multiply_shapes_cyclic(cls, shapes, number, sep):
    #     for shape in shapes:
    #         cls.multiply_shaps_cyclic(shape, number, sep)
    #
    # @classmethod
    # def multiply_shaps_cyclic(cls, shape, number, sep):
    #     shapes = [shape]
    #
    #     # distance circle-midpoint to shape-midpoint is given by sep
    #     height = 2*(max(shape.height, shape.width)/2 + sep)
    #     width = height
    #
    #     midpoint = [ shape.left + shape.width/2 , shape.top + shape.height/2 + height/2 ]
    #
    #     # generate shapes and arrange
    #     for row_idx in range(number-1):
    #         shape_copy = shape.duplicate()
    #         shape_copy.select(False)
    #         shapes.append(shape_copy)
    #     CircularArrangement.arrange_circular_wargs(shapes, midpoint, width, height)
        


split_shapes_group = bkt.ribbon.Group(
    id="bkt_splitshapes_group",
    label="Teilen/Vervielfachen",
    image_mso='TableRowsDistribute',
    children=[
        #bkt.ribbon.Label(label="Zeilen"),
        bkt.ribbon.Box(
            box_style="horizontal",
            children = [
                bkt.ribbon.Button(
                    id='shape_split_horizontal',
                    label="Horizontal teilen",
                    show_label=False,
                    image="split_horizontal",
                    screentip="Horizontal teilen",
                    supertip="Shape horizontal in mehrere Shapes teilen, entsprechend der angegebenen Anzahl und mit angegebenem Abstand zwischen den Shapes.",
                    on_action = bkt.Callback(lambda shapes: SplitShapes.split_shapes(shapes, SplitShapes.default_rows, 1, SplitShapes.default_row_sep, 0 )),
                    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                ),
                bkt.ribbon.Button(
                    id='shape_split_vertical',
                    label="Vertikal teilen",
                    show_label=False,
                    image="split_vertical",
                    screentip="Vertikal teilen",
                    supertip="Shape vertikal in mehrere Shapes teilen, entsprechend der angegebenen Anzahl und mit angegebenem Abstand zwischen den Shapes.",
                    on_action = bkt.Callback(lambda shapes: SplitShapes.split_shapes(shapes, 1, SplitShapes.default_cols, 0, SplitShapes.default_col_sep )),
                    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                ),
                bkt.ribbon.Button(
                    id='shape_mult_vertical',
                    label="Vertikal vervielfachen",
                    show_label=False,
                    image="multiply_vertical",
                    screentip="Vertikal vervielfachen",
                    supertip="Shape mehrfach dublizieren, entsprechend der angegebenen Anzahl. Shapes werden untereinander angeordnet mit dem angegebenem Abstand zwischen den Shapes.",
                    on_action = bkt.Callback(lambda shapes: MultiplyShapes.multiply_shapes(shapes, SplitShapes.default_rows, 1, SplitShapes.default_row_sep, 0 )),
                    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                ),
                bkt.ribbon.Button(
                    id='shape_mult_horizontal',
                    label="Horizontal vervielfachen",
                    show_label=False,
                    image="multiply_horizontal",
                    screentip="Horizontal vervielfachen",
                    supertip="Shape mehrfach dublizieren, entsprechend der angegebenen Anzahl. Shapes werden nebeneinander angeordnet mit dem angegebenem Abstand zwischen den Shapes.",
                    on_action = bkt.Callback(lambda shapes: MultiplyShapes.multiply_shapes(shapes, 1, SplitShapes.default_cols, 0, SplitShapes.default_col_sep )),
                    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                ),
            ]
        ),
        bkt.ribbon.RoundingSpinnerBox(
            id = 'shape_slit_rows',
            label="Anzahl Zeilen/Spalten",
            supertip="Angestrebte Shapeanzahl für das Teilen/Vervielfachen von Shapes.",
            show_label=False,
            imageMso="TableRowsDistribute",
            #TableRowsDistribute, TableStyleBandedRowsWord, TableRowsSelect
            on_change = bkt.Callback(lambda value: [setattr(SplitShapes, 'default_rows', max(0, int(value))), setattr(SplitShapes, 'default_cols', max(0, int(value)))]),
            get_text  = bkt.Callback(lambda: SplitShapes.default_rows),
            big_step = 1,
            small_step = 1,
            round_at = 0
        ),
        bkt.ribbon.RoundingSpinnerBox(
            id = 'shape_slit_row_sep',
            label="Zeilen-/Spaltenabstand",
            supertip="Abstand zwischen Shapes zur Berücksichtigung beim Teilen/Vervielfachen von Shapes.\n\nBei Kreisanordnung wird hiermit der vertikale/horizontale Abstand zum Mittelpunkt angegeben.",
            show_label=False,
            image_mso="RowHeight",
            on_change = bkt.Callback(lambda value: [setattr(SplitShapes, 'default_row_sep', cm_to_pt(value)), setattr(SplitShapes, 'default_col_sep', cm_to_pt(value))]),
            get_text  = bkt.Callback(lambda: round(pt_to_cm(SplitShapes.default_row_sep),2)),
            round_cm = True
        ),
        # bkt.ribbon.Button(
        #     id='shape_mult_circular',
        #     #label="mult.",
        #     image="multiply_circular",
        #     screentip="Shape kreisförmig vervielfachen",
        #     supertip="Shape vervielfachen und kreisförmig anordnen.",
        #     on_action = bkt.Callback(lambda shapes: MultiplyShapes.multiply_shapes_cyclic(shapes, SplitShapes.default_rows, SplitShapes.default_row_sep )),
        #     get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
        # )
    ]
)




# Context menu if multiple connectors are selected
class CtxVerbinder(object):
    @staticmethod
    def ctx_connectors_reroute_enabled(shapes):
        return all(shape.Connector == -1 and shape.ConnectorFormat.BeginConnected == -1 and shape.ConnectorFormat.EndConnected == -1 for shape in shapes)

    @staticmethod
    def ctx_connectors_visible(shapes):
        return all(shape.Connector == -1 for shape in shapes)

    @staticmethod
    def set_connector_type(shapes, con_type):
        for shape in shapes:
            if shape.Connector == -1: #msoTrue
                shape.ConnectorFormat.Type = con_type

    @staticmethod
    def reroute_connectors(shapes):
        for shape in shapes:
            if shape.Connector == -1 and shape.ConnectorFormat.BeginConnected == -1 and shape.ConnectorFormat.EndConnected == -1: #msoTrue
                shape.RerouteConnections()

    @staticmethod
    def invert_direction(shapes):
        for shape in shapes:
            if shape.Connector == -1: #msoTrue
                #Swap begin and end styles
                shape.Line.BeginArrowheadLength, shape.Line.EndArrowheadLength = shape.Line.EndArrowheadLength, shape.Line.BeginArrowheadLength
                shape.Line.BeginArrowheadStyle, shape.Line.EndArrowheadStyle = shape.Line.EndArrowheadStyle, shape.Line.BeginArrowheadStyle
                shape.Line.BeginArrowheadWidth, shape.Line.EndArrowheadWidth = shape.Line.EndArrowheadWidth, shape.Line.BeginArrowheadWidth
    



class ShapeFormats(object):
    transparencies = list(range(0, 110, 10))

    @classmethod
    def _attr_setter(cls, shape, value, shp_object, attribute):
        try:
            if attribute == "Transparency":
                value = min(max(0, value/100),100)
            else:
                value = max(0, value)
            shp_object = getattr(shape, shp_object)
            setattr(shp_object, "visible", -1)
            setattr(shp_object, attribute, value)
        except:
            logging.exception("Setting %s attribute %s to value %s failed!", shp_object, attribute, value)
    @classmethod
    def _attr_getter(cls, shape, shp_object, attribute):
        try:
            shp_object = getattr(shape, shp_object)
            value = max(0, getattr(shp_object, attribute))
            if attribute == "Transparency":
                value = value*100
            return value
        except:
            logging.exception("Getting %s attribute %s failed!", shp_object, attribute)
            return 0

    ### Fill properties ###
    @classmethod
    def get_fill_enabled(cls, context):
        #TESTME: is fill implemented for all shape types? (see also problem with line)
        # shape = next(pplib.iterate_shape_subshapes(shapes))
        # return shape.Fill.visible == -1
        
        # copy enabled status of fill-button
        return context.app.commandbars.GetEnabledMso("ShapeFillColorPicker")

    @classmethod
    def get_fill_transparency(cls, shapes):
        shapes = pplib.iterate_shape_subshapes(shapes)
        for shape in shapes:
            try:
                return max(0, round(shape.fill.transparency*100))
            except:
                continue
        return None
    
    @classmethod
    def set_fill_transparency(cls, shapes, value):
        value = min(max(0, value),100) #min=0, max=100
        shapes = list(pplib.iterate_shape_subshapes(shapes))
        bkt.apply_delta_on_ALT_key(
            # lambda shape, value: setattr(shape.Fill, 'Transparency', min(max(0, value/100),100)), 
            cls._attr_setter,
            cls._attr_getter,
            shapes, value, shp_object="Fill", attribute="Transparency")

    ### Line properties ###
    @classmethod
    def get_line_enabled(cls, context):
        # return len(cls._line_filter(shapes)) > 0
        # shape = next(pplib.iterate_shape_subshapes(shapes))
        # try:
        #     return hasattr(shape.line, "visible")
        # except ValueError:
        #     return False

        # copy enabled status of line-button
        return context.app.commandbars.GetEnabledMso("ShapeOutlineColorPicker")

    @classmethod
    def get_line_transparency(cls, shapes):
        shapes = pplib.iterate_shape_subshapes(shapes, exclude=[pplib.MsoShapeType['msoTable']])
        #IMPORTANT: if tables are not excluded, Powerpoint will crash if a table is selected and this function is executed
        for shape in shapes:
            try:
                return max(0, round(shape.line.transparency*100))
            except:
                continue
        return None
    
    @classmethod
    def set_line_transparency(cls, shapes, value):
        value = min(max(0, value),100) #min=0, max=100
        shapes = pplib.iterate_shape_subshapes(shapes)
        bkt.apply_delta_on_ALT_key(
            # lambda shape, value: setattr(shape.Line, 'Transparency', min(max(0, value/100),100)), 
            cls._attr_setter,
            cls._attr_getter,
            shapes, value, shp_object="Line", attribute="Transparency")

    @classmethod
    def get_line_weight(cls, shapes):
        shapes = pplib.iterate_shape_subshapes(shapes, exclude=[pplib.MsoShapeType['msoTable']])
        #IMPORTANT: if tables are not excluded, Powerpoint will crash if a table is selected and this function is executed
        for shape in shapes:
            try:
                return max(0, shape.line.weight)
            except:
                continue
        return None
    
    @classmethod
    def set_line_weight(cls, shapes, value):
        value = max(0, value)
        shapes = list(pplib.iterate_shape_subshapes(shapes))
        bkt.apply_delta_on_ALT_key(
            # lambda shape, value: setattr(shape.Line, 'weight', max(0, value)), 
            cls._attr_setter,
            cls._attr_getter,
            shapes, value, shp_object="Line", attribute="weight")

    ### GALLERY ###
    @classmethod
    def get_item_count(cls):
        return len(cls.transparencies)
    
    @classmethod
    def get_item_label(cls, index):
        return "%s%%" % cls.transparencies[index]
    
    @classmethod
    def get_item_image(cls, index, context):
        return cls._get_image_for_transp(cls.transparencies[index], context)
    
    @classmethod
    def _get_image_for_transp(cls, transp, context):
        return context.python_addin.load_image( "transp_%s" % int(round(transp/10.0)*10) )


    @classmethod
    def fill_on_action_indexed(cls, selected_item, index, shapes):
        value = float(cls.transparencies[index])
        cls.set_fill_transparency(shapes, value)
    
    @classmethod
    def fill_get_selected_item_index(cls, context):
        try:
            return cls.transparencies.index(cls.get_fill_transparency(context.shapes))
        except:
            return -1

    @classmethod
    def line_on_action_indexed(cls, selected_item, index, shapes):
        value = float(cls.transparencies[index])
        cls.set_line_transparency(shapes, value)
    
    @classmethod
    def line_get_selected_item_index(cls, context):
        try:
            return cls.transparencies.index(cls.get_line_transparency(context.shapes))
        except:
            return -1


format_group = bkt.ribbon.Group(
    id="bkt_format_group",
    label="Format",
    image_mso='BehindText',
    children=[
        bkt.ribbon.RoundingSpinnerBox(
            id = 'fill_transparency',
            label="Transparenz Hintergrund",
            supertip="Ändere die Transparenz vom Hintergrund",
            show_label=False,
            round_int = True,
            image="fill_transparency",
            on_change = bkt.Callback(ShapeFormats.set_fill_transparency, shapes=True),
            get_text  = bkt.Callback(ShapeFormats.get_fill_transparency, shapes=True),
            get_enabled = bkt.Callback(ShapeFormats.get_fill_enabled),
        ),
        bkt.ribbon.RoundingSpinnerBox(
            id = 'line_transparency',
            label="Transparenz Linie/Rahmen",
            supertip="Ändere die Transparenz vom Rahmen bzw. der Linie",
            show_label=False,
            round_int = True,
            image="line_transparency",
            on_change = bkt.Callback(ShapeFormats.set_line_transparency, shapes=True),
            get_text  = bkt.Callback(ShapeFormats.get_line_transparency, shapes=True),
            get_enabled = bkt.Callback(ShapeFormats.get_line_enabled),
        ),
        bkt.ribbon.RoundingSpinnerBox(
            id = 'line_weight',
            label="Dicke Linie/Rahmen",
            supertip="Ändere die Dicke vom Rahmen bzw. der Linie",
            show_label=False,
            round_pt = True,
            rounding_factor=0.25,
            huge_step=1,
            big_step=0.5,
            small_step=0.25,
            image_mso="LineThickness",
            on_change = bkt.Callback(ShapeFormats.set_line_weight, shapes=True),
            get_text  = bkt.Callback(ShapeFormats.get_line_weight, shapes=True),
            get_enabled = bkt.Callback(ShapeFormats.get_line_enabled),
        ),
        bkt.ribbon.DialogBoxLauncher(idMso='ObjectFormatDialog')
    ]
)

fill_transparency_gallery = bkt.ribbon.Gallery(
    id="bkt_fill_transparency_menu",
    label="Transparenz Hintergrund",
    supertip="Setzt die Hintergrund-Transparenz auf den gewählten Wert.",
    show_label=False,
    show_item_label=True,
    image="fill_transparency",
    columns="1",
    get_enabled = bkt.Callback(ShapeFormats.get_fill_enabled),
    on_action_indexed = bkt.Callback(ShapeFormats.fill_on_action_indexed, shapes=True),
    get_selected_item_index = bkt.Callback(ShapeFormats.fill_get_selected_item_index, context=True),
    item_height="16",
    get_item_count=bkt.Callback(ShapeFormats.get_item_count),
    get_item_label=bkt.Callback(ShapeFormats.get_item_label),
    get_item_image=bkt.Callback(ShapeFormats.get_item_image, context=True),
    ### static definition of children has disadvantage that get_selected_item_index is called even if nothing 
    ### is selected, leading to an error message on ppt startup if UI error are cative.
    # children=[
    #     bkt.ribbon.Item(label="%s%%" % transp, image="transp_%s" % transp)
    #     for transp in ShapeFormats.transparencies
    # ]
)

line_transparency_gallery = bkt.ribbon.Gallery(
    id="bkt_line_transparency_menu",
    label="Transparenz Linie/Rahmen",
    supertip="Setzt die Linien-Transparenz auf den gewählten Wert.",
    show_label=False,
    show_item_label=True,
    image="line_transparency",
    columns="1",
    get_enabled = bkt.Callback(ShapeFormats.get_line_enabled),
    on_action_indexed = bkt.Callback(ShapeFormats.line_on_action_indexed, shapes=True),
    get_selected_item_index = bkt.Callback(ShapeFormats.line_get_selected_item_index, context=True),
    item_height="16",
    get_item_count=bkt.Callback(ShapeFormats.get_item_count),
    get_item_label=bkt.Callback(ShapeFormats.get_item_label),
    get_item_image=bkt.Callback(ShapeFormats.get_item_image, context=True),
    ### static definition of children has disadvantage that get_selected_item_index is called even if nothing 
    ### is selected, leading to an error message on ppt startup if UI error are cative.
    # children=[
    #     bkt.ribbon.Item(label="%s%%" % transp, image="transp_%s" % transp)
    #     for transp in ShapeFormats.transparencies
    # ]
)


class PictureFormat(object):
    @staticmethod
    def make_img_transparent(slide, shapes, transparency=0.5):
        if not bkt.message.confirmation("Das bestehende Bild wird dabei ersetzt. Fortfahren?"):
            return

        import tempfile, os
        filename = os.path.join(tempfile.gettempdir(), "bktimgtransp.png")

        for shape in shapes:
            if shape.Type != pplib.MsoShapeType["msoPicture"]:
                continue

            shape.Export(filename, 2) #2=ppShapeFormatPNG

            pic_shape = slide.Shapes.AddShape(
                shape.AutoShapeType,
                shape.Left, shape.Top,
                shape.Width, shape.Height
                )
            pic_shape.LockAspectRatio = -1
            pic_shape.Rotation = shape.Rotation
            pplib.set_shape_zorder(pic_shape, value=shape.ZOrderPosition)
            shape.PickUp()
            pic_shape.Apply()
            pic_shape.line.visible = shape.line.visible # line is not properly transferred by pickup-apply

            pic_shape.fill.UserPicture(filename)
            pic_shape.fill.transparency = transparency
            pic_shape.Select(replace=False)

            shape.Delete()
            os.remove(filename)



class PlaceholderConverter(object):
    @staticmethod
    def is_text_placeholder(shape):
        # return shape.Type == pplib.MsoShapeType["msoPlaceholder"] and shape.PlaceholderFormat.ContainedType in (pplib.MsoShapeType['msoTextBox'],pplib.MsoShapeType['msoAutoShape'] )
        return shape.Type == pplib.MsoShapeType["msoPlaceholder"]
    
    @classmethod
    def convert_placeholder(cls, shape):
        # new = pplib.replicate_shape(shape)
        new = shape.Duplicate()
        new.top, new.left = shape.top, shape.left
        shape.Delete()
        new.select(False)

    @classmethod
    def convert_shapes(cls, shapes):
        success=False
        for shape in shapes:
            if cls.is_text_placeholder(shape):
                try:
                    cls.convert_placeholder(shape)
                    success = True
                except:
                    logging.exception("placeholder conversion failed")

        if not success:
            bkt.message.warning("Aktuelle Auswahl enthält keine Platzhalter!")


class ShapeTableGallery(bkt.ribbon.Gallery):
    
    # item-settings for gallery
    #items = [ dict(label=l, style=s, shape_type=t) for l in ['1', 'a', 'A', 'I'] for t in ['circle', 'square'] for s in ['dark', 'light']  ]
    _columns = 6
    _rows = 8
    
    
    def __init__(self, **kwargs):
        self._margin = 0
        parent_id = kwargs.get('id') or ""
        my_kwargs = dict(
            label = 'Shape-Tabelle einfügen',
            columns = ShapeTableGallery._columns,
            image = 'shapetable',
            # image_mso = 'SlidesPerPage4Slides',
            screentip="Shape-Tabelle einfügen",
            supertip="Füge eine Tabelle aus Standard-Shapes ein",
            description="Füge eine Tabelle aus Standard-Shapes ein",
            children=[
                bkt.ribbon.Button(id=parent_id + "_margin0", label="Ohne Abstand", supertip="Abstand bei Shape-Tabelle deaktivieren", on_action=bkt.Callback(lambda: setattr(self, "_margin", 0)), get_image=bkt.Callback(lambda: self.get_toggle_image(0))),
                bkt.ribbon.Button(id=parent_id + "_margin10", label="Kleiner Abstand", supertip="Abstand bei Shape-Tabelle auf klein setzen", on_action=bkt.Callback(lambda: setattr(self, "_margin", 10)), get_image=bkt.Callback(lambda: self.get_toggle_image(10))),
                bkt.ribbon.Button(id=parent_id + "_margin20", label="Großer Abstand", supertip="Abstand bei Shape-Tabelle auf groß setzen", on_action=bkt.Callback(lambda: setattr(self, "_margin", 20)), get_image=bkt.Callback(lambda: self.get_toggle_image(20))),
            ]
        )
        my_kwargs.update(kwargs)

        super(ShapeTableGallery, self).__init__(**my_kwargs)
    
    
    def on_action_indexed(self, selected_item, index, slide):
        ''' create numberd shape according of settings in clicked element '''
        n_rows, n_cols = self.get_rows_cols_from_index(index)
        self.create_shape_table(slide, n_rows, n_cols)
    
    
    def create_shape_table(self, slide, rows, columns):
        
        ref_left,ref_top,ref_width,ref_height = pplib.slide_content_size(slide)
        target_width = ref_width + self._margin
        target_height = ref_height + self._margin
        
        shape_width = target_width/columns
        shape_height = target_height/rows
        
        for r in range(rows):
            for c in range(columns):
                slide.shapes.AddShape(
                    1, #msoShapeRectangle
                    ref_left+c*shape_width, ref_top+r*shape_height,
                    shape_width-self._margin, shape_height-self._margin)
        
        shapes = pplib.last_n_shapes_on_slide(slide, rows*columns)
        shapes.select()
        
    
    def get_rows_cols_from_index(self, index):
        n_cols = index%self._columns
        n_rows = (index-n_cols)//self._columns + 1
        n_cols += 1
        return n_rows, n_cols
    
    def get_item_count(self):
        return self._rows * self._columns
        
    def get_item_label(self, index):
        n_rows, n_cols = self.get_rows_cols_from_index(index)
        return "%sx%s" % (n_cols, n_rows)
    
    def get_item_screentip(self, index):
        return "Shape-Tabelle einfügen"
        
    def get_item_supertip(self, index):
        n_rows, n_cols = self.get_rows_cols_from_index(index)
        return "Füge eine %sx%s-Tabelle aus Standard-Shapes ein (%s Spalten, %s Zeilen)" % (n_cols, n_rows, n_cols, n_rows)
    
    def get_item_image(self, index):
        ''' creates an item image with numberd shape according to settings in the specified item '''
        n_rows, n_cols = self.get_rows_cols_from_index(index)
        
        # create bitmap, define pen/brush
        size_w = 60 #16*3
        size_h = round(size_w/16*9) #9*3
        img = Drawing.Bitmap(size_w, size_h)
        g = Drawing.Graphics.FromImage(img)
        # color_black = Drawing.ColorTranslator.FromOle(0)
        #color_light_grey  = Drawing.ColorTranslator.FromOle(14540253)
        # color_grey  = Drawing.ColorTranslator.FromHtml('#666')
        color_grey  = Drawing.Brushes.Gray
        pen = Drawing.Pen(color_grey,1)
        #brush = Drawing.SolidBrush(color_black)
        
        #Draw smooth rectangle/ellipse
        g.SmoothingMode = Drawing.Drawing2D.SmoothingMode.AntiAlias
        
        #square
        #g.DrawRectangle(pen, Drawing.Rectangle(0,0, size-1, size-1))
        
        width = round(size_w/n_cols-1)
        height = round(size_h/n_rows-1)
        for r in range(n_rows):
            for c in range(n_cols):
                g.DrawRectangle(pen, Drawing.Rectangle(c*width,r*height, width, height))
        
        return img
    
    def get_toggle_image(self, margin):
        if self._margin == margin:
            return self.get_check_image()
        else:
            return self.get_check_image(checked=False)
    

class ChessTableGallery(ShapeTableGallery):
    
    def __init__(self, **kwargs):
        parent_id = kwargs.get('id') or ""
        my_kwargs = dict(
            label = 'Shape-Schachbrett einfügen',
            image = 'shapechessboard',
            screentip="Shape-Schachbrett einfügen",
            supertip="Füge ein Schachbrett aus Standard-Shapes ein",
            description="Füge ein Schachbrett aus Standard-Shapes ein",
        )
        my_kwargs.update(kwargs)
        super(ChessTableGallery, self).__init__(**my_kwargs)

        #overwrite attributes
        self._margin = 10
        #new attributes
        self._insert_textboxes = True
        self.children.append(
            bkt.ribbon.Button(id=parent_id + "_txtboxes", label="Textboxen in Zellen", supertip="Abstand bei Shape-Tabelle auf groß setzen", on_action=bkt.Callback(lambda: setattr(self, "_insert_textboxes", not self._insert_textboxes)), get_image=bkt.Callback(lambda: self.get_check_image(self._insert_textboxes)))
            )
    
    def create_shape_table(self, slide, rows, columns):
        
        ref_left,ref_top,ref_width,ref_height = pplib.slide_content_size(slide)
        target_width = ref_width
        target_height = ref_height
        
        shape_width = (target_width-self._margin)/columns
        shape_height = (target_height-self._margin)/rows
        
        for c in range(columns):
            shp = slide.shapes.AddShape(
                1, #msoShapeRectangle
                ref_left+self._margin+c*shape_width, ref_top,
                shape_width-self._margin, target_height)
            # shp.Fill.Transparency = 0.5
        
        for r in range(rows):
            shp = slide.shapes.AddShape(
                1, #msoShapeRectangle
                ref_left, ref_top+self._margin+r*shape_height,
                target_width, shape_height-self._margin)
            shp.Fill.Transparency = 0.5
        
        num_to_sel = rows+columns

        if self._insert_textboxes:
            for r in range(rows):
                for c in range(columns):
                    shpTxt = slide.shapes.AddTextbox(
                        1, #msoTextOrientationHorizontal
                        ref_left+self._margin+c*shape_width, ref_top+self._margin+r*shape_height,
                        shape_width-self._margin, shape_height-self._margin)
                    shpTxt.TextFrame2.AutoSize = 0 #ppAutoSizeNone
                    shpTxt.TextFrame2.WordWrap = -1 #msoTrue
                    shpTxt.TextFrame2.TextRange.Text = "tbd"
            num_to_sel += rows*columns

        shapes = pplib.last_n_shapes_on_slide(slide, num_to_sel)
        shapes.select()



picture_format_tab = bkt.ribbon.Tab(
    idMso = "TabPictureToolsFormat",
    children = [
        bkt.ribbon.Group(
            id="bkt_pictureformat_group",
            label="Format",
            insert_after_mso="GroupPictureTools",
            children = [
                bkt.ribbon.Button(
                    id = 'make_img_transparent',
                    label="Transparenz ermöglichen",
                    supertip="Ersetzt das Bild durch ein Shape mit Bildfüllung, welches nativ transparent gemacht werden kann. Dabei wird das bestehende Bild exportiert und dann gelöscht, d.h. etwaige zugeschnittene Bereiche gehen verloren und Bildformate können nicht rückgängig gemacht werden.",
                    size="large",
                    show_label=True,
                    image_mso='PictureRecolorWashout',
                    on_action=bkt.Callback(PictureFormat.make_img_transparent),
                    # get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                ),
            ]
        )
    ]
)




# default ui for shape styling
styles_group = bkt.ribbon.Group(
    id="bkt_style_group",
    label='Stile',
    image_mso='ShapeFillColorPicker',
    children = [
        bkt.mso.control.ShapeFillColorPicker,
        bkt.mso.control.ShapeOutlineColorPicker,
        bkt.mso.control.ShapeEffectsMenu,
        bkt.mso.control.TextFillColorPicker,
        bkt.mso.control.TextOutlineColorPicker,
        bkt.mso.control.TextEffectsMenu,
        bkt.mso.control.OutlineWeightGallery,
        bkt.mso.control.OutlineDashesGallery,
        bkt.mso.control.ArrowStyleGallery,
        fill_transparency_gallery,
        line_transparency_gallery,
        bkt.mso.control.ShapeQuickStylesHome, #if ppt_customformats is active, this button is replaced
        bkt.ribbon.DialogBoxLauncher(idMso='ObjectFormatDialog')
    ]
)



shapes_group = bkt.ribbon.Group(
    id="bkt_shapes_group",
    label='Formen',
    image_mso='ShapesInsertGallery',
    children = [
        bkt.mso.control.ShapesInsertGallery,
        text.text_splitbutton,
        bkt.ribbon.Menu(
            image_mso='TableInsertGallery',
            label="Tabelle einfügen",
            show_label=False,
            supertip="Einfügen von Standard- oder Shape-Tabellen",
            item_size="large",
            children=[
                bkt.ribbon.MenuSeparator(title="PowerPoint-Tabelle"),
                bkt.mso.control.TableInsertGallery,
                bkt.ribbon.MenuSeparator(title="Shape-Tabelle"),
                ShapeTableGallery(id="insert_shape_table"),
                ChessTableGallery(id="insert_shape_chessboard")
            ]
        ),
        
        #bkt.mso.control.PictureInsertFromFilePowerPoint,
        shapelib_button,
        text.symbol_insert_splitbutton,
        bkt.ribbon.Menu(
            label='Spezialformen',
            show_label=False,
            image_mso='SmartArtInsert',
            screentip="Spezielle und Interaktive Formen ",
            supertip="Interaktive BKT-Shapes und spezielle zusammengesetzte Shapes einfügen, die sonst nur umständlich zu erstellen sind.",
            children = [
                bkt.ribbon.MenuSeparator(title="Einfügehilfen"),
                bkt.ribbon.Button(
                    id = 'segmented_circle',
                    label = "Kreissegmente…",
                    image = "segmented_circle",
                    screentip="Kreissegmente einfügen",
                    supertip="Erstelle Kreis mit Segmenten oder Chevrons.",
                    on_action=bkt.Callback(ShapeDialogs.show_segmented_circle_dialog)
                ),
                bkt.ribbon.Button(
                    id='agenda_textbox',
                    label="Agenda-Textbox einfügen",
                    supertip="Standard Agenda-Textbox einfügen, um daraus eine aktualisierbare Agenda zu generieren.",
                    imageMso="TextBoxInsert",
                    on_action=bkt.Callback(ToolboxAgenda.create_agenda_textbox_on_slide)
                ),
                NumberShapesGallery(id='number-labels-gallery'),
                bkt.ribbon.Menu(
                    label='Grafik-Tracker',
                    image = "Tracker",
                    screentip="Tracker erstellen oder ausrichten",
                    supertip="Einen Tracker aus einer Auswahl als Bild erstellen, verteilen und ausrichten.",
                    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                    children = [
                        bkt.ribbon.Button(
                            id = 'tracker',
                            label = "Tracker aus Auswahl erstellen",
                            #image = "Tracker",
                            screentip="Tracker aus Auswahl erstellen",
                            supertip="Erstelle aus den markierten Shapes einen Tracker.\nDer Shape-Stil für Highlights wird aus dem zuerst markierten Shape (in der Regel oben links) bestimmt. Der Shape-Stil für alle anderen Shapes wird aus dem als zweites markierten Shape bestimmt.",
                            on_action=bkt.Callback(TrackerShape.generateTracker, shapes=True, shapes_min=2, context=True),
                            get_enabled = bkt.apps.ppt_shapes_min2_selected,
                        ),
                        bkt.ribbon.Button(
                            id = 'tracker_distribute',
                            label = "Tracker auf Folien verteilen",
                            #image = "Tracker",
                            screentip="Alle Tracker verteilen",
                            supertip="Verteilen der ausgewählten Tracker auf die Folgefolien und ausrichten.",
                            on_action=bkt.Callback(TrackerShape.distributeTracker, shapes=True, shapes_min=2, context=True),
                            get_enabled = bkt.apps.ppt_shapes_min2_selected,
                        ),
                        bkt.ribbon.MenuSeparator(),
                        bkt.ribbon.Button(
                            id = 'tracker_align',
                            label = "Alle Tracker ausrichten",
                            #image = "Tracker",
                            screentip="Alle Tracker ausrichten",
                            supertip="Ausrichten (Position, Größe, Rotation) aller Tracker (auf allen Folien) anhand des ausgewählten Tracker.",
                            on_action=bkt.Callback(TrackerShape.alignTracker, shape=True, context=True),
                            get_enabled = bkt.Callback(TrackerShape.isTracker, shape=True),
                        ),
                        bkt.ribbon.Button(
                            id = 'tracker_remove',
                            label = "Alle Tracker löschen",
                            #image = "Tracker",
                            screentip="Alle Tracker löschen",
                            supertip="Löschen aller Tracker (auf allen Folien) anhand des ausgewählten Tracker.",
                            on_action=bkt.Callback(TrackerShape.removeTracker, shape=True, context=True),
                            get_enabled = bkt.Callback(TrackerShape.isTracker, shape=True),
                        ),
                    ]
                ),
                bkt.ribbon.MenuSeparator(title="Interaktive Formen"),
                bkt.ribbon.Button(
                    id = 'standard_process',
                    label = "Prozesspfeile…",
                    image = "process_chevrons",
                    screentip="Prozess-Pfeile einfügen",
                    supertip="Erstelle Standard Prozess-Pfeile.",
                    on_action=bkt.Callback(ShapeDialogs.show_process_chevrons_dialog)
                ),
                bkt.ribbon.Button(
                    id = 'headered_pentagon',
                    label = "Prozessschritt mit Kopfzeile",
                    image = "headered_pentagon",
                    screentip="Prozess-Schritt-Shape mit Kopfzeile erstellen",
                    supertip="Erstelle einen Prozess-Pfeil mit Header-Shape. Das Header-Shape kann im Prozess-Pfeil über Kontext-Menü des Header-Shapes passend angeordnet werden.",
                    on_action=bkt.Callback(ShapeDialogs.create_headered_pentagon)
                ),
                bkt.ribbon.Button(
                    id = 'headered_chevron',
                    label = "2. Prozessschritt mit Kopfzeile",
                    image = "headered_chevron",
                    screentip="Prozess-Schritt-Shape mit Kopfzeile erstellen",
                    supertip="Erstelle einen Prozess-Pfeil mit Header-Shape. Das Header-Shape kann im Prozess-Pfeil über Kontext-Menü des Header-Shapes passend angeordnet werden.",
                    on_action=bkt.Callback(ShapeDialogs.create_headered_chevron)
                ),
                harvey.harvey_create_button,
                bkt.ribbon.Menu(
                    id="traffic_light_menu",
                    label="Ampel",
                    image="traffic_light",
                    screentip='Status-Ampel erstellen',
                    children=[
                        bkt.ribbon.Button(
                            id="traffic_light",
                            label="Ampel vertikal",
                            image="traffic_light",
                            screentip='Status-Ampel vertikal erstellen',
                            supertip="Füge eine Status-Ampel ein. Die Status-Farbe der Ampel kann per Kontext-Dialog konfiguriert werden.",
                            on_action=bkt.Callback(lambda slide: ShapeDialogs.create_traffic_light(slide, "vertical"), slide=True)
                        ),
                        bkt.ribbon.Button(
                            label="Ampel horizontal",
                            image="traffic_light2",
                            screentip='Status-Ampel horizontal erstellen',
                            supertip="Füge eine Status-Ampel ein. Die Status-Farbe der Ampel kann per Kontext-Dialog konfiguriert werden.",
                            on_action=bkt.Callback(lambda slide: ShapeDialogs.create_traffic_light(slide, "horizontal"), slide=True)
                        ),
                        bkt.ribbon.Button(
                            label="Ampel Punkt",
                            image="traffic_light3",
                            screentip='Status-Ampel einfach erstellen',
                            supertip="Füge eine Status-Ampel ein. Die Status-Farbe der Ampel kann per Kontext-Dialog konfiguriert werden.",
                            on_action=bkt.Callback(lambda slide: ShapeDialogs.create_traffic_light(slide, "simple"), slide=True)
                        ),
                    ]
                ),
                stateshapes.likert_button,
                stateshapes.checkbox_button,
                bkt.ribbon.MenuSeparator(title="Verbindungsflächen"),
                bkt.ribbon.Button(
                    id = 'connector_h',
                    label = "Horizontale Verbindungsfläche",
                    image = "ConnectorHorizontal",
                    supertip="Erstelle eine horizontale Verbindungsfläche zwischen den vertikalen Seiten (links/rechts) von zwei Shapes.",
                    on_action=bkt.Callback(ShapeConnectors.addHorizontalConnector, context=True, shapes=True, shapes_min=2, shapes_max=2),
                    get_enabled = bkt.apps.ppt_shapes_exactly2_selected,
                ),
                bkt.ribbon.Button(
                    id = 'connector_v',
                    label = "Vertikale Verbindungsfläche",
                    image = "ConnectorVertical",
                    supertip="Erstelle eine vertikale Verbindungsfläche zwischen den horizontalen Seiten (oben/unten) von zwei Shapes.",
                    on_action=bkt.Callback(ShapeConnectors.addVerticalConnector, context=True, shapes=True, shapes_min=2, shapes_max=2),
                    get_enabled = bkt.apps.ppt_shapes_exactly2_selected,
                ),
                bkt.ribbon.MenuSeparator(),
                bkt.ribbon.Button(
                    id = 'connector_update',
                    label = "Verbindungsfläche neu verbinden",
                    image = "ConnectorUpdate",
                    supertip="Aktualisiere die Verbindungsfläche nachdem sich die verbundenen Shapes geändert haben.",
                    on_action=bkt.Callback(ShapeConnectors.update_connector_shape, context=True, shape=True),
                    get_enabled = bkt.Callback(ShapeConnectors.is_connector, shape=True),
                ),
            ]
        ),
        bkt.mso.control.ShapeChangeShapeGallery,
        bkt.ribbon.Menu(
            image_mso='CombineShapesMenu',
            label="Shape verändern",
            supertip="Funktionen um Shape-Punkte zu manipulieren, Shapes zu duplizieren, und Text in Symbol/Grafik umzuwandeln",
            show_label=False,
            children=[
                bkt.ribbon.MenuSeparator(title="Formen manipulieren"),
                bkt.ribbon.Button(
                    label="Shapes teilen/vervielfachen…",
                    image="split_horizontal",
                    screentip="Shapes teilen oder vervielfachen",
                    supertip="Shape horizontal/vertikal in mehrere Shapes teilen oder verfielfachen.",
                    on_action=bkt.Callback(ShapeDialogs.shape_split),
                    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                ),
                bkt.ribbon.Button(
                    label="Shapes skalieren…",
                    image_mso="DiagramScale",
                    screentip="Shapes skalieren",
                    supertip="Shape-Größe inkl. aller Elemente/Eigenschaften (Schriftgröße, Konturen, etc.) gleichmäßig ändern.",
                    on_action=bkt.Callback(ShapeDialogs.shape_scale),
                    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                ),
                bkt.ribbon.Button(
                    label="Platzhalter in Textbox umwandeln",
                    image_mso="ConvertTableToText",
                    supertip="Wandelt alle markierten Text-Platzhalter in echte Textboxen um, die u.A. eine Gruppierung erlauben.",
                    on_action=bkt.Callback(PlaceholderConverter.convert_shapes),
                    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                ),
                bkt.mso.control.ObjectEditPoints,
                bkt.ribbon.Button(
                    label="Text/Symbol zu Shapes umwandeln",
                    image_mso="TextEffectTransformGallery",
                    screentip="Texte bzw. Symbole werden in Standardshapes umgewandelt",
                    supertip="Ersetzt den Text einer Textbox in Shapes. Damit kann man bspw. einen Icon-Font in echte Icons umwandeln.",
                    on_action=bkt.Callback(ShapesMore.texts_to_shapes),
                    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                ),
                bkt.ribbon.MenuSeparator(title="Formen zusammenführen"),
                bkt.mso.control.ShapesUnion,
                bkt.mso.control.ShapesCombine,
                bkt.mso.control.ShapesFragment,
                bkt.mso.control.ShapesIntersect,
                bkt.mso.control.ShapesSubtract
            ]
        ),
        bkt.ribbon.Menu(
            label='Mehr',
            show_label=False,
            image_mso='TableDesign',
            screentip="Weitere Funktionen",
            supertip="Standardobjekte (Bilder, Smart-Art, etc.) einfügen, Shapes verstecken und wieder anzeigen, Kopf- und Fußzeile anpassen",
            children = [
                bkt.ribbon.MenuSeparator(title="Bilder und Objekte"),
                bkt.mso.control.PictureInsertFromFilePowerPoint,
                bkt.mso.control.OleObjectctInsert,
                bkt.mso.control.ClipArtInsertDialog,
                bkt.mso.control.SmartArtInsert,
                bkt.mso.control.ChartInsert,
                # bkt.mso.control.IconInsertFromFile, #only available in Office 2016 with 365 subscription
                bkt.ribbon.MenuSeparator(title="Text & Beschriftungen"),
                bkt.mso.control.HeaderFooterInsert,
                bkt.mso.control.DateAndTimeInsert,
                bkt.mso.control.NumberInsert,
                bkt.mso.control.InsertNewComment,
                bkt.ribbon.MenuSeparator(title="Ein-/Ausblenden"),
                bkt.ribbon.Button(
                    id = 'toggle_shapes_visibility',
                    label = "Shapes vertecken/einblenden",
                    image_mso="ShapesSubtract",
                    supertip="Wenn Shapes ausgewählt sind, verstecke alle markierten Shapes (visible=False), anderen falls mache Shapes wieder sichtbar (visible=True). Gibt es keine unsichtbaren Shapes, werden die zuletzt versteckten Shapes erneut versteckt.",
                    on_action=bkt.Callback(ShapesMore.toggle_shapes_visibility),
                    # get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                ),
                # bkt.ribbon.Button(
                #     id = 'hide_shape',
                #     label = u"Shapes verstecken",
                #     image_mso="ShapesSubtract",
                #     supertip="Verstecke alle markierten Shapes (visible=False).",
                #     on_action=bkt.Callback(ShapesMore.hide_shapes),
                #     get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                # ),
                bkt.ribbon.Button(
                    id = 'show_shapes',
                    label = "Alle versteckten Shapes einblenden",
                    image_mso="VisibilityVisible",
                    supertip="Mache alle versteckten Shapes (visible=False) wieder sichtbar.",
                    on_action=bkt.Callback(ShapesMore.show_invisible_shapes)
                ),

            ]
        ),
    ]
)
