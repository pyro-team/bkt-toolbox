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

# from bkt.library.algorithms import get_bounding_nodes, mid_point

# from bkt import dotnet
# Drawing = dotnet.import_drawing()
# office = dotnet.import_officecore()

# other toolbox modules
from .chartlib import shapelib_button
# from .agenda import ToolboxAgenda
from . import text
# from . import harvey
# from . import stateshapes






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
        bkt.ribbon.DynamicMenu(
            image_mso='TableInsertGallery',
            label="Tabelle einfügen",
            show_label=False,
            supertip="Einfügen von Standard- oder Shape-Tabellen",
            # item_size="large", #not supported by dynamic menu
            get_content=bkt.CallbackLazy("toolbox.models.shapes_menu", "shapes_table_menu"),
        ),
        
        #bkt.mso.control.PictureInsertFromFilePowerPoint,
        shapelib_button,
        text.symbol_insert_splitbutton,
        bkt.ribbon.DynamicMenu(
            label='Spezialformen',
            show_label=False,
            image_mso='SmartArtInsert',
            screentip="Spezielle und Interaktive Formen ",
            supertip="Interaktive BKT-Shapes und spezielle zusammengesetzte Shapes einfügen, die sonst nur umständlich zu erstellen sind.",
            get_content=bkt.CallbackLazy("toolbox.models.shapes_menu", "shapes_interactive_menu"),
        ),
        bkt.mso.control.ShapeChangeShapeGallery,
        bkt.ribbon.DynamicMenu(
            image_mso='CombineShapesMenu',
            label="Shape verändern",
            supertip="Funktionen um Shape-Punkte zu manipulieren, Shapes zu duplizieren, und Text in Symbol/Grafik umzuwandeln",
            show_label=False,
            get_content=bkt.CallbackLazy("toolbox.models.shapes_menu", "shapes_change_menu"),
        ),
        bkt.ribbon.DynamicMenu(
            label='Mehr',
            show_label=False,
            image_mso='TableDesign',
            screentip="Weitere Funktionen",
            supertip="Standardobjekte (Bilder, Smart-Art, etc.) einfügen, Shapes verstecken und wieder anzeigen, Kopf- und Fußzeile anpassen",
            get_content=bkt.CallbackLazy("toolbox.models.shapes_menu", "shapes_more_menu"),
        ),
    ]
)
