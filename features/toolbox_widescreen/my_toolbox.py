# -*- coding: utf-8 -*-
'''
Created on 24.07.2014

@author: fstsallmann
'''


import bkt

# import toolbox modules with ui
from toolbox import arrange, harvey, info, language, shape_adjustments, shape_selection, shapes as mod_shapes, slides, stateshapes, text


# FIXME: this should be easier
from os.path import dirname, join, normpath, realpath
bkt.apps.Resources.root_folders.append(normpath(join(dirname(realpath(__file__)), '..', 'toolbox', 'resources')))



# define context-menus and context-tabs
from toolbox import context_menus


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
        mod_shapes.fill_transparency_gallery,
        mod_shapes.line_transparency_gallery,
        bkt.mso.control.ShapeQuickStylesHome, #TODO: replace this with customformats feature
        bkt.ribbon.DialogBoxLauncher(idMso='ObjectFormatDialog')
    ]
)


arrange_advanced_small_group = bkt.ribbon.Group(
    id="bkt_arrage_adv_small_group",
    label=u'Erw. Anordnen',
    image='arrange_left_at_left',
    children=[
        arrange.arrange_advaced.get_button("arrange_left_at_left", "-small",        label="Links an Links"),
        arrange.arrange_advaced.get_button("arrange_right_at_right", "-small",      label="Rechts an Rechts"),
        arrange.arrange_advaced.get_button("arrange_middle_at_middle", "-small",    label="Mitte an Mitte"),
        arrange.arrange_advaced.get_button("arrange_top_at_top", "-small",          label="Oben an Oben"),
        arrange.arrange_advaced.get_button("arrange_bottom_at_bottom", "-small",    label="Unten an Unten"),
        arrange.arrange_advaced.get_button("arrange_vmiddle_at_vmiddle", "-small",  label="Mitte an Mitte"),

        bkt.ribbon.Button(
            id="arrange_quick_position-small",
            label="Position",
            show_label=False,
            image_mso="ControlAlignToGrid",
            on_action=bkt.Callback(arrange.arrange_advaced.arrange_quick_position),
            get_enabled=bkt.Callback(arrange.arrange_advaced.enabled),
            screentip="Gleiche Position wie Master",
        ),
        bkt.ribbon.Button(
            id="arrange_quick_size-small",
            label="Größe",
            show_label=False,
            image_mso="SizeToControlHeightAndWidth",
            on_action=bkt.Callback(arrange.arrange_advaced.arrange_quick_size),
            get_enabled=bkt.Callback(arrange.arrange_advaced.enabled),
            screentip="Gleiche Größe wie Master",
        ),

        arrange.arrange_advaced.get_master_button("-small", show_label=False)
    ]
)


# ==============================
# = Definition of Toolbox-Tabs =
# ==============================

bkt.powerpoint.add_tab(bkt.ribbon.Tab(
    id="my_bkt_powerpoint_toolbox",
    label=u'BKT 1/2',
    insert_before_mso="TabHome",
    # get_visible defaults to False during async-startup
    get_visible=bkt.Callback(lambda:True),
    children = [
        shape_selection.clipboard_group,
        slides.slides_group,
        bkt.mso.group.GroupFont,
        bkt.mso.group.GroupParagraph,
        mod_shapes.shapes_group,
        styles_group,
        #mod_shapes.size_group,
        mod_shapes.pos_size_group,
        arrange.arrange_group,
        arrange_advanced_small_group,
        arrange.distance_rotation_group,
        text.innenabstand_gruppe,
        text.paragraph_group,
        text.paragraph_indent_group,
        shape_adjustments.adjustments_group,
        info.info_group
    ]
))

bkt.powerpoint.add_tab(bkt.ribbon.Tab(
    id="bkt_powerpoint_toolbox_extensions",
    #id_q="nsBKT:powerpoint_toolbox_extensions",
    #insert_after_q="nsBKT:powerpoint_toolbox_advanced",
    insert_before_mso="TabHome",
    label=u'BKT 2/2',
    # get_visible defaults to False during async-startup
    get_visible=bkt.Callback(lambda:True),
    children = [
        arrange.arrange_advanced_group,
        arrange.euclid_angle_group,
        mod_shapes.format_group,
        #arrange.arrange_adv_easy_group,
        #mod_shapes.split_shapes_group,
        #language.sprachen_gruppe,
        stateshapes.stateshape_gruppe,
    ]
), extend=True)



