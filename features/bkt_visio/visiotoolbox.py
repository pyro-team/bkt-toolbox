# -*- coding: utf-8 -*-
'''
Created on 2016-04-27
@author: Tobias Schickling, Florian Stallmann
'''

from __future__ import absolute_import

import bkt

from . import arrange
from . import connection
from . import text

# from . import tests

# reuse settings-menu from bkt-framework
import modules.settings as settings

version_short = 'v0.4b'
version_long  = 'Visio Toolbox v0.4 beta'

class DuplicateOnTarget(object):

    @staticmethod
    def copy_on_target(selection):
        selection.Copy("&H1")

    @staticmethod
    def paste_on_target(context):
        context.app.ActivePage.Paste("&H1")


ablage_gruppe = bkt.ribbon.Group(
    label='Ablage',
    image_mso='Copy',
    children=[
        bkt.mso.control.Copy,
        bkt.ribbon.Button(
            id = 'copy_on_target',
            label="Kopieren mit Koordinaten",
            show_label=False,
            image_mso='Copy',
            screentip="Kopieren mit Koordinaten",
            supertip="Shapes werden mit ihren ursprünglichen Koordinatenpositionen kopiert",
            on_action=bkt.Callback(DuplicateOnTarget.copy_on_target, selection=True)
        ),
        bkt.mso.control.Cut,
        bkt.mso.control.PasteMenu,
        bkt.ribbon.Button(
            id = 'paste_on_target',
            label="Einfügen auf Koordinaten",
            show_label=False,
            image_mso='PasteAsNestedTable',
            screentip="Einfügen auf Koordinaten",
            supertip="Shapes werden an ihren ursprünglichen Koordinatenpositionen eingefügt",
            on_action=bkt.Callback(DuplicateOnTarget.paste_on_target, context=True)
        ),
        bkt.mso.control.FormatPainter
    ]
)

# tools_gruppe = bkt.ribbon.Group(
    # label="Tools",
    # image_mso='GroupTools',
    # children=[
        # bkt.mso.control.PointerTool(show_label=True),
        # bkt.mso.control.ConnectorTool(show_label=True),
        # bkt.mso.control.TextTool(show_label=True),
        # bkt.mso.control.DrawingToolsMenu,
        # bkt.mso.control.ConnectionPointTool,
        # bkt.mso.control.TextBlockTool
    # ]
# )

formarten_gruppe = bkt.ribbon.Group(
    label="Formarten",
    image_mso='ShapeFillColorPicker',
    children=[
        bkt.mso.control.FillColorGallery,
        bkt.mso.control.LineColorGallery,
        bkt.mso.control.ShapeEffectsMenu,
        bkt.mso.control.LineWeightGallery,
        bkt.mso.control.LinePatternGallery,
        bkt.mso.control.LineArrowsGallery,
        bkt.mso.control.TextBoxInsertMenu,
        bkt.mso.control.ConnectorsGallery,
        #bkt.mso.control.ContainerGallery,
        bkt.mso.control.CalloutGallery,
        bkt.ribbon.DialogBoxLauncher(idMso='ObjectFormatDialog')
    ]
)

container_gruppe = bkt.ribbon.Group(
    label="Container",
    image_mso='ContainerGallery',
    children=[
        bkt.mso.control.ContainerGallery(size="large"),
        #bkt.mso.control.ContainerStyles,
        bkt.mso.control.SelectContents,
        bkt.mso.control.ContainerLock,
        bkt.mso.control.ContainerDisband,
    ]
)

formen_gruppe = bkt.ribbon.Group(
    label='Formen',
    image_mso='ShapesObjectIntersect',
    children = [
        bkt.mso.control.ShapesObjectUnion,
        bkt.mso.control.ShapesObjectIntersect,
        bkt.mso.control.ShapesObjectSubtract,
        bkt.mso.control.ShapesObjectCombine,
        bkt.mso.control.ShapesObjectFragment,
        bkt.ribbon.Separator(),
        bkt.mso.control.ShapesObjectOffset,
        bkt.mso.control.ShapesObjectTrim,
        bkt.mso.control.ShapesObjectJoin,
        #bkt.mso.control.ShapesObjectUpdateAlignmentBox,
    ]
)

bearbeiten_gruppe = bkt.ribbon.Group(
    label='Bearbeiten',
    image_mso='FindMenu',
    children = [
        bkt.mso.control.ReplaceShape(show_label=True, size="large"),
        bkt.mso.control.FindMenu(show_label=True),
        bkt.mso.control.LayersMenu(show_label=True),
        bkt.mso.control.SelectMenu(show_label=True),
    ]
)

kleben_gruppe = bkt.ribbon.Group(
    label='Kleben/Ausrichten',
    image_mso='GlueToggle',
    children = [
        bkt.mso.control.GlueToggle(show_label=False),
        bkt.mso.control.GlueToConnectionPoints(show_label=False),
        bkt.mso.control.GlueToGuides(show_label=False),
        bkt.mso.control.GlueToShapeGeometry(show_label=False),
        bkt.mso.control.GlueToShapeVertices(show_label=False),
        bkt.mso.control.GlueToShapeHandles(show_label=False),
        bkt.ribbon.Separator(),
        bkt.mso.control.SnapToggle(show_label=False),
        bkt.mso.control.SnapToDrawingAids(show_label=False),
        bkt.mso.control.SnapToGridVisio(show_label=False),
        bkt.mso.control.SnapToShapeIntersections(show_label=False),
        bkt.mso.control.SnapToRulerSubdivisions(show_label=False),
        bkt.mso.control.SnapToAlignmentBox(show_label=False),
        bkt.ribbon.DialogBoxLauncher(idMso='SnapAndGlueDialog')
    ]
)

kreuzung_gruppe = bkt.ribbon.Group(
    label='Kreuzung',
    image_mso='JumpStyleArc',
    children = [
        bkt.mso.control.JumpStyleArc(show_label=False),
        bkt.mso.control.JumpStyleGap(show_label=False),
        bkt.mso.control.JumpStyleSquare(show_label=False),
        bkt.mso.control.JumpStyle2Sides(show_label=False),
        bkt.mso.control.JumpStylePageDefault(show_label=False),
        bkt.mso.control.JumpsAddTo(show_label=False),
        bkt.ribbon.DialogBoxLauncher(idMso='PageSetupLayoutAndRoutingDialog')
    ]
)

shapedesign_gruppe = bkt.ribbon.Group(
    label='Shape-Design',
    image_mso='ShowShapeSheet',
    children = [
        bkt.mso.control.ShowShapeSheet(show_label=True, size="large"),
        bkt.mso.control.ShapeName(show_label=True),
        bkt.mso.control.ShapeBehavior(show_label=True),
        bkt.mso.control.ShapeProtection(show_label=True)
    ]
)

info_gruppe = bkt.ribbon.Group(
    label="Settings",
    children=[
        settings.settings_menu,
        bkt.ribbon.Button(label=version_short, screentip="Toolbox", supertip=version_long + "\n" + bkt.__release__),
    ]
)




bkt.visio.add_tab(
    bkt.ribbon.Tab(
        id="bkt_visio_toolbox",
        #id_q="nsBKT:visio_toolbox",
        label=u"Toolbox 1/2",
        insert_before_mso="TabHome",
        get_visible=bkt.Callback(lambda: True),
        children = [
            ablage_gruppe,
            bkt.mso.group.GroupFont,
            bkt.mso.group.GroupParagraph,
            bkt.mso.group.GroupTools,
            # tools_gruppe,
            formarten_gruppe,
            container_gruppe,
            arrange.anordnen_gruppe,
            arrange.objektabstand_gruppe,
            arrange.pos_size_group,
            bearbeiten_gruppe,
            shapedesign_gruppe,
            info_gruppe,
        ]
    )
)

bkt.visio.add_tab(
    bkt.ribbon.Tab(
        id="bkt_visio_toolbox_advanced",
        #id_q="nsBKT:visio_toolbox_advanced",
        label=u"Toolbox 2/2",
        insert_before_mso="TabHome",
        get_visible=bkt.Callback(lambda: True),
        children = [
            text.innenabstand_gruppe,
            formen_gruppe,
            kleben_gruppe,
            kreuzung_gruppe,
            connection.verbindungspunkte_gruppe,
        ]
    )
)