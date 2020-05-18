# -*- coding: utf-8 -*-
'''
Created on 2017-07-18
@author: Florian Stallmann
'''

from __future__ import absolute_import

import bkt
import bkt.library.excel.helpers as xllib

# reuse settings-menu from bkt-framework
import modules.settings as settings

from . import sheets
from . import cells
from . import selection
from . import books

version_short = 'v0.6b'
version_long  = 'Excel Toolbox v0.6 beta'


info_gruppe = bkt.ribbon.Group(
    label="Settings",
    children=[
        settings.settings_menu,
        bkt.ribbon.Button(label=version_short, screentip="Toolbox", supertip=version_long + "\n" + bkt.__release__, on_action=bkt.Callback(settings.BKTInfos.show_version_dialog)),
    ]
)


# ===============================
# = Definition des Toolbox-Tabs =
# ===============================

bkt.excel.add_tab(bkt.ribbon.Tab(
    id='bkt_excel_toolbox',
    #id_q='nsBKT:excel_toolbox',
    label=u'Toolbox 1/3',
    insert_before_mso="TabHome",
    get_visible=bkt.Callback(lambda: True),
    children = [
        #bkt.mso.group.GroupClipboard,
        bkt.ribbon.Group(
            label="Ablage",
            image_mso="GroupClipboard",
            children=[
                bkt.mso.control.PasteMenu,
                bkt.mso.control.CopySplitButton,
                bkt.mso.control.FormatPainter,
            ]
        ),
        bkt.mso.group.GroupFont,
        bkt.ribbon.Group(
            label="Sonst.",
            children=[
                bkt.ribbon.ToggleButton(
                    id = 'halign_centeracross_toggle',
                    label="Über Auswahl zentrieren",
                    show_label=False,
                    image_mso='MergeCellsAcross',
                    supertip="Ausgewählte zellen 'Über Auswahl zentriert' ausrichten, d.h es werden verbundene Zellen simuliert.",
                    on_toggle_action=bkt.Callback(lambda selection, pressed: cells.Format.horiz_align(selection, 7, pressed), selection=True), #xlHAlignCenterAcrossSelection 
                    get_pressed=bkt.Callback(lambda selection: cells.Format.horiz_align_pressed(selection, 7), selection=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                # bkt.mso.control.MergeCellsAcross(show_label=False),
                #bkt.mso.control.SymbolInsert(show_label=False),
                bkt.mso.control.AlignJustify(show_label=False),
                #bkt.mso.control.ParagraphDistributed(show_label=False),
                bkt.mso.control.Strikethrough(show_label=False)
            ]
        ),
        bkt.mso.group.GroupAlignmentExcel,
        bkt.mso.group.GroupNumber,
        bkt.mso.group.GroupStyles,
        # bkt.ribbon.Group(
        #     label="Formatvorlagen",
        #     image_mso="ConditionalFormattingMenu",
        #     children=[
        #         bkt.mso.control.ConditionalFormattingMenu(size="large"),
        #         bkt.mso.control.FormatAsTableGallery(size="large")
        #     ]
        # ),
        bkt.mso.group.GroupCells,
        bkt.mso.group.GroupEditingExcel,
        #bkt.mso.group.GroupOutline,
        bkt.ribbon.Group(
            label="Gliederung",
            image_mso="GroupOutline",
            children=[
                bkt.mso.control.OutlineGroupMenu(size="large"),
                bkt.mso.control.OutlineUngroupMenu(size="large"),
                #bkt.mso.control.OutlineSubtotals,
                bkt.mso.control.OutlineShowDetail,
                bkt.mso.control.OutlineHideDetail,
                bkt.mso.control.OutlineSymbolsShowHide,
                bkt.ribbon.DialogBoxLauncher(idMso='OutlineSettings')
            ]
        ),
        info_gruppe
    ]
))

bkt.excel.add_tab(bkt.ribbon.Tab(
    id='bkt_excel_toolbox_p2',
    #id_q='nsBKT:excel_toolbox',
    label=u'Toolbox 2/3 BETA',
    insert_before_mso="TabHome",
    get_visible=bkt.Callback(lambda: True),
    children = [
        cells.zellen_inhalt_gruppe,
        cells.zellen_format_gruppe,
        cells.comments_gruppe,
        bkt.ribbon.Group(
            label="Rahmen",
            image_mso="BorderDrawMenu",
            children=[
                bkt.mso.control.BorderTop(),
                bkt.mso.control.BorderLeft(),
                bkt.mso.control.BorderInsideHorizontal(),

                bkt.mso.control.BorderRight(),
                bkt.mso.control.BorderBottom(),
                bkt.mso.control.BorderInsideVertical(),

                #bkt.mso.control.BorderOutside(),
                #bkt.mso.control.BorderInside(),

                bkt.mso.control.BordersAll(),
                bkt.mso.control.BorderNone(),
                #bkt.mso.control.BordersGallery(),
                bkt.mso.control.BorderColorPickerExcel(),
                bkt.ribbon.DialogBoxLauncher(idMso='BordersMoreDialog')
            ]
        ),
        bkt.ribbon.Group(
            label="Definierte Namen",
            image_mso="NameManager",
            children=[
                bkt.mso.control.NameDefineMenu(show_label=True),
                bkt.mso.control.NameUseInFormula(show_label=True),
                bkt.mso.control.NameCreateFromSelection(show_label=True),
                bkt.ribbon.DialogBoxLauncher(idMso='NameManager')
            ]
        ),
        bkt.ribbon.Group(
            label="Datentools",
            image_mso="RemoveDuplicates",
            children=[
                bkt.mso.control.PivotTableInsertMenu(size="large"),
                bkt.mso.control.ConvertTextToTable(show_label=True),
                bkt.mso.control.RemoveDuplicates(show_label=True),
                bkt.mso.control.DataValidationMenu(show_label=True)
            ]
        ),
        bkt.ribbon.Group(
            label="Fenster",
            image_mso="WindowNew",
            children=[
                bkt.mso.control.WindowsArrangeAll(show_label=True),
                bkt.mso.control.ViewFreezePanesGallery(show_label=True),
                bkt.mso.control.WindowSideBySideSynchronousScrolling(show_label=True)
            ]
        ),
        bkt.ribbon.Group(
            label="Drucken",
            image_mso="PrintAreaMenu",
            children=[
                bkt.mso.control.PageOrientationGallery(show_label=True),
                bkt.mso.control.PageScaleToFitWidth(show_label=True),
                bkt.mso.control.PageScaleToFitHeight(show_label=True),
                bkt.ribbon.DialogBoxLauncher(idMso='PageSetupPageDialog')
            ]
        ),
    ]
))

bkt.excel.add_tab(bkt.ribbon.Tab(
    id='bkt_excel_toolbox_advanced',
    #id_q='nsBKT:excel_toolbox_advanced',
    label=u'Toolbox 3/3 BETA',
    insert_before_mso="TabHome",
    get_visible=bkt.Callback(lambda: True),
    children = [
        selection.selektion_gruppe,
        sheets.blatt_gruppe,
        books.mappen_gruppe,
    ]
))
