# -*- coding: utf-8 -*-
'''
Created on 06.02.2018

@author: rdebeerst
'''

import bkt


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
                        bkt.mso.button.PasteAsPicture,
                        bkt.ribbon.MenuSeparator(title="Einfügen-Spezial"),
                        bkt.ribbon.Button(
                            id='paste_to_slides',
                            label="Auf ausgewählte Folien einfügen",
                            supertip="Zwischenablage auf allen ausgewählten Folien gleichzeitig einfügen.",
                            image_mso='PasteDuplicate',
                            on_action=bkt.CallbackLazy("toolbox.models.copy_paste_format", "SlidesMore", "paste_to_slides", slides=True),
                        ),
                        bkt.ribbon.Button(
                            id='paste_as_link',
                            label="Als Verknüpfung einfügen",
                            supertip="Zwischenablage als verknüpftes Element (bspw. Bild, OLE-Objekt) einfügen.",
                            image_mso='PasteLink',
                            on_action=bkt.CallbackLazy("toolbox.models.copy_paste_format", "SlidesMore", "paste_as_link", slide=True),
                        ),
                        bkt.ribbon.Button(
                            id='paste_and_replace',
                            label="Mit Zwischenablage ersetzen",
                            supertip="Markierte Shapes mit dem Inhalt der Zwischenablage ersetzen und dabei Größe und Position erhalten.",
                            image_mso='PasteSingleCellExcelTableDestinationFormatting',
                            on_action=bkt.CallbackLazy("toolbox.models.copy_paste_format", "SlidesMore", "paste_and_replace_shapes", slide=True, shapes=True),
                            get_enabled=bkt.apps.ppt_shapes_or_text_selected,
                        ),
                        bkt.ribbon.Button(
                            id='paste_and_distribute',
                            label="Text auf Auswahl verteilen",
                            supertip="Jeden Paragraphen (bzw. Zelle) aus der Zwischenablage einzeln auf die markierten Shapes verteilen (von links nach rechts, und von oben nach unten). Überflüssige Paragraphen werden verworfen.",
                            image_mso='PasteMergeList',
                            on_action=bkt.CallbackLazy("toolbox.models.copy_paste_format", "SlidesMore", "paste_and_distribute", slide=True, shapes=True),
                            get_enabled=bkt.apps.ppt_shapes_or_text_selected,
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
                            id="copy_texts",
                            label="Shape-Text kopieren",
                            supertip="Kopiert den Text aller markierten Shapes in die Zwischenablage.",
                            image_mso='DrawTextBox',
                            on_action=bkt.CallbackLazy("toolbox.models.copy_paste_format", "SlidesMore", "copy_texts", shapes=True),
                            get_enabled=bkt.get_enabled_auto
                        ),
                        bkt.ribbon.Button(
                            id="copy_slide_hq",
                            label="Folie als HQ-Bild kopieren",
                            supertip="Kopiert die aktuelle Folie in hoher Qualität in die Zwischenablage.",
                            image_mso='CopyPicture',
                            on_action=bkt.CallbackLazy("toolbox.models.copy_paste_format", "SlidesMore", "copy_in_highquality", slide=True),
                            get_enabled=bkt.get_enabled_auto
                        ),
                    ]
                )
            ]
        ),
        #bkt.mso.control.PasteSpecialDialog,
        #bkt.mso.control.Cut,
        #bkt.mso.control.CopySplitButton,
        
        bkt.ribbon.DynamicMenu(
            label='Auswahl',
            screentip='Auswahl von Shapes',
            supertip='Auswahl von Shapes, die dem aktuellem Shape bzgl. Typ/Hintergrund/Rahmen ähneln',
            show_label=False,
            image_mso='ObjectsMultiSelect',
            get_content=bkt.CallbackLazy("toolbox.models.shape_selection", "selection_menu"),
        ),
        
        bkt.mso.control.PasteApplyStyle,
        bkt.mso.control.PickUpStyle,
        bkt.ribbon.Button(
            id="select_by_fill",
            image_mso = 'ColorBlue',
            label='Auswahl von Shapes mit gleichem Hintergrund',
            show_label=False,
            on_action=bkt.CallbackLazy("toolbox.models.shape_selection", "ShapeSelector", "selectByFill", context=True),
            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
            screentip="Shape-Objekte mit gleichem Hintergrund markieren",
            supertip="Selektiere alle Shapes auf dem aktuellen Slide, die den gleichen Hintergrund (Farbe) haben wie eine der selektierten Shapes",
        ),


        bkt.mso.control.FormatPainter,
        bkt.ribbon.Button(
            id="format_syncer",
            label="Format Syncer",
            supertip="Alle Shapes so formatieren wie das zuerst ausgewählte Shape",
            image_mso="ShapeFillEffectMoreTexturesDialogClassic",
            show_label=False,
            get_enabled=bkt.apps.ppt_shapes_min2_selected,
            on_action=bkt.CallbackLazy("toolbox.models.copy_paste_format", "FormatPainter" "sync_shapes", shapes=True)
        ),
        bkt.ribbon.Button(
            id="select_by_border",
            image_mso = 'ColorWhite',
            label='Auswahl von Shapes mit gleichem Rahmen',
            show_label=False,
            on_action=bkt.CallbackLazy("toolbox.models.shape_selection", "ShapeSelector", "selectByLine", context=True),
            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
            screentip="Shape-Objekte mit gleichem Rahmen markieren",
            supertip="Selektiere alle Shapes auf dem aktuellen Slide, die den gleichen Rahmen (Farbe, Strichtyp) haben wie eine der selektierten Shapes",
        ),

        #dirty hack to show only one of the following two buttons:
        # bkt.ribbon.Box(get_visible=bkt.Callback(FormatPainter.fp_visible, context=True), children=[
        #     bkt.mso.control.FormatPainter
        # ]),
        
    ]
)


