# -*- coding: utf-8 -*-
'''

@author: fstallmann
'''

from __future__ import absolute_import

import re #regex
import locale #to format number for clipboard and supertip

import bkt.dotnet as dotnet
Forms = dotnet.import_forms() #required to copy text to clipboard

import bkt
import bkt.library.powerpoint as pplib


class Statistics(object):
    comma_langs = {
        1031: "Deutsch",
        3079: "Deutsch (Österreich)",
        1040: "Italienisch",
        1036: "Französisch",
        3082: "Spanisch",
        1049: "Russisch",
        1029: "Tschechisch",
        1030: "Dänisch",
        1043: "Holländisch",
        1045: "Polnisch",
        2070: "Portugisisch",
        1053: "Schwedisch",
        1055: "Türkisch",
    }
    dot_langs = {
        1033: "US English",
        2057: "UK English",
    }
    _refresh1 = False
    _refresh2 = False
    _refresh3 = False

    _res1 = None
    _res2 = None
    _res3 = None

    _list_shapes = []
    _list_numbers = []

    @classmethod
    def shape_no(cls):
        if not cls._refresh1:
            return cls._res1
        cls._refresh1 = False

        # cls._res1 = len([shp for shp in pplib.iterate_shape_subshapes(shapes)])
        cls._res1 = len(cls._list_shapes)
        return cls._res1
    
    @classmethod
    def shape_num_no(cls):
        if not cls._refresh2:
            return cls._res2
        cls._refresh2 = False

        cls._res2 = len(cls._list_numbers)
        return cls._res2

        # res_no = 0
        # for textframe in pplib.iterate_shape_textframes(shapes):
        #     res_no += len( list(cls.get_numbers_from_textframe(textframe)) )
        # return res_no

    @classmethod
    def shape_sum(cls):
        if not cls._refresh3:
            return cls._res3
        cls._refresh3 = False

        cls._res3 = sum(cls._list_numbers)
        return cls._res3

        # res_sum = 0
        # for textframe in pplib.iterate_shape_textframes(shapes):
        #     res_sum += sum( list(cls.get_numbers_from_textframe(textframe)) )
        # return res_sum

    @classmethod
    def get_numbers_from_textframe(cls, textframe):
        try:
            if textframe.TextRange.LanguageID in cls.comma_langs.keys():
                regex = r'[^\-0-9,]'
            else: #EN, US, others
                regex = r'[^\-0-9.]'
            
            # units = iter(textframe.TextRange.Words()) #issue: splits -X in 2 words: [-, X]
            units = re.split(r'[\s;]', textframe.TextRange.Text)
            for unit in units:
                try:
                    # yield float(re.sub(regex, "", unit.Text).replace(",", "."))
                    yield float(re.sub(regex, "", unit).replace(",", "."))
                except ValueError:
                    # print(unit.Text)
                    continue
        except:
            pass

    @classmethod
    def reload_all(cls, shapes):
        def get_shp_name(shp):
            try:
                return shp.Name
            except:
                return "<No name, maybe table cell>"
        cls._list_shapes = [get_shp_name(shp) for shp in pplib.iterate_shape_subshapes(shapes)]
        cls._list_numbers = []
        for textframe in pplib.iterate_shape_textframes(shapes):
            cls._list_numbers.extend(cls.get_numbers_from_textframe(textframe))
        
        cls._refresh1 = True
        cls._refresh2 = True
        cls._refresh3 = True

        #button click already invalidates ribbon, no need to extra invalidate
        # try:
        #     context.python_addin.invalidate_ribbon()
        # except:
        #     pass
    
    @classmethod
    def copy_no(cls):
        if cls._res1:
            Forms.Clipboard.SetText(locale.format("%d",cls._res1))
            bkt.message("Anzahl Shapes in Zwischenablage kopiert")
        
    @classmethod
    def copy_num_no(cls):
        if cls._res2:
            Forms.Clipboard.SetText(locale.format("%d",cls._res2))
            bkt.message("Anzahl Zahlen in Zwischenablage kopiert")
        
    @classmethod
    def copy_sum(cls):
        if cls._res3:
            Forms.Clipboard.SetText(locale.format("%f",cls._res3))
            bkt.message("Summe in Zwischenablage kopiert")

    @classmethod
    def get_supertip_no(cls):
        if len(cls._list_shapes) > 30:
            part1 = "\n".join(cls._list_shapes[:15])
            part2 = "\n".join(cls._list_shapes[-15:])
            return part1 + "\n\n...\n\n" + part2
        else:
            return "\n".join(cls._list_shapes)

    @classmethod
    def get_supertip_num_no(cls):
        if len(cls._list_numbers) > 30:
            part1 = "\n".join("{:.15n}".format(n) for n in cls._list_numbers[:15])
            part2 = "\n".join("{:.15n}".format(n) for n in cls._list_numbers[-15:])
            return part1 + "\n\n...\n\n" + part2
        else:
            return "\n".join("{:.15n}".format(n) for n in cls._list_numbers)

    @classmethod
    def get_supertip_sum(cls):
        if len(cls._list_numbers) > 30:
            part1 = "\n+ ".join("{:.15n}".format(n) for n in cls._list_numbers[:15])
            part2 = "\n+ ".join("{:.15n}".format(n) for n in cls._list_numbers[-15:])
            return part1 + "\n\n...\n\n+ " + part2
        else:
            return "\n+ ".join("{:.15n}".format(n) for n in cls._list_numbers)


statistik_gruppe = bkt.ribbon.Group(
    id="bkt_statistics_group",
    label='Statistik',
    supertip="Ermöglicht die Anzeige einfacher Statistiken zur schnellen Überprüfung von zahlenlastigen Folien. Das Feature `ppt_statistics` muss installiert sein.",
    image_mso='RecordsTotals',
    children = [
        # bkt.ribbon.LabelControl(
        #     label="Statistiken für Auswahl"
        # ),
        bkt.ribbon.Box(
            box_style="horizontal",
            children=[
                bkt.ribbon.Button(
                    label="Anzahl Shapes",
                    image_mso='Repaginate',
                    show_label=False,
                    supertip="Zählt die Anzahl der ausgewählten Shapes (inkl. Shapes innerhalb von Gruppen bzw. Zellen innerhalb von Tabellen).",
                    on_action=bkt.Callback(Statistics.copy_no),
                ),
                bkt.ribbon.Box(children=[
                    #innerbox required to avoid space before EditBox
                    bkt.ribbon.EditBox(
                        # label=u"#",
                        show_label=False,
                        screentip="Ergebnis Anzahl Shapes",
                        get_supertip=bkt.Callback(Statistics.get_supertip_no),
                        size_string="#######",
                        get_text=bkt.Callback(Statistics.shape_no),
                        # get_enabled = bkt.Callback(lambda: getattr(Statistics, "_refresh1")),
                    ),
                ]),
                bkt.ribbon.Button(
                    label="Neu laden",
                    show_label=False,
                    supertip="Statistiken für aktuelle Auswahl neu berechnen",
                    image_mso='AccessRefreshAllLists',
                    on_action=bkt.Callback(Statistics.reload_all, shapes=True),
                    # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
            ]
        ),
        bkt.ribbon.Box(
            box_style="horizontal",
            children=[
                bkt.ribbon.Button(
                    label="Anzahl Zahlen",
                    image_mso='NumberStyleGallery',
                    show_label=False,
                    supertip="Zählt die Anzahl der erkennbaren Zahlen in der aktuellen Auswahl, die im Feld Summe aufsummiert angezeigt werden.",
                    on_action=bkt.Callback(Statistics.copy_num_no),
                ),
                bkt.ribbon.Box(children=[
                    #innerbox required to avoid space before EditBox
                    bkt.ribbon.EditBox(
                        # label=u"1",
                        show_label=False,
                        screentip="Ergebnis für Anzahl Zahlen",
                        get_supertip=bkt.Callback(Statistics.get_supertip_num_no),
                        size_string="#######",
                        get_text=bkt.Callback(Statistics.shape_num_no),
                        # get_enabled = bkt.Callback(lambda: getattr(Statistics, "_refresh2")),
                    ),
                ]),
                bkt.ribbon.Button(
                    label="Neu laden",
                    show_label=False,
                    supertip="Statistiken für aktuelle Auswahl neu berechnen",
                    image_mso='AccessRefreshAllLists',
                    on_action=bkt.Callback(Statistics.reload_all, shapes=True),
                    # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
            ]
        ),
        bkt.ribbon.Box(
            box_style="horizontal",
            children=[
                bkt.ribbon.Button(
                    label="Summe",
                    image_mso='RecordsTotals',
                    show_label=False,
                    supertip="Summiert alle erkennbaren Zahlen in der aktuellen Auswahl. Je nach Sprache der Rechtschreibkorrektur wird Punkt oder Komma als Dezimaltrenner genommen. Negative Zahlen werden abgezogen.",
                    on_action=bkt.Callback(Statistics.copy_sum),
                ),
                bkt.ribbon.Box(children=[
                    #innerbox required to avoid space before EditBox
                    bkt.ribbon.EditBox(
                        # label=u"\u2211",
                        show_label=False,
                        screentip="Ergebnis für Summe",
                        get_supertip=bkt.Callback(Statistics.get_supertip_sum),
                        size_string="#######",
                        get_text=bkt.Callback(Statistics.shape_sum),
                        # get_enabled = bkt.Callback(lambda: getattr(Statistics, "_refresh3")),
                    ),
                ]),
                bkt.ribbon.Button(
                    label="Neu laden",
                    show_label=False,
                    supertip="Statistiken für aktuelle Auswahl neu berechnen",
                    image_mso='AccessRefreshAllLists',
                    on_action=bkt.Callback(Statistics.reload_all, shapes=True),
                    # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
            ]
        )
    ]
)


bkt.powerpoint.add_tab(bkt.ribbon.Tab(
    id="bkt_powerpoint_toolbox_extensions",
    insert_before_mso="TabHome",
    label=u'Toolbox 3/3',
    # get_visible defaults to False during async-startup
    get_visible=bkt.Callback(lambda:True),
    children = [
        statistik_gruppe
    ]
), extend=True)

