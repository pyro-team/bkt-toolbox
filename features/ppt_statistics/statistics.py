# -*- coding: utf-8 -*-
'''

@author: fstallmann
'''

from __future__ import absolute_import

import re #regex

import bkt
import bkt.library.powerpoint as pplib


class Statistics(object):
    do_refresh1=False
    do_refresh2=False
    do_refresh3=False

    @classmethod
    def shape_no(cls, shapes):
        if not cls.do_refresh1:
            return
        cls.do_refresh1 = False

        # return len(shapes)
        return len([shp for shp in pplib.iterate_shape_subshapes(shapes)])

    @classmethod
    def shape_sum(cls, shapes):
        if not cls.do_refresh2:
            return
        cls.do_refresh2 = False

        res_sum = 0
        for textframe in pplib.iterate_shape_textframes(shapes):
            res_sum += sum( list(cls.get_numbers_from_textframe(textframe)) )
        return res_sum
    
    @classmethod
    def shape_num_no(cls, shapes):
        if not cls.do_refresh3:
            return
        cls.do_refresh3 = False

        res_no = 0
        for textframe in pplib.iterate_shape_textframes(shapes):
            res_no += len( list(cls.get_numbers_from_textframe(textframe)) )
        return res_no

    @classmethod
    def get_numbers_from_textframe(cls, textframe):
        try:
            if textframe.TextRange.LanguageID == 1031: #DE
                regex = r'[^\-0-9,]'
            else: #EN, US, others
                regex = r'[^\-0-9.]'
            
            units = iter(textframe.TextRange.Words())
            for unit in units:
                try:
                    yield float(re.sub(regex, "", unit.Text).replace(",", "."))
                except ValueError:
                    # print(unit.Text)
                    continue
        except:
            pass

    @classmethod
    def reload(cls, context):
        cls.do_refresh1 = True
        cls.do_refresh2 = True
        cls.do_refresh3 = True
        try:
            context.python_addin.invalidate_ribbon()
        except:
            pass


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
                bkt.ribbon.EditBox(
                    # label=u"#",
                    image_mso='Repaginate',
                    screentip="Anzahl Shapes",
                    supertip="Zählt die Anzahl der ausgewählten Shapes (inkl. Shapes innerhalb von Gruppen bzw. Zellen innerhalb von Tabellen).",
                    size_string="#######",
                    get_text=bkt.Callback(Statistics.shape_no, shapes=True),
                    # get_enabled = bkt.Callback(lambda: getattr(Statistics, "do_refresh1")),
                ),
                bkt.ribbon.Button(
                    label="Neu laden",
                    show_label=False,
                    supertip="Statistiken für aktuelle Auswahl neu berechnen",
                    image_mso='AccessRefreshAllLists',
                    on_action=bkt.Callback(Statistics.reload, context=True),
                    # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
            ]
        ),
        bkt.ribbon.Box(
            box_style="horizontal",
            children=[
                bkt.ribbon.EditBox(
                    # label=u"1",
                    image_mso='NumberStyleGallery',
                    screentip="Anzahl Zahlen",
                    supertip="Zählt die Anzahl der erkennbaren Zahlen in der aktuellen Auswahl, die im Feld Summe aufsummiert angezeigt werden.",
                    size_string="#######",
                    get_text=bkt.Callback(Statistics.shape_num_no, shapes=True),
                    # get_enabled = bkt.Callback(lambda: getattr(Statistics, "do_refresh2")),
                ),
                bkt.ribbon.Button(
                    label="Neu laden",
                    show_label=False,
                    supertip="Statistiken für aktuelle Auswahl neu berechnen",
                    image_mso='AccessRefreshAllLists',
                    on_action=bkt.Callback(Statistics.reload, context=True),
                    # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
            ]
        ),
        bkt.ribbon.Box(
            box_style="horizontal",
            children=[
                bkt.ribbon.EditBox(
                    # label=u"\u2211",
                    image_mso='RecordsTotals',
                    screentip="Summe",
                    supertip="Summiert alle erkennbaren Zahlen in der aktuellen Auswahl. Je nach Sprache der Rechtschreibkorrektur wird Punkt oder Komma als Dezimaltrenner genommen. Negative Zahlen werden abgezogen.",
                    size_string="#######",
                    get_text=bkt.Callback(Statistics.shape_sum, shapes=True),
                    # get_enabled = bkt.Callback(lambda: getattr(Statistics, "do_refresh3")),
                ),
                bkt.ribbon.Button(
                    label="Neu laden",
                    show_label=False,
                    supertip="Statistiken für aktuelle Auswahl neu berechnen",
                    image_mso='AccessRefreshAllLists',
                    on_action=bkt.Callback(Statistics.reload, context=True),
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

