# -*- coding: utf-8 -*-
'''

@author: fstallmann
'''

from __future__ import absolute_import

import bkt
import bkt.library.powerpoint as pplib


class Statistics(object):
    
    @staticmethod
    def show_dialog(context):
        from .dialog import StatisticsWindow
        # StatisticsWindow.create_and_show_dialog(context)
        dialog = StatisticsWindow(context)
        dialog.show_dialog(False)


statistik_gruppe = bkt.ribbon.Group(
    id="bkt_statistics_group",
    label='Statistik',
    supertip="Ermöglicht die Anzeige einfacher Statistiken zur schnellen Überprüfung von zahlenlastigen Folien. Das Feature `ppt_statistics` muss installiert sein.",
    image_mso='RecordsTotals',
    children = [
        bkt.ribbon.Button(
            label="Statistik laden",
            image_mso='RecordsTotals',
            size="large",
            supertip="Öffnet einen Dialog zur Anzeige der Anzahl der markierten Shapes, Zahlen, Summe der Zahlen, Anzahl Zeichen, Wörter, Zeilen und Absätze.",
            on_action=bkt.Callback(Statistics.show_dialog),
        ),
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

