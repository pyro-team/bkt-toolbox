# -*- coding: utf-8 -*-
'''
Created on 26.02.2020

@author: fstallmann
'''

from __future__ import absolute_import

import bkt

from . import contextmenu_ids

from .common import common_groups
from .powerpoint import powerpoint_groups


bkt.powerpoint.add_tab(bkt.ribbon.Tab(
    idMso="TabDeveloper",
    # id="bkt_powerpoint_devkit",
    # label="BKT DEVKIT",
    # insert_before_mso="TabHome",
    children = common_groups + powerpoint_groups
), extend=True)


bkt.excel.add_tab(bkt.ribbon.Tab(
    idMso="TabDeveloper",
    # id="bkt_excel_devkit",
    # label="BKT DEVKIT",
    # insert_before_mso="TabHome",
    children = common_groups
), extend=True)


bkt.word.add_tab(bkt.ribbon.Tab(
    idMso="TabDeveloper",
    # id="bkt_word_devkit",
    # label="BKT DEVKIT",
    # insert_before_mso="TabHome",
    children = common_groups
), extend=True)


bkt.visio.add_tab(bkt.ribbon.Tab(
    idMso="TabDeveloper",
    # id="bkt_visio_devkit",
    # label="BKT DEVKIT",
    # insert_before_mso="TabHome",
    children = common_groups
), extend=True)