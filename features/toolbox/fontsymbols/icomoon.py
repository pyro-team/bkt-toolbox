# -*- coding: utf-8 -*-

# https://fontawesome.com/v4.7.0/



import os.path
import io
import json
from collections import OrderedDict
from itertools import chain

import bkt
from bkt.library.powerpoint import PPTSymbolsGallery



file = os.path.join(os.path.dirname(os.path.realpath(__file__)), "icomoon-free.json")
with io.open(file, 'r', encoding='utf-8') as json_file:
    icons = json.load(json_file, object_pairs_hook=OrderedDict)
    icons = icons["icons"]

    symbols1 = []
    symbols2 = []

    i = 0
    for icon in icons:
        if not "tags" in icon["icon"]:
            continue
        if i < 250:
            symbols1.append(("IcoMoon-Free", chr(int(icon['properties']['code'])), icon['properties']['name'], ", ".join(icon['icon']['tags'])))
        else:
            symbols2.append(("IcoMoon-Free", chr(int(icon['properties']['code'])), icon['properties']['name'], ", ".join(icon['icon']['tags'])))
        i += 1


def update_search_index(search_engine):
    search_writer = search_engine.writer()

    for char in chain(symbols1, symbols2):
        search_writer.add_document(
            module="icomoon",
            fontlabel="IcoMoon Free",
            fontname=char[0],
            unicode=char[1],
            label=char[2],
            # keywords=search_writer.get_keywords_from_string(char[3])
            keywords=char[3].split(", ")
        )
    search_writer.commit()


# define the menu parts
menu_title = 'IcoMoon Free'

menus = [
    PPTSymbolsGallery(label="Seite 1 ({})".format(len(symbols1)), symbols=symbols1, columns=16),
    PPTSymbolsGallery(label="Seite 2 ({})".format(len(symbols2)), symbols=symbols2, columns=16)
]
