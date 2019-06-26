# -*- coding: utf-8 -*-

# https://fontawesome.com/v4.7.0/

import bkt
from bkt.library.powerpoint import PPTSymbolsGallery

import os.path
import io
import json
from collections import OrderedDict,defaultdict


# define the menu parts

menu_title = 'Material Icons'

menu_settings = [
    # menu label,          list of symbols,       icons per row
    # ('IT-System',          symbols_itsystems,           6  ),
    # ('Kommunikation',      symbols_communication,       6  ),
    # ('Dateien',            symbols_files,               6  ),
    # ('Analyse/Bewertung',  symbols_analysis,            6  ),
    # ('Mixed',              symbols_mixed,               6  ),
    # ('All',                symbols_all,                16  )
]

def get_content_categories():
    # Automatically generate categories from json file from https://gist.github.com/AmirOfir/daee915574b1ba0d877da90777dc2181
    file = os.path.join(os.path.dirname(os.path.realpath(__file__)), "materialicons.json")
    with io.open(file, 'r') as json_file:
        chars = json.load(json_file, object_pairs_hook=OrderedDict)

    # categories = OrderedDict()
    categories = defaultdict(list)
    for char in chars['categories']:
        for ico in char['icons']:
            t=("Material Icons", unichr(int(ico['codepoint'], 16)), ico['name'])
            categories[char['name'].capitalize()].append(t)
    
    return bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None,
                children=[
                    PPTSymbolsGallery(label="{} ({})".format(cat, len(categories[cat])), symbols=categories[cat], columns=16)
                    for cat in sorted(categories.keys())
                ]
            )

menus = [
    PPTSymbolsGallery(label="{} ({})".format(label, len(symbollist)), symbols=symbollist, columns=columns)
    for (label, symbollist, columns) in menu_settings
] + [
    # submenu for categories
    bkt.ribbon.DynamicMenu(label="All Categories", get_content=bkt.Callback(get_content_categories))
]
