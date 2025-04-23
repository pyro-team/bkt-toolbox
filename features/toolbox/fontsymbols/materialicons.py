# -*- coding: utf-8 -*-

# https://fontawesome.com/v4.7.0/



import os.path
import io
import json
from collections import OrderedDict,defaultdict

import bkt
from bkt.library.powerpoint import PPTSymbolsGallery


# define the menu parts

menu_title = 'Material Icons'

symbols_common = [
    ("Material Icons", "\uE853", "account circle", "Material Icons > Wichtige"),
    ("Material Icons", "\uE897", "lock", "Material Icons > Wichtige"),
    ("Material Icons", "\uE8DC", "thumbs up", "Material Icons > Wichtige"),
    ("Material Icons", "\uE8DB", "thumbs down", "Material Icons > Wichtige"),
    ("Material Icons", "\uE0B0", "call", "Material Icons > Wichtige"),
    ("Material Icons", "\uE0B7", "chat", "Material Icons > Wichtige"),
    ("Material Icons", "\uE0BE", "email", "Material Icons > Wichtige"),
    ("Material Icons", "\uE2BD", "cloud", "Material Icons > Wichtige"),
    ("Material Icons", "\uE7EF", "group", "Material Icons > Wichtige"),
    ("Material Icons", "\uE7FD", "person", "Material Icons > Wichtige"),
    ("Material Icons", "\uE55F", "place", "Material Icons > Wichtige"),
    ("Material Icons", "\uE80B", "public", "Material Icons > Wichtige"),
]

menu_settings = [
    # menu label,          list of symbols,       icons per row
    ('Wichtige',          symbols_common,           6  ),
]

# full font names
font_names = [
    "Material Icons",
    # "Material Icons Outlined",
    # "Material Icons Rounded",
    # "Material Icons Sharp"
    ]

cache_menu = {}

def get_content_categories(current_control):
    global cache_menu

    font = current_control["tag"]

    if font in cache_menu:
        return cache_menu[font]

    # Automatically generate categories from json file (based on yaml file provided by fontawesome)
    file = os.path.join(os.path.dirname(os.path.realpath(__file__)), "materialicons.json")
    with io.open(file, 'r', encoding='utf-8') as json_file:
        chars = json.load(json_file, object_pairs_hook=OrderedDict)

    # categories = OrderedDict()
    categories = defaultdict(list)
    for char in chars['icons']:
        if font in char['unsupported_families']:
            continue
        for cat in char['categories']:
            supertip = "{} > {}\n{}".format(font, cat, ", ".join(char['tags']))
            categories[cat].append((font, chr(int(char['codepoint'])), char['name'].capitalize(), supertip))
    
    cache_menu[font] = bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None,
                children=[
                    PPTSymbolsGallery(label="{} ({})".format(cat.capitalize(), len(categories[cat])), symbols=categories[cat], columns=16)
                    for cat in sorted(categories.keys())
                ]
            )
    return cache_menu[font]

def update_search_index(search_engine):
    search_writer = search_engine.writer()

    # Automatically generate categories from json file from https://gist.github.com/AmirOfir/daee915574b1ba0d877da90777dc2181
    file = os.path.join(os.path.dirname(os.path.realpath(__file__)), "materialicons.json")
    with io.open(file, 'r', encoding='utf-8') as json_file:
        chars = json.load(json_file, object_pairs_hook=OrderedDict)
        
        for font in font_names:
            for char in chars['icons']:
                if font in char['unsupported_families']:
                    continue

                search_writer.add_document(
                    module="materialicons",
                    fontlabel=font,
                    fontname=font,
                    unicode=chr(int(char['codepoint'])),
                    label="{} > {}".format(font, char['name'].capitalize()),
                    keywords=char['tags']
                )
    
    search_writer.commit()


menus = [
    PPTSymbolsGallery(label="{} ({})".format(label, len(symbollist)), symbols=symbollist, columns=columns)
    for (label, symbollist, columns) in menu_settings
] + [
    # submenu for categories
    # bkt.ribbon.DynamicMenu(label="All Categories", get_content=bkt.Callback(get_content_categories))
    bkt.ribbon.DynamicMenu(label=font, tag=font, get_content=bkt.Callback(get_content_categories))
    for font in font_names
]
