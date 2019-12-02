# -*- coding: utf-8 -*-

# https://fontawesome.com

import bkt
from bkt.library.powerpoint import PPTSymbolsGallery

import os.path
import io
import json
from collections import OrderedDict


### How to get json files?
# The font awesome archive contains metadata/categories.yml and metadata/icons.yml
# Use https://www.json2yaml.com/ to convert yml to json
# DO NOT USE metadata/icons.json!
###


# full font names
font_name_hash = {
    'regular': "Font Awesome 5 Free Regular",
    'solid':   "Font Awesome 5 Free Solid",
    'brands':  "Font Awesome 5 Brands Regular"
}

all_fonts = {
    'regular': [],
    'solid': [],
    'brands': []
}

file = os.path.join(os.path.dirname(os.path.realpath(__file__)), "fa5-icons.json")
with io.open(file, 'r') as json_file:
    all_icons = json.load(json_file, object_pairs_hook=OrderedDict)

    for label, icon in all_icons.iteritems():
        for font in icon["styles"]:
            symbol = (
                font_name_hash[font],
                unichr(int(icon['unicode'], 16)),
                icon["label"],
                "{}\n{}".format(font_name_hash[font], ", ".join(icon["search"]["terms"]))
            )
            all_fonts[font].append(symbol)


cache_menu = None

def get_content_categories():
    global cache_menu

    if cache_menu:
        return cache_menu
    
    categories = []
    file = os.path.join(os.path.dirname(os.path.realpath(__file__)), "fa5-categories.json")
    with io.open(file, 'r') as json_file:
        cats = json.load(json_file, object_pairs_hook=OrderedDict)

        for key, value in cats.iteritems():
            catname = value["label"].replace("&", "and")
            symbols = []
            for ico in value["icons"]:
                icon = all_icons[ico]
                for font in icon["styles"]:
                    symbol = (
                        font_name_hash[font],
                        unichr(int(icon['unicode'], 16)),
                        icon["label"],
                        "{} > {}\n{}".format(font_name_hash[font], catname, ", ".join(icon["search"]["terms"]))
                    )
                    symbols.append(symbol)

            categories.append(
                PPTSymbolsGallery(
                    label="{} ({})".format(catname,len(symbols)),
                    symbols=symbols,
                    columns=16
                )
            )

    cache_menu = bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None,
                children=categories
            )
    return cache_menu


    # # hash for fa5-icons
    # # { icon_name: symbol-gallery-item, ...}
    # # symbol-gallery-item = [fontname, character-id, label]
    # fa5_icons_hash = {
    #     icon_name: [ [ font_name_hash[style], uid, "%s, unicode character %s, %s" % (font_name_hash[style], format(ord(uid), '02x'), icon_name)  ] for style in styles]
    #     for (icon_name, uid, styles) in fa5_icons
    # }


# define the menu parts

menu_title = 'Font Awesome 5 Free'

menu_settings = [
    # menu label,          list of symbols,       icons per row
    ('All Regular',            all_fonts['regular'],          16  ),
    ('All Solid',              all_fonts['solid'],            16  ),
    ('All Brands',             all_fonts['brands'],           16  ),
]

menus = [
    PPTSymbolsGallery(label="{} ({})".format(label, len(symbollist)), symbols=symbollist, columns=columns)
    for (label, symbollist, columns) in menu_settings
] + [
    # submenu for categories
    bkt.ribbon.DynamicMenu(label="All Categories", get_content=bkt.Callback(get_content_categories))
]

