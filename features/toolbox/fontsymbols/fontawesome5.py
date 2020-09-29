# -*- coding: utf-8 -*-

# https://fontawesome.com

from __future__ import absolute_import

import os.path
import io
import json
from collections import OrderedDict

import bkt
from bkt.library.powerpoint import PPTSymbolsGallery


### How to get json files?
# The font awesome archive on https://github.com/FortAwesome/Font-Awesome/releases contains metadata/categories.yml and metadata/icons.yml
# Use https://www.json2yaml.com/ to convert yml to json
# DO NOT USE metadata/icons.json!
###

version_of_fontawesome_json = "5.15.0"
menu_title = 'Font Awesome 5 Free v' + version_of_fontawesome_json

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
with io.open(file, 'r', encoding='utf-8') as json_file:
    all_icons = json.load(json_file, object_pairs_hook=OrderedDict)

    for _, icon in all_icons.iteritems():
        for font in icon["styles"]:
            symbol = (
                font_name_hash[font],
                unichr(int(icon['unicode'], 16)),
                icon["label"],
                "{}\n{}".format(font_name_hash[font], ", ".join(icon["search"]["terms"]))
            )
            all_fonts[font].append(symbol)

#cache for category menu
cache_menu = None

def get_content_categories():
    global cache_menu

    if cache_menu:
        return cache_menu
    
    categories = []
    file = os.path.join(os.path.dirname(os.path.realpath(__file__)), "fa5-categories.json")
    with io.open(file, 'r', encoding='utf-8') as json_file:
        cats = json.load(json_file, object_pairs_hook=OrderedDict)

        for _, value in cats.iteritems():
            catname = value["label"]
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

def update_search_index(search_engine):
    search_writer = search_engine.writer()
    full_icon_infos = {
        'regular': OrderedDict(),
        'solid': OrderedDict(),
        'brands': OrderedDict(),
    }

    def _add_icon(ident, font, unicode, label, keywords):
        full_icon_infos[font][ident] = {
            "module":    "fontawesome5",
            "fontlabel": font_name_hash[font],
            "fontname":  font_name_hash[font],
            "unicode":   unichr(int(unicode, 16)),
            "label":     label,
            "keywords":  set(keywords+label.lower().split()),
        }
    
    #first add all icons, as not all icons are part of a category
    for ident, icon in all_icons.iteritems():
        for font in icon["styles"]:
            _add_icon(ident, font, icon['unicode'], icon["label"], icon["search"]["terms"])

    #second consolidate category names into keywords
    file = os.path.join(os.path.dirname(os.path.realpath(__file__)), "fa5-categories.json")
    with io.open(file, 'r', encoding='utf-8') as json_file:
        cats = json.load(json_file, object_pairs_hook=OrderedDict)

        for _, value in cats.iteritems():
            for ident in value["icons"]:
                for font in all_icons[ident]["styles"]:
                    full_icon_infos[font][ident]["keywords"].update( value["label"].lower().replace("&", " ").split() )

    for icon in full_icon_infos['regular'].itervalues():
        search_writer.add_document(**icon)
    for icon in full_icon_infos['solid'].itervalues():
        search_writer.add_document(**icon)
    for icon in full_icon_infos['brands'].itervalues():
        search_writer.add_document(**icon)
    search_writer.commit()


    # # hash for fa5-icons
    # # { icon_name: symbol-gallery-item, ...}
    # # symbol-gallery-item = [fontname, character-id, label]
    # fa5_icons_hash = {
    #     icon_name: [ [ font_name_hash[style], uid, "%s, unicode character %s, %s" % (font_name_hash[style], format(ord(uid), '02x'), icon_name)  ] for style in styles]
    #     for (icon_name, uid, styles) in fa5_icons
    # }


# define the menu parts
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

