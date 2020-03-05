# -*- coding: utf-8 -*-

# https://docs.microsoft.com/de-de/windows/uwp/design/style/segoe-ui-symbol-font

import bkt
from bkt.library.powerpoint import PPTSymbolsGallery

import os.path
import io
import json

font_list = ["Wingdings", "Wingdings 2", "Wingdings 3", "Webdings"]
all_fonts = {
    'Wingdings': [],
    'Wingdings 2': [],
    'Wingdings 3': [],
    'Webdings': [],
}


file = os.path.join(os.path.dirname(os.path.realpath(__file__)), "wingdings.json")
with io.open(file, 'r', encoding='utf-8') as json_file:
    chars = json.load(json_file)

    for ico in chars:
        all_fonts[ico["font"]].append(
            (ico["font"], unichr(ico['dec']), ico["name"])
        )


def update_search_index(search_engine):
    search_writer = search_engine.writer()

    for font in font_list:
        for ico in all_fonts[font]:
            search_writer.add_document(
                module="wingdings",
                fontlabel=ico[0],
                fontname=ico[0],
                unicode=ico[1],
                label=ico[2],
                keywords=ico[2].lower().split()
            )

    search_writer.commit()


# define the menu parts
menu_title = "Wingdings"

menus = [
    PPTSymbolsGallery(label="{} ({})".format(font, len(all_fonts[font])),     symbols=all_fonts[font], columns=16)
    for font in font_list
    # PPTSymbolsGallery(label="Wingdings ({})".format(len(all_fonts["Wingdings"])),     symbols=all_fonts["Wingdings"], columns=16),
    # PPTSymbolsGallery(label="Wingdings 2 ({})".format(len(all_fonts["Wingdings 2"])), symbols=all_fonts["Wingdings 2"], columns=16),
    # PPTSymbolsGallery(label="Wingdings 3 ({})".format(len(all_fonts["Wingdings 3"])), symbols=all_fonts["Wingdings 3"], columns=16),
    # PPTSymbolsGallery(label="Webdings ({})".format(len(all_fonts["Webdings"])),       symbols=all_fonts["Webdings"], columns=16),
]


