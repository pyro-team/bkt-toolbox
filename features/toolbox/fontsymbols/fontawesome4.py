# -*- coding: utf-8 -*-

# https://fontawesome.com/v4.7.0/



import os.path
import io
import json
from collections import OrderedDict,defaultdict

import bkt
from bkt.library.powerpoint import PPTSymbolsGallery


symbols_communication = [
    ("FontAwesome", "\uf2ba", "address book", "Font Awesome 4"),
    ("FontAwesome", "\uf2bc", "address card", "Font Awesome 4"),
    ("FontAwesome", "\uf2c1", "id badge", "Font Awesome 4"),
    ("FontAwesome", "\uf2c3", "id card", "Font Awesome 4"),
    ("FontAwesome", "\uf183", "man", "Font Awesome 4"),
    ("FontAwesome", "\uf0c0", "users", "Font Awesome 4"),
    ("FontAwesome", "\uf2be", "user circle", "Font Awesome 4"),
    ("FontAwesome", "\uf2c0", "user", "Font Awesome 4"),
    ("FontAwesome", "\uf007", "user black", "Font Awesome 4"),
    ("FontAwesome", "\uf2b5", "handshake", "Font Awesome 4"),
    ("FontAwesome", "\uf0e5", "comment", "Font Awesome 4"),
    ("FontAwesome", "\uf27b", "commenting", "Font Awesome 4"),
    ("FontAwesome", "\uf0e6", "comments", "Font Awesome 4"),
    ("FontAwesome", "\uf086", "comments", "Font Awesome 4"),
]

symbols_itsystems = [
    ("FontAwesome", "\uf108", "desktop", "Font Awesome 4"),
    ("FontAwesome", "\uf109", "laptop", "Font Awesome 4"),
    ("FontAwesome", "\uf10a", "tablet", "Font Awesome 4"),
    ("FontAwesome", "\uf10b", "mobile", "Font Awesome 4"),
    ("FontAwesome", "\uf095", "phone", "Font Awesome 4"),
    ("FontAwesome", "\uf1ac", "fax", "Font Awesome 4"),
    ("FontAwesome", "\uf003", "mail", "Font Awesome 4"),
    ("FontAwesome", "\uf01c", "inbox", "Font Awesome 4"),
    ("FontAwesome", "\uf11c", "keyboard", "Font Awesome 4"),
    ("FontAwesome", "\uf0c2", "cloud", "Font Awesome 4"),
    ("FontAwesome", "\uf09e", "rss", "Font Awesome 4"),
    ("FontAwesome", "\uf1eb", "wifi", "Font Awesome 4"),
    ("FontAwesome", "\uf090", "sign in", "Font Awesome 4"),
    ("FontAwesome", "\uf084", "key", "Font Awesome 4"),
    ("FontAwesome", "\uf023", "lock", "Font Awesome 4"),
    ("FontAwesome", "\uf09c", "unlock", "Font Awesome 4"),
    ("FontAwesome", "\uf13e", "unlock", "Font Awesome 4"),
    ("FontAwesome", "\uf132", "shield", "Font Awesome 4"),
    ("FontAwesome", "\uf19c", "university/bank", "Font Awesome 4"),
    ("FontAwesome", "\uf015", "home", "Font Awesome 4"),
    ("FontAwesome", "\uf1ad", "building", "Font Awesome 4"),
    ("FontAwesome", "\uf1b2", "cube", "Font Awesome 4"),
    ("FontAwesome", "\uf1b3", "cubes", "Font Awesome 4"),
    ("FontAwesome", "\uf1c0", "database", "Font Awesome 4"),
    ("FontAwesome", "\uf233", "server", "Font Awesome 4"),
    ("FontAwesome", "\uf2db", "microchip", "Font Awesome 4"),
    ("FontAwesome", "\uf188", "bug", "Font Awesome 4"),
]

symbols_files = [
    ("FontAwesome", "\uf114", "folder", "Font Awesome 4"),
    ("FontAwesome", "\uf115", "folder open", "Font Awesome 4"),
    ("FontAweSome", "\uf016", "file", "Font Awesome 4"),
    ("FontAweSome", "\uf0c5", "files", "Font Awesome 4"),
    ("FontAweSome", "\uf0f6", "text", "Font Awesome 4"),
    ("FontAweSome", "\uf1c9", "code", "Font Awesome 4"),
    ("FontAweSome", "\uf1c6", "archive", "Font Awesome 4"),
    ("FontAweSome", "\uf1c5", "image", "Font Awesome 4"),
    ("FontAweSome", "\uf1c8", "video", "Font Awesome 4"),
    ("FontAweSome", "\uf1c7", "audio", "Font Awesome 4"),
    ("FontAweSome", "\uf1c1", "pdf", "Font Awesome 4"),
    ("FontAweSome", "\uf1c3", "excel", "Font Awesome 4"),
    ("FontAweSome", "\uf1c4", "powerpoint", "Font Awesome 4"),
    ("FontAweSome", "\uf1c2", "word", "Font Awesome 4"),
]

symbols_analysis = [
    ("FontAwesome", "\uf080", "bar chart", "Font Awesome 4"),
    ("FontAwesome", "\uf200", "pie chart", "Font Awesome 4"),
    ("FontAwesome", "\uf201", "line chart", "Font Awesome 4"),
    ("FontAwesome", "\uf1fe", "area chart", "Font Awesome 4"),
    ("FontAwesome", "\uf02a", "barcode", "Font Awesome 4"),
    ("FontAwesome", "\uf24e", "balance", "Font Awesome 4"),

    ("FontAwesome", "\uf005", "star", "Font Awesome 4"),
    ("FontAwesome", "\uf123", "star half", "Font Awesome 4"),
    ("FontAwesome", "\uf006", "star empty", "Font Awesome 4"),
    ("FontAwesome", "\uf251", "hourglass start", "Font Awesome 4"),
    ("FontAwesome", "\uf252", "hourglass half", "Font Awesome 4"),
    ("FontAwesome", "\uf253", "hourglass end", "Font Awesome 4"),

    ("FontAwesome", "\uf244", "battery empty", "Font Awesome 4"),
    ("FontAwesome", "\uf243", "battery 1/4", "Font Awesome 4"),
    ("FontAwesome", "\uf242", "battery 1/2", "Font Awesome 4"),
    ("FontAwesome", "\uf241", "battery 3/4", "Font Awesome 4"),
    ("FontAwesome", "\uf240", "battery full", "Font Awesome 4"),
    ("FontAwesome", "\uf05e", "ban", "Font Awesome 4"),

    ("FontAwesome", "\uf087", "thumbs up", "Font Awesome 4"),
    ("FontAwesome", "\uf088", "thumbs down", "Font Awesome 4"),
    ("FontAwesome", "\uf164", "thumbs up", "Font Awesome 4"),
    ("FontAwesome", "\uf165", "thumbs down", "Font Awesome 4"),

    ("FontAwesome", "\uf046", "check", "Font Awesome 4"),
    ("FontAwesome", "\uf05d", "check circle", "Font Awesome 4"),
    ("FontAwesome", "\uf00c", "check", "Font Awesome 4"),
    ("FontAwesome", "\uf00d", "x", "Font Awesome 4"),
]


symbols_mixed = [
    ("FontAwesome", "\uf0d0", "magic", "Font Awesome 4"),
    ("FontAwesome", "\uf02d", "book", "Font Awesome 4"),
    ("FontAwesome", "\uf02e", "bookmark", "Font Awesome 4"),
    ("FontAwesome", "\uf0b1", "briefcase", "Font Awesome 4"),
    ("FontAwesome", "\uf140", "bullseye", "Font Awesome 4"),
    ("FontAwesome", "\uf073", "calendar", "Font Awesome 4"),
    ("FontAwesome", "\uf0a3", "certificate", "Font Awesome 4"),
    ("FontAwesome", "\uf017", "clock", "Font Awesome 4"),
    ("FontAwesome", "\uf013", "cog", "Font Awesome 4"),
    ("FontAwesome", "\uf085", "cogs", "Font Awesome 4"),
    ("FontAwesome", "\uf134", "fire extinguisher", "Font Awesome 4"),
    ("FontAwesome", "\uf277", "map signs", "Font Awesome 4"),
    ("FontAwesome", "\uf041", "map marker", "Font Awesome 4"),
    ("FontAwesome", "\uf1ea", "newspaper", "Font Awesome 4"),
    ("FontAwesome", "\uf0eb", "lightbulb", "Font Awesome 4"),
    ("FontAwesome", "\uf08d", "pin", "Font Awesome 4"),
    ("FontAwesome", "\uf074", "random", "Font Awesome 4"),
    ("FontAwesome", "\uf1b8", "recycle", "Font Awesome 4"),
    ("FontAwesome", "\uf021", "refresh", "Font Awesome 4"),
    ("FontAwesome", "\uf135", "rocket", "Font Awesome 4"),
    ("FontAwesome", "\uf002", "search", "Font Awesome 4"),
    ("FontAwesome", "\uf0e4", "tachometer", "Font Awesome 4"),
    ("FontAwesome", "\uf02b", "tag", "Font Awesome 4"),
    ("FontAwesome", "\uf02c", "tags", "Font Awesome 4"),
    ("FontAwesome", "\uf014", "trash", "Font Awesome 4"),
    ("FontAwesome", "\uf0ad", "wrench", "Font Awesome 4"),
]


# # create gallery for all symbols
# # symbols in uid-range, f000 to f299
# hex_start_end = ('f000', 'f300')
# int_start_end = (int(x, 16) for x in hex_start_end)
# uids = range(*int_start_end)

# # [ symbol-gallery-item, symbol-gallery-item, ... ]
# # symbol-gallery-item = [fontname, character-id, label]
# symbols_all = [ 
#     [
#         'FontAwesome',
#         unichr(uid),
#         'FontAwesome, unicode character %s' % format(uid, '02x')
#     ]
#     for uid in uids
# ]


# define the menu parts

menu_title = 'Font Awesome 4.7'

menu_settings = [
    # menu label,          list of symbols,       icons per row
    ('IT-System',          symbols_itsystems,           6  ),
    ('Kommunikation',      symbols_communication,       6  ),
    ('Dateien',            symbols_files,               6  ),
    ('Analyse/Bewertung',  symbols_analysis,            6  ),
    ('Mixed',              symbols_mixed,               6  ),
    # ('All',                symbols_all,                16  )
]

# menus = [
#     PPTSymbolsGallery(label=cat, symbols=categories[cat], columns=16)
#     for cat in sorted(categories.keys())
# ]


cache_menu = None

def get_content_categories():
    global cache_menu

    if cache_menu:
        return cache_menu
    
    # Automatically generate categories from json file (based on yaml file provided by fontawesome)
    file = os.path.join(os.path.dirname(os.path.realpath(__file__)), "fontawesome4.json")
    with io.open(file, 'r', encoding='utf-8') as json_file:
        chars = json.load(json_file, object_pairs_hook=OrderedDict)

    # categories = OrderedDict()
    categories = defaultdict(list)
    for char in chars['icons']:
        uc = chr(int(char['unicode'], 16))
        for cat in char['categories']:
            try:
                supertip = "Font Awesome 4 > {}\n{}".format(cat, ", ".join(char['filter']))
            except:
                supertip = "Font Awesome 4 > {}".format(cat)
            categories[cat].append(("FontAwesome", uc, char['name'], supertip))
    
    cache_menu = bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None,
                children=[
                    PPTSymbolsGallery(label="{} ({})".format(cat, len(categories[cat])), symbols=categories[cat], columns=16)
                    for cat in sorted(categories.keys())
                ]
            )
    return cache_menu


def update_search_index(search_engine):
    search_writer = search_engine.writer()

    # Automatically generate categories from json file (based on yaml file provided by fontawesome)
    file = os.path.join(os.path.dirname(os.path.realpath(__file__)), "fontawesome4.json")
    with io.open(file, 'r', encoding='utf-8') as json_file:
        chars = json.load(json_file, object_pairs_hook=OrderedDict)
        
        for char in chars['icons']:
            keywords = set(char['name'].lower().split())
            try:
                #filter can be non-existent or null
                keywords.update(char['filter'])
            except:
                pass
            for cat in char['categories']:
                keywords.update(cat.lower().replace("icons", "").split())
            search_writer.add_document(
                module="fontawesome4",
                fontlabel="Font Awesome 4",
                fontname="FontAwesome",
                unicode=chr(int(char['unicode'], 16)),
                label=char['name'],
                keywords=keywords
            )
    search_writer.commit()


menus = [
    PPTSymbolsGallery(label="{} ({})".format(label, len(symbollist)), symbols=symbollist, columns=columns)
    for (label, symbollist, columns) in menu_settings
] + [
    # submenu for categories
    bkt.ribbon.DynamicMenu(label="All Categories", get_content=bkt.Callback(get_content_categories))
]
