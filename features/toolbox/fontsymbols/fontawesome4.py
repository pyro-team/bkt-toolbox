# -*- coding: utf-8 -*-

# https://fontawesome.com/v4.7.0/

import bkt
from bkt.library.powerpoint import PPTSymbolsGallery


symbols_communication = [
    ["FontAwesome", u"\uf2ba", "address book"],
    ["FontAwesome", u"\uf2bc", "address card"],
    ["FontAwesome", u"\uf2c1", "id badge"],
    ["FontAwesome", u"\uf2c3", "id card"],
    ["FontAwesome", u"\uf183", "man"],
    ["FontAwesome", u"\uf0c0", "users"],
    ["FontAwesome", u"\uf2be", "user circle"],
    ["FontAwesome", u"\uf2c0", "user"],
    ["FontAwesome", u"\uf007", "user black"],
    ["FontAwesome", u"\uf2b5", "handshake"],
    ["FontAwesome", u"\uf0e5", "comment"],
    ["FontAwesome", u"\uf27b", "commenting"],
    ["FontAwesome", u"\uf0e6", "comments"],
    ["FontAwesome", u"\uf086", "comments"],
]

symbols_itsystems = [
    
    ["FontAwesome", u"\uf108", "desktop"],
    ["FontAwesome", u"\uf109", "laptop"],
    ["FontAwesome", u"\uf10a", "tablet"],
    ["FontAwesome", u"\uf10b", "mobile"],
    ["FontAwesome", u"\uf095", "phone"],
    ["FontAwesome", u"\uf1ac", "fax"],
    ["FontAwesome", u"\uf003", "mail"],
    ["FontAwesome", u"\uf01c", "inbox"],
    ["FontAwesome", u"\uf11c", "keyboard"],
    
    ["FontAwesome", u"\uf0c2", "cloud"],
    ["FontAwesome", u"\uf09e", "rss"],
    ["FontAwesome", u"\uf1eb", "wifi"],
    
    ["FontAwesome", u"\uf090", "sign in"],
    ["FontAwesome", u"\uf084", "key"],
    ["FontAwesome", u"\uf023", "lock"],
    ["FontAwesome", u"\uf09c", "unlock"],
    ["FontAwesome", u"\uf13e", "unlock"],
    ["FontAwesome", u"\uf132", "shield"],
    
    ["FontAwesome", u"\uf19c", "university/bank"],
    ["FontAwesome", u"\uf015", "home"],
    ["FontAwesome", u"\uf1ad", "building"],
    
    ["FontAwesome", u"\uf1b2", "cube"],
    ["FontAwesome", u"\uf1b3", "cubes"],
    ["FontAwesome", u"\uf1c0", "database"],
    ["FontAwesome", u"\uf233", "server"],
    ["FontAwesome", u"\uf2db", "microchip"],
    
    ["FontAwesome", u"\uf188", "bug"],
]

symbols_files = [
    ["FontAwesome", u"\uf114", "folder"],
    ["FontAwesome", u"\uf115", "folder open"],
    ["FontAweSome", u"\uf016", "file"],
    ["FontAweSome", u"\uf0c5", "files"],
    ["FontAweSome", u"\uf0f6", "text"],
    ["FontAweSome", u"\uf1c9", "code"],
    ["FontAweSome", u"\uf1c6", "archive"],
    ["FontAweSome", u"\uf1c5", "image"],
    ["FontAweSome", u"\uf1c8", "video"],
    ["FontAweSome", u"\uf1c7", "audio"],
    ["FontAweSome", u"\uf1c1", "pdf"],
    ["FontAweSome", u"\uf1c3", "excel"],
    ["FontAweSome", u"\uf1c4", "powerpoint"],
    ["FontAweSome", u"\uf1c2", "word"],
]

symbols_analysis = [
    ["FontAwesome", u"\uf080", "bar chart"],
    ["FontAwesome", u"\uf200", "pie chart"],
    ["FontAwesome", u"\uf201", "line chart"],
    ["FontAwesome", u"\uf1fe", "area chart"],
    ["FontAwesome", u"\uf02a", "barcode"],
    ["FontAwesome", u"\uf24e", "balance"],
    
    ["FontAwesome", u"\uf005", "star"],
    ["FontAwesome", u"\uf123", "star half"],
    ["FontAwesome", u"\uf006", "star empty"],
    ["FontAwesome", u"\uf251", "hourglass start"],
    ["FontAwesome", u"\uf252", "hourglass half"],
    ["FontAwesome", u"\uf253", "hourglass end"],
    
    ["FontAwesome", u"\uf244", "battery empty"],
    ["FontAwesome", u"\uf243", "battery 1/4"],
    ["FontAwesome", u"\uf242", "battery 1/2"],
    ["FontAwesome", u"\uf241", "battery 3/4"],
    ["FontAwesome", u"\uf240", "battery full"],
    ["FontAwesome", u"\uf05e", "ban"],

    ["FontAwesome", u"\uf087", "thumbs up"],
    ["FontAwesome", u"\uf088", "thumbs down"],
    ["FontAwesome", u"\uf164", "thumbs up"],
    ["FontAwesome", u"\uf165", "thumbs down"],
    
    ["FontAwesome", u"\uf046", "check"],
    ["FontAwesome", u"\uf05d", "check circle"],
    ["FontAwesome", u"\uf00c", "check"],
    ["FontAwesome", u"\uf00d", "x"],
]


symbols_mixed = [
    ["FontAwesome", u"\uf0d0", "magic"],
    ["FontAwesome", u"\uf02d", "book"],
    ["FontAwesome", u"\uf02e", "bookmark"],
    ["FontAwesome", u"\uf0b1", "briefcase"],
    ["FontAwesome", u"\uf140", "bullseye"],
    ["FontAwesome", u"\uf073", "calendar"],
    ["FontAwesome", u"\uf0a3", "certificate"],
    ["FontAwesome", u"\uf017", "clock"],
    ["FontAwesome", u"\uf013", "cog"],
    ["FontAwesome", u"\uf085", "cogs"],
    ["FontAwesome", u"\uf134", "fire extinguisher"],
    ["FontAwesome", u"\uf277", "map signs"],
    ["FontAwesome", u"\uf041", "map marker"],
    ["FontAwesome", u"\uf1ea", "newspaper"],
    ["FontAwesome", u"\uf0eb", "lightbulb"],
    ["FontAwesome", u"\uf08d", "pin"],
    ["FontAwesome", u"\uf074", "random"],
    ["FontAwesome", u"\uf1b8", "recycle"],
    ["FontAwesome", u"\uf021", "refresh"],
    ["FontAwesome", u"\uf135", "rocket"],
    ["FontAwesome", u"\uf002", "search"],
    ["FontAwesome", u"\uf0e4", "tachometer"],
    ["FontAwesome", u"\uf02b", "tag"],
    ["FontAwesome", u"\uf02c", "tags"],
    ["FontAwesome", u"\uf014", "trash"],
    ["FontAwesome", u"\uf0ad", "wrench"],
]


# create gallery for all symbols
# symbols in uid-range, f000 to f299
hex_start_end = ('f000', 'f300')
int_start_end = (int(x, 16) for x in hex_start_end)
uids = range(*int_start_end)

# [ symbol-gallery-item, symbol-gallery-item, ... ]
# symbol-gallery-item = [fontname, character-id, label]
symbols_all = [ 
    [
        'FontAwesome',
        unichr(uid),
        'FontAwesome, unicode character %s' % format(uid, '02x')
    ]
    for uid in uids
]


# define the menu parts

menu_title = 'Font Awesome 4.7'

menu_settings = [
    # menu label,          list of symbols,       icons per row
    ('IT-System',          symbols_itsystems,           6  ),
    ('Kommunikation',      symbols_communication,       6  ),
    ('Dateien',            symbols_files,               6  ),
    ('Analyse/Bewertung',  symbols_analysis,            6  ),
    ('Mixed',              symbols_mixed,               6  ),
    ('All',                symbols_all,                16  )
    
]

menus = [
    PPTSymbolsGallery(label=label, symbols=symbollist, columns=columns)
    for (label, symbollist, columns) in menu_settings
]


