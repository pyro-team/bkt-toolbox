# -*- coding: utf-8 -*-

# https://docs.microsoft.com/de-de/windows/uwp/design/style/segoe-ui-symbol-font


from __future__ import absolute_import

import os.path
import io
import json
from math import ceil

import bkt
from bkt.library.powerpoint import PPTSymbolsGallery

#not used anymore:
mono_icons = [
    "AADLogo",
    "AccessLogo",
    "ATPLogo",
    "AzureLogo",
    "BingLogo",
    "BookingsLogo",
    "ClassroomLogo",
    "DelveAnalyticsLogo",
    "DelveLogo",
    "DynamicSMBLogo",
    "EdgeLogo",
    "ExcelDocument",
    "ExcelLogo",
    "ExchangeLogo",
    "LyncLogo",
    "MSNLogo",
    "OfficeAssistantLogo",
    "OfficeLogo",
    "OfficeStoreLogo",
    "OfficeVideoLogo",
    "OneDriveLogo",
    "OneNoteLogo",
    "OutlookLogo",
    "PowerBILogo",
    "PowerPointDocument",
    "PowerPointLogo",
    "SharepointLogo",
    "SkypeLogo",
    "SocialListeningLogo",
    "StoreLogo",
    "StoreLogoMed",
    "VisioLogo",
    "WindowsLogo",
    "WordDocument",
    "WordLogo",
    "YammerLogo",
]


file = os.path.join(os.path.dirname(os.path.realpath(__file__)), "fabricmdl2.json")
with io.open(file, 'r', encoding='utf-8') as json_file:
    chars = json.load(json_file)
    chars = chars["glyphs"]

    fabric_symbols1 = []
    fabric_symbols2 = []
    fabric_symbols3 = []
    fabric_symbols4 = []
    # fabric_symbols_mono = []
    chunk_size = ceil(len(chars)/4.0)
    for i,char in enumerate(chars):
        if not "unicode" in char:
            continue
        if 0 < i <= chunk_size:
            fabric_symbols1.append(("Fabric MDL2 Assets", unichr(int(char['unicode'], 16)), char['name'], "Fabric MDL2 Assets > 1"))
        elif chunk_size < i <= 2*chunk_size:
            fabric_symbols2.append(("Fabric MDL2 Assets", unichr(int(char['unicode'], 16)), char['name'], "Fabric MDL2 Assets > 2"))
        elif 2*chunk_size < i <= 3*chunk_size:
            fabric_symbols3.append(("Fabric MDL2 Assets", unichr(int(char['unicode'], 16)), char['name'], "Fabric MDL2 Assets > 3"))
        else:
            fabric_symbols4.append(("Fabric MDL2 Assets", unichr(int(char['unicode'], 16)), char['name'], "Fabric MDL2 Assets > 4"))
        # if char['name'] in mono_icons:
        #     fabric_symbols_mono.append(t)


# define the menu parts
menu_title = "Fabric MDL2 Assets"

menus = [
    PPTSymbolsGallery(label="Fabric MDL2 Assets 1 ({})".format(len(fabric_symbols1)), symbols=fabric_symbols1, columns=16),
    PPTSymbolsGallery(label="Fabric MDL2 Assets 2 ({})".format(len(fabric_symbols2)), symbols=fabric_symbols2, columns=16),
    PPTSymbolsGallery(label="Fabric MDL2 Assets 3 ({})".format(len(fabric_symbols3)), symbols=fabric_symbols3, columns=16),
    PPTSymbolsGallery(label="Fabric MDL2 Assets 4 ({})".format(len(fabric_symbols4)), symbols=fabric_symbols4, columns=16),
    # PPTSymbolsGallery(label="Fabric MDL2 Assets Mono ({})".format(len(fabric_symbols_mono)), symbols=fabric_symbols_mono, columns=16),
]

