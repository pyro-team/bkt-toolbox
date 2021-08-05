# -*- coding: utf-8 -*-
'''
Created on 10.02.2017

@author: rdebeerst
'''

from __future__ import absolute_import

import importlib
from collections import namedtuple, deque

import bkt
from bkt.library.powerpoint import PPTSymbolsGallery


FontSymbol = namedtuple("FontSymbol", "module fontlabel fontname unicode label keywords")

class Fontawesome(object):
    installed_fonts = None
    fontsettings = [
            # module-name,      font-filename, suppress-font-not-installed-message
            ('fabricmdl2',      'Fabric MDL2 Assets',           True),
            ('fontawesome4',    'FontAwesome',                  True),
            ('fontawesome5',    'Font Awesome 5 Free Regular',  False),
            ('icomoon',         'IcoMoon-Free',                 False),
            ('materialicons',   'Material Icons',               False),
            ('segoemdl2',       'Segoe MDL2 Assets',            True),
            ('segoeui',         'Segoe UI',                     False),
            ('wingdings',       'Wingdings',                    True),
            # ('foobar', 'Non-existing test font', True),
        ]
    search_engine = None
    searchable_fonts = []

    @classmethod
    def get_installed_fonts(cls):
        if not cls.installed_fonts:
            # Method 1 (returns Font Awesome 5 Free Regular)
            import System.Drawing.Text
            font_collection = System.Drawing.Text.InstalledFontCollection()
            cls.installed_fonts = [font.Name for font in font_collection.Families]
            # Method 2 (return Font Awesome 5 Free)
            # import System.Windows.Media
            # font_families = System.Windows.Media.Fonts.SystemFontFamilies
            # cls.installed_fonts = [font.Source for font in font_families]
        return cls.installed_fonts
    
    # helper to check system-fonts
    @classmethod
    def font_exists(cls, fontname):
        return fontname in cls.get_installed_fonts()

    @classmethod
    def get_symbol_galleries(cls):
        symbol_galleries = []
        for font_module, font_name, suppress_hint in cls.fontsettings:
            # check if font exists
            if cls.font_exists(font_name):
                # import the corresponding font-symbol-module from 'fontsymbols'-folder
                fontsymbolmodule = importlib.import_module('toolbox.fontsymbols.%s' % font_module)
                
                # add menu seperator with title
                if fontsymbolmodule.menu_title:
                    symbol_galleries += [
                        bkt.ribbon.MenuSeparator(title="" + fontsymbolmodule.menu_title),
                    ]
                else:
                    symbol_galleries += [
                        bkt.ribbon.MenuSeparator(),
                    ]
                
                # add font-symbol-galleries
                symbol_galleries += fontsymbolmodule.menus
            elif not suppress_hint:
                symbol_galleries += [
                    bkt.ribbon.MenuSeparator(title=font_name),
                    bkt.ribbon.Button(
                        label="Font nicht installiert",
                        enabled=False
                    )
                ]
        return symbol_galleries

    @classmethod
    def get_search_engine(cls):
        if cls.search_engine:
            return cls.search_engine

        from bkt.library.search import get_search_engine
        cls.search_engine = get_search_engine("fonticons", FontSymbol)
        # initialize search index on first use of engine
        cls.update_search_index(cls.search_engine)
        return cls.search_engine
    
    @classmethod
    def update_search_index(cls, engine):
        for font_module, font_name, _ in cls.fontsettings:
            # check if font exists
            if cls.font_exists(font_name):
                # import the corresponding font-symbol-module from 'fontsymbols'-folder
                fontsymbolmodule = importlib.import_module('toolbox.fontsymbols.%s' % font_module)
                try:
                    fontsymbolmodule.update_search_index(engine)
                    cls.searchable_fonts.append(fontsymbolmodule.menu_title)
                except AttributeError:
                    continue
    @classmethod
    def get_text_fontawesome(cls):
        return bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None,
                children=cls.get_symbol_galleries()
            )


class FontSearch(object):
    search_term = ""
    search_results = None
    search_exact = True
    search_and = True #True = AND-search, False=OR-search

    _cache_menu_infos = None

    @classmethod
    def _perform_search(cls):
        if len(cls.search_term) > 0:
            if len(cls.search_term) < 3:
                cls.search_exact = True
            cls.search_results = None
            engine = cls.get_search_engine()
            with engine.searcher() as searcher:
                if cls.search_exact:
                    cls.search_results = searcher.search_exact(cls.search_term, cls.search_and)
                else:
                    cls.search_results = searcher.search(cls.search_term, cls.search_and)
        else:
            cls.search_results = None

    @classmethod
    def toggle_search_exact(cls, pressed):
        cls.search_exact = not cls.search_exact
        cls._perform_search()

    @classmethod
    def checked_search_exact(cls):
        return cls.search_exact

    @classmethod
    def set_search_term(cls, value):
        cls.search_term = value
        cls._perform_search()

    @classmethod
    def get_search_term(cls):
        return cls.search_term

    @classmethod
    def get_search_engine(cls):
        return Fontawesome.get_search_engine()
    
    @classmethod
    def get_symbol_galleries(cls):
        if not cls.search_results or len(cls.search_results) == 0:
            fontmodules = [
                bkt.ribbon.Button(
                    label="Keine Ergebnisse für »{}«".format(cls.search_term),
                    enabled=False
                )
            ]
        
        else:
            fontmodules = [
                bkt.ribbon.MenuSeparator(title="{} Ergebnisse für »{}«".format(len(cls.search_results), cls.search_term))
            ]
            for fontlabel, icons in cls.search_results.groupedby("fontlabel"):
                fontmodules.append(
                    PPTSymbolsGallery(
                        label="{} ({})".format(fontlabel, len(icons)),
                        symbols=[
                            (
                                ico.fontname,
                                ico.unicode,
                                # unichr(int(ico.unicode, 16)),
                                ico.label,
                                ', '.join(sorted(ico.keywords))
                            ) for ico in icons
                        ],
                        columns=16
                    )
                )

        fontmodules.extend(cls._get_symbol_galleries_infos())
        
        return bkt.ribbon.Menu(
                    xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                    id=None,
                    children=fontmodules
                )
    
    @classmethod
    def _get_symbol_galleries_infos(cls):
        if cls._cache_menu_infos:
            return cls._cache_menu_infos

        engine = cls.get_search_engine()
        cls._cache_menu_infos = [
            bkt.ribbon.MenuSeparator(title="Informationen"),
            bkt.ribbon.ToggleButton(
                label="Exakte Suche ein/aus",
                supertip="Wenn die exakte Suche deaktiviert ist, wird bei 'person' auch 'personality', 'impersonal', usw. gefunden.",
                on_toggle_action=bkt.Callback(cls.toggle_search_exact),
                get_pressed=bkt.Callback(cls.checked_search_exact),
            ),
            bkt.ribbon.Button(
                label="Indizierte Icons: {}".format(engine.count_documents()),
                enabled=False,
            ),
            bkt.ribbon.Button(
                label="Indizierte Keywords: {}".format(engine.count_keywords()),
                enabled=False,
            ),
            bkt.ribbon.Button(
                label="Durchsuchbare Fonts: {}".format(len(Fontawesome.searchable_fonts)),
                enabled=False,
                supertip=", ".join(Fontawesome.searchable_fonts)
            ),
        ]
        return cls._cache_menu_infos
    
    @classmethod
    def get_enabled_results(cls):
        return cls.search_results is not None
    
    @classmethod
    def get_results_label(cls):
        if cls.search_results is not None:
            return "{} Icons".format(len(cls.search_results))
        else:
            return "Ergebnis"


# Font search
fontsearch_gruppe = bkt.ribbon.Group(
    id="bkt_fontsearch_group",
    label="Icon-Suche",
    image_mso='GroupSearch',
    children=[
        bkt.ribbon.DynamicMenu(
            id="fontsearch_all_symbols",
            label="Alle Icons",
            size="large",
            supertip="Zeigt Icons für verfügbare Icon-Fonts an, die als Textsymbol oder Grafik eingefügt werden können.\n\nHinweis: Die Icon-Fonts müssen auf dem Rechner installiert sein.",
            image_mso="Call",
            get_content = bkt.Callback(Fontawesome.get_text_fontawesome)
        ),
        bkt.ribbon.Separator(),
        bkt.ribbon.Label(
            label="Suchwort:",
        ),
        # bkt.ribbon.EditBox(
        bkt.ribbon.ComboBox(
            label="Suchwort",
            show_label=False,
            sizeString = '#######',
            get_text = bkt.Callback(FontSearch.get_search_term),
            on_change = bkt.Callback(FontSearch.set_search_term),
            supertip="Suchwort eingeben und ENTER klicken",
            get_item_count=bkt.Callback(lambda: FontSearch.get_search_engine().count_recent_searches()),
            get_item_label=bkt.Callback(lambda index: FontSearch.get_search_engine().get_recent_searches()[index]),
        ),
        bkt.ribbon.DynamicMenu(
            get_label=bkt.Callback(FontSearch.get_results_label),
            get_content=bkt.Callback(FontSearch.get_symbol_galleries),
            get_enabled=bkt.Callback(FontSearch.get_enabled_results),
            screentip="Suchergebnisse",
            supertip="Zeigt die Suchergebnisse der Icon-Suche nach dem gewünschten Suchwort an",
        ),
        # bkt.ribbon.Box(children=[
        #     bkt.ribbon.Button(
        #         label="sync",
        #         on_action=bkt.Callback(FontSearch.sync_cache),
        #     ),
        #     bkt.ribbon.Button(
        #         label="clear",
        #         on_action=bkt.Callback(FontSearch.clear_cache),
        #     ),
        # ]),
    ]
)