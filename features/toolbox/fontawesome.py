# -*- coding: utf-8 -*-
'''
Created on 10.02.2017

@author: rdebeerst
'''

import importlib
from collections import namedtuple

import bkt
from bkt.library.powerpoint import PPTSymbolsGallery


FontSymbol = namedtuple("FontSymbol", "module fontname unicode label keywords")

class Fontawesome(object):
    installed_fonts = None
    fontsettings = [
            # module-name,      font-filename, suppress-font-not-installed-message
            ('fontawesome4',    'FontAwesome',                  True),
            ('fontawesome5',    'Font Awesome 5 Free Regular',  False),
            ('segoeui',         'Segoe UI',                     False),
            ('segoemdl2',       'Segoe MDL2 Assets',            True),
            ('materialicons',   'Material Icons',               False),
            ('fabricmdl2',      'Fabric MDL2 Assets',           True),
            # ('foobar', 'Non-existing test font', True),
        ]
    search_engine = None

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

        from bkt.search import get_search_engine
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
                except AttributeError:
                    continue


# initialize galleries
# symbol_galleries = Fontawesome.get_symbol_galleries()

# initialize search
# Fontawesome.update_search_index()



class FontSearch(object):
    search_term = ""
    search_results = None
    search_exact = True

    @classmethod
    def set_search_term(cls, value):
        cls.search_term = value
        if len(cls.search_term) > 1:
            engine = cls.get_search_engine()
            with engine.searcher() as searcher:
                if cls.search_exact:
                    cls.search_results = searcher.search_exact(cls.search_term)
                else:
                    cls.search_results = searcher.search(cls.search_term)
        else:
            cls.search_results = None

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
                    label="Keine Ergebnisse",
                    enabled=False
                )
            ]
        
        else:
            fontmodules = [
                bkt.ribbon.MenuSeparator(title="{} Ergebnisse für »{}«".format(len(cls.search_results), cls.search_term))
            ]
            for fontname, icons in cls.search_results.groupedby("fontname").iteritems():
                fontmodules.append(
                    PPTSymbolsGallery(
                        label="{} ({})".format(fontname, len(icons)),
                        symbols=[
                            (
                                fontname,
                                # ico.unicode,
                                unichr(int(ico.unicode, 16)),
                                ico.label,
                                ', '.join(sorted(ico.keywords))
                            ) for ico in icons
                        ],
                        columns=16
                    )
                )

        engine = cls.get_search_engine()
        fontmodules.extend([
            bkt.ribbon.MenuSeparator(title="Infos"),
            # bkt.ribbon.Menu(
            #     label="Unterstützte Fonts",
            #     children=[
            #         bkt.ribbon.Button(
            #             label="{}: {}".format(font_name, Fontawesome.font_exists(font_name)),
            #             # enabled=False
            #         )
            #         for _, font_name, _ in Fontawesome.fontsettings
            #     ]
            # ),
            bkt.ribbon.Button(
                label="Indizierte Icons: {}".format(engine.count_documents())
            ),
            bkt.ribbon.Button(
                label="Indizierte Keywords: {}".format(engine.count_keywords())
            ),
        ])
        
        return bkt.ribbon.Menu(
                    xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                    id=None,
                    children=fontmodules
                )
    
    @classmethod
    def get_enabled_results(cls):
        return cls.search_results is not None
    
    @classmethod
    def get_results_label(cls):
        if cls.search_results is not None:
            return "{} Ergebnisse".format(len(cls.search_results))
        else:
            return "Ergebnisse"
    
    @classmethod
    def sync_cache(cls):
        engine = cls.get_search_engine()
        engine.cache_sync()
    
    @classmethod
    def clear_cache(cls):
        engine = cls.get_search_engine()
        engine.cache_clear()

# Font search
fontsearch_gruppe = bkt.ribbon.Group(
    id="bkt_fontsearch_group",
    label="Icon-Suche",
    image_mso='SearchUI',
    children=[
        bkt.ribbon.EditBox(
            label="Suche:",
            sizeString = '#######',
            get_text = bkt.Callback(FontSearch.get_search_term),
            on_change = bkt.Callback(FontSearch.set_search_term),
        ),
        bkt.ribbon.DynamicMenu(
            get_label=bkt.Callback(FontSearch.get_results_label),
            get_content=bkt.Callback(FontSearch.get_symbol_galleries),
            get_enabled=bkt.Callback(FontSearch.get_enabled_results),
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