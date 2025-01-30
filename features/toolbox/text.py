# -*- coding: utf-8 -*-
'''
Created on 06.07.2016

@author: rdebeerst
'''

import bkt
import bkt.library.powerpoint as pplib



class Characters(object):
    @staticmethod
    def symbol_insert(context):
        if bkt.get_key_state(bkt.KeyCodes.SHIFT):
            Characters.add_protected_hyphen(context.app.ActiveWindow.Selection)
        elif bkt.get_key_state(bkt.KeyCodes.CTRL):
            Characters.add_protected_space(context.app.ActiveWindow.Selection)
        else:
            context.app.commandbars.ExecuteMso("SymbolInsert")

    
    ### TYPOGRAPHY ###
    @staticmethod
    def add_protected_hyphen(selection):
        selection.TextRange2.text='\xad'

    @staticmethod
    def add_protected_space(selection):
        selection.TextRange2.text='\xa0'

    @staticmethod
    def add_protected_narrow_space(selection):
        selection.TextRange2.text='\u202F'

    ### TYPOGRAPHY ###
    typography = [
        (None, '\xbb', "Linkes Guillemets"),
        (None, '\xab', "Rechtes Guillemets"),
        (None, '\xb6', "Paragraph"),
        (None, '\u2026', "Auslassungspunkte (Ellipse)", "Auslassungspunkte sind drei kurz aufeinanderfolgende Punkte. Meistens zeigen diese eine Ellipse (Auslassung eines Textteils) an."),
        (None, '\u2013', "Gedankenstrich (Halbgeviertstrich/En-Dash)", "Ein Gedankenstrich (sog. Halbgeviertstrich) wie er von Office teilweise automatisch gesetzt wird. Verwendet als Bis-Strich oder Streckenstrich."),
        (None, '\u2014', "Waagerechter Strich (Geviertstrich/Em-Dash)"),
        (None, '\u2020', "Kreuz"),
        (None, '\u2021', "Doppelkreuz"),
        (None, '\u25A0', "Schwarzes Quadrat"),
        (None, '\u25A1', "Weißes Quadrat"),
        (None, '\u2423', "Leerzeichen-Symbol"),
        (None, '\xa9',   "Copyright"),
        (None, '\xae',   "Registered Trade Mark"),
        (None, '\u2122', "Trade Mark"),
        (None, '\u2030', "Per mil"),
        (None, '\u20AC', "Euro"),
        (None, '\u1E9E', "Großes Eszett"),
    ]

    ### MATH ###
    math = [
        (None, '\xb1',   "Plus-Minus-Zeichen", "Ein Plus-Minus-Zeichen einfügen."),
        (None, '\u2212', "Echtes Minuszeichen", "Ein echtes Minuszeichen (kein Bindestrich) einfügen."),
        (None, '\xd7',   "Echtes Malzeichen (Kreuz)", "Ein echtes Kreuz-Multiplikatorzeichen einfügen."),
        (None, '\u22c5', "Echtes Malzeichen (Punkt)", "Ein echtes Punkt-Multiplikatorzeichen einfügen."),
        (None, '\u2044', "Echter Bruchstrich", "Einen echten Bruchstrich (ähnlich Schrägstrich) einfügen."),
        (None, '\u2248', "Ungefähr Gleich", "Ein Ungefähr Gleich Zeichen einfügen."),
        (None, '\u2260', "Ungleich", "Ein Ungleich-Zeichen einfügen."),
        (None, '\u2206', "Delta", "Ein Deltazeichen einfügen."), #alt: \u0394 griechisches Delta
        (None, '\u2300', "Mittelwert/Durchmesser", "Ein Durchmesserzeichen bzw. Durchschnittszeichen einfügen."), #alt: \xD8 leere menge
        (None, '\u2211', "Summenzeichen", "Ein Summenzeichen einfügen."),
        (None, '\u221A', "Wurzelzeichen", "Ein Wurzelzeichen einfügen."),
        (None, '\u221E', "Unendlich-Zeichen", "Ein Unendlich-Zeichen (liegende Acht) einfügen."),
        (None, '\u2264', "Kleiner-Gleich", "Ein kleiner oder gleich Zeichen einfügen."),
        (None, '\u2265', "Größer-Gleich", "Ein größer oder gleich Zeichen einfügen."),
    ]

    ### LIST ###
    lists = [
        (None, '\u2022', "Aufzählungszeichen (Kreis)", "Ein Aufzählungszeichen (schwarzer Punkt) einfügen."),
        (None, '\u2023', "Aufzählungszeichen (Dreieck)", "Ein Aufzählungszeichen (schwarzes Dreieck) einfügen."),
        (None, '\u25AA', "Aufzählungszeichen (Quadrat)", "Ein Aufzählungszeichen (schwarzes Quadrat) einfügen."),
        (None, '\u2043', "Aufzählungszeichen (Strich)", "Ein Aufzählungszeichen (Bindestrich) einfügen."),
        (None, '\u2212', "Echtes Minuszeichen", "Ein echtes Minuszeichen (kein Bindestrich) einfügen."),
        (None, '\x2b',   "Pluszeichen", "Ein Pluszeichen einfügen."),
        (None, '\u2610', "Box leer"),
        (None, '\u2611', "Box Häkchen"),
        (None, '\u2612', "Box Kreuzchen"),
        ("Wingdings", 'J', "Wingdings Smiley gut"),
        ("Wingdings", 'K', "Wingdings Smiley neutral"),
        ("Wingdings", 'L', "Wingdings Smiley schlecht"),
        (None, '\u2713', "Häkchen", "Ein Häkchen-Symbol einfügen."),
        (None, '\u2714', "Häkchen fett", "Ein fettes Häkchen-Symbol einfügen."),
        (None, '\u2717', "Kreuzchen geschwungen", "Ein geschwungenes Kreuzchen (passend zu Häkchen) einfügen."),
        (None, '\u2718', "Kreuzchen geschwungen fett", "Ein fettes geschwungenes Kreuzchen (passend zu Häkchen) einfügen."),
        (None, '\u2715', "Kreuzchen symmetrisch", "Ein symmetrisches Kreuzchen (ähnlich Malzeichen) einfügen."),
        (None, '\u2716', "Kreuzchen symmetrisch fett", "Ein fettes symmetrisches Kreuzchen (ähnlich Malzeichen) einfügen."),
        (None, '\u2605', "Stern schwarz"),
        (None, '\u2606', "Stern weiß"),
        (None, '\u261B', "Zeigefinger schwarz"),
        (None, '\u261E', "Zeigefinger weiß"),
        ("Wingdings", 'C', "Wingdings Thumbs-Up"),
        ("Wingdings", 'D', "Wingdings Thumbs-Down"),
        ### Default list symbol:
        # ("Arial",       u'\u2022', "Arial Bullet"),
        ("Courier New", 'o', "Courier New Kreis"),
        ("Wingdings",   '\xa7', "Wingdings Rechteck"),
        ("Symbol",      '-', "Symbol Strich"),
        ("Wingdings",   'v', "Wingdings Stern"),
        ("Wingdings",   '\xd8', "Wingdings Pfeil"),
        ("Wingdings",   '\xfc', "Wingdings Häckchen"),
    ]

    ### ARROWS ###
    arrows = [
        (None, '\u2192', "Pfeil rechts"),
        (None, '\u2190', "Pfeil links"),
        (None, '\u2191', "Pfeil oben"),
        (None, '\u2193', "Pfeil unten"),
        (None, '\u2194', "Pfeil links und rechts"),
        (None, '\u21C4', "Pfeil links und rechts"),
        (None, '\u2197', "Pfeil rechts oben"),
        (None, '\u2196', "Pfeil links oben"),
        (None, '\u2198', "Pfeil rechts unten"),
        (None, '\u2199', "Pfeil links unten"),
        (None, '\u2195', "Pfeil oben und unten"),
        (None, '\u21C5', "Pfeil oben und unten"),
        (None, '\u21E8', "Weißer Pfeil rechts"),
        (None, '\u21E6', "Weißer Pfeil links"),
        (None, '\u21E7', "Weißer Pfeil oben"),
        (None, '\u21E9', "Weißer Pfeil unten"),
        (None, '\u21AF', "Blitz"),
        (None, '\u21BA', "Kreispfeil gegen den Uhrzeigersinn"),
    ]

    @staticmethod
    def text_selection(selection):
        return selection.Type == 3


    @classmethod
    def get_text_fontawesome(cls):
        from .fontawesome import Fontawesome

        return bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None,
                children=Fontawesome.get_symbol_galleries()
            )

    @classmethod
    def get_text_fontawesome_exclusion(cls):
        from .fontawesome import Fontawesome

        return bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None,
                children=Fontawesome.get_exclusions()
            )
        
    @classmethod
    def get_text_unicodefont(cls):
        def _unicode_font_button(font):
            return bkt.ribbon.ToggleButton(
                label=font,
                screentip="Unicode-Schriftart "+font,
                supertip=font+" als Unicode-Schriftart verwenden.",
                on_toggle_action=bkt.Callback(lambda pressed: pplib.PPTSymbolsSettings.switch_unicode_font(font)),
                get_pressed=bkt.Callback(lambda: pplib.PPTSymbolsSettings.unicode_font == font),
                get_image=bkt.Callback(lambda:bkt.ribbon.SymbolsGallery.create_symbol_image(font, "\u2192"))
            )

        return bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None,
                children=[
                            bkt.ribbon.ToggleButton(
                                label='Theme-Schriftart (Standard)',
                                screentip="Unicode-Schrift entspricht Theme-Schriftart",
                                supertip="Es wird keine spezielle Unicode-Schriftart verwendet, sondern die Standard-Schriftart des Themes.",
                                on_toggle_action=bkt.Callback(lambda pressed: pplib.PPTSymbolsSettings.switch_unicode_font(None)),
                                get_pressed=bkt.Callback(lambda: pplib.PPTSymbolsSettings.unicode_font is None),
                            ),
                        ] + [
                            _unicode_font_button(font)
                            for font in ["Arial", "Arial Unicode MS", "Calibri", "Lucida Sans Unicode", "Segoe UI"]
                        ]
            )

    @classmethod
    def get_text_menu(cls):
        recent_symbols = pplib.PPTSymbolsGalleryRecent(
            id="symbols_recent_gallery",
            label="Zuletzt verwendet",
        )

        return bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                # id=None,
                id="symbols_splitbutton",
                label="Symbol-Menü",
                show_label=False,
                image_mso="SymbolInsert",
                screentip="Symbol",
                supertip="Öffnet ein Menü mit verschiedenen Gallerien zum schnellen Einfügen von Symbolen und speziellen Zeichen.",
                children=[
                    bkt.mso.button.SymbolInsert,
                    bkt.ribbon.MenuSeparator(title="Zuletzt verwendet"),
                    recent_symbols.get_index_as_button(2),
                    recent_symbols.get_index_as_button(1),
                    recent_symbols.get_index_as_button(0),
                    bkt.ribbon.MenuSeparator(title="Symbole"),
                    bkt.ribbon.Button(
                        id='symbols_add_protected_hyphen',
                        label='Geschützter Trennstrich',
                        supertip='Ein geschützter Trennstrich ist ein Symbol zur optionalen Silbentrennung. Der Trennstrich erscheint nur am Zeilenende und bleibt sonst unsichtbar.',
                        on_action=bkt.Callback(cls.add_protected_hyphen, selection=True),
                        get_enabled = bkt.Callback(cls.text_selection, selection=True),
                        get_image=bkt.Callback(lambda:bkt.ribbon.SymbolsGallery.create_symbol_image("Arial", "-"))
                    ),
                    bkt.ribbon.Button(
                        id='symbols_add_protected_space',
                        label='Geschütztes Leerzeichen',
                        supertip='Ein geschütztes Leerzeichen erlaubt keinen Zeilenumbruch.',
                        on_action=bkt.Callback(cls.add_protected_space, selection=True),
                        get_enabled = bkt.Callback(cls.text_selection, selection=True),
                        get_image=bkt.Callback(lambda:bkt.ribbon.SymbolsGallery.create_symbol_image("Arial", "\u23B5")) #alt: 2423
                    ),
                    bkt.ribbon.Button(
                        id='symbols_add_protected_narrow_space',
                        label='Schmales geschütztes Leerzeichen',
                        supertip='Ein schmales geschütztes Leerzeichen erlaubt keinen Zeilenumbruch und ist bspw. zwischen Buchstaben von Abkürzungen zu verwenden.',
                        on_action=bkt.Callback(cls.add_protected_narrow_space, selection=True),
                        get_enabled = bkt.Callback(cls.text_selection, selection=True),
                        get_image=bkt.Callback(lambda:bkt.ribbon.SymbolsGallery.create_symbol_image("Arial", "\u02FD"))
                    ),

                    pplib.PPTSymbolsGallery(
                        id="symbols_typo_gallery",
                        label="Typografiesymbole",
                        supertip="Verschiedene Typografiesymbole einfügen",
                        symbols = cls.typography,
                    ),
                    bkt.ribbon.MenuSeparator(),

                    pplib.PPTSymbolsGallery(
                        id="symbols_math_gallery",
                        label="Mathesymbole",
                        supertip="Verschiedene Mathesymbole einfügen",
                        symbols = cls.math,
                    ),
                    pplib.PPTSymbolsGallery(
                        id="symbols_lists_gallery",
                        label="Listensymbole",
                        supertip="Verschiedene Listensymbole einfügen",
                        symbols = cls.lists,
                    ),
                    pplib.PPTSymbolsGallery(
                        id="symbols_arrow_gallery",
                        label="Pfeile",
                        supertip="Verschiedene Pfeile einfügen",
                        symbols = cls.arrows,
                    ),
                # ] + fontawesome.symbol_galleries + [
                    bkt.ribbon.MenuSeparator(),
                    bkt.ribbon.DynamicMenu(
                        id="symbols_icon_fonts",
                        label="Icon-Fonts",
                        supertip="Zeigt Icons für verfügbare Icon-Fonts an, die als Textsymbol oder Grafik eingefügt werden können.\n\nHinweis: Die Icon-Fonts müssen auf dem Rechner installiert sein.",
                        image_mso="Call",
                        get_content = bkt.Callback(cls.get_text_fontawesome)
                    ),
                    bkt.ribbon.MenuSeparator(title="Einstellungen"),
                    bkt.ribbon.DynamicMenu(
                        label="Unicode-Schriftart wählen",
                        image_mso='FontDialogPowerPoint',
                        supertip="Unicode-Zeichen können entweder mit der Standard-Schriftart oder einer speziellen Unicode-Schriftart eingefügt werden. Diese kann hier ausgewählt werden.",
                        get_content = bkt.Callback(cls.get_text_unicodefont)
                    ),
                    bkt.ribbon.DynamicMenu(
                        label="Icons-Fonts ausschließen",
                        # image_mso='FontDialogPowerPoint',
                        supertip="xxx",
                        get_content = bkt.Callback(cls.get_text_fontawesome_exclusion)
                    ),
                    bkt.ribbon.ToggleButton(
                        label='Als Text einfügen (Standard)',
                        image_mso='TextTool',
                        screentip="Als Text einfügen ein/aus",
                        supertip='Wenn kein Text ausgewählt und diese Option aktiviert ist, wird das Symbol als Unicode-Zeichen eingefügt. Dies ist der Standard wenn keine Taste gedrückt wird.',
                        on_toggle_action=bkt.Callback(pplib.PPTSymbolsSettings.switch_convert_into_text),
                        get_pressed=bkt.Callback(pplib.PPTSymbolsSettings.convert_into_text), #convert into text is a function!
                    ),
                    bkt.ribbon.ToggleButton(
                        label='Als Shapes einfügen [Shift]',
                        image_mso='TextEffectTransformGallery',
                        screentip="Als Shape einfügen ein/aus",
                        supertip='Wenn kein Text ausgewählt und diese Option aktiviert ist, wird das Symbol in ein Shape konvertiert. Dies geht auch bei Klick auf ein Symbol mit gedrückter Shift-Taste.',
                        on_toggle_action=bkt.Callback(pplib.PPTSymbolsSettings.switch_convert_into_shape),
                        get_pressed=bkt.Callback(lambda: pplib.PPTSymbolsSettings.convert_into_shape),
                    ),
                    bkt.ribbon.ToggleButton(
                        label='Als Bild einfügen [Strg]',
                        image_mso='PictureRecolorBlackAndWhite',
                        screentip="Als Bild einfügen ein/aus",
                        supertip='Wenn kein Text ausgewählt und diese Option aktiviert ist, wird das Symbol als Raster-Grafik eingefügt. Dies geht auch bei Klick auf ein Symbol mit gedrückter Strg-Taste.',
                        on_toggle_action=bkt.Callback(pplib.PPTSymbolsSettings.switch_convert_into_bitmap),
                        get_pressed=bkt.Callback(lambda: pplib.PPTSymbolsSettings.convert_into_bitmap),
                    ),
                ]
            )


#TODO: Use MouseKeyHook to register Strg+-/Space key combination in order to add special chars

#OPTION 1: Dynamic Menu - Cons: Buttons (e.g. hyphen) cannot be added to quick access toolbar
# symbol_insert_splitbutton = bkt.ribbon.DynamicMenu(
#     id="symbols_splitbutton",
#     label="Symbol-Menü",
#     show_label=False,
#     image_mso="SymbolInsert",
#     screentip="Symbol",
#     supertip="Öffnet ein Menü mit verschiedenen Gallerien zum schnellen Einfügen von Symbolen und speziellen Zeichen.",
#     get_content = bkt.Callback(
#         Characters.get_text_menu
#     ),
# )

#OPTION 2: Splitbutton with regular menu - Cons: Splitbutton is not intuitive and not compatible with dynamic menu
# symbol_insert_splitbutton = bkt.ribbon.SplitButton(
#     id="symbols_splitbutton",
#     show_label=False,
#     children=[
#         bkt.ribbon.Button(
#             label="Symbol",
#             image_mso="SymbolInsert",
#             screentip="Symbol",
#             supertip="Öffnet den Dialog zum Einfügen von Symbolen.\n\nMit gedrückter Umschalt-Taste wird direkt ein geschützter Trennstrich eingefügt.\n\nMit gedrückter Strg-Taste wird in geschütztes Leerzeichen eingefügt.",
#             on_action=bkt.Callback(Characters.symbol_insert, context=True),
#             get_enabled=bkt.Callback(lambda context: context.app.commandbars.GetEnabledMso("SymbolInsert"), context=True),
#         ),
#         #bkt.mso.button.SymbolInsert,
#         # character_menu
#         Characters.get_text_menu()
#     ]
# )

#OPTION 3: Regular menu with dynamic menu only for icons fonts
symbol_insert_splitbutton = Characters.get_text_menu()





class InnerMargin(pplib.TextframeSpinnerBox):
    
    ### class methods ###
    
    all_equal = False

    @classmethod
    def toggle_all_equal(cls, pressed):
        cls.all_equal = pressed

    @classmethod
    def get_all_equal(cls):
        return cls.all_equal
    
    ### set margin to 0
    
    @classmethod
    def set_to_0(cls, shapes, context):
        for textframe in pplib.iterate_shape_textframes(shapes):
            textframe.MarginTop    = 0
            textframe.MarginBottom = 0
            textframe.MarginLeft   = 0
            textframe.MarginRight  = 0


    ### Setter methods ###
    
    def set_attr_for_textframe(self, textframe, value):
        setattr(textframe, self.attr, value)
        if InnerMargin.all_equal:
            textframe.MarginTop    = value
            textframe.MarginBottom = value
            textframe.MarginLeft   = value
            textframe.MarginRight  = value



inner_margin_top    = InnerMargin(attr="MarginTop",    id='textFrameMargin-top-2',    image_button=False, show_label=False, image_mso='FillDown' , label="Innenabstand oben",   screentip='Innenabstand oben',   supertip='Ändere den oberen Innenabstand des Textfelds auf das angegebene Maß (in cm).')
inner_margin_bottom = InnerMargin(attr="MarginBottom", id='textFrameMargin-bottom-2', image_button=False, show_label=False, image_mso='FillUp'   , label="Innenabstand unten",  screentip='Innenabstand unten',  supertip='Ändere den unteren Innenabstand des Textfelds auf das angegebene Maß (in cm).')
inner_margin_left   = InnerMargin(attr="MarginLeft",   id='textFrameMargin-left-2',   image_button=False, show_label=False, image_mso='FillRight', label="Innenabstand links",  screentip='Innenabstand links',  supertip='Ändere den linken Innenabstand des Textfelds auf das angegebene Maß (in cm).')
inner_margin_right  = InnerMargin(attr="MarginRight",  id='textFrameMargin-right-2',  image_button=False, show_label=False, image_mso='FillLeft' , label="Innenabstand rechts", screentip='Innenabstand rechts', supertip='Ändere den rechten Innenabstand des Textfelds auf das angegebene Maß (in cm).')



innenabstand_gruppe = bkt.ribbon.Group(
    id="bkt_innermargin_group",
    label="Textfeld Innenabstand",
    image_mso='ObjectNudgeRight',
    children=[
    bkt.ribbon.Box(id='textFrameMargin-box-top', children=[
        bkt.ribbon.LabelControl(id='textFrameMargin-label-top', label='         \u200b'),
        #create_margin_spinner('MarginTop', imageMso='ObjectNudgeDown'),
        inner_margin_top,
        bkt.ribbon.LabelControl(label='   \u200b'),
        bkt.ribbon.Button(
            id='textFrameMargin-zero',
            label="=\u202F0",
            screentip="Innenabstand auf Null",
            supertip="Ändere in Innenabstand des Textfelds an allen Seiten auf Null.",
            on_action=bkt.Callback( InnerMargin.set_to_0, shapes=True, context=True )
        )
    ]),
    bkt.ribbon.Box(id='textFrameMargin-box-LR', children=[
        #create_margin_spinner('MarginLeft',  imageMso='ObjectNudgeRight'),
        #create_margin_spinner('MarginRight', imageMso='ObjectNudgeLeft')
        inner_margin_left,
        #bkt.ribbon.LabelControl(label=u' '),
        inner_margin_right,
    ]),
    bkt.ribbon.Box(id='textFrameMargin-box-bottom', children=[
        bkt.ribbon.LabelControl(id='textFrameMargin-label-bottom', label='         \u200b'),
        #create_margin_spinner('MarginBottom', imageMso='ObjectNudgeUp'),
        inner_margin_bottom,
        bkt.ribbon.LabelControl(label='   \u200b'),
        bkt.ribbon.ToggleButton(
            id='textFrameMargin-equal',
            label="==",
            screentip="Einheitlicher Innenabstand",
            supertip="Bei Änderung des Textfeld-Innenabstand einer Seite wird der Innenabstand aller Seiten geändert.",
            on_toggle_action=bkt.Callback( InnerMargin.toggle_all_equal ),
            get_pressed=bkt.Callback( InnerMargin.get_all_equal )
        )
    ]),
    bkt.ribbon.DialogBoxLauncher(idMso='TextAlignMoreOptionsDialog')
    #bkt.ribbon.DialogBoxLauncher(idMso='WordArtFormatDialog')
])



class ParSpaceBefore(pplib.ParagraphFormatSpinnerBox):
    attr = 'SpaceBefore'
    _attributes = dict(
        label="Absatzabstand oben",
        image_mso='WordOpenParaAbove',
        screentip="Oberen Absatzabstand",
        supertip="Ändere den Absatzabstand vor dem Absatz auf das angegebene Maß (entweder in Abstand Zeilen oder in pt).",
    )

class ParSpaceAfter(pplib.ParagraphFormatSpinnerBox):
    attr = 'SpaceAfter'
    _attributes = dict(
        label="Absatzabstand unten",
        image_mso='WordOpenParaBelow',
        screentip="Unteren Absatzabstand",
        supertip="Ändere den Absatzabstand hinter dem Absatz auf das angegebene Maß (entweder in Abstand Zeilen oder in pt).",
    )

class ParSpaceWithin(pplib.ParagraphFormatSpinnerBox):
    attr = 'SpaceWithin'
    _attributes = dict(
        label="Zeilenabstand",
        image_mso='LineSpacing',
        screentip="Zeilenabstand",
        supertip="Ändere den Zeilenabstand (entweder in Abstand Zeilen oder in pt).",
        fallback_value = 1,
    )

class ParFirstLineIndent(pplib.ParagraphFormatSpinnerBox):
    attr = 'FirstLineIndent'
    _attributes = dict(
        label="Einzug 1. Zeile",
        image='first_line_indent',
        screentip="Sondereinzug",
        supertip="Ändere den Sondereinzug (1. Zeile, hängend) auf das angegebene Maß (in cm).",
    )

class ParLeftIndent(pplib.ParagraphFormatSpinnerBox):
    attr = 'LeftIndent'
    _attributes = dict(
        label="Einzug links",
        image_mso='ParagraphIndentLeft',
        screentip="Absatzeinzug links",
        supertip="Ändere den linken Absatzeinzug auf das angegebene Maß (in cm).",
    )

class ParRightIndent(pplib.ParagraphFormatSpinnerBox):
    attr = 'RightIndent'
    _attributes = dict(
        label="Einzug rechts",
        image_mso='ParagraphIndentRight',
        screentip="Absatzeinzug rechts",
        supertip="Ändere den rechten Absatzeinzug auf das angegebene Maß (in cm).",
    )


class Absatz(object):

    @staticmethod
    def set_word_wrap(shapes, pressed):
        for textframe in pplib.iterate_shape_textframes(shapes):
            try:
                textframe.WordWrap = -1 if pressed else 0
            except:
                continue

    @staticmethod
    def get_word_wrap(shapes):
        for textframe in pplib.iterate_shape_textframes(shapes):
            try:
                return (textframe.WordWrap == -1) # msoTrue
            except:
                continue
        return None


    @staticmethod
    def set_auto_size(shapes, pressed):
        for textframe in pplib.iterate_shape_textframes(shapes):
            try:
                textframe.AutoSize = 1 if pressed else 0
                # 1 = ppAutoSizeShapeToFitText
                # 0 = ppAutoSizeNone
            except:
                continue

    @staticmethod
    def get_auto_size(shapes):
        for textframe in pplib.iterate_shape_textframes(shapes):
            try:
                return (textframe.AutoSize == 1)
            except:
                continue
        return None

    # def set_par_indent(self, shapes, value):
    #     # pt_value = cm_to_pt(value)
    #     # delta = pt_value - shapes[0].TextFrame.Ruler.Levels(1).LeftMargin
    #     for shape in shapes:
    #         shape.TextFrame.Ruler.Levels(1).LeftMargin = cm_to_pt(value)
    #         # shape.TextFrame.Ruler.Levels(1).LeftMargin = pt_value
    #         # shape.TextFrame.Ruler.Levels(1).LeftMargin  = shp.TextFrame.Ruler.Levels(1).LeftMargin + delta
    #
    # def get_par_indent(self, shapes):
    #     return round(pt_to_cm(shapes[0].TextFrame.Ruler.Levels(1).LeftMargin), 2)

    # @staticmethod
    # def set_par_sep_before(shapes, selection, value):
    #     value = max(0,value)
    #     if selection.Type == 2:
    #         # shapes selected
    #         for shape in shapes:
    #             # distance in points, not in number of lines
    #             shape.TextFrame.TextRange.ParagraphFormat.LineRuleBefore = 0
    #             # set distance
    #             shape.TextFrame.TextRange.ParagraphFormat.SpaceBefore = value
    #     elif selection.Type == 3:
    #         # text selected
    #         selection.TextRange2.ParagraphFormat.LineRuleBefore = 0
    #         selection.TextRange2.ParagraphFormat.SpaceBefore = value 

    # @staticmethod
    # def get_par_sep_before(shapes, selection):
    #     if selection.Type == 2:
    #         # shapes selected
    #         return shapes[0].TextFrame.TextRange.ParagraphFormat.SpaceBefore
    #     elif selection.Type == 3:
    #         # text selected
    #         try:
    #             # produces error if no text is selected
    #             return selection.TextRange2.Paragraphs(1,1).ParagraphFormat.SpaceBefore
    #         except:
    #             return selection.TextRange2.ParagraphFormat.SpaceBefore


    # @staticmethod
    # def set_par_sep_after(shapes, selection, value):
    #     value = max(0,value)
    #     if selection.Type == 2:
    #         # shapes selected
    #         for shape in shapes:
    #             # distance in points, not in number of lines
    #             shape.TextFrame.TextRange.ParagraphFormat.LineRuleAfter = 0
    #             # set distance
    #             shape.TextFrame.TextRange.ParagraphFormat.SpaceAfter = value
    #     elif selection.Type == 3:
    #         # text selected
    #         selection.TextRange2.ParagraphFormat.LineRuleAfter = 0
    #         selection.TextRange2.ParagraphFormat.SpaceAfter = value 

    # @staticmethod
    # def get_par_sep_after(shapes, selection):
    #     if selection.Type == 2:
    #         # shapes selected
    #         return shapes[0].TextFrame.TextRange.ParagraphFormat.SpaceAfter
    #     elif selection.Type == 3:
    #         # text selected
    #         try:
    #             # produces error if no text is selected
    #             return selection.TextRange2.Paragraphs(1,1).ParagraphFormat.SpaceAfter
    #         except:
    #             return selection.TextRange2.ParagraphFormat.SpaceAfter
    
    
    # @staticmethod
    # def set_left_indent(shapes, selection, value):
    #     # FIXME: apply text-selection-logic from set_par_sep_after
    #     if type(value) == str:
    #         value = float(value.replace(',', '.'))
    #     value = float(value) / pt_to_cm_factor
        
    #     if selection.Type == 2:
    #         # shapes selected
    #         for shape in shapes:
    #             shape.TextFrame2.TextRange.ParagraphFormat.LeftIndent = value
    #     elif selection.Type == 3:
    #         # text selected
    #         selection.TextRange2.ParagraphFormat.LeftIndent = value


    # @staticmethod
    # def get_left_indent(shapes, selection):
    #     if selection.Type == 2:
    #         # shapes selected
    #         value = shapes[0].TextFrame2.TextRange.ParagraphFormat.LeftIndent
    #     elif selection.Type == 3:
    #         # text selected
    #         try:
    #             # produces error if no text is selected
    #             value = selection.TextRange2.Paragraphs(1,1).ParagraphFormat.LeftIndent 
    #         except:
    #             value = selection.TextRange2.ParagraphFormat.LeftIndent 

    #     return round(value * pt_to_cm_factor, 2)
    
    
    # @staticmethod
    # def set_first_line_indent(shapes, selection, value):
    #     if type(value) == str:
    #         value = float(value.replace(',', '.'))
    #     value = float(value) / pt_to_cm_factor
        
    #     if selection.Type == 2:
    #         # shapes selected
    #         for shape in shapes:
    #             shape.TextFrame2.TextRange.ParagraphFormat.FirstLineIndent = value
    #     elif selection.Type == 3:
    #         # text selected
    #         selection.TextRange2.ParagraphFormat.FirstLineIndent = value

    # @staticmethod
    # def get_first_line_indent(shapes, selection):
    #     if selection.Type == 2:
    #         # shapes selected
    #         value = shapes[0].TextFrame2.TextRange.ParagraphFormat.FirstLineIndent
    #     elif selection.Type == 3:
    #         # text selected
    #         try:
    #             # produces error if no text is selected
    #             value = selection.TextRange2.Paragraphs(1,1).ParagraphFormat.FirstLineIndent 
    #         except:
    #             value = selection.TextRange2.ParagraphFormat.FirstLineIndent 
        
    #     return round(value * pt_to_cm_factor, 2)



text_menu = bkt.ribbon.Menu(
    label="Textfeld zeichnen Menü",
    supertip="Sticker einfügen, Bullet Points angleichen, sowie weitere Text-bezogene Funktionen",
    children=[
        bkt.ribbon.MenuSeparator(title="Textformen einfügen"),
        bkt.mso.control.TextBoxInsert,
        bkt.ribbon.DynamicMenu(
            id="sticker_splitbutton", #2023-01-26 not a splitbutton anymore to make it dynamic
            label="Sticker Menü",
            supertip="Verschiedene Sticker einfügen",
            get_content=bkt.CallbackLazy("toolbox.models.text_menu", "sticker_menu")
        ),
        bkt.ribbon.Button(
            id = 'underlined_textbox',
            label = "Unterstrichene Textbox",
            image = "underlined_textbox",
            screentip="Unterstrichene Textbox einfügen",
            supertip="Füge eine Textbox mit Linie unten am Shape auf dem aktuellen Slide ein.",
            on_action=bkt.CallbackLazy("toolbox.models.text_menu", "TextShapes", "addUnderlinedTextbox", slide=True, presentation=True)
        ),
        bkt.ribbon.MenuSeparator(title="Aufzählungszeichen"),
        bkt.ribbon.Button(
            id="bullet_fixing",
            label="Aufzählungszeichen korrigieren",
            supertip="Aufzählungszeichen werden korrigiert. Der Stil wird vom Textplatzhalter auf dem Masterslide übernommen. Betrifft: Symbol, Symbol-/Textfarbe, Absatzeinzug/-abstand",
            image_mso='MultilevelListGallery',
            on_action=bkt.CallbackLazy("toolbox.models.text_menu", "BulletStyle", "fix_bullet_style", shapes=True),
            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
        ),
        bkt.ribbon.MenuSeparator(),
        bkt.ribbon.ColorGallery(
            id = 'bullet_color',
            label='Farbe ändern',
            screentip="Bullet Point Farbe ändern",
            supertip="Ändert die Farbe der gewählten Bullet Points.",
            on_rgb_color_change = bkt.CallbackLazy("toolbox.models.text_menu", "BulletStyle", "set_bullet_color_rgb", selection=True, shapes=True),
            on_theme_color_change = bkt.CallbackLazy("toolbox.models.text_menu", "BulletStyle", "set_bullet_theme_color", selection=True, shapes=True),
            get_selected_color = bkt.CallbackLazy("toolbox.models.text_menu", "BulletStyle", "get_bullet_color_rgb", selection=True, shapes=True),
            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
            children=[
                bkt.ribbon.Button(
                    id="bullet_color_auto",
                    label="Automatisch",
                    screentip="Bullet Point Farbe automatisch bestimmen",
                    supertip="Bullet Point Farbe wird automatisch anhand der Textfarbe bestimmt.",
                    on_action=bkt.CallbackLazy("toolbox.models.text_menu", "BulletStyle", "set_bullet_color_auto", selection=True, shapes=True),
                    image_mso="ColorBlack",
                ),
            ]
        ),
        bkt.ribbon.SymbolsGallery(
            id="bullet_symbol",
            label="Symbol ändern",
            screentip="Bullet Point Symbol ändern",
            supertip="Ändert das Symbol der gewählten Bullet Points.",
            symbols = Characters.lists,
            on_symbol_change = bkt.CallbackLazy("toolbox.models.text_menu", "BulletStyle", "set_bullet_symbol", selection=True, shapes=True),
            get_selected_symbol = bkt.CallbackLazy("toolbox.models.text_menu", "BulletStyle", "get_bullet_symbol", selection=True, shapes=True),
            get_enabled = bkt.apps.ppt_shapes_or_text_selected
        ),
        bkt.ribbon.MenuSeparator(title="Textoperationen"),
        bkt.ribbon.Button(
            id = 'text_in_shape',
            label = "Text in Shape",
            image_mso = "TextBoxInsert",
            screentip="Text in Shape kombinieren",
            supertip="Kopiere den Text eines Text-Shapes in das zweite markierte Shape und löscht das Text-Shape.",
            on_action=bkt.CallbackLazy("toolbox.models.text_menu", "TextOnShape", "textIntoShape", shapes=True, shapes_min=2),
            get_enabled = bkt.apps.ppt_shapes_min2_selected,
        ),
        bkt.ribbon.Button(
            id = 'text_on_shape',
            label = "Text auf Shape",
            image_mso = "TableCellCustomMarginsDialog",
            screentip="Text auf Shape zerlegen",
            supertip="Überführe jeweils den Textinhalt der markierten Shapes in ein separates Text-Shape.",
            on_action=bkt.CallbackLazy("toolbox.models.text_menu", "TextOnShape", "textOutOfShape", shapes=True, slide=True),
            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
        ),
        bkt.ribbon.MenuSeparator(),
        bkt.ribbon.Button(
            id = 'decompose_text',
            label = "Shape-Text zerlegen",
            image_mso = "TraceDependents",
            supertip="Zerlege die markierten Shapes anhand der Text-Absätze in mehrere Shapes. Pro Absatz wird ein Shape mit dem entsprechenden Text angelegt.",
            on_action=bkt.CallbackLazy("toolbox.models.text_menu", "SplitTextShapes", "splitShapesByParagraphs", shapes=True, context=True),
            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
        ),
        bkt.ribbon.Button(
            id = 'compose_text',
            label = "Shape-Text zusammenführen",
            image_mso = "TracePrecedents",
            supertip="Führe die markierten Shapes in ein Shape zusammen. Der Text aller Shapes wird übernommen und aneinandergehängt.",
            on_action=bkt.CallbackLazy("toolbox.models.text_menu", "SplitTextShapes", "joinShapesWithText", shapes=True, shapes_min=2),
            get_enabled = bkt.apps.ppt_shapes_min2_selected,
        ),
        bkt.ribbon.MenuSeparator(),
        bkt.ribbon.Button(
            id = 'text_truncate',
            label="Shape-Texte löschen",
            image_mso='ReviewDeleteMarkup',
            supertip="Text aller gewählten Shapes löschen.",
            on_action=bkt.CallbackLazy("toolbox.models.text_menu", "TextPlaceholder", "text_truncate", shapes=True),
            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
        ),
        bkt.ribbon.SplitButton(
            id = 'text_replace_splitbutton',
            get_enabled=bkt.apps.ppt_shapes_or_text_selected,
            children=[
                bkt.ribbon.Button(
                    id = 'text_replace',
                    label="Shape-Texte ersetzen…",
                    image_mso='ReplaceDialog',
                    supertip="Text aller gewählten Shapes mit im Dialogfeld eingegebenen Text ersetzen.",
                    on_action=bkt.CallbackLazy("toolbox.models.text_menu", "TextPlaceholder", "text_replace", shapes=True, presentation=True),
                    get_enabled=bkt.apps.ppt_shapes_or_text_selected,
                ),
                bkt.ribbon.Menu(label="Shape-Texte ersetzen Menü", supertip="Text mit vordefinierten Standard-Platzhaltern ersetzen", children=[
                    bkt.ribbon.Button(
                        id = 'text_tbd',
                        label="… mit »tbd«",
                        image_mso='TextDialog',
                        screentip="Text mit »tbd« ersetzen",
                        supertip="Text aller gewählten Shapes mit 'tbd' ersetzen.",
                        on_action=bkt.CallbackLazy("toolbox.models.text_menu", "TextPlaceholder", "text_tbd", shapes=True),
                    ),
                    bkt.ribbon.Button(
                        id = 'text_lorem',
                        label="… mit Lorem ipsum",
                        image_mso='TextDialog',
                        screentip="Text mit Lorem ipsum ersetzen",
                        supertip="Text aller gewählten Shapes mit langem 'Lorem ipsum'-Platzhaltertext ersetzen.",
                        on_action=bkt.CallbackLazy("toolbox.models.text_menu", "TextPlaceholder", "text_lorem", shapes=True),
                    ),
                    bkt.ribbon.Button(
                        id = 'text_counter',
                        label="… mit Nummerierung",
                        image_mso='TextDialog',
                        screentip="Text mit Nummerierung ersetzen",
                        supertip="Text aller gewählten Shapes durch Nummerierung ersetzen.",
                        on_action=bkt.CallbackLazy("toolbox.models.text_menu", "TextPlaceholder", "text_counter", shapes=True),
                    ),
                    bkt.ribbon.MenuSeparator(),
                    bkt.ribbon.Button(
                        id = 'text_replace2',
                        label="… mit benutzerdefiniertem Text",
                        image_mso='ReplaceDialog',
                        screentip="Text mit eigener Eingabe ersetzen",
                        supertip="Text aller gewählten Shapes mit im Dialogfeld eingegebenen Text ersetzen.",
                        on_action=bkt.CallbackLazy("toolbox.models.text_menu", "TextPlaceholder", "text_replace", shapes=True, presentation=True),
                        get_enabled=bkt.apps.ppt_shapes_or_text_selected,
                    ),
                ])
            ]
        ),
    ]
)


class TextBox(object):
    @staticmethod
    def textbox_insert(context, pressed):
        from .models.text_menu import TextShapes
        if bkt.get_key_state(bkt.KeyCodes.SHIFT):
            TextShapes.addUnderlinedTextbox(context.slide, context.presentation)
        elif bkt.get_key_state(bkt.KeyCodes.CTRL):
            TextShapes.addSticker(context.slide, context.presentation)
        else:
            # NOTE: idMso is different on some machines, see: https://answers.microsoft.com/en-us/msoffice/forum/msoffice_powerpoint-msoffice_custom-mso_2007/powerpoint-2007-textboxinsert-vs/52f12b52-7e1c-4d7c-86a7-bded312437b0
            try:
                context.app.commandbars.ExecuteMso("TextBoxInsert")
            except:
                context.app.commandbars.ExecuteMso("TextBoxInsertHorizontal")
    
    @staticmethod
    def textbox_enabled(context):
        try:
            return context.app.commandbars.GetEnabledMso("TextBoxInsert")
        except:
            return context.app.commandbars.GetEnabledMso("TextBoxInsertHorizontal")
    
    @staticmethod
    def textbox_pressed(context):
        try:
            return context.app.commandbars.GetPressedMso("TextBoxInsert")
        except:
            return context.app.commandbars.GetPressedMso("TextBoxInsertHorizontal")


text_splitbutton = bkt.ribbon.SplitButton(
    id="textbox_insert_splitbutton",
    show_label=False,
    children=[
        bkt.ribbon.ToggleButton(
            id="textbox_insert",
            label="Textfeld zeichnen",
            image_mso="TextBoxInsert",
            supertip="Zeichnen Sie ein Textfeld an einer beliebigen Stelle.\n\nMit gedrückter Umschalt-Taste wird eine unterstrichene Textbox eingefügt.\n\nMit gedrückter Strg-Taste wird ein Sticker eingefügt.",
            on_toggle_action=bkt.Callback(TextBox.textbox_insert, context=True),
            get_pressed=bkt.Callback(TextBox.textbox_pressed, context=True),
            get_enabled=bkt.Callback(TextBox.textbox_enabled, context=True),
        ),
        # bkt.mso.toggleButton.TextBoxInsert,
        text_menu
    ]
)


paragraph_group = bkt.ribbon.Group(
    id="bkt_paragraph_group",
    label = "Absatz",
    image_mso='FormattingMarkDropDown',
    children = [
        bkt.ribbon.Menu(
            label="Einst.",
            imageMso="FormattingMarkDropDown",
            supertip="Einstellungen für die Textbox ändern",
            children = [
                bkt.ribbon.ToggleButton(
                    id = 'wordwrap',
                    label="WordWrap",
                    image_mso="FormattingMarkDropDown",
                    screentip="Text in Form umbrechen",
                    supertip="Konfiguriere die Textoption auf 'Text in Form umbrechen'.",
                    on_toggle_action=bkt.Callback(Absatz.set_word_wrap , shapes=True),
                    get_pressed=bkt.Callback(Absatz.get_word_wrap , shapes=True),
                    get_enabled = bkt.apps.ppt_selection_contains_textframe,
                ),
                bkt.ribbon.ToggleButton(
                    id = 'autosize',
                    label="AutoSize",
                    image_mso="SmartArtLargerShape",
                    screentip="Größe der Form anpassen",
                    supertip="Konfiguriere die Textoption auf 'Größe der Form dem Text anpassen' bzw. 'Größe nicht automatisch anpassen'.",
                    on_toggle_action=bkt.Callback(Absatz.set_auto_size , shapes=True),
                    get_pressed=bkt.Callback(Absatz.get_auto_size , shapes=True),
                    get_enabled = bkt.apps.ppt_selection_contains_textframe,
                ),
                bkt.ribbon.MenuSeparator(),
                bkt.mso.control.TextAlignMoreOptionsDialog
            ]
        ),
        ParSpaceBefore(
            id = 'par_sep_top',
            show_label=False,
            size_string = '##',
            # label=u"Absatzabstand oben",
            # image_mso='WordOpenParaAbove',
            # screentip="Oberen Absatzabstand",
            # supertip="Ändere den Absatzabstand vor dem Absatz auf das angegebene Maß (in pt).",
            #attr='SpaceBefore'
        ),
        ParSpaceAfter(
            id = 'par_sep_bottom',
            show_label=False,
            size_string = '##',
            # label=u"Absatzabstand unten",
            # image_mso='WordOpenParaBelow',
            # screentip="Unteren Absatzabstand",
            # supertip="Ändere den Absatzabstand hinter dem Absatz auf das angegebene Maß (in pt).",
            #attr='SpaceAfter'
        ),
        bkt.ribbon.DialogBoxLauncher(idMso='PowerPointParagraphDialog')
    ]
)

paragraph_indent_group = bkt.ribbon.Group(
    id="bkt_paragraph_adv_group",
    label = "Absatzeinzug",
    image_mso='ViewRulerPowerPoint',
    #ViewRulerPowerPoint
    children = [
        ParFirstLineIndent(
            id = 'first_line_indent',
            show_label=False,
            # label=u"Einzug 1. Zeile",
            # image='first_line_indent',
            # screentip="Sondereinzug",
            # supertip="Ändere den Sondereinzug (1. Zeile, hängend) auf das angegebene Maß (in cm).",
            # attr='FirstLineIndent',
            # big_step = 0.25,
            # small_step = 0.125,
            # rounding_factor = 0.125,
            # size_string = '-###',
        ),
        ParLeftIndent(
            id = 'left_indent',
            show_label=False,
            # label=u"Einzug links",
            # image_mso='IndentClassic',
            # screentip="Absatzeinzug",
            # supertip="Ändere den Absatzeinzug auf das angegebene Maß (in cm).",
            # attr='LeftIndent',
            # big_step = 0.25,
            # small_step = 0.125,
            # rounding_factor = 0.125,
            # size_string = '-###',
        ),
        ParRightIndent(
            id = 'right_indent',
            show_label=False,
            # label=u"Einzug links",
            # image_mso='IndentClassic',
            # screentip="Absatzeinzug",
            # supertip="Ändere den Absatzeinzug auf das angegebene Maß (in cm).",
            # attr='LeftIndent',
            # big_step = 0.25,
            # small_step = 0.125,
            # rounding_factor = 0.125,
            # size_string = '-###',
        ),
        ParSpaceWithin(
            id = 'par_sep_within',
            show_label=False,
            # label=u"Zeilenabstand",
            # image_mso='LineSpacing',
            # screentip="Zeilenabstand",
            # supertip="Ändere den Zeilenabstand (entweder in Abstand Zeilen oder in pt).",
            # attr='SpaceWithin',
            # size_string = '-###',
            # fallback_value = 1,
        ),
        bkt.ribbon.CheckBox(
            id = 'wordwrap2',
            label="WordWrap",
            # image_mso="FormattingMarkDropDown",
            screentip="Text in Form umbrechen",
            supertip="Konfiguriere die Textoption auf 'Text in Form umbrechen'.",
            on_toggle_action=bkt.Callback(Absatz.set_word_wrap , shapes=True, require_text=True),
            get_pressed=bkt.Callback(Absatz.get_word_wrap , shapes=True, require_text=True),
            get_enabled = bkt.get_enabled_auto,
        ),
        bkt.ribbon.CheckBox(
            id = 'autosize2',
            label="AutoSize",
            # image_mso="SmartArtLargerShape",
            screentip="Größe der Form anpassen",
            supertip="Konfiguriere die Textoption auf 'Größe der Form dem Text anpassen' bzw. 'Größe nicht automatisch anpassen'.",
            on_toggle_action=bkt.Callback(Absatz.set_auto_size , shapes=True, require_text=True),
            get_pressed=bkt.Callback(Absatz.get_auto_size , shapes=True, require_text=True),
            get_enabled = bkt.get_enabled_auto,
        ),
        bkt.ribbon.DialogBoxLauncher(idMso='PowerPointParagraphDialog')
    ]
)


compact_font_group = bkt.ribbon.Group(
    id="bkt_compact_font_group",
    label = "Schriftart",
    image_mso='GroupFont',
    children = [
        #NOTE: horizontal box layout leads to spacing between Font and FontSize ComboBox!
        bkt.mso.comboBox.Font(sizeString="WWWWWWWI"),
        bkt.ribbon.ButtonGroup(children=[
            bkt.mso.control.Bold,
            bkt.mso.control.Italic,
            bkt.mso.control.Underline,
            # bkt.mso.control.Shadow,
            bkt.mso.control.Strikethrough,
        ]),
        bkt.ribbon.Box(box_style="horizontal", children=[
            bkt.mso.control.CharacterSpacingGallery,
            bkt.mso.control.ChangeCaseGallery,
            bkt.mso.control.ClearFormatting,
        ]),

        bkt.mso.control.FontSize,
        bkt.ribbon.ButtonGroup(children=[
            bkt.mso.control.FontSizeIncrease,
            bkt.mso.control.FontSizeDecrease,
        ]),
        bkt.ribbon.ButtonGroup(children=[
            bkt.mso.control.Superscript,
            bkt.mso.control.Subscript,
        ]),
        bkt.ribbon.DialogBoxLauncher(idMso='FontDialogPowerPoint')
    ]
)

compact_paragraph_group = bkt.ribbon.Group(
    id="bkt_compact_paragraph_group",
    label = "Absatz",
    image_mso='GroupParagraph',
    children = [
        bkt.ribbon.Box(box_style="horizontal", children=[
            bkt.mso.control.BulletsGallery,
            bkt.mso.control.NumberingGallery,
            bkt.ribbon.ButtonGroup(children=[
                bkt.mso.control.IndentDecrease,
                bkt.mso.control.IndentIncrease,
            ]),
            # bkt.mso.control.ConvertToSmartArt,
        ]),
        bkt.ribbon.Box(box_style="horizontal", children=[
            bkt.ribbon.ButtonGroup(children=[
                bkt.mso.control.AlignLeft,
                bkt.mso.control.AlignCenter,
                bkt.mso.control.AlignRight,
                bkt.mso.control.AlignJustify,
                bkt.mso.control.AlignJustifyMenu,
            ]),
            # bkt.mso.control.ParagraphDistributed,
            # bkt.mso.control.AlignJustifyThai,
            # bkt.mso.control.TextDirectionLeftToRight,
            # bkt.mso.control.TextDirectionRightToLeft,
            bkt.mso.control.TableColumnsGallery,
        ]),

        bkt.ribbon.Box(box_style="horizontal", children=[
            bkt.mso.control.LineSpacingGalleryPowerPoint,
            # bkt.mso.control.FontColorPicker,
            bkt.mso.control.TextDirectionGallery,
            bkt.mso.control.TextAlignGallery,
        ]),
        bkt.ribbon.DialogBoxLauncher(idMso='PowerPointParagraphDialog')
    ]
)