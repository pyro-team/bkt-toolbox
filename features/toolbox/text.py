# -*- coding: utf-8 -*-
'''
Created on 06.07.2016

@author: rdebeerst
'''

from __future__ import absolute_import

import bkt
import bkt.library.powerpoint as pplib

# import fontawesome


# pt_to_cm_factor = 2.54 / 72;
# def pt_to_cm(pt):
#     return float(pt) * pt_to_cm_factor;
# def cm_to_pt(cm):
#     return float(cm) / pt_to_cm_factor;

class TextPlaceholder(object):
    recent_placeholder = bkt.settings.get("toolbox.recent_placeholder", "…")
    #labels for counter, but max 0..25
    label_a = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z']
    label_A = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
    label_I = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI', 'XII', 'XIII', 'XIV', 'XV', 'XVI', 'XVII', 'XVIII', 'XIX', 'XX', 'XXI', 'XXII', 'XXIII', 'XXIV', 'XXV', 'XXVI']

    @staticmethod
    def set_text_for_shape(textframe, text=None): #None=delete text
        if text is not None:
            if type(text) == list:
                if text[0] != '':
                    textframe.TextRange.InsertBefore( text[0] )
                if text[1] != '':
                    textframe.TextRange.InsertAfter( text[1] )
            else:
                textframe.TextRange.Text = text
        else:
            textframe.TextRange.Delete()

    @classmethod
    def text_truncate(cls, shapes):
        for textframe in pplib.iterate_shape_textframes(shapes):
            cls.set_text_for_shape(textframe, None)

    @classmethod
    def text_replace(cls, shapes):
        input_text = bkt.ui.show_user_input("Text eingeben, der auf alle Shapes angewendet werden soll (Platzhalter [text] kann für bestehenden Text und [counter], [counter-a], [counter-A], [counter-I] zur Nummerierung verwendet werden):", "Text ersetzen", cls.recent_placeholder, True)

        # user_form = bkt.ui.UserInputBox("Text eingeben, der auf alle Shapes angewendet werden soll (Platzhalter [counter] kann zur Nummerierung verwendet werden):", "Text ersetzen")
        # user_form._add_textbox("new_text", "…", True)
        # user_form._add_checkbox("keep_text", "Text anhängen/ Vorhandenen Text erhalten")
        # form_return = user_form.show()
        # if len(form_return) == 0:
        if input_text is None:
            return
        cls.recent_placeholder = bkt.settings["toolbox.recent_placeholder"] = input_text

        placeholder_count = input_text.count("[text]")
        counter = 1
        for textframe in pplib.iterate_shape_textframes(shapes):
            new_text = input_text.replace("[counter]", str(counter))
            new_text = new_text.replace("[counter-a]", cls.label_a[counter-1%26])
            new_text = new_text.replace("[counter-A]", cls.label_A[counter-1%26])
            new_text = new_text.replace("[counter-I]", cls.label_I[counter-1%26])
            
            if placeholder_count > 1:
                #replace placeholder with text, might loose existing formatting
                new_text = new_text.replace("[text]", textframe.TextRange.Text)
            elif placeholder_count == 1:
                #only one occurence of text-placeholder, make use of InsertBefore/After to keep formatting
                new_text = new_text.split("[text]", 1)
            
            cls.set_text_for_shape(textframe, new_text)
            
            # TextPlaceholder.set_text_for_shape(textframe, new_text, form_return["keep_text"])
            counter += 1

    @classmethod
    def text_tbd(cls, shapes):
        for textframe in pplib.iterate_shape_textframes(shapes):
            cls.set_text_for_shape(textframe, "tbd")

    @classmethod
    def text_counter(cls, shapes):
        counter = 1
        for textframe in pplib.iterate_shape_textframes(shapes):
            cls.set_text_for_shape(textframe, str(counter))
            counter += 1

    @classmethod
    def text_lorem(cls, shapes):
        lorem_text = '''Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat, sed diam voluptua.
At vero eos et accusam et justo duo dolores et ea rebum. Stet clita kasd gubergren, no sea takimata sanctus est Lorem ipsum dolor sit amet.
Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat, sed diam voluptua.
At vero eos et accusam et justo duo dolores et ea rebum. Stet clita kasd gubergren, no sea takimata sanctus est Lorem ipsum dolor sit amet.'''
        lorem_text = bkt.ui.endings_to_windows(lorem_text)

        for textframe in pplib.iterate_shape_textframes(shapes):
            cls.set_text_for_shape(textframe, lorem_text)

    @staticmethod
    def remove_placeholders(slides):
        for sld in slides:
            for plh in list(iter(sld.Shapes.Placeholders)):
                if plh.HasTextFrame == -1 and plh.TextFrame2.HasText == 0:
                    #placeholder is a text placeholder and has no text. note: placeholder can also be a picture, table or diagram without text!
                    plh.Delete()


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
        selection.TextRange2.text=u'\xad'

    @staticmethod
    def add_protected_space(selection):
        selection.TextRange2.text=u'\xa0'

    @staticmethod
    def add_protected_narrow_space(selection):
        selection.TextRange2.text=u'\u202F'

    ### TYPOGRAPHY ###
    typography = [
        (None, u'\xbb', "Linkes Guillemets"),
        (None, u'\xab', "Rechtes Guillemets"),
        (None, u'\xb6', "Paragraph"),
        (None, u'\u2026', "Auslassungspunkte (Ellipse)", "Auslassungspunkte sind drei kurz aufeinanderfolgende Punkte. Meistens zeigen diese eine Ellipse (Auslassung eines Textteils) an."),
        (None, u'\u2013', "Gedankenstrich (Halbgeviertstrich/En-Dash)", "Ein Gedankenstrich (sog. Halbgeviertstrich) wie er von Office teilweise automatisch gesetzt wird. Verwendet als Bis-Strich oder Streckenstrich."),
        (None, u'\u2014', "Waagerechter Strich (Geviertstrich/Em-Dash)"),
        (None, u'\u2020', "Kreuz"),
        (None, u'\u2021', "Doppelkreuz"),
        (None, u'\u25A0', "Schwarzes Quadrat"),
        (None, u'\u25A1', "Weißes Quadrat"),
        (None, u'\u2423', "Leerzeichen-Symbol"),
        (None, u'\xa9',   "Copyright"),
        (None, u'\xae',   "Registered Trade Mark"),
        (None, u'\u2122', "Trade Mark"),
        (None, u'\u2030', "Per mil"),
        (None, u'\u20AC', "Euro"),
        (None, u'\u1E9E', "Großes Eszett"),
    ]

    ### MATH ###
    math = [
        (None, u'\xb1',   "Plus-Minus-Zeichen", "Ein Plus-Minus-Zeichen einfügen."),
        (None, u'\u2212', "Echtes Minuszeichen", "Ein echtes Minuszeichen (kein Bindestrich) einfügen."),
        (None, u'\xd7',   "Echtes Malzeichen (Kreuz)", "Ein echtes Kreuz-Multiplikatorzeichen einfügen."),
        (None, u'\u22c5', "Echtes Malzeichen (Punkt)", "Ein echtes Punkt-Multiplikatorzeichen einfügen."),
        (None, u'\u2044', "Echter Bruchstrich", "Einen echten Bruchstrich (ähnlich Schrägstrich) einfügen."),
        (None, u'\u2248', "Ungefähr Gleich", "Ein Ungefähr Gleich Zeichen einfügen."),
        (None, u'\u2260', "Ungleich", "Ein Ungleich-Zeichen einfügen."),
        (None, u'\u2206', "Delta", "Ein Deltazeichen einfügen."), #alt: \u0394 griechisches Delta
        (None, u'\u2300', "Mittelwert/Durchmesser", "Ein Durchmesserzeichen bzw. Durchschnittszeichen einfügen."), #alt: \xD8 leere menge
        (None, u'\u2211', "Summenzeichen", "Ein Summenzeichen einfügen."),
        (None, u'\u221A', "Wurzelzeichen", "Ein Wurzelzeichen einfügen."),
        (None, u'\u221E', "Unendlich-Zeichen", "Ein Unendlich-Zeichen (liegende Acht) einfügen."),
        (None, u'\u2264', "Kleiner-Gleich", "Ein kleiner oder gleich Zeichen einfügen."),
        (None, u'\u2265', "Größer-Gleich", "Ein größer oder gleich Zeichen einfügen."),
    ]

    ### LIST ###
    lists = [
        (None, u'\u2022', "Aufzählungszeichen (Kreis)", "Ein Aufzählungszeichen (schwarzer Punkt) einfügen."),
        (None, u'\u2023', "Aufzählungszeichen (Dreieck)", "Ein Aufzählungszeichen (schwarzes Dreieck) einfügen."),
        (None, u'\u25AA', "Aufzählungszeichen (Quadrat)", "Ein Aufzählungszeichen (schwarzes Quadrat) einfügen."),
        (None, u'\u2043', "Aufzählungszeichen (Strich)", "Ein Aufzählungszeichen (Bindestrich) einfügen."),
        (None, u'\u2212', "Echtes Minuszeichen", "Ein echtes Minuszeichen (kein Bindestrich) einfügen."),
        (None, u'\x2b',   "Pluszeichen", "Ein Pluszeichen einfügen."),
        (None, u'\u2610', "Box leer"),
        (None, u'\u2611', "Box Häkchen"),
        (None, u'\u2612', "Box Kreuzchen"),
        ("Wingdings", u'J', "Wingdings Smiley gut"),
        ("Wingdings", u'K', "Wingdings Smiley neutral"),
        ("Wingdings", u'L', "Wingdings Smiley schlecht"),
        (None, u'\u2713', "Häkchen", "Ein Häkchen-Symbol einfügen."),
        (None, u'\u2714', "Häkchen fett", "Ein fettes Häkchen-Symbol einfügen."),
        (None, u'\u2717', "Kreuzchen geschwungen", "Ein geschwungenes Kreuzchen (passend zu Häkchen) einfügen."),
        (None, u'\u2718', "Kreuzchen geschwungen fett", "Ein fettes geschwungenes Kreuzchen (passend zu Häkchen) einfügen."),
        (None, u'\u2715', "Kreuzchen symmetrisch", "Ein symmetrisches Kreuzchen (ähnlich Malzeichen) einfügen."),
        (None, u'\u2716', "Kreuzchen symmetrisch fett", "Ein fettes symmetrisches Kreuzchen (ähnlich Malzeichen) einfügen."),
        (None, u'\u2605', "Stern schwarz"),
        (None, u'\u2606', "Stern weiß"),
        (None, u'\u261B', "Zeigefinger schwarz"),
        (None, u'\u261E', "Zeigefinger weiß"),
        ("Wingdings", u'C', "Wingdings Thumbs-Up"),
        ("Wingdings", u'D', "Wingdings Thumbs-Down"),
        ### Default list symbol:
        # ("Arial",       u'\u2022', "Arial Bullet"),
        ("Courier New", u'o', "Courier New Kreis"),
        ("Wingdings",   u'\xa7', "Wingdings Rechteck"),
        ("Symbol",      u'-', "Symbol Strich"),
        ("Wingdings",   u'v', "Wingdings Stern"),
        ("Wingdings",   u'\xd8', "Wingdings Pfeil"),
        ("Wingdings",   u'\xfc', "Wingdings Häckchen"),
    ]

    ### ARROWS ###
    arrows = [
        (None, u'\u2192', "Pfeil rechts"),
        (None, u'\u2190', "Pfeil links"),
        (None, u'\u2191', "Pfeil oben"),
        (None, u'\u2193', "Pfeil unten"),
        (None, u'\u2194', "Pfeil links und rechts"),
        (None, u'\u21C4', "Pfeil links und rechts"),
        (None, u'\u2197', "Pfeil rechts oben"),
        (None, u'\u2196', "Pfeil links oben"),
        (None, u'\u2198', "Pfeil rechts unten"),
        (None, u'\u2199', "Pfeil links unten"),
        (None, u'\u2195', "Pfeil oben und unten"),
        (None, u'\u21C5', "Pfeil oben und unten"),
        (None, u'\u21E8', "Weißer Pfeil rechts"),
        (None, u'\u21E6', "Weißer Pfeil links"),
        (None, u'\u21E7', "Weißer Pfeil oben"),
        (None, u'\u21E9', "Weißer Pfeil unten"),
        (None, u'\u21AF', "Blitz"),
        (None, u'\u21BA', "Kreispfeil gegen den Uhrzeigersinn"),
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
    def get_text_menu(cls):
        # import fontawesome

        def _unicode_font_button(font):
            return bkt.ribbon.ToggleButton(
                label=font,
                screentip="Unicode-Schriftart "+font,
                supertip=font+" als Unicode-Schriftart verwenden.",
                on_toggle_action=bkt.Callback(lambda pressed: pplib.PPTSymbolsSettings.switch_unicode_font(font)),
                get_pressed=bkt.Callback(lambda: pplib.PPTSymbolsSettings.unicode_font == font),
                get_image=bkt.Callback(lambda:bkt.ribbon.SymbolsGallery.create_symbol_image(font, u"\u2192"))
            )

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
                        get_image=bkt.Callback(lambda:bkt.ribbon.SymbolsGallery.create_symbol_image("Arial", u"\u23B5")) #alt: 2423
                    ),
                    bkt.ribbon.Button(
                        id='symbols_add_protected_narrow_space',
                        label='Schmales geschütztes Leerzeichen',
                        supertip='Ein schmales geschütztes Leerzeichen erlaubt keinen Zeilenumbruch und ist bspw. zwischen Buchstaben von Abkürzungen zu verwenden.',
                        on_action=bkt.Callback(cls.add_protected_narrow_space, selection=True),
                        get_enabled = bkt.Callback(cls.text_selection, selection=True),
                        get_image=bkt.Callback(lambda:bkt.ribbon.SymbolsGallery.create_symbol_image("Arial", u"\u02FD"))
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
                    bkt.ribbon.Menu(
                        label="Unicode-Schriftart wählen",
                        image_mso='FontDialogPowerPoint',
                        supertip="Unicode-Zeichen können entweder mit der Standard-Schriftart oder einer speziellen Unicode-Schriftart eingefügt werden. Diese kann hier ausgewählt werden.",
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
    
    
    
    
    
inner_margin_top    = InnerMargin(attr="MarginTop",    id='textFrameMargin-top-2',    image_button=True, show_label=False, image_mso='FillDown' , label="Innenabstand oben",   screentip='Innenabstand oben',   supertip='Ändere den oberen Innenabstand des Textfelds auf das angegebene Maß (in cm).')
inner_margin_bottom = InnerMargin(attr="MarginBottom", id='textFrameMargin-bottom-2', image_button=True, show_label=False, image_mso='FillUp'   , label="Innenabstand unten",  screentip='Innenabstand unten',  supertip='Ändere den unteren Innenabstand des Textfelds auf das angegebene Maß (in cm).')
inner_margin_left   = InnerMargin(attr="MarginLeft",   id='textFrameMargin-left-2',   image_button=True, show_label=False, image_mso='FillRight', label="Innenabstand links",  screentip='Innenabstand links',  supertip='Ändere den linken Innenabstand des Textfelds auf das angegebene Maß (in cm).')
inner_margin_right  = InnerMargin(attr="MarginRight",  id='textFrameMargin-right-2',  image_button=True, show_label=False, image_mso='FillLeft' , label="Innenabstand rechts", screentip='Innenabstand rechts', supertip='Ändere den rechten Innenabstand des Textfelds auf das angegebene Maß (in cm).')



innenabstand_gruppe = bkt.ribbon.Group(
    id="bkt_innermargin_group",
    label="Textfeld Innenabstand",
    image_mso='ObjectNudgeRight',
    children=[
    bkt.ribbon.Box(id='textFrameMargin-box-top', children=[
        bkt.ribbon.LabelControl(id='textFrameMargin-label-top', label=u'         \u200b'),
        #create_margin_spinner('MarginTop', imageMso='ObjectNudgeDown'),
        inner_margin_top,
        bkt.ribbon.LabelControl(label=u'   \u200b'),
        bkt.ribbon.Button(
            id='textFrameMargin-zero',
            label=u"=\u202F0",
            screentip="Innenabstand auf Null",
            supertip="Ändere in Innenabstand des Textfelds an allen Seiten auf Null.",
            on_action=bkt.Callback( InnerMargin.set_to_0 )
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
        bkt.ribbon.LabelControl(id='textFrameMargin-label-bottom', label=u'         \u200b'),
        #create_margin_spinner('MarginBottom', imageMso='ObjectNudgeUp'),
        inner_margin_bottom,
        bkt.ribbon.LabelControl(label=u'   \u200b'),
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
        label=u"Absatzabstand oben",
        image_mso='WordOpenParaAbove',
        screentip="Oberen Absatzabstand",
        supertip="Ändere den Absatzabstand vor dem Absatz auf das angegebene Maß (entweder in Abstand Zeilen oder in pt).",
    )

class ParSpaceAfter(pplib.ParagraphFormatSpinnerBox):
    attr = 'SpaceAfter'
    _attributes = dict(
        label=u"Absatzabstand unten",
        image_mso='WordOpenParaBelow',
        screentip="Unteren Absatzabstand",
        supertip="Ändere den Absatzabstand hinter dem Absatz auf das angegebene Maß (entweder in Abstand Zeilen oder in pt).",
    )

class ParSpaceWithin(pplib.ParagraphFormatSpinnerBox):
    attr = 'SpaceWithin'
    _attributes = dict(
        label=u"Zeilenabstand",
        image_mso='LineSpacing',
        screentip="Zeilenabstand",
        supertip="Ändere den Zeilenabstand (entweder in Abstand Zeilen oder in pt).",
        fallback_value = 1,
    )

class ParFirstLineIndent(pplib.ParagraphFormatSpinnerBox):
    attr = 'FirstLineIndent'
    _attributes = dict(
        label=u"Einzug 1. Zeile",
        image='first_line_indent',
        screentip="Sondereinzug",
        supertip="Ändere den Sondereinzug (1. Zeile, hängend) auf das angegebene Maß (in cm).",
    )

class ParLeftIndent(pplib.ParagraphFormatSpinnerBox):
    attr = 'LeftIndent'
    _attributes = dict(
        label=u"Einzug links",
        image_mso='ParagraphIndentLeft',
        screentip="Absatzeinzug links",
        supertip="Ändere den linken Absatzeinzug auf das angegebene Maß (in cm).",
    )

class ParRightIndent(pplib.ParagraphFormatSpinnerBox):
    attr = 'RightIndent'
    _attributes = dict(
        label=u"Einzug rechts",
        image_mso='ParagraphIndentRight',
        screentip="Absatzeinzug rechts",
        supertip="Ändere den rechten Absatzeinzug auf das angegebene Maß (in cm).",
    )


class Absatz(object):

    @staticmethod
    def set_word_wrap(shapes, pressed):
        for shape in shapes:
            shape.TextFrame.WordWrap = -1 if pressed else 0

    @staticmethod
    def get_word_wrap(shapes):
        if not shapes[0].TextFrame:
            return None
        return (shapes[0].TextFrame.WordWrap == -1) # msoTrue


    @staticmethod
    def set_auto_size(shapes, pressed):
        for shape in shapes:
            shape.TextFrame.AutoSize = 1 if pressed else 0
            # 1 = ppAutoSizeShapeToFitText
            # 0 = ppAutoSizeNone

    @staticmethod
    def get_auto_size(shapes):
        if not shapes[0].TextFrame:
            return None
        return (shapes[0].TextFrame.AutoSize == 1)

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



class BulletStyle(object):
    
    @staticmethod
    def set_bullet_color_rgb(selection, shapes, color):
        if selection.Type == 3:
            # text selected
            selection.TextRange2.ParagraphFormat.Bullet.Font.Fill.ForeColor.RGB = color

        else:
            for textframe in pplib.iterate_shape_textframes(shapes):
                textframe.TextRange.ParagraphFormat.Bullet.Font.Fill.ForeColor.RGB = color

        # for shape in shapes:
        #     shape.TextFrame2.TextRange.ParagraphFormat.Bullet.Font.Fill.ForeColor.RGB = color

    @staticmethod
    def set_bullet_theme_color(selection, shapes, color_index, brightness):
        if selection.Type == 3:
            # text selected
            selection.TextRange2.ParagraphFormat.Bullet.Font.Fill.ForeColor.ObjectThemeColor = color_index
            selection.TextRange2.ParagraphFormat.Bullet.Font.Fill.ForeColor.Brightness = brightness

        else:
            for textframe in pplib.iterate_shape_textframes(shapes):
                textframe.TextRange.ParagraphFormat.Bullet.Font.Fill.ForeColor.ObjectThemeColor = color_index
                textframe.TextRange.ParagraphFormat.Bullet.Font.Fill.ForeColor.Brightness = brightness

        # for shape in shapes:
        #     shape.TextFrame2.TextRange.ParagraphFormat.Bullet.Font.Fill.ForeColor.ObjectThemeColor = color_index
        #     shape.TextFrame2.TextRange.ParagraphFormat.Bullet.Font.Fill.ForeColor.Brightness = brightness

    @staticmethod
    def set_bullet_color_auto(selection, shapes):
        if selection.Type == 3:
            # text selected
            selection.TextRange2.ParagraphFormat.Bullet.UseTextColor = -1

        else:
            for textframe in pplib.iterate_shape_textframes(shapes):
                textframe.TextRange.ParagraphFormat.Bullet.UseTextColor = -1

    @staticmethod
    def set_bullet_symbol(selection, shapes, symbol):
        def _set_bullet(par):
            par.Bullet.Character = ord(symbol[1])
            if symbol[0]:
                par.Bullet.Font.Name = symbol[0]
            else:
                par.Bullet.UseTextFont = -1


        if selection.Type == 3:
            # text selected
            _set_bullet(selection.TextRange2.ParagraphFormat)

        else:
            for textframe in pplib.iterate_shape_textframes(shapes):
                _set_bullet(textframe.TextRange.ParagraphFormat)




    @classmethod
    def get_bullet_color_rgb(cls, selection, shapes):
        return cls._get_from_par_format(selection, shapes, cls._get_bullet_color_from_par)

    @classmethod
    def get_bullet_symbol(cls, selection, shapes):
        return cls._get_from_par_format(selection, shapes, cls._get_bullet_symbol_from_par)


    @classmethod
    def _get_from_par_format(cls, selection, shapes, getter_method):
        if selection.Type == 3:
            # text selected
            try:
                # produces error if no text is selected
                return getter_method(selection.TextRange2.Paragraphs(1,1).ParagraphFormat)
            except ValueError: #ValueError: Der Index in der angegebenen Sammlung ist außerhalb des zulässigen Bereichs.
                return getter_method(selection.TextRange2.ParagraphFormat)
        
        else:
            # shapes selected
            for textframe in pplib.iterate_shape_textframes(shapes):
                return getter_method(textframe.TextRange.ParagraphFormat)


    @classmethod
    def _get_bullet_color_from_par(cls, par_format):
        return [par_format.Bullet.Font.Fill.ForeColor.ObjectThemeColor, par_format.Bullet.Font.Fill.ForeColor.Brightness, par_format.Bullet.Font.Fill.ForeColor.RGB]

    @classmethod
    def _get_bullet_symbol_from_par(cls, par_format):
        if par_format.Bullet.Visible:
            return unichr(par_format.Bullet.Character)
        return None

    
    @classmethod
    def fix_bullet_style(cls, shapes):
        shape = shapes[0]
        slide = shape.Parent
        placeholders = [shape for shape in slide.Master.Shapes if shape.Type == 14 and shape.PlaceholderFormat.Type == 2]
        ref_shape = placeholders[0]
        cls.fix_bullet_style_by_reference(shapes, ref_shape)
    
    
    @staticmethod
    def fix_bullet_style_by_reference(shapes, ref_shape):
        # shape = shapes[0]
        # slide = shape.Parent
        # placeholders = [shape for shape in slide.Master.Shapes if shape.Type == 14 and shape.PlaceholderFormat.Type == 2]
        # textph = placeholders[0]
        ph_paragraphs = [p for p in ref_shape.TextFrame2.TextRange.Paragraphs() ]
        ph_paragraphs = [[p for p in ph_paragraphs if p.ParagraphFormat.IndentLevel == indent_level] for indent_level in range(1,6) ] #IndentLevel is between 1 and 5
        ph_paragraphs = [ None if len(ph_list) == 0 else ph_list[0] for ph_list in ph_paragraphs ]
        
        ph_formats = [None if p==None else p.ParagraphFormat for p in ph_paragraphs ]
        
        # for shape in shapes:
            # for par in shape.TextFrame2.TextRange.Paragraphs() :
        for textframe in pplib.iterate_shape_textframes(shapes):
            for par in textframe.TextRange.Paragraphs() :
                indent_level = par.ParagraphFormat.IndentLevel
            
                if ph_paragraphs[indent_level] is None:
                    continue

                par.ParagraphFormat.Bullet.Style = ph_formats[indent_level].Bullet.Style
                par.ParagraphFormat.Bullet.Type = ph_formats[indent_level].Bullet.Type
                par.ParagraphFormat.Bullet.RelativeSize = ph_formats[indent_level].Bullet.RelativeSize
                par.ParagraphFormat.Bullet.Character = ph_formats[indent_level].Bullet.Character

                if ph_formats[indent_level].Bullet.UseTextFont == -1:
                    par.ParagraphFormat.Bullet.UseTextFont = -1
                else:
                    par.ParagraphFormat.Bullet.Font.Name = ph_formats[indent_level].Bullet.Font.Name
                
                if ph_formats[indent_level].Bullet.UseTextColor == -1:
                    par.ParagraphFormat.Bullet.UseTextColor = -1
                else:
                    if ph_formats[indent_level].Bullet.Font.Fill.ForeColor.ObjectThemeColor == 0:
                        par.ParagraphFormat.Bullet.Font.Fill.ForeColor.RGB = ph_formats[indent_level].Bullet.Font.Fill.ForeColor.RGB
                    else:
                        par.ParagraphFormat.Bullet.Font.Fill.ForeColor.ObjectThemeColor = ph_formats[indent_level].Bullet.Font.Fill.ForeColor.ObjectThemeColor
                        par.ParagraphFormat.Bullet.Font.Fill.ForeColor.Brightness = ph_formats[indent_level].Bullet.Font.Fill.ForeColor.Brightness
            
                # par.ParagraphFormat.Bullet.UseTextColor = ph_formats[indent_level].Bullet.UseTextColor
                # par.ParagraphFormat.Bullet.UseTextFont = ph_formats[indent_level].Bullet.UseTextFont

                par.ParagraphFormat.Bullet.Visible = ph_formats[indent_level].Bullet.Visible
            
                par.ParagraphFormat.LeftIndent = ph_formats[indent_level].LeftIndent
                par.ParagraphFormat.RightIndent = ph_formats[indent_level].RightIndent
                par.ParagraphFormat.FirstLineIndent = ph_formats[indent_level].FirstLineIndent
                par.ParagraphFormat.HangingPunctuation = ph_formats[indent_level].HangingPunctuation
            
                par.ParagraphFormat.BaselineAlignment = ph_formats[indent_level].BaselineAlignment
                par.ParagraphFormat.LineRuleBefore    = ph_formats[indent_level].LineRuleBefore
                par.ParagraphFormat.SpaceBefore       = ph_formats[indent_level].SpaceBefore
                par.ParagraphFormat.LineRuleAfter     = ph_formats[indent_level].LineRuleAfter
                par.ParagraphFormat.SpaceAfter        = ph_formats[indent_level].SpaceAfter
                par.ParagraphFormat.LineRuleWithin    = ph_formats[indent_level].LineRuleWithin
                par.ParagraphFormat.SpaceWithin       = ph_formats[indent_level].SpaceWithin
            
                # Text format
                par.Font.Name   = ph_paragraphs[indent_level].Font.Name
                par.Font.Size   = ph_paragraphs[indent_level].Font.Size
                par.Font.Bold   = ph_paragraphs[indent_level].Font.Bold
                par.Font.Italic = ph_paragraphs[indent_level].Font.Italic
                par.Font.Caps   = ph_paragraphs[indent_level].Font.Caps
                if ph_paragraphs[indent_level].Font.Fill.ForeColor.ObjectThemeColor == 0:
                    par.Font.Fill.ForeColor.RGB = ph_paragraphs[indent_level].Font.Fill.ForeColor.RGB
                else:
                    par.Font.Fill.ForeColor.ObjectThemeColor = ph_paragraphs[indent_level].Font.Fill.ForeColor.ObjectThemeColor
                    par.Font.Fill.ForeColor.Brightness = ph_paragraphs[indent_level].Font.Fill.ForeColor.Brightness
    


class TextShapes(object):
    sticker_alignment = bkt.settings.get("toolbox.sticker_alignment", "right")
    sticker_fontsize = bkt.settings.get("toolbox.sticker_fontsize", 14)
    sticker_custom = bkt.settings.get("toolbox.sticker_custom", None)

    @classmethod
    def settings_setter(cls, name, value):
        setattr(cls, name, value)
        bkt.settings["toolbox."+name] = value

    @staticmethod
    def textbox_insert(context, pressed):
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
    
    @staticmethod
    def addUnderlinedTextbox(slide, presentation):
        # Textbox erstellen, damit Standardformatierung der Textbox genommen wird
        shp = slide.shapes.AddTextbox( 1 #msoTextOrientationHorizontal
            , 100, 100, 200, 50)
        # Shape-Typ ist links-rechts-Pfeil, weil es die passenden Connector-Ecken hat
        shp.AutoShapeType = pplib.MsoAutoShapeType['msoShapeLeftRightArrow']
        # Shape-Anpassung, so dass es wie ein Rechteck aussieht
        shp.Adjustments.item[1] = 1
        shp.Adjustments.item[2] = 0
        # Text
        shp.TextFrame.TextRange.text = "Lorem ipsum"

        # Mitting ausrichten
        shp.Top = (presentation.PageSetup.SlideHeight - shp.height) /2
        shp.Left = (presentation.PageSetup.SlideWidth - shp.width) /2

        # Connectoren erstellen und mit Connector-Ecken des Shapes verbinden
        connector = slide.shapes.AddConnector(Type=1 #msoConnectorStraight
            , BeginX=0, BeginY=0, EndX=100, EndY=100)
        connector.ConnectorFormat.BeginConnect(ConnectedShape=shp, ConnectionSite=5)
        connector.ConnectorFormat.EndConnect(ConnectedShape=shp, ConnectionSite=7)
        
        # Default Formatierung
        color = shp.TextFrame.TextRange.Font.Color
        if color.Type == pplib.MsoColorType["msoColorTypeScheme"]:
            connector.Line.ForeColor.ObjectThemeColor = color.ObjectThemeColor
            connector.Line.ForeColor.Brightness = color.Brightness
        else:
            connector.Line.ForeColor.RGB = color.RGB
        connector.Line.Weight = 1.5
        # shp.TextFrame.MarginBottom = 0
        # shp.TextFrame.MarginTop    = 0
        # shp.TextFrame.MarginLeft   = 0
        # shp.TextFrame.MarginRight  = 0

        # Text auswählen
        shp.TextFrame2.TextRange.Select()
    
    
    @classmethod
    def addSticker(cls, slide, presentation, sticker_text="DRAFT", select_text=True):
        # Textbox erstellen, damit Standardformatierung der Textbox genommen wird
        shp = slide.shapes.AddTextbox( 1 #msoTextOrientationHorizontal
            , 0, 60, 100, 20)
        # Shape-Typ ist links-rechts-Pfeil, weil es die passenden Connector-Ecken hat
        shp.AutoShapeType = pplib.MsoAutoShapeType['msoShapeLeftRightArrow']
        # Shape-Anpassung, so dass es wie ein Rechteck aussieht
        shp.Adjustments.item[1] = 1
        shp.Adjustments.item[2] = 0
        # Shape-Stil
        # shp.Line.Weight = 0.75
        shp.Fill.Visible = 0 #msoFalse
        shp.Line.Visible = 0 #msoFalse
        # Text-Stil
        # shp.TextFrame.TextRange.Font.Color.RGB = 0
        shp.TextFrame.TextRange.Font.Size = cls.sticker_fontsize
        shp.TextFrame.TextRange.ParagraphFormat.Bullet.Visible = False
        # Autosize / Text nicht umbrechen
        shp.TextFrame.WordWrap = 0 #msoFalse
        shp.TextFrame.AutoSize = 1 #ppAutoSizeShapeToFitText
        # Innenabstand
        shp.TextFrame.MarginBottom = 0
        shp.TextFrame.MarginTop    = 0
        shp.TextFrame.MarginLeft   = 0
        shp.TextFrame.MarginRight  = 0
        # Text
        shp.TextFrame.TextRange.text = sticker_text
        # Top-Position
        shp.Top = 15

        # Text alignment + Left-Position
        if cls.sticker_alignment == "left":
            shp.TextFrame.TextRange.ParagraphFormat.Alignment = 1 #ppAlignLeft
            shp.Left = 15
        elif cls.sticker_alignment == "center":
            shp.TextFrame.TextRange.ParagraphFormat.Alignment = 2 #ppAlignCenter
            shp.Left = presentation.PageSetup.SlideWidth/2 - shp.width/2
        else: #right
            shp.TextFrame.TextRange.ParagraphFormat.Alignment = 3 #ppAlignRight
            shp.Left = presentation.PageSetup.SlideWidth - shp.width - 15

        # Connectoren erstellen und mit Connector-Ecken des Shapes verbinden
        connector1 = slide.shapes.AddConnector(Type=1 #msoConnectorStraight
            , BeginX=0, BeginY=0, EndX=100, EndY=100)
        connector1.ConnectorFormat.BeginConnect(ConnectedShape=shp, ConnectionSite=1)
        connector1.ConnectorFormat.EndConnect(ConnectedShape=shp, ConnectionSite=3)
        connector1.Line.ForeColor.RGB = 0
        connector1.Line.Weight = 0.75

        connector2 = slide.shapes.AddConnector(Type=1 #msoConnectorStraight
            , BeginX=0, BeginY=0, EndX=100, EndY=100)
        connector2.ConnectorFormat.BeginConnect(ConnectedShape=shp, ConnectionSite=5)
        connector2.ConnectorFormat.EndConnect(ConnectedShape=shp, ConnectionSite=7)
        connector2.Line.ForeColor.RGB = 0
        connector2.Line.Weight = 0.75

        color = shp.TextFrame.TextRange.Font.Color
        if color.Type == pplib.MsoColorType["msoColorTypeScheme"]:
            connector1.Line.ForeColor.ObjectThemeColor = color.ObjectThemeColor
            connector1.Line.ForeColor.Brightness = color.Brightness
            connector2.Line.ForeColor.ObjectThemeColor = color.ObjectThemeColor
            connector2.Line.ForeColor.Brightness = color.Brightness
        else:
            connector1.Line.ForeColor.RGB = color.RGB
            connector2.Line.ForeColor.RGB = color.RGB

        if select_text:
            # Text auswählen
            shp.Select()
            shp.TextFrame2.TextRange.Select()
    
    
    @classmethod
    def add_sticker_to_slides(cls, slides, presentation, sticker_text="DRAFT"):
        select_text = len(slides) == 1
        for slide in slides:
            cls.addSticker(slide, presentation, sticker_text, select_text)
    
    @classmethod
    def own_sticker_enabled(cls):
        return cls.sticker_custom is not None
    
    @classmethod
    def own_sticker_label(cls):
        if cls.sticker_custom:
            return cls.sticker_custom + "-Sticker"
        else:
            return "Noch nicht definiert"
    
    @classmethod
    def own_sticker_insert(cls, slides, presentation):
        if cls.sticker_custom and not bkt.get_key_state(bkt.KeyCodes.SHIFT):
            cls.add_sticker_to_slides(slides, presentation, cls.sticker_custom)
        else:
            cls.own_sticker_edit()
    
    @classmethod
    def own_sticker_edit(cls):
        res = bkt.ui.show_user_input("Selbst definierten Sticker-Text eingeben:", "Sticker bearbeiten", cls.sticker_custom)
        cls.sticker_custom = bkt.settings["toolbox.sticker_custom"] = res


class TextOnShape(object):

    @classmethod
    def find_shape_on_shape(cls, master_shape, shapes):
        if master_shape.HasTextFrame == 0:
            return None
        for s in shapes:
            #shape on top of master shape and shape midpoint within master shape
            if s != master_shape and s.HasTextFrame == -1 and s.ZOrderPosition > master_shape.ZOrderPosition and s.left+s.width/2 >= master_shape.left and s.left+s.width/2 <= master_shape.left+master_shape.width and s.top+s.height/2 >= master_shape.top and s.top+s.height/2 <= master_shape.top+master_shape.height:
                return s
        return None
    
    @classmethod
    def merge_shapes(cls, master_shape, text_shape):
        # Kein TextFrame, bspw. bei Linien
        if master_shape.HasTextFrame == 0 or text_shape.HasTextFrame == 0:
            return

        # Text kopieren
        # text_shape.TextFrame2.TextRange.Copy()
        # master_shape.TextFrame2.TextRange.Paste()
        pplib.transfer_textrange(text_shape.TextFrame2.TextRange, master_shape.TextFrame2.TextRange)
        # Textbox loeschen
        text_shape.Delete()

    @classmethod
    def textIntoShape(cls, shapes):
        if len(shapes) == 2 and shapes[0].Type == pplib.MsoShapeType['msoTextBox']:
            cls.merge_shapes(shapes[1], shapes[0])
        elif len(shapes) == 2 and shapes[1].Type == pplib.MsoShapeType['msoTextBox']:
            cls.merge_shapes(shapes[0], shapes[1])
        else:
            shapes = sorted(shapes, key=lambda s: s.ZOrderPosition) #important due to removal of items in for loops
            for shape in shapes:
                inner_shp = cls.find_shape_on_shape(shape, shapes)
                if inner_shp is not None:
                    cls.merge_shapes(shape, inner_shp)
                    shapes.remove(inner_shp)


    @staticmethod
    def textOutOfShape(shapes, context):
        for shp in shapes:
            #if shp.TextFrame.TextRange.text != "":
            if shp.HasTextFrame == -1 and shp.TextFrame.HasText == -1:
                shpTxt = context.slide.shapes.AddTextbox(
                    1, #msoTextOrientationHorizontal
                    shp.Left, shp.Top, shp.Width, shp.Height)
                # WordWrap / AutoSize
                # shpTxt.TextFrame2.WordWrap = -1 #msoTrue
                shpTxt.TextFrame2.WordWrap = shp.TextFrame2.WordWrap
                shpTxt.TextFrame2.AutoSize = 0 #ppAutoSizeNone
                shpTxt.Height   = shp.Height
                shpTxt.Rotation = shp.Rotation
                # Seitenraender
                shpTxt.TextFrame2.MarginBottom = shp.TextFrame2.MarginBottom
                shpTxt.TextFrame2.MarginTop    = shp.TextFrame2.MarginTop
                shpTxt.TextFrame2.MarginLeft   = shp.TextFrame2.MarginLeft
                shpTxt.TextFrame2.MarginRight  = shp.TextFrame2.MarginRight
                # Ausrichtung
                shpTxt.TextFrame2.Orientation      = shp.TextFrame2.Orientation
                shpTxt.TextFrame2.HorizontalAnchor = shp.TextFrame2.HorizontalAnchor
                shpTxt.TextFrame2.VerticalAnchor   = shp.TextFrame2.VerticalAnchor
                # Text kopieren
                shp.TextFrame2.TextRange.Copy()
                shpTxt.TextFrame2.TextRange.Paste()
                shp.TextFrame2.DeleteText()
                # Größe wiederherstellen
                shp.Height = shpTxt.Height
                shp.Width = shpTxt.Width
                # Textfeld selektieren
                shpTxt.Select(0)
    
    ### context menu callbacks ###

    @staticmethod
    def is_outable(shape):
        return shape.HasTextFrame == -1 and shape.TextFrame.HasText == -1 and shape.Type not in [pplib.MsoShapeType['msoTextBox'], pplib.MsoShapeType['msoPlaceholder']]
    
    @staticmethod
    def is_mergable(shapes):
        return len(shapes) == 2 and (shapes[0].Type == pplib.MsoShapeType['msoTextBox'] or shapes[1].Type == pplib.MsoShapeType['msoTextBox'])
    

class SplitTextShapes(object):


    ### This method was required in Office 2007 were BoundXXX methods were not available
    # @classmethod
    # def paragraph_height(cls, par, with_par_spaces=True):
    #     parHeight = par.Lines().Count * cls.line_height(par) * 1.0
    #     if with_par_spaces:
    #         parHeight = parHeight + max(0, par.ParagraphFormat.SpaceBefore) + max(0, par.ParagraphFormat.SpaceAfter)
    #     return parHeight
    

    ### This method was required in Office 2007 were BoundXXX methods were not available
    # @staticmethod
    # def line_height(par):
    #     if par.ParagraphFormat.LineRuleWithin == -1:
    #         # spacing = number of lines
    #         # Annahme zur Korrektur der Abstände: Abstand zwischen zwei Zeilen ist 0.2pt
    #         return par.Font.Size * (max(0, par.ParagraphFormat.SpaceWithin) + 0.2)
    #     else:
    #         # spacing = number of pt
    #         # Annahme zur Korrektur der Abstände: Abstand zwischen zwei Zeilen ist 0.2pt
    #         return par.ParagraphFormat.SpaceWithin #+ 0.1 * .Font.Size
    
    
    @staticmethod
    def trim_newline_character(par):
        if par.Characters(par.Length, 1).Text == "\r":
            par.Characters(par.Length, 1).Delete()
    
    
    @classmethod
    def splitShapesByParagraphs(cls, shapes, context):
        for shp in shapes:
            # if shp.TextFrame2.TextRange.Text != "":
            if shp.HasTextFrame == -1 and shp.TextFrame.HasText == -1 and shp.TextFrame.TextRange.Paragraphs().Count > 1:
                #Shape exklusiv markieren (alle anderen deselektieren)
                # shp.Select(-1) # msoTrue

                par_count = shp.TextFrame2.TextRange.Paragraphs().Count
                for par_index in range(2, par_count+1):
                    par = shp.TextFrame2.TextRange.Paragraphs(par_index)
                    # Leere Paragraphen überspringen
                    if par.text in ["", "\r"]:
                        continue
                    # Shape dublizieren
                    shpCopy = shp.Duplicate()
                    shpCopy.Select(0) # msoFalse
                    shpCopy.Top  = par.BoundTop - shp.TextFrame2.MarginTop + par.ParagraphFormat.SpaceBefore
                    shpCopy.Left = shp.Left
                    
                    # Absaetze 1..i-1 entfernen und Shape entsprechend verschieben
                    shpCopy.TextFrame2.TextRange.Paragraphs(1, par_index-1).Delete()
                    # for index in range(1, par_index):
                    #     # Textbox Position entsprechend Absatzhoehe anpassen
                    #     # shpCopy.Top = shpCopy.Top + cls.paragraph_height(shpCopy.TextFrame2.TextRange.Paragraphs(1))
                    #     # Absatz entfernen
                    #     shpCopy.TextFrame2.TextRange.Paragraphs(1).Delete()
                    
                    # Absaetze i+1..n entfernen
                    shpCopy.TextFrame2.TextRange.Paragraphs(2, shpCopy.TextFrame2.TextRange.Paragraphs().Count-1).Delete()
                    # for index in range(par_index + 1, shp.TextFrame2.TextRange.Paragraphs().Count + 1):
                    #     shpCopy.TextFrame2.TextRange.Paragraphs(2).Delete()

                    # Letztes CR-Zeichen loeschen
                    cls.trim_newline_character(shpCopy.TextFrame2.TextRange)

                    # Shape Hoehe abhaengig von Absaetzhoehe
                    # shpCopy.Height = cls.paragraph_height(shpCopy.TextFrame.TextRange.Paragraphs(1)) + shpCopy.TextFrame2.MarginTop + shpCopy.TextFrame2.MarginBottom
                    shpCopy.Height = par.BoundHeight + shp.TextFrame2.MarginTop + shp.TextFrame2.MarginBottom - par.ParagraphFormat.SpaceBefore - par.ParagraphFormat.SpaceAfter
                    if par_index == par_count:
                        #last paragraph does not have spaceafter
                        shpCopy.Height += par.ParagraphFormat.SpaceAfter

                    # --> ein Absatz bleibt übrig

                # letzten Shape nach unten schieben
                # shpCopy.Top = max(shpCopy.Top, shp.Top + shp.Height - shpCopy.Height)

                # Absaetze 2..n im Original-Shape entfernen
                shp.TextFrame2.TextRange.Paragraphs(2, par_count-1).Delete()

                # Letztes CR-Zeichen loeschen
                cls.trim_newline_character(shp.TextFrame2.TextRange)
                # Textbox Hoehe an Absatzhoehe anpassen
                # shp.Height = cls.paragraph_height(shp.TextFrame2.TextRange.Paragraphs(1)) + shp.TextFrame2.MarginTop + shp.TextFrame2.MarginBottom
                shp.Height = shp.TextFrame2.TextRange.Paragraphs(1).BoundHeight + shp.TextFrame2.MarginTop + shp.TextFrame2.MarginBottom
                
                #Verteilung bei 2 Shapes führt zu Fehler
                # if context.app.ActiveWindow.Selection.ShapeRange.Count > 2:
                #     # Objekte vertikal verteilen
                #     context.app.ActiveWindow.Selection.ShapeRange.Distribute(
                #         1, #msoDistributeVertically
                #         0) #msoFalse)

    @classmethod
    def joinShapesWithText(cls, shapes):
        # Shapes nach top sortieren
        shapes = sorted(shapes, key=lambda shape: shape.Top)
        # Anapssung Größe des ersten Shapes (Master-Shape) mit TextFrame
        for i in range(len(shapes)):
            shpMaster = shapes.pop(i) #shapes[0]
            if shpMaster.HasTextFrame:
                break
        else:
            # no shape with textframe found
            return
        
        shpMaster.Height = max(shpMaster.Height, shapes[-1].Top + shapes[-1].Height - shpMaster.Top)

        for shp in shapes: #[1:]:
            # Text aus Shape kopieren
            shp.TextFrame2.TextRange.Copy()
            # neuen Absatz in Master-Shape erstellen
            # Bug in PowerPoint: machmal muss InsertAfter mehrmals aufgerufen werden
            parCount = shpMaster.TextFrame2.TextRange.Paragraphs().Count
            for i in range(10):
                txtRange = shpMaster.TextFrame2.TextRange.Paragraphs().InsertAfter("\r")
                if parCount < shpMaster.TextFrame2.TextRange.Paragraphs().Count:
                    break
            # Text in Master-Shape einfuegen
            txtRange.Paste()
            # Letztes CR-Zeichen loeschen
            cls.trim_newline_character(txtRange)
            # Shape loeschen
            shp.Delete()
    
    ### context menu callbacks ###

    @staticmethod
    def is_splitable(shape):
        return shape.HasTextFrame == -1 and shape.TextFrame2.TextRange.Paragraphs().Count>1
    
    @staticmethod
    def is_joinable(shapes):
        return any(shp.HasTextFrame == -1 and shp.TextFrame2.HasText == -1 for shp in shapes)
    


text_menu = bkt.ribbon.Menu(
    label="Textfeld zeichnen Menü",
    supertip="Sticker einfügen, Bullet Points angleichen, sowie weitere Text-bezogene Funktionen",
    children=[
        bkt.ribbon.MenuSeparator(title="Textformen einfügen"),
        bkt.mso.control.TextBoxInsert,
        bkt.ribbon.SplitButton(
            id="sticker_splitbutton",
            children=[
                bkt.ribbon.Button(
                    id = 'sticker',
                    label = u"Sticker",
                    image = "Sticker",
                    screentip="Sticker einfügen",
                    supertip="Füge ein Sticker-Shape oben rechts auf dem aktuellen Slide ein.",
                    on_action=bkt.Callback(TextShapes.add_sticker_to_slides, slides=True, presentation=True)
                ),
                bkt.ribbon.Menu(label="Sticker Menü", supertip="Verschiedene Sticker einfügen", children=[
                    bkt.ribbon.Button(
                        id="sticker_draft",
                        label = u"DRAFT-Sticker",
                        image = "Sticker",
                        screentip="DRAFT-Sticker einfügen",
                        supertip="Füge ein Sticker-Shape oben rechts auf dem aktuellen Slide mit Text DRAFT ein.",
                        on_action=bkt.Callback(TextShapes.add_sticker_to_slides, slides=True, presentation=True)
                    ),
                    bkt.ribbon.Button(
                        id="sticker_backup",
                        label = u"BACKUP-Sticker",
                        screentip="BACKUP-Sticker einfügen",
                        supertip="Füge ein Sticker-Shape oben rechts auf dem aktuellen Slide mit Text BACKUP ein.",
                        on_action=bkt.Callback(lambda slides, presentation: TextShapes.add_sticker_to_slides(slides, presentation, "BACKUP"), slides=True, presentation=True)
                    ),
                    bkt.ribbon.Button(
                        id="sticker_discussion",
                        label = u"FOR DISCUSSION-Sticker",
                        screentip="FOR DISCUSSION-Sticker einfügen",
                        supertip="Füge ein Sticker-Shape oben rechts auf dem aktuellen Slide mit Text FOR DISCUSSION ein.",
                        on_action=bkt.Callback(lambda slides, presentation: TextShapes.add_sticker_to_slides(slides, presentation, "FOR DISCUSSION"), slides=True, presentation=True)
                    ),
                    bkt.ribbon.Button(
                        id="sticker_illustrative",
                        label = u"ILLUSTRATIVE-Sticker",
                        screentip="ILLUSTRATIVE-Sticker einfügen",
                        supertip="Füge ein Sticker-Shape oben rechts auf dem aktuellen Slide mit Text ILLUSTRATIVE ein.",
                        on_action=bkt.Callback(lambda slides, presentation: TextShapes.add_sticker_to_slides(slides, presentation, "ILLUSTRATIVE"), slides=True, presentation=True)
                    ),
                    bkt.ribbon.Button(
                        id="sticker_confidential",
                        label = u"CONFIDENTIAL-Sticker",
                        screentip="CONFIDENTIAL-Sticker einfügen",
                        supertip="Füge ein Sticker-Shape oben rechts auf dem aktuellen Slide mit Text CONFIDENTIAL ein.",
                        on_action=bkt.Callback(lambda slides, presentation: TextShapes.add_sticker_to_slides(slides, presentation, "CONFIDENTIAL"), slides=True, presentation=True)
                    ),
                    bkt.ribbon.MenuSeparator(),
                    bkt.ribbon.Button(
                        id="sticker_own",
                        get_label=bkt.Callback(TextShapes.own_sticker_label),
                        screentip="Selbst definierten Sticker einfügen",
                        supertip="Füge ein Sticker-Shape oben rechts auf dem aktuellen Slide mit selbst definiertem ein.",
                        on_action=bkt.Callback(TextShapes.own_sticker_insert, slides=True, presentation=True),
                        get_enabled=bkt.Callback(TextShapes.own_sticker_enabled)
                    ),
                    bkt.ribbon.Button(
                        id="sticker_own_edit",
                        label = u"Sticker-Text ändern",
                        screentip="Selbst definierten Sticker bearbeiten",
                        supertip="Ändere des Text des selbst definierten Stickers.",
                        on_action=bkt.Callback(TextShapes.own_sticker_edit)
                    ),
                    bkt.ribbon.MenuSeparator(),
                    bkt.ribbon.Menu(
                        label="Ausrichtung",
                        supertip="Ausrichtungsoptionen für Sticker einstellen",
                        children=[
                            bkt.ribbon.ToggleButton(
                                label="Links",
                                screentip="Sticker-Ausrichtung Links",
                                supertip="Setzt die Position beim Einfügen der Sticker auf links.",
                                get_pressed=bkt.Callback(lambda: TextShapes.sticker_alignment == "left"),
                                on_toggle_action=bkt.Callback(lambda pressed: TextShapes.settings_setter("sticker_alignment", "left")),
                            ),
                            bkt.ribbon.ToggleButton(
                                label="Mitte",
                                screentip="Sticker-Ausrichtung Mitte",
                                supertip="Setzt die Position beim Einfügen der Sticker auf mittig.",
                                get_pressed=bkt.Callback(lambda: TextShapes.sticker_alignment == "center"),
                                on_toggle_action=bkt.Callback(lambda pressed: TextShapes.settings_setter("sticker_alignment", "center")),
                            ),
                            bkt.ribbon.ToggleButton(
                                label="Rechts",
                                screentip="Sticker-Ausrichtung Rechts",
                                supertip="Setzt die Position beim Einfügen der Sticker auf rechts.",
                                get_pressed=bkt.Callback(lambda: TextShapes.sticker_alignment == "right"),
                                on_toggle_action=bkt.Callback(lambda pressed: TextShapes.settings_setter("sticker_alignment", "right")),
                            ),
                        ]
                    ),
                    bkt.ribbon.Menu(
                        label="Schriftgröße",
                        supertip="Schriftgrößenoptionen für Sticker einstellen",
                        children=[
                            bkt.ribbon.ToggleButton(
                                label="10",
                                screentip="Sticker-Schriftgröße 10",
                                supertip="Setzt die Schriftgröße beim Einfügen der Sticker auf 10",
                                get_pressed=bkt.Callback(lambda: TextShapes.sticker_fontsize == 10),
                                on_toggle_action=bkt.Callback(lambda pressed: TextShapes.settings_setter("sticker_fontsize", 10)),
                            ),
                            bkt.ribbon.ToggleButton(
                                label="11",
                                screentip="Sticker-Schriftgröße 11",
                                supertip="Setzt die Schriftgröße beim Einfügen der Sticker auf 11",
                                get_pressed=bkt.Callback(lambda: TextShapes.sticker_fontsize == 11),
                                on_toggle_action=bkt.Callback(lambda pressed: TextShapes.settings_setter("sticker_fontsize", 11)),
                            ),
                            bkt.ribbon.ToggleButton(
                                label="12",
                                screentip="Sticker-Schriftgröße 12",
                                supertip="Setzt die Schriftgröße beim Einfügen der Sticker auf 12",
                                get_pressed=bkt.Callback(lambda: TextShapes.sticker_fontsize == 12),
                                on_toggle_action=bkt.Callback(lambda pressed: TextShapes.settings_setter("sticker_fontsize", 12)),
                            ),
                            bkt.ribbon.ToggleButton(
                                label="14",
                                screentip="Sticker-Schriftgröße 14",
                                supertip="Setzt die Schriftgröße beim Einfügen der Sticker auf 14",
                                get_pressed=bkt.Callback(lambda: TextShapes.sticker_fontsize == 14),
                                on_toggle_action=bkt.Callback(lambda pressed: TextShapes.settings_setter("sticker_fontsize", 14)),
                            ),
                        ]
                    ),
                ])
            ]
        ),
        bkt.ribbon.Button(
            id = 'underlined_textbox',
            label = u"Unterstrichene Textbox",
            image = "underlined_textbox",
            screentip="Unterstrichene Textbox einfügen",
            supertip="Füge eine Textbox mit Linie unten am Shape auf dem aktuellen Slide ein.",
            on_action=bkt.Callback(TextShapes.addUnderlinedTextbox)
        ),
        bkt.ribbon.MenuSeparator(title="Aufzählungszeichen"),
        bkt.ribbon.Button(
            id="bullet_fixing",
            label=u"Aufzählungszeichen korrigieren",
            supertip=u"Aufzählungszeichen werden korrigiert. Der Stil wird vom Textplatzhalter auf dem Masterslide übernommen. Betrifft: Symbol, Symbol-/Textfarbe, Absatzeinzug/-abstand",
            image_mso='MultilevelListGallery',
            on_action=bkt.Callback(BulletStyle.fix_bullet_style),
            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
        ),
        bkt.ribbon.MenuSeparator(),
        bkt.ribbon.ColorGallery(
            id = 'bullet_color',
            label=u'Farbe ändern',
            screentip="Bullet Point Farbe ändern",
            supertip="Ändert die Farbe der gewählten Bullet Points.",
            on_rgb_color_change = bkt.Callback(BulletStyle.set_bullet_color_rgb, selection=True, shapes=True),
            on_theme_color_change = bkt.Callback(BulletStyle.set_bullet_theme_color, selection=True, shapes=True),
            get_selected_color = bkt.Callback(BulletStyle.get_bullet_color_rgb, selection=True, shapes=True),
            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
            children=[
                bkt.ribbon.Button(
                    id="bullet_color_auto",
                    label="Automatisch",
                    screentip="Bullet Point Farbe automatisch bestimmen",
                    supertip="Bullet Point Farbe wird automatisch anhand der Textfarbe bestimmt.",
                    on_action=bkt.Callback(BulletStyle.set_bullet_color_auto, selection=True, shapes=True),
                    image_mso="ColorBlack",
                ),
            ]
        ),
        bkt.ribbon.SymbolsGallery(
            id="bullet_symbol",
            label=u"Symbol ändern",
            screentip="Bullet Point Symbol ändern",
            supertip="Ändert das Symbol der gewählten Bullet Points.",
            symbols = Characters.lists,
            on_symbol_change = bkt.Callback(BulletStyle.set_bullet_symbol, selection=True, shapes=True),
            get_selected_symbol = bkt.Callback(BulletStyle.get_bullet_symbol, selection=True, shapes=True),
            get_enabled = bkt.apps.ppt_shapes_or_text_selected
        ),
        bkt.ribbon.MenuSeparator(title="Textoperationen"),
        bkt.ribbon.Button(
            id = 'text_in_shape',
            label = u"Text in Shape",
            image_mso = "TextBoxInsert",
            screentip="Text in Shape kombinieren",
            supertip="Kopiere den Text eines Text-Shapes in das zweite markierte Shape und löscht das Text-Shape.",
            on_action=bkt.Callback(TextOnShape.textIntoShape, shapes=True, shapes_min=2),
            get_enabled = bkt.apps.ppt_shapes_min2_selected,
        ),
        bkt.ribbon.Button(
            id = 'text_on_shape',
            label = u"Text auf Shape",
            image_mso = "TableCellCustomMarginsDialog",
            screentip="Text auf Shape zerlegen",
            supertip="Überführe jeweils den Textinhalt der markierten Shapes in ein separates Text-Shape.",
            on_action=bkt.Callback(TextOnShape.textOutOfShape, shapes=True, context=True),
            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
        ),
        bkt.ribbon.MenuSeparator(),
        bkt.ribbon.Button(
            id = 'decompose_text',
            label = u"Shape-Text zerlegen",
            image_mso = "TraceDependents",
            supertip="Zerlege die markierten Shapes anhand der Text-Absätze in mehrere Shapes. Pro Absatz wird ein Shape mit dem entsprechenden Text angelegt.",
            on_action=bkt.Callback(SplitTextShapes.splitShapesByParagraphs, shapes=True, context=True),
            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
        ),
        bkt.ribbon.Button(
            id = 'compose_text',
            label = u"Shape-Text zusammenführen",
            image_mso = "TracePrecedents",
            supertip="Führe die markierten Shapes in ein Shape zusammen. Der Text aller Shapes wird übernommen und aneinandergehängt.",
            on_action=bkt.Callback(SplitTextShapes.joinShapesWithText, shapes=True, shapes_min=2),
            get_enabled = bkt.apps.ppt_shapes_min2_selected,
        ),
        bkt.ribbon.MenuSeparator(),
        bkt.ribbon.Button(
            id = 'text_truncate',
            label="Shape-Texte löschen",
            image_mso='ReviewDeleteMarkup',
            supertip="Text aller gewählten Shapes löschen.",
            on_action=bkt.Callback(TextPlaceholder.text_truncate, shapes=True),
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
                    on_action=bkt.Callback(TextPlaceholder.text_replace, shapes=True),
                    get_enabled=bkt.apps.ppt_shapes_or_text_selected,
                ),
                bkt.ribbon.Menu(label="Shape-Texte ersetzen Menü", supertip="Text mit vordefinierten Standard-Platzhaltern ersetzen", children=[
                    bkt.ribbon.Button(
                        id = 'text_tbd',
                        label="… mit »tbd«",
                        image_mso='TextDialog',
                        screentip="Text mit »tbd« ersetzen",
                        supertip="Text aller gewählten Shapes mit 'tbd' ersetzen.",
                        on_action=bkt.Callback(TextPlaceholder.text_tbd, shapes=True),
                    ),
                    bkt.ribbon.Button(
                        id = 'text_lorem',
                        label="… mit Lorem ipsum",
                        image_mso='TextDialog',
                        screentip="Text mit Lorem ipsum ersetzen",
                        supertip="Text aller gewählten Shapes mit langem 'Lorem ipsum'-Platzhaltertext ersetzen.",
                        on_action=bkt.Callback(TextPlaceholder.text_lorem, shapes=True),
                    ),
                    bkt.ribbon.Button(
                        id = 'text_counter',
                        label="… mit Nummerierung",
                        image_mso='TextDialog',
                        screentip="Text mit Nummerierung ersetzen",
                        supertip="Text aller gewählten Shapes durch Nummerierung ersetzen.",
                        on_action=bkt.Callback(TextPlaceholder.text_counter, shapes=True),
                    ),
                    bkt.ribbon.MenuSeparator(),
                    bkt.ribbon.Button(
                        id = 'text_replace2',
                        label="… mit benutzerdefiniertem Text",
                        image_mso='ReplaceDialog',
                        screentip="Text mit eigener Eingabe ersetzen",
                        supertip="Text aller gewählten Shapes mit im Dialogfeld eingegebenen Text ersetzen.",
                        on_action=bkt.Callback(TextPlaceholder.text_replace, shapes=True),
                        get_enabled=bkt.apps.ppt_shapes_or_text_selected,
                    ),
                ])
            ]
        ),
    ]
)

text_splitbutton = bkt.ribbon.SplitButton(
    id="textbox_insert_splitbutton",
    show_label=False,
    children=[
        bkt.ribbon.ToggleButton(
            id="textbox_insert",
            label="Textfeld zeichnen",
            image_mso="TextBoxInsert",
            supertip="Zeichnen Sie ein Textfeld an einer beliebigen Stelle.\n\nMit gedrückter Umschalt-Taste wird eine unterstrichene Textbox eingefügt.\n\nMit gedrückter Strg-Taste wird ein Sticker eingefügt.",
            on_toggle_action=bkt.Callback(TextShapes.textbox_insert, context=True),
            get_pressed=bkt.Callback(TextShapes.textbox_pressed, context=True),
            get_enabled=bkt.Callback(TextShapes.textbox_enabled, context=True),
        ),
        # bkt.mso.toggleButton.TextBoxInsert,
        text_menu
    ]
)


paragraph_group = bkt.ribbon.Group(
    id="bkt_paragraph_group",
    label = u"Absatz",
    image_mso='FormattingMarkDropDown',
    children = [
        bkt.ribbon.Menu(
            label=u"Textbox",
            imageMso="FormattingMarkDropDown",
            supertip="Einstellungen für die Textbox ändern",
            children = [
                bkt.ribbon.ToggleButton(
                    id = 'wordwrap',
                    label="WordWrap",
                    image_mso="FormattingMarkDropDown",
                    screentip="Text in Form umbrechen",
                    supertip="Konfiguriere die Textoption auf 'Text in Form umbrechen'.",
                    on_toggle_action=bkt.Callback(Absatz.set_word_wrap , shapes=True, require_text=True),
                    get_pressed=bkt.Callback(Absatz.get_word_wrap , shapes=True, require_text=True),
                    get_enabled = bkt.get_enabled_auto,
                ),
                bkt.ribbon.ToggleButton(
                    id = 'autosize',
                    label="AutoSize",
                    image_mso="SmartArtLargerShape",
                    screentip="Größe der Form anpassen",
                    supertip="Konfiguriere die Textoption auf 'Größe der Form dem Text anpassen' bzw. 'Größe nicht automatisch anpassen'.",
                    on_toggle_action=bkt.Callback(Absatz.set_auto_size , shapes=True, require_text=True),
                    get_pressed=bkt.Callback(Absatz.get_auto_size , shapes=True, require_text=True),
                    get_enabled = bkt.get_enabled_auto,
                ),
                bkt.ribbon.MenuSeparator(),
                bkt.mso.control.TextAlignMoreOptionsDialog
            ]
        ),
        ParSpaceBefore(
            id = 'par_sep_top',
            show_label=False,
            # label=u"Absatzabstand oben",
            # image_mso='WordOpenParaAbove',
            # screentip="Oberen Absatzabstand",
            # supertip="Ändere den Absatzabstand vor dem Absatz auf das angegebene Maß (in pt).",
            #attr='SpaceBefore'
        ),
        ParSpaceAfter(
            id = 'par_sep_bottom',
            show_label=False,
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
    label = u"Absatzeinzug",
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


