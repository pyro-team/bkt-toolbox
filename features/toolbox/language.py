# -*- coding: utf-8 -*-
'''
Created on 06.09.2018

@author: fstallmann
'''



import bkt
import bkt.library.powerpoint as pplib


class LangSetter(object):
    active_langs = bkt.settings.get("toolbox.languages_checked", ['de', 'us', 'gb'])

    #list: https://docs.microsoft.com/de-de/office/vba/api/office.msolanguageid
    langs = {
        #[key, id, name, image]
        'de': ('de', 1031, "Deutsch",               "flag_de"),
        'us': ('us', 1033, "US English",            "flag_us"),
        'gb': ('uk', 2057, "UK English",            "flag_gb"), #keep uk as id for button for backwards compatibility
        'at': ('at', 3079, "Deutsch (Österreich)",  "flag_at"),
        'it': ('it', 1040, "Italienisch",           "flag_it"),
        'fr': ('fr', 1036, "Französisch",           "flag_fr"),
        'es': ('es', 3082, "Spanisch",              "flag_es"),
        'ru': ('ru', 1049, "Russisch",              "flag_ru"),
        'cz': ('cz', 1029, "Tschechisch",           "flag_cz"),
        'dk': ('dk', 1030, "Dänisch",               "flag_dk"),
        'nl': ('nl', 1043, "Holländisch",           "flag_nl"),
        'pl': ('pl', 1045, "Polnisch",              "flag_pl"),
        'pt': ('pt', 2070, "Portugisisch",          "flag_pt"),
        'se': ('se', 1053, "Schwedisch",            "flag_se"),
        'tr': ('tr', 1055, "Türkisch",              "flag_tr"),
    }

    @classmethod
    def show_lang_dialog(cls, context):
        from .dialogs.language import LanguageWindow
        LanguageWindow.create_and_show_dialog(context, cls)

    @classmethod
    def get_languages(cls):
        for lang in cls.active_langs:
            yield cls.langs[lang]
    
    @classmethod
    def edit_active_language(cls):
        lang_list = bkt.ui.show_user_input("Liste möglicher Sprachen bearbeiten.\nVerfügbar sind: {}.".format(",".join(sorted(cls.langs.keys()))), "Sprachen-Liste", ",".join(cls.active_langs))
        if lang_list is None:
            return
        
        cls.active_langs = [lang for lang in lang_list.split(",") if lang in cls.langs]
        if len(cls.active_langs) == 0:
            cls.active_langs = ['de', 'us', 'gb']
        bkt.settings["toolbox.languages_checked"] = cls.active_langs

        bkt.message("Die Änderungen werden nach einem PowerPoint-Neustart sichtbar.")

    @classmethod
    def get_button(cls, language, idtag=""):
        return bkt.ribbon.Button(
                id = 'lang_'+language[0]+idtag,
                label=language[2],
                image=language[3],
                screentip="Sprache auf " + language[2] + " ändern",
                supertip="Setze Sprache für ausgewählten Text bzw. alle ausgewählten Shapes.\nWenn mehrere Folien ausgewählt sind, werden alle Shapes der gewählten Folien geändert.\nWenn nichts ausgewählt ist, werden alle Shapes in der Präsentation sowie die Standardsprache geändert.",
                on_action=bkt.Callback(lambda context, selection, presentation: cls.set_language(context, selection, presentation, language[1]), context=True, selection=True, presentation=True)
            )

    @classmethod
    def _get_words_in_selection(cls, selection):
        cursor_start = selection.TextRange2.Start
        cursor_end   = cursor_start + selection.TextRange2.Length
        words = selection.TextRange2.Parent.TextRange.Words()
        word_first = words.Count #set last word as default (count is always min 1)
        word_last  = words.Count #set last word as default (count is always min 1)
        for i,word in enumerate(words):
            word_end = word.Start+word.Length
            if i < word_first and cursor_start < word_end:
                word_first = i+1
            if i < word_last and cursor_end < word_end:
                word_last = i+1
                break #we can stop loop here
        return selection.TextRange2.Parent.TextRange.Words(word_first, word_last-word_first+1)

    @classmethod
    def set_language(cls, context, selection, presentation, lang_code):
        shapes = pplib.get_shapes_from_selection(selection)
        slides = pplib.get_slides_from_selection(selection)

        # Set language for selected text, shapes, slides or whole presentation
        if selection.Type == 3: #text selected
            textrange = cls._get_words_in_selection(selection)
            textrange.LanguageID = lang_code
            # selection.TextRange2.LanguageID = lang_code
        elif len(shapes) > 0:
            #bkt.message("Setze Sprache für Shapes: " + str(len(shapes)))
            cls.set_language_for_shapes(shapes, lang_code)
        elif len(slides) != presentation.slides.count and (len(slides) > 1 or context.app.ActiveWindow.ActivePane.ViewType in [7, 11]): #7=ppViewSlideSorter, 11=ppViewThumbnails
            #bkt.message("Setze Sprache für Slides: " + str(len(slides)))
            if len(slides) > 1 and not bkt.message.confirmation("Sprache aller Shapes auf ausgewählten Folien ändern?"):
                return
            cls.set_language_for_slides(slides, lang_code)
        else:
            #bkt.message("Setze Sprache für Präsentation")
            if not bkt.message.confirmation("Sprache aller Shapes auf allen Folien (inkl. Standardsprache der Präsentation) ändern?"):
                return
            cls.set_language_for_presentation(presentation, lang_code)

    @classmethod
    def set_language_for_presentation(cls, presentation, lang_code):
        presentation.DefaultLanguageID = lang_code
        cls.set_language_for_slides(presentation.slides, lang_code)

    @classmethod
    def set_language_for_slides(cls, slides, lang_code):
        for slide in slides:
            cls.set_language_for_shapes(slide.shapes, lang_code, False)

    @classmethod
    def set_language_for_shapes(cls, shapes, lang_code, from_selection=True):
        for textframe in pplib.iterate_shape_textframes(shapes, from_selection):
            try:
                textframe.TextRange.LanguageID = lang_code
            except:
                #skip errors, e.g. for certain chart types
                continue
    
    @classmethod
    def get_dynamicmenu_content(cls):
        return bkt.ribbon.Menu(
            xmlns="http://schemas.microsoft.com/office/2009/07/customui",
            id=None,
            children=[
                cls.get_button(lang)
                for lang in cls.get_languages()
            ]
        )


sprachen_gruppe = bkt.ribbon.Group(
    id="bkt_language_group",
    label="Sprache",
    image_mso="GroupLanguage",
    auto_scale=True,
    children=[
        LangSetter.get_button(lang, "_group")
        for lang in LangSetter.get_languages()
    ] + [
        bkt.ribbon.DialogBoxLauncher(
            label="Wählbare Sprachen editieren…",
            supertip="Öffnet Dialog um wählbare Sprachen zu ändern.",
            on_action=bkt.Callback(LangSetter.edit_active_language),
        )
        # bkt.ribbon.DialogBoxLauncher(idMso='SetLanguage')
    ]
)

sprachen_menu = bkt.ribbon.SplitButton(
    id="lang_change_menu",
    children=[
        bkt.ribbon.Button(
            label='Sprache ändern…',
            image_mso='GroupLanguage',
            supertip="Zeigt Dialog zur Auswahl der anzuwendenden Sprache auf die Präsentation oder die ausgewählten Folien.",
            on_action=bkt.Callback(LangSetter.show_lang_dialog)
        ),
        bkt.ribbon.Menu(
            label="Sprache ändern",
            supertip="Sprache der Rechtschreibkorrektur für mehrere Shapes, Folien oder die ganze Präsentation anpassen",
            image_mso="GroupLanguage",
            children=[
                bkt.ribbon.MenuSeparator(title="Sprache von Shapes oder Folien ändern"),
                bkt.ribbon.Button(
                    label="Alle Sprachen anzeigen…",
                    image_mso='GroupLanguage',
                    supertip="Zeigt Dialog zur Auswahl der anzuwendenden Sprache auf die Präsentation oder die ausgewählten Folien.",
                    on_action=bkt.Callback(LangSetter.show_lang_dialog),
                ),
            ] + [
                LangSetter.get_button(lang)
                for lang in LangSetter.get_languages()
            ] + [
                bkt.ribbon.Button(
                    label="Wählbare Sprachen editieren…",
                    supertip="Öffnet Dialog um wählbare Sprachen zu ändern.",
                    on_action=bkt.Callback(LangSetter.edit_active_language),
                ),
                bkt.ribbon.MenuSeparator(),
                bkt.mso.button.SetLanguage,
                bkt.mso.button.Spelling,
            ]
        )
    ]
)