# -*- coding: utf-8 -*-
'''
Created on 2018-05-29
@author: Florian Stallmann
'''



import System
from System.Collections.ObjectModel import Collection

from collections import namedtuple

import bkt.ui
notify_property = bkt.ui.notify_property

from bkt.helpers import Resources

Language = namedtuple("Language", "Flag Label Tag")


class ViewModel(bkt.ui.ViewModelSingleton):
    
    def __init__(self, lang_dict):
        super(ViewModel, self).__init__()

        self.languages = Collection[Language]()
        for key in sorted(lang_dict):
            lang = lang_dict[key]
            self.languages.Add(Language(
                System.Uri(Resources.images.locate(lang[3])),
                lang[2],
                key,
            ))


class LanguageWindow(bkt.ui.WpfWindowAbstract):
    _xamlname = 'language'
    # _vm_class = ViewModel
    
    def __init__(self, context, model):
        self._model = model
        self._vm = ViewModel(model.langs)
        super(LanguageWindow, self).__init__(context)
    
    def get_lang_code(self, tag):
        return self._model.langs[tag][1]
    
    def setPresentation(self, sender, event):
        lang_code = self.get_lang_code(sender.Tag)
        self._model.set_language_for_presentation(self._context.presentation, lang_code)
        self.Close()
    
    def setSlides(self, sender, event):
        lang_code = self.get_lang_code(sender.Tag)
        self._model.set_language_for_slides(self._context.slides, lang_code)
        self.Close()