# -*- coding: utf-8 -*-
'''
Created on 2020-08-24
@author: Florian Stallmann
'''

from __future__ import absolute_import

import logging
import re #regex
import locale #to format number for clipboard and supertip

import bkt.dotnet as dotnet
Forms = dotnet.import_forms() #required to copy text to clipboard

import bkt
import bkt.library.powerpoint as pplib

from bkt.library.algorithms import mean

import bkt.ui
notify_property = bkt.ui.notify_property


class Statistics(object):
    comma_langs = {
        1031: "Deutsch",
        3079: "Deutsch (Österreich)",
        1040: "Italienisch",
        1036: "Französisch",
        3082: "Spanisch",
        1049: "Russisch",
        1029: "Tschechisch",
        1030: "Dänisch",
        1043: "Holländisch",
        1045: "Polnisch",
        2070: "Portugisisch",
        1053: "Schwedisch",
        1055: "Türkisch",
    }
    dot_langs = {
        1033: "US English",
        2057: "UK English",
    }

    _list_shapes = []
    _list_subshapes = []
    _list_txtshapes = []
    _list_numbers = []

    _num_chars = 0
    _num_pars = 0
    _num_lines = 0
    _num_words = 0

    @classmethod
    def shape_no(cls):
        return locale.str(len(cls._list_shapes))
    @classmethod
    def shape_sub_no(cls):
        return locale.str(len(cls._list_subshapes))
    @classmethod
    def shape_txt_no(cls):
        return locale.str(len(cls._list_txtshapes))
    @classmethod
    def shape_num_no(cls):
        return locale.str(len(cls._list_numbers))
    @classmethod
    def shape_sum(cls):
        return "{:.15n}".format(float(sum(cls._list_numbers)))
    @classmethod
    def shape_avg(cls):
        if not cls._list_numbers:
            return 0
        return "{:.15n}".format(float(mean(cls._list_numbers)))
    @classmethod
    def shape_chars(cls):
        return locale.str(cls._num_chars)
    @classmethod
    def shape_pars(cls):
        return locale.str(cls._num_pars)
    @classmethod
    def shape_lines(cls):
        return locale.str(cls._num_lines)
    @classmethod
    def shape_words(cls):
        return locale.str(cls._num_words)

    @classmethod
    def get_supertip_no(cls):
        if len(cls._list_shapes) > 30:
            part1 = "\n".join(cls._list_shapes[:15])
            part2 = "\n".join(cls._list_shapes[-15:])
            return part1 + "\n\n...\n\n" + part2
        else:
            return "\n".join(cls._list_shapes)

    @classmethod
    def get_supertip_sub_no(cls):
        if len(cls._list_subshapes) > 30:
            part1 = "\n".join(cls._list_subshapes[:15])
            part2 = "\n".join(cls._list_subshapes[-15:])
            return part1 + "\n\n...\n\n" + part2
        else:
            return "\n".join(cls._list_subshapes)

    @classmethod
    def get_supertip_txt_no(cls):
        if len(cls._list_txtshapes) > 30:
            part1 = "\n".join(cls._list_txtshapes[:15])
            part2 = "\n".join(cls._list_txtshapes[-15:])
            return part1 + "\n\n...\n\n" + part2
        else:
            return "\n".join(cls._list_txtshapes)

    @classmethod
    def get_supertip_num_no(cls):
        if len(cls._list_numbers) > 30:
            part1 = "\n".join("{:.15n}".format(n) for n in cls._list_numbers[:15])
            part2 = "\n".join("{:.15n}".format(n) for n in cls._list_numbers[-15:])
            return part1 + "\n\n...\n\n" + part2
        else:
            return "\n".join("{:.15n}".format(n) for n in cls._list_numbers)

    @classmethod
    def get_supertip_sum(cls):
        if len(cls._list_numbers) > 30:
            part1 = "\n+ ".join("{:.15n}".format(n) for n in cls._list_numbers[:15])
            part2 = "\n+ ".join("{:.15n}".format(n) for n in cls._list_numbers[-15:])
            return part1 + "\n\n...\n\n+ " + part2
        else:
            return "\n+ ".join("{:.15n}".format(n) for n in cls._list_numbers)

    @classmethod
    def clear_all(cls):
        cls._list_shapes = []
        cls._list_subshapes = []
        cls._list_txtshapes = []
        cls._list_numbers = []
        cls._num_chars = 0
        cls._num_pars = 0
        cls._num_lines = 0
        cls._num_words = 0

    @classmethod
    def reload_all(cls, shapes):
        def get_shp_name(shp):
            try:
                return shp.Name
            except:
                return "<No name, maybe table cell>"
        cls.clear_all()
        cls._list_shapes = [get_shp_name(shp) for shp in shapes]

        for shp in pplib.iterate_shape_subshapes(shapes):
            shp_name = get_shp_name(shp)
            cls._list_subshapes.append(shp_name)
            try:
                if shp.HasTextFrame and shp.TextFrame2.HasText:
                    cls._list_txtshapes.append(shp_name)
            except:
                pass
        for textframe in pplib.iterate_shape_textframes(shapes):
            cls._list_numbers.extend(cls.get_numbers_from_textframe(textframe))
            try:
                # cls._num_chars += textframe.TextRange.Characters().Count
                cls._num_chars += textframe.TextRange.Length
                cls._num_pars += textframe.TextRange.Paragraphs().Count
                cls._num_lines += textframe.TextRange.Lines().Count
                cls._num_words += textframe.TextRange.Words().Count
            except:
                pass

    @classmethod
    def get_numbers_from_textframe(cls, textframe):
        try:
            if textframe.TextRange.LanguageID in cls.comma_langs.keys():
                regex = r'[^\-0-9,]'
            else: #EN, US, others
                regex = r'[^\-0-9.]'
            
            # units = iter(textframe.TextRange.Words()) #issue: splits -X in 2 words: [-, X]
            units = re.split(r'[\s;]', textframe.TextRange.Text)
            for unit in units:
                try:
                    # yield float(re.sub(regex, "", unit.Text).replace(",", "."))
                    yield float(re.sub(regex, "", unit).replace(",", "."))
                except ValueError:
                    # print(unit.Text)
                    continue
        except:
            pass



class ViewModel(bkt.ui.ViewModelSingleton):
    
    @notify_property
    def num_shapes(self):
        return Statistics.shape_no()
    @notify_property
    def num_shapes_tooltip(self):
        return Statistics.get_supertip_no()
    
    @notify_property
    def num_subshapes(self):
        return Statistics.shape_sub_no()
    @notify_property
    def num_subshapes_tooltip(self):
        return Statistics.get_supertip_sub_no()
    
    @notify_property
    def num_txtshapes(self):
        return Statistics.shape_txt_no()
    @notify_property
    def num_txtshapes_tooltip(self):
        return Statistics.get_supertip_txt_no()
    
    @notify_property
    def num_numbers(self):
        return Statistics.shape_num_no()
    @notify_property
    def num_numbers_tooltip(self):
        return Statistics.get_supertip_num_no()
    
    @notify_property
    def sum_numbers(self):
        return Statistics.shape_sum()
    @notify_property
    def sum_numbers_tooltip(self):
        return Statistics.get_supertip_sum()
    
    @notify_property
    def avg_numbers(self):
        return Statistics.shape_avg()
    
    @notify_property
    def num_chars(self):
        return Statistics.shape_chars()
    
    @notify_property
    def num_pars(self):
        return Statistics.shape_pars()
    
    @notify_property
    def num_lines(self):
        return Statistics.shape_lines()
    
    @notify_property
    def num_words(self):
        return Statistics.shape_words()


class StatisticsWindow(bkt.ui.WpfWindowAbstract):
    _xamlname = 'statistics'
    _vm_class = ViewModel

    def __init__(self, context):
        try:
            Statistics.reload_all(context.shapes)
        except:
            logging.exception("error loading statistics")
        super(StatisticsWindow, self).__init__(context)
    
    def update(self, sender=None, event=None):
        Statistics.reload_all(self._context.shapes)
        self._vm.OnPropertyChanged('num_shapes')
        self._vm.OnPropertyChanged('num_shapes_tooltip')
        self._vm.OnPropertyChanged('num_subshapes')
        self._vm.OnPropertyChanged('num_subshapes_tooltip')
        self._vm.OnPropertyChanged('num_txtshapes')
        self._vm.OnPropertyChanged('num_txtshapes_tooltip')
        self._vm.OnPropertyChanged('num_numbers')
        self._vm.OnPropertyChanged('num_numbers_tooltip')
        self._vm.OnPropertyChanged('sum_numbers')
        self._vm.OnPropertyChanged('sum_numbers_tooltip')
        self._vm.OnPropertyChanged('avg_numbers')
        self._vm.OnPropertyChanged('num_chars')
        self._vm.OnPropertyChanged('num_pars')
        self._vm.OnPropertyChanged('num_lines')
        self._vm.OnPropertyChanged('num_words')
    
    def cancel(self, sender, event):
        self.Close()
        Statistics.clear_all()
    
    def copy(self, sender, event):
        try:
            Forms.Clipboard.SetText(sender.Tag)
            bkt.message("In Zwischenablage kopiert!")
        except:
            logging.exception("error copying data")
            bkt.message.error("Fehler beim Kopieren!")