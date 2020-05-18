# -*- coding: utf-8 -*-

from __future__ import absolute_import

class BktFeature(object):
    name            = "PowerPoint Shape-Tabellen"
    relevant_apps   = ["Microsoft PowerPoint"]
    
    @staticmethod
    def contructor():
        from . import shape_tables
