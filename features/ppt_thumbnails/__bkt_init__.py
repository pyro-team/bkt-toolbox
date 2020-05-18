# -*- coding: utf-8 -*-

from __future__ import absolute_import

class BktFeature(object):
    name            = "PowerPoint Thumbnails"
    relevant_apps   = ["Microsoft PowerPoint"]
    
    @staticmethod
    def contructor():
        from . import thumbnails
