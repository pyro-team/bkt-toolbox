# -*- coding: utf-8 -*-



class BktFeature(object):
    name            = "PowerPoint Custom Format Styles"
    relevant_apps   = ["Microsoft PowerPoint"]
    
    @staticmethod
    def contructor():
        from . import customformats
