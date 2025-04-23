# -*- coding: utf-8 -*-



class BktFeature(object):
    name            = "PowerPoint Consolidation-Split"
    relevant_apps   = ["Microsoft PowerPoint"]
    
    @staticmethod
    def contructor():
        from . import consolsplit
