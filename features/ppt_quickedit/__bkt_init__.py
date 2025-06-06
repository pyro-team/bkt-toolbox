# -*- coding: utf-8 -*-



class BktFeature(object):
    name            = "PowerPoint QuickEdit"
    relevant_apps   = ["Microsoft PowerPoint"]
    
    @staticmethod
    def contructor():
        from . import quickedit