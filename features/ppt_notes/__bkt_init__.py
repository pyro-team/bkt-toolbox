# -*- coding: utf-8 -*-



class BktFeature(object):
    name            = "PowerPoint Notes"
    relevant_apps   = ["Microsoft PowerPoint"]
    
    @staticmethod
    def contructor():
        from . import notes