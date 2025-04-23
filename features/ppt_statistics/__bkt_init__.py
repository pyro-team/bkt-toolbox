# -*- coding: utf-8 -*-



class BktFeature(object):
    name            = "PowerPoint Statistics"
    relevant_apps   = ["Microsoft PowerPoint"]
    
    @staticmethod
    def contructor():
        from . import statistics