# -*- coding: utf-8 -*-



class BktFeature(object):
    name            = "PowerPoint Toolbox"
    relevant_apps   = ["Microsoft PowerPoint"]
    
    @staticmethod
    def contructor():
        from . import toolbox_powerpoint