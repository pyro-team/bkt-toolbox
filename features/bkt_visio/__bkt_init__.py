# -*- coding: utf-8 -*-



class BktFeature(object):
    name            = "Visio Toolbox"
    relevant_apps   = ["Microsoft Visio"]
    
    @staticmethod
    def contructor():
        from . import visiotoolbox