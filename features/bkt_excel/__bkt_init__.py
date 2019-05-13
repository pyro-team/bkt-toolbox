# -*- coding: utf-8 -*-

class BktFeature(object):
    name            = "Excel Toolbox"
    relevant_apps   = ["Microsoft Excel"]
    
    @staticmethod
    def contructor():
        import exceltoolbox
