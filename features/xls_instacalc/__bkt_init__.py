# -*- coding: utf-8 -*-

class BktFeature(object):
    name            = "Excel InstaCalc"
    relevant_apps   = ["Microsoft Excel"]
    
    @staticmethod
    def contructor():
        from . import calc