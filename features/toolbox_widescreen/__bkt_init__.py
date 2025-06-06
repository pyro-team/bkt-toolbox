# -*- coding: utf-8 -*-



class BktFeature(object):
    name            = "PowerPoint Toolbox Widescreen"
    relevant_apps   = ["Microsoft PowerPoint"]

    conflicts       = ["toolbox", "toolbox_variation"]
    dependencies    = []
    
    @staticmethod
    def contructor():
        from . import my_toolbox