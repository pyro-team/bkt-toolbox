# -*- coding: utf-8 -*-

class BktFeature(object):
    name            = "PowerPoint Toolbox Variation"
    relevant_apps   = ["Microsoft PowerPoint"]

    conflicts       = ["toolbox", "toolbox_widescreen"]
    dependencies    = []
    
    @staticmethod
    def contructor():
        import my_toolbox
