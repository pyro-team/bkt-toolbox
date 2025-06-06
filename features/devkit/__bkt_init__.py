# -*- coding: utf-8 -*-



class BktFeature(object):
    name            = "BKT Development Kit"
    relevant_apps   = ["Microsoft PowerPoint", "Microsoft Excel", "Microsoft Visio", "Microsoft Word", "Outlook"]
    
    @staticmethod
    def contructor():
        from . import devkit