# -*- coding: utf-8 -*-

class BktFeature(object):
    name            = "PowerPoint Shape-Tabellen"
    relevant_apps   = ["Microsoft PowerPoint"]
    
    @staticmethod
    def contructor():
        import shape_tables
