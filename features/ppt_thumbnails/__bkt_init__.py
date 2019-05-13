# -*- coding: utf-8 -*-

class BktFeature(object):
    name            = "PowerPoint Thumbnails"
    relevant_apps   = ["Microsoft PowerPoint"]
    
    @staticmethod
    def contructor():
        import thumbnails
