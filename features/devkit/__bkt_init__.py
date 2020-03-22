# -*- coding: utf-8 -*-

from __future__ import absolute_import

class BktFeature(object):
    name            = "BKT Development Kit"
    relevant_apps   = ["Microsoft PowerPoint", "Microsoft Excel", "Microsoft Visio", "Microsoft Word"]
    
    @staticmethod
    def contructor():
        from . import devkit