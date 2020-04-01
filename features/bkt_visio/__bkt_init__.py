# -*- coding: utf-8 -*-

from __future__ import absolute_import

class BktFeature(object):
    name            = "Visio Toolbox"
    relevant_apps   = ["Microsoft Visio"]
    
    @staticmethod
    def contructor():
        from . import visiotoolbox