# -*- coding: utf-8 -*-

from __future__ import absolute_import

class BktFeature(object):
    name            = "Excel Toolbox"
    relevant_apps   = ["Microsoft Excel"]
    
    @staticmethod
    def contructor():
        from . import exceltoolbox
