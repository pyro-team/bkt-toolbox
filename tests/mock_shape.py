# -*- coding: utf-8 -*-
'''
Created on 14.07.2020

@author: fstallmann
'''

from __future__ import absolute_import

from .mock import Mock


class Shape(Mock):
    def __init__(self, autoshapetype=1, left=0, top=0, width=1, height=1):
        self._attributes = {
            "type": 1, #msoAutoShape
            "autoshapetype": autoshapetype,
            "left": left,
            "top": top,
            "width": width,
            "height": height,
            "rotation": 0,
        }
    
    def __setattr__(self, name, value):
        if name.startswith("_"):
            super(Shape, self).__setattr__(name, value)
        elif name.lower() in self._attributes:
            self._attributes[name.lower()] = value
        else:
            raise AttributeError
     
    def __getattr__(self, name):
        try:
            return self._attributes[name.lower()]
        except KeyError:
            raise AttributeError
