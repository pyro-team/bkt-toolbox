# -*- coding: utf-8 -*-
'''
Created on 14.07.2020

@author: fstallmann
'''



from .mock import OfficeMock

class TextFrame(OfficeMock):
    def __init__(self, **kwargs):
        self._attributes = {
            "autosize": 1,
            "wordwrap": -1,
            "hastext": -1,
        }
        self._attributes.update(kwargs)

class Shape(OfficeMock):
    def __init__(self, **kwargs):
        self._attributes = {
            "name": "Shape %s" % id(self),
            "type": 1, #msoAutoShape
            "autoshapetype": 1,
            "left": 0,
            "top": 0,
            "width": 1,
            "height": 1,
            "rotation": 0,
            "lockaspectratio": 0,
            "hastextframe": -1,
            "textframe": TextFrame(),
        }
        self._attributes.update(kwargs)
        self._attributes["textframe2"] = self._attributes["textframe"]
