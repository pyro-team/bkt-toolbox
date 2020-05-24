# -*- coding: utf-8 -*-
'''
Created on 21.12.2017

@author: fstallmann
'''

from __future__ import absolute_import

import logging

import bkt.ui

from ..stateshapes import StateShape


class StateShapePopup(bkt.ui.WpfWindowAbstract):
    # _filename = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'popups', 'stateshapes.xaml')
    _xamlname = 'stateshapes'
    '''
    class representing a popup-dialog for a stateshape
    '''
    
    def __init__(self, context=None):
        self.IsPopup = True

        super(StateShapePopup, self).__init__(context)

    def btnprev(self, sender, event):
        try:
            #always use ShapeRange, never ChildShapeRange
            shapes = list(iter(self._context.selection.ShapeRange))
            StateShape.previous_state(shapes)
        except:
            logging.exception("Error in StateShape popup: %s")

    def btnnext(self, sender, event):
        try:
            #always use ShapeRange, never ChildShapeRange
            shapes = list(iter(self._context.selection.ShapeRange))
            StateShape.next_state(shapes)
        except:
            logging.exception("Error in StateShape popup: %s")


#initialization function called by contextdialogs.py
create_window = StateShapePopup

def trigger_doubleclick(shape, context):
    try:
        StateShape.next_state([shape])
    except:
        logging.exception("Error in StateShape popup: %s")