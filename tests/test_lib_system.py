# -*- coding: utf-8 -*-
'''
Created on 13.05.2020

@author: fstallmann
'''

from __future__ import absolute_import

import unittest

from bkt.library import system


class ShapeMock(object):
    def __init__(self, x=0):
        self.x = x


class DeltaApplyTests(unittest.TestCase):
    def setUp(self):
        self.shapes = [
            ShapeMock(1),
            ShapeMock(1),
            ShapeMock(4),
            ShapeMock(0),
        ]
    
    def test_apply_delta_noalt(self):
        #simulate alt key not pressed
        system.get_key_state = lambda x: False

        self.assertListEqual([s.x for s in self.shapes], [1,1,4,0])
        system.apply_delta_on_ALT_key(lambda shape, value: setattr(shape, "x", value), lambda shape: shape.x, self.shapes, 3)
        self.assertListEqual([s.x for s in self.shapes], [3,3,3,3])
    
    def test_apply_delta_alt(self):
        #simulate alt key pressed
        system.get_key_state = lambda x: True

        self.assertListEqual([s.x for s in self.shapes], [1,1,4,0])
        system.apply_delta_on_ALT_key(lambda shape, value: setattr(shape, "x", value), lambda shape: shape.x, self.shapes, 3)
        self.assertListEqual([s.x for s in self.shapes], [3,3,6,2])

        system.apply_delta_on_ALT_key(lambda shape, value: setattr(shape, "x", value), lambda shape: shape.x, self.shapes, 1.5)
        self.assertListEqual([s.x for s in self.shapes], [1.5,1.5,4.5,0.5])


class MessageTests(unittest.TestCase):
    def test_message(self):
        self.assertTrue(system.message.confirmation("Confirm this is a confirmation message box with question mark!", "MessageBox Tests"))