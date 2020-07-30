# -*- coding: utf-8 -*-
'''
Created on 13.05.2020

@author: fstallmann
'''

from __future__ import absolute_import

import unittest

from tests.mock_shape import Shape

from bkt.library import algorithms


class ShapeMockTests(unittest.TestCase):
    def setUp(self):
        self.shape = Shape()
    
    def test_mock_getters(self):
        self.assertEqual(self.shape.type, 1)
        self.assertEqual(self.shape.AutoShapeType, 1)
        self.assertEqual(self.shape.left, 0)
        self.assertEqual(self.shape.Width, 1)
        self.assertEqual(self.shape.HEIGHt, 1)

        with self.assertRaises(AttributeError):
            self.shape.does_not_exist
        
    def test_mock_setters(self):
        self.shape.left = 234.4
        self.assertAlmostEqual(self.shape.left, 234.4)
        self.assertAlmostEqual(self.shape.Left, 234.4)

        self.shape.Width = 72.4
        self.assertAlmostEqual(self.shape.width, 72.4)
        self.assertAlmostEqual(self.shape.Width, 72.4)

        with self.assertRaises(AttributeError):
            self.shape.does_not_exist = True
