# -*- coding: utf-8 -*-
'''
Created on 13.05.2020

@author: fstallmann
'''



import unittest

from tests.mock_shape import Shape

from bkt.library import algorithms


def float_to_int(var):
    if isinstance(var, list):
        return [float_to_int(f) for f in var]
    elif isinstance(var, tuple):
        return tuple(float_to_int(f) for f in var)
    else:
        return round(var)


class MathTests(unittest.TestCase):
    def test_median(self):
        self.assertAlmostEqual(algorithms.median([4,2,5,8]), 4.5)
        self.assertAlmostEqual(algorithms.median([4,2,6.3,8,5.21]), 5.21)
        self.assertAlmostEqual(algorithms.median([4,2,6,8,5,4,4,4,4,4]), 4)
        
    def test_mean(self):
        self.assertAlmostEqual(algorithms.mean([4,2,5,8]), 4.75)
        self.assertAlmostEqual(algorithms.mean([40,20,30]), 30)
        self.assertAlmostEqual(algorithms.mean([0.3,0.6,0.9]), 0.6)
        
    def test_mid_point(self):
        self.assertTupleEqual(algorithms.mid_point([(0,0),(1,1)]), (0.5,0.5))
        self.assertTupleEqual(algorithms.mid_point([(1,1),(1,3),(3,1),(3,3)]), (2,2))
        
    def test_is_close(self):
        self.assertTrue(algorithms.is_close(0.01, 0.0099999999999999))
        self.assertFalse(algorithms.is_close(0.01, 0.015))
        
    def test_get_bounds(self):
        self.assertTupleEqual(algorithms.get_bounds([(0,0),(1,1)]), (0,0,1,1))
        self.assertTupleEqual(algorithms.get_bounds([(1,1),(1,3),(3,1),(3,3)]), (1,1,2,2))
        
    def test_rotate_point(self):
        self.assertTupleEqual(float_to_int(algorithms.rotate_point(1,1, 0,0, 0)), (1,1))
        self.assertTupleEqual(float_to_int(algorithms.rotate_point(1,1, 0,0, 90)), (1,-1))
        self.assertTupleEqual(float_to_int(algorithms.rotate_point(1,1, 0,0, -90)), (-1,1))
        self.assertTupleEqual(float_to_int(algorithms.rotate_point(1,1, 0,0, 180)), (-1,-1))
        self.assertTupleEqual(float_to_int(algorithms.rotate_point(1,1, 0,0, 360)), (1,1))
        self.assertTupleEqual(float_to_int(algorithms.rotate_point(2,2, 1,1, 180)), (0,0))


class ShapeTests(unittest.TestCase):
    def setUp(self):
        s0 = Shape(left=1,top=2,width=3,height=4)
        s1 = Shape(left=1,top=2,width=3,height=4)
        s1.rotation = 180
        s2 = Shape(left=1,top=2,width=3,height=5)
        s2.rotation = 90
        s3 = Shape(left=0,top=5,width=6,height=3)
        s4 = Shape(left=7,top=5,width=1,height=25)

        self.shapes = [s0,s1,s2,s3,s4]
        
    def test_get_bounds_shapse(self):
        self.assertTupleEqual(algorithms.get_bounds_shapes(self.shapes), (0,2,8,28))
        
    def test_rotate_point_by_shape_rotation(self):
        self.assertTupleEqual(float_to_int(algorithms.rotate_point_by_shape_rotation(1,1, self.shapes[0])), (1,1))
        self.assertTupleEqual(float_to_int(algorithms.rotate_point_by_shape_rotation(1,1, self.shapes[1])), (-1,-1))
        self.assertTupleEqual(float_to_int(algorithms.rotate_point_by_shape_rotation(1,1, self.shapes[2])), (-1,1))
        
    def test_get_bounding_nodes(self):
        self.assertListEqual(float_to_int(algorithms.get_bounding_nodes(self.shapes[0])), [[1,2],[1,6],[4,6],[4,2]])
        self.assertListEqual(float_to_int(algorithms.get_bounding_nodes(self.shapes[1])), [[4,6],[4,2],[1,2],[1,6]])
        self.assertListEqual(float_to_int(algorithms.get_bounding_nodes(self.shapes[2])), [[5,3],[0,3],[0,6],[5,6]])


class ColorTests(unittest.TestCase):
    def test_rgb_from_ole(self):
        self.assertTupleEqual(algorithms.get_rgb_from_ole(0), (0,0,0))
        self.assertTupleEqual(algorithms.get_rgb_from_ole(8421504), (128,128,128))
        self.assertTupleEqual(algorithms.get_rgb_from_ole(16777215), (255,255,255))
        
    def test_ole_from_rgb(self):
        self.assertEqual(algorithms.get_ole_from_rgb(0,0,0), 0)
        self.assertEqual(algorithms.get_ole_from_rgb(128,128,128), 8421504)
        self.assertEqual(algorithms.get_ole_from_rgb(255,255,255), 16777215) #white

    def test_brightness_from_rgb(self):
        self.assertEqual(algorithms.get_brightness_from_rgb(0,0,0), 0)
        self.assertEqual(algorithms.get_brightness_from_rgb(128,128,128), 128)
        self.assertEqual(algorithms.get_brightness_from_rgb(255,255,255), 255)