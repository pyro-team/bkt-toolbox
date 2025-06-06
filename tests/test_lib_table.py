# -*- coding: utf-8 -*-
'''
Created on 13.05.2020

@author: fstallmann
'''



import unittest

from tests.mock_shape import Shape

from bkt.library.table import TableRecognition


class TableTestsEasy(unittest.TestCase):
    def setUp(self):
        s0 = Shape(left=1,top=1, width=2,height=2)
        s1 = Shape(left=1,top=4, width=2,height=2)
        s2 = Shape(left=4,top=1, width=2,height=2)
        s3 = Shape(left=4,top=4, width=2,height=2)
        s4 = Shape(left=1,top=7, width=2,height=2)
        s5 = Shape(left=4,top=7, width=2,height=2)

        self.shapes = [s0,s1,s2,s3,s4,s5]
        
    def test_table_recognition(self):
        tr = TableRecognition(self.shapes)
        tr.run()

        self.assertTupleEqual(tr.dimension, (3,2)) #rows,columns
        self.assertEqual(tr.first_top, self.shapes[0])
        self.assertEqual(tr.first_left, self.shapes[0])

        self.assertEqual(tr.min_spacing_rows(), 1)
        self.assertEqual(tr.min_spacing_cols(), 1)
        self.assertEqual(tr.median_spacing(), 1)
        self.assertEqual(tr.column_count(), 2)

        self.assertTupleEqual(tr.get_bounds(), (1,1,5,8))

        self.assertListEqual(list(tr.column(0)), [self.shapes[0], self.shapes[1], self.shapes[4]])
        self.assertListEqual(list(tr.column(1)), [self.shapes[2], self.shapes[3], self.shapes[5]])

    def test_table_align_no_change(self):
        tr = TableRecognition(self.shapes)
        tr.run()

        prev_values = [(s.left, s.top, s.width, s.height) for s in self.shapes]
        tr.align(tr.median_spacing())
        after_values = [(s.left, s.top, s.width, s.height) for s in self.shapes]
        
        self.assertListEqual(prev_values, after_values)

    
    #TODO:
    #already table aligned
    #shapes have minor changes in position or width
    #two shapes at the same position
    #missing cell
    #cell alignment
    #cell size change
    #distribute rows/cols