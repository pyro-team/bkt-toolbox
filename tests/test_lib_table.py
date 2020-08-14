# -*- coding: utf-8 -*-
'''
Created on 13.05.2020

@author: fstallmann
'''

from __future__ import absolute_import

import unittest

from tests.mock_shape import Shape

from bkt.library.table import TableRecognition, TableData


class TableDataTest(unittest.TestCase):
    def setUp(self):
        self.table = TableData.from_list(range(10), 4)
    
    def test_table_standards(self):
        table = self.table

        self.assertEqual(table.columns, 4)
        self.assertEqual(table.rows, 3)
        self.assertTupleEqual(table.dimension, (3,4))

        self.assertEqual(table.get_cell(0,0), 0)
        self.assertEqual(table.get_cell(1,3), 7)
        self.assertEqual(table.get_cell(2,3), None)
        self.assertSequenceEqual(table.get_row(2), [8,9,None,None])
        self.assertSequenceEqual(list(table.get_column(1)), [1,5,9])

        self.assertEqual(table[0,0], 0)
        self.assertEqual(table[1,3], 7)
        self.assertEqual(table[2,3], None)
        self.assertSequenceEqual(table[2], [8,9,None,None])
        self.assertSequenceEqual(list(table[None,1]), [1,5,9])

        self.assertSequenceEqual(list(iter(table)), [
            (0,0,0),(0,1,1),(0,2,2),(0,3,3),
            (1,0,4),(1,1,5),(1,2,6),(1,3,7),
            (2,0,8),(2,1,9)
            ])
    
    def test_table_errors(self):
        table = self.table

        with self.assertRaises(IndexError):
            table.get_cell(5,5)
        with self.assertRaises(IndexError):
            table.get_cell(0,5)
        with self.assertRaises(IndexError):
            table.get_cell(5,0)
    
    def test_table_transpose(self):
        table = self.table
        table.transpose()

        self.assertSequenceEqual(list(iter(table)), [
            (0,0,0),(0,1,4),(0,2,8),
            (1,0,1),(1,1,5),(1,2,9),
            (2,0,2),(2,1,6),
            (3,0,3),(3,1,7)
            ])
    
    def test_table_variation(self):
        table_rows = [
            [None, "a", "b", "c"],
            ["d", None, "e"],
            ]
        table = TableData(table_rows)

        self.assertTupleEqual(table.dimension, (2,4))
        self.assertSequenceEqual(list(table.get_column(1)), ["a", None])
        self.assertSequenceEqual(list(table.get_column(3)), ["c", None])

        table.add_rows(["f", "g", "h", None, "i"], ["j"])
        self.assertTupleEqual(table.dimension, (4,5))
        self.assertSequenceEqual(list(table.get_column(1)), ["a", None,"g", None])
        self.assertSequenceEqual(list(table.get_column(3)), ["c", None, None, None])



class TableAlignmentTest(unittest.TestCase):
    def setUp(self):
        self.table = TableData.from_list(range(10), 4)



class TableRecognitionTest(unittest.TestCase):
    def setUp(self):
        s0 = Shape(1, 1,1, 2,2)
        s1 = Shape(1, 1,4, 2,2)
        s2 = Shape(1, 4,1, 2,2)
        s3 = Shape(1, 4,4, 2,2)
        s4 = Shape(1, 1,7, 2,2)
        s5 = Shape(1, 4,7, 2,2)

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