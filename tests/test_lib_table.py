# -*- coding: utf-8 -*-
'''
Created on 13.05.2020

@author: fstallmann
'''

from __future__ import absolute_import, print_function

import unittest

from tests.mock_shape import Shape

from bkt.library.table import TableRecognition, TableData, ShapeTableAlignment


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
        self.shapes = [Shape(left=1+i,top=1, width=2+(i%3),height=2+(i%4)) for i in range(10)]
        td = TableData.from_list(self.shapes, 4)
        self.table = ShapeTableAlignment(td)

    def test_simple_align(self):
        self.table.spacing = 5
        self.table.align()

        self.assertEqual(self.shapes[2].left, 19)
        self.assertEqual(self.shapes[4].top, 11)
        self.assertEqual(self.shapes[7].left, 25)
        self.assertEqual(self.shapes[9].top, 21)

        self.table.spacing = 2,3
        self.table.align()

        self.assertEqual(self.shapes[2].left, 15)
        self.assertEqual(self.shapes[4].top, 8)
        self.assertEqual(self.shapes[7].left, 19)
        self.assertEqual(self.shapes[9].top, 15)

        self.table.cell_fit = True
        self.table.align()

        self.assertEqual(self.shapes[2].left, 15)
        self.assertEqual(self.shapes[4].top, 8)
        self.assertEqual(self.shapes[7].left, 19)
        self.assertEqual(self.shapes[9].top, 15)

        self.assertEqual(self.shapes[2].width, 4)
        self.assertEqual(self.shapes[4].height, 5)
        self.assertEqual(self.shapes[7].width, 3)
        self.assertEqual(self.shapes[9].height, 3)
    
    def test_align_zero(self):
        self.table.spacing = 0
        self.table.align()

        self.assertEqual(self.shapes[2].left, 9)
        self.assertEqual(self.shapes[4].top, 6)
        self.assertEqual(self.shapes[7].left, 10)
        self.assertEqual(self.shapes[9].top, 11)
    
    def test_getter(self):
        self.assertTupleEqual(self.table.get_bounds(), (1,1,12,5))
        self.assertAlmostEqual(self.table.get_median_spacing(), 0)

        self.table.spacing = 1.5
        self.table.align()

        self.assertTupleEqual(self.table.get_bounds(), (1,1,16.5,16))
        self.assertAlmostEqual(self.table.get_median_spacing(), 2.5)
    
    def test_bounds_align(self):
        self.table.in_bounds = True
        self.table.spacing = 0.5
        self.table.align()

        self.assertAlmostEqual(self.shapes[2].left, 7.6, places=3)
        self.assertAlmostEqual(self.shapes[4].top, 3.0384, places=3)
        self.assertAlmostEqual(self.shapes[7].left, 8.8, places=3)
        self.assertAlmostEqual(self.shapes[9].top, 5.0769, places=3)

        self.assertAlmostEqual(self.shapes[2].width, 2.8, places=3)
        self.assertAlmostEqual(self.shapes[4].height, 0.6153, places=3)
        self.assertAlmostEqual(self.shapes[7].width, 2.1, places=3)
        self.assertAlmostEqual(self.shapes[9].height, 0.923, places=3)

        self.table.cell_fit = True
        self.table.align()

        self.assertAlmostEqual(self.shapes[2].left, 7.6, places=3)
        self.assertAlmostEqual(self.shapes[4].top, 3.0384, places=3)
        self.assertAlmostEqual(self.shapes[7].left, 8.8, places=3)
        self.assertAlmostEqual(self.shapes[9].top, 5.0769, places=3)

        self.assertAlmostEqual(self.shapes[2].width, 2.8, places=3)
        self.assertAlmostEqual(self.shapes[4].height, 1.5384, places=3)
        self.assertAlmostEqual(self.shapes[7].width, 2.1, places=3)
        self.assertAlmostEqual(self.shapes[9].height, 0.9230, places=3)

    def test_rows_cols_only_align(self):
        self.table.spacing = None,3
        self.table.align()

        self.assertEqual(self.shapes[2].left, 15)
        self.assertEqual(self.shapes[4].top, 1)
        self.assertEqual(self.shapes[7].left, 19)
        self.assertEqual(self.shapes[9].top, 1)

        self.table.spacing = 3,None
        self.table.align()

        self.assertEqual(self.shapes[2].left, 15)
        self.assertEqual(self.shapes[4].top, 9)
        self.assertEqual(self.shapes[7].left, 19)
        self.assertEqual(self.shapes[9].top, 17)
    
    def test_equalize(self):
        self.table.equalize_cols = True
        self.table.equalize_rows = True
        self.table.spacing = 0.5
        self.table.align()

        self.assertAlmostEqual(self.shapes[2].left, 10, places=3)
        self.assertAlmostEqual(self.shapes[4].top, 6.5, places=3)
        self.assertAlmostEqual(self.shapes[7].left, 14.5, places=3)
        self.assertAlmostEqual(self.shapes[9].top, 12, places=3)

        self.assertAlmostEqual(self.shapes[2].width, 4, places=3)
        self.assertAlmostEqual(self.shapes[4].height, 2, places=3)
        self.assertAlmostEqual(self.shapes[7].width, 3, places=3)
        self.assertAlmostEqual(self.shapes[9].height, 3, places=3)

        self.table.in_bounds = True
        self.table.spacing = 1
        self.table.align()

        self.assertAlmostEqual(self.shapes[2].left, 11, places=3)
        self.assertAlmostEqual(self.shapes[4].top, 7, places=3)
        self.assertAlmostEqual(self.shapes[7].left, 16, places=3)
        self.assertAlmostEqual(self.shapes[9].top, 13, places=3)

        self.assertAlmostEqual(self.shapes[2].width, 4, places=3)
        self.assertAlmostEqual(self.shapes[4].height, 2, places=3)
        self.assertAlmostEqual(self.shapes[7].width, 3, places=3)
        self.assertAlmostEqual(self.shapes[9].height, 3, places=3)

        self.table.cell_fit = True
        self.table.align()

        self.assertAlmostEqual(self.shapes[2].left, 11, places=3)
        self.assertAlmostEqual(self.shapes[4].top, 7, places=3)
        self.assertAlmostEqual(self.shapes[7].left, 16, places=3)
        self.assertAlmostEqual(self.shapes[9].top, 13, places=3)

        self.assertAlmostEqual(self.shapes[2].width, 3.75, places=3)
        self.assertAlmostEqual(self.shapes[4].height, 4.3333, places=3)
        self.assertAlmostEqual(self.shapes[7].width, 3.75, places=3)
        self.assertAlmostEqual(self.shapes[9].height, 4.3333, places=3)

        # print("\n\n2\t %s" % self.shapes[2].left)
        # print("4\t %s" % self.shapes[4].top)
        # print("7\t %s" % self.shapes[7].left)
        # print("9\t %s\n\n" % self.shapes[9].top)

        # print("\n\n2\t %s" % self.shapes[2].width)
        # print("4\t %s" % self.shapes[4].height)
        # print("7\t %s" % self.shapes[7].width)
        # print("9\t %s\n\n" % self.shapes[9].height)

    def test_cell_alignment(self):
        self.table.cell_alignment_x = "center"
        self.table.cell_alignment_y = "right"
        self.table.align()
        self.fail()

    def test_transpose(self):
        self.fail()


class TableRecognitionTest(unittest.TestCase):
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