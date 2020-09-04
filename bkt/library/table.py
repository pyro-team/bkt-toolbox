# -*- coding: utf-8 -*-
'''
Created on 11.09.2013

@author: cschmitt
'''

from __future__ import absolute_import, division #always force float-division, for int divison use //

import math
import logging

from .algorithms import median, get_bounds_shapes


class TableData(object):
    "Generic class to store tabular data"
    def __init__(self, rows=[]):
        self._rows = rows
        self._num_cols = None

    @property
    def rows(self):
        "return number of rows in the table"
        return len(self._rows)

    @property
    def columns(self):
        "return number of columns in the table"
        if self._num_cols is None:
            self._num_cols = max(len(row) for row in self._rows)
        return self._num_cols

    @property
    def dimension(self):
        "return the table dimension as tuple with rows, columns"
        return self.rows, self.columns

    def add_rows(self, *args):
        "add row(s) to table by passing each row-list as argument"
        for arg in args:
            self._rows.append(arg)
        self._num_cols = None #force column recalc on use
        # self._num_cols = max(self.columns, len(row))

    def get_row(self, index):
        "return all cell-shapes of a specific row"
        return self._rows[index]
    
    def get_rows(self):
        "return all rows"
        return self._rows

    def get_column(self, index, fill=None):
        "return all cell-shapes of a specific column"
        if index >= self.columns:
            raise IndexError
        for line in self._rows:
            if index < len(line):
                yield line[index]
            else:
                yield fill
    
    def get_columns(self, fill=None):
        "return all columns"
        return (self.get_column(i, fill) for i in range(self.columns))

    def get_cell(self, row, column, fill=None):
        "return a cell-shape by given row and column"
        list_row = self._rows[row]
        try:
            return list_row[column]
        except IndexError:
            if column >= self.columns:
                raise IndexError #reraise error
            else:
                return fill

    def __getitem__(self, key):
        "if key is an index, return according row, if key is a tuple of ("
        if type(key) is tuple:
            row,col = key
            if row is None:
                return self.get_column(col)
            else:
                return self.get_cell(row, col)
        else:
            return self.get_row(key)

    def __iter__(self):
        "iterate over all cell-shapes by returning the tuple (row index, col index, cell shape)"
        rows, cols = self.dimension
        for i in range(rows):
            for j in range(cols):
                cell = self.get_cell(i,j)
                if cell is None:
                    continue
                yield i,j,cell

    def transpose(self):
        "transpose the table cells"
        self._rows = zip(*self._rows)
    
    @classmethod
    def from_list(cls, cells, max_columns=None):
        "create table from list of of cells, optionally divide in rows with max number of given columns"
        from itertools import izip_longest
        if max_columns:
            args = [iter(cells)] * int(max_columns)
            return cls(list(izip_longest(*args)))
        else:
            return cls(cells)


class ShapeTableAlignment(object):
    def __init__(self, table):
        self._table = table

        self.cell_alignment_x = "left" #center, right
        self.cell_alignment_y = "top" #middle, bottom
        self.cell_fit = False #set shape size to cell size

        self.spacing_rows = None #spacing between rows
        self.spacing_cols = None #spacing between cols

        self.equalize_rows = False #equalize row heights
        self.equalize_cols = False #equalize col widths

        self.bounds = None #set bounds to keep table within current size

    @property
    def spacing(self):
        return self.spacing_rows, self.spacing_cols
    @spacing.setter
    def spacing(self, value):
        if type(value) is tuple:
            self.spacing_rows, self.spacing_cols = value
        else:
            self.spacing_rows = self.spacing_cols = value

    @property
    def in_bounds(self):
        return self.bounds is not None
    @in_bounds.setter
    def in_bounds(self, value):
        if value:
            self.bounds = self.get_bounds()
        else:
            self.bounds = None

    def _column_left(self,col):
        "return the left-most coordinate of all cell-shapes of the given column"
        return min(s.left for s in self._table.get_column(col) if s is not None)
    
    def _column_width(self,col):
        "return the maximum width of all cell-shapes of the given column"
        return max(s.width for s in self._table.get_column(col) if s is not None)
    
    def _row_top(self,row):
        "return the top-most coordinate of all cell-shapes of the given row"
        return min(s.top for s in self._table.get_row(row) if s is not None)
    
    def _row_height(self,row):
        "return the maximum height of all cell-shapes of the given row"
        return max(s.height for s in self._table.get_row(row) if s is not None)


    def _set_cell_size(self, cell, width=None, height=None):
        "set cell size to given width and/or height"
        #FIXME: this is probably powerpoint specific
        if cell.HasTextFrame == -1 and cell.TextFrame.AutoSize == 1:
            cell.TextFrame.AutoSize = 0

        if not cell.LockAspectRatio or width is None or height is None:
            if width is not None:
                cell.width = width
            if height is not None:
                cell.height = height
        else:
            ratio = float(cell.width)/float(cell.height)
            if ratio > 1:
                cell.width = width
            else:
                cell.height = height

    def set_bounds(self, left,right,width,height):
        self.bounds = left,right,width,height

    def get_bounds(self):
        "return the bounds of the whole table as tuple with left, top, width, height"
        return get_bounds_shapes(c[2] for c in self._table)

    def get_median_spacing(self):
        "return the median spacing between all rows and columns together"
        def iterate_spacings():
            for i,j,cell in self._table:
                if cell is None:
                    continue
                if j > 0:
                    left = self._table.get_cell(i, j-1)
                    if left is not None:
                        spacing = cell.left - left.left - left.width
                        if spacing >= 0:
                            yield spacing
                if i > 0:
                    top = self._table.get_cell(i-1, j)
                    if top is not None:
                        spacing = cell.top - top.top - top.height
                        if spacing >= 0:
                            yield spacing
                            
        spacings = list(iterate_spacings())
        if not spacings:
            return 0
        return median(spacings)


    def align(self, xstart=None, ystart=None):
        """
        Align table shape with given spacing. If spacing is a tuple, first value define row-spacing, second column-spacing.
        The whole table can be moved with given xstart and ystart coordinates.
        If fit_cells is set to True, the cell-shapes will fill the whole cell size.
        Alignment within cells is specified with align_x and align_y, default is top-left.
        """

        if self.in_bounds:
            self._fit_content()
            xstart, ystart, _, _ = self.bounds

        #set x-start coordinate
        if xstart is None:
            x = self._column_left(0)
        else:
            x = xstart

        #set y-start coordinate
        if ystart is None:
            # y = self.first_top.Top
            y = self._row_top(0)
        else:
            y = ystart

        #get column left coordinates and widths for each column
        widths = []
        lefts = []
        for col in range(self._table.columns):
            col_width = self._column_width(col)
            widths.append(col_width)
            # if self.spacing_cols is not None:
            #     lefts.append(x)
            #     x += col_width + self.spacing_cols
        
        if self.equalize_cols:
            new_width = max(widths)
            widths = [new_width] * len(widths)
        
        if self.spacing_cols is not None:
            lefts = [x + (col_width+self.spacing_cols)*i for i,col_width in enumerate(widths)]
        
        heights = [self._row_height(row) for row in range(self._table.rows)]
        if self.equalize_rows:
            new_height = max(heights)
            heights = [new_height] * len(heights)
        
        #iterate lines
        for row, line in enumerate(self._table.get_rows()):
            #get height of current line
            # height = self._row_height(row)
            
            for j, shape in enumerate(line):
                if shape is None:
                    continue
                
                #fill the whole cell
                if self.cell_fit:
                    shape.width = widths[j]
                    shape.height = heights[row]

                #vertical alignment
                if self.spacing_rows is not None:
                    y_pos = y
                    y_pos += (heights[row] - shape.height) if self.cell_alignment_y == "bottom" else 0
                    y_pos += (heights[row] - shape.height)/2 if self.cell_alignment_y == "middle" else 0
                    shape.top = y_pos
                
                #horizontal alignment
                if self.spacing_cols is not None:
                    x_pos = lefts[j]
                    x_pos += (widths[j] - shape.width) if self.cell_alignment_x == "right" else 0
                    x_pos += (widths[j] - shape.width)/2 if self.cell_alignment_x == "center" else 0
                    shape.left = x_pos
            
            if self.spacing_rows is not None:
                y += heights[row] + self.spacing_rows

    def _fit_content(self):
        """
        Align the table within the given bounds and with given spacing. If spacing is a tuple, first value define row-spacing, second column-spacing.
        If fit_cells is set to True, the cell-shapes will fill the whole cell size.
        If distribute_cols or distribute_rows is given, the col/row size is equalized within given bounds.
        """

        assert self.bounds is not None

        _,_,width,height = self.bounds
        rows, cols = self._table.dimension

        widths  = []
        heights = []
        scale_x = 1
        scale_y = 1

        if self.spacing_cols is not None:
            widths = [self._column_width(col) for col in range(cols)]
            if self.equalize_cols:
                widths = [float(sum(widths)) / len(widths)] * len(widths)
            else:
                remaining_width = width - (len(widths)-1)*self.spacing_cols
                scale_x = float(remaining_width) / float(sum(widths))
        
        if self.spacing_rows is not None:
            heights = [self._row_height(row) for row in range(rows)]
            if self.equalize_rows:
                heights = [float(sum(heights)) / len(heights)] * len(heights)
            else:
                remaining_height = height - (len(heights)-1)*self.spacing_rows
                scale_y = float(remaining_height) / float(sum(heights))

        for i,j,cell in self._table:
            width = None
            height = None

            if self.spacing_cols is not None:
                width = widths[j]*scale_x if self.cell_fit else cell.Width*scale_x
            if self.spacing_rows is not None:
                height = heights[i]*scale_y if self.cell_fit else cell.Height*scale_y

            self._set_cell_size(cell, width, height)



class ShapeTableRecognition(object):
    def __init__(self, shapes):
        self._shapes = set(shapes)
    #run, get spacing




class TableRecognition(object):
    "Recognize table of shapes and manipulate all shapes at once, e.g. increase spacing between rows and columns."

    def __init__(self,shapes):
        self.shapes = set(shapes) #TODO: better use heapq?
        self.table = None
        
    def run(self):
        "run the table recognation algorithm and save result internally. does not change any given shapes."
        table = []
        while self.shapes:
            seed = self._collect_seed()
            line = [seed]
            while self.shapes:
                nxt = self._next_in_row(line)
                if nxt is None:
                    break
                self.shapes.discard(nxt)
                line.append(nxt)
            table.append(line)
        self.table = table
        
        self.table = sorted(table, key=lambda l : min(s.Top for s in l))
        self._correct_columns()
        
    def _left_edge(self,col):
        "return median of all left-coordinates of given column"
        return median(s.Left for s in self.column(col) if s is not None)
    
    def _get_column_edges(self):
        num_cols = self.column_count()
        #median_edges = 
        edges = set(self._left_edge(col) for col in range(num_cols))
        for col in range(num_cols):
            #median_edge = self.left_edge(col)
            #edges.add(median_edge)
            for cell in self.column(col):
                if cell is None:
                    continue
                new_edge = True
                for edge in edges:
                    if abs(edge-cell.Left) <= cell.Width * 0.5:
                        new_edge = False
                        break
                        #MessageBox.Show("adding new edge %r:\r\nmedian_edge=%r, \r\ndist=%r, \r\nwidth/2=%r" % (cell.Left, edge, abs(edge-cell.Left), cell.Width*0.5))
                        #edges.add(cell.Left)
                if new_edge:
                    edges.add(cell.Left)
        
        return list(enumerate(sorted(edges)))

    @property
    def dimension(self):
        "return the table dimension as tuple with rows, columns"
        return len(self.table), self.column_count()

    def _cell(self,i,j):
        "return a cell-shape by given row i and column j"
        row = self.table[i]
        if j < len(row):
            return row[j]
    
    def _iter_cells(self):
        "iterate over all cell-shapes by returning the tuple (row index, col index, cell shape)"
        rows, cols = self.dimension
        for i in range(rows):
            for j in range(cols):
                cell = self._cell(i,j)
                if cell is None:
                    continue
                yield i,j,cell
     
    def get_bounds(self):
        "return the bounds of the whole table as tuple with left, top, width, height"
        def iter_points():
            for _, _, cell in self._iter_cells():
                x0 = cell.Left
                y0 = cell.Top
                yield (x0,y0)
                x1 = x0 + cell.Width
                y1 = y0 + cell.Height
                yield (x1,y1)
        
        points = list(iter_points())
        x = [p[0] for p in points]
        y = [p[1] for p in points]
        
        left = min(x)
        top = min(y)
        width = max(x)-left
        height = max(y)-top
        
        return left,top,width,height
    
    def change_spacing_in_bounds(self,delta):
        "change the spacing of the table by given delta, but do not change the overall table size"
        spacing_old = self.median_spacing()
        self.align(spacing_old)
        spacing = max(0,spacing_old+delta)
        bounds = self.get_bounds()
        self.fit_content(*bounds, spacing=spacing)
    
    def median_spacing(self):
        "return the median spacing between all rows and columns together"
        def iterate_spacings():
            rows, cols = self.dimension
            for i in range(rows):
                for j in range(cols):
                    cell = self._cell(i,j)
                    if cell is None:
                        continue
                    if j > 0:
                        left = self._cell(i, j-1)
                        if left is not None:
                            spacing = cell.Left - left.Left - left.Width
                            if spacing > 0:
                                yield spacing
                    if i > 0:        
                        top = self._cell(i-1, j)
                        if top is not None:
                            spacing = cell.Top - top.Top - top.Height
                            if spacing > 0:
                                yield spacing
                            
        spacings = list(iterate_spacings())
        #MessageBox.Show(str(spacings))
        if not spacings:
            return 0
        return median(spacings)
    
    def min_spacing_rows(self, max_rows=None):
        "return the minimum spacing between all rows, or max_rows if given"
        rows, cols = self.dimension
        if rows < 2:
            return 0
        cur_min = float('inf')
        for i in range(1,max_rows or rows): #for speed improvement, we can only iterate first 2 rows
            row1_bottom = []
            row2_top    = []
            for j in range(cols):
                row1_cell = self._cell(i-1,j)
                if row1_cell is not None:
                    row1_bottom.append(row1_cell.Top+row1_cell.Height)
                row2_cell = self._cell(i,j)
                if row2_cell is not None:
                    row2_top.append(row2_cell.Top)
            cur_min = min(cur_min, min(row2_top) - max(row1_bottom))
        return cur_min
    
    def min_spacing_cols(self, max_cols=None):
        "return the minimum spacing between all columns, or max_cols if given"
        rows, cols = self.dimension
        if cols < 2:
            return 0
        cur_min = float('inf')
        for j in range(1,max_cols or cols): #for speed improvement, we can only iterate first 2 cols
            col1_right = []
            col2_left  = []
            for i in range(rows):
                col1_cell = self._cell(i,j-1)
                if col1_cell is not None:
                    col1_right.append(col1_cell.Left+col1_cell.Width)
                col2_cell = self._cell(i,j)
                if col2_cell is not None:
                    col2_left.append(col2_cell.Left)
            cur_min = min(cur_min, min(col2_left) - max(col1_right))
        return cur_min
    
    def _correct_columns(self):
        edges = self._get_column_edges()
        #MessageBox.Show(str(edges))
        num_cols = len(edges)
        
        
        def get_index(cell):
            left = cell.Left
            res = min(edges, key=lambda t : abs(left-t[1]))
            #MessageBox.Show(str(res))
            return res[0]
        
        table = []
        for i,line in enumerate(self.table):
            line_new = [None]*num_cols
            for cell in line:
                if cell is None:
                    continue
                index = get_index(cell)
                if line_new[index] is not None:
                    # logging.debug("cell index %d is duplicated in line %d\r\nedges: %r", index,i,edges)
                    line_new = list(line)
                    while len(line_new) < num_cols:
                        line_new.append(None)
                    break
                line_new[index] = cell
            table.append(line_new)
            
        col = 0
        while col < num_cols:
            cellset = set(line[col] for line in table)
            if cellset == set([None]):
                for line in table:
                    del line[col]
                num_cols -= 1
            else:
                col += 1
        
        self.table = table
                        
    
    def column_count(self):
        "return number of columns in the table"
        return max(len(line) for line in self.table)
            
    def column(self,col):
        "return all cell-shapes of a specific column"
        for line in self.table:
            if col < len(line):
                yield line[col]
            else:
                yield None
    
    def transpose(self):
        "transpose the table cells (internally). align needs to be called afterwards."
        cols = self.column_count()
        table = []
        for j in range(cols):
            table.append(list(self.column(j)))
        self.table = table
        
    def transpose_cell_size(self):
        "transpose the cell sizes of the table"
        for line in self.table:
            for cell in line:
                if cell is None:
                    continue
                cell.Width, cell.Height = cell.Height, cell.Width
    
    @property
    def first_top(self):
        "return the first cell-shape in the first row"
        for cell in self.table[0]:
            if cell is not None:
                return cell
    
    @property
    def first_left(self):
        "return the first cell-shape in the first column"
        for cell in self.column(0):
            if cell is not None:
                return cell
    
    def _column_left(self,col):
        "return the left-most coordinate of all cell-shapes of the given column"
        shapes = [s for s in self.column(col) if s is not None]
        if not shapes:
            return 0
        return min(s.Left for s in shapes)
    
    def _column_width(self,col):
        "return the maximum width of all cell-shapes of the given column"
        #FIXME: shouldnt this be: max(left+width) - col_left?
        shapes = [s for s in self.column(col) if s is not None]
        if not shapes:
            return 0
        return max(s.Width for s in shapes)
    
    def _row_top(self,row):
        "return the top-most coordinate of all cell-shapes of the given row"
        shapes = [s for s in self.table[row] if s is not None]
        if not shapes:
            return 0
        return min(s.Top for s in shapes)
    
    def _row_height(self,row):
        "return the maximum height of all cell-shapes of the given row"
        #FIXME: shouldnt this be: max(top+height) - row_top?
        shapes = [s for s in self.table[row] if s is not None]
        if not shapes:
            return 0
        return max(s.Height for s in shapes)
    
    def fit_content(self,left,top,width,height,spacing,fit_cells=False, distribute_cols=False, distribute_rows=False):
        """
        Align the table within the given bounds and with given spacing. If spacing is a tuple, first value define row-spacing, second column-spacing.
        If fit_cells is set to True, the cell-shapes will fill the whole cell size.
        If distribute_cols or distribute_rows is given, the col/row size is equalized within given bounds.
        """
        rows, cols = self.dimension

        #tuple = (row spacing, column spacing)
        if type(spacing) == tuple:
            spacing_rows, spacing_cols = spacing
        else:
            spacing_rows = spacing
            spacing_cols = spacing

        if spacing_cols is not None:
            widths = [self._column_width(col) for col in range(cols)]
            if distribute_cols:
                widths = [float(sum(widths)) / len(widths)] * len(widths)
                scale_x = 1
            else:
                remaining_width = width - (len(widths)-1)*spacing_cols
                scale_x = float(remaining_width) / float(sum(widths))
        
        if spacing_rows is not None:
            heights = [self._row_height(row) for row in range(rows)]
            if distribute_rows:
                heights = [float(sum(heights)) / len(heights)] * len(heights)
                scale_y = 1
            else:
                remaining_height = height - (len(heights)-1)*spacing_rows
                scale_y = float(remaining_height) / float(sum(heights))
        
        def set_size(cell, width=None, height=None):
            if not cell.LockAspectRatio or width is None or height is None:
                if width is not None:
                    cell.Width = width
                if height is not None:
                    cell.Height = height
            else:
                ratio = float(cell.Width)/float(cell.Height)
                if ratio > 1:
                    cell.Width = width
                else:
                    cell.Height = height


        for i in range(rows):
            for j in range(cols):
                cell = self._cell(i, j)
                if cell is None:
                    continue
                
                #FIXME: this is probably powerpoint specific
                if cell.HasTextFrame == -1 and cell.TextFrame.AutoSize == 1:
                    cell.TextFrame.AutoSize = 0

                width = None
                height = None

                if spacing_cols is not None:
                    width = widths[j]*scale_x if fit_cells else cell.Width*scale_x
                if spacing_rows is not None:
                    height = heights[i]*scale_y if fit_cells else cell.Height*scale_y

                set_size(cell, width, height)
        
        self.align(spacing, left, top)
            
    def align(self, spacing=10, xstart=None, ystart=None, fit_cells=False, align_x="left", align_y="top"):
        """
        Align table shape with given spacing. If spacing is a tuple, first value define row-spacing, second column-spacing.
        The whole table can be moved with given xstart and ystart coordinates.
        If fit_cells is set to True, the cell-shapes will fill the whole cell size.
        Alignment within cells is specified with align_x and align_y, default is top-left.
        """
        num_columns = self.column_count()
        widths = [self._column_width(col) for col in range(num_columns)]
        
        #tuple = (row spacing, column spacing)
        if type(spacing) == tuple:
            spacing_rows, spacing_cols = spacing
        else:
            spacing_rows = spacing
            spacing_cols = spacing

        #set x-start coordinate
        left = []
        if xstart is None:
            # x = self.first_left.Left
            x = self._column_left(0)
        else:
            x = xstart

        #set y-start coordinate
        if ystart is None:
            # y = self.first_top.Top
            y = self._row_top(0)
        else:
            y = ystart

        #calculate x coordinates (columns)
        if spacing_cols is not None:
            for w in widths:
                left.append(x)
                x += w + spacing_cols
        
        #iterate lines
        for line in self.table:
            height = max(s.Height for s in line if s is not None)
            
            for j, shape in enumerate(line):
                if shape is None:
                    continue
                
                if fit_cells:
                    shape.Width = widths[j]
                    shape.Height = height

                if spacing_rows is not None:
                    y_pos = y
                    y_pos += (height - shape.Height) if align_y == "bottom" else 0
                    y_pos += (height - shape.Height)/2 if align_y == "middle" else 0
                    shape.Top = y_pos
                
                if spacing_cols is not None:
                    x_pos = left[j]
                    x_pos += (widths[j] - shape.Width) if align_x == "right" else 0
                    x_pos += (widths[j] - shape.Width)/2 if align_x == "center" else 0
                    shape.Left = x_pos
            
            if spacing_rows is not None:
                y += height + spacing_rows
        
    def _next_in_row(self, line):
        "return next suitable shape in row by given line, or None if no shape in line is found"
        ref = line[0]
        refx = ref.Left
        refy = ref.Top
        bound = ref.Height * 0.5
        
        selection = []
        for s in self.shapes:
            if s.Left < refx:
                continue
            if abs(s.Top-refy) > bound:
                continue
            selection.append(s)
            
        if selection:
            return min(selection, key=lambda s : abs(s.Left-refx))
        else:
            return None
        
#        refx = shape.Left + shape.Width
#        refy = shape.top
#        
#        selection = []
#        ref = shape.Left
#        bound = shape.
#        for s in self.shapes:
#            if s is shape:
#                continue
#            m
#        
#        return min(self.shapes, key=self.distfun(refx, refy))

    # def _collect_next_in_row(self, shape):
    #     nxt = self.next_in_row(shape)
    #     self.shapes.discard(nxt)
    #     return nxt
    
    def distfun(self,refx,refy):
        "return the distance function for a shape used as key function"
        def dist(s):
            return math.hypot(refx-s.Left, refy-s.Top)
        return dist
    
    def _collect_seed(self):
        "get the first shape of a new row, which is the shape with the least distance from the top-left-most edge of all remaining shapes"
        leftmost = min(s.Left for s in self.shapes)
        topmost =  min(s.Top for s in self.shapes)
        
        seed = min(self.shapes,key=self.distfun(leftmost, topmost))
        self.shapes.discard(seed)
        return seed
