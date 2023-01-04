# -*- coding: utf-8 -*-



import logging

import bkt.ui
notify_property = bkt.ui.notify_property

from bkt.library.powerpoint import pt_to_cm, cm_to_pt


# =================
# = FUNCTIONALITY =
# =================


class SplitShapes(object):
    default_row_sep = cm_to_pt(0.2)
    default_col_sep = cm_to_pt(0.2)
    default_rows = 6
    default_cols = 6
    
    @classmethod
    def split_shapes(cls, shapes, rows, cols, row_sep, col_sep):
        for shape in shapes:
            cls.split_shape(shape, rows, cols, row_sep, col_sep)
    
    @classmethod
    def split_shape(cls, shape, rows, cols, row_sep, col_sep):
        shape_width = shape.width
        shape_height = shape.height
        if cols > 1:
            shape_width = (shape.width - (cols-1)*col_sep)/cols
        if rows > 1:
            shape_height = (shape.height - (rows-1)*row_sep)/rows
        
        #shape.width = shape_width
        #shape.height = shape_height
        
        last_lock_aspect_ratio = shape.LockAspectRatio
        shape.LockAspectRatio = 0
        
        for row_idx in range(rows):
            for col_idx in range(cols):
                if row_idx == 0 and col_idx == 0:
                    shape_copy = shape
                else:
                    shape_copy = shape.duplicate()
                shape_copy.left = shape.left + col_idx*(shape_width+col_sep)
                shape_copy.top = shape.top + row_idx*(shape_height+row_sep)
                shape_copy.width = shape_width
                shape_copy.height = shape_height
                shape_copy.LockAspectRatio = last_lock_aspect_ratio
                shape_copy.select(False)
        #shape.Delete()




class MultiplyShapes(object):
    
    @classmethod
    def multiply_shapes(cls, shapes, rows, cols, row_sep, col_sep):
        for shape in shapes:
            cls.multiply_shape(shape, rows, cols, row_sep, col_sep)
    
    @classmethod
    def multiply_shape(cls, shape, rows, cols, row_sep, col_sep):
        shape_width = shape.width
        shape_height = shape.height
        
        for row_idx in range(rows):
            for col_idx in range(cols):
                if row_idx == 0 and col_idx == 0:
                    continue
                shape_copy = shape.duplicate()
                shape_copy.left = shape.left + col_idx*(shape_width+col_sep)
                shape_copy.top = shape.top + row_idx*(shape_height+row_sep)
                shape_copy.width = shape_width
                shape_copy.height = shape_height
                shape_copy.select(False)



# =======================
# = UI MODEL AND WINDOW =
# =======================


class ViewModel(bkt.ui.ViewModelSingleton):
    
    def __init__(self):
        super(ViewModel, self).__init__()
        
        self.rows = 2
        self.columns = 2
        self.rowsep = 0.2
        self.columnsep = 0.2
        self.row_col_sep_equal = True
        # self.selected_method_is_split = True
    
    
    @notify_property
    def rowsep(self):
        return self._rowsep

    @rowsep.setter
    def rowsep(self, value):
        self._rowsep = value
        if self.row_col_sep_equal:
            self._columnsep = value
            self.OnPropertyChanged('columnsep')

    @notify_property
    def columnsep(self):
        return self._columnsep

    @columnsep.setter
    def columnsep(self, value):
        self._columnsep = value
        if self.row_col_sep_equal:
            self._rowsep = value
            self.OnPropertyChanged('rowsep')

    @notify_property
    def row_col_sep_equal(self):
        return self._row_col_sep_equal

    @row_col_sep_equal.setter
    def row_col_sep_equal(self, value):
        self._row_col_sep_equal = value
        if value==True:
            self._columnsep=self._rowsep
            self.OnPropertyChanged('columnsep')
        self.update_toggle_link_sep_image()


    # @notify_property
    # def selected_method_is_split(self):
    #     return self._selected_method_is_split
    
    # @selected_method_is_split.setter
    # def selected_method_is_split(self, value):
    #     self._selected_method_is_split = value
    #     self.OnPropertyChanged('method_split')
    #     self.OnPropertyChanged('method_multiply')

    
    ## getters/setters for radio buttons
    
    # @notify_property
    # def method_split(self):
    #     return self._selected_method_is_split == True
    
    # @method_split.setter
    # def method_split(self, value):
    #     # always called with value=True
    #     self.selected_method_is_split = True
        
    # @notify_property
    # def method_multiply(self):
    #     return self._selected_method_is_split == False

    # @method_multiply.setter
    # def method_multiply(self, value):
    #     # always called with value=True
    #     self.selected_method_is_split = False

    
    def update_toggle_link_sep_image(self):
        if self.row_col_sep_equal:
            self.toggle_link_sep_image = bkt.ui.load_bitmapimage("TextBoxLinkCreate")
        else:
            self.toggle_link_sep_image = bkt.ui.load_bitmapimage("TextBoxLinkBreak")
    
    
    @notify_property
    def toggle_link_sep_image(self):
        return self._toggle_link_sep_image
        
    @toggle_link_sep_image.setter
    def toggle_link_sep_image(self, value):
        self._toggle_link_sep_image = value
        


#class ShapeSplitWindow(MetroWindow):
class ShapeSplitWindow(bkt.ui.WpfWindowAbstract):
    _xamlname = 'shape_split'
    _vm_class = ViewModel
    
    def __init__(self, context, shapes):
        self.ref_shapes = shapes
        super(ShapeSplitWindow, self).__init__(context)

    def cancel(self, sender, event):
        self.Close()
    
    # def split_multiply_shapes(self, sender, event):
    #     if self._vm.method_split:
    #         method = SplitShapes.split_shapes
    #     else:
    #         method = MultiplyShapes.multiply_shapes
    #     method(self.ref_shapes, self._vm.rows, self._vm.columns, cm_to_pt(self._vm.rowsep), cm_to_pt(self._vm.columnsep))
    #     self.Close()
    
    def multiply_shapes(self, sender, event):
        MultiplyShapes.multiply_shapes(self.ref_shapes, self._vm.rows, self._vm.columns, cm_to_pt(self._vm.rowsep), cm_to_pt(self._vm.columnsep))
        self.Close()
    
    def split_shapes(self, sender, event):
        SplitShapes.split_shapes(self.ref_shapes, self._vm.rows, self._vm.columns, cm_to_pt(self._vm.rowsep), cm_to_pt(self._vm.columnsep))
        self.Close()
