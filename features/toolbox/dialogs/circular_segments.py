# -*- coding: utf-8 -*-



import logging

import bkt.ui
notify_property = bkt.ui.notify_property

import bkt.library.powerpoint as pplib
pt_to_cm = pplib.pt_to_cm
cm_to_pt = pplib.cm_to_pt

from ..models.segmentedcircle import SegmentedCircle


# =======================
# = UI MODEL AND WINDOW =
# =======================


class SegmentedCircleViewModel(bkt.ui.ViewModelSingleton):
    
    def __init__(self):
        super(SegmentedCircleViewModel, self).__init__()
        
        self._update_enabled = False
        self._num_segments = 3
        self._radius = 4.0
        self._width = 25
        self._use_arrow_shape = True
        self._spacing = 0

    def set_values_based_on_shape(self, shape):
        if not SegmentedCircle.is_segmented_circle(shape):
            raise ValueError("not a segmented circle")
        self.num_segments, self.width, self._radius, arrow_shape, spacing = SegmentedCircle.determine_from_shape(shape)
        self.radius = pt_to_cm(self._radius)
        if arrow_shape:
            self.use_arrow_shape = True
        else:
            self.use_segment_shape = True
        if spacing >= 3:
            self.spacing_big = True
        elif spacing >= 1:
            self.spacing_small = True
        else:
            self.spacing_none = True
        self.update_enabled = self.num_segments > 1 #only allow update if shape is already a group
    
    
    @notify_property
    def update_enabled(self):
        return self._update_enabled and self._num_segments > 1

    @update_enabled.setter
    def update_enabled(self, value):
        self._update_enabled = value
    
    @notify_property
    def num_segments(self):
        return self._num_segments

    @num_segments.setter
    def num_segments(self, value):
        self._num_segments = value
        self.OnPropertyChanged('update_enabled')
    
    @notify_property
    def radius(self):
        return self._radius

    @radius.setter
    def radius(self, value):
        self._radius = value
    
    @notify_property
    def width(self):
        return self._width

    @width.setter
    def width(self, value):
        self._width = value
    
    ## getters/setters for radio buttons arrows
    
    @notify_property
    def use_arrow_shape(self):
        return self._use_arrow_shape

    @use_arrow_shape.setter
    def use_arrow_shape(self, value):
        self._use_arrow_shape = True
    
    @notify_property
    def use_segment_shape(self):
        return not self._use_arrow_shape

    @use_segment_shape.setter
    def use_segment_shape(self, value):
        self._use_arrow_shape = False
    
    ## getters/setters for radio buttons spacing
    
    @notify_property
    def spacing_none(self):
        return self._spacing == 0

    @spacing_none.setter
    def spacing_none(self, value):
        self._spacing = 0
    
    @notify_property
    def spacing_small(self):
        return self._spacing == 5

    @spacing_small.setter
    def spacing_small(self, value):
        self._spacing = 5
    
    @notify_property
    def spacing_big(self):
        return self._spacing == 10

    @spacing_big.setter
    def spacing_big(self, value):
        self._spacing = 10



class SegmentedCircleWindow(bkt.ui.WpfWindowAbstract):
    _xamlname = 'circular_segments'
    _vm_class = SegmentedCircleViewModel
    
    def __init__(self, context, slide):
        super(SegmentedCircleWindow, self).__init__(context)
        
        self.ref_slide = slide
        self.ref_shape = None

        try:
            self._vm.update_enabled = False
            self._vm.set_values_based_on_shape(context.shape)
            self.ref_shape = context.shape
        except:
            pass #e.g. nothing selected

    def cancel(self, sender, event):
        self.Close()
    
    def update_segments(self, sender, event):
        try:
            if self.ref_shape:
                SegmentedCircle.updated_segmented_circle(self.ref_shape, self._vm.num_segments, self._vm.width, cm_to_pt(self._vm.radius), self._vm.use_arrow_shape, self._vm._spacing)
            else:
                SegmentedCircle.create_segmented_circle(self.ref_slide, self._vm.num_segments, self._vm.width, cm_to_pt(self._vm.radius), self._vm.use_arrow_shape, self._vm._spacing)
        except:
            logging.exception("Dialog action failed")
        finally:
            self.Close()
    
    def create_segments(self, sender, event):
        try:
            SegmentedCircle.create_segmented_circle(self.ref_slide, self._vm.num_segments, self._vm.width, cm_to_pt(self._vm.radius), self._vm.use_arrow_shape, self._vm._spacing)
        except:
            logging.exception("Dialog action failed")
        finally:
            self.Close()