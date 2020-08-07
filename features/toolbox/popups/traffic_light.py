# -*- coding: utf-8 -*-

from __future__ import absolute_import

import logging

import bkt
import bkt.library.powerpoint as pplib

from bkt.callbacks import WpfActionCallback



# ================
# = dialog logic =
# ================

class Ampel(object):
    #BKT_CONTEXTDIALOG = 'BKT_CONTEXTDIALOG'
    BKT_DIALOG_AMPEL = 'BKT_DIALOG_AMPEL3'

    color_states = ['red', 'yellow', 'green']
    color_rgb    = [255, 65535, 5287936]
    color_white  = 16777215

    @classmethod
    def create(cls, slide, style="vertical", border=True):
        logging.debug("create ampel 3")

        if style == "simple":
            shapes = [
                slide.shapes.addshape(9, 100, 100, 20, 20), #circle
            ]
        elif style == "horizontal":
            shapes = [
                slide.shapes.addshape(1, 100, 100, 80, 30), #rect: left, top, width, height
                slide.shapes.addshape(9, 105, 105, 20, 20), #red
                slide.shapes.addshape(9, 130, 105, 20, 20), #yellow
                slide.shapes.addshape(9, 155, 105, 20, 20)  #green
            ]
        else:
            shapes = [
                slide.shapes.addshape(1, 100, 100, 30, 80), #rect: left, top, width, height
                slide.shapes.addshape(9, 105, 105, 20, 20), #red
                slide.shapes.addshape(9, 105, 130, 20, 20), #yellow
                slide.shapes.addshape(9, 105, 155, 20, 20)  #green
            ]
        for shape in shapes:
            shape.fill.ForeColor.RGB = 14277081
            if border:
                shape.line.weight = 0.75
                shape.line.ForeColor.RGB = 0
            else:
                shape.line.visible = 0
        
        # gruppieren
        if len(shapes) == 1:
            grp = shapes[0]
        else:
            grp = pplib.last_n_shapes_on_slide(slide, len(shapes)).group()
        
        grp.LockAspectRatio = -1 #msoTrue
        grp.Tags.Add(bkt.contextdialogs.BKT_CONTEXTDIALOG_TAGKEY, cls.BKT_DIALOG_AMPEL)
        grp.Name = "[BKT] Traffic Light %s" % grp.id

        grp.select()
        cls.set_color(grp)
        
    
    @classmethod
    def set_color(cls, shape, color="red"):
        try:
            index = cls.color_states.index(color)
        except ValueError:
            index = -1
        # shape to change depending on traffic light type
        if shape.Type != pplib.MsoShapeType['msoGroup']:
            shape_to_change = shape
            shape_to_change.fill.ForeColor.RGB = cls.color_white # white
        else:
            colors = [shp for shp in shape.GroupItems if shp.AutoShapeType == 9]
            colors.sort(key=lambda shp: (shp.Top, shp.Left))
            colors[0].fill.ForeColor.RGB = cls.color_white # white
            colors[1].fill.ForeColor.RGB = cls.color_white # white
            colors[2].fill.ForeColor.RGB = cls.color_white # white
            shape_to_change = colors[index]
        # if index is -1 nothing to change (light stays white)
        if index >= 0:
            shape_to_change.fill.ForeColor.RGB = cls.color_rgb[index]
        
        
    @classmethod
    def get_color(cls, shape):
        if shape.Type != pplib.MsoShapeType['msoGroup']:
            try:
                return cls.color_states[cls.color_rgb.index(shape.fill.ForeColor.RGB)]
            except ValueError:
                return "green"
        else:
            colors = [shp for shp in shape.GroupItems if shp.AutoShapeType == pplib.MsoAutoShapeType['msoShapeOval']]
            colors.sort(key=lambda shp: (shp.Top, shp.Left))
            if colors[0].fill.ForeColor.RGB == cls.color_rgb[0]:
                return "red"
            elif colors[1].fill.ForeColor.RGB == cls.color_rgb[1]:
                return "yellow"
            else:
                return "green"


    @classmethod
    def next_color(cls, shape):
        current_color = cls.get_color(shape)
        next_color_index = (cls.color_states.index(current_color)+1) % len(cls.color_states)
        cls.set_color(shape, cls.color_states[next_color_index])



# ==========
# = window =
# ==========

# class TrafficPopup(bkt.ui.WpfPopupAbstract):
class TrafficPopup(bkt.ui.WpfWindowAbstract):
    # _filename = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'traffic_light_dialog2.xaml')
    _xamlname = 'traffic_light_dialog2'
    '''
    class representing a popup-dialog for a traffic-light-shape
    '''
    
    def __init__(self, context=None):
        self.IsPopup = True

        super(TrafficPopup, self).__init__(context)

    # def __init__(self, context=None):
    #     filename=os.path.join(os.path.dirname(os.path.realpath(__file__)), 'traffic_light_dialog.xaml')
    #     wpf.LoadComponent(self, filename)
    #     self._vm = ViewModel()
    #     self._context = context
    #     self.DataContext = self._vm

    @WpfActionCallback
    def btnred(self, sender, event):
        try:
            shapes = list(iter(self._context.selection.shaperange))
            for shape in shapes:
                Ampel.set_color(shape, "red")
            # self._context.app.ActiveWindow.Activate()
        except:
            logging.exception("traffic light exception")

    @WpfActionCallback
    def btnyellow(self, sender, event):
        try:
            shapes = list(iter(self._context.selection.shaperange))
            for shape in shapes:
                Ampel.set_color(shape, "yellow")
            # self._context.app.ActiveWindow.Activate()
        except:
            logging.exception("traffic light exception")

    @WpfActionCallback
    def btngreen(self, sender, event):
        try:
            shapes = list(iter(self._context.selection.shaperange))
            for shape in shapes:
                Ampel.set_color(shape, "green")
            # self._context.app.ActiveWindow.Activate()
        except:
            logging.exception("traffic light exception")

    @WpfActionCallback
    def btnwhite(self, sender, event):
        try:
            shapes = list(iter(self._context.selection.shaperange))
            for shape in shapes:
                Ampel.set_color(shape, "white")
            # self._context.app.ActiveWindow.Activate()
        except:
            logging.exception("traffic light exception")


#initialization function called by contextdialogs.py
create_window = TrafficPopup

def trigger_doubleclick(shape, context):
    try:
        Ampel.next_color(shape)
    except:
        logging.exception("traffic light exception")