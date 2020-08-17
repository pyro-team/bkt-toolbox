# -*- coding: utf-8 -*-

from __future__ import absolute_import

import logging

from colorsys import rgb_to_hls

import bkt
import bkt.library.powerpoint as pplib

from bkt.library.algorithms import get_rgb_from_ole
from bkt.callbacks import WpfActionCallback



# ================
# = dialog logic =
# ================

class AmpelTags(pplib.BKTTag):
    TAG_NAME = "BKT_AMPEL3"

class Ampel(object):
    #BKT_CONTEXTDIALOG = 'BKT_CONTEXTDIALOG'
    BKT_DIALOG_AMPEL = 'BKT_DIALOG_AMPEL3'

    color_states = ['red', 'yellow', 'green', 'white']
    color_rgb    = [255, 65535, 5287936, 16777215]
    color_box    = 14277081

    @classmethod
    def get_color_white(cls, shape):
        return cls.get_color_rgb(shape)[-1]

    @classmethod
    def get_color_rgb(cls, shape):
        with AmpelTags(shape.Tags) as tags:
            try:
                return tags["color_rgb"]
            except KeyError:
                return cls.color_rgb

    @classmethod
    def set_color_rgb(cls, shape, color):
        with AmpelTags(shape.Tags) as tags:
            tags["color_rgb"] = color


    @classmethod
    def create(cls, slide, style="vertical", border=True):
        logging.debug("traffic light: create ampel 3")

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
            shapes[0].fill.ForeColor.RGB = cls.color_box
        else:
            shapes = [
                slide.shapes.addshape(1, 100, 100, 30, 80), #rect: left, top, width, height
                slide.shapes.addshape(9, 105, 105, 20, 20), #red
                slide.shapes.addshape(9, 105, 130, 20, 20), #yellow
                slide.shapes.addshape(9, 105, 155, 20, 20)  #green
            ]
            shapes[0].fill.ForeColor.RGB = cls.color_box
        
        color_white = cls.color_rgb[-1]
        for shape in shapes:
            if shape.AutoShapeType == pplib.MsoAutoShapeType['msoShapeOval']:
                shape.fill.ForeColor.RGB = color_white
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
        cls.save_color_to_shape(shape)
        try:
            index = cls.color_states.index(color)
        except ValueError:
            logging.debug("traffic light: unknown index")
            index = -1
        # shape to change depending on traffic light type
        if shape.Type != pplib.MsoShapeType['msoGroup']:
            shape_to_change = shape
            shape_to_change.fill.ForeColor.RGB = cls.get_color_white(shape) # white
        else:
            colors = sorted((shp for shp in shape.GroupItems if shp.AutoShapeType == pplib.MsoAutoShapeType['msoShapeOval']), key=lambda shp: (shp.Top, shp.Left))
            color_white = cls.get_color_white(shape)
            colors[0].fill.ForeColor.RGB = color_white # white
            colors[1].fill.ForeColor.RGB = color_white # white
            colors[2].fill.ForeColor.RGB = color_white # white
            try:
                shape_to_change = colors[index]
            except IndexError:
                logging.debug("traffic light: index not in range")
                return #white
        # if index is -1 nothing to change (light stays white)
        if index >= 0:
            shape_to_change.fill.ForeColor.RGB = cls.get_color_rgb(shape)[index]


    @classmethod
    def save_color_to_shape(cls, shape):
        current_color = cls.get_color_state(shape)
        logging.debug("traffic light: current color state %s", current_color)
        try:
            index = cls.color_states.index(current_color)
        except ValueError:
            logging.debug("traffic light: unknown current color")
            return

        if shape.Type != pplib.MsoShapeType['msoGroup']:
            color = shape.fill.ForeColor.RGB
        else:
            colors = sorted((shp for shp in shape.GroupItems if shp.AutoShapeType == pplib.MsoAutoShapeType['msoShapeOval']), key=lambda shp: (shp.Top, shp.Left))
            try:
                color = colors[index].fill.ForeColor.RGB
            except IndexError:
                logging.debug("traffic light: color white not saved")
                return #white

        color_rgb = cls.get_color_rgb(shape)
        if color != color_rgb[index]:
            new_color_rgb = list(color_rgb)
            new_color_rgb[index] = color
            cls.set_color_rgb(shape, new_color_rgb)
            logging.debug("traffic light: new colors saved %s", new_color_rgb)


    @classmethod
    def get_color_state(cls, shape):
        if shape.Type != pplib.MsoShapeType['msoGroup']:
            color = shape.fill.ForeColor.RGB
            try:
                return cls.color_states[cls.get_color_rgb(shape).index(color)]
            except ValueError:
                try:
                    #determine color based on saturation and hue values
                    logging.debug("traffic light: determine color based on hue")
                    r,g,b = get_rgb_from_ole(color)
                    h,_,s = rgb_to_hls(r/255.,g/255.,b/255.)
                    if s < 0.15:
                        return "white"
                    elif h <= 0.09 or h > 0.8:
                        return "red"
                    elif h <= 0.19:
                        return "yellow"
                    else:
                        return "green"
                except:
                    return "green"
        else:
            colors = sorted((shp for shp in shape.GroupItems if shp.AutoShapeType == pplib.MsoAutoShapeType['msoShapeOval']), key=lambda shp: (shp.Top, shp.Left))
            color_white = cls.get_color_white(shape)
            if colors[0].fill.ForeColor.RGB != color_white:
                return "red"
            elif colors[1].fill.ForeColor.RGB != color_white:
                return "yellow"
            elif colors[2].fill.ForeColor.RGB != color_white:
                return "green"
            else:
                return "white"


    @classmethod
    def next_color(cls, shape):
        current_color = cls.get_color_state(shape)
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
