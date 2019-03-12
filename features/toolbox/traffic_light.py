# -*- coding: utf-8 -*-

import logging
import os.path

import bkt

# wpf basics
# import clr
# clr.AddReference("IronPython.Wpf")
# import wpf

# import System
# from System.Windows import Controls, Window

# property binding
# import bkt
# from bkt.library.wpf.notify import NotifyPropertyChangedBase, notify_property



# ================
# = dialog logic =
# ================

class Ampel(object):
    #BKT_CONTEXTDIALOG = 'BKT_CONTEXTDIALOG'
    BKT_DIALOG_AMPEL = 'BKT_DIALOG_AMPEL3'

    @classmethod
    def create(cls, slide):
        logging.debug("create ampel 3")

        from System import Array

        # slide=cls.context.app.activewindow.selection.sliderange[1]

        shapeCount = slide.shapes.count
        shapes = [
            slide.shapes.addshape(1, 100, 100, 30, 80), #rect
            slide.shapes.addshape(9, 105, 105, 20, 20), #red
            slide.shapes.addshape(9, 105, 130, 20, 20), #yellow
            slide.shapes.addshape(9, 105, 155, 20, 20) #green
        ]
        for shape in shapes:
            shape.fill.ForeColor.RGB = 14277081
            shape.line.weight = 0.75
            shape.line.ForeColor.RGB = 0
        # gruppieren
        grp = slide.Shapes.Range(Array[int](range(shapeCount+1, shapeCount+5))).group()
        grp.select()
        grp.LockAspectRatio = -1 #msoTrue
        grp.Tags.Add(bkt.contextdialogs.BKT_CONTEXTDIALOG_TAGKEY, cls.BKT_DIALOG_AMPEL)

        cls.set_color(grp)
        
    
    @classmethod
    def set_color(cls, shape, color="red"):
        colors = [shp for shp in shape.GroupItems if shp.AutoShapeType == 9]
        colors.sort(key=lambda shp: shp.Top)
        colors[0].fill.ForeColor.RGB = 16777215 # white
        colors[1].fill.ForeColor.RGB = 16777215 # white
        colors[2].fill.ForeColor.RGB = 16777215 # white
        if color == "red":
            colors[0].fill.ForeColor.RGB = 255 #red
        elif color == "yellow":
            colors[1].fill.ForeColor.RGB = 65535 #yellow
        elif color == "green":
            colors[2].fill.ForeColor.RGB = 5287936 #green
        
        
    @classmethod
    def get_color(cls, shape):
        colors = [shp for shp in shape.GroupItems if shp.AutoShapeType == 9]
        colors.sort(key=lambda shp: shp.Top)
        if colors[0].fill.ForeColor.RGB == 255:
            return "red"
        elif colors[1].fill.ForeColor.RGB == 65535:
            return "yellow"
        else:
            return "green"





# ==============
# = view model =
# ==============

# class ViewModel(NotifyPropertyChangedBase):
#     '''
#     empty view model for traffic-light popup-window
#     '''

#     def __init__(self):
#         super(ViewModel, self).__init__()


# ==========
# = window =
# ==========

# class TrafficPopup(bkt.ui.WpfPopupAbstract):
class TrafficPopup(bkt.ui.WpfWindowAbstract):
    _filename = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'traffic_light_dialog2.xaml')
    '''
    class representing a opup-dialog for a traffic-light-shape
    '''
    
    def __init__(self, context=None):
        self.IsPopup = True
        self._context = context

        super(TrafficPopup, self).__init__()

    # def __init__(self, context=None):
    #     filename=os.path.join(os.path.dirname(os.path.realpath(__file__)), 'traffic_light_dialog.xaml')
    #     wpf.LoadComponent(self, filename)
    #     self._vm = ViewModel()
    #     self._context = context
    #     self.DataContext = self._vm

    def btnred(self, sender, event):
        try:
            shapes = list(iter(self._context.app.activewindow.selection.shaperange))
            for shape in shapes:
                Ampel.set_color(shape, "red")
            self._context.app.ActiveWindow.Activate()
        except:
            logging.error(traceback.format_exc())

    def btnyellow(self, sender, event):
        try:
            shapes = list(iter(self._context.app.activewindow.selection.shaperange))
            for shape in shapes:
                Ampel.set_color(shape, "yellow")
            self._context.app.ActiveWindow.Activate()
        except:
            logging.error(traceback.format_exc())

    def btngreen(self, sender, event):
        try:
            shapes = list(iter(self._context.app.activewindow.selection.shaperange))
            for shape in shapes:
                Ampel.set_color(shape, "green")
            self._context.app.ActiveWindow.Activate()
        except:
            logging.error(traceback.format_exc())

    def btnwhite(self, sender, event):
        try:
            shapes = list(iter(self._context.app.activewindow.selection.shaperange))
            for shape in shapes:
                Ampel.set_color(shape, "white")
            self._context.app.ActiveWindow.Activate()
        except:
            logging.error(traceback.format_exc())
    


# ======================
# = register UI in BKT =
# ======================

# register dialog
logging.debug("traffic light: registering dialog")
bkt.powerpoint.context_dialogs.register_dialog(
    bkt.contextdialogs.ContextDialog(
        id=Ampel.BKT_DIALOG_AMPEL,
        module=None,
        window_class=TrafficPopup
    )
)


# ribbon button, for usage elsewhere
traffic_light_create_button = bkt.ribbon.Button(
    label="Ampel",
    image="traffic_light",
    screentip='Status-Ampel erstellen',
    supertip="Füge eine Status-Ampel ein. Die Status-Farbe der Ampel kann per Kontext-Dialog konfiguriert werden.",
    on_action=bkt.Callback(Ampel.create)
)


