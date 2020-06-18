# -*- coding: utf-8 -*-

from __future__ import absolute_import, division

import logging

import bkt.ui
notify_property = bkt.ui.notify_property

import bkt.library.powerpoint as pplib


# =================
# = FUNCTIONALITY =
# =================


class ScaleShapes(object):
    excluded_shapes = [
        pplib.MsoShapeType["msoChart"],
        pplib.MsoShapeType["msoDiagram"],
        pplib.MsoShapeType["msoSmartArt"],
        pplib.MsoShapeType["msoTable"],
    ]
    
    @classmethod
    def scale_shapes(cls, shapes, value, scale="percent", settings=['size', 'line', 'shadow', 'font', 'margin', 'indent']):
        shapes_exluded_message = False
        for shape in shapes:
            try:
                if shape.Type in cls.excluded_shapes:
                    raise TypeError("shape type not supported")
                cls.scale_shape(shape, value, scale, settings)
            except TypeError:
                logging.warning("scale shape failed due to unsupported shape type")
                shapes_exluded_message = True
            except:
                logging.exception("scale shape failed")
        
        if shapes_exluded_message:
            bkt.message.warning("Shape-Skalierung wird nicht unterstützt für Diagramme, Tabellen und Smart-Arts!")
    
    @classmethod
    def scale_shape(cls, shape, value, scale, settings):
        if scale == "height":
            rel = value / shape.height
        elif scale == "width":
            rel = value / shape.width
        else:
            rel = value
            logging.warning("scale percent: %s", rel)

        if 'size' in settings:
            shape.ScaleWidth(rel,0,0)
            if not shape.LockAspectRatio:
                shape.ScaleHeight(rel,0,0)

        for shape in pplib.iterate_shape_subshapes([shape]):
            if 'line' in settings and shape.Line.Visible:
                shape.Line.Weight *= rel

            if 'shadow' in settings and shape.Shadow.Visible:
                shape.shadow.Blur *= rel
                shape.shadow.OffsetX *= rel
                shape.shadow.OffsetY *= rel

            if shape.HasTextFrame:
                textframe = shape.TextFrame2
                if 'margin' in settings:
                    textframe.MarginTop *= rel
                    textframe.MarginBottom *= rel
                    textframe.MarginLeft *= rel
                    textframe.MarginRight *= rel
                #per run
                textrange = textframe.TextRange
                for run in textrange.Runs():
                    if 'font' in settings:
                        run.Font.Size *= rel
                        if run.Font.Line.Visible:
                            run.Font.Line.Weight *= rel

                    if 'indent' in settings:
                        parf = run.ParagraphFormat
                        if not parf.LineRuleBefore:
                            parf.SpaceBefore *= rel
                        if not parf.LineRuleAfter:
                            parf.SpaceAfter *= rel
                        if not parf.LineRuleWithin:
                            parf.SpaceWithin *= rel

                        parf.FirstLineIndent *= rel
                        parf.LeftIndent *= rel
                        parf.RightIndent *= rel



# =======================
# = UI MODEL AND WINDOW =
# =======================


class ViewModel(bkt.ui.ViewModelSingleton):
    
    def __init__(self):
        super(ViewModel, self).__init__()
        
        self._scale_target = "percent"
        
        self._target_percent = 1.0
        self._target_width = 10
        self._target_height = 10

        self._orig_width = 10
        self._orig_height = 10

        self._settings = set(['size', 'line', 'shadow', 'font', 'margin', 'indent'])
    
    def set_dimensions(self, width, height):
        self._orig_width = width
        self._orig_height = height

        if self._scale_target == "height":
            self._target_percent = self._target_height/height
            self._target_width = width * self._target_percent
        elif self._scale_target == "width":
            self._target_percent = self._target_width/width
            self._target_height = height * self._target_percent
        else:
            self._target_width = width * self._target_percent
            self._target_height = height * self._target_percent

        self.OnPropertyChanged('target_percent')
        self.OnPropertyChanged('target_height')
        self.OnPropertyChanged('target_width')
    
    @notify_property
    def scale_percent(self):
        return self._scale_target == "percent"
    @scale_percent.setter
    def scale_percent(self, value):
        if value:
            self._scale_target = "percent"
    
    @notify_property
    def scale_width(self):
        return self._scale_target == "width"
    @scale_width.setter
    def scale_width(self, value):
        if value:
            self._scale_target = "width"
    
    @notify_property
    def scale_height(self):
        return self._scale_target == "height"
    @scale_height.setter
    def scale_height(self, value):
        if value:
            self._scale_target = "height"
    
    
    @notify_property
    def target_percent(self):
        return self._target_percent*100
    @target_percent.setter
    def target_percent(self, value):
        value = value/100
        self._target_percent = value
        self._target_width = self._orig_width*value
        self._target_height = self._orig_height*value
        self.OnPropertyChanged('target_height')
        self.OnPropertyChanged('target_width')
        self.scale_percent = True
    
    @notify_property
    def target_width(self):
        return self._target_width
    @target_width.setter
    def target_width(self, value):
        self._target_width = value
        self._target_percent = value/self._orig_width
        self._target_height = self._target_percent*self._orig_height
        self.OnPropertyChanged('target_percent')
        self.OnPropertyChanged('target_height')
        self.scale_width = True
    
    @notify_property
    def target_height(self):
        return self._target_height
    @target_height.setter
    def target_height(self, value):
        self._target_height = value
        self._target_percent = value/self._orig_height
        self._target_width = self._target_percent*self._orig_width
        self.OnPropertyChanged('target_percent')
        self.OnPropertyChanged('target_width')
        self.scale_height = True
    
    @notify_property
    def settings_size(self):
        return 'size' in self._settings
    @settings_size.setter
    def settings_size(self, value):
        if value:
            self._settings.add('size')
        else:
            self._settings.remove('size')
    
    @notify_property
    def settings_line(self):
        return 'line' in self._settings
    @settings_line.setter
    def settings_line(self, value):
        if value:
            self._settings.add('line')
        else:
            self._settings.remove('line')
    
    @notify_property
    def settings_shadow(self):
        return 'shadow' in self._settings
    @settings_shadow.setter
    def settings_shadow(self, value):
        if value:
            self._settings.add('shadow')
        else:
            self._settings.remove('shadow')
    
    @notify_property
    def settings_font(self):
        return 'font' in self._settings
    @settings_font.setter
    def settings_font(self, value):
        if value:
            self._settings.add('font')
        else:
            self._settings.remove('font')
    
    @notify_property
    def settings_margin(self):
        return 'margin' in self._settings
    @settings_margin.setter
    def settings_margin(self, value):
        if value:
            self._settings.add('margin')
        else:
            self._settings.remove('margin')
    
    @notify_property
    def settings_indent(self):
        return 'indent' in self._settings
    @settings_indent.setter
    def settings_indent(self, value):
        if value:
            self._settings.add('indent')
        else:
            self._settings.remove('indent')


class ShapeScaleWindow(bkt.ui.WpfWindowAbstract):
    _xamlname = 'shape_scale'
    _vm_class = ViewModel
    
    def __init__(self, context, shapes):
        super(ShapeScaleWindow, self).__init__(context)

        self.ref_shapes = shapes
        self._vm.set_dimensions(pplib.pt_to_cm(shapes[0].width), pplib.pt_to_cm(shapes[0].height))

    def reset(self, sender, event):
        self._vm.target_percent = 1.0
    
    def scale(self, sender, event):
        vm = self._vm
        if vm._scale_target == "height":
            value = pplib.cm_to_pt(vm._target_height)
        elif vm._scale_target == "width":
            value = pplib.cm_to_pt(vm._target_width)
        else:
            value = vm._target_percent
        ScaleShapes.scale_shapes(self.ref_shapes, value, vm._scale_target, vm._settings)
        self.Close()