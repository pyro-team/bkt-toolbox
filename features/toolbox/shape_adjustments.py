# -*- coding: utf-8 -*-
'''
Created on 06.02.2018

@author: rdebeerst
'''

from __future__ import absolute_import

import logging

import bkt
import bkt.library.powerpoint as pplib
pt_to_cm = pplib.pt_to_cm
cm_to_pt = pplib.cm_to_pt

class ShapeAdjustments(object):
    adjustment_nums_all = [(1,2), (3,4), (5,6), (7,8)]
    adjustment_nums = (1,2)

    @classmethod
    def set_adjustment_nums(cls, value, context):
        cls.adjustment_nums = value
        context.ribbon.InvalidateControl("anfasser1")
        context.ribbon.InvalidateControl("anfasser2")

    @classmethod
    def adjustment_nums_next(cls, context):
        cur_index = cls.adjustment_nums_all.index(cls.adjustment_nums)
        cls.set_adjustment_nums(cls.adjustment_nums_all[(cur_index+1)%4], context)
    @classmethod
    def adjustment_nums_prev(cls, context):
        cur_index = cls.adjustment_nums_all.index(cls.adjustment_nums)
        cls.set_adjustment_nums(cls.adjustment_nums_all[cur_index-1], context)



    allowed_shape_types = [
        pplib.MsoShapeType['msoAutoShape'],
        pplib.MsoShapeType['msoTextBox'],
        pplib.MsoShapeType['msoCallout'],
        pplib.MsoShapeType['msoPicture'],
    ]
    
    auto_shape_type_settings = {
        # process-arrows
        pplib.MsoAutoShapeType['msoShapeChevron']             : [dict(ref='min(hw)', min=0, max='w')],                                        # type=52
        pplib.MsoAutoShapeType['msoShapePentagon']            : [dict(ref='min(hw)', min=0, max='w')],                                        # type=51
        pplib.MsoAutoShapeType['msoShapeHexagon']             : [dict(ref='min(hw)', min=0, max='w/2')],                                      # type=10
        # rounded/sniped rectangles, max=min(h, w)/2 = 0.5
        pplib.MsoAutoShapeType['msoShapeRoundedRectangle']    : [dict(ref='min(hw)', min=0, max=0.5)],                                        # type=5
        pplib.MsoAutoShapeType['msoShapeSnip1Rectangle']      : [dict(ref='min(hw)', min=0, max=0.5)],
        pplib.MsoAutoShapeType['msoShapeSnip2DiagRectangle']  : [dict(ref='min(hw)', min=0, max=0.5), dict(ref='min(hw)', min=0, max=0.5)],   # type=155
        pplib.MsoAutoShapeType['msoShapeSnip2SameRectangle']  : [dict(ref='min(hw)', min=0, max=0.5), dict(ref='min(hw)', min=0, max=0.5)],   # type=156
        pplib.MsoAutoShapeType['msoShapeSnipRoundRectangle']  : [dict(ref='min(hw)', min=0, max=0.5), dict(ref='min(hw)', min=0, max=0.5)],   # type=157
        pplib.MsoAutoShapeType['msoShapeRound1Rectangle']     : [dict(ref='min(hw)', min=0, max=0.5)],                                        # type=151
        pplib.MsoAutoShapeType['msoShapeRound2DiagRectangle'] : [dict(ref='min(hw)', min=0, max=0.5), dict(ref='min(hw)', min=0, max=0.5)],   # type=153
        pplib.MsoAutoShapeType['msoShapeRound2SameRectangle'] : [dict(ref='min(hw)', min=0, max=0.5), dict(ref='min(hw)', min=0, max=0.5)],   # type=152
        pplib.MsoAutoShapeType['msoShapeBevel']               : [dict(ref='min(hw)', min=0, max=0.5)],                                        # type=15
        pplib.MsoAutoShapeType['msoShapeCross']               : [dict(ref='min(hw)', min=0, max=0.5)],                                        # type=11
        pplib.MsoAutoShapeType['msoShapePlaque']              : [dict(ref='min(hw)', min=0, max=0.5)],                                        # type=29
        pplib.MsoAutoShapeType['msoShapeFoldedCorner']        : [dict(ref='min(hw)', min=0, max=0.5)],                                        # type=16
        # arrows, anchor-1 thickness, anchor-2 arrow-size
        pplib.MsoAutoShapeType['msoShapeRightArrow']          : [dict(ref='h', min=0, max=1), dict(ref='h', min=0, max='w')],                 # type=33
        pplib.MsoAutoShapeType['msoShapeLeftArrow']           : [dict(ref='h', min=0, max=1), dict(ref='h', min=0, max='w')],                 # type=34
        pplib.MsoAutoShapeType['msoShapeLeftRightArrow']      : [dict(ref='h', min=0, max=1), dict(ref='h', min=0, max='w/2')],               # type=37
        pplib.MsoAutoShapeType['msoShapeUpArrow']             : [dict(ref='w', min=0, max=1), dict(ref='w', min=0, max='h')],                 # type=35
        pplib.MsoAutoShapeType['msoShapeDownArrow']           : [dict(ref='w', min=0, max=1), dict(ref='w', min=0, max='h')],                 # type=36
        pplib.MsoAutoShapeType['msoShapeUpDownArrow']         : [dict(ref='w', min=0, max=1), dict(ref='w', min=0, max='h/2')],               # type=38
        pplib.MsoAutoShapeType['msoShapeBentArrow']           : [dict(ref='min(hw)', min=0, max=1), dict(ref='min(hw)*2', min=0, max=0.5), dict(ref='min(hw)', min=0, max=0.5), dict(ref='min(hw)', min=0, max=1)],             # type=41
        pplib.MsoAutoShapeType['msoShapeQuadArrow']           : [dict(ref='min(hw)', min=0, max=1), dict(ref='min(hw)*2', min=0, max=0.5), dict(ref='min(hw)', min=0, max=0.5)],             # type=39
        pplib.MsoAutoShapeType['msoShapeLeftRightUpArrow']    : [dict(ref='min(hw)', min=0, max=1), dict(ref='min(hw)*2', min=0, max=0.5), dict(ref='min(hw)', min=0, max=0.5)],             # type=40
        pplib.MsoAutoShapeType['msoShapeLeftUpArrow']         : [dict(ref='min(hw)', min=0, max=1), dict(ref='min(hw)*2', min=0, max=0.5), dict(ref='min(hw)', min=0, max=0.5)],             # type=43
        pplib.MsoAutoShapeType['msoShapeBentUpArrow']         : [dict(ref='min(hw)', min=0, max=1), dict(ref='min(hw)*2', min=0, max=0.5), dict(ref='min(hw)', min=0, max=0.5)],             # type=44
        pplib.MsoAutoShapeType['msoShapeParallelogram']       : [dict(ref='h', min=0, max='w')],                                              # type=2
        pplib.MsoAutoShapeType['msoShapeTrapezoid']           : [dict(ref='min(hw)', min=0, max="w/2")],                                      # type=3
        pplib.MsoAutoShapeType['msoShapeStripedRightArrow']   : [dict(ref='h', min=0, max="h"), dict(ref='min(hw)', min=0, max='w')],         # type=49
        pplib.MsoAutoShapeType['msoShapeNotchedRightArrow']   : [dict(ref='h', min=0, max="h"), dict(ref='min(hw)', min=0, max='w')],         # type=50
        # database, box
        pplib.MsoAutoShapeType['msoShapeCan']                 : [dict(ref='min(hw)', min=0.001, max='h/2')],                                  # type=13
        pplib.MsoAutoShapeType['msoShapeCube']                : [dict(ref='min(hw)', min=0, max=1)],                                          # type=14
        # donut
        pplib.MsoAutoShapeType['msoShapeDonut']               : [dict(ref='min(hw)', min=0, max=0.5)],                                        # type=18
        pplib.MsoAutoShapeType['msoShapeNoSymbol']            : [dict(ref='min(hw)', min=0, max=0.5)],                                        # type=19
        # moon
        pplib.MsoAutoShapeType['msoShapeMoon']                : [dict(ref='w', min=0, max=0.875)],                                            # type=24
        # tear
        pplib.MsoAutoShapeType['msoShapeTear']                : [dict(ref='w/2', min=0, max='w')],                                            # type=160
        # stripes
        pplib.MsoAutoShapeType['msoShapeDiagonalStripe']      : [dict(ref='h', min=0, max='h')],                                              # type=141
        # math
        pplib.MsoAutoShapeType['msoShapeMathMinus']           : [dict(ref='h', min=0, max='h')],                                              # type=164
        pplib.MsoAutoShapeType['msoShapeMathPlus']            : [dict(ref='min(hw)', min=0, max=2)],                                          # type=163
        pplib.MsoAutoShapeType['msoShapeMathMultiply']        : [dict(ref='min(hw)', min=0, max=1)],                                          # type=165
        #triangle, x-agons
        pplib.MsoAutoShapeType['msoShapeIsoscelesTriangle']   : [dict(ref='w', min=0, max='w')],                                              # type=7
        pplib.MsoAutoShapeType['msoShapeOctagon']             : [dict(ref='min(hw)', min=0, max=0.5)],                                        # type=6
        # brackes, braces
        pplib.MsoAutoShapeType['msoShapeDoubleBracket']       : [dict(ref='min(hw)', min=0, max=0.5)],                                        # type=26
        pplib.MsoAutoShapeType['msoShapeDoubleBrace']         : [dict(ref='min(hw)', min=0, max=0.5)],                                        # type=27
        pplib.MsoAutoShapeType['msoShapeLeftBracket']         : [dict(ref='w', min=0, max='h/2')],                                            # type=29
        pplib.MsoAutoShapeType['msoShapeRightBracket']        : [dict(ref='w', min=0, max='h/2')],                                            # type=30
        pplib.MsoAutoShapeType['msoShapeLeftBrace']           : [dict(ref='w', min=0, max='h/4'), dict(ref='h', min=0, max=1)],               # type=31
        pplib.MsoAutoShapeType['msoShapeRightBrace']          : [dict(ref='w', min=0, max='h/4'), dict(ref='h', min=0, max=1)],               # type=32
        # frame, frame-corner
        pplib.MsoAutoShapeType['msoShapeFrame']               : [dict(ref='min(hw)', min=0, max=0.5)],                                        # type=158
        pplib.MsoAutoShapeType['msoShapeHalfFrame']           : [dict(ref='min(hw)', min=0, max='h'), dict(ref='min(hw)', min=0, max='w')],   # type=159
        pplib.MsoAutoShapeType['msoShapeCorner']              : [dict(ref='min(hw)', min=0, max='h'), dict(ref='min(hw)', min=0, max='w')],   # type=162
        # arc-shape, pie
        pplib.MsoAutoShapeType['msoShapeArc']                 : [dict(ref='deg', min=-180, max=180), dict(ref='deg', min=-180, max=180)],     # type=25
        pplib.MsoAutoShapeType['msoShapeBlockArc']            : [dict(ref='deg', min=-180, max=180), dict(ref='deg', min=-180, max=180), dict(ref='min(hw)', min=0, max=0.5)],     # type=20
        pplib.MsoAutoShapeType['msoShapeChord']               : [dict(ref='deg', min=-180, max=180), dict(ref='deg', min=-180, max=180)],     # type=161
        pplib.MsoAutoShapeType['msoShapePie']                 : [dict(ref='deg', min=-180, max=180), dict(ref='deg', min=-180, max=180)],     # type=142
        #callouts
        pplib.MsoAutoShapeType['msoShapeRectangularCallout']            : [dict(ref='w', min=None, max=None), dict(ref='h', min=None, max=None)],                                          # type=105
        pplib.MsoAutoShapeType['msoShapeOvalCallout']                   : [dict(ref='w', min=None, max=None), dict(ref='h', min=None, max=None)],                                          # type=107
        pplib.MsoAutoShapeType['msoShapeCloudCallout']                  : [dict(ref='w', min=None, max=None), dict(ref='h', min=None, max=None)],                                          # type=108
        pplib.MsoAutoShapeType['msoShapeRoundedRectangularCallout']     : [dict(ref='w', min=None, max=None), dict(ref='h', min=None, max=None), dict(ref='min(hw)', min=0, max=0.5)],     # type=106
        # connector line (auto shape type = -2)
        "connector"                                           : [dict(ref='w', min=None, max=None)],
    }
    
        
    @classmethod
    def set_shapes_rounded_corner_size(cls, shapes, num, value):
        # for shape in shapes:
        for shape in pplib.iterate_shape_subshapes(shapes):
            cls.set_adjustment(shape, num, value)

    @classmethod
    def get_shapes_rounded_corner_size(cls, shapes, num):
        try:
            for shape in pplib.iterate_shape_subshapes(shapes):
                if shape.adjustments.count >= num:
                    return cls.get_adjustment(shape, num)
                else:
                    return None
        except:
            return None

    @classmethod
    def get_enabled(cls, shapes, num):
        try:
            for shape in pplib.iterate_shape_subshapes(shapes):
                return cls.get_shape_type(shape) in cls.allowed_shape_types and (shape.adjustments.count >= num)
        except:
            return False

    @classmethod
    def get_label(cls):
        return "%s und %s" % cls.adjustment_nums
        # return "Werte %s u. %s" % cls.adjustment_nums

    @classmethod
    def get_shape_type(cls, shape):
        try:
            if shape.Type == pplib.MsoShapeType['msoPlaceholder']:
                return shape.PlaceholderFormat.ContainedType
            else:
                return shape.Type
        except:
            return None

    @classmethod
    def get_shape_autotype(cls, shape):
        if shape.AutoShapeType == pplib.MsoAutoShapeType['msoShapeMixed'] and shape.Connector == -1:
            return "connector"
        else:
            return shape.AutoShapeType


    @classmethod
    def set_adjustment(cls, shape, num, value):
        ''' sets n's adjustment of shape, where value is assumed to be cm-value '''
        if cls.get_shape_type(shape) in cls.allowed_shape_types and shape.adjustments.count >= num:
            if type(value) == str:
                value = float(value.replace(',', '.'))
            # if cls.get_shape_autotype(shape) in cls.auto_shape_type_settings.keys():
            try:
                ref, minimum, maximum = cls.get_ref_min_max(shape, num)
                pt_value = cm_to_pt(value)
                logging.debug('shape adjustment pt value=%s' % pt_value)
                if minimum != None:
                    pt_value = max(minimum, pt_value)
                if maximum != None:
                    pt_value = min(maximum, pt_value)
                shape.adjustments.item[num] = pt_value / ref
            # else:
            except (KeyError, IndexError): #KeyError = shape type is not in database, IndexError = adjustment number is not in database
                shape.adjustments.item[num] = value / 100

    @classmethod
    def get_adjustment(cls, shape, num):
        ''' returns n's adjustment of shape, transformed to cm '''
        if cls.get_shape_type(shape) in cls.allowed_shape_types and shape.adjustments.count >= num:
            # if cls.get_shape_autotype(shape) in cls.auto_shape_type_settings.keys():
            try:
                ref, _, _ = cls.get_ref_min_max(shape, num)
                return round(pt_to_cm( shape.adjustments.item[num] * ref ), 2)
            # else:
            except (KeyError, IndexError): #KeyError = shape type is not in database, IndexError = adjustment number is not in database
                return round( shape.adjustments.item[num] * 100, 2)
    
    @classmethod
    def get_ref_min_max(cls, shape, num):
        ''' returns reference-values (minimum, maxium) for adjustments depending on shape-type '''
        ref_settings = cls.auto_shape_type_settings[cls.get_shape_autotype(shape)][num-1]
        ref = cls.get_ref_value(shape, ref_settings['ref'])
        return ref, cls.get_ref_value(shape, ref_settings['min'], ref=ref), cls.get_ref_value(shape, ref_settings['max'], ref=ref)
    
    @classmethod
    def get_ref_value(cls, shape, ref_key, ref=None):
        ''' computes reference-values (minimum, maxium) for adjustments depending on ref-key '''
        if ref != None and type(ref_key) in [int,float]:
            value = ref_key*ref
        elif ref_key == 'deg':
            value = cm_to_pt(1) #convert to pt as setter/getter are converting back -> "double conversion"
        elif ref_key == 'h':
            value = shape.Height
        elif ref_key == 'h/2':
            value = shape.Height/2
        elif ref_key == 'h/4':
            value = shape.Height/4
        elif ref_key == 'w':
            value = shape.Width
        elif ref_key == 'w/2':
            value = shape.Width/2
        elif ref_key == 'min(hw)/2':
            value = min(shape.height, shape.width)/2
        elif ref_key == 'min(hw)*2':
            value = min(shape.height, shape.width)*2
        elif ref_key == 'min(hw)':
            value = min(shape.height, shape.width)
        else:
            value = None
        return value
        
    @classmethod
    def equalize_adjustments(cls, shapes):
        master = shapes[0]
        for shape in shapes:
            for i in range(1,master.adjustments.count+1):
                try:
                    cls.set_adjustment(shape, i, cls.get_adjustment(master, i))
                except:
                    logging.exception("error equalizing shape adjustments")
                # if i <= shape.adjustments.count:
                #     shape.adjustments.item[i] = master.adjustments.item[i]

    @classmethod
    def reset_adjustments(cls, shapes):
        for shape in shapes:
            shape_db = pplib.GlobalShapeDb.get_by_shape(shape)
            default_adj = shape_db["adjustments"]
            for i in range(shape.adjustments.count):
                shape.adjustments.item[i+1] = default_adj[i]["default"]


    @classmethod
    def get_shape_details(cls, shape):
        def _get_key_by_value(search_dict, value):
            try:
                return search_dict.keys()[search_dict.values().index(value)]
            except:
                return "?"

        import bkt.console
        infomsg  = "{:16} {:3} - {}\n".format("ID+Name:", shape.id, shape.name)

        shape_type = shape.Type
        if shape_type == pplib.MsoShapeType['msoPlaceholder']:
            shape_type = shape.PlaceholderFormat.ContainedType
            infomsg += "{:16} {:3} - Placeholder: {}\n".format("Shape-Typ:", shape_type, _get_key_by_value(pplib.MsoShapeType, shape_type))
        else:
            infomsg += "{:16} {:3} - {}\n".format("Shape-Typ:", shape_type, _get_key_by_value(pplib.MsoShapeType, shape_type))

        shape_autotype = cls.get_shape_autotype(shape)
        if shape_autotype == "connector":
            infomsg += "{:16}   {}\n".format("Auto-Shape-Typ:", "Connector")
        else:
            infomsg += "{:16} {:3} - {}\n".format("Auto-Shape-Typ:", shape_autotype, _get_key_by_value(pplib.MsoAutoShapeType, shape_autotype))
        
        try:
            shape_adjustments = shape.adjustments.count
            shape_db = pplib.GlobalShapeDb.get_by_shape(shape)
            default_adj = shape_db["adjustments"]
        except: #ValueError for Groups
            shape_adjustments = 0
        infomsg += "{:16} {:3}\n\n".format("Adjust.-Werte:", shape_adjustments)

        infomsg += "│ {:2} {:>8} │ {:10} {:>6} {:>6} {:>8} │\n".format("#", "Wert", "ref", "min", "max", "default")
        infomsg += "├─────────────┼───────────────────────────────────┤\n"
        for i in range(1,shape_adjustments+1):
            try:
                bktlib = cls.auto_shape_type_settings[shape_autotype][i-1]
            except:
                bktlib = dict(ref="", min="", max="")
            infomsg += "│ {:<2} {:8.4} │ {:10} {:>6} {:>6} {:8.4} │\n".format(i, shape.adjustments.item[i], bktlib["ref"], bktlib["min"], bktlib["max"], default_adj[i-1]["default"])

        bkt.console.show_message(bkt.helpers.endings_to_windows(infomsg))



adjustments_group = bkt.ribbon.Group(
    id="bkt_adjustments_group",
    label = u"Fine-Tuning",
    image_mso='ShapeArc',
    children=[
        # bkt.ribbon.Label(label="Rund./Spitzen"),
        bkt.ribbon.Box(box_style="horizontal",
            children=[
                bkt.ribbon.Button(
                    label="<", #"#«",
                    screentip="Vorherige Anfasser-Werte anzeigen",
                    supertip="Springe zu vorheriger Seite für die Anfasser-Werte.",
                    on_action=bkt.Callback(ShapeAdjustments.adjustment_nums_prev, context=True),
                ),
                bkt.ribbon.Menu(
                    # label=u"Rund./Spitz.",
                    get_label=bkt.Callback(ShapeAdjustments.get_label),
                    screentip="Anfasser-Werte wechseln",
                    supertip="Anfasser beeinflussen die Shape-Form, bspw. die Rundung an Ecken oder die Spitze von Pfeilen. Je nach Shape-Typ kann es bis zu 8 Anfasser-Werte geben, die manuell gesetzt werden können. Hier kann paarweise zwischen den Werten umgeschaltet werden.",
                    children = [
                        bkt.ribbon.MenuSeparator(title="Rundungen/Spitzen/Ecken"),
                        bkt.ribbon.ToggleButton(
                            label="Werte 1 und 2",
                            supertip="Anfasser-Werte Nr. 1 und 2 in Spinner-Boxen anpassen",
                            on_toggle_action=bkt.Callback(lambda pressed, context: ShapeAdjustments.set_adjustment_nums((1,2), context), context=True),
                            get_pressed=bkt.Callback(lambda: ShapeAdjustments.adjustment_nums == (1,2)),
                        ),
                        bkt.ribbon.ToggleButton(
                            label="Werte 3 und 4",
                            supertip="Anfasser-Werte Nr. 3 und 4 in Spinner-Boxen anpassen",
                            on_toggle_action=bkt.Callback(lambda pressed, context: ShapeAdjustments.set_adjustment_nums((3,4), context), context=True),
                            get_pressed=bkt.Callback(lambda: ShapeAdjustments.adjustment_nums == (3,4)),
                        ),
                        bkt.ribbon.ToggleButton(
                            label="Werte 5 und 6",
                            supertip="Anfasser-Werte Nr. 5 und 6 in Spinner-Boxen anpassen",
                            on_toggle_action=bkt.Callback(lambda pressed, context: ShapeAdjustments.set_adjustment_nums((5,6), context), context=True),
                            get_pressed=bkt.Callback(lambda: ShapeAdjustments.adjustment_nums == (5,6)),
                        ),
                        bkt.ribbon.ToggleButton(
                            label="Werte 7 und 8",
                            supertip="Anfasser-Werte Nr. 7 und 8 in Spinner-Boxen anpassen",
                            on_toggle_action=bkt.Callback(lambda pressed, context: ShapeAdjustments.set_adjustment_nums((7,8), context), context=True),
                            get_pressed=bkt.Callback(lambda: ShapeAdjustments.adjustment_nums == (7,8)),
                        ),
                        bkt.ribbon.MenuSeparator(),
                        bkt.ribbon.Button(
                            label="Alle Shapes zurücksetzen",
                            image_mso='ResetCurrentView',
                            supertip="Shape-Anfasser-Werte für alle ausgewählten Shapes auf die Originalwerte zurücksetzen",
                            on_action=bkt.Callback(ShapeAdjustments.reset_adjustments, shapes=True),
                        ),
                        bkt.ribbon.Button(
                            label="Alle Shapes angleichen",
                            image_mso='ShapeArc',
                            supertip="Shape-Anfasser-Werte für alle ausgewählten Shapes entsprechend des zuerst gewählten Shapes angleichen",
                            on_action=bkt.Callback(ShapeAdjustments.equalize_adjustments, shapes=True),
                        ),
                        bkt.ribbon.MenuSeparator(),
                        bkt.ribbon.Button(
                            label="Shape-Details anzeigen",
                            image_mso='PropertiesForm',
                            supertip="Zeigt ein Fenster mit ausführlichen Informationen über das gewählte Shape an",
                            on_action=bkt.Callback(ShapeAdjustments.get_shape_details, shape=True),
                            get_enabled=bkt.apps.ppt_shapes_exactly1_selected,
                        ),
                    ]
                ),
                bkt.ribbon.Button(
                    label=">", #"»",
                    screentip="Nächste Anfasser-Werte anzeigen",
                    supertip="Springe zu nächsten Seite für die Anfasser-Werte.",
                    on_action=bkt.Callback(ShapeAdjustments.adjustment_nums_next, context=True),
                ),
            ]
        ),
        bkt.ribbon.RoundingSpinnerBox(
            id=u"anfasser1",
            label=u"Anfasser 1",
            show_label=False,
            image_mso='ShapeArc',
            screentip="Breite von Rundung/Pfeilspitzen/Ecken/etc.",
            supertip="Ändere die Breite von Rundungen (bspw. abgerundetes Rechteck), Pfeilspitzen (bspw. Richtungspfeil) oder Ecken (bspw. abgeschnittenes Rechteck) auf das angegebene Maß (je nach Shape-Typ in cm oder %).",
            on_change   = bkt.Callback(lambda shapes, value: ShapeAdjustments.set_shapes_rounded_corner_size(shapes, ShapeAdjustments.adjustment_nums[0], value)),
            get_text    = bkt.Callback(lambda shapes : ShapeAdjustments.get_shapes_rounded_corner_size(shapes, ShapeAdjustments.adjustment_nums[0])),
            get_enabled = bkt.Callback(lambda shapes : ShapeAdjustments.get_enabled(shapes, ShapeAdjustments.adjustment_nums[0])),
            round_cm = True,
            size_string = '####',
        ),
        bkt.ribbon.RoundingSpinnerBox(
            id=u"anfasser2",
            label=u"Anfasser 2",
            show_label=False,
            image_mso='ShapeCurve',
            screentip="Breite von Rundung/Pfeilspitzen/Ecken/etc.",
            supertip="Ändere die Breite von Rundungen (bspw. abgerundetes Rechteck), Pfeilspitzen (bspw. Richtungspfeil) oder Ecken (bspw. abgeschnittenes Rechteck) auf das angegebene Maß (je nach Shape-Typ in cm oder %).",
            on_change   = bkt.Callback(lambda shapes, value: ShapeAdjustments.set_shapes_rounded_corner_size(shapes, ShapeAdjustments.adjustment_nums[1], value)),
            get_text    = bkt.Callback(lambda shapes : ShapeAdjustments.get_shapes_rounded_corner_size(shapes, ShapeAdjustments.adjustment_nums[1])),
            get_enabled = bkt.Callback(lambda shapes : ShapeAdjustments.get_enabled(shapes, ShapeAdjustments.adjustment_nums[1])),
            round_cm = True,
            size_string = '####',
        )
    ]
)
