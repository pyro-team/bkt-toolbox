# -*- coding: utf-8 -*-
'''
Created on 2017-07-24
@author: Florian Stallmann
'''



import logging
from collections import OrderedDict

import bkt
import bkt.library.powerpoint as pplib
import bkt.library.algorithms as algos

import bkt.dotnet
F = bkt.dotnet.import_forms()
D = bkt.dotnet.import_drawing()

# import pycountry
# import gettext
# german = gettext.translation('iso3166', pycountry.LOCALES_DIR, languages=['de'])
# german.install()

from . import wpftest



class ColorSelectorWindow(object):
    def __init__(self, application):
        prompt = F.Form();
        prompt.Width = 80;
        prompt.Height = 600;
        prompt.Text = "Farben";
        prompt.StartPosition = F.FormStartPosition.CenterScreen;
        prompt.AutoSize = False
        prompt.MinimizeBox = False
        prompt.MaximizeBox = False
        prompt.ShowInTaskbar = False
        prompt.FormBorderStyle = F.FormBorderStyle.FixedToolWindow
        prompt.TopMost = True
        prompt.TopLevel = True
        prompt.SizeGripStyle = F.SizeGripStyle.Hide

        testbtn = F.Button()
        testbtn.Text = "OK"
        testbtn.Width = 20
        testbtn.Height = 20
        testbtn.BackColor = D.Color.LightGreen
        testbtn.Click += self.testbtn

        buttonsPanel = F.FlowLayoutPanel()
        buttonsPanel.AutoSize = True
        buttonsPanel.WrapContents = False
        buttonsPanel.FlowDirection = F.FlowDirection.TopDown
        buttonsPanel.Controls.Add(testbtn)
        buttonsPanel.Top = 5
        buttonsPanel.Left = 5
        buttonsPanel.Height = 200

        prompt.Controls.Add(buttonsPanel)

        self.prompt = prompt
        self.is_shown = False
        self.application = application

    def testbtn(self, sender, e):
        import bkt.console
        bkt.console.show_message("test")

    def show(self):
        self.prompt.Show()
        self.is_shown = True

    def close(self):
        self.prompt.Close()
        self.is_shown = False

    def switchWindow(self):
        if self.is_shown:
            self.close()
        else:
            self.show()


class WindowTests(object):

    @staticmethod
    def open_color_dialog(shape):
        cd = F.ColorDialog()
        cd.Color = D.ColorTranslator.FromOle(shape.Fill.ForeColor.RGB)
        cd.FullOpen = True
        if cd.ShowDialog() == F.DialogResult.OK:
            color = D.ColorTranslator.ToOle(cd.Color)
            shape.Fill.ForeColor.RGB = color
            # bkt.message("Farbe: %r" % color)

    @staticmethod
    def colorwindow(application):
        ColorSelectorWindow(application).switchWindow()


class FormattingTestFunctions(object):
    shape_db = None

    @classmethod
    def build_shape_db(cls, shapes):
        import io
        import json

        from bkt import console

        from json import encoder
        encoder.FLOAT_REPR = lambda o: format(o, '.4f')

        from toolbox.shape_adjustments import ShapeAdjustments

        cls.shape_db = OrderedDict()

        def _get_key_by_value(search_dict, value):
            try:
                return list(search_dict.keys())[list(search_dict.values()).index(value)]
            except:
                return "?"
        
        for shape in sorted(shapes, key=lambda s: s.AutoShapeType):
            dba = cls.shape_db[shape.AutoShapeType] = OrderedDict()
            dba['id']                   = shape.AutoShapeType
            dba['name']                 = _get_key_by_value(pplib.MsoAutoShapeType, shape.AutoShapeType)
            dba['ratio']                = float(shape.width/shape.height)
            dba['adjustments_count']    = shape.adjustments.count
            # dba['adjustments']          = [float(a) for a in shape.adjustments]
            dba['adjustments']          = []

            if shape.adjustments.count == 0:
                continue

            adj_settings = ShapeAdjustments.auto_shape_type_settings.get(shape.AutoShapeType, [])
            if not adj_settings:
                logging.warning("Shape DB: No adj settings for type %s", shape.AutoShapeType)
                shape.fill.forecolor.rgb = 3243501
            for i,a in enumerate(shape.adjustments):
                try:
                    d = adj_settings[i]
                except:
                    # logging.warning("Shape DB: Missing adj setting for type %s, i %s", shape.AutoShapeType, i)
                    d = dict()
                d['default'] = float(a)
                dba['adjustments'].append(d)

            dba['adjustments']
        
        file = bkt.helpers.file_base_path_join(__file__, "shapedb.json")
        with io.open(file, 'w', encoding='utf-8') as json_file:
            json.dump(cls.shape_db, json_file, ensure_ascii=False, indent=2)

        # console.show_message(json.dumps(cls.shape_db, indent=2))

    @staticmethod
    def export_as_png(presentation, slides):
        from System import Array, Int32
        import bkt.library.graphics as glib

        for slide in slides:
            shape_indices = []
            shape_index = 1
            for shape in slide.shapes:
                if shape.type != 14 and shape.visible == -1:
                    # shape is not a placeholder and visible
                    shape_indices.append(shape_index)
                shape_index+=1
            # select shapes
            shape_range = slide.shapes.Range(Array[Int32](shape_indices))
            path = presentation.Path + "\\" + str(slide.SlideIndex) + ".png"
            path2 = presentation.Path + "\\" + str(slide.SlideIndex) + "-sq.png"
            shape_range.Export(path, 2) #2=ppShapeFormatPNG

            #make square
            glib.make_thumbnail(path, 32, 32, path2)

    @staticmethod
    def table_formatter(shape):
        colors = [13936767, 15388863, 15388863]
        index = 0
        for row in shape.table.rows:
            cell_in_row_selected = False
            for cell in row.cells:
                if cell.selected:
                    cell_in_row_selected = True
                    cell.shape.fill.forecolor.rgb = colors[index]
            if cell_in_row_selected:
                index = (index+1) % len(colors)
    

    current_control = None
    customui_control = None
    @classmethod
    def control_position(cls, context, current_control, customui_control):
        cls.current_control = current_control
        cls.customui_control = customui_control
        print((current_control.id))
        try:
            print(("test1: %s" % context.app.CommandBars.ActionControl))
        except:
            pass
        try:
            print(("test2: %s" % context.app.Caller))
        except:
            pass


class CustomFontStyles(object):
    @staticmethod
    def get_pressed(selection):
        try:
            return selection.TextRange2.Font.Name == "Lucida Sans DemiBold"
        except:
            shapes = pplib.get_shapes_from_selection(selection)
            for textframe in pplib.iterate_shape_textframes(shapes):
                return textframe.TextRange.Font.Name == "Lucida Sans DemiBold"

    @staticmethod
    def on_toggle_action(pressed, selection):
        font = "Lucida Sans DemiBold" if pressed else "Lucida Sans"
        try:
            selection.TextRange2.Font.Name = font
            #fails for selected table cells
        except:
            shapes = pplib.get_shapes_from_selection(selection)
            for textframe in pplib.iterate_shape_textframes(shapes):
                textframe.TextRange.Font.Name = font

        # if selection.Type == 3: #text selected
        #     selection.TextRange2.Font = font
        # elif selection.Type == 2: #shapes selected
        #     shapes = pplib.get_shapes_from_selection(selection)
        #     for textframe in pplib.iterate_shape_textframes(shapes):
        #         textframe.TextRange.Font = font


class AppEventTester(object):
    @staticmethod
    def load():
        logging.debug("bkt load")
    @staticmethod
    def unload():
        logging.debug("bkt unload")

    @staticmethod
    def sld_begin(**kwargs):
        logging.debug("sld_begin: %r"%kwargs)
    @staticmethod
    def sld_end(**kwargs):
        logging.debug("sld_end: %r"%kwargs)


bkt.AppEvents.bkt_load += bkt.Callback(AppEventTester.load)
bkt.AppEvents.bkt_unload += bkt.Callback(AppEventTester.unload)

bkt.AppEvents.slideshow_begin += bkt.Callback(AppEventTester.sld_begin)
bkt.AppEvents.slideshow_end += bkt.Callback(AppEventTester.sld_end)


testfenster_gruppe = bkt.ribbon.Group(
    label='Colordialog Tests',
    image_mso='HappyFace',
    children = [
        bkt.ribbon.Button(
            id = 'colorwindow',
            label="Fenster mit Farbauswahl",
            show_label=True,
            image_mso='HappyFace',
            on_action=bkt.Callback(WindowTests.colorwindow, application=True),
            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            size="large",
        ),
        bkt.ribbon.Button(
            id = 'color_dialog',
            label="Windows Farbdialog",
            show_label=True,
            image_mso='HappyFace',
            on_action=bkt.Callback(WindowTests.open_color_dialog, shape=True),
            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            size="large",
        ),
    ]
)

testformatting_gruppe = bkt.ribbon.Group(
    label='Formatting Tests',
    image_mso='HappyFace',
    children = [
        bkt.ribbon.Button(
            id = 'export_as_png',
            label="Folie als PNG speichern",
            show_label=True,
            image_mso='HappyFace',
            on_action=bkt.Callback(FormattingTestFunctions.export_as_png),
            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            size="large",
        ),
        bkt.ribbon.Button(
            id = 'buttonpos',
            label="Control pos",
            show_label=True,
            image_mso='HappyFace',
            on_action=bkt.Callback(FormattingTestFunctions.control_position),
            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            size="large",
        ),
        bkt.ribbon.Button(
            id = 'shapetypedb',
            label="Shape DB",
            show_label=True,
            image_mso='HappyFace',
            on_action=bkt.Callback(FormattingTestFunctions.build_shape_db),
            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            size="large",
        ),
        bkt.ribbon.Separator(),
        bkt.ribbon.Button(
            id = 'table_formatter',
            label="Tabelle formatieren",
            show_label=True,
            image_mso='HappyFace',
            on_action=bkt.Callback(FormattingTestFunctions.table_formatter),
            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            size="large",
        ),
        bkt.ribbon.ToggleButton(
            id = 'custom_bold',
            label="Custom Fett",
            show_label=True,
            image_mso='Bold',
            get_pressed=bkt.Callback(CustomFontStyles.get_pressed),
            on_toggle_action=bkt.Callback(CustomFontStyles.on_toggle_action),
            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
            size="large",
        ),
    ]
)


def get_content_symbols_test():
    return bkt.ribbon.Menu(
        xmlns="http://schemas.microsoft.com/office/2009/07/customui",
        id=None,
        children = [
            bkt.ribbon.Button(
                label="Sonder&amp;&amp;zeichen &amp;&amp; bla test",
                screentip="sjdfs &amp;&amp;sdf &amp;&amp; sdfsdf",
                supertip="sjdfs &amp;sdf &amp; sdfsdf",
                tag="TAG er&amp;zeichen &amp; bla test",
                on_action=bkt.Callback(lambda current_control: bkt.message("here: %s"%current_control["tag"])),
            ),
        ],
    )

ampersand_gruppe = bkt.ribbon.Group(
    label='Sonderzeichen',
    image_mso='HappyFace',
    children = [
        bkt.ribbon.Button(
            label="Sonder&&zeichen && bla test",
            show_label=True,
            image_mso='HappyFace',
            screentip="sjdfs &&sdf && sdfsdf",
            supertip="sjdfs &sdf & sdfsdf",
            tag="TAG er&zeichen & bla test",
            on_action=bkt.Callback(lambda current_control: bkt.message("here: %s"%current_control["tag"])),
        ),
        bkt.ribbon.Button(
            get_label=bkt.Callback(lambda: "Sonder&&zeichen && bla test"),
            show_label=True,
            image_mso='HappyFace',
            get_screentip=bkt.Callback(lambda: "sjdfs &&sdf && sdfsdf"),
            get_supertip=bkt.Callback(lambda: "sjdfs &sdf & sdfsdf"),
        ),
        bkt.ribbon.DynamicMenu(
            label="Test dynamic",
            show_label=True,
            image_mso='HappyFace',
            get_content=bkt.Callback(get_content_symbols_test),
        )
    ]
)


class ShapesMore(object):
    @classmethod
    def convert_to_architecture_shape(cls, shapes):
        bkt.message("Gewinkelte Konnektoren funktioniert leider nicht mit benutzerdefinierten Shape-Punkten.")
        for shp in shapes:
            cls._convert_to_architecture_shape(shp)
    
    @staticmethod
    def _convert_to_architecture_shape(shape, n=5):
        # from math import atan2
        shape_nodes = algos.get_bounding_nodes(shape)

        #convert shape into freeform by adding and deleting node (not sure if this is required)
        shape.Nodes.Insert(1, 0, 0, 0, 0) #msoSegmentLine, msoEditingAuto, x, y
        shape.Nodes.Delete(2)

        if shape.nodes.count != 5:
            return bkt.message.error("Diese Funktion steht nur bei Rechtecken zur Verfügung.")

        # set nodes (rectangle has 5 nodes as start and end node are the same)
        num_points = n-1
        x_dis = (shape_nodes[3][0] - shape_nodes[0][0]) / num_points
        y_dis = (shape_nodes[1][1] - shape_nodes[0][1]) / num_points

        for i in range(1,num_points):
            f = num_points-i
            shape.nodes.insert(1, 0, 0, shape_nodes[0][0] + x_dis*f, shape_nodes[0][1])
            shape.nodes.insert(2+i, 0, 0, shape_nodes[3][0], shape_nodes[3][1] + y_dis*f)
            shape.nodes.insert(3+2*i, 0, 0, shape_nodes[1][0] + x_dis*f, shape_nodes[1][1])
            shape.nodes.insert(4+3*i, 0, 0, shape_nodes[0][0], shape_nodes[0][1] + y_dis*f)

        # shape.nodes.insert(1+0, 0, 0, shape_nodes[0][0] + x_dis*2, shape_nodes[0][1])
        # shape.nodes.insert(2+1, 0, 0, shape_nodes[3][0], shape_nodes[3][1] + y_dis*2)
        # shape.nodes.insert(3+2, 0, 0, shape_nodes[1][0] + x_dis*2, shape_nodes[1][1])
        # shape.nodes.insert(4+3, 0, 0, shape_nodes[0][0], shape_nodes[0][1] + y_dis*2)

        # shape.nodes.insert(1+0, 0, 0, shape_nodes[0][0] + x_dis*1, shape_nodes[0][1])
        # shape.nodes.insert(2+2, 0, 0, shape_nodes[3][0], shape_nodes[3][1] + y_dis*1)
        # shape.nodes.insert(3+4, 0, 0, shape_nodes[1][0] + x_dis*1, shape_nodes[1][1])
        # shape.nodes.insert(4+6, 0, 0, shape_nodes[0][0], shape_nodes[0][1] + y_dis*1)

architecture_gruppe = bkt.ribbon.Group(
    label='Shapes',
    image_mso='HappyFace',
    children = [
        bkt.ribbon.Button(
            label="3 zusätzliche Punkte je Rechtecktseite",
            size="large",
            image_mso="ObjectEditPoints",
            screentip="Jeder Rechteck-Seite 3 zusätzliche Punkte für Konnektoren hinzufügen",
            supertip="Fügt jeder Seite 3 zusätzliche Shape-Punkte hinzu, die für Konntektoren genutzt werden können. Funktioniert nur mit Rechtecken.",
            on_action=bkt.Callback(ShapesMore.convert_to_architecture_shape),
            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
        ),
    ]
)


bkt.powerpoint.add_tab(bkt.ribbon.Tab(
    id="FlorianTab",
    label='DEV FST',
    children = [
        ampersand_gruppe,
        testfenster_gruppe,
        testformatting_gruppe,
        wpftest.xamltest_gruppe,
        architecture_gruppe,
    ]
))



##############################
### Backstage area testing ###
##############################

bkt.powerpoint.add_backstage_control(
    bkt.ribbon.Tab(
        label="BKT Flo 1",
        title="BKT Florian Test 1",
        columnWidthPercent="30",
        insertAfterMso="TabInfo",
        children=[
            bkt.ribbon.FirstColumn(children=[
                bkt.ribbon.Group(label="testgruppe", children=[
                    bkt.ribbon.PrimaryItem(children=[
                        bkt.ribbon.Menu(
                            label="Test1",
                            image_mso="HappyFace",
                            children=[
                                bkt.ribbon.MenuGroup(
                                    #label="Test1",
                                    item_size="large",
                                    children=[
                                        bkt.ribbon.Button(
                                            label="Test2",
                                            description="Beschreibungstext bla bla",
                                            image_mso="HappyFace",
                                            on_action=bkt.Callback(lambda: True),
                                        )
                                    ]
                                ),
                                bkt.ribbon.MenuGroup(
                                    #label="Test1",
                                    # item_size="large",
                                    children=[
                                        bkt.ribbon.Button(
                                            label="Test3",
                                            image_mso="HappyFace",
                                            on_action=bkt.Callback(lambda: True),
                                        )
                                    ]
                                )
                            ]
                        )
                    ]),
                    bkt.ribbon.TopItems(children=[
                        bkt.ribbon.Label(
                            label="Test TEXT bla bla",
                        ),
                        bkt.ribbon.Label(
                            label="Test TEXT bla bla",
                        ),
                    ]),
                    bkt.ribbon.BottomItems(children=[
                        bkt.ribbon.Button(
                            label="Test2",
                            image_mso="HappyFace",
                            on_action=bkt.Callback(lambda: True),
                        )
                    ])
                ])
            ]),
            bkt.ribbon.SecondColumn(children=[
                bkt.ribbon.Group(label="Testgruppe", children=[
                    bkt.ribbon.TopItems(children=[
                        bkt.ribbon.Button(
                            label="Test1 Close",
                            image_mso="HappyFace",
                            on_action=bkt.Callback(lambda: True),
                            is_definitive=True,
                        )
                    ]),
                    bkt.ribbon.BottomItems(children=[
                        bkt.ribbon.Button(
                            label="Test2",
                            image_mso="HappyFace",
                            on_action=bkt.Callback(lambda: True),
                        )
                    ])
                ]),
                bkt.ribbon.Group(label="Testgruppe 2", children=[
                    bkt.ribbon.TopItems(children=[
                        bkt.ribbon.Button(
                            label="Test1",
                            image_mso="HappyFace",
                            on_action=bkt.Callback(lambda: True),
                        )
                    ]),
                    bkt.ribbon.BottomItems(children=[
                        bkt.ribbon.Button(
                            label="Test2",
                            image_mso="HappyFace",
                            on_action=bkt.Callback(lambda: True),
                        )
                    ])
                ]),
            ])
        ]
    )
)

bkt.powerpoint.add_backstage_control(
    bkt.ribbon.Tab(
        label="BKT Flo 2",
        title="BKT Florian Test 2",
        insertAfterMso="TabInfo",
        firstColumnMinWidth="500",
        firstColumnMaxWidth="500",
        children=[
            bkt.ribbon.FirstColumn(children=[
                bkt.ribbon.TaskFormGroup(label="Task Form Group Label", children=[
                    bkt.ribbon.Category(label="Category Label", children=[

                        bkt.ribbon.Task(label="test 1", description="bla bla", image_mso="HappyFace", children=[
                            bkt.ribbon.Group(label="Testgruppe 1", children=[
                                bkt.ribbon.TopItems(children=[
                                    bkt.ribbon.Button(
                                        label="Test1",
                                        image_mso="HappyFace",
                                        on_action=bkt.Callback(lambda: True),
                                    ),
                                    bkt.ribbon.Button(
                                        label="Test2",
                                        image_mso="HappyFace",
                                        on_action=bkt.Callback(lambda: True),
                                    )
                                ]),
                            ])
                        ]),
                        bkt.ribbon.Task(label="test 2", description="bla bla", image_mso="HappyFace", children=[
                            bkt.ribbon.Group(label="Testgruppe 2", children=[
                                bkt.ribbon.TopItems(children=[
                                    bkt.ribbon.Button(
                                        label="Test3",
                                        image_mso="HappyFace",
                                        on_action=bkt.Callback(lambda: True),
                                    ),
                                    bkt.ribbon.Button(
                                        label="Test4",
                                        image_mso="HappyFace",
                                        on_action=bkt.Callback(lambda: True),
                                    )
                                ]),
                            ])
                        ]),

                    ])
                ])
            ])
        ]
    )
)