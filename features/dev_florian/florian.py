# -*- coding: utf-8 -*-
'''
Created on 2017-07-24
@author: Florian Stallmann
'''

import bkt
import bkt.library.powerpoint as pplib
import bkt.library.algorithms as algos

import logging

# import clr
# clr.AddReference('System.Windows.Forms')
# import System.Windows.Forms as F

F = bkt.dotnet.import_forms()
D = bkt.dotnet.import_drawing()

# import pycountry
# import gettext
# german = gettext.translation('iso3166', pycountry.LOCALES_DIR, languages=['de'])
# german.install()

import os
from collections import OrderedDict

# import quickedit
import wpftest


import ctypes

# import clr
# clr.AddReference("BKT")
# import BKT



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
        testbtn.BackColor = F.Color.LightGreen
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

    @staticmethod
    def colorwindow(application):
        ColorSelectorWindow(application).switchWindow()


class TestsFST(object):
    @staticmethod
    def open_color_dialog(shape):
        cd = F.ColorDialog()
        cd.Color = D.ColorTranslator.FromOle(shape.Fill.ForeColor.RGB)
        cd.FullOpen = True
        if cd.ShowDialog() == F.DialogResult.OK:
            color = D.ColorTranslator.ToOle(cd.Color)
            shape.Fill.ForeColor.RGB = color
            # bkt.helpers.message("Farbe: %r" % color)

    # types_hierarchy = [bkt.ribbon.Tab, bkt.ribbon.Group]
    types_include = [bkt.ribbon.Button, bkt.ribbon.ToggleButton, bkt.ribbon.Gallery, bkt.ribbon.DynamicMenu, bkt.ribbon.EditBox, bkt.ribbon.SpinnerBox]
    types_exclude = [bkt.ribbon.DialogBoxLauncher]
    types_haschildren = [bkt.ribbon.Menu, bkt.ribbon.SplitButton, bkt.ribbon.Box, bkt.ribbon.Gallery]
    types_haslabel    = [bkt.ribbon.Group, bkt.ribbon.Menu, bkt.ribbon.Gallery, bkt.ribbon.DynamicMenu]

    @classmethod
    def gen_overview(cls, context):
        import json
        import io
        import bkt.console

        all_controls = OrderedDict()
        
        def _getattr(object, name, default=None):
            try:
                return object[name]
            except:
                return default

        def _get_screentip_or_label(control):
            return _getattr(control, 'label', _getattr(control, 'screentip'))
        
        def _get_supertip_or_description(control):
            return _getattr(control, 'supertip', _getattr(control, 'description'))

        def _get_image(control):
            return _getattr(control, 'image', _getattr(control, 'image_mso'))
        
        def _add_control(tab_id, child_control, current_location):
            if any(isinstance(child_control, t) for t in cls.types_include) and type(child_control) not in cls.types_exclude:
                c_name = _get_screentip_or_label(child_control)
                #skip controls where label AND screentip are not given, i.e. callback functions
                if c_name:
                    all_controls[tab_id].append({
                        'id':           _getattr(child_control, "id"),
                        'image':        _get_image(child_control),
                        'name':         c_name,
                        'description':  _get_supertip_or_description(child_control),
                        'location':     " > ".join(current_location),
                        'is_standard':  False,
                    })

            elif isinstance(child_control, bkt.ribbon.MSOControl):
                #standard office functionalities
                idmso = _getattr(child_control, "id_mso")
                all_controls[tab_id].append({
                    'id':           idmso,
                    'image':        idmso,
                    'name':         context.app.commandbars.GetScreentipMso(idmso),
                    'description':  context.app.commandbars.GetSupertipMso(idmso) or "-",
                    'location':     " > ".join(current_location),
                    'is_standard':  True,
                })

            if any(isinstance(child_control, t) for t in cls.types_haschildren):
                _iterate_over_children(tab_id, child_control, current_location)
        
        def _iterate_over_children(tab_id, control, current_location):
            if any(isinstance(control, t) for t in cls.types_haslabel):
                current_location = current_location + [_getattr(control, 'label', _getattr(control, 'screentip'))]
            
            if isinstance(control, bkt.ribbon.SpinnerBox):
                all_controls[tab_id].append({
                    'id':           _getattr(control, "id"),
                    'image':        _get_image(control.txt_box),
                    'name':         _get_screentip_or_label(control.txt_box),
                    'description':  _get_supertip_or_description(control.txt_box),
                    'location':     " > ".join(current_location),
                    'is_standard':  False,
                })
                
                if control.image_element:
                    _add_control(tab_id, control.image_element, current_location)
                
            else:
                for child_control in control.children:
                    _add_control(tab_id, child_control, current_location)

        #TODO: iterate over context.python_addin.app_ui.tabs and its children
        # for c_id, control in context.python_addin.callback_manager.ribbon_controls.iteritems():
        for _, tab in context.python_addin.app_ui.tabs.iteritems():
            tab_label = _getattr(tab, "label")
            if tab_label is None:
                continue
            all_controls[tab_label] = []
            for group in tab.children:
                try:
                    _iterate_over_children(tab_label, group, [])
                except:
                    pass
        
        #TODO: add context menu functions
        #TODO: add backstage functions
        
        file = os.path.join(os.path.dirname(__file__), "all_controls.json")
        with io.open(file, 'w', encoding='utf-8') as json_file:
            # bkt.console.show_message(json.dumps(all_controls, ensure_ascii=False))
            json.dump(all_controls, json_file, ensure_ascii=False, indent=2)
        
        file2 = os.path.join(os.path.dirname(__file__), "all_controls.md")
        with io.open(file2, 'w', encoding='utf-8') as md_file:
            for tab_label, controls in all_controls.iteritems():
                if len(controls) == 0:
                    continue
                md_file.write("## {}\n".format(tab_label))
                md_file.write("| {:50} | {:50} | {:50} |\n".format("Name", "Beschreibung", "Gruppe > Untermenü(s)"))
                md_file.write("| {:-<50} | {:-<50} | {:-<50} |\n".format("-", "-", "-"))
                for control in controls:
                    if not control["name"] and not control["description"]:
                        continue
                    if control["description"]:
                        control["description"] = control["description"].replace("\n", "<br>")
                    if control["is_standard"]:
                        control["name"] = "*{}*".format(control["name"])
                    md_file.write("| {name:50} | {description:50} | {location:50} |\n".format(**control))
                md_file.write("\n\n")
        

    @staticmethod
    def export_as_png(presentation, slides):
        import System
        for slide in slides:
            shape_indices = []
            shape_index = 1
            for shape in slide.shapes:
                if shape.type != 14 and shape.visible == -1:
                    # shape is not a placeholder and visible
                    shape_indices.append(shape_index)
                shape_index+=1
            # select shapes
            shape_range = slide.shapes.Range(System.Array[int](shape_indices))
            path = presentation.Path + "\\" + str(slide.SlideIndex) + ".png"
            shape_range.Export(path, 2) #2=ppShapeFormatPNG

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
    
    @staticmethod
    def hook_events():
        import mousehook
        mousehook.mouse_msg_loop()
        # HC_ACTION = 0
        # WH_MOUSE_LL = 14
        # WM_LBUTTONDOWN = 0x0201

        # def my_callback(nCode, wParam, lParam):
        #     if nCode == HC_ACTION and wParam == WM_LBUTTONDOWN:
        #         print("hier i am mouse")
        #     return ctypes.windll.user32.CallNextHookEx(None, nCode, wParam, lParam)
        
        # hook = ctypes.windll.user32.SetWindowsHookExW(WH_MOUSE_LL, my_callback, None, 0)

    current_control = None
    customui_control = None
    @classmethod
    def control_position(cls, context, current_control, customui_control):
        cls.current_control = current_control
        cls.customui_control = customui_control
        print(current_control.id)
        try:
            print("test1: %s" % context.app.CommandBars.ActionControl)
        except:
            pass
        try:
            print("test2: %s" % context.app.Caller)
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


bkt.AppEvents.bkt_load += bkt.Callback(AppEventTester.load)
bkt.AppEvents.bkt_unload += bkt.Callback(AppEventTester.unload)



testfenster_gruppe = bkt.ribbon.Group(
    label='Tests FST',
    image_mso='HappyFace',
    children = [
        bkt.ribbon.Button(
            id = 'export_as_png',
            label="Folie als PNG speichern",
            show_label=True,
            image_mso='HappyFace',
            supertip="XXX",
            on_action=bkt.Callback(TestsFST.export_as_png),
            # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            size="large",
        ),
        bkt.ribbon.Separator(),
        bkt.ribbon.Button(
            id = 'colorwindow',
            label="Fenster mit Farbauswahl",
            show_label=True,
            image_mso='HappyFace',
            supertip="XXX",
            on_action=bkt.Callback(ColorSelectorWindow.colorwindow, application=True),
            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
        ),
        bkt.ribbon.Button(
            id = 'color_dialog',
            label="Windows Farbdialog",
            show_label=True,
            image_mso='HappyFace',
            supertip="XXX",
            on_action=bkt.Callback(TestsFST.open_color_dialog, shape=True),
            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
        ),
        bkt.ribbon.Button(
            id = 'generate_overview',
            label="Übersicht generieren",
            show_label=True,
            image_mso='HappyFace',
            supertip="XXX",
            on_action=bkt.Callback(TestsFST.gen_overview, context=True),
            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
        ),
        bkt.ribbon.Button(
            id = 'table_formatter',
            label="Tabelle formatieren",
            show_label=True,
            image_mso='HappyFace',
            supertip="XXX",
            on_action=bkt.Callback(TestsFST.table_formatter),
            # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
        ),
        bkt.ribbon.Button(
            id = 'hook_mouse',
            label="Maushook",
            show_label=True,
            image_mso='HappyFace',
            supertip="XXX",
            on_action=bkt.Callback(TestsFST.hook_events),
            # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
        ),
        bkt.ribbon.Button(
            id = 'buttonpos',
            label="Control pos",
            show_label=True,
            image_mso='HappyFace',
            supertip="XXX",
            on_action=bkt.Callback(TestsFST.control_position),
            # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
        ),
        bkt.ribbon.ToggleButton(
            id = 'custom_bold',
            label="Custom Fett",
            show_label=True,
            image_mso='Bold',
            get_pressed=bkt.Callback(CustomFontStyles.get_pressed),
            on_toggle_action=bkt.Callback(CustomFontStyles.on_toggle_action),
            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
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
                on_action=bkt.Callback(lambda current_control: bkt.helpers.message("here: %s"%current_control["tag"])),
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
            on_action=bkt.Callback(lambda current_control: bkt.helpers.message("here: %s"%current_control["tag"])),
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

bkt.powerpoint.add_tab(bkt.ribbon.Tab(
    id="FlorianTab",
    label=u'DEV FST',
    children = [
        ampersand_gruppe,
        testfenster_gruppe,
        wpftest.xamltest_gruppe,
        # quickedit.color_selector_gruppe,
    ]
))

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