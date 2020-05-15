# -*- coding: utf-8 -*-
'''
Created on 29.07.2015

@author: rdebeerst
'''

from __future__ import absolute_import

import unittest

import bkt
# import bkt.addin
import bkt.appui
import bkt.ribbon

from bkt.xml import RibbonXMLFactory
from bkt.callbacks import CallbackTypes, Callback
# from bkt.apps import ApplicationRibbonInformation

bkt.helpers.settings = bkt.settings = dict()
RibbonXMLFactory.namespace = ""

def ctrl_to_str(ctrl):
    return RibbonXMLFactory.to_normalized_string(ctrl.xml()).strip()


class UIControlTest(unittest.TestCase):
    
    def test_uicontrol_args(self):
        bkt.ribbon.RibbonControl.no_id = True
        ctrl = bkt.ribbon.RibbonControl('button')
        self.assertEqual(ctrl_to_str(ctrl), '<button />')
        
        ctrl = bkt.ribbon.RibbonControl('button', label='test label')
        self.assertEqual(ctrl_to_str(ctrl), '<button label="test label" />')

        ctrl = bkt.ribbon.RibbonControl('button')
        ctrl['label'] = 'test label'
        self.assertEqual(ctrl_to_str(ctrl), '<button label="test label" />')
        
        # FIXME: ist das so gewollt? Oder eher rausfiltern?
        ctrl = bkt.ribbon.RibbonControl('button')
        ctrl['_label'] = 'test label'
        self.assertEqual(ctrl_to_str(ctrl), '<button Label="test label" />')
    
    def test_uicontrol_arg_cases(self):
        bkt.ribbon.RibbonControl.no_id = True
        ctrl = bkt.ribbon.RibbonControl('button')
        ctrl['button_size'] = 'large'
        self.assertEqual(ctrl_to_str(ctrl), '<button buttonSize="large" />')
        
        ctrl = bkt.ribbon.RibbonControl('button')
        ctrl['buttonSize'] = 'large'
        self.assertEqual(ctrl_to_str(ctrl), '<button buttonSize="large" />')
        
        ctrl = bkt.ribbon.RibbonControl('button')
        ctrl['ButtonSize'] = 'large'
        self.assertEqual(ctrl_to_str(ctrl), '<button buttonSize="large" />')
        
    
    def test_uicontrol_ids(self):
        bkt.ribbon.RibbonControl.no_id = False
        ctrl = bkt.ribbon.RibbonControl('button')
        ctrl.set_id('my_id')
        self.assertEqual(ctrl_to_str(ctrl), '<button id="my_id" />')
        
        ctrl = bkt.ribbon.RibbonControl('button', id_mso='msoid')
        self.assertEqual(ctrl_to_str(ctrl), '<button idMso="msoid" />')
        ctrl.set_id('my_id')
        self.assertEqual(ctrl_to_str(ctrl), '<button idMso="msoid" />')
        
        
        ctrl = bkt.ribbon.RibbonControl('button')
        ctrl.uuid = 'some-long-uuid'
        # fallback_id is unused
        ctrl.set_id('my_id')
        self.assertEqual(ctrl_to_str(ctrl), '<button id="_some-long-uuid" />')
        ctrl.set_id(fallback_id='my_id')
        self.assertEqual(ctrl_to_str(ctrl), '<button id="_some-long-uuid" />')
        # uuid appended to ribbon-short-id 
        ctrl.set_id(ribbon_short_id='my_id')
        self.assertEqual(ctrl_to_str(ctrl), '<button id="my_id__some-long-uuid" />')
        # idtag appended at end
        ctrl.set_id(ribbon_short_id='my_id', id_tag='test-tag')
        self.assertEqual(ctrl_to_str(ctrl), '<button id="my_id__some-long-uuid_test-tag" />')
        
        
    def test_uicontrol_creator(self):
        bkt.ribbon.RibbonControl.no_id = True
        x = bkt.ribbon.create_ribbon_control_class('group')
        ctrl = x()
        self.assertEqual(ctrl_to_str(ctrl), '<group />')
        
        
    def test_uicontrol_no_id(self):
        bkt.ribbon.RibbonControl.no_id = False
        x = bkt.ribbon.create_ribbon_control_class('tab')
        x.no_id = True
        
        ctrl = x()
        ctrl.set_id('my_id')
        self.assertEqual(ctrl_to_str(ctrl), '<tab />')
        
        x.no_id = False
        ctrl = x()
        ctrl.set_id('my_id')
        self.assertEqual(ctrl_to_str(ctrl), '<tab id="my_id" />')
        
        # FIXME: ist das so gewollt?
        ctrl = x(id='some-id')
        self.assertEqual(ctrl_to_str(ctrl), '<tab id="some-id" />')
    
    def test_grouped_controls(self):
        bkt.ribbon.RibbonControl.no_id = False
        btn = bkt.ribbon.create_ribbon_control_class('button')(id='some-button', label='first button')
        grp = bkt.ribbon.create_ribbon_control_class('group')(id='grp-id', label='first group', children=[btn,btn])
        self.assertEqual(ctrl_to_str(grp), 
"""<group id="grp-id" label="first group">
<button id="some-button" label="first button" />
<button id="some-button" label="first button" />
</group>""")
        
    
    def test_default_uicontrols(self):
        bkt.ribbon.RibbonControl.no_id = True
        ctrl = bkt.ribbon.Button()
        self.assertEqual(ctrl_to_str(ctrl), '<button />')
        ctrl = bkt.ribbon.ToggleButton()
        self.assertEqual(ctrl_to_str(ctrl), '<toggleButton />')
        ctrl = bkt.ribbon.EditBox()
        self.assertEqual(ctrl_to_str(ctrl), '<editBox />')
        ctrl = bkt.ribbon.Gallery()
        self.assertEqual(ctrl_to_str(ctrl), '<gallery />')

        ctrl = bkt.ribbon.Box()
        self.assertEqual(ctrl_to_str(ctrl), '<box />')
        ctrl = bkt.ribbon.Group()
        self.assertEqual(ctrl_to_str(ctrl), '<group />')
        
        ctrl = bkt.ribbon.Tab()
        self.assertEqual(ctrl_to_str(ctrl), '<tab />')
        ctrl = bkt.ribbon.Tabs()
        self.assertEqual(ctrl_to_str(ctrl), '<tabs />')
        ctrl = bkt.ribbon.Ribbon()
        self.assertEqual(ctrl_to_str(ctrl), '<ribbon />')
    
    
    def test_spinner_box(self):
        bkt.ribbon.RibbonControl.no_id = True
        ctrl = bkt.ribbon.SpinnerBox()
        self.assertEqual(ctrl_to_str(ctrl), u'<box>\n<editBox sizeString="####" />\n<button label="\xab" />\n<button label="\xbb" />\n</box>')
        
        bkt.ribbon.RibbonControl.no_id = False
        ctrl = bkt.ribbon.SpinnerBox()
        ctrl.set_id('test-id')
        self.assertEqual(ctrl_to_str(ctrl), u'<box id="test-id">\n<editBox id="test-id_text" sizeString="####" />\n<button id="test-id_decrement" label="\xab" />\n<button id="test-id_increment" label="\xbb" />\n</box>')
        
        ctrl = bkt.ribbon.SpinnerBox(id='test-id')
        self.assertEqual(ctrl_to_str(ctrl), u'<box id="test-id">\n<editBox id="test-id_text" sizeString="####" />\n<button id="test-id_decrement" label="\xab" />\n<button id="test-id_increment" label="\xbb" />\n</box>')
        
        
        def on_change():
            return 'onchange'
        def get_text():
            return 'gettext'
        def increment():
            return 'increment'
        def decrement():
            return 'decrement'
        
        ctrl = bkt.ribbon.SpinnerBox(id='test-id')
        ctrl.add_callback(Callback(on_change, CallbackTypes.on_change))
        ctrl.add_callback(Callback(get_text, CallbackTypes.get_text))
        ctrl.add_callback(Callback(increment, CallbackTypes.increment))
        ctrl.add_callback(Callback(decrement, CallbackTypes.decrement))
        
        callbacks = ctrl.collect_callbacks()
        lst = { (cb.control.id, cb.callback_type):cb.method  for cb in callbacks}
        
        # returns a list of Callbacks
        self.assertEqual(lst[('test-id_text', CallbackTypes.on_change )](), 'onchange')
        self.assertEqual(lst[('test-id_text', CallbackTypes.get_text )](), 'gettext')
        self.assertEqual(lst[('test-id_increment', CallbackTypes.on_action )](), 'increment')
        self.assertEqual(lst[('test-id_decrement', CallbackTypes.on_action )](), 'decrement')
        
        
        
    def test_color_gallery(self):
        bkt.ribbon.RibbonControl.no_id = True
        self.maxDiff = None
        ctrl = bkt.ribbon.ColorGallery()
        self.assertEqual(ctrl_to_str(ctrl), u'<gallery columns="10" getItemCount="PythonGetItemCount" getItemImage="PythonGetItemImage" getItemLabel="PythonGetItemLabel" imageMso="SmartArtChangeColorsGallery" itemHeight="14" itemWidth="14" onAction="PythonOnActionIndexed" showItemLabel="false" />')
        
        def set_rgb(color):
            return 'set rgb: ' + str(color)
        def set_theme(color_index, brightness):
            return 'set theme: ' + str(color_index)+ '/' + str(brightness)
            
        ctrl.add_callback(Callback(set_rgb, CallbackTypes.on_rgb_color_change))
        self.assertEqual(ctrl_to_str(ctrl), u'<gallery columns="10" getItemCount="PythonGetItemCount" getItemImage="PythonGetItemImage" getItemLabel="PythonGetItemLabel" imageMso="SmartArtChangeColorsGallery" itemHeight="14" itemWidth="14" onAction="PythonOnActionIndexed" showItemLabel="false" />')
        
        callbacks = ctrl.collect_callbacks()
        lst = { cb.callback_type:cb.method  for cb in callbacks}
        self.assertEqual( lst[CallbackTypes.on_action_indexed](selected_item=1, index=60, context=None), 'set rgb: 0')
        self.assertEqual( lst[CallbackTypes.on_action_indexed](selected_item=1, index=1, context=None), 'set rgb: 0')
        ctrl.add_callback(Callback(set_theme, CallbackTypes.on_theme_color_change))
        self.assertEqual( lst[CallbackTypes.on_action_indexed](selected_item=1, index=60, context=None), 'set rgb: 0')
        self.assertEqual( lst[CallbackTypes.on_action_indexed](selected_item=1, index=1, context=None), 'set theme: 0/0')
        
    
    
    def test_tab_definition(self):
        bkt.ribbon.RibbonControl.no_id = True
        
        # myofficeapp = ApplicationRibbonInformation('myofficeapp', 'app')
        myofficeapp = bkt.appui.AppUIs.get_app_ui("MyOfficeApp")
        
        myofficeapp.add_tab(bkt.ribbon.Tab(label="TestTab", children=[
            bkt.ribbon.Group(label="TestGroup", children=[
                bkt.ribbon.Button(label="test button")
            ])
        ]))
        
        # ctrl = bkt.addin.AddinCustomUI('myofficeapp').create_base_control(ribbon_info=myofficeapp)
        ctrl = myofficeapp.get_customui_control()
        # ctrl is customUI-control
        # TODO: check just ribbon-child of customUI, not customUI attributes
        self.assertEqual(ctrl_to_str(ctrl), u'<customUI loadImage="PythonLoadImage" onLoad="PythonOnRibbonLoad" {http://www.w3.org/2000/xmlns/}nsBKT="http://www.business-kasper-toolbox.com/toolbox">\n<ribbon startFromScratch="false">\n<tabs>\n<tab label="TestTab">\n<group label="TestGroup">\n<button label="test button" />\n</group>\n</tab>\n</tabs>\n</ribbon>\n</customUI>')
    
        
        
        

class LargeButton(bkt.ribbon.Button):
    _attributes = dict(
        size = "large"
    )

class LabledLargeButton(LargeButton):
    _attributes = dict(
        label = "mylabel"
    )

class ButtonWithAction(bkt.ribbon.Button):
    def on_action(self):
        print "this will do something"

class UIControlSubclassingTest(unittest.TestCase):
                
    def test_ui_control_subclassing(self):
        bkt.ribbon.RibbonControl.no_id = True
        llb = LabledLargeButton()
        self.assertEqual(ctrl_to_str(llb), u'<button label="mylabel" size="large" />')
        
    def test_ui_control_auto_callbacks(self):
        bkt.ribbon.RibbonControl.no_id = True
        b = ButtonWithAction()
        self.assertEqual(ctrl_to_str(b), u'<button onAction="PythonOnAction" />')
        
        
        
        
        
        
        
        
