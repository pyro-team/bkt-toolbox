# -*- coding: utf-8 -*-
'''
Created on 18.08.2015

@author: rdebeerst
'''

import bkt
import bkt.ribbon
import unittest
from bkt.xml import RibbonXMLFactory


from bkt.callbacks import CallbackTypes
import bkt.factory
# from bkt.ribbon import Box, EditBox, Button
from bkt.ribbon import Box
from test_uicontrol_definition import WorkingSpinner


RibbonXMLFactory.namespace = ""

def ctrl_to_str(ctrl):
    return RibbonXMLFactory.to_normalized_string(ctrl.xml()).strip()



@bkt.configure(label='Dummy Group')
@bkt.group
class DummyGroup(bkt.FeatureContainer):
    
    @bkt.button
    def first_button(self):
        return 'first button'


@bkt.configure(label='Dummy Button')
@bkt.decorators.UIControlAnnotationCommand('button') 
class DummyButton(bkt.FeatureContainer):
    @bkt.callback_type(CallbackTypes.on_action)
    def on_action(self):
        return 'dummy action'


@bkt.configure(label='Dummy Spinner')
@bkt.decorators.UIControlAnnotationCommand(bkt.ribbon.SpinnerBox) 
class DummySpinner(bkt.FeatureContainer):
    @bkt.callback_type(CallbackTypes.on_change)
    def on_change(self):
        return 'dummy change spinner'
    
    @bkt.callback_type(CallbackTypes.get_text)
    def get_text(self):
        return 'dummy get text'
    
    @bkt.callback_type(CallbackTypes.increment)
    def increment(self):
        return 'inc +1'

    @bkt.callback_type(CallbackTypes.decrement)
    def decrement(self):
        return 'dec -1'
    





    



class FeatureContainerTest(unittest.TestCase):

    def test_button(self):
        f = bkt.factory.ControlFactory(DummyButton)
        btn = f.create_control()
        self.assertEqual(ctrl_to_str(btn), '<button id="DummyButton" label="Dummy Button" onAction="PythonOnAction" />')
        
        callbacks = btn.collect_callbacks()
        # returns a list of Callbacks
        self.assertEqual(len(callbacks), 1)
        self.assertEqual(type(callbacks[0]), bkt.callbacks.Callback)
        # Callback-Type and method are as expected
        self.assertEqual(callbacks[0].callback_type, CallbackTypes.on_action)
        #self.assertEqual(callbacks[0].method.__func__, DummyButton.on_action.__func__)
        self.assertEqual(callbacks[0].method(), 'dummy action')
        
    
    
    def test_group(self):
        f = bkt.factory.ControlFactory(DummyGroup)
        grp = f.create_control()
        self.assertEqual(ctrl_to_str(grp), '<group id="DummyGroup" label="Dummy Group">\n<button id="first_button" onAction="PythonOnAction" />\n</group>')
        
        callbacks = grp.collect_callbacks()
        # returns a list of Callbacks
        self.assertEqual(len(callbacks), 1)
        self.assertEqual(type(callbacks[0]), bkt.callbacks.Callback)
        # Callback-Type and method are as expected
        self.assertEqual(callbacks[0].callback_type, CallbackTypes.on_action)
        #self.assertEqual(callbacks[0].method.__func__, DummyButton.on_action.__func__)
        self.assertEqual(callbacks[0].method(), 'first button')

        
        
    def test_spinner(self):
        #self.maxDiff = None
        f = bkt.factory.ControlFactory(DummySpinner)
        ctrl = f.create_control()
        #self.assertEqual(ctrl_to_str(ctrl), u'<box id="DummySpinner">\n<editBox onChange="PythonOnChange" sizeString="###" label="Dummy Spinner" />\n<button label="\xab" />\n<button label="\xbb" />\n</box>')
        self.assertEqual(ctrl_to_str(ctrl), '<box id="DummySpinner">\n<editBox getText="PythonGetText" id="DummySpinner_text" label="Dummy Spinner" onChange="PythonOnChange" sizeString="####" />\n<button id="DummySpinner_decrement" label="«" onAction="PythonOnAction" />\n<button id="DummySpinner_increment" label="»" onAction="PythonOnAction" />\n</box>')
        
        callbacks = ctrl.collect_callbacks()
        lst = { (cb.control.id, cb.callback_type):cb.method  for cb in callbacks}
        
        # returns a list of Callbacks
        self.assertEqual(len(lst), 4)
        self.assertEqual(lst[('DummySpinner_text', CallbackTypes.on_change )](), 'dummy change spinner')
        self.assertEqual(lst[('DummySpinner_text', CallbackTypes.get_text )](), 'dummy get text')
        self.assertEqual(lst[('DummySpinner_increment', CallbackTypes.on_action )](), 'inc +1')
        self.assertEqual(lst[('DummySpinner_decrement', CallbackTypes.on_action )](), 'dec -1')
        
        
        
        
        
        
        
        
