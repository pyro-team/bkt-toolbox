# -*- coding: utf-8 -*-
'''
Created on 07.08.2015

@author: rdebeerst
'''

import unittest
import bkt
import bkt.ribbon
from bkt.xml import RibbonXMLFactory
from bkt.callbacks import CallbackType, CallbackTypes

def ctrl_to_str(ctrl):
    return RibbonXMLFactory.to_string(ctrl.xml()).strip()





class CallbackTypesTest(unittest.TestCase):
    
    def test_callback_type_naming_convention(self):
        cbtype = CallbackType(python_name='on_action')
        self.assertEqual(cbtype.xml_name, 'onAction') 
        self.assertEqual(cbtype.dotnet_name, 'PythonOnAction') 
        self.assertEqual(cbtype.xml(), 'PythonOnAction') 
        
    def test_callback_type_custom_names(self):
        cbtype = CallbackType(xml_name = 'myName')
        self.assertEqual(cbtype.xml_name, 'myName') 
        self.assertEqual(cbtype.dotnet_name, None) 
        self.assertEqual(cbtype.python_name, None) 
        
    def test_callback_type_set_attr_name(self):
        cbtype = CallbackType()
        cbtype.set_attribute('on_action')
        self.assertEqual(cbtype.python_name, 'on_action') 
        self.assertEqual(cbtype.xml_name, 'onAction') 
        self.assertEqual(cbtype.dotnet_name, 'PythonOnAction') 
        
    def test_callback_type_set_attr_name_custom(self):
        cbtype = CallbackType(xml_name='myName')
        cbtype.set_attribute('on_action')
        self.assertEqual(cbtype.python_name, 'on_action') 
        self.assertEqual(cbtype.xml_name, 'myName') 
        self.assertEqual(cbtype.dotnet_name, 'PythonOnAction') 
    
    def test_callback_types(self):
        CallbackTypes.my_type = CallbackType()
        cbtype = CallbackTypes.my_type
        self.assertEqual(cbtype.python_name, 'my_type') 
        self.assertEqual(cbtype.xml_name, 'myType') 

    def test_callback_types_custom_xml_name(self):
        CallbackTypes.my_type = CallbackType(xml_name = 'myName')
        cbtype = CallbackTypes.my_type
        self.assertEqual(cbtype.python_name, 'my_type') 
        self.assertEqual(cbtype.xml_name, 'myName') 
        
    def test_callback_types_custom_name(self):
        CallbackTypes.my_type = CallbackType(python_name='on_action')
        cbtype = CallbackTypes.my_type
        self.assertEqual(cbtype.python_name, 'on_action') 
        self.assertEqual(cbtype.xml_name, 'onAction') 
        
        CallbackTypes.my_type = CallbackType(python_name='on_action', xml_name='xmlName')
        cbtype = CallbackTypes.my_type
        self.assertEqual(cbtype.python_name, 'on_action') 
        self.assertEqual(cbtype.xml_name, 'xmlName') 
    
    def test_callback_types_custom(self):
        cbtype = CallbackTypes.not_defined
        self.assertEqual(cbtype.python_name, 'not_defined') 
        
        cbtype2 = CallbackTypes.not_defined
        self.assertEqual(cbtype, cbtype2)
    
        
        
        
        
