# -*- coding: utf-8 -*-
'''
Created on 07.08.2015

@author: rdebeerst
'''



import unittest

import bkt
import bkt.ribbon

from bkt.xml import RibbonXMLFactory
from bkt.callbacks import CallbackType, CallbackTypes, Callback, CallbackLazy, InvocationContext

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



class ModelMock(object):
    @staticmethod
    def on_action():
        return 'onaction'

    @classmethod
    def on_change(cls, context, shape, non_existing_parameter=None):
        return 'nochange'


class InvocationContextTest(unittest.TestCase):
    def test_context_init(self):
        context = InvocationContext()

        self.assertFalse(context.shape)
        self.assertFalse(context.context)
        self.assertFalse(context.slide)
    
    def test_context_from_method(self):
        context = InvocationContext.from_method(ModelMock.on_change)

        self.assertTrue(context.shape)
        self.assertTrue(context.context)
        self.assertFalse(context.slide)


class CallbackTest(unittest.TestCase):
    def test_callback_init_method1(self):
        cb = Callback(ModelMock.on_action, CallbackTypes.on_action, InvocationContext())

        self.assertEqual(cb.method, ModelMock.on_action)
        self.assertEqual(cb.method(), 'onaction')
        self.assertEqual(cb.callback_type, CallbackTypes.on_action)

        self.assertTrue(cb.is_transactional)
        self.assertFalse(cb.is_cacheable)

    def test_callback_init_method2(self):
        cb = Callback(ModelMock, "on_action", CallbackTypes.on_action, InvocationContext())

        self.assertEqual(cb.method, ModelMock.on_action)
        self.assertEqual(cb.callback_type, CallbackTypes.on_action)
        self.assertTrue(cb.is_transactional)
        self.assertFalse(cb.is_cacheable)

    def test_callback_init_method2_kwargs_copy(self):
        cb = Callback(ModelMock, "on_action", CallbackTypes.on_action, InvocationContext(), transactional=False, cacheable=True)
        cb_copy = cb.copy()

        self.assertEqual(cb_copy.method, ModelMock.on_action)
        self.assertEqual(cb_copy.method(), 'onaction')
        self.assertEqual(cb_copy.callback_type, CallbackTypes.on_action)

        self.assertFalse(cb_copy.is_transactional)
        self.assertTrue(cb_copy.is_cacheable)

    def test_callback_init_method2_cbtype_auto(self):
        cb = Callback(ModelMock, "on_action", None, InvocationContext())

        self.assertEqual(cb.callback_type, CallbackTypes.on_action)

    def test_callback_init_method3(self):
        cb = Callback(ModelMock.on_change, CallbackTypes.on_change, shape=True, context=True)

        self.assertEqual(cb.method, ModelMock.on_change)
        self.assertEqual(cb.callback_type, CallbackTypes.on_change)
        self.assertTrue(cb.is_transactional)
        self.assertFalse(cb.is_cacheable)

        self.assertTrue(cb.invocation_context.shape)
        self.assertTrue(cb.invocation_context.context)

    def test_callback_init_method3_kwargs(self):
        cb = Callback(ModelMock.on_change, CallbackTypes.on_change, shape=True, context=True, transactional=False, cacheable=True)

        self.assertFalse(cb.is_transactional)
        self.assertTrue(cb.is_cacheable)

    def test_callback_init_method4(self):
        cb = Callback(ModelMock.on_change, CallbackTypes.on_change)

        self.assertEqual(cb.method, ModelMock.on_change)
        self.assertEqual(cb.callback_type, CallbackTypes.on_change)
        self.assertTrue(cb.is_transactional)
        self.assertFalse(cb.is_cacheable)

        self.assertTrue(cb.invocation_context.shape)
        self.assertTrue(cb.invocation_context.context)

    def test_callback_init_method4_kwargs(self):
        cb = Callback(ModelMock.on_change, CallbackTypes.on_change, transactional=False, cacheable=True)

        self.assertFalse(cb.is_transactional)
        self.assertTrue(cb.is_cacheable)

    def test_callback_init_method5(self):
        cb = Callback(ModelMock.on_change)

        self.assertEqual(cb.method, ModelMock.on_change)
        self.assertEqual(cb.callback_type, None)

        self.assertTrue(cb.invocation_context.shape)
        self.assertTrue(cb.invocation_context.context)

        cb.set_callback_type(CallbackTypes.on_action)
        self.assertEqual(cb.callback_type, CallbackTypes.on_action)

        self.assertTrue(cb.is_transactional) #default for on_action
        self.assertFalse(cb.is_cacheable) #default for on_action


    def test_callback_init_method5_kwargs(self):
        cb = Callback(ModelMock.on_change, shape=True, context=True, transactional=False, cacheable=True)

        self.assertEqual(cb.method, ModelMock.on_change)
        self.assertEqual(cb.callback_type, None)

        self.assertTrue(cb.invocation_context.shape)
        self.assertTrue(cb.invocation_context.context)

        cb.set_callback_type(CallbackTypes.on_action)
        self.assertEqual(cb.callback_type, CallbackTypes.on_action)

        self.assertFalse(cb.is_transactional)
        self.assertTrue(cb.is_cacheable)


    def test_callback_init_method5_kwargs_copy(self):
        cb = Callback(ModelMock.on_change, shape=True, context=True, transactional=False, cacheable=True)
        cb_copy = cb.copy()

        self.assertEqual(cb_copy.method, ModelMock.on_change)
        self.assertEqual(cb_copy.callback_type, None)

        self.assertTrue(cb_copy.invocation_context.shape)
        self.assertTrue(cb_copy.invocation_context.context)

        cb_copy.set_callback_type(CallbackTypes.on_action)
        self.assertEqual(cb_copy.callback_type, CallbackTypes.on_action)

        self.assertFalse(cb_copy.is_transactional)
        self.assertTrue(cb_copy.is_cacheable)



class CallbackLazyTest(unittest.TestCase):
    def test_callbacklazy_init_method1(self):
        cb = CallbackLazy("tests.mock_callback", "MockCallback", "on_action")
        cb.set_callback_type(CallbackTypes.on_action)
        
        self.assertEqual(cb.method(), 'onaction')

        self.assertTrue(cb.is_transactional) #default for on_action
        self.assertFalse(cb.is_cacheable) #default for on_action
    
    def test_callbacklazy_init_method1_kwargs(self):
        cb = CallbackLazy("tests.mock_callback", "MockCallback", "on_action", shape=True, transactional=False, cacheable=True)
        cb.set_callback_type(CallbackTypes.on_action)

        self.assertTrue(cb.invocation_context.shape)
        self.assertEqual(cb.method(), 'onaction')

        self.assertFalse(cb.is_transactional)
        self.assertTrue(cb.is_cacheable)
    
    def test_callbacklazy_init_method1_kwargs(self):
        cb = CallbackLazy("tests.mock_callback", "MockCallback", "on_action", shape=True, transactional=False, cacheable=True)
        cb.set_callback_type(CallbackTypes.on_action)
        cb_copy = cb.copy()

        self.assertTrue(cb_copy.invocation_context.shape)
        self.assertEqual(cb_copy.method(), 'onaction')
        self.assertEqual(cb_copy.callback_type, CallbackTypes.on_action)

        self.assertFalse(cb_copy.is_transactional)
        self.assertTrue(cb_copy.is_cacheable)

        
    def test_callbacklazy_init_method2(self):
        cb = CallbackLazy("tests.mock_callback", "do_something")
        
        self.assertEqual(cb.method(), 'dosomething')
    
    def test_callbacklazy_init_method2_kwargs(self):
        cb = CallbackLazy("tests.mock_callback", "do_something", shape=True, transactional=False, cacheable=True)
        cb.set_callback_type(CallbackTypes.on_action)

        self.assertTrue(cb.invocation_context.shape)
        self.assertEqual(cb.method(), 'dosomething')

        self.assertFalse(cb.is_transactional)
        self.assertTrue(cb.is_cacheable)
    
    def test_callbacklazy_init_method2_kwargs(self):
        cb = CallbackLazy("tests.mock_callback", "do_something", shape=True, transactional=False, cacheable=True)
        cb.set_callback_type(CallbackTypes.on_action)
        cb_copy = cb.copy()

        self.assertTrue(cb_copy.invocation_context.shape)
        self.assertEqual(cb_copy.method(), 'dosomething')
        self.assertEqual(cb_copy.callback_type, CallbackTypes.on_action)

        self.assertFalse(cb_copy.is_transactional)
        self.assertTrue(cb_copy.is_cacheable)