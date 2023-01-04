# -*- coding: utf-8 -*-
'''
Created on 07.08.2015

@author: rdebeerst
'''



import unittest

import bkt
import bkt.ribbon

from bkt.xml import RibbonXMLFactory
from bkt.callbacks import CallbackTypes, Callback

RibbonXMLFactory.namespace = ""

def ctrl_to_str(ctrl):
    return RibbonXMLFactory.to_normalized_string(ctrl.xml()).strip()


# define a new control
from bkt.ribbon import Box, EditBox, Button

class SpinnerBox(Box):
    #
    def __init__(self, **kwargs):
        self.txt_box = EditBox(size_string='###', **kwargs)
        self.inc_button = Button(label="»")
        self.dec_button = Button(label="«")
        super(SpinnerBox, self).__init__(children = [self.txt_box, self.dec_button, self.inc_button])
        #self._local_callbacks = {}
        #
    
    def add_callback(self, rc):
        if rc.callback_type.python_name == 'increment':
            self.inc_button.add_callback(rc, callback_type=CallbackTypes.on_action)
        elif rc.callback_type.python_name == 'decrement':
            self.dec_button.add_callback(rc, callback_type=CallbackTypes.on_action)
        elif rc.callback_type.python_name in ['get_enabled', 'get_visible']:
            self.inc_button.add_callback(rc)
            self.dec_button.add_callback(rc)
            self.txt_box.add_callback(rc)
        else:
            self.txt_box.add_callback(rc)
        #
        #self._local_callbacks[rc.callback_type.python_name] = rc
    def set_id(self, fallback_id=None, ribbon_short_id=None, id_tag=None):
        Box.set_id(self,
                   fallback_id=fallback_id,
                   ribbon_short_id=ribbon_short_id,
                   id_tag=id_tag)
        box_id = self.args.id

        self.txt_box['id']    = box_id + '_text'
        self.inc_button['id'] = box_id + '_increment'
        self.dec_button['id'] = box_id + '_decrement'


class WorkingSpinner(SpinnerBox):
    
    def __init__(self, **kwargs):
        super(WorkingSpinner, self).__init__()
        
        # inc_rc = Callback( self.increment, CallbackTypes.increment, {'context': True} )
        # dec_rc = Callback( self.decrement, CallbackTypes.decrement, {'context': True} )
        inc_rc = Callback( self.increment, python_name='increment', context=True )
        dec_rc = Callback( self.decrement, python_name='decrement', context=True )
        self.add_callback( inc_rc )
        self.add_callback( dec_rc )
        self.step = 1
    
    def increment(self, context):
        #return context.invoke_callback( self._local_callbacks['on_change'], context.invoke_callback( self._local_callbacks['get_text']) + self._step )
        return context.invoke_callback( self.txt_box.on_change, context.invoke_callback( self.txt_box.get_text) + self.step )

    def decrement(self, context):
        return context.invoke_callback( self.txt_box.on_change, context.invoke_callback( self.txt_box.get_text) - self.step )
    
    # def add_callback(self, callback):
    #     if callback.callback_type ==
    


# @uicontrol(SpinnerBox)
# class WorkingSpinner2(FeatureContainer):
#
#     @context({'ctx_resolution': True, 'uicontrol': True})
#     @callback('increment')  # --> automatisch ?
#     def increment(self, ctx_resolution, uicontrol):
#         return ctx_resolution.invoke( uicontrol._callbacks['on_change'], ctx_resolution.invoke( uicontrol._callbacks['get_text']) + 5 )
#
#     @context({'ctx_resolution': True, 'uicontrol': True})
#     @callback('decrement')  # --> automatisch ?
#     def decrement(self, ctx_resolution, uicontrol):
#         return ctx_resolution.invoke( uicontrol._callbacks['on_change'], ctx_resolution.invoke( uicontrol._callbacks['get_text']) - 5 )


class DummyContext(object):
    ''' simply dummy object to resolve context for on_change and get_text in WorkingSpinner '''
    def invoke_callback(self, callback, *args):
        kwargs = {}
        if callback.invocation_context.context:
            kwargs['context'] = self
        if len(callback.callback_type.pos_args) > 0:
            kwargs['value'] = args[0]
        
        #print "will invoke %s with params %s" % (callback, kwargs)
        
        return callback.method(**kwargs)


class UIDefinitionTest(unittest.TestCase):
    
    def test_simple_spinner(self):
        bkt.ribbon.RibbonControl.no_id = True
        
        self.maxDiff = None
        ctrl = SpinnerBox()
        self.assertEqual(ctrl_to_str(ctrl), '<box>\n<editBox sizeString="###" />\n<button label="\xab" />\n<button label="\xbb" />\n</box>')

        def my_gettext():
            return 10
        def my_onchange(value):
            return "changed: " + value

        def my_inc():
            return "incremented"
        def my_dec():
            return "decremented"

        cb_onchange = Callback(my_onchange, CallbackTypes.on_change)
        cb_gettext = Callback(my_gettext, CallbackTypes.get_text)
        cb_inc = Callback(my_inc, CallbackTypes.increment)
        cb_dec = Callback(my_dec, CallbackTypes.decrement)

        ctrl.add_callback( cb_onchange )
        ctrl.add_callback( cb_gettext )
        ctrl.add_callback( cb_inc )
        self.assertEqual(ctrl_to_str(ctrl), '<box>\n<editBox getText="PythonGetText" onChange="PythonOnChange" sizeString="###" />\n<button label="\xab" />\n<button label="\xbb" onAction="PythonOnAction" />\n</box>')

        ctrl.add_callback( cb_dec )
        self.assertEqual(ctrl_to_str(ctrl), '<box>\n<editBox getText="PythonGetText" onChange="PythonOnChange" sizeString="###" />\n<button label="\xab" onAction="PythonOnAction" />\n<button label="\xbb" onAction="PythonOnAction" />\n</box>')



    def test_simple_spinner_callbacks(self):
        ctrl = SpinnerBox()

        def my_gettext():
            return 10
        def my_onchange(value):
            return "changed: " + value

        def my_inc():
            return "incremented"
        def my_dec():
            return "decremented"

        cb_onchange = Callback(my_onchange, CallbackTypes.on_change)
        cb_gettext = Callback(my_gettext, CallbackTypes.get_text)
        cb_inc = Callback(my_inc, CallbackTypes.increment)
        cb_dec = Callback(my_dec, CallbackTypes.decrement)

        ctrl.add_callback( cb_onchange )
        ctrl.add_callback( cb_gettext )
        
        # FIXME: Callback aus ctrl hat eine control-Referenz
        # self.assertEqual(ctrl.collect_callbacks(), [cb_onchange, cb_gettext])
        
        ctrl.add_callback( cb_inc )
        ctrl.add_callback( cb_dec )
        #cb_inc.callback_type = CallbackTypes.on_action
        #cb_dec.callback_type = CallbackTypes.on_action
        # FIXME: Callback aus ctrl hat eine control-Referenz
        # self.assertEqual(ctrl.collect_callbacks(), [cb_inc, cb_dec, cb_onchange, cb_gettext])

    
    def test_working_spinner(self):
        bkt.ribbon.RibbonControl.no_id = True
        
        ctrl = WorkingSpinner()
        
        self.assertEqual(ctrl_to_str(ctrl), '<box>\n<editBox sizeString="###" />\n<button label="\xab" onAction="PythonOnAction" />\n<button label="\xbb" onAction="PythonOnAction" />\n</box>')
        
        def my_gettext():
            return 10
        def my_onchange(value):
            return "changed: %s" % value

        cb_onchange = Callback(my_onchange, CallbackTypes.on_change)
        cb_gettext = Callback(my_gettext, CallbackTypes.get_text)
        ctrl.add_callback( cb_onchange )
        ctrl.add_callback( cb_gettext )
        
        dec = ctrl.children[1].collect_callbacks()[0]
        inc = ctrl.children[2].collect_callbacks()[0]
        
        context = DummyContext()
        
        self.assertEqual(context.invoke_callback( cb_gettext ), 10)
        self.assertEqual(context.invoke_callback( dec ), "changed: 9")
        self.assertEqual(context.invoke_callback( inc ), "changed: 11")

        ctrl.step = 5
        self.assertEqual(context.invoke_callback( cb_gettext ), 10)
        self.assertEqual(context.invoke_callback( dec ), "changed: 5")
        self.assertEqual(context.invoke_callback( inc ), "changed: 15")
        
        
        
        
                
    
        
        
        
        
        
        
        
        
        
        
