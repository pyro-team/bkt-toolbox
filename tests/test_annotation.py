# -*- coding: utf-8 -*-
'''
Created on 20.07.2015

@author: rdebeerst
'''

import bkt
import unittest


class ClassAnnotationTest(unittest.TestCase):
    
    
    def test_Annotating_Method(self):
        pass
    
        
    def test_Annotating_Class(self):
        BlankFC = type('BlankFC', (bkt.FeatureContainer,), {})
        myClass = bkt.group(BlankFC)
        x = myClass()
        self.assertTrue(isinstance(x, bkt.FeatureContainer))
    
    def test_Annotating_Class_MultipleAnotations(self):
        BlankFC = type('BlankFC', (bkt.FeatureContainer,), {})
        myClass = bkt.uuid('c3973689-0aec-4922-9846-80d1fdeed457')(bkt.configure(label="BKT Dev Options")(bkt.group(BlankFC)))
        #myClass = bkt.configure(label="BKT Dev Options")(bkt.group(BlankFC))
        x = myClass()
        self.assertTrue(isinstance(x, bkt.FeatureContainer))
    
    
    
    # def test_AnnotatedMethods(self):
    #     x = AnnotatedMethods()
    #     self.assertEqual(type(x), FeatureContainer)
    #
    # @command_2_with_prio_300
    # @command_1_with_prio_300
    # def method():
    #    pass
    # --> sicherstellen, dass command_1 erst aufgerufen wird
    
    # @button(label='button with default on_action')
    # def on_action():
    #     pass
    #
    # @button(label='button with default callback not on_action')
    # @callback('get_image')
    # def get_image():
    #     pass
    
    
    def test(self):
        pass


if __name__ == '__main__':
    unittest.main()
    
    

