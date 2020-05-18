# -*- coding: utf-8 -*-
'''
Created on 23.07.2015

@author: rdebeerst
'''

import bkt
import unittest


class ClassAnnotationAbstractTest(unittest.TestCase):
    
    # ===============================
    # = abstract annotation command =
    # ===============================
    
    def test_annotation_command(self):
        def foo(param):
            pass
        
        cmd = bkt.annotation.AnnotationCommand(foo)
        self.assertTrue(isinstance(cmd, bkt.annotation.AnnotationCommand))
        self.assertEqual(cmd._annotation_method, foo)
        
    def test_annotation_command_list(self):
        def foo(param):
            pass
        
        cmd_list = bkt.annotation.AnnotationCommandList(foo)
        self.assertTrue(isinstance(cmd_list, bkt.annotation.AnnotationCommandList))
        self.assertEqual(cmd_list._annotation_commands[0]._annotation_method, foo)

        cmd = bkt.annotation.AnnotationCommand(foo)
        cmd_list = bkt.annotation.AnnotationCommandList(cmd)
        self.assertTrue(isinstance(cmd_list, bkt.annotation.AnnotationCommandList))
        self.assertTrue(cmd_list._annotation_commands[0], cmd)
        
        cmd_list = bkt.annotation.AnnotationCommandList([cmd])
        self.assertTrue(isinstance(cmd_list, bkt.annotation.AnnotationCommandList))
        self.assertTrue(cmd_list._annotation_commands[0], cmd)
        
        
    
    
    # ====================
    # = command chaining =
    # ====================
    
    def test_annotation_command_chaining(self):
        def foo(param):
            pass
        
        cmd1 = bkt.annotation.AnnotationCommand(foo)
        cmd2 = bkt.annotation.AnnotationCommand(foo)
        
        cmd_list = cmd2.chain(cmd1)
        self.assertEqual(cmd_list._annotation_commands[0], cmd1)
        self.assertEqual(cmd_list._annotation_commands[1], cmd2)
        
        #lb = bkt.button(bkt.configure(size='large'))
    
    def test_annotation_command_chaining_mult(self):
        def foo(param):
            pass
        
        cmd1 = bkt.annotation.AnnotationCommand(foo)
        cmd2 = bkt.annotation.AnnotationCommand(foo)
        
        cmd_list = cmd2 * cmd1
        self.assertEqual(cmd_list._annotation_commands[0], cmd1)
        self.assertEqual(cmd_list._annotation_commands[1], cmd2)

    
    def test_annotation_command_list_chain_associative(self):
        def foo(param):
            pass

        cmd1 = bkt.annotation.AnnotationCommand(foo)
        cmd2 = bkt.annotation.AnnotationCommand(foo)
        cmd3 = bkt.annotation.AnnotationCommand(foo)
        
        cmd_list_1 = cmd3.chain( cmd2.chain( cmd1) )
        self.assertEqual(cmd_list_1._annotation_commands, [cmd1,cmd2,cmd3])
        
        cmd_list_2 = cmd3.chain(cmd2).chain( cmd1 ) 
        self.assertEqual(cmd_list_1._annotation_commands, cmd_list_2._annotation_commands)
    
    
    def test_annotation_command_list_chaining(self):
        def foo(param):
            pass

        cmd1 = bkt.annotation.AnnotationCommand(foo)
        cmd2 = bkt.annotation.AnnotationCommand(foo)
        cmd3 = bkt.annotation.AnnotationCommand(foo)
        cmd_list_1 = cmd1.chain( cmd2 )

        # chain list with command
        cmd_list_2 = cmd_list_1.chain( cmd3 )
        # chaining didn't change the original command-list
        self.assertEqual(len(cmd_list_1._annotation_commands), 2)
        self.assertEqual(len(cmd_list_2._annotation_commands), 3)
        
        # create chain of chains
        cmd4 = bkt.annotation.AnnotationCommand(foo)
        cmd_list_2 = cmd3.chain( cmd4 )
        self.assertEqual(cmd_list_1._annotation_commands, [cmd2, cmd1])
        self.assertEqual(cmd_list_2._annotation_commands, [cmd4, cmd3])
        cmd_list_3 = cmd_list_1.chain( cmd_list_2 )
        # chain of chain working
        self.assertEqual(cmd_list_3._annotation_commands, [cmd4, cmd3, cmd2, cmd1])
        # original chains unchanged
        self.assertEqual(cmd_list_1._annotation_commands, [cmd2, cmd1])
        self.assertEqual(cmd_list_2._annotation_commands, [cmd4, cmd3])
        
        
        

    def test_annotation_command_list_chaining_mult(self):
        def foo(param):
            pass

        cmd1 = bkt.annotation.AnnotationCommand(foo)
        cmd2 = bkt.annotation.AnnotationCommand(foo)
        cmd3 = bkt.annotation.AnnotationCommand(foo)
        
        cmd_list_1 = cmd1 * cmd2
        cmd_list_2 = cmd_list_1 * cmd3
        
        self.assertEqual(cmd_list_1._annotation_commands, [cmd2,cmd1])
        self.assertEqual(cmd_list_2._annotation_commands, [cmd3, cmd2, cmd1])
        
        cmd_list_3 = cmd1 * cmd2 * cmd3 * cmd1
        self.assertEqual(cmd_list_3._annotation_commands, [cmd1, cmd3, cmd2, cmd1])
        cmd_list_3 = (cmd1 * cmd2) * (cmd3 * cmd1)
        self.assertEqual(cmd_list_3._annotation_commands, [cmd1, cmd3, cmd2, cmd1])
        
        
        
        
        
    
    # ======================================
    # = apply arguments / partial commands =
    # ======================================
    
    def test_annotation_command_apply_args(self):
        def foo(param):
            pass
        
        cmd = bkt.annotation.AnnotationCommand(foo)
        self.assertEqual(cmd._args, ())
        self.assertEqual(cmd._kwargs, {})
        
        cmd2 = cmd.apply_arguments('test', label='foo')
        self.assertEqual(cmd2._args , ('test',))
        self.assertEqual(cmd2._kwargs , {'label':'foo'})
        # cmd is also changed
        self.assertEqual(cmd._args , ('test',))
        self.assertEqual(cmd._kwargs, {'label':'foo'})
    
    def test_partial_commands(self):
        def foo(param):
            pass
        
        cmd = bkt.annotation.AnnotationCommand(foo)
        self.assertEqual(cmd._args, ())
        self.assertEqual(cmd._kwargs, {})
        
        cmd2 = cmd.partial('test', label='foo')
        self.assertEqual(cmd2._args , ('test',))
        self.assertEqual(cmd2._kwargs , {'label':'foo'})
        # cmd unchanged
        self.assertEqual(cmd._args , ())
        self.assertEqual(cmd._kwargs, {})
        
        cmd3 = cmd('test2', label='foo2')
        self.assertEqual(cmd3._args , ('test2',))
        self.assertEqual(cmd3._kwargs , {'label':'foo2'})
        # cmd2 unchanged
        self.assertEqual(cmd2._args , ('test',))
        self.assertEqual(cmd2._kwargs , {'label':'foo'})
        # cmd unchanged
        self.assertEqual(cmd._args , ())
        self.assertEqual(cmd._kwargs, {})
        
        cmd4 = cmd3('another_arg', key='value')
        self.assertEqual(cmd4._args , ('test2', 'another_arg'))
        self.assertEqual(cmd4._kwargs , {'label':'foo2', 'key':'value'})
        # cmd3 unchanged
        self.assertEqual(cmd3._args , ('test2',))
        self.assertEqual(cmd3._kwargs , {'label':'foo2'})
        
    
        
    def test_partial_command_lists(self):
        def foo1(param):
            pass
        def foo2(param):
            pass
        
        cmd1 = bkt.annotation.AnnotationCommand(foo1)
        cmd2 = bkt.annotation.AnnotationCommand(foo2)
        cmd_list = cmd1.chain(cmd2)
        
        partial = cmd_list('arg1')
        
        self.assertEqual(cmd_list._annotation_commands, [cmd2, cmd1])
        self.assertEqual(partial._annotation_commands[1]._args, ('arg1',))
        self.assertEqual(cmd_list._annotation_commands[1]._args, ())
        self.assertEqual(cmd1._args, ())
        self.assertEqual(cmd2._args, ())
    
        partial2 = cmd_list(key='value')
        self.assertEqual(cmd_list._annotation_commands, [cmd2, cmd1])
        self.assertEqual(partial2._annotation_commands[1]._kwargs, {'key':'value'})
        self.assertEqual(cmd_list._annotation_commands[1]._kwargs, {})
        self.assertEqual(cmd1._kwargs, {})
        self.assertEqual(cmd2._kwargs, {})
    
    # ======================
    # = annotation methods =
    # ======================
    
    def test_annotating_method(self):
        def bar(param):
            pass
        
        cmd1 = bkt.annotation.AnnotationCommand(bar)
        am = cmd1(bar)
        
        self.assertTrue(isinstance(am, bkt.annotation.AnnotatedMethod))
        self.assertEqual(am._target, bar)
        self.assertEqual(am._annotation_commands[0], cmd1)
    
    
    def test_annotating_method_chained(self):
        def foo1(param):
            pass
        def foo2(param):
            pass
        def bar(param):
            pass
            
        cmd1 = bkt.annotation.AnnotationCommand(foo1)
        cmd2 = bkt.annotation.AnnotationCommand(foo2)
        
        cmd_list = cmd2.chain(cmd1) # = cmd2 * cmd1
        am = cmd_list(bar)
        self.assertTrue(isinstance(am, bkt.annotation.AnnotatedMethod))
        self.assertEqual(am._target, bar)
        self.assertEqual(am._annotation_commands, [cmd1, cmd2])
        
        am = cmd2(cmd1(bar))
        self.assertTrue(isinstance(am, bkt.annotation.AnnotatedMethod))
        self.assertEqual(am._target, bar)
        self.assertEqual(am._annotation_commands, [cmd1, cmd2])

        am1 = cmd1(bar)
        am2 = cmd2(am1)
        self.assertTrue(isinstance(am2, bkt.annotation.AnnotatedMethod))
        self.assertEqual(am2._target, bar)
        self.assertEqual(am2._annotation_commands, [cmd1, cmd2])
        
    
    def test_annotation_command_list_operations(self):
        def foo1(param):
            pass
        def foo2(param):
            pass
        def foo3(param):
            pass
        def bar(param):
            pass


        cmd1 = bkt.annotation.AnnotationCommand(foo1)
        cmd2 = bkt.annotation.AnnotationCommand(foo2)
        cmd3 = bkt.annotation.AnnotationCommand(foo3)
        
        cmd_list_1 = cmd3.chain(cmd2).chain(cmd1) # = (cmd3 * cmd2) * cmd1
        cmd_list_2 = cmd3.chain( cmd2.chain(cmd1) ) # = cmd3 * (cmd2 * cmd1) 
        
        annotated_foo = cmd3(cmd2(cmd1(bar)))
        self.assertEqual(cmd_list_1._annotation_commands, [cmd1, cmd2, cmd3])
        self.assertEqual(cmd_list_2._annotation_commands, [cmd1, cmd2, cmd3])
        self.assertEqual(annotated_foo._annotation_commands, [cmd1, cmd2, cmd3])        

        annoted_foo_1 = cmd_list_1(bar)
        annoted_foo_2 = cmd_list_2(bar)
        self.assertEqual(annoted_foo_1._annotation_commands, [cmd1, cmd2, cmd3])
        self.assertEqual(annoted_foo_2._annotation_commands, [cmd1, cmd2, cmd3])
        
    
    # ======================
    # = annotation classes =
    # ======================
    
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
    
    



    # ==================
    # = sub annotators =
    # ==================

    # def test_sub_annotator(self):
    #     raise NotImplementedError
    #
    # def test_default_sub_annotator(self):
    #     raise NotImplementedError

    
    
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
    
    

