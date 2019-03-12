# -*- coding: utf-8 -*-
'''
Created on 23.11.2014

@author: cschmitt
'''


#FIXME
#from callbacks import CallbackType
import bkt.helpers as _h
from bkt.callbacks import InvocationContext

import types
import collections
import logging

PRIO_LOWEST = 999
PRIO_CALLBACK_DEFINITION = 100
PRIO_NODE_DEFINITION = 200
PRIO_CONFIGURATION = 300

'''
Contains known subdecorators of annotated objects (accessible via LazyAnnotation.__getattr__)
Keys: name of decorator
Values: instance of MagicLazyAnnotator 
'''

class DecorationError(Exception):
    pass

def infinite_counter():
    def seq():
        value = 0
        while True:
            value += 1
            yield value
            
    s = seq()
    
    def _next():
        return next(s)
    
    return _next

_declaration_counter = infinite_counter()




# =============================
# = Base-Type for annotations =
# =============================


class Annotation(object):
    ''' The Annotation-Object is created when methods are decorated with AnnotationCommands '''
    def __init__(self, target, target_name=None):
        self.target = target
        self.target_order = 0
        self.last_prio = None
        
        # filled when class is unwrapped
        self.target_name = None
        self.container_cls = None
        self.parents = None
        self.children = []
        
        # BKT-sepcific
        #self.callback_info = None
        #self.ui_info = None
        
    def __getattr__(self, attr):
        ''' Returns None for all non-existing attributes.
            Makes the Annotation-class BKT-independent since BKT-specific attributes (callback_info, ui_info) don't have to be initialized.
        '''
        return None
    
    @property
    def is_root(self):
        return self.target is not None and self.target is self.container_cls
    
    def __repr__(self):
        return '<%s: target_name=%s, target=%s, order=%s>' % (type(self).__name__, self.target_name, self.target, self.target_order)



# =======================
# = Annotations-Objekte =
# =======================

class AnnotationCommand(object):
    '''
    Instancses of this class are used as decorators for methods or classes.
    Class will contain which is called when code is iterpreted by python. 
    
    Example:
        @some_annotation_command
        def some_method(...):
            pass
    '''
    
    def __init__(self, annotation_method, prio=None):
        ''' AnnotationCommand(method, prio)
            AnnotationCommand(prio)(method)
            AnnotationCommand(method)
        '''
    #FIXME: def __init__(self, lazy_annotator, args=None, kwargs=None, context=None):
        self._args = ()
        self._kwargs = {}
        self._priority = prio or PRIO_LOWEST
        self._annotation_method = annotation_method
        #FIXME: self.context = context
        # parent is set by AnnotatedMethod if AnnotationCommand is used in its context
        self._parent_AnnotatedMethod = None
    
    
    def __call__(self, *args, **kwargs):
        ''' convenience method to specify arguments or to apply the AnnotationCommand to methods/classes '''
        
        #logging.debug('call of AnnotationCommand %s (method %s) with args: args=%s, kwargs=%s' % (self, self._annotation_method, args, kwargs))
        
        if len(args) == 0 and len(kwargs) == 0:
            #raise TypeError('AnnotationCommand called without arguments, expected AnnotatedMethod, FeatureContainer, FunctionType or arbitrary arguments')
            return self.apply_arguments(*args, **kwargs)
        elif len(kwargs) > 0 or len(args) > 1:
            return self.partial(*args, **kwargs)
            
        elif isinstance(args[0], AnnotatedType):
            return self.decorate_class(args[0])
            
        elif isinstance(args[0], types.FunctionType):
            return self.decorate_method(args[0])
        
        elif isinstance(args[0], AnnotatedMethod):
            args[0].append_annotation_command(self)
            return args[0]
            
        elif isinstance(args[0], AnnotationCommand):
            # called: self(args[0])
            #return args[0].append(self)
            raise Exception('argument AnnotationCommand deprecated. Use .chain instead.\n %s' % (self._annotation_method))
            
        
        else:
            return self.partial(*args, **kwargs)
    
    def chain(self, other_command):
        return other_command.append(self)
    
    def __mul__(self, other):
        return self.chain(other)
        
    def invoke(self, annotation):
        ''' invoke the annotation on the target '''
        #logging.debug('will invoke %s with args: args=%s, kwargs=%s' % (self._annotation_method, self._args, self._kwargs))
        self._annotation_method(annotation, *self._args, **self._kwargs)
    
    def apply_arguments(self, *args, **kwargs):
        ''' arguments are saved and passed, when annotation is invoked '''
        self._args += args
        self._kwargs.update(kwargs)
        return self
    
    def copy(self):
        return AnnotationCommand(self._annotation_method, prio=self._priority)
    
    def partial(self, *args, **kwargs):
        ''' create a new AnnotationCommand where the passed arguments are predefined '''
        cmd = self.copy()
        #cmd = AnnotationCommand(self._annotation_method, prio=self._priority)
        cmd._args = tuple(self._args)
        cmd._kwargs = dict(self._kwargs)
        cmd.apply_arguments(*args, **kwargs)
        return cmd
    
    
    def decorate_class(self, target):
        ''' class is decorated with the AnnotationCommand, returns the class itself.
        Example:
        @some_annotation_command
        class some_class(...)
            pass
        '''
        
        # logging.debug('AnnotationCommand.decorate_class: self=%s, target=%s' % (self, target))
        
        if target._annotation_self_lazy is None:
            target._annotation_self_lazy = Annotation(target)
            # FIXME: Annotation sollte von AbstractAnnotationObject ableiten, worüber target_order automatisch gesetzt wird
            target._annotation_self_lazy.target_order = _declaration_counter()
        
        annotation = target._annotation_self_lazy
        if annotation.last_prio is not None and annotation.last_prio > self._priority:
            raise DecorationError('%s annotated in wrong order' % target)
        
        # FIXME: lazy class decoration erlauben. annotation-commands merken, und in AnnotatedType.__new__ invoke ausführen
        self.invoke(annotation)
        annotation.last_prio = self._priority
        
        return target
    
    
    def decorate_method(self, method):
        ''' method is decorated with the AnnotationCommand, returns an AnnotatedMethod.
        Example:
        @some_annotation_command
        def some_method(...)
            pass
        '''
        return AnnotatedMethod(method, self)
    
    
    def append(self, command):
        return AnnotationCommandList([self, command])
    
    def __repr__(self):
        return '<%s: with method %s and arguments: args=%s, kwargs%s>' % (type(self).__name__, self._annotation_method, self._args, self._kwargs)


        

class AnnotationCommandList(AnnotationCommand):
    ''' Class representes an AnnotationCommand created by chaining other AnnotationCommands
        Internally, the AnnotationCommand to invoke first is listed first.
        Example:
            cmd_list = cmd3 * cmd2 * cmd1
            cmd_list._annotation_commands = [cmd1, cmd2, cmd3]
    '''
    
    def __init__(self, method_cmd_or_cmdlist=None):
        ''' AnnotationCommandList(method) equivalent to AnnotationCommandList(AnnotationCommand(method))
            AnnotationCommandList(cmd1) 
            AnnotationCommandList([cmd2, cmd1])
        '''
        
        if method_cmd_or_cmdlist==None:
            self._annotation_commands = []
        elif isinstance(method_cmd_or_cmdlist, list):
            self._annotation_commands = method_cmd_or_cmdlist
        elif isinstance(method_cmd_or_cmdlist, AnnotationCommand):
            self._annotation_commands = [method_cmd_or_cmdlist]
        elif isinstance(method_cmd_or_cmdlist, types.FunctionType):
            self._annotation_commands = [AnnotationCommand(method_cmd_or_cmdlist)]
        else:
            raise TypeError('AnnotationCommandList initialized with wrong arguments. Expected FunctionType, AnnotationCommand or list of AnnotationCommands.')
    
    def __call__(self, *args, **kwargs):
        ''' convenience method to specify arguments or to apply the AnnotationCommand to methods/classes '''
        
        # logging.debug('call of AnnotationCommandList %s (#cmds %d) with args: args=%s, kwargs=%s' % (self, len(self._annotation_commands), args, kwargs))
        
        
        if len(args) == 1 and (
            isinstance(args[0], types.FunctionType) or
            isinstance(args[0], AnnotatedType) or
            isinstance(args[0], AnnotatedMethod) 
            ):
            # one argument, which is of the accepted type, then handle the commands sequentially
            # assume _annotation_commands = [a,b,c]
            # then return c(b(a( args[0] )))
            
            value = args[0]
            for cmd in self._annotation_commands:
                value = cmd(value)
            return value
            
        elif len(args) == 1 and isinstance(args[0], AnnotationCommand):
            # chain the command
            raise Exception('argument AnnotationCommand deprecated. Use .chain instead')
            #copy = AnnotationCommandList([args[0]] + self._annotation_commands)
            #return copy
            
        else:
            # if any other type of argument accurs, then pass it to the last AnnotationCommand
            return self.partial(*args, **kwargs)
    
    def chain(self, other_command):
        if isinstance(other_command, AnnotationCommandList):
            # empty list [] ensures that the copy starts with a new list-object
            copy = AnnotationCommandList([] + other_command._annotation_commands + self._annotation_commands)
        else:
            copy = AnnotationCommandList([other_command] + self._annotation_commands)
        return copy
    
    def __mul__(self, other_command):
        return self.chain(other_command)
        
    
    def invoke(self, annotation):
        ''' invoke the annotations on the target in reversed order '''
        for command in sorted(self._annotation_commands, key=lambda c : c._priority):
            command.invoke(annotation)
    
    def apply_arguments(self, *args, **kwargs):
        ''' arguments are saved passed to the first AnnotationCommand '''
        if len(self._annotation_commands) == 0:
            raise Warning('Can\'t apply arguments to an empty AnnotationCommandList')
        else:
            self._annotation_commands[0] = self._annotation_commands[0].apply_arguments(*args, **kwargs)
        return self
    
    def copy(self):
        return AnnotationCommandList(list(self._annotation_commands))
        
    def partial(self, *args, **kwargs):
        ''' Create a new AnnotationCommand where the arguments are predefined in the last AnnotationCommand.
        Only the first AnnotionCommand becomes a new instance. '''
        if len(self._annotation_commands) == 0:
            raise Warning('Can\'t compute partial of an empty AnnotationCommandList')
            return self
        else:
            cmdlist = AnnotationCommandList(list(self._annotation_commands))
            cmdlist._annotation_commands[-1] = cmdlist._annotation_commands[-1].partial(*args, **kwargs)
            return cmdlist
    
    
    def decorate_class(self, target):
        ''' class is decorated with the AnnotationCommand, returns the class itself.
        Example:
        @some_annotation_command
        class some_class(...)
            pass
        '''
        return self(target)
    
    
    def decorate_method(self, method):
        ''' method is decorated with the AnnotationCommand, returns an AnnotatedMethod.
        Example:
        @some_annotation_command
        def some_method(...)
            pass
        '''
        return self(method)
    
    
    def append(self, command):
        ''' equivalend to command(self)
        AnnotationCommands and AnnotationCommandLists are chainable, such that
        command*self(params) = command(self(params))
        '''
        copy = self.copy()
        copy._annotation_commands.append(command)
        return copy

    def __repr__(self):
        return '<%s: %d annotation commands>' % (type(self).__name__, len(self._annotation_commands))
    


class AnnotatedMethod(object):
    ''' Class represents a method which is decorated by an AnnotationCommand '''
    _sub_annotation_commands = {}
    
    def __init__(self, method, command=None):
        self._target = method
        # if command:
        #     self.append_annotation_command(command)
        self._annotation_commands = [command] if command else []
        self._order = _declaration_counter()
    
    def append_annotation_command(self, command):
        # logging.debug('AnnotatedMethod.append_annotation_command: self=%s, command=%s' % (self, command))
        if not isinstance(command, AnnotationCommand):
            raise TypeError
        # if isinstance(command, AnnotationCommandList):
        #     self._annotation_commands.extend(command._annotation_commands)
        self._annotation_commands.append(command)
    
    def invoke_annotations(self):
        ''' invoke the annotation commands '''
        annotation = Annotation(self._target)
        annotation.target_order = self._order
        parents = set()
        for command in sorted(self._annotation_commands, key=lambda c : c._priority):
            command.invoke(annotation)
            if command._parent_AnnotatedMethod is not None:
                parents.add(command._parent_AnnotatedMethod)
        
        # FIXME: allow multiple parents, e.g. to reuse an enabled-callback
        parents = list(parents)
        
        return annotation, parents
        
    def __repr__(self):
        return '<%s: %d annotations pending for %s>' % (type(self).__name__, len(self._annotation_commands), self._target)
    
    
    def __getattr__(self, attr):
        ''' access sub-AnnotationCommands '''
        if attr in ['__init__', '__repr__', 'append_annotation_command', 'invoke_annotations']:
            raise AttributeError(attr)
        # elif type(self)._sub_annotation_commands.has_key(attr):
        #     # return AnnotationCommand
        #     pass
        elif type(self)._default_sub_annotation_command:
            # return AnnotationCommand with param attr
            cmd = type(self)._default_sub_annotation_command(attr)
            cmd._parent_AnnotatedMethod = self 
            return cmd
        else:
            raise AttributeError(attr)
    




# ======================================
# = Objects to handle AnnotatedMethods =
# ======================================


class AnnotatedType(type):
    '''
    Metaclass to help the use of decorators as method annotators (i.e. to add some meta
    information to the method) without changing the actual method of the constructed class
    and without losing the method context.
    This is achieved by using decorators which wrap the target in an idempotent
    fashion into a annotation object (see Annotation) during class body execution
    and by unwrapping the target before constructing the class. 
    '''
    def __new__(cls, clsname, bases, dct):
        ''' This method is called when a new class with __metaclass__ = AnnotatedType is created.
        The definitions in dct (interpreted class body) are analyzed to unwrap AnnotatedMethods etc.
        '''
        
        # logging.debug('AnnotatedType.__new__: cls=%s, clsname=%s, bases=%s, dct=%s' % (cls, clsname, bases, dct))
        
        raw_annotations = [(k, v) for k, v in dct.items() if isinstance(v, AnnotatedMethod)]
        raw_annotations.sort(key=lambda t : t[1]._order)
        
        annotations = []
        
        anno_to_method = {} # list of AnnotatedMethod-Objects
        method_to_anno = {} # list of Annotation-Objects
        method_parents = {} # list of AnnotatedMethod-Objects used as parents
        
        for attr, annotated_method in raw_annotations:
            # uwrap annotated functions and save annotation information
            # based on the class attribute name
            annotation, parents = annotated_method.invoke_annotations()
            annotation.target_name = attr

            dct[attr] = annotation.target
            annotations.append(annotation)

            anno_to_method[annotation] = annotated_method
            method_to_anno[annotated_method] = annotation
            method_parents[annotation] = parents
        
        # child-parent-relation from AnnotatedMethod-Objects is rebuild for Annotation-Objects
        for annotation in annotations:
            annotation.parents = []
            for parent in method_parents[annotation]:
                parent_annotation = method_to_anno[parent]
                parent_annotation.children.append(annotation)
                annotation.parents.append(parent_annotation)
                
        annotations = collections.OrderedDict((a.target_name, a) for a in sorted(annotations, key=lambda a : a.target_order))
        
        
        
        usage = []
        mso_controls = []
        for attr, value in dct.items():
            #print(attr, value)
            if isinstance(value, ContainerUsage):
                value.attribute = attr
                usage.append(value)
            elif isinstance(value, AbstractAnnotationObject):
                value.attribute = attr
                mso_controls.append(value)
        
        new_cls = type.__new__(cls, clsname, bases, dct)
        # process annotations either after the class creation has completed
        # using a special class method of the constructed class itself
        # or persist the annotations in a class attribute for later use
        new_cls._annotation_members = annotations
        new_cls._annotation_self_lazy = None
        new_cls._annotation_self_real = None
        new_cls._annotation = ContainerAnnotationDescriptor()
        new_cls._usage = usage
        new_cls._mso_controls = mso_controls

        for annotation in annotations.values():
            annotation.container_cls = new_cls
        
        def dump_method(self):
            text  = '=== dump of ' + str(self) + ' ==='
            text += "\n" + '-- class members'
            text += "\n" + '   usage:        ' + str(self._usage)
            text += "\n" + '   mso-controls: ' + str(self._mso_controls)
            text += "\n" + '-- annotation of class '
            try:
                text += "\n" + '   annotation: ' + str(self._annotation)
                text += "\n" + '   ui_info:    ' + str(self._annotation.ui_info)
            except:
                pass
            text += "\n" + '-- annotation of members'
            for anno_name in self._annotation_members:
                text += "\n" + '   -- ' + anno_name + ' --'
                anno = self._annotation_members[anno_name]
                text += "\n" + '      target:        ' + str(anno.target)
                text += "\n" + '      target_name:   ' + anno.target_name
                text += "\n" + '      target_order:  ' + str(anno.target_order)
                text += "\n" + '      container_cls: ' + str(anno.container_cls)
                text += "\n" + '      parents:        ' + str(anno.parents)
                text += "\n" + '      children:      ' + str(anno.children)
                text += "\n" + '      callback_info: ' + str(anno.callback_info)
                text += "\n" + '      ui_info:       ' + str(anno.ui_info)
            text += "\n" + "=== end dump ===\n"
            return text
        
        new_cls._dump = dump_method
        
        return new_cls
    
class ContainerAnnotationDescriptor(object):
    ''' Descriptor-class to access the Annotation-Object of a class. '''
    
    def __get__(self, obj, cls):
        if cls._annotation_self_real is not None:
            return cls._annotation_self_real
        
        if cls._annotation_self_lazy is None:
            return None
        
        self.initialize_class_annotation(cls)
        return cls._annotation_self_real
    
    
    def initialize_class_annotation(self, cls):
        ''' called on first usage of Descriptor-class '''
        # set meta information
        cls_anno = cls._annotation_self_lazy
        cls_anno.target_name = cls.__name__
        cls_anno.container_cls = cls
        
        # set parent-child-reference
        anno_children = [a for a in cls._annotation_members.values() if len(a.parents) ==0 ]
        cls_anno.children = anno_children
        for child_anno in anno_children:
            child_anno.parents = [cls_anno]
        
        cls._annotation_self_real = cls_anno
        cls._annotation_self_lazy = None
        
        
        
        

class AbstractAnnotationObject(object):
    ''' Base class for ordered annotations and objects. Attributes of this type are considered while unwrapping annotated classes.
        See AnnotatedType and Factory.
    '''
    def __init__(self):
        self.target_order = _declaration_counter()
        self.attribute = None


class ContainerUsage(AbstractAnnotationObject):
    def __init__(self, container, id_tag=None):
        super(ContainerUsage, self).__init__()
        if not issubclass(container, FeatureContainer):
            raise TypeError
        self.container = container
        self.id_tag = id_tag
        self.attribute = None


class FeatureContainer(object):
    '''
    Intended as the parent class of all implementation classes which organize addin logic via annotated
    (decorated) methods. The decorators to be used may add various types of meta information,
    e.g. the context necessary for invocation (see @arg_ decorators) and the type/appearance of
    the associated UI element.
    '''
    __metaclass__ = AnnotatedType



def is_feature_container_class(obj):
    try:
        return issubclass(obj, FeatureContainer)
    except TypeError:
        return False








# ===================================
# = BKT-specific annotation classes =
# ===================================

class CallbackInformation(object):
    ''' Object to store information about a callback.
        Information is filled by AnnotationCommands and unwrapped by the Factory-class. 
    '''
    def __init__(self):
        self.callback_type = None
        self.invocation_context = None
        
        self.control = None
        self.target_container_cls = None
        self.target_method_name = None
    
    def __repr__(self):
        fmt = '<%s callback=%s invocation_context=%s control=%s resolution=%s.%s>'
        return fmt % (type(self).__name__, self.callback_type, self.invocation_context, self.control, self.target_container_cls, self.target_method_name)
        
class UIInformation(object):
    ''' Object to store information about a ui-controls (i.e. Ribbon-Controls).
        Information is filled by AnnotationCommands and unwrapped by the Factory-class. 
    '''
    def __init__(self, node_type):
        self.node_type = node_type
        self.control_args = {}
        
    def __repr__(self):
        return '<%s node_type=%s, args=%s>' % (type(self).__name__, self.node_type, self.control_args)




# =========================
# = Annotation-decorators =
# =========================
# Annotations-Dekoratoren werden verwendet um Annotations-Methoden zu definieren
# und dabei sicherzustellen, dass bestimmte Objekte in der Annotation initialisiert sind.

def require_ui(func):
    def require_ui_wrapper(annotation, *args, **kwargs):
        ui_nfo = annotation.ui_info
        if ui_nfo is None:
            raise DecorationError('wrong annotation order: ' + require_ui_wrapper.__name__ + ' called before ui-info was initialized')
        if not isinstance(ui_nfo, UIInformation):
            raise TypeError('got %s, expected' % (type(ui_nfo), UIInformation))
        return func(annotation, *args, **kwargs)
    
    require_ui_wrapper.__name__ = func.__name__ + '_require_ui'
    return require_ui_wrapper

def ensure_callback_nfo(func):
    def wrapper(annotation, *args, **kwargs):
        cb_nfo = annotation.callback_info
        if cb_nfo is None:
            annotation.callback_info = CallbackInformation()
            
        return func(annotation, *args, **kwargs)
    
    wrapper.__name__ = func.__name__ + '_ensure_callback_nfo'
    return wrapper

def ensure_invocation_context(func):
    def wrapper(annotation, *args, **kwargs):
        cb_nfo = annotation.callback_info
        if cb_nfo is None:
            cb_nfo = CallbackInformation()
            annotation.callback_info = cb_nfo
        if cb_nfo.invocation_context is None:
            cb_nfo.invocation_context = InvocationContext()
            
        return func(annotation, *args, **kwargs)
    
    wrapper.__name__ = func.__name__ + '_ensure_invocation_context'
    return wrapper



class UIControlAnnotationCommand(AnnotationCommand):
    ''' Base-class for annotation-commands defining ui-controls (i.e. Ribbon-Controls) such as tab, group, button.
        Example:
            button = UIControlAnnotationCommand('button')
            @button
            def methode_to_be_decorated(...):
                pass
    '''
    def __init__(self, node_type):
        super(UIControlAnnotationCommand, self).__init__(None, prio=PRIO_NODE_DEFINITION)
        self.node_type = node_type
    
    def copy(self):
        ''' creates a copy of itself '''
        return UIControlAnnotationCommand(self.node_type)
    
    def invoke(self, annotation):
        ''' invokes the annotation-command on the annotation-object '''
        ui_nfo = annotation.ui_info
        
        if ui_nfo is None:
            ui_nfo = UIInformation(self.node_type)
            annotation.ui_info = ui_nfo
        elif not isinstance(ui_nfo, UIInformation):
            raise TypeError
        elif ui_nfo.node_type is not None:
            raise ValueError('node_type already set')
        
    def __repr__(self):
        return '<%s: nodetype=%s>' % (type(self).__name__, self.node_type)




