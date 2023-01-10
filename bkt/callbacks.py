# -*- coding: utf-8 -*-
'''
Definition of callbacks and invocation context

Created on 13.11.2014
@authors: cschmitt, rdebeerst
'''



import importlib
import logging
from functools import wraps

import bkt.helpers as _h


def get_dotnet_callback_name(python_name):
    # example: on_action --> OnAction
    # dotnet_name = ''.join([x[0].upper() + x[1:] for x in python_name.split('_')])
    return 'Python' + _h.snake_to_upper_camelcase(python_name)

def get_xml_callback_name(python_name):
    # # example: on_action --> OnAction
    # dotnet_name = ''.join([x[0].upper() + x[1:] for x in python_name.split('_')])
    # # OnAction --> onAction
    # return  (dotnet_name[0].lower() + dotnet_name[1:])
    return _h.snake_to_lower_camelcase(python_name)



class CallbackType(object):
    ''' A CallbackType represents a specific callback type as used in the Ribbon-MSCustomXML,
        e.g. a button-click (no arguments) or a on-change-event (with a single text-argument)
    '''
    def __init__(self, python_name=None, dotnet_name=None, xml_name=None, pos_args=None, custom=False, transactional=False, cacheable=False):
        self.python_name = python_name
        if python_name:
            self.dotnet_name = dotnet_name or get_dotnet_callback_name(python_name)
            self.xml_name = xml_name or get_xml_callback_name(python_name)
        else:
            self.dotnet_name = dotnet_name
            self.xml_name = xml_name
        
        self.pos_args = pos_args or []
        self.custom = custom
        self.transactional = transactional
        self.cacheable = cacheable
    
    def __repr__(self):
        args = ', '.join(self.pos_args)
        return '<%s %s(%s) .NET=%s xml=%s>' % (type(self).__name__, self.python_name, args, self.dotnet_name, self.xml_name)
    
    
    def xml(self):
        #return '%s="%s"' % (self.xml_name, self.dotnet_name)
        return self.dotnet_name
    
    
    def set_attribute(self, attr):
        ''' applies naming-conventions when callback is value of attribute attr, i.e. object.attr=Callback(...) '''
        if not self.python_name:
            self.python_name = attr
        if not self.xml_name:
            self.xml_name = get_xml_callback_name(attr)
        if not self.dotnet_name:
            self.dotnet_name = get_dotnet_callback_name(attr)


class CallbackTypesCatalog(object):
    ''' Manages a repository of CallbackTypes. '''
    def __init__(self):
        self._callback_types = {}
    
    def __getattr__(self, attr):
        try:
            return self._callback_types[attr]
        except:
            # define custom callback
            custom = CallbackType(custom=True,python_name=attr)
            # save custom callback-type. On the second access the same object will be returned
            self._callback_types[attr] = custom
            return custom
    
    def __setattr__(self, attr, value):
        if isinstance(value, CallbackType):
            value.set_attribute(attr)
            self._callback_types[attr] = value
        else:
            super(CallbackTypesCatalog, self).__setattr__(attr, value)
    
    def has_callback_type(self, name):
        return name in self._callback_types

    def get_callback_type(self, name):
        return self._callback_types[name]
    
    def callback_map(self):
        return self._callback_types

    def callback_list(self):
        return self._callback_types.values()


def callback_type(*pos_args, **kw_args):
    return CallbackType(pos_args=pos_args, **kw_args)

def tx_callback_type(*pos_args, **kw_args):
    return CallbackType(pos_args=pos_args, transactional=True, **kw_args)


CallbackTypes = CallbackTypesCatalog()

# information callbacks (no arguments)
CallbackTypes.get_content     = callback_type()
CallbackTypes.get_description = callback_type()
CallbackTypes.get_enabled     = callback_type(cacheable=True)
CallbackTypes.get_image       = callback_type()
CallbackTypes.get_keytip      = callback_type()
CallbackTypes.get_label       = callback_type()
CallbackTypes.get_pressed     = callback_type()
CallbackTypes.get_screentip   = callback_type()
CallbackTypes.get_show_image  = callback_type()
CallbackTypes.get_show_label  = callback_type()
CallbackTypes.get_size        = callback_type()
CallbackTypes.get_supertip    = callback_type()
CallbackTypes.get_text        = callback_type()
CallbackTypes.get_title       = callback_type()
CallbackTypes.get_visible     = callback_type(cacheable=True)

# Callbacks for Gallery/ComboBox
CallbackTypes.get_item_count          = callback_type()
CallbackTypes.get_selected_item_index = callback_type()
CallbackTypes.get_selected_item_id    = callback_type(xml_name='getSelectedItemID', dotnet_name='PythonGetSelectedItemID')
# indexed callbacks
CallbackTypes.get_item_height    = callback_type()
CallbackTypes.get_item_id        = callback_type('index', xml_name='getItemID', dotnet_name='PythonGetItemID')
CallbackTypes.get_item_image     = callback_type('index')
CallbackTypes.get_item_label     = callback_type('index')
CallbackTypes.get_item_screentip = callback_type('index')
CallbackTypes.get_item_supertip  = callback_type('index')
CallbackTypes.get_item_width     = callback_type()

# action Callbacks
CallbackTypes.on_action            = tx_callback_type()
CallbackTypes.on_action_indexed    = tx_callback_type('selected_item', 'index', xml_name='onAction')
CallbackTypes.on_action_repurposed = tx_callback_type(xml_name='onAction')
CallbackTypes.on_toggle_action     = tx_callback_type('pressed', xml_name='onAction')
CallbackTypes.on_change            = tx_callback_type('value')

# callbacks loadImage/onLoad unused
# CallbackTypes.loadImage = callback_type('image')
# CallbackTypes.onLoad = callback_type()


#FIXME: custom-callback-types sollten nicht explizit definiert werden müssen
CallbackTypes.increment = CallbackType(custom=True, transactional=True)
CallbackTypes.decrement = CallbackType(custom=True, transactional=True)



# WPF general callback
CallbackTypes.wpf_event = callback_type(xml_name=None, dotnet_name='WPFEvent')
CallbackTypes.wpf_action = tx_callback_type(xml_name=None, dotnet_name='WPFAction')

# BKT general events
CallbackTypes.bkt_event = callback_type(xml_name=None, dotnet_name='BKTEvent')



class Callback(object):
    # CALLBACK_KEYS = set(('python_name', 'dotnet_name', 'xml_name', 'pos_args', 'custom', 'transactional', 'cacheable'))
    CALLBACK_KEYS = set(('transactional', 'cacheable'))

    ''' Represents a callback-method with information about the method-arguments which need to be passed to invoke the method. '''
    def __init__(self, *args, **kwargs):
        ''' Initialization method, use on of the following options
             1) Callback( method, callback_type, invocation_context)
             2) Callback( container_class, method_name, callback_type, invocation_context )
                    The method is obtained from the container_class
             3) Callback( method, callback_type, **kwargs)
                    The invocation_context is then build from **kwargs
             4) Callback( method, callback_type)
                    The invocation_context is then build from method params.
             5) Callback( method, **kwargs)
                    The invocation_context and the callback_type are then build from **kwargs.
                    callback_type uses the arguments: python_name, dotnet_name, xml_name, pos_args, custom, transactional, cacheable
                    All other arguments are used for the invocation_context
             6) Callback( method)
                    The invocation_context ist build from method params.
        
        '''
        self.container = None
        self.method = None
        self.control = None
        self.callback_type = None

        # extract kwargs that are related to the callback and not passed to invocation_context
        self._callback_args = { key:kwargs.pop(key) for key in list(kwargs.keys()) if key in self.CALLBACK_KEYS }

        len_args = len(args)
        
        if len_args == 4:
            #FIXME: if len(kwargs) != 0: ERROR
            container, method_name, callback_type, invocation_context = args
            self.init_container_method(container, method_name, callback_type, invocation_context)
            
        elif len_args == 3:
            # no logic, alls objects as arguments
            self.method, self.callback_type, self.invocation_context = args
            
        elif len_args == 2:
            method, callback_type = args
            if kwargs:
                self.init_method_callback(method, callback_type, **kwargs)
            else:
                self.init_method_callback_auto(method, callback_type)
            
        elif len_args == 1:
            if kwargs:
                self.init_method(args[0], **kwargs)
            else:
                self.init_method_auto(args[0])
    
    
    def init_container_method(self, container, method_name, callback_type, invocation_context):
        ''' initialization method. Obtains method from container_class '''
        #self.container = container
        self.method = None
        if method_name:
            self.method = getattr(container, method_name)
        if callback_type is None:
            # Fallback, if method_name is a known callback
            if CallbackTypes.has_callback_type(method_name):
                callback_type = CallbackTypes.get_callback_type(method_name)
        self.callback_type = callback_type
        self.invocation_context = invocation_context
    
    def init_method_callback(self, method, callback_type, **kwargs):
        ''' initialization method. Builds invocation_context from keyword-arguments '''
        self.method = method
        self.callback_type = callback_type
        self.invocation_context = InvocationContext(**kwargs)
    
    def init_method_callback_auto(self, method, callback_type):
        ''' initialization method. Builds invocation_context from varnames of method '''
        self.method = method
        self.callback_type = callback_type
        self.invocation_context = InvocationContext.from_method(method)
        
    def init_method(self, method, **kwargs):
        ''' initialization method. Builds invocation_context from keyword-arguments  '''
        # split kwargs for callback_type and invocation_context
        # callback_args = { key:kwargs.pop(key) for key in list(kwargs.keys()) if key in self.CALLBACK_KEYS }
        # callback_args = { key:value for key, value in kwargs.items() if key in self.CALLBACK_KEYS }
        # kwargs        = { key:value for key, value in kwargs.items() if key not in self.CALLBACK_KEYS }
        # callback_type = CallbackType(**callback_args)
        self.method = method
        # self.callback_type = CallbackType(**callback_args)
        self.invocation_context = InvocationContext(**kwargs)
    
    def init_method_auto(self, method):
        ''' initialization method. Builds invocation_context from varnames of method '''
        self.method = method
        self.invocation_context = InvocationContext.from_method(method)


    def set_callback_type(self, callback_type):
        self.callback_type = callback_type

    @property
    def is_cacheable(self):
        return self._callback_args.get("cacheable", self.callback_type.cacheable)
    
    @property
    def is_transactional(self):
        return self._callback_args.get("transactional", self.callback_type.transactional)
    

    def __repr__(self):
        return '<%s container=%s, method=%s, invocation_context=%s, callback=%s, control=%s>' % (type(self).__name__,
                                                                  self.container,
                                                                  self.method,
                                                                  self.invocation_context,
                                                                  self.callback_type,
                                                                  self.control)
    
    def copy(self):
        cb = Callback(self.method, self.callback_type, self.invocation_context, **self._callback_args)
        cb.container = self.container
        return cb
    
    
    def xml(self):
        if self.callback_type:
            return self.callback_type.xml()
        else:
            return self.__class__.__name__
    


class CallbackLazy(Callback):
    '''
    Same as Callback, but imports module only on first use
    '''

    def __init__(self, *args, **kwargs):
        super().__init__(self._load_and_execute, **kwargs)
        
        # self.container = None
        # self.control = None
        # self.callback_type = None

        # self._callback_args = { key:kwargs.pop(key) for key in list(kwargs.keys()) if key in self.CALLBACK_KEYS }

        self.module_name = None
        self.method_name = None

        if len(args) == 3:
            self.module_name, self.container, self.method_name = args
        elif len(args) == 2:
            self.module_name, self.method_name = args
            self.container = None
        else:
            raise AttributeError('not enough arguments')

        self._module = None
        self._method = None

        # self.init_method(self._load_and_execute, **kwargs)

    def __repr__(self):
        return '<%s container=%s, method=%s, invocation_context=%s, callback=%s, control=%s>' % (type(self).__name__,
                                                                  self.container,
                                                                  self.method_name,
                                                                  self.invocation_context,
                                                                  self.callback_type,
                                                                  self.control)
    
    def copy(self):
        cb = CallbackLazy(self.module_name, self.container, self.method_name, **self._callback_args)
        cb.invocation_context = self.invocation_context
        cb.callback_type = self.callback_type
        return cb
    
    def _load_and_execute(self, **kwargs):
        ''' load module and requested method in module '''
        logging.debug('CallbackLazy._load_and_execute')
        try:
            if not self._method:
                self._import_module()
                if self.container is None:
                    self._method = getattr(self._module, self.method_name)
                else:
                    class_ = getattr(self._module, self.container)
                    self._method = getattr(class_, self.method_name)
            
            return self._method(**kwargs)
            
        except:
            logging.exception("error in contextdialog window creation")

    def _import_module(self):
        '''
        equivalent to: import <<module_name>>
        will not reload if module was already loaded
        '''
        if not self._module:
            logging.debug('CallbackLazy._import_module importing %s' % self.module_name)
            #do an import equivalent to:  import <<module_name>>
            self._module = importlib.import_module(self.module_name)



def WpfActionCallback(function):
    '''
    This decorator can be used to convert function for WPF windows into a BKT callback. This way
    invalidate and begin/end undo is automatically handled as for other callbacks. The window 
    needs to have a _context attribute.
    '''
    @wraps(function)
    def wrapper(self,*args,**kwargs):
        if hasattr(self, "_context") and self._context is not None:
            # print "Doing something with self.var1==%s" % self.var1
            method = Callback(lambda: function(self,*args,**kwargs), CallbackTypes.wpf_action)
            return_value = self._context.app_callbacks.invoke_callback(self._context, method)
            self._context.python_addin.invalidate_ribbon()
            return return_value
        else:
            logging.error("no context found; cannot convert wpf function into wpf action callback")
            return function(self,*args,**kwargs)
    return wrapper






class InvocationContext(object):
    '''
    Instances of this class encode information about the invocation context of a target function.
    Usually, the target function is some custom addin logic which is triggered by the user
    (e.g. clicking a button, entering information in a text field). The context information enables
    event dispatching code to to construct the actual execution context of the target function
    (usually the arguments passed to the target function). 
    '''
    def __init__(self, raise_error=True, **kwargs):
        # generic
        self.python_addin = False
        self.ribbon_id = False
        self.customui_control = False
        self.context = False
        self.application = False
        self.current_control = False
        
        # powerpoint/visio
        self.shapes = False
        self.shape = False
        self.shapes_min = None
        self.shapes_max = None
        self.wrap_shapes = False

        # powerpoint/excel/visio
        self.selection = False
        
        # powerpoint
        self.slide_of_shapes = False
        self.slides = False
        self.slide = False
        self.slides_min = None
        self.slides_max = None
        
        self.presentation = False
        self.require_text = False
        
        # visio
        self.page = False
        self.page_shapes = False

        # excel
        self.workbook = False
        self.sheet = False
        self.require_worksheet = False
        self.sheets = False
        self.selected_sheets = False
        self.cell = False
        self.cells = False
        self.areas = False
        self.areas_min = None
        self.areas_max = None
        
        for key in kwargs:
            if hasattr(self, key):
                setattr(self, key, kwargs[key])
            elif raise_error:
                raise AttributeError("%s has no attribute '%s' " % (type(self).__name__,key))
    
    def copy(self):
        ctx = InvocationContext()
        for key in [ 'python_addin', 'ribbon_id', 'customui_control', 'context', 'application', 'current_control', 'shapes', 'shape', 'shapes_min', 'shapes_max', 'wrap_shapes', 'selection', 'slide_of_shapes', 'slides', 'slide', 'slides_min', 'slides_max', 'presentation', 'require_text', 'page', 'page_shapes', 'workbook', 'sheet', 'require_worksheet', 'sheets', 'selected_sheets', 'cell', 'cells', 'areas', 'areas_min', 'areas_max' ]:
            setattr(ctx, key, getattr(self, key))
        return ctx
    
    @staticmethod
    def from_method(method):
        ''' Alternative constructor. Derives InvocationContext-settings from methods's parameter-names  '''
        kwargs = {}
        for var_name in list(method.__code__.co_varnames)[:method.__code__.co_argcount]:
            if var_name != "self":
                kwargs[var_name] = True
        
        return InvocationContext(raise_error=False, **kwargs)
        
        
        
        
        
        