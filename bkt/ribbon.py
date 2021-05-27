# -*- coding: utf-8 -*-
'''
Standard ribbon controls and bkt-specific ribbon controls

Created on 17.11.2014
@author: cschmitt
'''

from __future__ import absolute_import

import logging
import uuid #for getting random id

from itertools import count

import bkt.helpers as _h #for snake-to-camelcase
import bkt.library.system as lib_sys #for getting key-states in spinner

from bkt.callbacks import CallbackTypes, CallbackType, Callback
from bkt.xml import RibbonXMLFactory, linq

# Abhängigkeit zu annotation nicht gewünscht, siehe RibbonControl-Klasse
# from .annotation import AbstractAnnotationObject

from bkt import dotnet
Drawing = dotnet.import_drawing()
Bitmap = Drawing.Bitmap



class ArgAccessor(object):
    ''' allow acces to dictionary values via attributes, i.e. dict.key instead of dict['key'] '''
    def __init__(self, attributes):
        self._attributes = attributes
        
    def __getattr__(self, attr):
        try:
            return self._attributes[attr]
        except KeyError:
            raise AttributeError(attr)
        
    def __contains__(self, key):
        return key in self._attributes


#FIXME: Nutzung von AbstractAnnotationObject führt zu Abhängigkeit vom annotation-Modul.
#       Aktuell ist diese Abhängigkeit notwendig, damit RibbonControl-Instanzen in Klassenattributen bei der Control-Erstellung einer
#       FeatureContainer-Klasse berücksichtigt werden.
#       Sauberer wäre, die Logik vom AbstractAnnotationObject von außerhalb zu injezieren; diese Logik wird hier nicht weiter verwendet.
# class RibbonControl(AbstractAnnotationObject):
class RibbonControl(object):
    ''' Base class to represent any element from MSCustomUI.
        Holds attributes of the xml-element and callbacks associated to the element.
        Attributes and callbacks can be accessed in dict-style (i.e. button['label']).
    '''
    
    # default_callback wird in Addin.fallback_get_enabled verwendet, um Controls nur dann enabled-state zu geben, wenn der Context der Callback-Methode gefüllt werden kann
    default_callback = None
    no_id = False
    _attributes = {}
    _id_attribute_key = "id"
    _auto_id_counter = count()
    _predefined_ids = set(["id_mso", "idMso", "id_q", "idQ"])

    #NOTE: DEPRECATED: this counter is only used to set target_order for legacy annotations syntax
    _order_counter = count()
    
    #def __init__(self, node_type, xml_name, id_tag=None, attributes=None, **kwargs):
    def __init__(self, xml_name, id_tag=None, attributes={}, **kwargs):
        # AbstractAnnotationObject.__init__(self)
        self.target_order = next(RibbonControl._order_counter)
        
        #self.node_type = node_type
        self.xml_name = xml_name
        
        # attributes-dict accessible in class-attribute-style, e.g. self.attributes.some_key
        self._attributes = {} #attributes or {}
        self.attributes = ArgAccessor(self._attributes)
        
        # default control to use in get_enabled fallback-callback
        self.default_callback_control = self
        
        # init children from kwargs, fallback is empty list
        self.children = kwargs.pop('children', [])
        
        # init callbacks from kwargs, fallback is empty dict
        self._callbacks = {}
        if kwargs.has_key('callbacks'):
            self._callbacks = kwargs.pop('callbacks', {})
        else:
            # auto-identify callbacks
            for ct in CallbackTypes.callback_map().keys():
                if hasattr(self, ct):
                    if callable(getattr(self, ct)):
                        cb = Callback(getattr(self, ct))
                        cb.callback_type = getattr(CallbackTypes, ct)
                        self.add_callback(cb)
                    
        
        # class attributes
        # initialize attributes from all attributes defined in class hierarchy
        attributes_from_classes = [ t.__dict__.get('_attributes') for t in type(self).mro()]
        attributes_from_classes.reverse()
        for cls_attr in attributes_from_classes:
            if cls_attr:
                self.set_attributes(**cls_attr)
        
        
        # process kwargs
        self.set_attributes(**attributes)
        self.set_attributes(**kwargs)
        
        # init id
        self.id_tag = id_tag
        self.uuid = self.pop_attr('uuid')
        if self._id_attribute_key in self._attributes:
            self.user_defined_id = True
            self.set_id()
        else:
            self.user_defined_id = False
            if not self.no_id:
                pre_id = self.check_predefined_ids()
                if pre_id is None:
                    self._attributes[self._id_attribute_key] = self.create_persisting_id()
                # elif isinstance(self, Tab):
                else:
                    logging.debug("Predefined ID found: %s (%s)" % (pre_id, self._attributes[pre_id]))
                    # enable callbacks for idMso tab (e.g. powerpoint contextual tabs)
                    self._id_attribute_key = pre_id
            
            #if not (self.no_id or ('id_mso' in self._attributes) or ('idMso' in self._attributes) or ('id_q' in self._attributes) or ('idQ' in self._attributes)):
                #self._attributes[self._id_attribute_key] = self.create_persisting_id()
                #self._attributes[self._id_attribute_key] = ("_" + (self.uuid or str(uuid.uuid4()))).replace('-','_')
        
        # allow subclasses to do other stuff
        self.on_create()
    
    
    
    def set_attributes(self, **kwargs):
        ''' sets class-attributes, control-attributes and callbacks '''
        for key, value in kwargs.iteritems():
            if isinstance(value, Callback):
                # add callback
                if value.callback_type is None or value.callback_type.python_name is None:
                    # use fallback callback-type
                   value.callback_type = getattr(CallbackTypes, key)
                self.add_callback(value)
                
            elif hasattr(self, key):
                # set class-attribute
                setattr(self, key, value)
                
            else:
                # set control-attribute
                self.set_control_attributes(**{key:value})
        
    
    def set_control_attributes(self, **kwargs):
        ''' set attributes of ribbon control '''
        for key, value in kwargs.iteritems():
            self._attributes[key] = value
    
    
    @property
    def id(self):
        try:
            return self._attributes[self._id_attribute_key]
        except:
            return str(uuid.uuid4()).replace('-','_')
    
    @id.setter
    def id(self, value):
        self._attributes[self._id_attribute_key] = value
    
    
    def on_create(self):
        pass
    

    def check_predefined_ids(self):
        pre_id = set(self._attributes.keys()) & self._predefined_ids
        if len(pre_id) > 0:
            return pre_id.pop()
        else:
            return None

    
    def create_persisting_id(self):
        ''' generates an id, which will be identical on every addin-start '''
        return "_auto_id_" + str(next(RibbonControl._auto_id_counter))
        
    
    def set_id(self, fallback_id=None, ribbon_short_id=None, id_tag=None):
        id_tag = id_tag or self.id_tag
        
        if self.no_id:
            # don't use ids for elements such as <ribbon> or <tabs>
            if self._id_attribute_key in self._attributes:
                del self._attributes[self._id_attribute_key]
            return
        
        # if 'id_mso' in self._attributes or 'idMso' in self._attributes or 'id_q' in self._attributes or 'idQ'  in self._attributes:
        if self.check_predefined_ids() != None:
            # don't use ids for mso-controls or idq-controls
            return
        
        if self.uuid:
            # use uuid if available
            control_id = '_' + self.uuid
        elif self.user_defined_id and self._id_attribute_key in self._attributes:
            # use attributes.id if available and defined by user
            control_id = self.id
        elif not fallback_id is None:
            # use fallback id
            control_id = fallback_id
        else:
            # create new id
            control_id = self.create_persisting_id()
        
        # add ribbon-id and id-tag
        if ribbon_short_id is not None:
            control_id = ribbon_short_id + '_' + control_id
        if id_tag is not None:
            control_id += '_' + id_tag
        
        # save id as argument, so it is used on xml-generation
        self._attributes[self._id_attribute_key] = control_id
        
        # apply set_id on children
        for child in self.children:
            if isinstance(child, RibbonControl):
                child.set_id(ribbon_short_id=ribbon_short_id, id_tag=id_tag)
        
        
        
    def __getattr__(self, attr):
        ''' access callbacks as properties, e.g. button.on_action '''
        if self._callbacks.has_key(attr):
            return self._callbacks[attr]
        else:
            raise AttributeError(attr)
    
    def __getitem__(self, arg):
        ''' access ribbon-attributes in dict-style, e.g. button['label'] '''
        return self._attributes[arg]
    
    def __setitem__(self, arg, value):
        ''' access ribbon-attributes in dict-style, e.g. button['label'] = 'foo' '''
        if arg is None or value is None:
            raise ValueError
        
        if isinstance(value, Callback):
            # apply naming conventions on callback-names
            value.callback_type.set_attribute(arg)
            self._callbacks[arg] = value
        else:
            self._attributes[arg] = value
        

    def pop_attr(self, attr):
        ''' Returns value of 'attr' from ribbon attributes (self[attr]) and removes the key 'attr' from the attributes-dictionary (del self._attributes[attr]). '''
        return self.pop_attrs(attr)[0]
    
    def pop_attrs(self, *attrs):
        ''' Same as pop_attr for multiple keys. '''
        result = []
        for attr in attrs:
            v = None
            if attr in self._attributes:
                v = self._attributes[attr]
                del self._attributes[attr]
            result.append(v)
        return result
    
    
    
    def get_attributes_xml_dict(self):
        '''
        returns dictionary for the controls attributes for usage in xml-representation of control.
        
        keys are converted as follows:
            underscore to camelcase: my_key --> myKey
            
        values are converted as follows:
            boolean to strings: 'true', 'false'
            objects x to x.xml() or str(x)
        '''
        
        return {key:value for key, value in self._attributes.iteritems() if value != None}
        
        
        
    
    
    def xml(self):
        ''' Returns xml-representation of the element.
            Attaches all attributes to the xml-node and uses the CallbackType (i.e. 'onAction') of every Callback.
        '''
        f = RibbonXMLFactory()
        #print('RibbonControl._attributes: ' + str(self._attributes))
        node = f.node(self.xml_name, **convert_dict_to_ribbon_xml_style(self.get_attributes_xml_dict()))
        for cb in self._callbacks.values():
            if not cb.callback_type.custom:
                node.SetAttributeValue(cb.callback_type.xml_name, cb.callback_type.dotnet_name)
        
        if hasattr(self, 'children'):
            for child in self.children:
                if child:
                    node.Add(child.xml())
        return node
    
    def xml_string(self):
        #FIXME: get xml-string without hacking RibbonXMLFactory.namespace
        ns = RibbonXMLFactory.namespace
        RibbonXMLFactory.namespace = ""
        xmlstring = RibbonXMLFactory.to_normalized_string(self.xml())
        RibbonXMLFactory.namespace = ns
        return xmlstring
    
    def __repr__(self):
        return '<%s uuid=%s, id=%s>' % (type(self).__name__, self.uuid, self.id)
    
    
    def add_callback(self, original_callback, callback_type=None):
        ''' adds a callback to the element'''
        if isinstance(original_callback, CallbackType):
            # FIXME: is this ever used? should fail in xml-method
            callback = original_callback
            callback = callback_type or callback
            self._callbacks[callback.python_name] = callback
            callback.control = self
        elif isinstance(original_callback, Callback):
            callback = original_callback.copy()
            callback.callback_type = callback_type or callback.callback_type
            if not callback.callback_type is None:
                self._callbacks[callback.callback_type.python_name] = callback
                callback.control = self
            else:
                raise ValueError('unexpected callback w/o callback_type: ' + str(callback))
        else:
            raise ValueError('unexpected type in add_callback: ' + str(type(original_callback)))
    
    
    def collect_callbacks(self):
        ''' returns callback-map of the element and all its children '''
        stack = [self]
        res = []
        while stack:
            current = stack.pop()
            if isinstance(current, RibbonControl):
                res.extend(current._callbacks.values())
                stack.extend(current.children)
                # RibbonControls in attribute-values
                for ctrl in [attr for attr in current._attributes.values() if isinstance(attr, RibbonControl)]:
                    stack.extend([ctrl])
        return res



# ========================================
# = Default Ribbon Control in MSCustomUI =
# ========================================

class TypedRibbonControl(RibbonControl):
    ''' RibbonControl which initializes with xml-name defined as class attribute _xml_name '''
    def __init__(self, *args, **kwargs):
        RibbonControl.__init__(self, self._xml_name, *args, **kwargs)

class ActionRibbonControl(TypedRibbonControl):
    default_callback = CallbackTypes.on_action

    ''' TypedRibbonControl with default get_enabled callback '''
    def on_create(self):
        if not ('id_mso' in self._attributes or 'idMso' in self._attributes):
            pass
            #self._attributes['get_enabled'] = CallbackTypes.get_enabled.dotnet_name


def create_ribbon_control_class(python_name, base_cls=TypedRibbonControl, cls_name=None, xml_name=None, attributes={}):
    # example: toggle_button --> ToggleButton
    cls_name = cls_name or ''.join([x[0].upper() + x[1:] for x in python_name.split('_')])
    # example: ToggleButton --> toggleButton
    xml_name = xml_name or (cls_name[0].lower() + cls_name[1:])
    # create new class at runtime
    class_attributes = {'_python_name':python_name, '_xml_name': xml_name}
    class_attributes.update(attributes)
    return type(cls_name, (base_cls,), class_attributes)


#CustomUI    = create_ribbon_control_class('custom_ui', attributes={'no_id':True}, xml_name="customUI")
Ribbon      = create_ribbon_control_class('ribbon',    attributes={'no_id':True})
Tabs        = create_ribbon_control_class('tabs',      attributes={'no_id':True})
Tab         = create_ribbon_control_class('tab')
ContextualTabs = create_ribbon_control_class('contextual_tabs', attributes={'no_id':True})
TabSet         = create_ribbon_control_class('tab_set')
Qat              = create_ribbon_control_class('qat', attributes={'no_id':True})
DocumentControls = create_ribbon_control_class('document_controls', attributes={'no_id':True})
SharedControls   = create_ribbon_control_class('shared_controls', attributes={'no_id':True})

Group       = create_ribbon_control_class('group')
Box         = create_ribbon_control_class('box')
Menu        = create_ribbon_control_class('menu')
SplitButton = create_ribbon_control_class('split_button')

Control      = create_ribbon_control_class('control')
Button       = create_ribbon_control_class('button',        base_cls=ActionRibbonControl)
ToggleButton = create_ribbon_control_class('toggle_button', base_cls=ActionRibbonControl, attributes={'default_callback':CallbackTypes.on_toggle_action})
CheckBox     = create_ribbon_control_class('check_box',     base_cls=ActionRibbonControl, attributes={'default_callback':CallbackTypes.on_toggle_action})
EditBox      = create_ribbon_control_class('edit_box',      base_cls=ActionRibbonControl, attributes={'default_callback':CallbackTypes.on_change})
ComboBox     = create_ribbon_control_class('combo_box',     base_cls=ActionRibbonControl, attributes={'default_callback':CallbackTypes.on_change})
DropDown     = create_ribbon_control_class('drop_down',     base_cls=ActionRibbonControl, attributes={'default_callback':CallbackTypes.on_action_indexed})
GalleryMso   = create_ribbon_control_class('gallery',       base_cls=ActionRibbonControl, attributes={'default_callback':CallbackTypes.on_action_indexed})
DynamicMenu  = create_ribbon_control_class('dynamic_menu',  base_cls=ActionRibbonControl, attributes={'default_callback':CallbackTypes.get_content})

Label        = LabelControl = create_ribbon_control_class('label_control')
ButtonGroup  = create_ribbon_control_class('button_group')
Separator    = create_ribbon_control_class('separator')
MenuSeparator = create_ribbon_control_class('menu_separator')
Item         = create_ribbon_control_class('item')

ContextMenus = create_ribbon_control_class('context_menus',   attributes={'no_id':True})
ContextMenu  = create_ribbon_control_class('context_menu')
Commands     = CommandList = create_ribbon_control_class('commands',   attributes={'no_id':True})
Command      = create_ribbon_control_class('command')

#Special backstage elements: see https://msdn.microsoft.com/de-de/library/office/ee691833(v=office.14).aspx and https://msdn.microsoft.com/de-de/library/office/ee815851(v=office.14).aspx
Backstage        = create_ribbon_control_class('backstage',      attributes={'no_id':True}) # > tab, button
#Tab > firstColumn, secondColumn
FirstColumn      = create_ribbon_control_class('first_column',   attributes={'no_id':True})
SecondColumn     = create_ribbon_control_class('second_column',  attributes={'no_id':True})
PrimaryItem      = create_ribbon_control_class('primary_item',   attributes={'no_id':True}) # > button, menu
TopItems         = create_ribbon_control_class('top_items',      attributes={'no_id':True})
BottomItems      = create_ribbon_control_class('bottom_items',   attributes={'no_id':True})
LayoutContainer  = create_ribbon_control_class('layout_container') # > button, checkbox, editBox, dropdown, radioGroup, comboBox, hyperlink, labelControl, groupBox, layoutContainer, imageControl
#Menu > menuGroup
MenuGroup        = create_ribbon_control_class('menu_group') # > button, checkbox, menu
#Group > primaryItem, topItems, bottomItems
GroupBox         = create_ribbon_control_class('group_box') # > button, checkbox, editBox, dropdown, radioGroup, comboBox, hyperlink, labelControl, groupBox, layoutContainer, imageControl
RadioGroup       = create_ribbon_control_class('radio_group') # > radioButton
RadioButton      = create_ribbon_control_class('radio_button')
TaskGroup        = create_ribbon_control_class('task_group') # > category
TaskFormGroup    = create_ribbon_control_class('task_form_group') # > category
Category         = create_ribbon_control_class('category') # > task
Task             = create_ribbon_control_class('task') # > group
Hyperlink        = create_ribbon_control_class('hyperlink')
Image            = ImageControl = create_ribbon_control_class('image_control')



class DialogBoxLauncher(Button):
    _python_name = 'dialog_box_launcher'
    
    def xml(self):
        f = RibbonXMLFactory()
        node = f.node('dialogBoxLauncher')
        node.Add( super(DialogBoxLauncher, self).xml() )
        return node



class CustomUI(TypedRibbonControl):
    _python_name = 'custom_ui'
    _xml_name = 'customUI'
    no_id=True
    
    def xml(self):
        node = super(CustomUI, self).xml()
        # FIXME: namespaces variabel definierbar machen
        node.Add(linq.XAttribute(linq.XNamespace.Xmlns + "nsBKT", "http://www.business-kasper-toolbox.com/toolbox"))
        return node



# ===============================
# = Specialized Ribbon Controls =
# ===============================

class Gallery(GalleryMso):
    @staticmethod
    def get_check_image(checked=True):
        if not checked:
            return None

        size = 32
        img = Bitmap(size, size)
        g = Drawing.Graphics.FromImage(img)

        text_brush = Drawing.Brushes.DimGray
        strFormat = Drawing.StringFormat()
        strFormat.Alignment = Drawing.StringAlignment.Center
        strFormat.LineAlignment = Drawing.StringAlignment.Center
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAlias
        g.DrawString(u"\uE10B",
                    Drawing.Font("Segoe UI Symbol", 24, Drawing.GraphicsUnit.Pixel), text_brush,
                    # Drawing.RectangleF(2, 3, size, size),
                    Drawing.RectangleF(1, 2, size, size-1), 
                    strFormat)

        return img


class SpinnerBox(Box):
    _python_name = 'spinner_box'
    _xml_name = 'box'
    
    def __init__(self, **user_kwargs):
        
        # initialize children
        self.image_element = user_kwargs.pop('image_element', None) #image can be user defined element (e.g. button); otherwise image is given to textbox
        self.txt_box = EditBox()
        self.inc_button = Button(label=u"»")
        self.dec_button = Button(label=u"«")

        # default attributes

        kwargs = { 'size_string': '####' }
        if self.image_element is not None:
            # kwargs['children'] = [self.image_element, self.txt_box, self.dec_button, self.inc_button]
            # yet another box required to avoid space between image element and edit box
            self.inner_box = Box(children=[self.txt_box, self.dec_button, self.inc_button])
            kwargs['children'] = [self.image_element, self.inner_box]
        else:
            kwargs['children'] = [self.txt_box, self.dec_button, self.inc_button]
        kwargs.update(user_kwargs or {})
        
        # init Box-control
        super(SpinnerBox, self).__init__(**kwargs)
        
        # route fallback callbacks to textbox
        self.inc_button.default_callback_control = self.txt_box
        self.dec_button.default_callback_control = self.txt_box
        
        
    def set_control_attributes(self, **kwargs):
        button_args = { key: kwargs[key] for key in ['get_enabled', 'get_visible', 'screentip', 'supertip'] if key in kwargs }
        
        # if image element is used, some control-attributes are passed to image element
        if self.image_element is not None:
            image_args = { key: kwargs.pop(key) for key in ['label', 'show_label', 'image', 'image_mso', 'show_image'] if key in kwargs }
            image_args.update(button_args)
            if type(self.image_element) == SplitButton:
                self.image_element.children[0].set_control_attributes(**image_args)
            else:
                self.image_element.set_control_attributes(**image_args)
            # avoid space before textbox
            kwargs['show_label'] = False

        # other control-attributes are passed to edit-box
        self.txt_box.set_control_attributes(**kwargs)
        
        if len(button_args) > 0:
            self.inc_button.set_control_attributes(**button_args)
            self.dec_button.set_control_attributes(**button_args)
        
        
    
    def set_id(self, fallback_id=None, ribbon_short_id=None, id_tag=None):
        # init box-id
        Box.set_id(self,
                   fallback_id=fallback_id,
                   ribbon_short_id=ribbon_short_id,
                   id_tag=id_tag)
        # use box-id for children
        box_id = self.attributes.id
        if self.image_element is not None:
            self.image_element['id']   = box_id + '_image'
            self.inner_box['id']       = box_id + '_innerbox'
        self.txt_box['id']    = box_id + '_text'
        self.inc_button['id'] = box_id + '_increment'
        self.dec_button['id'] = box_id + '_decrement'

    def add_callback(self, callback):
        if callback.callback_type.python_name == 'increment':
            # pass increment-callback to button
            self.inc_button.add_callback(callback, callback_type=CallbackTypes.on_action)
        elif callback.callback_type.python_name == 'decrement':
            # pass decrement-callback to button
            self.dec_button.add_callback(callback, callback_type=CallbackTypes.on_action)
        elif callback.callback_type.python_name in ['get_enabled', 'get_visible']:
            # pass enabled/visible-callback to all children
            self.inc_button.add_callback(callback)
            self.dec_button.add_callback(callback)
            self.txt_box.add_callback(callback)
            if self.image_element is not None:
                self.image_element.add_callback(callback)
        elif self.image_element is not None and callback.callback_type.python_name in ['on_action', 'on_toggle_action']:
            # pass on_action to image element
            self.image_element.add_callback(callback)
        else:
            # pass other callbacks to editbox
            self.txt_box.add_callback(callback)


class RoundingSpinnerBox(SpinnerBox):
    _python_name = 'rounding_spinner_box'
    _xml_name = 'box'
    
    def __init__(self, **user_kwargs):
        ''' constructor.
            initializes big_step/small_step and default-step/round-settings
        '''
        # initialize class-attributes
        self.huge_step = 10
        self.big_step = 3
        self.small_step = 1
        self.round_at = None
        self.rounding_factor = None
        self.convert = None
        self.reset_value = 0

        self.ambiguous_text = None #u"\u2260"
        self.ambiguous_fallback_value = 0
        
        # default attributes
        kwargs = { 'size_string': '####' }

        # default-param-sets
        if user_kwargs.pop('round_cm', False):
            kwargs['huge_step'] = 1
            kwargs['big_step'] = 0.2
            kwargs['small_step'] = 0.1
            kwargs['rounding_factor'] = 0.1 #for setting values
            kwargs['round_at'] = 2 #for displaying values
        if user_kwargs.pop('round_pt', False):
            kwargs['huge_step'] = 10
            kwargs['big_step'] = 3
            kwargs['small_step'] = 1
            kwargs['round_at'] = 0
        if user_kwargs.pop('round_int', False):
            kwargs['huge_step'] = 10
            kwargs['big_step'] = 5
            kwargs['small_step'] = 1
            kwargs['round_at'] = 0
        
        if user_kwargs.pop('image_button', False):
            kwargs['image_element'] = Button()
            self.reset_value = user_kwargs.pop('reset_value', 0)
        
        #update with user attributes
        #user can also overwrite default-params in param-sets (e.g. small-step, big-step)
        kwargs.update(user_kwargs)
        
        # init SpinnerBox
        #print "init RoundingSpinnerBox with kwargs " + str(kwargs)
        super(RoundingSpinnerBox, self).__init__(**kwargs)
        self.init_spinner_callbacks()


    def add_callback(self, callback):
        ''' overwrites RibbonControl.add_callback.
            wraps on_change/get_text-methods with converter-methods and initializes spinner-callbacks (see init_spinner_callbacks-method)
        '''
        #print 'RoundingSpinnerBox.add_callback: ' + str(callback)
        cb = callback

        #print 'RoundingSpinnerBox.add_callback: self.convert=' + str(self.convert)
        #if self.convert != None and cb.callback_type in [CallbackTypes.on_change, CallbackTypes.get_text]:
        # for on_change and get_text-callbacks:
        # copy callback and wrap methods with converter-method

        if cb.callback_type == CallbackTypes.on_change:
            #print 'RoundingSpinnerBox.add_callback: alter on_change'
            # convert value, then call writer method
            cb = cb.copy()
            old_method = cb.method
            cb.method = lambda value, **kwargs: old_method(value=self.convert_before_write(value), **kwargs)

        elif cb.callback_type == CallbackTypes.get_text:
            #print 'RoundingSpinnerBox.add_callback: alter get_text'
            # call get_text-method and convert the returned value
            cb = cb.copy()
            old_method = cb.method
            cb.method = lambda **kwargs: self.convert_after_read(old_method(**kwargs))

        super(RoundingSpinnerBox, self).add_callback(cb)
        self.init_spinner_callbacks()

    def init_spinner_callbacks(self):
        ''' initializes the increment- and decrement-callback, if the editbox's on_change- and get_text-callback are defined.
            adds self._inc and self._dec as callbacks to the inc/dec-button
        '''
        if not ( self.txt_box._callbacks.has_key('get_text') and self.txt_box._callbacks.has_key('on_change') ):
            return

        invocation_context = self.txt_box._callbacks['on_change'].invocation_context.copy()
        invocation_context.context=True

        # add increment/decrement-Callbacks
        super(RoundingSpinnerBox, self).add_callback( Callback(self._dec, CallbackTypes.decrement, invocation_context) )
        super(RoundingSpinnerBox, self).add_callback( Callback(self._inc, CallbackTypes.increment, invocation_context) )

        # add reset-Callback is image_element is button and does not have on_action-Callback already
        if self.image_element is not None and type(self.image_element) == Button and not self.image_element._callbacks.has_key('on_action'):
            super(RoundingSpinnerBox, self).add_callback( Callback(self._res, CallbackTypes.on_action, invocation_context) )


    def _round(self, value):
        ''' rounding-function using the control's rounding-settings.
            either rounds value to a multiple of self.rounding_factor, e.g. multiples of .25
            or rounds the value to a multiple of 10^(-self.round_at), e.g. multiples of 0.1, 1, 10, ...
        '''
        if self.rounding_factor != None:
            return round(float(value) / self.rounding_factor) * self.rounding_factor
        elif self.round_at != None:
            return round(value, self.round_at)
        else:
            return value

    def _get(self, context):
        value = context.invoke_callback(self.txt_box._callbacks['get_text'])
        if value == self.ambiguous_text:
            value = self.ambiguous_fallback_value
        return self._round(value)

    def _dec(self, context, **kwargs):
        ''' decrement-callback using the on_change/get_text-callback from the editbox '''
        value = self._get(context)
        ctrl_pressed = lib_sys.get_key_state(lib_sys.KeyCodes.CTRL)
        shift_pressed = lib_sys.get_key_state(lib_sys.KeyCodes.SHIFT)
        step = self.big_step if not ctrl_pressed else self.small_step
        step = step if not shift_pressed else self.huge_step
        context.invoke_callback(self.txt_box._callbacks['on_change'], value=value-step)

    def _inc(self, context, **kwargs):
        ''' increment-callback using the on_change/get_text-callback from the editbox '''
        value = self._get(context)
        ctrl_pressed = lib_sys.get_key_state(lib_sys.KeyCodes.CTRL)
        shift_pressed = lib_sys.get_key_state(lib_sys.KeyCodes.SHIFT)
        step = self.big_step if not ctrl_pressed else self.small_step
        step = step if not shift_pressed else self.huge_step
        context.invoke_callback(self.txt_box._callbacks['on_change'], value=value+step)

    def _res(self, context, **kwargs):
        ''' reset-callback using the on_change/get_text-callback from the editbox '''
        # ctrl_pressed = lib_sys.get_key_state(lib_sys.KeyCodes.CTRL)
        shift_pressed = lib_sys.get_key_state(lib_sys.KeyCodes.SHIFT)
        value = self.big_step if shift_pressed else self.reset_value
        context.invoke_callback(self.txt_box._callbacks['on_change'], value=value)

    def convert_after_read(self, value):
        ''' general convert-function using the control's convert-setting.
            converts the value after calling get_text-callback
        '''
        ambiguous = False
        
        if value is None:
            return None
        elif type(value) is tuple:
            #tuple with 2 values ambiguous bool and fallback value
            assert len(value) == 2, "Ambiguity tuple must have exactly two values"
            ambiguous, value = value
        elif type(value) is list:
            # list means ambiguous values
            value_0 = value[0]
            if not value or value_0 is None:
                return None
            elif value.count(value_0) == len(value):
                # https://stackoverflow.com/questions/3844801/check-if-all-elements-in-a-list-are-identical
                value = value_0 # all values are the same, use first value
            else:
                value = value_0 #fallback value for dec/inc methods
                ambiguous = True

        convert_func_name = 'convert_' + str(self.convert) + '_A'
        convert_func = getattr(self, convert_func_name, lambda value: value)
        
        if ambiguous:
            self.ambiguous_fallback_value = convert_func(value)
            return self.ambiguous_text
        else:
            return convert_func(value)

    def convert_before_write(self, value):
        ''' general convert-function using the control's convert-setting.
            converts the value before calling on_change-callback
        '''
        if type(value) == str:
            try:
                value = float(value.replace(',', '.'))
            except:
                pass
        
        convert_func_name = 'convert_' + str(self.convert) + '_B'
        convert_func = getattr(self, convert_func_name, lambda value: value)
        return convert_func(value)


    ####  functions to convert pt to cm values
    pt_to_cm_factor = 2.54 / 72

    def convert_pt_to_cm_A(self, pt):
        ''' convert pt-value to cm-value, round to 4 digits '''
        round_at = self.round_at if self.round_at != None else 4
        return round(float(pt) * self.pt_to_cm_factor, round_at)

    def convert_pt_to_cm_B(self, cm):
        ''' convert cm-value to pt-value '''
        return float(cm) / self.pt_to_cm_factor




class ColorGallery(Gallery):
    item_size = 14
    color_helper = None
    
    def __init__(self, color_helper=None, **user_kwargs):
        # default attributes
        kwargs = {
            'show_item_label': 'false',
            'columns': 10,
            'on_action_indexed':  Callback(self.on_action_indexed, context=True),
            'get_item_count':     Callback(self.get_item_count,    context=True),
            'get_item_image':     Callback(self.get_item_image,    context=True),
            # 'get_item_screentip': Callback(self.get_item_name),
            'get_item_label':     Callback(self.get_item_name,     context=True),
            # 'get_selected_item_index':  Callback(self.get_selected_item_index, context=True),
            'item_width': self.item_size,
            'item_height': self.item_size,
        }
        #only add callback 'get_selected_item_index' if 'get_selected_color' is provided to avoid UI error messages
        if "get_selected_color" in user_kwargs:
            kwargs["get_selected_item_index"] = Callback(self.cb_get_selected_item_index, context=True)
        kwargs.update(user_kwargs)
        
        if not 'image' in kwargs and not 'image_mso' in kwargs:
            kwargs['image_mso'] = 'SmartArtChangeColorsGallery'
        
        # init gallery attributes and callbacks
        # this also includes the custom ColorGallery-callbacks: on_rgb_color_change, on_theme_color_change, get_selected_color
        super(ColorGallery, self).__init__(**kwargs)
        
        # reset gallery_colors, later initialized by get_item_image
        self.gallery_colors = [[0,0,0] for _ in range(80)]
        self.theme_colors = [None]*60
        
        #allow to pass color helper to make this element also available for other office apps than powerpoint
        #powerpoint color helper is fallback for backwards compatibility
        if color_helper:
            self.color_helper = color_helper
        elif ColorGallery.color_helper:
            self.color_helper = ColorGallery.color_helper
        else:
            import bkt.library.powerpoint as pplib
            ColorGallery.color_helper = self.color_helper = pplib.ColorHelper #4 functions required: get_theme_color, get_theme_index, get_recent_color, get_recent_colors_count
    

    def on_action_indexed(self, selected_item, index, context, **kwargs):
        '''
            Method is called when a color was clicked. Invokes the event defined when instance was created.
        '''
        if 'on_theme_color_change' in self._callbacks and index < 60:
            # a theme color was selected
            # set color by theme-index/brightness
            if context:
                # callback aufrufen; vom callback benötigte Parameter werden durch invoke_callback aufgelöst
                return context.invoke_callback(self._callbacks['on_theme_color_change'], color_index=self.gallery_colors[index][0], brightness=self.gallery_colors[index][1] , **kwargs)
            else:
                # Kein context, um callback aufzurufen; versuche methode direkt aufzurufen
                return self._callbacks['on_theme_color_change'].method(color_index=self.gallery_colors[index][0], brightness=self.gallery_colors[index][1] , **kwargs)
        elif 'on_rgb_color_change' in self._callbacks :
            # set color by rgb-value
            rgb = self.gallery_colors[index][2]
            if rgb != None:
                if context:
                    # callback aufrufen; vom callback benötigte Parameter werden durch invoke_callback aufgelöst
                    return context.invoke_callback(self._callbacks['on_rgb_color_change'], color=rgb, **kwargs)
                else:
                    # Kein context, um callback aufzurufen; versuche methode direkt aufzurufen
                    return self._callbacks['on_rgb_color_change'].method(color=rgb, **kwargs)
            else:
                return None

    def get_item_count(self, context):
        # we use this callback to reset the cache which is important if design or presentation is changed
        self.theme_colors = [None]*60 #reset theme colors cache
        return 70 + self.recent_count(context)
    
    def cb_get_selected_item_index(self, context, **kwargs):
        #if self.get_selected_color == None:
        if not 'get_selected_color' in self._callbacks:
            return -1

        theme_brightness_rgb = context.invoke_callback(self._callbacks['get_selected_color'], **kwargs)
        
        if not theme_brightness_rgb:
            return -1
        
        if theme_brightness_rgb[0] == 0:
            # check rgb-value only
            gallery_colors_list = [ None if x is None or x[0] != 0 else x[2]  for x in self.gallery_colors]
            check = theme_brightness_rgb[2]

        else:
            gallery_colors_list = [ None if x is None else [x[0], int(x[1]*100)]  for x in self.gallery_colors]
            check = [theme_brightness_rgb[0], int(theme_brightness_rgb[1]*100)]

        try:
            return gallery_colors_list.index(check)
        except ValueError:
            return -1

    def get_item_image(self, context, index):
        '''
            Returns image as System.Drawing.Bitmap-object for the given gallery-item-index.
        '''
        rgb = None
        themecolor = 0
        brightness = 0

        # get rgb-value and brightness
        if index/10 == 0:
            # theme colors
            color = self.get_theme_color(context, index)
            rgb, themecolor = color.rgb, color.theme_index
        elif index/10 < 6:
            # theme color shades
            color = self.get_theme_color_shade(context, index, index/10-1)
            rgb, themecolor, brightness = color.rgb, color.theme_index, color.brightness
        elif index/10 == 6:
            # standard colors
            rgb = self.get_standard_color(index % 10)
        elif index/10 == 7:
            # recent colors
            rgb = self.get_recent_color(context, index % 10)

        # save values for item_action
        self.gallery_colors[index] = [themecolor, brightness, rgb]

        if rgb is None:
            return None
        else:
            return self.get_color_image(rgb)

    def get_item_name(self, context, index):
        '''
            Returns the label for the given gallery-item-index.
        '''
        if index/10 == 0:
            return self.get_theme_color_name(context, index)
        elif index/10 < 6:
            return self.get_theme_color_shade_name(context, index, index/10-1)
        elif index/10 == 6:
            return self.get_standard_color_name(index % 10)
        elif index/10 == 7:
            return 'recently used color'
        else:
            return ''


    #### Theme colors ####

    def get_theme_color(self, context, index):
        '''
            Returns Office-RGB-value for the given theme-color-index
        '''
        if not self.theme_colors[index]:
            self.theme_colors[index] = self.color_helper.get_theme_color(context, self.color_helper.get_theme_index(index))
        return self.theme_colors[index]

    def get_theme_color_name(self, context, index):
        '''
            Returns name for the given theme-color-index
        '''
        return self.get_theme_color(context, index).name


    #### Theme color shades ####

    def get_theme_color_shade(self, context, index, shade_index):
        '''
            Returns Office-RGB-value and brightness-factor for the given theme-color-index and shade-index.
            Note that RGB-values can differ from PowerPoint-RGB-values by ~1 (in each color-dimension)
        '''
        if not self.theme_colors[index]:
            self.theme_colors[index] = self.color_helper.get_theme_color(context, self.color_helper.get_theme_index(index), shade_index=shade_index)
        return self.theme_colors[index]

    def get_theme_color_shade_name(self, context, index, shade_index):
        '''
            Returns name for the given theme-color-index and shade-index
        '''
        return self.get_theme_color_shade(context, index, shade_index).name



    #### Standard colors ####
    _standard_colors = ['Dark Red', 'Red', 'Orange', 'Yellow', 'Light Green', 'Green', 'Light Blue', 'Blue', 'Dark Blue', 'Purple']
    _standard_colors_rgb = [192, 255, 49407, 65535, 5296274, 5287936, 15773696, 12611584, 6299648, 10498160]

    def get_standard_color(self, index):
        '''
            Returns Office-RGB-value for the given standard-color-index
        '''
        return self._standard_colors_rgb[index]

    def get_standard_color_name(self, index):
        '''
            Returns name for the given standard-color-index
        '''
        return self._standard_colors[index]


    #### Recent colors ####

    def get_recent_color(self, context, index):
        '''
            Returns Office-RGB-value for the given recent-color-index
        '''
        if index < self.recent_count(context):
            return self.color_helper.get_recent_color(context, index+1)
        else:
            return None, None

    def recent_count(self, context):
        '''
            Returns number of recent colors
        '''
        return self.color_helper.get_recent_colors_count(context)


    #### image creation ####

    def get_color_image(self, rgb):
        '''
            Returns image as System.Drawing.Bitmap-object for the given Office-RGB-value
        '''
        # Farbe on the fly erstellen
        # http://msdn.microsoft.com/en-us/library/aa287582(v=vs.71).aspx
        size = self.item_size*2
        img = Bitmap(size, size)
        g = Drawing.Graphics.FromImage(img)

        if rgb is not None:
            color = Drawing.ColorTranslator.FromOle(rgb)
            g.Clear(color)

        color_grey = Drawing.Color.FromArgb(230,230,230)
        pen = Drawing.Pen(color_grey,1)
        # pen.Alignment = Drawing.Drawing2D.PenAlignment.Inset
        g.DrawRectangle(pen, 0,0, size-1, size-1) #left, top, width, height

        # for x in range(0, img.Height):
        #     for y in range(0, img.Width):
        #         img.SetPixel(x, y, color);

        return img


class SymbolsGallery(Gallery):
    fallback_font = "Arial"
    item_size = 32
    
    def __init__(self, symbols=None, **user_kwargs):
        self.symbols = symbols or [] #list([font, symbol, screentip, supertip], [...])

        # default attributes
        kwargs = {
            'show_item_label': 'false',
            'columns': 6,
            # 'on_action_indexed':  Callback(self.on_action_indexed, context=True),
            # 'get_item_count':     Callback(self.get_item_count),
            # 'get_item_image':     Callback(self.get_item_image),
            # 'get_item_screentip': Callback(self.get_item_screentip),
            # 'get_item_supertip':  Callback(self.get_item_supertip),
            'get_image':                Callback(lambda: self.get_item_image(0) ),
            # 'get_selected_item_index':  Callback(self.get_selected_item_index, context=True),
            'item_width': self.item_size,
            'item_height': self.item_size,
        }
        #only add callback 'get_selected_item_index' if 'get_selected_symbol' is provided to avoid UI error messages
        if "get_selected_symbol" in user_kwargs:
            kwargs["get_selected_item_index"] = Callback(self.cb_get_selected_item_index, context=True)
        kwargs.update(user_kwargs)

        super(SymbolsGallery, self).__init__(**kwargs)
    

    def on_action_indexed(self, selected_item, index, context, **kwargs):
        '''
            Method is called when a symbol was clicked. Invokes the event defined when instance was created.
        '''
        # if 'on_symbol_change' in self._callbacks:

        if context:
            # callback aufrufen; vom callback benötigte Parameter werden durch invoke_callback aufgelöst
            return context.invoke_callback(self._callbacks['on_symbol_change'], symbol=self.symbols[index], **kwargs)
        else:
            # Kein context, um callback aufzurufen; versuche methode direkt aufzurufen
            return self._callbacks['on_symbol_change'].method(symbol=self.symbols[index], **kwargs)

    def get_item_count(self):
        return len(self.symbols)
    
    def get_item_image(self, index):
        ''' creates an item image with numberd shape according to settings in the specified item '''
        # retrieve item-settings
        item = self.symbols[index]
        font = item[0] or self.fallback_font #Fallback font
        return SymbolsGallery.create_symbol_image(font, item[1])

    def get_item_screentip(self, index):
        ''' creates an item image with numberd shape according to settings in the specified item '''
        # retrieve item-settings
        item = self.symbols[index]
        try:
            return item[2]
        except:
            return "Symbol einfügen"

    def get_item_supertip(self, index):
        ''' creates an item image with numberd shape according to settings in the specified item '''
        # retrieve item-settings
        item = self.symbols[index]
        try:
            return item[3]
        except:
            return "Fügt das Symbol in aktuellen Text oder neues Shape ein."
    
    def cb_get_selected_item_index(self, context, **kwargs):
        if not 'get_selected_symbol' in self._callbacks:
            return -1

        symbol = context.invoke_callback(self._callbacks['get_selected_symbol'], **kwargs)
        
        if not symbol:
            return -1
        
        symbols_list = [x[1] for x in self.symbols]

        if symbol in symbols_list:
            return symbols_list.index(symbol)
        else:
            return -1
        
    @staticmethod
    def create_symbol_image(font, text, fontsize=54, size=64):
        # define font and pen/brush
        font = Drawing.Font(font, fontsize, Drawing.GraphicsUnit.Pixel)
        text_brush = Drawing.Brushes.Black
        # estimate size
        if size is None:
            g = Drawing.Graphics.FromImage(Bitmap(1, 1))
            size_f = g.MeasureString(text, font)
            width  = size_f.Width
            height = size_f.Height
        else:
            width=height=size
        # set string format
        strFormat = Drawing.StringFormat()
        strFormat.Alignment = Drawing.StringAlignment.Center
        strFormat.LineAlignment = Drawing.StringAlignment.Center
        # create bitmap
        img = Bitmap(width, height)
        g = Drawing.Graphics.FromImage(img)
        # draw string
        # g.Clear(Color.White)
        # g.TextRenderingHint = Drawing.Text.TextRenderingHint.ClearTypeGridFit
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAlias
        g.DrawString(text,
                     font, text_brush,
                     Drawing.RectangleF(1, 2, width, height-1), 
                     strFormat)
        
        return img



# =======================
# = Ribbon MSO Controls =
# =======================

class MSOControl(RibbonControl):
    ''' describes mso controls to be reused in the custom ui'''
    # no_id = True
    
    def __init__(self, control_type, id_mso, **kwargs):
        if "attributes" in kwargs:
            kwargs["attributes"]["id_mso"] = id_mso
        else:
            kwargs["attributes"] = {"id_mso": id_mso}
        super(MSOControl, self).__init__(control_type, **kwargs)
        
    def configure(self, **kwargs):
        '''
        Generic method to change attributes. Returns self.
        Intention: The definition and modification of MSOControls using special factories
        Example: paste = mso.button.Paste.configure(show_label=True)
        '''
        self._attributes.update(kwargs)
        return self
        
    def __call__(self, **kwargs):
        self.configure(**kwargs)
        return self


class MSOFactory(object):
    ''' Creates MSOControls via attribute access. Pass default constructor arguments as keyword args.
        Example: MSOFactory.ShapesInsertGallery(control_type='control', show_label=False) '''
    def __init__(self, **kwargs):
        self._attributes = kwargs
    
    def __getattr__(self, attr):
        return MSOControl(id_mso=attr, **self._attributes)


class MSOFactoryAccess(object):
    ''' Similar to MSOFactory. Creates MSOControls via attribute access as control/group/button.
        Example: MSOFactoryAccess.control.ShapesInsertGallery(show_label=False) '''
    def __init__(self):
        self.group   = MSOFactory(control_type='group')
        self.control = MSOFactory(control_type='control', show_label=False)
        self.button  = MSOFactory(control_type='button', show_label=False)
    
    def __getattr__(self, attr):
        return MSOFactory(control_type=attr)


''' make factories available via mso.button, mso.group etc.'''
mso = MSOFactoryAccess()





# ===============================
# = escaping labels, tites, etc =
# ===============================


def escape_field(value, field=None):
    if field in ("label", "screentip"):
        return value.replace("&", "&&")

    #e.g. supertip, tag, description do not need escaping
    return value




# =======================================
# = dictionary and key/value conversion =
# =======================================



def convert_value_to_string(v, key=None):
    if v is True:
        return 'true'
    elif v is False:
        return 'false'
    elif isinstance(v, (str, unicode)):
        return escape_field(v, key)
    else:
        try:
            return v.xml()
        except:
            return str(v)

# def convert_key_to_lower_camelcase(key):
    # if not '_' in key:
    #     return key
    # parts = key.split('_')
    # parts_new = []
    # for i, part in enumerate(parts):
    #     if len(part) > 1:
    #         if i > 0:
    #             p = part[0].upper() + part[1:]
    #         else:            
    #             p = part[0].lower() + part[1:]
    #     else:
    #         if i > 0:
    #             p = part.upper()
    #         else:            
    #             p = part.lower()
            
    #     parts_new.append(p)
    # return ''.join(parts_new)

def convert_dict_to_ribbon_xml_style(d):
    return {_h.snake_to_lower_camelcase(k):convert_value_to_string(v, k) for k, v in d.iteritems()}


