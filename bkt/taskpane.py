# -*- coding: utf-8 -*-
'''
Taskpane controls

Created on 11.11.2019
@author: rdebeerst
'''

from __future__ import absolute_import

import logging

import bkt.ribbon
from bkt.helpers import Resources
from bkt.xml import WpfXMLFactory, linq



# ===========================
# = General Control Classes =
# ===========================

class TaskPaneControl(bkt.ribbon.TypedRibbonControl):
    _xml_name = 'TaskPaneControl'
    _id_attribute_key = "Name"
    xml_namespace = 'http://schemas.microsoft.com/winfx/2006/xaml/presentation'
    no_id=False

    def __init__(self, *args, **kwargs):
        # image attributes
        self.image = None
        self.large_image = None
        super(TaskPaneControl, self).__init__(*args, **kwargs)
    
    # def __init__(self, xml_name, *args, **kwargs):
    #     bkt.ribbon.RibbonControl.__init__(self, xml_name, *args, **kwargs)
    def collect_image_resources(self):
        stack = [self]
        result = []
        while stack:
            current = stack.pop()
            if isinstance(current, TaskPaneControl):
                if current.image != None:
                    result.append(current.image)
                if current.large_image != None:
                    result.append(current.large_image)
                stack.extend(current.children)
        return list(set(result))
    
    
    def wpf_xml(self):
        ''' Returns xml-representation of the element.
            Attaches all attributes to the xml-node and uses the CallbackType (i.e. 'onAction') of every Callback.
        '''
        f = WpfXMLFactory(namespace=self.xml_namespace)
        #print('RibbonControl._attributes: ' + str(self._attributes))
        node = f.node(self.xml_name, **convert_dict_to_ribbon_xml_style(self._attributes))
        
        # parse XamlPropertyElement-attributes
        for key,value in self._attributes.iteritems():
            if isinstance(value, XamlPropertyElement):
                value.xml_namespace = self.xml_namespace
                node.Add(value.wpf_xml(type_name=self.xml_name))
            elif isinstance(value, TaskPaneControl):
                property_node = XamlPropertyElement(
                    xml_namespace = self.xml_namespace,
                    type_name=self.xml_name,
                    property_name=convert_key_to_upper_camelcase(key)
                ).wpf_xml()
                property_node.Add(value.wpf_xml())
                node.Add(property_node)
        
        # FIXME: callbacks for task pane controls
        # for cb in self._callbacks.values():
        #     if not cb.callback_type.custom:
        #         node.SetAttributeValue(cb.callback_type.xml_name, cb.callback_type.dotnet_name)
        
        if hasattr(self, 'children'):
            for child in self.children:
                if child:
                    node.Add(child.wpf_xml())
        return node



class XmlPart(TaskPaneControl):
    ''' TaskPaneControl defined by xml-string '''
    
    def __init__(self, xml_string):
        self._xml_name = xml_string
        super(XmlPart, self).__init__()
        #self.xml_string = xml_string
        # dummy-element for xmlns-attributes
        document = linq.XDocument.Parse("<dummy xmlns=\"http://schemas.microsoft.com/winfx/2006/xaml/presentation\" xmlns:x=\"http://schemas.microsoft.com/winfx/2006/xaml\" xmlns:po=\"http://schemas.microsoft.com/winfx/2006/xaml/presentation/options\" xmlns:r=\"clr-namespace:System.Windows.Controls.Ribbon;assembly=System.Windows.Controls.Ribbon\" xmlns:fr=\"urn:fluent-ribbon\" xmlns:sys=\"clr-namespace:System;assembly=mscorlib\" >" + xml_string + "</dummy>")
        # obtain first child
        self.root = document.Root.FirstNode

    def wpf_xml(self):
        return self.root


class FluentRibbonControl(TaskPaneControl):
    _xml_name = 'FluentRibbonControl'
    _id_attribute_key = "Name"
    xml_namespace = 'urn:fluent-ribbon'
    no_id=False
    
    def __init__(self, *args, **kwargs):
        # screentip-attributes
        self.screentip = None
        if self._xml_name != 'ScreenTip':
            self.screentip_title = None
            self.screentip_image = None
            self.disable_reason = None
            self.help_topic = None
        
        super(FluentRibbonControl, self).__init__(*args, **kwargs)
        
    
    # def set_screentip(title=None, help_topic=None, text=None, image=None, disable_reason=None):
    #     pass
    
    def wpf_xml(self):
        ''' Returns xml-representation of the element.
            Attaches all attributes to the xml-node and uses the CallbackType (i.e. 'onAction') of every Callback.
        '''
        
        # Handle Screentip
        if self.screentip:
            self.pop_attr('ToolTip')
            self.pop_attr('tool_tip')
        
        # Create XML-Node
        node = super(FluentRibbonControl, self).wpf_xml()
        
        # Handle Images
        if self.image:
            if self._xml_name != 'ScreenTip':
                node.SetAttributeValue("Icon", "{StaticResource " +  self.image + "}")
            else:
                node.SetAttributeValue("Image", "{StaticResource " +  self.image + "}")
        if self.large_image:
            node.SetAttributeValue("LargeIcon", "{StaticResource " + self.large_image + "}")
        
        # Handle Screentip
        if self.screentip:
            tooltip = ToolTip(self.xml_name,
                children=[ FluentRibbon.ScreenTip(
                    text=self.screentip,
                    title=self.screentip_title,
                    image=self.screentip_image,
                    disable_reason=self.disable_reason,
                    help_topic=self.help_topic,
                    IsRibbonAligned="False"
                )]
            )
            
            node.Add( tooltip.wpf_xml() )
        
        return node
        
        
        
    


# =========================================
# = XAML Property Element Representations =
# =========================================


NOTSPECIFIED = u'NotSpecified'

class XamlPropertyElement(TaskPaneControl):
    _property_name = None
    _type_name = None
    
    def __init__(self, type_name=None, property_name=None, *args, **kwargs):
        self.type_name = type_name or type(self)._type_name or NOTSPECIFIED
        self.property_name = property_name or type(self)._property_name or NOTSPECIFIED
        self._xml_name = self.type_name + '.' + self.property_name
        self.no_id=True
        super(XamlPropertyElement, self).__init__(*args, **kwargs)
    
    def wpf_xml(self, type_name=None):
        type_name = type_name or self.type_name or NOTSPECIFIED
        self.xml_name = type_name + "." + self.property_name
        return super(XamlPropertyElement, self).wpf_xml()
        


class XamlPropertyElementGenerator(object):
    
    def __init__(self, xmlns=None):
        self.xmlns = xmlns or 'http://schemas.microsoft.com/winfx/2006/xaml/presentation'
    
    def __getattr__(self, attr):
        cls_name = "XamlPropertyElement_" + attr
        class_attributes = {
            'xml_namespace': self.xmlns,
            '_property_name': attr
        }
        return type(cls_name, (XamlPropertyElement,), class_attributes)




# Wpf
XamlPropertyElements = XamlPropertyElementGenerator()
Resources         = XamlPropertyElements.Resources
RowDefinition     = XamlPropertyElements.RowDefinition
ColumnDefinition  = XamlPropertyElements.ColumnDefinition
Filters           = XamlPropertyElements.Filters

# FluentRibbon
FluentRibbonPropertyElements = XamlPropertyElementGenerator(xmlns='urn:fluent-ribbon')
Menu              = FluentRibbonPropertyElements.Menu
ToolTip           = FluentRibbonPropertyElements.ToolTip
Icon              = FluentRibbonPropertyElements.Icon




# ==============================
# = Factory-access to controls =
# ==============================


class WpfControlFactory(object):
    '''
    factory class to create WPF-controls.
    WPF-control-classes are created by attribute access
    
    example:
        Wpf = WpfControlFactory()
        Wpf.ScrollViewer()
    
    '''
    def __init__(self, xmlns=None, base_control=None):
        self.xmlns = xmlns
        self.base_control = base_control or TaskPaneControl
        #self.xmlns_prefix = ( xmlns + ':' if xmlns else '')

    def __getattr__(self, attr):
        ''' access wpf-controls as properties, e.g. factory.SomeElementName '''
        cls_name = "WpfControl_" + attr
        xml_name = attr
        python_name = attr
        # create new class at runtime
        class_attributes = {
            '_python_name': python_name,
            '_xml_name': xml_name,
            'xml_namespace': self.xmlns
        }
        #class_attributes.update(attributes)
        return type(cls_name, (self.base_control,), class_attributes)
    
    def __getitem__(self, item):
        ''' access wpf-controls as dict-items, e.g. factory['SomeElementName'] '''
        return self.__getattr__(item)
    
    


Wpf = WpfControlFactory()

# xmlns:r="clr-namespace:System.Windows.Controls.Ribbon;assembly=System.Windows.Controls.Ribbon"
Ribbon = WpfControlFactory('clr-namespace:System.Windows.Controls.Ribbon;assembly=System.Windows.Controls.Ribbon')

FluentRibbon = WpfControlFactory(xmlns='urn:fluent-ribbon', base_control=FluentRibbonControl)





# ===============================
# = Specific Control Definition =
# ===============================


class BaseScrollViewer(TaskPaneControl):
    _xml_name = "ScrollViewer"
    
    def __init__(self, *args, **user_kwargs):
        self.image_resources = {}
        
        kwargs = dict(
            Margin="0",
            CanContentScroll="False",
            VerticalScrollBarVisibility="Auto" 
        )
        kwargs.update(user_kwargs)
        super(BaseScrollViewer, self).__init__(*args, **kwargs)
        
        # load scrollbar styling
        xaml_filename=Resources.xaml.locate("scrollbar_style")
        xml_string = ""
        for line in open(xaml_filename, "r"):
            xml_string += line + "\n"
        
        resources_node = XmlPart(xml_string)
        
        # add image resources
        for image_name, image_path in self.image_resources.items():
            logging.debug('image resource %s=%s' % (image_name, image_path))
            resources_node.root.Add(
                XmlPart('<BitmapImage x:Key="%s" UriSource="%s"/>' % (image_name, image_path)).wpf_xml()
            )
        
        self.children.insert(0, resources_node )
        
    
    def wpf_xml(self):
        node = super(BaseScrollViewer, self).wpf_xml()
        # xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        # Add other xml prefixes
        for prefix, url in WpfXMLFactory.namespace_prefixes.items():
            node.Add(linq.XAttribute(linq.XNamespace.Xmlns + prefix, url))
        
        return node




class ExpanderStackPanel(Wpf.Expander):
    '''
    Simplified definition of StackPanel within Expander
    '''
    def __init__(self, *args, **userkwargs):
        super(ExpanderStackPanel, self).__init__()
        
        kwargs = dict(Orientation="Vertical")
        kwargs.update(userkwargs)
        self.children = [
            Wpf.StackPanel(
                *args, **kwargs
            )
        ]


class ExpanderWrapPanel(Wpf.Expander):
    '''
    Simplified definition of WrapPanel within Expander
    '''
    def __init__(self, *args, **userkwargs):
        super(ExpanderWrapPanel, self).__init__()
        kwargs = dict(Orientation="Horizontal")
        kwargs.update(userkwargs)
        self.children = [
            Wpf.WrapPanel(
                *args, **kwargs
            )
        ]

class Expander(Wpf.Expander):
    '''
    Simplified definition of WrapPanel within Expander
    '''
    def __init__(self, *args, **userkwargs):
        
        self.auto_stack  = userkwargs.pop('auto_stack', False)
        self.auto_wrap   = userkwargs.pop('auto_wrap', False)
        
        if self.auto_stack:
            super(Expander, self).__init__(Header=userkwargs.pop('header', None), IsExpanded=userkwargs.pop('IsExpanded', False))
            kwargs = dict(Orientation="Vertical")
            kwargs.update(userkwargs)
            self.children = [Wpf.StackPanel(*args, **kwargs)]
        elif self.auto_wrap:
            super(Expander, self).__init__(Header=userkwargs.pop('header', None), IsExpanded=userkwargs.pop('IsExpanded', False))
            kwargs = dict(Orientation="Horizontal")
            kwargs.update(userkwargs)
            self.children = [Wpf.WrapPanel(*args, **kwargs)]
        else:
            super(Expander, self).__init__(*args, **userkwargs)
        

        
class GroupSeparator(TaskPaneControl):
    _xml_name="Grid"
    _attributes = {
        "margin":"0,10,0,5"
    }
    
    def __init__(self, *args, **kwargs):
        self.label = kwargs.pop('label', None) or kwargs.pop('Label', None)
        super(GroupSeparator, self).__init__(*args, **kwargs)
    
    def wpf_xml(self):
        if self.label:
            self.children=[
                XmlPart(
                    """<Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>"""),
                XmlPart('<Border Height="1" Background="{StaticResource BKTDivider}" HorizontalAlignment="Stretch" SnapsToDevicePixels="True" Margin="7,3,0,3" />'),
                Wpf.Label(attributes={"Grid.Column":"1"}, Padding="6,1,0,0", Foreground="{StaticResource GroupLabel}", Content=self.label),
                XmlPart('<Border Grid.Column="2" Height="1" Background="{StaticResource BKTDivider}" HorizontalAlignment="Stretch" SnapsToDevicePixels="True" Margin="10,3,10,3" />'),
            ]
        else:
            self.children=[
                XmlPart(
                    """<Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>"""),
                XmlPart('<Border Height="1" Background="{StaticResource BKTDivider}" HorizontalAlignment="Stretch" SnapsToDevicePixels="True" Margin="7,3,10,3" />'),
            ]
        
        return super(GroupSeparator,self).wpf_xml()


class Group(TaskPaneControl):
    _xml_name= "StackPanel"
    _attributes = {
        "orientation":"Vertical",
    }
    
    def __init__(self, *args, **userkwargs):
        self.auto_wrap = userkwargs.pop('auto_wrap', False)
        self.label     = userkwargs.pop('label', None) or userkwargs.pop('Label', None)
        self.show_separator = userkwargs.pop('show_separator', True)
        
        if self.auto_wrap:
            super(Group, self).__init__()
            kwargs = dict(Orientation="Horizontal")
            kwargs.update(userkwargs)
            self.children = [Wpf.WrapPanel(*args, **kwargs)]
        else:
            super(Group, self).__init__(*args, **userkwargs)
    
    def wpf_xml(self):
        stackpanel = super(Group,self).wpf_xml()
        if self.show_separator:
            stackpanel.AddFirst(GroupSeparator(label=self.label).wpf_xml())
        return stackpanel


class Button(FluentRibbon.Button):
    _attributes = {
        "size":"Middle"
    }







# =======================================
# = dictionary and key/value conversion =
# =======================================


def convert_value_to_string(v):
    if v == True:
        return 'true'
    elif v == False:
        return 'false'
    elif isinstance(v, (str, unicode)):
        return v
    else:
        try:
            return v.xml()
        except:
            return str(v)

def convert_key_to_upper_camelcase(key):
    parts = key.split('_')
    parts_new = []
    for i, part in enumerate(parts):
        if len(part) > 1:
            p = part[0].upper() + part[1:]
        else:
            p = part.upper()
            
        parts_new.append(p)
    return ''.join(parts_new)

def convert_dict_to_ribbon_xml_style(d):
    return {convert_key_to_upper_camelcase(k):convert_value_to_string(v) for k, v in d.items() if v != None and not isinstance(v, XamlPropertyElement) and not isinstance(v, TaskPaneControl)}



