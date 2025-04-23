# -*- coding: utf-8 -*-
'''
XML factory

Created on 23.11.2014
@author: cschmitt
'''



import logging

from bkt import dotnet
linq = dotnet.import_linq()


class RibbonXMLFactory(object):
    ''' convenience class to save some space in ribbon XML conversions'''
    
    namespace = "http://schemas.microsoft.com/office/2009/07/customui"
    
    namespace_prefixes = {}
    
    def __init__(self, namespace=None):
        self.namespace = namespace or type(self).namespace
        self.ns_office = linq.XNamespace.Get(self.namespace)
        
    def node(self, node_tag, **kwargs):
        element = linq.XElement(self.ns_office + node_tag)
        for k in sorted(kwargs):
            v = kwargs[k]
            
            key_parts = k.split("_")
            if len(key_parts) == 2:
                ns, key = key_parts
                if ns in self.namespace_prefixes:
                    attr = linq.XAttribute(linq.XNamespace.Get(self.namespace_prefixes[ns]) + key, v)
                else:
                    attr = linq.XAttribute(k, v)
            else:
                attr = linq.XAttribute(k, v)
            element.Add(attr)
            
        return element

    def pnode(self, parent, node_tag, **kwargs):
        node = self.node(node_tag, **kwargs)
        parent.Add(node)
        return node
    
    def attr(self, key, value):
        return linq.XAttribute(key, value)
    
    def nattr(self, node, key, value):
        # attr = self.attr(key, value)
        # node.Add(attr)
        # return attr
        node.SetAttributeValue(key, value)
        return None
    
    @staticmethod
    def to_normalized_string(x):
        text = "<" + x.Name.ToString()
        if x.HasAttributes:
            attrs = {attr.Name.ToString():attr.Value for attr in x.Attributes() }
            keys = sorted(attrs.keys())
            for k in keys:
                text += " " + k + "=\"" + attrs[k] + "\""
        
        if x.HasElements:
            text += ">\n"
            for child in x.Elements():
                text += RibbonXMLFactory.to_normalized_string(child)
            text += "</" + x.Name.ToString() + ">\n"
        else:
            text += " />\n"
        
        return text
        
    @staticmethod
    def to_string(x):
        doc = linq.XDocument()
        doc.Add(x)
        return doc.ToString()


class WpfXMLFactory(RibbonXMLFactory):
    namespace = "http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    
    namespace_prefixes = {
        'x': 'http://schemas.microsoft.com/winfx/2006/xaml',
        'po': 'http://schemas.microsoft.com/winfx/2006/xaml/presentation/options',
        'sys': 'clr-namespace:System;assembly=mscorlib',
        'r': 'clr-namespace:System.Windows.Controls.Ribbon;assembly=System.Windows.Controls.Ribbon',
        'fr': 'urn:fluent-ribbon'
    }
    







