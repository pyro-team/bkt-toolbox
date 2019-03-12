# -*- coding: utf-8 -*-
'''
Minimal compatibility module to support XML generation if BKT is not run IronPython. 

Created on 20.11.2014

@author: cschmitt
'''

class XNamespace(object):
    Xmlns = "{http://www.w3.org/2000/xmlns/}"
    
    def __init__(self, name):
        self.name = name
    
    @classmethod
    def Get(cls, name):
        return cls(name)
    
    def __add__(self, other):
        return other
    

class XName(object):
    
    def __init__(self, element_name):
        self.element_name = element_name
    
    def ToString(self):
        return self.element_name
    
class XElement(object):
    def __init__(self, tag):
        self.tag = tag
        self.attrs = {}
        self.children = []
        
    def Add(self, obj):
        if isinstance(obj, XAttribute):
            self.attrs[obj.key] = obj.value
        else:
            self.children.append(obj)
    
    @property
    def Name(self):
        return XName(self.tag)
    
    @property
    def HasAttributes(self):
        return len(self.attrs) > 0
    
    @property
    def HasElements(self):
        return len(self.children) > 0
    
    def Elements(self):
        return self.children

    def Attributes(self):
        return [ XAttribute(k,v) for  k,v in self.attrs.items()]
    
    def SetAttributeValue(self, key, value):
        self.attrs[key] = value    
            
    def ToString(self, indent=''):
        s = indent + '<%s ' % self.tag
        for k, v in self.attrs.items():
            s += '%s="%s" ' % (k, v)
        s = s.rstrip()
        if self.children:
            s += '>\n'
            for c in self.children:
                s += c.ToString(indent=indent+'  ')
                s += '\n'
            s += indent + '</%s>' % self.tag
        else:
            s += ' />'
        return s

class XDocument(object):
    def __init__(self):
        self.children = []

    def Add(self, obj):
        self.children.append(obj)
        
    def ToString(self):
        s = ''
        for child in self.children:
            s += child.ToString()
            s += '\n'
        return s

class XAttribute(object):
    def __init__(self, key, value):
        self.key = key
        self.value = value
    
    @property
    def Value(self):
        return self.value
    
    @property
    def Name(self):
        return XName(self.key)