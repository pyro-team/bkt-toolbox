# -*- coding: utf-8 -*-
'''
This module provides an import hook (for sys.meta_path) to mockup the IronPython clr module
and .NET references specified via clr.AddReference.


Created on 20.11.2014

@author: cschmitt
'''

from __future__ import absolute_import, print_function
import sys
import types

class GenericMockModule(types.ModuleType):
    def __getitem__(self, key):
        return self
    
    def __call__(self, *args, **kwargs):
        return self
    
    def __getattr__(self, attr):
        return self
    
    def __iter__(self):
        return iter([])

    def __iadd__(self, other):
        return self
    
class MockCLR(object):
    def __init__(self):
        self.references = set()
        
    def AddReference(self, ref):
        parts = ref.split('.')
        for l in range(len(parts)):
            sub_ref = '.'.join(parts[:l+1])
            self.references.add(sub_ref)
        print(sorted(self.references))
        
class MockCLRMetaPath(object):
    def __init__(self):
        self.clr = MockCLR()
        
    def has_mod(self, name):
        for ref in self.clr.references:
            return ref in self.clr.references
        
    def load_module(self, full_name):
        if full_name == 'clr':
            return self.create_module('clr',
                                      dict(AddReference=self.clr.AddReference))
        elif full_name == 'System.Xml.Linq':
            from . import linq
            sys.modules[full_name] = linq
            return linq
        elif full_name in self.clr.references:
            return self.create_module(full_name)
        raise ImportError
        
    def find_module(self, name, foo):
        if name == 'clr' or name in self.clr.references:
            #print(name, foo)
            return self
    
    def create_module(self, name, d=None):
        if d is None:
            d = {}
        mod = GenericMockModule(name)
        mod.__path__ = None
        for k, v in d.items():
            setattr(mod, k, v)
        sys.modules[name] = mod
        return mod

def inject_mock():
    sys.meta_path.append(MockCLRMetaPath())
