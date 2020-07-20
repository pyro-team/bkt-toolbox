# -*- coding: utf-8 -*-
'''
Created on 14.07.2020

@author: fstallmann
'''

class Mock(object):
    def __init__(self, *args, **kwargs):
        pass
    
    def __getitem__(self, key):
        return self
    
    def __setitem__(self, key, value):
        pass
    
    def __call__(self, *args, **kwargs):
        return self
    
    def __getattr__(self, attr):
        return self
    
    def __iter__(self):
        return iter([])

    def __iadd__(self, other):
        return self
    
    def __repr__(self):
        return "<Mock class=%s id=%s>" % (self.__class__.__name__, id(self))

class OfficeMock(Mock):
    unknown_raises_error = True

    def __init__(self, *args, **kwargs):
        self._attributes = dict()
    
    def __getattr__(self, name):
        try:
            return self._attributes[name.lower()]
        except KeyError:
            raise AttributeError(name)
    
    def __setattr__(self, name, value):
        if name.startswith("_"):
            super(OfficeMock, self).__setattr__(name, value)
        elif name.lower() in self._attributes or not self.unknown_raises_error:
            self._attributes[name.lower()] = value
        else:
            raise AttributeError(name)