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