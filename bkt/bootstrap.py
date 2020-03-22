# -*- coding: utf-8 -*-
'''
Bootstrapper for the BKT python addin

Created on 03.09.2013
@author: cschmitt
'''

from __future__ import absolute_import

# import helpers --> ~2 sec
#from bkt.helpers import exception_as_message

def create_addin():
    try: 
        import bkt.addin as addin
        return addin.AddIn()
    except:
        import bkt.helpers
        bkt.helpers.exception_as_message()
