# -*- coding: utf-8 -*-

from __future__ import absolute_import

import ctypes

_User32 = None
_GetKeyState = None

class key_code(object):
    # More Key Codes: http://msdn.microsoft.com/en-us/library/dd375731(v=vs.85).aspx
    SHIFT = 0x10
    CTRL  = 0x11
    ALT   = 0x12


def get_key_state(code):
    global _User32, _GetKeyState
    if _User32 == None:
        _User32 = ctypes.CDLL("User32.dll")
    if _GetKeyState == None:
        _GetKeyState = _User32.__getattr__("GetKeyState")
    
    return (_GetKeyState(code) & 128) == 128
    


