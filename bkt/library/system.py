# -*- coding: utf-8 -*-

from __future__ import absolute_import

import ctypes

# _User32 = None
# _GetKeyState = None

# def get_key_state(code):
#     global' _User32, _GetKeyState
#     if _User32 == None:
#         _User32 = ctypes.CDLL("User32.dll")
#     if _GetKeyState == None:
#         _GetKeyState = _User32.__getattr__("GetKeyState")
    
#     return' (_GetKeyState(code) & 128) == 128


class KeyCodes(object):
    # More Key Codes: http://msdn.microsoft.com/en-us/library/dd375731(v=vs.85).aspx
    SHIFT = 0x10
    CTRL  = 0x11
    ALT   = 0x12

class KeyState(object):
    def __call__(self, code):
        return (ctypes.windll.user32.GetKeyState(code) & 128) == 128

    def __getattr__(self, attr):
        return self(getattr(KeyCodes, attr))
    
get_key_state = KeyState()


def apply_delta_on_ALT_key(setter_method, getter_method, shapes, value, **kwargs):
    '''
    If the ALT key is pressed, the setter-method is called shifting all values by the same delta for every shape,
    i.e. setter_method([shape], old_value + delta, **kwargs) is called for every shape

    The delta-value is obtained by comparing getter_method([shapes[0]]) and value.
    For every shape, old_value is obtained using getter_method([shape]).
    
    If the ALT key is not pressed, setter_method(shapes, value, **kwargs) is called
    '''
    
    alt_state = get_key_state(KeyCodes.ALT)
    
    if not alt_state:
        for shape in shapes:
            setter_method(shape=shape, value=value, **kwargs)
        
    else:
        delta = value - getter_method(shape=shapes[0], **kwargs)
        for shape in shapes:
            old_value = getter_method(shape=shape, **kwargs)
            setter_method(shape=shape, value=old_value + delta, **kwargs)

    return None