# -*- coding: utf-8 -*-



from ctypes import windll, Structure, c_long, byref

# _User32 = None
# _GetKeyState = None

def get_key_state(code):
    '''
    Get the current state of a specified key using windll
    https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getkeystate
    '''
    # global' _User32, _GetKeyState
    # if _User32 == None:
    #     _User32 = ctypes.CDLL("User32.dll")
    # if _GetKeyState == None:
    #     _GetKeyState = _User32.__getattr__("GetKeyState")
    
    # return' (_GetKeyState(code) & 128) == 128

    return (windll.user32.GetKeyState(code) & 128) == 128


class KeyCodes(object):
    '''
    Prefefined key codes commonly used in bkt
    More Key Codes: http://msdn.microsoft.com/en-us/library/dd375731(v=vs.85).aspx
    '''
    SHIFT = 0x10
    CTRL  = 0x11
    ALT   = 0x12


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


class POINT(Structure):
    _fields_ = [("x", c_long), ("y", c_long)]

def get_mouse_position():
    pt = POINT()
    windll.user32.GetCursorPos(byref(pt))
    return { "x": pt.x, "y": pt.y}


class MessageBox(object):
    '''
    Definition of standard windows message box
    https://github.com/asweigart/PyMsgBox/blob/master/src/pymsgbox/_native_win.py
    https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-messageboxw
    Can be used by calling instance (mesage("test")) or by calling specific messages boxes (message.confirmation("test"))
    '''
    MB_OK = 0x0
    MB_OKCANCEL = 0x1
    MB_ABORTRETRYIGNORE = 0x2
    MB_YESNOCANCEL = 0x3
    MB_YESNO = 0x4
    MB_RETRYCANCEL = 0x5
    MB_CANCELTRYCONTINUE = 0x6

    NO_ICON = 0
    STOP = MB_ICONHAND = MB_ICONSTOP = MB_ICONERRPR = 0x10
    QUESTION = MB_ICONQUESTION = 0x20
    WARNING = MB_ICONEXCLAIMATION = 0x30
    INFO = MB_ICONASTERISK = MB_ICONINFOMRAITON = 0x40

    MB_DEFAULTBUTTON1 = 0x0
    MB_DEFAULTBUTTON2 = 0x100
    MB_DEFAULTBUTTON3 = 0x200
    MB_DEFAULTBUTTON4 = 0x300

    MB_APPLMODAL = 0
    MB_SYSTEMMODAL = 0x1000
    MB_TASKMODAL = 0x2000

    MB_SETFOREGROUND = 0x10000
    MB_TOPMOST = 0x40000

    IDABORT = 0x3
    IDCANCEL = 0x2
    IDCONTINUE = 0x11
    IDIGNORE = 0x5
    IDNO = 0x7
    IDOK = 0x1
    IDRETRY = 0x4
    IDTRYAGAIN = 0x10
    IDYES = 0x6

    @staticmethod
    def _get_hwnd():
        try:
            return windll.user32.GetForegroundWindow()
        except:
            return 0

    @staticmethod
    def _show_message_box(*args):
        return windll.user32.MessageBoxW(*args)
    

    # Easy access to standard message box types
    def __call__(self, text, title="BKT", icon=INFO, buttons=MB_OK):
        return MessageBox._show_message_box(MessageBox._get_hwnd(), text, title, buttons | icon | MessageBox.MB_TASKMODAL | MessageBox.MB_SETFOREGROUND)
    
    def confirmation(self, text, title="BKT", buttons=MB_OKCANCEL, icon=QUESTION):
        result = self(text, title, buttons, icon)
        if buttons in (MessageBox.MB_OKCANCEL, MessageBox.MB_YESNO):
            return result in (MessageBox.IDOK, MessageBox.IDYES)
        else:
            return result

    def warning(self, text, title="BKT"):
        return self(text, title, icon=MessageBox.WARNING)

    def error(self, text, title="BKT"):
        return self(text, title, icon=MessageBox.STOP)

message = MessageBox()