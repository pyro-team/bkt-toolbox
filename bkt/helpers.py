# -*- coding: utf-8 -*-
# '''
# Created on 10.09.2013
# 
# @author: cschmitt
# '''
# 

from __future__ import print_function

import os.path
import logging

# import von clr/MessageBox --> ~0.05 sec

# import dotnet
# Forms = dotnet.import_forms()
# MessageBox = Forms.MessageBox
# MessageBoxButtons = Forms.MessageBoxButtons

import ctypes #required for messagebox

# importlib fuer lazy-load von anderen libs
# beschleunigt start / verlangsamt erste tatsaechliche Verwendung
import importlib

# import von ui --> ~1 sec
#import bkt.ui as _ui

log_as_messagebox = False
log_as_uibox = False

#class is compatible to systems.forms
class Forms(object):
    class MessageBoxButtons(object):
        OK =                     0x00000000L #OK
        OKCancel =               0x00000001L #OK | Cancel
        AbortRetryIgnore =       0x00000002L #Abort | Retry | Ignore
        YesNoCancel =            0x00000003L #Yes | No | Cancel
        YesNo =                  0x00000004L #Yes | No
        RetryCancel =            0x00000005L #Retry | Cancel 
        CancelTryAgainContinue = 0x00000006L #Cancel | Try Again | Continue

    class MessageBoxIcon(object):
        #None =        0x00000000L
        Stop =        0x00000010L
        Error =       0x00000010L
        Hand =        0x00000010L
        Question =    0x00000020L
        Exclamation = 0x00000030L
        Warning     = 0x00000030L
        Information = 0x00000040L
        Asterisk    = 0x00000040L

    class DialogResult(object):
        OK          = 1
        Yes         = 6
        No          = 7
        Cancel      = 2
        Abort       = 3
        Continue    = 11
        Ignore      = 5
        Retry       = 4
        TryAgain    = 10

    class MessageBox(object):
        @staticmethod
        def Show(text, title, buttons, icon):
            def _get_hwnd():
                try:
                    return ctypes.windll.user32.GetForegroundWindow()
                except:
                    return 0
            
            return ctypes.windll.user32.MessageBoxW(_get_hwnd(), text, title, buttons | icon | 0x00002000L | 0x00010000L) #TASKMODAL | SETFOREGROUND


def message(text, title="BKT"):
    #MessageBox.Show(text, title, buttons, icon, default button, options, help-string)
    Forms.MessageBox.Show(text, title, Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Information)

def confirmation(text, title="BKT", buttons=Forms.MessageBoxButtons.OKCancel):
    result = Forms.MessageBox.Show(text, title, buttons, Forms.MessageBoxIcon.Question)
    if buttons == Forms.MessageBoxButtons.OKCancel or buttons == Forms.MessageBoxButtons.YesNo:
        if result == Forms.DialogResult.OK or result == Forms.DialogResult.Yes:
            return True
        else:
            return False
    else:
        return result

def log(s):
    logging.warning(s)
    #print(s)
    if log_as_messagebox:
        message(s)
    elif log_as_uibox:
        _co = importlib.import_module('bkt.console')
        _co.show_message(s)

def exception_as_message(additional_message=None):
    import StringIO
    import traceback

    fd = StringIO.StringIO()
    if additional_message:
        print(additional_message,file=fd)
    traceback.print_exc(file=fd)
    traceback.print_exc()

    _co = importlib.import_module('bkt.console')
    _ui = importlib.import_module('bkt.ui')
    _co.show_message(_ui.endings_to_windows(fd.getvalue()))


def mjoin(*paths):
    ''' Joins multiple path components. Use it to avoid multiple calls of os.path.join() '''
    current = paths[0]
    for path in paths[1:]:
        current = os.path.join(current,path)
    return current


import ConfigParser


class BKTConfigParser(ConfigParser.ConfigParser):

    def __getattr__(self, attr):
        '''
        returns self.get("BKT", attr)
        Method is injected into ConfigParser-class as fallback __getattr__ to allow
        access to config-options through attribute notation, e.g. config.my_option
        Multiline options (starting with \n) are split into lists.
        '''
        try:
            value = self.get("BKT", attr)
        except Exception:
            return None
        if value == "":
            return value
        elif value in ['false', 'False']:
            return False
        elif value in ['true', 'True']:
            return True
        elif value[0] != "\n":
            return value
        else:
            return value[1:].split("\n")

    def get_smart(self, attr, default=None, attr_type=str):
        try:
            if attr_type==bool:
                return self.getboolean("BKT", attr)
            elif attr_type==int:
                return self.getint("BKT", attr)
            elif attr_type==float:
                return self.getfloat("BKT", attr)
            else:
                return attr_type(self.get("BKT", attr))
        except:
            return default

    def set_smart(self, option, value):
        '''
        Method is injected into ConfigParser-class.
        Sets the config-value for option in section 'BKT', converts lists-values
        to '\n'-seperated strings. List-values can be read from the config file
        using attribute notation (e.g. config.my_list_option).
        '''
        if type(value) == list:
            value_list = [str(v) for v in value]
            self.set('BKT', option, "\n" + "\n".join(value_list))
        else:
            self.set('BKT', option, str(value)) #always transform to string, otherwise cannot access the value in same session anymore

        # write config file
        #outfilename = os.path.join(os.path.dirname(os.path.realpath(__file__)), "..", "config_written.txt")
        with open(config_filename, "wb") as configfile:
            config.write(configfile)



config = BKTConfigParser()
config_filename=os.path.join(os.path.dirname(os.path.realpath(__file__)), "..", "config.txt")
if os.path.exists(config_filename):
    config.read(config_filename)
else:
    config.add_section('BKT')



def get_fav_folder():
    folder = config.local_fav_path or False
    if folder:
        return folder
    else:
        return os.path.join(os.path.expanduser("~"), "Documents", "BKT-Favoriten")