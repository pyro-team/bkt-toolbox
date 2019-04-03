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

import ctypes #required for messagebox

import ConfigParser #required for config.txt file
import shelve #required for global settings database


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
        import bkt.console
        bkt.console.show_message(s)

def exception_as_message(additional_message=None):
    from cStringIO import StringIO
    import traceback

    import bkt.console
    import bkt.ui

    fd = StringIO()
    if additional_message:
        print(additional_message,file=fd)
    traceback.print_exc(file=fd)
    traceback.print_exc()

    bkt.console.show_message(bkt.ui.endings_to_windows(fd.getvalue()))

#@deprecated #os.path.join can handle multiple arguments
def mjoin(*paths):
    ''' Joins multiple path components. Use it to avoid multiple calls of os.path.join() '''
    current = paths[0]
    for path in paths[1:]:
        current = os.path.join(current,path)
    return current



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


# load config
config = BKTConfigParser()
config_filename=os.path.join(os.path.dirname(os.path.realpath(__file__)), "..", "config.txt")
if os.path.exists(config_filename):
    config.read(config_filename)
else:
    config.add_section('BKT')


def ensure_folders_exist(folder_path):
    if not os.path.isdir(folder_path):
        from os import makedirs
        makedirs(folder_path)
    return folder_path


def get_fav_folder():
    folder = config.local_fav_path or False
    if folder:
        return folder
    else:
        return ensure_folders_exist(os.path.join(os.path.expanduser("~"), "Documents", "BKT-Favoriten"))

def get_cache_folder():
    folder = config.local_cache_path or False
    if folder:
        return folder
    else:
        return ensure_folders_exist(os.path.normpath( os.path.join( os.path.dirname(__file__), "../resources/cache") ))

def get_settings_folder():
    folder = config.local_settings_path or False
    if folder:
        return folder
    else:
        return ensure_folders_exist(os.path.normpath( os.path.join( os.path.dirname(__file__), "../resources/settings") ))


#lazy loading shelve
class BKTSettings(shelve.Shelf):

    def __init__(self):
        shelve.Shelf.__init__(self, shelve._ClosedDict())
    
    def open(self, filename):
        import anydbm
        try:
            self.dict = anydbm.open(os.path.join( get_settings_folder(), filename), 'c')
        except:
            logging.error("error reading bkt settings")
            logging.debug(traceback.format_exc())
            exception_as_message()
            self.dict = dict() #fallback to empty dict

#load global setting database
settings = BKTSettings()

# try:
#     settings = shelve.open(os.path.join( get_settings_folder(), "bkt.settings" ))
# except:
#     exception_as_message()
#     settings = dict() #fallback to empty dict
