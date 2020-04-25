# -*- coding: utf-8 -*-
'''
Various helper function, global config and settings parser

Created on 10.09.2013
@author: cschmitt
'''


from __future__ import absolute_import, print_function

import os.path
import logging

import ctypes #required for messagebox

import ConfigParser #required for config.txt file
import shelve #required for global settings database


BKT_BASE = os.path.realpath(os.path.join(os.path.dirname(__file__), ".."))

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
        Stop =        0x00000010L #=Error=Hand
        Error =       0x00000010L
        Hand =        0x00000010L
        Question =    0x00000020L
        Exclamation = 0x00000030L #=Warning
        Warning     = 0x00000030L
        Information = 0x00000040L #=Asterisk
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


def message(text, title="BKT", icon=Forms.MessageBoxIcon.Information):
    #MessageBox.Show(text, title, buttons, icon, default button, options, help-string)
    Forms.MessageBox.Show(text, title, Forms.MessageBoxButtons.OK, icon)

def confirmation(text, title="BKT", buttons=Forms.MessageBoxButtons.OKCancel, icon=Forms.MessageBoxIcon.Question):
    result = Forms.MessageBox.Show(text, title, buttons, icon)
    if buttons == Forms.MessageBoxButtons.OKCancel or buttons == Forms.MessageBoxButtons.YesNo:
        if result == Forms.DialogResult.OK or result == Forms.DialogResult.Yes:
            return True
        else:
            return False
    else:
        return result

def warning(text, title="BKT"):
    message(text, title, icon=Forms.MessageBoxIcon.Exclamation)

def error(text, title="BKT"):
    message(text, title, icon=Forms.MessageBoxIcon.Error)

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


def file_base_path_join(base_file, *args):
    return os.path.realpath(os.path.join(os.path.dirname(base_file), *args))

def bkt_base_path_join(*args):
    return os.path.realpath(os.path.join(BKT_BASE, *args))


class BKTConfigParser(ConfigParser.ConfigParser):
    config_filename = None

    def __init__(self, config_filename):
        self.config_filename = config_filename
        ConfigParser.ConfigParser.__init__(self)

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
        with open(self.config_filename, "wb") as configfile:
            self.write(configfile)


# load config
config_filename=bkt_base_path_join("config.txt")
config = BKTConfigParser(config_filename)
if os.path.exists(config_filename):
    config.read(config_filename)
else:
    config.add_section('BKT')


def ensure_folders_exist(folder_path):
    if not os.path.isdir(folder_path):
        from os import makedirs
        makedirs(folder_path)
    return folder_path


def get_fav_folder(*args):
    folder = config.local_fav_path or False
    if not folder:
        #FIXME: this doesnt work if Documents folder has been moved by user or by OneDrive installation
        folder = ensure_folders_exist( os.path.realpath(os.path.join(os.path.expanduser("~"), "Documents", "BKT-Favoriten")) )
    args = args or tuple()
    args = (folder,)+args
    return os.path.join(*args)

def get_cache_folder(*args):
    folder = config.local_cache_path or False
    if not folder:
        folder = ensure_folders_exist( bkt_base_path_join("resources","cache") )
    args = args or tuple()
    args = (folder,)+args
    return os.path.join(*args)

def get_settings_folder(*args):
    folder = config.local_settings_path or False
    if not folder:
        folder = ensure_folders_exist( bkt_base_path_join("resources","settings") )
    args = args or tuple()
    args = (folder,)+args
    return os.path.join(*args)


#lazy loading shelve
class BKTSettings(shelve.Shelf):

    def __init__(self):
        shelve.Shelf.__init__(self, shelve._ClosedDict(), protocol=2)
    
    def open(self, filename):
        import anydbm
        try:
            self.dict = anydbm.open(get_settings_folder(filename), 'c')
        except:
            logging.error("error reading bkt settings")
            # logging.debug(traceback.format_exc())
            exception_as_message()
            self.dict = dict() #fallback to empty dict
    
    def get(self, key, default=None):
        try:
            # super(BKTSettings, self).get(key, default) #doesnt work as Shelf is not a new-style object
            if key in self.dict:
                return self[key]
            return default
        except EOFError:
            logging.error("EOF-Error in settings for getting key {}. Reset to default value: {}".format(key, default))
            exception_as_message("Settings database corrupt for key {}. Trying to repair now.".format(key))

            #settings database corrupt, trying to fix it
            if default is None:
                del self[key]
            else:
                self[key] = default

            return default

#load global setting database
settings = BKTSettings()
