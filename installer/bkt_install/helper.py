# -*- coding: utf-8 -*-
'''
Created on 25.02.2019

@author: fstallmann
'''



import os
import io
import configparser
from ctypes import windll, POINTER
from ctypes.wintypes import LPWSTR, DWORD, BOOL


def log(message):
    # pass
    print("\t> %s" % message)


def is_admin():
    #test for admin
    try:
        return os.getuid() == 0
    except AttributeError:
        return windll.shell32.IsUserAnAdmin() != 0


def yes_no_question(question):
    reply = str(input(question + ' (y/n): ')).lower().strip()
    if reply[0] == 'y':
        return True
    else:
        return False


def is_64bit_os():
    import platform
    #https://stackoverflow.com/questions/2208828/detect-64bit-os-windows-in-python
    return platform.machine().endswith('64')


_GetBinaryType = windll.kernel32.GetBinaryTypeW
_GetBinaryType.argtypes = (LPWSTR, POINTER(DWORD))
_GetBinaryType.restype = BOOL

def is_64bit_exe(path):
    #https://stackoverflow.com/questions/1345632/determine-if-an-executable-or-library-is-32-or-64-bits-on-windows
    res = DWORD()
    if not _GetBinaryType(path, res):
        raise SystemError("could not get binary type")
    return res == 6 #SCS_64BIT_BINARY


def exception_as_message():
    import io
    import traceback

    fd = io.StringIO()
    traceback.print_exc(file=fd)
    traceback.print_exc()


class BKTConfigParser(configparser.ConfigParser):
    config_filename = None

    def __init__(self, config_filename):
        self.config_filename = config_filename
        configparser.ConfigParser.__init__(self)

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
        elif value.lower() in ['false', 'no', 'off', '0']:
            return False
        elif value.lower() in ['true', 'yes', 'on', '1']:
            return True
        elif value[0] != "\n":
            return value
        else:
            return value[1:].split("\n")
    
    def save_to_disk(self):
        '''
        Save the config back to disk.
        '''
        with io.open(self.config_filename, "w", encoding='utf-8') as configfile:
            self.write(configfile)

    def get_smart(self, attr, default=None, attr_type=str):
        '''
        Method to get config-values and force a particular data type, return
        default value on error. This method does not work for lists.
        '''
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

    def set_smart(self, option, value, write_back=True):
        '''
        Method is injected into ConfigParser-class.
        Sets the config-value for option in section 'BKT', converts lists-values
        to '\n'-seperated strings. List-values can be read from the config file
        using attribute notation (e.g. config.my_list_option).
        '''
        if type(value) == list:
            if value:
                self.set('BKT', option, "\n" + "\n".join(str(v) for v in value))
            else:
                #empty list
                self.set('BKT', option, '')
        else:
            self.set('BKT', option, str(value)) #always transform to string, otherwise cannot access the value in same session anymore

        # write config file
        if write_back:
            self.save_to_disk()

    def read_unicode(self):
        try:
            with io.open(self.config_filename, encoding='utf-8') as configfile:
                self.readfp(configfile)
        except UnicodeDecodeError:
            #before v2.7.5 config file was saved as ANSI, so open as ANSI here and save as Unicode later
            with io.open(self.config_filename) as configfile:
                self.readfp(configfile)

configs = {}
def get_config(config_filename):
    try:
        return configs[config_filename]
    except KeyError:
        configs[config_filename] = config = BKTConfigParser(config_filename)
        if os.path.exists(config_filename):
            config.read_unicode()
            if not config.has_section('BKT'):
                config.add_section('BKT')
        else:
            config.add_section('BKT')
        return config