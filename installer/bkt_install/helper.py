# -*- coding: utf-8 -*-
'''
Created on 25.02.2019

@author: fstallmann
'''

from __future__ import absolute_import, print_function

import os
import ConfigParser
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
    reply = str(raw_input(question + ' (y/n): ')).lower().strip()
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
    import StringIO
    import traceback

    fd = StringIO.StringIO()
    traceback.print_exc(file=fd)
    traceback.print_exc()


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

configs = {}
def get_config(config_filename):
    try:
        return configs[config_filename]
    except KeyError:
        configs[config_filename] = config = BKTConfigParser(config_filename)
        if os.path.exists(config_filename):
            config.read(config_filename)
        else:
            config.add_section('BKT')
        return config