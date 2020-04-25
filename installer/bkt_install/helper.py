# -*- coding: utf-8 -*-
'''
Created on 25.02.2019

@author: fstallmann
'''

from __future__ import absolute_import, print_function

import os
import ConfigParser
import ctypes


def is_admin():
    #test for admin
    try:
        return os.getuid() == 0
    except AttributeError:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0


def yes_no_question(question):
    reply = str(raw_input(question + ' (y/n): ')).lower().strip()
    if reply[0] == 'y':
        return True
    else:
        return False


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


def get_config(config_filename):
    config = BKTConfigParser(config_filename)
    if os.path.exists(config_filename):
        config.read(config_filename)
    else:
        config.add_section('BKT')
    return config