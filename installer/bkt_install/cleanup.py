# -*- coding: utf-8 -*-
'''
Created on 25.02.2019

@author: fstallmann
'''

from __future__ import absolute_import, print_function

import os

from . import helper
from .globals import INSTALL_BASE


class Cleaner(object):
    @classmethod
    def _purge_folder(cls, folder):
        if not os.path.isdir(folder):
            helper.log("%s not found" % folder)
            return
        
        if not helper.yes_no_question("Delete %s" % folder):
            return
        
        files = os.listdir(folder)
        for file in files:
            path = os.path.join(folder, file)
            if os.path.isfile(path):
                try:
                    os.remove(path)
                    helper.log("removed %s" % path)
                except:
                    helper.log("error removing %s" % path)
                    helper.exception_as_message()
    
    @classmethod
    def _get_from_config(cls, value, default=None):
        config_filename = os.path.join(INSTALL_BASE, 'config.txt')
        if os.path.exists(config_filename):
            return getattr(helper.get_config(config_filename), value) or default
        return default

    @classmethod
    def clear_cache(cls):
        cache_folder = cls._get_from_config("local_cache_path", os.path.join(INSTALL_BASE, 'resources', 'cache'))
        cls._purge_folder(cache_folder)

    @classmethod
    def clear_config(cls):
        config_filename = os.path.join(INSTALL_BASE, 'config.txt')
        if os.path.exists(config_filename):
            if helper.yes_no_question("Delete %s" % config_filename):
                os.remove(config_filename)
                helper.log("config.txt successfully removed")
                print("\nIMPORTANT: You need to run install command in order to generate new config.txt file!")
        else:
            helper.log("config.txt not found")

    @classmethod
    def clear_settings(cls):
        settings_folder = cls._get_from_config("local_cache_path", os.path.join(INSTALL_BASE, 'resources', 'settings'))
        cls._purge_folder(settings_folder)

    @classmethod
    def clear_xml(cls):
        xml_folder = os.path.join(INSTALL_BASE, 'resources', 'xml')
        cls._purge_folder(xml_folder)


def clean(args):
    if args.clear_cache:
        print("\nClearing cache...")
        Cleaner.clear_cache()

    if args.clear_config:
        print("\nClearing config.txt file...")
        Cleaner.clear_config()

    if args.clear_settings:
        print("\nClearing settings...")
        Cleaner.clear_settings()

    if args.clear_xml:
        print("\nClearing XML files...")
        Cleaner.clear_xml()