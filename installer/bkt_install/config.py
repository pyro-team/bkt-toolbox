# -*- coding: utf-8 -*-
'''
Created on 25.02.2019

@author: fstallmann
'''

from __future__ import absolute_import, print_function

import os.path

from . import helper
from .globals import INSTALL_BASE


class Configurator(object):
    @staticmethod
    def get_config():
        config_filename = os.path.join(INSTALL_BASE, "config.txt")
        if not os.path.exists(config_filename):
            raise SystemError("config file not found")
        return helper.get_config(config_filename)

    @classmethod
    def update_configuration(cls, values):
        config = cls.get_config()
        for k,v in values:
            print("Setting {} = {}".format(k,v))
            config.set_smart(k, v)

    @classmethod
    def add_folders(cls, folders):
        config = cls.get_config()
        existing_folders = config.feature_folders or []
        for folder in folders:
            folder = folder if os.path.isabs(folder) else os.path.normpath(os.path.join(INSTALL_BASE, folder))
            if not os.path.exists(folder):
                print("Folder does not exist: " + folder)
            elif folder in existing_folders:
                print("Folder already exists: " + folder)
            else:
                existing_folders.append( folder )
                print("Adding feature folder: " + folder)
        config.set_smart("feature_folders", existing_folders)

    @classmethod
    def remove_folders(cls, folders):
        config = cls.get_config()
        existing_folders = config.feature_folders or []
        for folder in folders:
            folder = folder if os.path.isabs(folder) else os.path.normpath(os.path.join(INSTALL_BASE, folder))
            if folder in existing_folders:
                existing_folders.remove( folder )
                print("Removing feature folder: " + folder)
            else:
                print("Folder not in config: " + folder)
        config.set_smart("feature_folders", existing_folders)
    
    @classmethod
    def migrate(cls, from_version):
        config = cls.get_config()
        if from_version == "2.4":
            folders = config.feature_folders or []
            if len(folders) > 0:
                print("Updating folder locations")
                old_path = os.path.normpath(os.path.join(INSTALL_BASE, '..', 'features'))
                new_path = os.path.normpath(os.path.join(INSTALL_BASE, 'features'))
                new_folders = [folder.replace(old_path+"\\", new_path+"\\") for folder in folders] #replace old path with new path
                new_folders = [folder for folder in new_folders if os.path.exists(folder)] #remove non-existent folders
                print("All folders updated")

                config.set_smart("feature_folders", new_folders)
            else:
                print("No folders to update")


def configure(args):
    if args.set_config:
        print("Updating config...")
        Configurator.update_configuration(args.set_config)

    if args.add_folders:
        print("Adding folders...")
        Configurator.add_folders(args.add_folders)

    if args.remove_folders:
        print("Removing folders...")
        Configurator.remove_folders(args.remove_folders)

    if args.migrate_from:
        print("Start migration...")
        Configurator.migrate(args.migrate_from)