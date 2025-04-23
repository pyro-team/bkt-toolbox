# -*- coding: utf-8 -*-
'''
Created on 25.02.2019

@author: fstallmann
'''



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
        try:
            config = cls.get_config()
        except SystemError:
            print("New installation, no migration required")
        
        from_tuple = tuple(int(x) for x in from_version.split("."))
        migration_funcs = [
            ((2,4), cls._mig_moved_feature_folders),
            ((2,6), cls._mig_remove_dev_module),
            ((2,6), cls._mig_new_fav_folder),
        ]

        for version, func in migration_funcs:
            if version >= from_tuple:
                try:
                    func(config)
                except:
                    helper.exception_as_message()
    
    @staticmethod
    def _mig_moved_feature_folders(config):
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
    
    @staticmethod
    def _mig_remove_dev_module(config):
        modules = config.modules or []
        try:
            modules.remove("modules.dev")
            print("Dev module deactivated")
            config.set_smart("modules", modules)
        except ValueError:
            print("No dev module activated")
    
    @staticmethod
    def _mig_new_fav_folder(config):
        folder = config.local_fav_path or False
        if folder:
            print("Local fav folder already defined, no need to migrate")
            return
        
        old_folder = os.path.realpath(os.path.join(os.path.expanduser("~"), "Documents", "BKT-Favoriten"))
        if not os.path.isdir(old_folder):
            print("Old fav folder does not exist, no need to migrate")
            return

        from System import Environment
        doc_folder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        new_folder = os.path.join(doc_folder, "BKT-Favoriten")

        if old_folder == new_folder:
            print("Old fav folder equals new folder, no need to migrate")
            return
        
        if os.path.isdir(new_folder):
            print("New folder already exists, migration not possible")
            return
        
        import shutil
        shutil.move(old_folder, doc_folder)

        print("Old fav folder moved to new location: %s" % new_folder)


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
        print("\nMigration finished...")