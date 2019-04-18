# -*- coding: utf-8 -*-

import sys
import os
import argparse

import helper

parser = argparse.ArgumentParser()
parser.add_argument('-af', '--add_folder', action='append', help='Add feature folder to config file')
parser.add_argument('-mg', '--migrate', help='Migrate config.txt from the given version to the current one')
args = parser.parse_args()


if args.add_folder is not None:
    root_path = os.path.normpath(os.path.join(os.path.dirname(__file__),'..'))
    folders = getattr(helper.config, "feature_folders") or []
    for folder in args.add_folder:
        folder = folder if os.path.isabs(folder) else os.path.join(root_path, folder)
        if not os.path.exists(folder):
            print("Folder does not exist: " + folder)
        elif folder in folders:
            print("Folder already exists: " + folder)
        else:
            folders.append( folder )
            print("Adding feature folder: " + folder)

    helper.config.set_smart("feature_folders", folders)


elif args.migrate is not None:
    if args.migrate == "2.4":
        print("Updating folder locations")
        old_path = os.path.normpath(os.path.join(os.path.dirname(__file__),'..', '..', 'features'))
        new_path = os.path.normpath(os.path.join(os.path.dirname(__file__),'..', 'features'))
        folders = getattr(helper.config, "feature_folders") or []
        new_folders = [folder.replace(old_path+"\\", new_path+"\\") for folder in folders] #replace old path with new path
        new_folders = [folder for folder in new_folders if os.path.exists(folder)] #remove non-existent folders
        print("All folders updated")

        helper.config.set_smart("feature_folders", new_folders)