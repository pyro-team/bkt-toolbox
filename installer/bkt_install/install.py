# -*- coding: utf-8 -*-
'''
Created on 12.09.2013
Refactored on 25.02.2017

@author: cschmitt
'''

from __future__ import absolute_import, division, print_function

import clr
import os
import traceback

import System.Environment

from . import reg
from . import helper

from .globals import INSTALL_BASE, default_config


class AppInfo(object):
    load_behavior = {
        'bkt' : 2,
        'bkt_dev': 0
        }
    register_addins = {'bkt'}


class PowerPoint(AppInfo):
    addins_regpath = reg.office_default_path('PowerPoint')
    register_addins = {'bkt', 'bkt_dev'}
    # load_behavior = {
    #     'bkt' : 3,
    #     'bkt_dev': 3,
    #     }


class Word(AppInfo):
    addins_regpath = reg.office_default_path('Word')


class Excel(AppInfo):
    addins_regpath = reg.office_default_path('Excel')


class Outlook(AppInfo):
    addins_regpath = reg.office_default_path('Outlook')


class Visio(AppInfo):
    addins_regpath = reg.PathString('Software') / 'Microsoft' / 'Visio' / 'Addins'


APPS = [
    PowerPoint,
    Excel,
    Word,
    Outlook,
    Visio,
    ] 


class AddinInfo(object):
    pass


class BKT(AddinInfo):
    key = 'bkt'
    prog_id = 'BKT.AddIn'
    uuid = '{8EA4071E-7BD4-48DA-B96D-21AD02E1C238}'
    name = 'BKT'
    description = 'Business Kasper Toolbox'
    dll = 'BKT.dll'


class BKTTaskPane(AddinInfo):
    key = 'bkt_taskpane'
    prog_id = 'BKT.TaskPane'
    uuid = '{76FD3062-86C8-11E4-BE43-6336340000B1}'
    name = 'BKT Task Pane'
    description = 'Business Kasper Toolbox Task Pane'
    dll = 'BKT.dll'


class BKTDev(AddinInfo):
    key = 'bkt_dev'
    prog_id = 'BKT.Dev.DevAddIn'
    uuid = '{FC4DBFDD-A8A2-4675-A32D-A56337844DC4}'
    name = 'BKT Dev'
    description = 'BKT Development Addin'
    dll = 'BKT.Dev.dll'


ALL_ADDINS = [
    BKT,
    BKTTaskPane,
    BKTDev,
    ]


INSTALL_ADDINS = [
    BKT,
    BKTTaskPane,
    BKTDev,
    ]



class RegistryInfoService(object):
    def __init__(self, apps=None, addins=None, install_base=None, uninstall=False):
        if apps is None:
            apps = list(APPS)
        if addins is None:
            if uninstall:
                addins = list(ALL_ADDINS)
            else:
                addins = list(INSTALL_ADDINS)
        if install_base is None:
            install_base = INSTALL_BASE
            
        self.apps = apps
        self.addins = {a.key: a for a in addins}
        self.install_base = install_base
        self.uninstall = uninstall
            
    def get_addin_assembly_info(self, addin_info):
        return dict(
            prog_id=addin_info.prog_id,
            uuid=addin_info.uuid,
            assembly_path=os.path.join(self.install_base, 'bin', addin_info.dll),
            )
    
    def iter_addin_assembly_infos(self):
        for addin in self.addins.values():
            yield self.get_addin_assembly_info(addin)
                
    def get_application_addin_info(self, app, addin):
        return dict(prog_id=addin.prog_id,
                   friendly_name=addin.name,
                   description=addin.description,
                   addins_regpath=app.addins_regpath,
                   load_behavior=app.load_behavior.get(addin.key, 0),
                   )

    def iter_application_addin_infos(self):
        all_addins = list(self.addins)
        for app in self.apps:
            if self.uninstall:
                addins = all_addins
            else:
                addins = app.register_addins
            
            for addin_key in addins:
                addin = self.addins[addin_key]
                yield self.get_application_addin_info(app, addin)


def check_wow6432():
    ''' returns true if office-32-bit is running on 64 bit machine '''
    iop_base = 'Microsoft.Office.Interop.'        
    
    apps = ['PowerPoint',
            'Excel']
    
    os_64 = System.Environment.Is64BitOperatingSystem
    if os_64 == False:
        return False
    
    office_is_32 = set()
    for app_name in apps:
        iop_name = iop_base + app_name
        try:
            clr.AddReference(iop_name)
            module = None
            # FIXME: this is ugly, but __import__(iop_name) does not seem to work
            exec 'import ' + iop_name + ' as module'
            app = module.ApplicationClass()
            try:
                office_is_32.add(app.OperatingSystem.startswith('Windows (32-bit)'))
            finally:
                app.Quit()
        except:
            traceback.print_exc()
            
    assert len(office_is_32) > 0, 'failed to get bitness of all tested office applications'
    
    return os_64 and (True in office_is_32)

def fmt_load_behavior(integer):
    return ('%08x' % integer).upper()


class Installer(object):
    def __init__(self, config=dict(), install_base=None, wow6432=None):
        if install_base is None:
            install_base = INSTALL_BASE

        self.install_base = install_base
        self.user_config = config
        
        if wow6432 is None:
            print('checking system and office for 32/64 bit')
            wow6432 = check_wow6432()
        
        self.wow6432 = wow6432
    
    def create_config_file(self):
        ''' creates the config file with the entries necessary to bootstrap the IronPython environment '''
        ipy_addin_path = self.install_base
        
        # fixed config values, will always be written into config
        install_config = dict(
            # ironpython_root = os.path.join(self.install_base, 'bin', 'ipy-2.7.9'),
            ironpython_root = os.path.join(self.install_base, 'bin'),
            ipy_addin_path = ipy_addin_path,
            ipy_addin_module = "bkt.bootstrap",
        )
        
        # default config values, existing values in config will not be overwritten
        default_config = dict(
            log_show_msgbox = False,
            log_write_file = False,
            log_level = 'WARNING',
            async_startup = False,
            # task_panes = False,
            show_exception = False,
            
            modules = [ 'modules.dev', 'modules.settings' ],
            
            feature_folders = [],
        )
        
        # allow Installer to be called with other default config values
        default_config.update(self.user_config)
        config_filename = os.path.join(self.install_base, 'config.txt')
        config = helper.get_config(config_filename)
        
        # change config
        # write default config values
        for key,value in default_config.items():
            existing_value = getattr(config, key)
            if existing_value == None or existing_value == '':
                new_value = value
            elif type(existing_value) == list and type(value) == list:
                new_value = existing_value + [v for v in value if not v in existing_value]
            else:
                new_value = existing_value
            config.set_smart(key, new_value)
        
        # write fixed config values
        for key,value in install_config.items():
            config.set_smart(key, value)
        
        config_example = """
######## CONFIG examples ########

#  ### Iron Python
#  ironpython_root = <installation_folder>\\bkt-framework\\bin
#  ipy_addin_path = <installation_folder>\\bkt-framework
#  ipy_addin_module = bkt.bootstrap
#  
#  ### Debugging
#  pydev_debug = True
#  pydev_codebase = <eclipse-folder>\\plugins\\org.python.pydev_<pydev-version>\\pysrc'
#  
#  ### Addin Configuration
#  log_write_file = True
#  log_level = False
#  log_show_msgbox = False
#  show_exception = False
#  async_startup = False
#  task_panes = False
#  
#  ### Modules',
#  modules = 
#       modules.settings
#       modules.tutorial
#       modules.demo.demo_customui
#       modules.demo.demo_bkt
#       modules.demo.demo_image_mso
#  	    modules.demo.demo_task_pane
#  
#  ### Feature-Folders
#  feature_folders = 
#       <some_feature_folder>
#       <another_feature_folder>
#  
#  ### Toolbox settings
#  chart_library_folders = 
#       <some_folder>
#       <another_folder>
#  chart_libraries = 
#       <some_file>
#       <another_file>
#  shape_library_folders = 
#       <some_folder>
#       <another_folder>
#  shape_libraries = 
#       <some_file>
#       <another_file>

"""
        # append config example
        with open(config_filename, 'a') as fd:
            fd.write(config_example.encode('utf-8'))


    
    def unregister(self):
        reginfo = RegistryInfoService(uninstall=True, install_base=self.install_base)
        for info in reginfo.iter_application_addin_infos():
            reg.AddinRegService(**info).unregister_addin()
            
        for info in reginfo.iter_addin_assembly_infos():
            reg.AssemblyRegService(wow6432=self.wow6432, **info).unregister_assembly()
            
    def install(self):
        self.unregister()
        try:
            print('create registry entries for addin assemblies')
            self.register()
            print('create/update config file')
            self.create_config_file()
        except Exception as e:
            self.unregister()
            raise e #re-raise exception
    
    def register(self):
        reginfo = RegistryInfoService(install_base=self.install_base)
        for info in reginfo.iter_addin_assembly_infos():
            reg.AssemblyRegService(wow6432=self.wow6432, **info).register_assembly()

        for info in reginfo.iter_application_addin_infos():
            reg.AddinRegService(**info).register_addin()


def uninstall(args):
    print('Uninstalling BKT from current directory...')
    try:
        Installer(wow6432=True).unregister()
        Installer(wow6432=False).unregister()
        print("\nBKT successfully uninstalled")
    except:
        helper.exception_as_message()

    if args.remove_config:
        try:
            print("\nRemoving BKT config file...")
            config_filename = os.path.join(INSTALL_BASE, 'config.txt')
            if os.path.exists(config_filename):
                os.remove(config_filename)
        except:
            helper.exception_as_message()


def install(args):
    if helper.is_admin() and not helper.yes_no_question('Are you sure to run BKT installer as admin?'):
        print('BKT installation cancelled')
        return

    print('Installing BKT in current directory...')

    #app load behaviour
    if "powerpoint" in args.apps:
        PowerPoint.load_behavior = { 'bkt' : 3, 'bkt_dev': 3 }
    if "excel" in args.apps:
        Excel.load_behavior = { 'bkt' : 3, 'bkt_dev': 3 }
    if "word" in args.apps:
        Word.load_behavior = { 'bkt' : 3, 'bkt_dev': 3 }
    if "visio" in args.apps:
        Visio.load_behavior = { 'bkt' : 3, 'bkt_dev': 3 }
    if "outlook" in args.apps:
        Outlook.load_behavior = { 'bkt' : 3, 'bkt_dev': 3 }

    # start installation
    try:
        if args.register_only:
            installer = Installer()
        else:
            installer = Installer(config=default_config)
        installer.install()

        print("\nInstallation ready -- addin available after Office restart")
    except:
        helper.exception_as_message()

