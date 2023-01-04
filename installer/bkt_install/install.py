# -*- coding: utf-8 -*-
'''
Created on 12.09.2013
Refactored on 25.02.2017

@author: cschmitt
'''



import os
from collections import OrderedDict

from . import reg
from . import helper
from .bitness import BitnessChecker

from .globals import INSTALL_BASE, default_config


class AppInfo(object):
    load_behavior = {
        'bkt' : 2,
        'bkt_dev': 0
        }
    register_addins = {'bkt'}


class PowerPoint(AppInfo):
    name = 'PowerPoint'
    addins_regpath = reg.office_default_path('PowerPoint')
    register_addins = {'bkt', 'bkt_dev'}
    # load_behavior = {
    #     'bkt' : 3,
    #     'bkt_dev': 3,
    #     }


class Word(AppInfo):
    name = 'Word'
    addins_regpath = reg.office_default_path('Word')


class Excel(AppInfo):
    name = 'Excel'
    addins_regpath = reg.office_default_path('Excel')


class Outlook(AppInfo):
    name = 'Outlook'
    addins_regpath = reg.office_default_path('Outlook')


class Visio(AppInfo):
    name = 'Visio'
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
        for app in self.apps:
            if self.uninstall:
                addins = list(self.addins)
            else:
                addins = app.register_addins
            
            for addin_key in addins:
                addin = self.addins[addin_key]
                yield self.get_application_addin_info(app, addin)
    
    def iter_active_application_addin_infos(self):
        for app in self.apps:
            if self.uninstall:
                addins = list(self.addins)
            else:
                addins = app.register_addins
            
            for addin_key in addins:
                addin = self.addins[addin_key]
                if app.load_behavior.get(addin.key, 0) == 3:
                    yield dict(
                        prog_id=addin.prog_id,
                        app_name=app.name
                    )


def fmt_load_behavior(integer):
    return ('%08x' % integer).upper()


class Installer(object):
    def __init__(self, config=dict(), install_base=None, wow6432=None, dndlist=False):
        if install_base is None:
            install_base = INSTALL_BASE

        self.install_base = install_base
        self.user_config = config
        
        if wow6432 is None:
            helper.log('checking system and office for 32/64 bit')
            wow6432 = BitnessChecker.get_bitness()
            helper.log('office is running in %s' % ("32-bit" if wow6432 else "64-bit"))
        
        self.wow6432 = wow6432
        self.dndlist = dndlist
    
    def create_config_file(self):
        ''' creates the config file with the entries necessary to bootstrap the IronPython environment '''
        ipy_addin_path = self.install_base
        
        # fixed config values, will always be written into config
        install_config = OrderedDict()
        install_config["ironpython_root"]  = os.path.join(self.install_base, 'bin')
        install_config["ipy_addin_path"]   = ipy_addin_path
        install_config["ipy_addin_module"] = "bkt.bootstrap"
        
        # default config values, existing values in config will not be overwritten
        default_config = install_config.copy() #use copy of install_config to preserve order
        default_config["log_write_file"]    = False
        default_config["log_level"]         = 'WARNING'
        default_config["log_show_msgbox"]   = False
        default_config["show_exception"]    = False
        default_config["async_startup"]     = False
        default_config["modules"]           = [ 'modules.settings' ]
        default_config["feature_folders"]   = []
        
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
            elif isinstance(existing_value, list) and isinstance(value, list):
                new_value = existing_value + [v for v in value if not v in existing_value]
            else:
                new_value = existing_value
            config.set_smart(key, new_value, False)
        
        # write fixed config values
        for key,value in install_config.items():
            config.set_smart(key, value, False)
        
        # remove example section (incl. all contents) as it is re-created later
        config.remove_section("EXAMPLE")

        # save file
        config.save_to_disk()
        
        # create example in section with dummy value (section starting with comment leads to parsing error)
        config_example = """
[EXAMPLE]
dummy_option = dummy_value
	#
	######## CONFIG examples ########
	#
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
	#  use_keymouse_hooks = True
	#  enable_legacy_syntax = False
	#  updates_auto_check_frequency = fridays-only
	#  
	#  ### Optional path settings
	#  local_fav_path = <path>
	#  local_cache_path = <path>
	#  local_settings_path = <path>
	#  
	#  ### App-specific config
	#  ppt_use_contextdialogs = True
	#  ppt_hide_format_tab = False
	#  ppt_activate_tab_on_new_shape = False
	#  excel_ignore_warnings = False
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
	#
"""
        # append config example
        with open(config_filename, 'a', encoding='utf-8') as fd:
            fd.write(config_example)


    
    def unregister(self):
        reginfo = RegistryInfoService(uninstall=True, install_base=self.install_base)
        for info in reginfo.iter_application_addin_infos():
            reg.AddinRegService(**info).unregister_addin()

        for info in reginfo.iter_addin_assembly_infos():
            reg.AssemblyRegService(wow6432=self.wow6432, **info).unregister_assembly()

        for info in reginfo.iter_active_application_addin_infos():
            reg.ResiliencyRegService(**info).remove_from_dndlist()
            
    def install(self):
        self.unregister()
        # try:
        helper.log('create registry entries')
        self.register()
        if self.user_config is not None:
            helper.log('create/update config file')
            self.create_config_file()
        # except Exception as e:
        #     self.unregister()
        #     raise e #re-raise exception
    
    def register(self):
        reginfo = RegistryInfoService(install_base=self.install_base)
        helper.log('register assemblies in registry')
        for info in reginfo.iter_addin_assembly_infos():
            reg.AssemblyRegService(wow6432=self.wow6432, **info).register_assembly()

        helper.log('register office addin in registry')
        for info in reginfo.iter_application_addin_infos():
            reg.AddinRegService(**info).register_addin()
        
        if self.dndlist:
            helper.log('register office addin in resiliency do-not-disable list')
            for info in reginfo.iter_active_application_addin_infos():
                reg.ResiliencyRegService(**info).add_to_dndlist()


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

    print('Deactivate any previous installation...')
    try:
        #this is required to avoid BKT loading when doing wow6232 check during installation
        Installer(wow6432=True).unregister()
        Installer(wow6432=False).unregister()
    except:
        helper.exception_as_message()

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

    wow6432 = None
    if args.force_office_bitness:
        wow6432 = args.force_office_bitness in ('32', 'x86')
        helper.log("forced office bitness to %s" % ("32-bit" if wow6432 else "64-bit"))

    # start installation
    try:
        if args.register_only:
            installer = Installer(wow6432=wow6432, config=None, dndlist=args.add_to_dndlist)
        else:
            installer = Installer(wow6432=wow6432, config=default_config, dndlist=args.add_to_dndlist)
        installer.install()

        print("\nInstallation ready -- addin available after Office restart")
    except:
        helper.exception_as_message()

