# -*- coding: utf-8 -*-
'''
Created on 19.02.2017

@author: chschmitt
'''

from __future__ import absolute_import, division, print_function

import os.path

from contextlib import contextmanager

import System
import Microsoft.Win32 as Win32

from System.Reflection import Assembly, AssemblyName

RegistryHive = Win32.RegistryHive
RegistryView = Win32.RegistryView
RegistryKey = Win32.RegistryKey
RegistryValueKind = Win32.RegistryValueKind


class PathString(str):
    def __truediv__(self, other):
        if not other or ('\\' in other) or ('/' in other):
            raise ValueError
        return type(self)(self + '\\' + other)

    __div__ = __truediv__
    __floordiv__ = __truediv__

@contextmanager
def open_key(base, path, *args, **kwargs):
    try:
        value = base.OpenSubKey(path, *args, **kwargs)
        if value is None:
            raise KeyError(str(base) + '\\' + path)
        yield value
    finally:
        if value:
            value.Close()

@contextmanager
def open_or_create(base, path, *args, **kwargs):
    try:
        value = base.CreateSubKey(path, *args, **kwargs)
        if value is None:
            raise KeyError(str(base) + '\\' + path)
        yield value
    finally:
        if value:
            value.Close()


class Properties(object):
    pass

class AssemblyRegService(object):
    def __init__(self, prog_id=None, uuid=None, assembly_path=None, wow6432=True):
        self.prog_id = prog_id
        self.uuid = uuid
        self.assembly_path = assembly_path
        self.wow6432 = wow6432

    def load_assembly_attributes(self):
        assembly = Assembly.ReflectionOnlyLoadFrom(self.assembly_path)
        assembly_name = AssemblyName(assembly.FullName)
        assembly_uri = u'file:///' + self.assembly_path.replace(os.path.sep, u'/')

        p = Properties()
        p.full_name = assembly.FullName
        p.version = str(assembly_name.Version)
        p.codebase_uri = assembly_uri
        p.runtime_version = assembly.ImageRuntimeVersion
        self.assembly_properties = p

    def get_hkcu(self, view=RegistryView.Default):
        return RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view)

    def get_hkcu_wow(self):
        if System.Environment.Is64BitOperatingSystem:
            if self.wow6432:
                view = RegistryView.Registry32
            else:
                view = RegistryView.Registry64
        else:
            view = RegistryView.Default
        return self.get_hkcu(view)

    def _define_prog_id(self, base, prog_id, uuid):
        prog_id_path = PathString('Software') / 'Classes' / prog_id
        with open_or_create(base, prog_id_path) as prog_id_key:
            prog_id_key.SetValue('', prog_id, RegistryValueKind.String)

            with open_or_create(base, prog_id_path / 'CLSID') as clsid:
                clsid.SetValue('', uuid, RegistryValueKind.String)

    def define_prog_id(self):
        with self.get_hkcu() as base:
            self._define_prog_id(base, self.prog_id, self.uuid)

    def define_wow_uuid_clsid(self):
        with self.get_hkcu_wow() as base:
            self._define_wow_uuid_clsid(base)

    def _define_wow_uuid_clsid(self, base):
        uuid_path = PathString('Software') / 'Classes' / 'CLSID' / self.uuid
        with open_or_create(base, uuid_path) as uuid:
            uuid.SetValue('', self.prog_id, RegistryValueKind.String)

        with open_or_create(base, uuid_path / 'ProgId') as uuid:
            uuid.SetValue('', self.prog_id, RegistryValueKind.String)

        with open_or_create(base, uuid_path / 'Implemented Categories' / '{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}'):
            pass

        with open_or_create(base, uuid_path / 'InprocServer32') as serv:
            serv.SetValue('', 'mscoree.dll')
            serv.SetValue('ThreadingModel', 'Both')
            serv.SetValue('Class', self.prog_id)

            p = self.assembly_properties
            serv.SetValue('Assembly', p.full_name)
            serv.SetValue('RuntimeVersion', p.runtime_version)
            serv.SetValue('CodeBase', p.codebase_uri)

        with open_or_create(base, uuid_path / 'InprocServer32' / self.assembly_properties.version) as version:
            version.SetValue('Class', self.prog_id)

            p = self.assembly_properties
            version.SetValue('Assembly', p.full_name)
            version.SetValue('RuntimeVersion', p.runtime_version)
            version.SetValue('CodeBase', p.codebase_uri)

    def register_assembly(self):
        self.load_assembly_attributes()
        self.define_prog_id()
        self.define_wow_uuid_clsid()

    def unregister_assembly(self):
        prog_id_path = PathString('Software') / 'Classes' / self.prog_id
        with self.get_hkcu() as base:
            base.DeleteSubKeyTree(prog_id_path, False)

        uuid_path = PathString('Software') / 'Classes' / 'CLSID' / self.uuid
        with self.get_hkcu_wow() as base:
            base.DeleteSubKeyTree(uuid_path, False)

def office_default_path(app_name):
    return PathString('Software') / 'Microsoft' / 'Office' / app_name / 'Addins'

class AddinRegService(object):
    def __init__(self, prog_id, friendly_name, description, addins_regpath, load_behavior):
        self.prog_id = prog_id
        self.friendly_name = friendly_name
        self.description = description
        self.addins_regpath = addins_regpath
        self.load_behavior = load_behavior

    def get_hkcu(self, view=RegistryView.Default):
        return RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view)

    def register_addin(self):
        with self.get_hkcu() as base:
            self._register_addin(base)

    def _register_addin(self, base):
        prog_id_path = self.addins_regpath / self.prog_id
        with open_or_create(base, prog_id_path) as prog_id:
            prog_id.SetValue('LoadBehavior', self.load_behavior, RegistryValueKind.DWord)
            prog_id.SetValue('FriendlyName', self.friendly_name)
            prog_id.SetValue('Description', self.description)

    def unregister_addin(self):
        prog_id_path = self.addins_regpath / self.prog_id
        with self.get_hkcu() as base:
            base.DeleteSubKeyTree(prog_id_path, False)



class QueryRegService(object):

    def get_hklm(self, view=RegistryView.Default):
        return RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view)

    # def get_hkcu(self, view=RegistryView.Default):
    #     return RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view)
    
    def _get_outlook_bitness_for_base(self, base, app_paths):
        with open_key(base, app_paths) as path:
            return path.GetValue('Bitness')

    def _get_path_for_base(self, base, app_name):
        app_paths = PathString('Software') / 'Microsoft' / 'Windows' / 'CurrentVersion' / 'App Paths' / app_name
        with open_key(base, app_paths) as path:
            return path.GetValue('')
    
    def get_app_path(self, app_name='excel.exe'):
        with self.get_hklm() as base:
            return self._get_path_for_base(base, app_name)

        # NOTE: If office is installed from Microsoft Store the app path exists in HKCU, but
        #       the path is under Program Files\WindowsApps\... which is not readable, so no need to check this
        # with self.get_hkcu() as base:
        #     try:
        #         return self._get_path_for_base(base, app_name)
        #     except KeyError:
        #         pass
        # raise KeyError("no path in registry found for %s" % app_name)

    def get_outlook_bitness(self):
        paths = [
            PathString('Software') / 'Microsoft' / 'Office' / 'ClickToRun' / 'REGISTRY' / 'MACHINE' / 'Software' / 'Microsoft' / 'Office' / '16.0' / 'Outlook',
            PathString('Software') / 'Microsoft' / 'Office' / '16.0' / 'Outlook',
            PathString('Software') / 'Microsoft' / 'Office' / '15.0' / 'Outlook',
            PathString('Software') / 'Microsoft' / 'Office' / '14.0' / 'Outlook',
        ]
        with self.get_hklm() as base:
            for path in paths:
                try:
                    return self._get_outlook_bitness_for_base(base, path)
                except KeyError:
                    continue
        return None