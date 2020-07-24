# -*- coding: utf-8 -*-
'''
Created on 24.07.2020

@author: fstallmann
'''

from __future__ import absolute_import, print_function

import os

import System.Environment

from . import reg
from . import helper


class BitnessChecker(object):
    @classmethod
    def get_bitness(cls):
        '''
        Returns true if office-32-bit is running on 64-bit windows machine, or if it is a 32-bit machine.
        Note: According to https://support.microsoft.com/en-us/help/2778964/addins-for-office-programs-may-be-registered-under-the-wow6432node
        this is not required for addin registration (we register in HKCU), but it is required for register of DLL in Classes/CLSID/* (refer to reg.py)
        '''
    
        # If os is 32-bit no need to do further check
        # Alternative python way: if not helper.is_64bit_os():
        if not System.Environment.Is64BitOperatingSystem:
            return False
        
        # Method 1: Get outlook bitness reg key (fastest if outlook is installed)
        try:
            return cls._check_outlook_reg()
        except Exception as e:
            # helper.exception_as_message()
            helper.log("bitness-method 1 failed: %s" % e)
        
        # Method 2: Try to find binaries via registry and get type
        try:
            return cls._check_binary_types()
        except Exception as e:
            # helper.exception_as_message()
            helper.log("bitness-method 2 failed: %s" % e)
        
        # Method 3: Load interop assemblies, start app instance and get product code GUID
        # This method is required if office is installed via Microsoft Store as no readable exe path exists!
        try:
            return cls._check_interop()
        except Exception as e:
            # helper.exception_as_message()
            helper.log("bitness-method 3 failed: %s" % e)
        
        #fallback if all methods failed
        return True

    @staticmethod
    def _check_outlook_reg():
        path_getter = reg.QueryRegService()
        outlook_bitness = path_getter.get_outlook_bitness()

        assert outlook_bitness in ("x86", "x64"), 'unkown outlook bitness key: %s' % outlook_bitness
        return outlook_bitness == "x86"

    @staticmethod
    def _check_binary_types():
        office_is_32 = set()
        path_getter = reg.QueryRegService()
        apps = ['powerpnt.exe',
                'excel.exe']
        for app_exe in apps:
            try:
                app_path = path_getter.get_app_path(app_exe)
            except KeyError:
                helper.log("no path in registry found for %s" % app_exe)
                continue
            if os.path.isfile(app_path):
                try:
                    office_is_32.add(not helper.is_64bit_exe(app_path))
                except:
                    helper.log("failed to get binary type for: %s" % app_path)
                else:
                    if True in office_is_32:
                        break
            else:
                helper.log("file not found: %s" % app_path)

        assert len(office_is_32) > 0, 'no result for binaries types'
        return True in office_is_32

    @staticmethod
    def _check_interop():
        import clr

        office_is_32 = set()
        iop_base = 'Microsoft.Office.Interop.'
        apps = ['PowerPoint',
                'Excel']
        for app_name in apps:
            iop_name = iop_base + app_name
            try:
                clr.AddReference(iop_name)
                # module = None
                # # FIXME: this is ugly, but __import__(iop_name) does not seem to work
                # exec 'import ' + iop_name + ' as module'
                import Microsoft #no need to import the whole iop name
                module = getattr(Microsoft.Office.Interop, app_name)
                app = module.ApplicationClass()
                try:
                    #NOTE: As of Office 2016 PPT will return "Windows (64-bit)" no matter of the Office bitness, but Excel returns "Windows (32-bit)"
                    # office_is_32.add(app.OperatingSystem.startswith('Windows (32-bit)'))
                    #NOTE: Using GUID should be more reliable, explanation: https://docs.microsoft.com/en-us/office/troubleshoot/miscellaneous/numbering-scheme-for-product-guid
                    office_is_32.add(app.ProductCode[20] == '0')
                except:
                    helper.log("failed to get bitness of %s" % app_name)
                else:
                    if True in office_is_32:
                        break
                finally:
                    app.Quit()
            except:
                # helper.exception_as_message()
                helper.log("failed to load app %s" % app_name)
        
        assert len(office_is_32) > 0, 'no result for interop method'
        return True in office_is_32