# -*- coding: utf-8 -*-
'''
Various helper function, global config and settings parser

Created on 10.09.2013
@author: cschmitt
'''


from __future__ import absolute_import, print_function

import os.path
import logging
import traceback

import ConfigParser #required for config.txt file
import shelve #required for BKTShelf

from functools import wraps

BKT_BASE = os.path.realpath(os.path.join(os.path.dirname(__file__), ".."))



# ==============================
# = Typical programming helpers =
# ==============================

def memoize(func):
    ''' Memoize a functions return value for each set of args (kwargs not supported) '''
    cache = {}
    @wraps(func)
    def memoizer(*args):
        try:
            return cache[args]
        except KeyError:
            result = cache[args] = func(*args)
            return result
    return memoizer

@memoize
def snake_to_lower_camelcase(string):
    ''' convert on_action to onAction, but leave onAction as is '''
    if '_' in string:
        parts = string.split('_')
        return parts[0].lower() + ''.join(x.title() for x in parts[1:])
    else:
        return string[0].lower() + string[1:]

@memoize
def snake_to_upper_camelcase(string):
    ''' convert on_action to OnAction, but leave OnAction as is '''
    if '_' in string:
        return ''.join(x.title() for x in string.split('_'))
    else:
        return string[0].upper() + string[1:]


def endings_to_windows(text, prepend="", prepend_first=""):
    ''' ensures that all line endings are using windows CRLF format '''
    def _iter():
        first = True
        for line in text.splitlines():
            if first:
                line = prepend_first + line
                first = False
            else:
                line = prepend + line
            yield line
    
    res = '\r\n'.join(_iter())
    #terminal line ending is removed by splitlines
    if text.endswith(("\r", "\n", "\r\n")):
        res += "\r\n"
    return res

def endings_to_unix(text):
    ''' ensures that all line endings are using unix LF format '''
    res = '\n'.join(text.splitlines())
    #terminal line ending is removed by splitlines
    if text.endswith(("\r", "\n", "\r\n")):
        res += "\n"
    return res


# ==================
# = Error messages =
# ==================

def message(*args, **kwargs):
    #only for backwards compatibility
    from bkt import message
    return message(*args, **kwargs)


def log(s):
    import bkt.console
    logging.warning(s)
    bkt.console.show_message(s)

def exception_as_message(additional_message=None):
    from cStringIO import StringIO
    import traceback

    import bkt.console
    import bkt.ui

    fd = StringIO()
    if additional_message:
        print(additional_message,file=fd)
    traceback.print_exc(file=fd)
    traceback.print_exc()

    bkt.console.show_message(endings_to_windows(fd.getvalue()))


# ==============================
# = Typical os.path operations =
# ==============================

def file_base_path_join(base_file, *args):
    return os.path.realpath(os.path.join(os.path.dirname(base_file), *args))

def bkt_base_path_join(*args):
    return os.path.realpath(os.path.join(BKT_BASE, *args))




# ========================
# = Load config.txt file =
# ========================

class BKTConfigParser(ConfigParser.ConfigParser):
    ''' Global configuration is stored in config.txt file '''
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
        elif value.lower() in ['false', 'no', 'off', '0']:
            return False
        elif value.lower() in ['true', 'yes', 'on', '1']:
            return True
        elif value[0] != "\n":
            return value
        else:
            return value[1:].split("\n")

    def get_smart(self, attr, default=None, attr_type=str):
        '''
        Method to get config-values and force a particular data type, return
        default value on error. This method does not work for lists.
        '''
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
            self.set('BKT', option, "\n" + "\n".join(str(v) for v in value))
        else:
            self.set('BKT', option, str(value)) #always transform to string, otherwise cannot access the value in same session anymore

        # write config file
        with open(self.config_filename, "wb") as configfile:
            self.write(configfile)


# load config
config_filename=bkt_base_path_join("config.txt")
config = BKTConfigParser(config_filename)
if os.path.exists(config_filename):
    config.read(config_filename)
else:
    config.add_section('BKT')





# ======================================
# = Helper to get configurable folders =
# ======================================

def ensure_folders_exist(folder_path):
    if not os.path.isdir(folder_path):
        from os import makedirs
        makedirs(folder_path)
    return folder_path


def get_fav_folder(*args):
    folder = config.local_fav_path or False
    if not folder:
        #FIXME: we could also get this with pure python using ctypes, refer to SHGetKnownFolderPath (https://docs.microsoft.com/en-us/windows/win32/api/shlobj_core/nf-shlobj_core-shgetknownfolderpath)
        from System import Environment
        folder = ensure_folders_exist( os.path.join(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "BKT-Favoriten") )
        # NOTE: os.path.expanduser("~")/Documents doesnt work if Documents folder has been moved by user or by OneDrive installation
        # folder = ensure_folders_exist( os.path.realpath(os.path.join(os.path.expanduser("~"), "Documents", "BKT-Favoriten")) )
    args = args or tuple()
    args = (folder,)+args
    return os.path.join(*args)

def get_cache_folder(*args):
    folder = config.local_cache_path or False
    if not folder:
        folder = ensure_folders_exist( bkt_base_path_join("resources","cache") )
    args = args or tuple()
    args = (folder,)+args
    return os.path.join(*args)

def get_settings_folder(*args):
    folder = config.local_settings_path or False
    if not folder:
        folder = ensure_folders_exist( bkt_base_path_join("resources","settings") )
    args = args or tuple()
    args = (folder,)+args
    return os.path.join(*args)





# =========================
# = BKT version of shelve =
# =========================

class BKTShelf(shelve.DbfilenameShelf):
    ''' BKT-style shelf with auto repair on corruption used for settings database and for caches '''

    def __init__(self, filename):
        self._filename = filename
        shelve.DbfilenameShelf.__init__(self, filename, protocol=2)
    
    def get(self, key, default=None):
        try:
            # super(BKTShelf, self).get(key, default) #doesnt work as Shelf is not a new-style object
            if key in self.dict:
                return self[key]
            return default
        except EOFError:
            logging.error("EOF-Error in shelf file %s for getting key %s. Reset to default value: %s", self._filename, key, default)
            if config.show_exception and not key.startswith("bkt.console."):
                #if key starts with bkt.console its not possible to show exception in console as error happended during console initialization
                exception_as_message("Shelf file {} corrupt for key {}. Trying to repair now.".format(self._filename, key))

            #shelf database corrupt, trying to fix it
            if default is None:
                del self[key]
            else:
                self[key] = default

            return default




# =============================================
# = Lazy loading app-specific settings shelve =
# =============================================

class BKTSettings(BKTShelf):
    ''' App-specific settings are stored as shelve object that supports various python data formats '''

    def __init__(self):
        self._filename = None
        shelve.Shelf.__init__(self, shelve._ClosedDict(), protocol=2)
    
    def open(self, filename):
        import anydbm
        self._filename = get_settings_folder(filename)
        try:
            self.dict = anydbm.open(self._filename, 'c')
        except:
            logging.exception("error reading bkt settings")
            exception_as_message()
            self.dict = dict() #fallback to empty dict

#load global setting database
settings = BKTSettings()




# =======================================
# = Helper to create caches with shelve =
# =======================================

class BKTCacheFactory(object):
    ''' Factory to create caches that are automatically closed on bkt unload '''

    def __init__(self):
        self._caches = dict()

    def get(self, name):
        try:
            return self._caches[name]
        except KeyError:
            cache_file = get_cache_folder("%s.cache" % name)
            self._caches[name] = cache = BKTShelf(cache_file)
            return cache
    
    def _close(self, name):
        try:
            self._caches[name].close()
            del self._caches[name]
        except KeyError:
            pass
    
    def close(self, name=None):
        if name is None:
            #close all
            for name in self._caches.keys():
                self._close(name)
        else:
            self._close(name)

caches = BKTCacheFactory()




# =======================================
# = Find resources (images, xaml files) =
# =======================================

class Resources(object):
    ''' Encapsulated path resolution for file resources (such as images) '''
    root_folders = []
    images = None
    xaml = None
    
    def __init__(self, category, suffix):
        self.category = category
        self.suffix = suffix
        
        try:
            self._cache = caches.get("resources.%s"%category)
        except:
            logging.exception("Loading resource cache failed")
            
    def locate(self, name):
        try:
            return self._cache[name]
        except KeyError:
            logging.info("Locate %s resource: %s.%s", self.category, name, self.suffix)
            for root_folder in self.root_folders:
                path = os.path.join(root_folder, self.category, name + '.' + self.suffix)
                if os.path.exists(path):
                    self._cache[name] = path
                    # self._cache.sync() #sync after each change as .close() is never called
                    return path
            return None
        except:
            logging.exception("Unknown error reading from resource cache")
            return None
    
    @staticmethod
    def bootstrap():
        Resources.root_folders = [ bkt_base_path_join('resources') ]
        Resources.images = Resources("images", "png")
        Resources.xaml = Resources("xaml", "xaml")

Resources.bootstrap()




# ===========================
# = Bitwise boolean storage =
# ===========================


class BitwiseValueAccessor(object):
    '''
    Provides an easy way to access boolean options stored as single integer value using bitwise operators.
    The integer bitvalue can be retrieved via get_bitvalue(). Attribute and item notation are supported.
    All options can be retrieved with as_dict() function. New options can be added with add_option(name, value).
    '''
    def __init__(self, bitvalue=0, attributes=[], settings_key=None):
        if settings_key:
            self.__bitvalue = settings.get(settings_key, 0)
        else:
            self.__bitvalue = bitvalue
        self._settings_key = settings_key
        self._attributes = attributes
        self._attr_dict = {k: 2**i for i, k in enumerate(attributes)}
    
    @property
    def _bitvalue(self):
        return self.__bitvalue
    @_bitvalue.setter
    def _bitvalue(self, value):
        self.__bitvalue = value
        if self._settings_key:
            settings[self._settings_key] = value

    def __repr__(self):
        return "<BitwiseValueAccessor bitvalue=%d attributes=%r>" % (self._bitvalue, self._attributes)
    
    def __getattr__(self, attr):
        try:
            return self.__getitem__(attr)
        except KeyError:
            raise AttributeError(attr)

    def __setattr__(self, attr, value):
        if attr.startswith("_"):
            super(BitwiseValueAccessor, self).__setattr__(attr, value)
        else:
            try:
                self.__setitem__(attr, value)
            except KeyError:
                raise AttributeError(attr)
    
    def __getitem__(self, key):
        option = self._attr_dict[key]
        return self._bitvalue & option == option

    def __setitem__(self, key, value):
        option = self._attr_dict[key]
        if value:
            self._bitvalue = self._bitvalue | option
        else:
            self._bitvalue = self._bitvalue ^ option

    def __dir__(self):
        return sorted(set( dir(type(self)) + self.__dict__.keys() + self._attributes ))
    
    def __contains__(self, value):
        return value in self._attributes

    def __len__(self):
        return len(self._attributes)
    
    def __iter__(self):
        for k in self._attributes:
            yield k, getattr(self, k)

    def get_bitvalue(self):
        return self._bitvalue

    def as_dict(self):
        return {k: getattr(self, k) for k in self._attributes}
    
    def add_option(self, name, value=False):
        self._attributes.append(name)
        self._attr_dict[name] = 2**(len(self._attributes)-1)
        if value:
            self.__setitem__(name, value)
