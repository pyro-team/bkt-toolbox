# -*- coding: utf-8 -*-
'''
Core of the BKT-Framework providing ribbon customization and library functions.
The frameworks requires IronPython. The .NET part of the addin is in ../dotnet/.

Created on 10.09.2013
@author: cschmitt, rdebeerst
'''



__author__ = "Christoph Schmitt, Ruben Debeerst, Thomas Weuffel, Florian Stallmann"
__copyright__ = "Copyright 2019 Christoph Schmitt, Ruben Debeerst, Thomas Weuffel, Florian Stallmann"
__license__ = "MIT"
__version__ = "3.0.1"
__release__ = "BKT r24-01-10"


# NOTE: Use StandardLib.dll as alternative to /bin/Lib, but seems to have problems with wpf/fluent
# import clr
# clr.AddReference("IronPython.StdLib")


# FIXME: clrmock needs to be fixed so this is currently commented out!
# import platform
# if not 'IronPython' in platform.python_implementation():
#     from bkt.compat import clrmock
#     clrmock.inject_mock()
#     del clrmock
# del platform


# set locale
import locale
locale.setlocale(locale.LC_ALL, '') #auto detect locale
#NOTE: better to set to office UI language? mapping from office lang id to locale string required...
#context.app.LanguageSettings.LanguageID(2) #MsoAppLanguageID=msoLanguageIDUI


# make the followig classes and decorators accessible 
# via bkt.xxx after 'import bkt'

# import modules with less dependencies first
from bkt.helpers import config, settings #no internal dependencies

from bkt.callbacks import CallbackTypes, Callback, CallbackLazy #imports helpers
get_enabled_auto = CallbackTypes.get_enabled.dotnet_name #convenience name

from bkt.apps import AppEvents #imports callbacks and context (imports helpers)

from bkt.ribbon import mso #imports callbacks, dotnet, xml (imports dotnet)
from bkt.appui import (excel,
                      outlook,
                      powerpoint,
                      word,
                      visio) #imports helpers, xml, ribbon, contextdialogs (imports library.algorithms, dotnet), taskpane (import helpers, ribbon, xml), ui (imports dotnet, library.wpf, helpers)

from bkt.library.system import (
    apply_delta_on_ALT_key,
    get_key_state,
    KeyCodes,
    message,
    MessageBox
)
