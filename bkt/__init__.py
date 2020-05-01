# -*- coding: utf-8 -*-
'''
Core of the BKT-Framework providing ribbon customization and library functions.
The frameworks requires IronPython. The .NET part of the addin is in ../dotnet/.

Created on 10.09.2013
@author: cschmitt, rdebeerst
'''

from __future__ import absolute_import

__author__ = "Christoph Schmitt, Ruben Debeerst, Florian Stallmann"
__copyright__ = "Copyright 2019 Christoph Schmitt, Ruben Debeerst, Florian Stallmann"
__license__ = "MIT"
__version__ = "2.7.0"
__release__ = "BKT r20-03-13"


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



# make the followig classes and decorators accessible 
# via bkt.xxx after 'import bkt'

# import modules with less dependencies first
from bkt.callbacks import CallbackTypes, Callback #no internal dependencies

from bkt.helpers import config, settings #no internal dependencies
from bkt.apps import AppEvents #imports callbacks and context (imports helpers)

from bkt.ribbon import mso #imports callbacks, dotnet, xml (imports dotnet)
from bkt.appui import (excel,
                      outlook,
                      powerpoint,
                      word,
                      visio) #imports helpers, xml, ribbon, contextdialogs (imports library.algorithms, dotnet), taskpane (import helpers, ribbon, xml), ui (imports dotnet, library.wpf, helpers)

from bkt.library.system import (
    apply_delta_on_ALT_key,
    get_key_state
)

# enable legacy annotations syntax with decorators
if config.enable_legacy_syntax or False:
    from bkt.annotation import FeatureContainer #@deprecated
    # @deprecated
    from bkt.decorators import (
        # public classes
        #EventHandler,
        #CallableContextInformation,
        #Resources,
        
        # decorators for ribbon-classes
        uicontrol,
        use,
        tab,
        group,
        menu,
        box,
        button,
        large_button,
        toggle_button,
        edit_box,
        spinner_box,
        gallery,
        combo_box,
        #dialog_box_launcher,
        #check_box,
        #incdec_edit_box,
        
        # decorators for ribbon attributes
        configure,
        image,
        image_mso,
        
        # decorators for callbacks
        callback_type,
        increment,
        decrement,
        on_change,
        get_text,
        
        # decorators for context
        arg_python_addin,
        arg_ribbon_id,
        arg_context,
        arg_presentation,
        arg_shape,
        arg_shapes,
        arg_shapes_limited,
        arg_slide,
        arg_slides,
        arg_slides_limited,
        arg_page_shapes,
        require_text,
        no_transaction,
        uuid
        # decorators for office apps
    )

