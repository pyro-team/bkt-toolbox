# -*- coding: utf-8 -*-
'''
Created on 10.09.2013

@author: cschmitt, rdebeerst
'''

#Use StandardLib.dll as alternative to /bin/Lib, but seems to have problems with wpf/fluent
# import clr
# clr.AddReference("IronPython.StdLib")

import sys

# print sys.version
if not 'IronPython' in sys.version:
    from bkt.compat import clrmock
    clrmock.inject_mock()
    del clrmock

del sys

full_version = 'BKT r19-07-05'

# import modules with less dependencies first
from bkt.annotation import FeatureContainer
from bkt.ribbon import mso

# make the followig classes and decorators accessible 
# via bkt.xxx after 'import bkt'

from bkt.helpers import (
    config,
    settings
)

from bkt.library.general import (
    apply_delta_on_ALT_key
)

#@deprecated
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

from bkt.appui import (excel,
                      outlook,
                      powerpoint,
                      word,
                      visio)

from bkt.callbacks import CallbackTypes, Callback
from bkt.apps import AppEvents

