# -*- coding: utf-8 -*-
'''
Created on 21.12.2017

@author: fstallmann
'''

from __future__ import absolute_import

import bkt

from bkt.callbacks import WpfActionCallback
from ..linkshapes import LinkedShapes





class LinkedShapePopup(bkt.ui.WpfWindowAbstract):
    # _filename = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'popups', 'linkedshape.xaml')
    _xamlname = 'linkshapes_popup'
    '''
    class representing a popup-dialog for a linked shape
    '''
    
    def __init__(self, context=None):
        self.IsPopup = True

        super(LinkedShapePopup, self).__init__(context)

    def btntab(self, sender, event):
        try:
            self._context.ribbon.ActivateTab('bkt_context_tab_linkshapes')
        except:
            bkt.message.error("Tab-Wechsel aus unbekannten Gründen fehlgeschlagen.")

    @WpfActionCallback
    def btnsync_text(self, sender, event):
        try:
            LinkedShapes.text_linked_shapes(self._context.shapes[-1], self._context)
        except:
            bkt.message.error("Aktualisierung aus unbekannten Gründen fehlgeschlagen.")

    @WpfActionCallback
    def btnsync_possize(self, sender, event):
        try:
            LinkedShapes.align_linked_shapes(self._context.shapes[-1], self._context)
            LinkedShapes.size_linked_shapes(self._context.shapes[-1], self._context)
        except:
            bkt.message.error("Aktualisierung aus unbekannten Gründen fehlgeschlagen.")

    @WpfActionCallback
    def btnsync_format(self, sender, event):
        try:
            LinkedShapes.format_linked_shapes(self._context.shapes[-1], self._context)
        except:
            bkt.message.error("Aktualisierung aus unbekannten Gründen fehlgeschlagen.")

    def btnnext(self, sender, event):
        try:
            LinkedShapes.goto_linked_shape(self._context.shapes[-1], self._context)
        except:
            bkt.message.error("Funktion aus unbekannten Gründen fehlgeschlagen.")

    @staticmethod
    def double_click(shape, context):
        try:
            context.ribbon.ActivateTab('bkt_context_tab_linkshapes')
        except:
            bkt.message.error("Tab-Wechsel aus unbekannten Gründen fehlgeschlagen.")



#initialization function called by contextdialogs.py
create_window = LinkedShapePopup
trigger_doubleclick = LinkedShapePopup.double_click


#old method:
# # register dialog
# bkt.powerpoint.context_dialogs.register_dialog(
#     bkt.contextdialogs.ContextDialog(
#         id=BKT_LINK_UUID,
#         module=None,
#         window_class=LinkedShapePopup,
#         dblclick_func=LinkedShapePopup.double_click,
#     )
# )