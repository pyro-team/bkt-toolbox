# -*- coding: utf-8 -*-
'''
Created on 21.12.2017

@author: fstallmann
'''



import logging

import bkt

from bkt.callbacks import WpfActionCallback
from ..harvey import harvey_balls



class HarveyPopup(bkt.ui.WpfWindowAbstract):
    # _filename = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'popups', 'harvey.xaml')
    _xamlname = 'harvey'
    '''
    class representing a popup-dialog for a harvey ball shape
    '''
    
    def __init__(self, context=None):
        self.IsPopup = True

        super(HarveyPopup, self).__init__(context)

    def btntab(self, sender, event):
        try:
            self._context.ribbon.ActivateTab('bkt_context_tab_harvey')
        except:
            bkt.message.error("Tab-Wechsel aus unbekannten Gründen fehlgeschlagen.")

    @WpfActionCallback
    def btnplus(self, sender, event):
        try:
            harvey_balls.harvey_percent_setter_popup(list(iter(self._context.selection.ShapeRange)))
            self._context.ribbon.Invalidate()
        except:
            bkt.message.error("Funktion aus unbekannten Gründen fehlgeschlagen.")
            # bkt.helpers.exception_as_message()

    @WpfActionCallback
    def btnminus(self, sender, event):
        try:
            harvey_balls.harvey_percent_setter_popup(list(iter(self._context.selection.ShapeRange)), inc=False)
            self._context.ribbon.Invalidate()
        except:
            bkt.message.error("Funktion aus unbekannten Gründen fehlgeschlagen.")
            # bkt.helpers.exception_as_message()

    @staticmethod
    def double_click(shape, context):
        try:
            context.ribbon.ActivateTab('bkt_context_tab_harvey')
        except:
            bkt.message.error("Tab-Wechsel aus unbekannten Gründen fehlgeschlagen.")


#initialization function called by contextdialogs.py
create_window = HarveyPopup
trigger_doubleclick = HarveyPopup.double_click


#old method:
# # register dialog
# bkt.powerpoint.context_dialogs.register_dialog(
#     bkt.contextdialogs.ContextDialog(
#         id=HarveyBalls.BKT_HARVEY_DIALOG_TAG,
#         module=None,
#         window_class=HarveyPopup,
#         dblclick_func=HarveyPopup.double_click,
#     )
# )