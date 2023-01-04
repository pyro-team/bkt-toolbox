# -*- coding: utf-8 -*-
'''
Created on 21.12.2017

@author: fstallmann
'''



import bkt

from bkt.callbacks import WpfActionCallback
from ..agenda import ToolboxAgenda


class AgendaPopup(bkt.ui.WpfWindowAbstract):
    # _filename = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'popups', 'agenda.xaml')
    _xamlname = 'agenda_popup'
    '''
    class representing a popup-dialog for a agenda textbox
    '''
    
    def __init__(self, context=None):
        self.IsPopup = True

        super(AgendaPopup, self).__init__(context)

    def btntab(self, sender, event):
        try:
            self._context.ribbon.ActivateTab('bkt_context_tab_agenda')
        except:
            bkt.message.error("Tab-Wechsel aus unbekannten Gründen fehlgeschlagen.")

    @WpfActionCallback
    def btnupdate(self, sender, event):
        try:
            ToolboxAgenda.update_or_create_agenda_from_slide(self._context.slide, self._context)
        except:
            bkt.message.error("Agenda-Update aus unbekannten Gründen fehlgeschlagen.")

    @staticmethod
    def double_click(shape, context):
        try:
            context.ribbon.ActivateTab('bkt_context_tab_agenda')
        except:
            bkt.message.error("Tab-Wechsel aus unbekannten Gründen fehlgeschlagen.")


#initialization function called by contextdialogs.py
create_window = AgendaPopup
trigger_doubleclick = AgendaPopup.double_click


#old method:
# # register dialog
# bkt.powerpoint.context_dialogs.register_dialog(
#     bkt.contextdialogs.ContextDialog(
#         id=TOOLBOX_AGENDA_POPUP,
#         module=None,
#         window_class=AgendaPopup,
#         # dblclick_func=AgendaPopup.double_click,
#     )
# )