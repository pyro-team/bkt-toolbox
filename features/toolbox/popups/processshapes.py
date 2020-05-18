# -*- coding: utf-8 -*-
'''
Created on 21.12.2017

@author: fstallmann
'''

from __future__ import absolute_import

import bkt

from ..processshapes import ProcessChevrons



class ProcessChevronsPopup(bkt.ui.WpfWindowAbstract):
    # _filename = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'popups', 'process_shapes.xaml')
    _xamlname = 'shape_process_popup'
    '''
    class representing a popup-dialog for a process chevron shapes
    '''
    
    def __init__(self, context=None):
        self.IsPopup = True
        self._context = context

        super(ProcessChevronsPopup, self).__init__()

    def btnplus(self, sender, event):
        try:
            ProcessChevrons.add_chevron(list(iter(self._context.selection.ShapeRange)))
        except:
            bkt.message.error("Funktion aus unbekannten Gründen fehlgeschlagen.")
            # bkt.helpers.exception_as_message()

    def btnminus(self, sender, event):
        try:
            ProcessChevrons.remove_chevron(list(iter(self._context.selection.ShapeRange)))
        except:
            bkt.message.error("Funktion aus unbekannten Gründen fehlgeschlagen.")
            # bkt.helpers.exception_as_message()

#initialization function called by contextdialogs.py
create_window = ProcessChevronsPopup