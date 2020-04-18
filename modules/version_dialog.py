# -*- coding: utf-8 -*-
'''
Created on 2020-04-18
@author: Florian Stallmann
'''

from __future__ import absolute_import

import os.path

import System

import modules.settings as settings

import bkt
import bkt.ui
# notify_property = bkt.ui.notify_property


class ViewModel(bkt.ui.ViewModelSingleton):
    def __init__(self):
        super(ViewModel, self).__init__()

        resource_path = os.path.normpath(os.path.join(os.path.dirname(os.path.realpath(__file__)), "..", "resources", "bkt_logo", "BKT Logo 1.0.png"))
        self.bkt_logo = System.Uri(resource_path)
        # self.bkt_logo = bkt.ui.load_bitmapimage("bkt_logo")
        self.bkt_version = "v" + bkt.version_tag_name
        self.bkt_update_label = settings.BKTUpdates.get_label_update()

class VersionDialog(bkt.ui.WpfWindowAbstract):
    _xamlname = 'version_dialog'
    _vm_class = ViewModel

    def __init__(self, context):
        super(VersionDialog, self).__init__(context)

    def open_website(self, sender, event):
        settings.BKTInfos.open_website()

    def check_for_updates(self, sender, event):
        settings.BKTUpdates.manual_check_for_updates(self._context)

    def show_debug_message(self, sender, event):
        settings.BKTInfos.show_debug_message(self._context)