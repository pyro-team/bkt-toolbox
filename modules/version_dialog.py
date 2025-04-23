# -*- coding: utf-8 -*-
'''
Created on 2020-04-18
@author: Florian Stallmann
'''



import os.path

import System
from System.Windows import Visibility

import modules.settings as settings

import bkt
import bkt.ui
# notify_property = bkt.ui.notify_property


class ViewModel(bkt.ui.ViewModelSingleton):
    def __init__(self):
        super(ViewModel, self).__init__()

        resource_path = bkt.helpers.bkt_base_path_join("resources", "bkt_logo", "BKT Logo 1.0.png")
        self.bkt_logo = System.Uri(resource_path)
        # self.bkt_logo = bkt.ui.load_bitmapimage("bkt_logo")
        self.bkt_version = "v" + bkt.__version__
        self.bkt_update_available = settings.BKTUpdates.is_update_available()
        self.bkt_update_label = settings.BKTUpdates.get_label_update()
        
        self.bkt_license_text = "Die BKT ist Open Source lizensiert unter der {}-Lizenz.".format(bkt.__license__)
        self.bkt_copyright_text = "{}.".format(bkt.__copyright__)

        self.bkt_branded_visible = Visibility.Collapsed
        self.bkt_branding_text = ""

        branding = settings.BKTInfos.get_branding_info()
        if branding.is_branded:
            self.bkt_branded_visible = Visibility.Visible
            self.bkt_branding_text = "Diese BKT-Version wurde modifiziert für {}.".format(branding.brand_name)

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