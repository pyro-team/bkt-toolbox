# -*- coding: utf-8 -*-
'''
Created on 26.02.2023

@author: fstallmann
'''


import bkt

class MasterShapeIndicator(bkt.ui.WpfWindowAbstract):
    # _filename = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'master_shape.xaml')
    _xamlname = 'master_shape'
    IsPopup = True

    # def __init__(self, context=None):
    #     self.IsPopup = True
    #     self._context = context

    #     super(MasterShapeIndicator, self).__init__()

class MasterShapeDialog(bkt.contextdialogs.ContextDialog):
    def __init__(self, arranger):
        super(MasterShapeDialog, self).__init__("MASTER")
        self.wnd = None
        self.arranger = arranger
    
    def create_window(self, context):
        # if not self.wnd:
        #     self.wnd = MasterShapeIndicator(context)
        # return self.wnd
        return MasterShapeIndicator(context)

    def get_master_shape(self, shapes):
        return self.arranger.get_master_for_indicator(shapes)