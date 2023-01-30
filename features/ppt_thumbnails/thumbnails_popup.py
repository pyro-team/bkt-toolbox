# -*- coding: utf-8 -*-
'''
Created on 2023-01-16
@author: Florian Stallmann
'''

import logging

from System.Windows import Visibility

import bkt

from .thumbnails_model import Thumbnailer
from bkt.callbacks import WpfActionCallback


class ThumbnailPopup(bkt.ui.WpfWindowAbstract):
    # _filename = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'thumbnail.xaml')
    _xamlname = 'thumbnail'
    '''
    class representing a popup-dialog for a thumbnail shape
    '''
    
    def __init__(self, context=None):

        self.IsPopup = True

        super(ThumbnailPopup, self).__init__(context)

        if context.app.activewindow.selection.shaperange.count > 1:
            self.btngoto.Visibility = Visibility.Collapsed

    @WpfActionCallback
    def btnrefresh(self, sender, event):
        try:
            shapes = self._context.shapes
            if len(shapes) == 1:
                Thumbnailer.shape_refresh(shapes[0], self._context.app)
            else:
                Thumbnailer.shapes_refresh(shapes, self._context.app)
        except:
            bkt.message.error("Thumbnail-Aktualisierung aus unbekannten Gründen fehlgeschlagen.", "BKT: Thumbnails")
            logging.exception("Thumbnails: Error in popup!")

    @WpfActionCallback
    def btngoto(self, sender, event):
        try:
            Thumbnailer.goto_ref(self._context.shape, self._context.app)
        except:
            bkt.message.error("Fehler beim Öffnen der Folienreferenz.", "BKT: Thumbnails")
            logging.exception("Thumbnails: Error in popup!")

    @WpfActionCallback
    def btntoggleco(self, sender, event):
        try:
            Thumbnailer.toggle_content_only(self._context.shape, self._context.app)
        except:
            bkt.message.error("Fehler beim Wechsel des Thumbnail-Inhalts.", "BKT: Thumbnails")
            logging.exception("Thumbnails: Error in popup!")

    @WpfActionCallback
    def btnfixar(self, sender, event):
        try:
            Thumbnailer.reset_aspect_ratio(self._context.shape)
        except:
            bkt.message.error("Fehler beim Zurücksetzen des Seitenverhältnisses.", "BKT: Thumbnails")
            logging.exception("Thumbnails: Error in popup!")


create_window = ThumbnailPopup