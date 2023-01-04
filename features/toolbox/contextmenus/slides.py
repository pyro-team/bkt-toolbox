# -*- coding: utf-8 -*-
'''
Created on 29.04.2021

@author: fstallmann
'''



import bkt

from .. import slides


class ContextSlides(object):
    @staticmethod
    def get_children():
        return bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None, 
                children=[
            bkt.ribbon.Button(
                label='Speichern',
                image_mso='SaveSelectionToTextBoxGallery',
                supertip="Speichert die ausgewählten Folien in einer neuen Präsentation.",
                on_action=bkt.Callback(slides.SlideMenu.save_slides_dialog)
            ),
            bkt.ribbon.Button(
                label='Senden',
                image_mso='FileSendAsAttachment',
                supertip="Sendet die ausgewählten Folien als E-Mail Anhang.",
                on_action=bkt.Callback(slides.SlideMenu.send_slides_dialog)
            ),
                    ]
            )