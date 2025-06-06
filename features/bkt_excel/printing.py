# -*- coding: utf-8 -*-
'''
Created on 2017-07-18
@author: Florian Stallmann
'''



import bkt


# class Printing(object):
    # @staticmethod
    # def fit_to_page(application, selected_sheets):
    #     application.PrintCommunication = False
    #     for sheet in selected_sheets:
    #         sheet.PageSetup.PrintArea = ""
    #         sheet.PageSetup.Zoom = False
    #         sheet.PageSetup.FitToPagesWide = 1
    #         sheet.PageSetup.FitToPagesTall = 1
    #     application.PrintCommunication = True

    # @staticmethod
    # def change_orientation(application, selected_sheets, orientation): #1=xlPortrait 2=xlLandscape
    #     application.PrintCommunication = False
    #     for sheet in selected_sheets:
    #         sheet.PageSetup.Orientation = orientation
    #     application.PrintCommunication = True

    # @staticmethod
    # def add_header_footer(application, selected_sheets):
    #     application.PrintCommunication = False
    #     for sheet in selected_sheets:
    #         sheet.PageSetup.OddAndEvenPagesHeaderFooter = False
    #         sheet.PageSetup.DifferentFirstPageHeaderFooter = False
    #         sheet.PageSetup.LeftHeader = '&N'
    #         sheet.PageSetup.CenterHeader = ''
    #         sheet.PageSetup.RightHeader = '&B'
    #         sheet.PageSetup.LeftFooter = ''
    #         sheet.PageSetup.CenterFooter = 'Seite &S von &A'
    #         sheet.PageSetup.RightFooter = ''
    #     application.PrintCommunication = True


drucken_gruppe = bkt.ribbon.Group(
    label="Drucken",
    image_mso="PrintAreaMenu",
    children=[
        bkt.mso.control.PageOrientationGallery(show_label=True),
        bkt.mso.control.PageScaleToFitWidth(show_label=True),
        bkt.mso.control.PageScaleToFitHeight(show_label=True),
        # bkt.ribbon.Button(
        #     id = 'fit_to_page',
        #     label="Blätter auf je 1 Seite skalieren",
        #     show_label=True,
        #     image_mso='PrintSetupDialog',
        #     supertip="Skaliert alle ausgewählten Blätter so, dass jedes auf genau einer Seite gedruckt wird. Die Seitengröße und -ausrichtung wird nicht geändert.",
        #     on_action=bkt.Callback(Printing.fit_to_page, application=True, selected_sheets=True),
        #     get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
        # ),
        # bkt.ribbon.Button(
        #     id = 'add_header_footer',
        #     label="Kopf-/Fußzeilen einfügen",
        #     show_label=True,
        #     image_mso='HeaderFooterInsert',
        #     supertip="Fügt Standard Kopfzeile (Dateiname und Tabellenblattname) und Fußzeile (Seite x von y) für alle ausgewählten Blätter ein.",
        #     on_action=bkt.Callback(Printing.add_header_footer, application=True, selected_sheets=True),
        #     get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
        # ),
        bkt.ribbon.DialogBoxLauncher(idMso='PageSetupPageDialog')
    ]
)