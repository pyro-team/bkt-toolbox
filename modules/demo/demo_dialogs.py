# -*- coding: utf-8 -*-

import logging
import bkt

import System



def show_dialog():
    logging.debug("show dialog")
    from .dialogs import dialog
    dialog.show(System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle)

def show_fluentdialog():
    logging.debug("show fluent dialog")
    from .dialogs import fluentdialog
    fluentdialog.show(System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle)

def show_mahappsdialog():
    logging.debug("show Mahapps.Metro dialog")
    from .dialogs import mahapps_dialog
    mahapps_dialog.show(System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle)




bkt.powerpoint.add_tab(
    bkt.ribbon.Tab(
        label="Demo WPF dialog",
        children = [
            bkt.ribbon.Group(
                label="WPF dialog",
                children=[
                    bkt.ribbon.Button(label="WPF dialog", on_action=bkt.Callback(show_dialog)),
                    bkt.ribbon.Button(label="WPF FluentRibbon-dialog", on_action=bkt.Callback(show_fluentdialog)),
                    bkt.ribbon.Button(label="MahApps.Metro dialog", on_action=bkt.Callback(show_mahappsdialog))
                ]
            ),
        ]
    )
)


