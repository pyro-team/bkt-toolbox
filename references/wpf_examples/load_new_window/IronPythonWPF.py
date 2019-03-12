# -*- coding: utf-8 -*-
# https://stackoverflow.com/questions/17504787/ironpython-wpf-load-new-window

import wpf

from System.Windows import Application, Window
from AboutWindow import *

class MyWindow(Window):
    def __init__(self):
        wpf.LoadComponent(self, 'IronPythonWPF.xaml')

    def MenuItem_Click(self, sender, e):   
        form = AboutWindow()
        form.Show()        

if __name__ == '__main__':
    Application().Run(MyWindow())
