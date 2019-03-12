# -*- coding: utf-8 -*-
# https://stackoverflow.com/questions/29690868/data-binding-wpf-ironpython

import wpf
import time
from System.Windows import Application, Window
from time import localtime

class MyWindow(Window):
    someMember = None

    def __init__(self):
        self.someMember = "Hello World"
        wpf.LoadComponent(self, 'mywindow.xaml')

    @property
    def SomeMember(self):
        return self.someMember 

    @SomeMember.setter
    def SomeMember(self, value):
        self.someMember = value 


if __name__ == '__main__':
    Application().Run(MyWindow())
    time1()



