# -*- coding: utf-8 -*-

# wpf basics
import wpf
from System.Windows import Application, Window

# property binding
from notify import NotifyPropertyChangedBase, notify_property

# for generate_image
from System.Windows.Media.Imaging import BitmapImage, BitmapSource
from System.Windows.Media import PixelFormats
import System

# for MahApps.Metro
import clr
clr.AddReferenceToFileAndPath('MahApps.Metro')
from MahApps.Metro.Controls import MetroWindow
# for FluentRibbon
clr.AddReferenceToFileAndPath('Fluent')


# for generate_rect_image
from System.Windows.Shapes import Rectangle
from System.Windows.Media import SolidColorBrush, ColorConverter

class ViewModel(NotifyPropertyChangedBase):

    def __init__(self):
        super(ViewModel, self).__init__()
        # must be string to two-way binding work correctly
        #self.define_notifiable_property("size")
        self.size = '10'
        self.fill_color = '#0066cc'
        self._red = 90
        self._green = 55
        self._blue = 200
        
  


class TestWindow(MetroWindow):
    
    def __init__(self):
        wpf.LoadComponent(self, 'MetroWindow.xaml')
        self._vm = ViewModel()
        self.DataPanel.DataContext = self._vm

    def __getattr__(self, name):
        # provides easy access to XAML elements (e.g. self.Button)
        return self.root.FindName(name)
    
    



if __name__ == '__main__':
    app = Application()
    print "create window"
    wnd = TestWindow()
    print "run window"
    app.Run(wnd)



