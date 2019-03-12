
# http://gui-at.blogspot.de/2009/11/inotifypropertychanged-in-ironpython.html

import clr
import System
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')

from System.Windows.Markup import XamlReader
from System.Windows import Application, Window
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs
import pyevent

XAML_str = """<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Width="250" Height="62">
    <StackPanel x:Name="DataPanel" Orientation="Horizontal">
        <Label Content="Size"/>
        <Label Content="{Binding size}"/>
        <TextBox x:Name="tbSize" Text="{Binding size, UpdateSourceTrigger=PropertyChanged}" />
        <Button x:Name="Button" Content="Set Initial Value"></Button>
    </StackPanel>
</Window>"""


class notify_property(property):

    def __init__(self, getter):
        def newgetter(slf):
            #return None when the property does not exist yet
            try:
                return getter(slf)
            except AttributeError:
                return None
        super(notify_property, self).__init__(newgetter)

    def setter(self, setter):
        def newsetter(slf, newvalue):
            # do not change value if the new value is the same
            # trigger PropertyChanged event when value changes
            oldvalue = self.fget(slf)
            if oldvalue != newvalue:
                setter(slf, newvalue)
                slf.OnPropertyChanged(setter.__name__)
        return property(
            fget=self.fget,
            fset=newsetter,
            fdel=self.fdel,
            doc=self.__doc__)


class NotifyPropertyChangedBase(INotifyPropertyChanged):
    PropertyChanged = None

    def __init__(self):
        self.PropertyChanged, self._propertyChangedCaller = pyevent.make_event()

    def add_PropertyChanged(self, value):
        self.PropertyChanged += value

    def remove_PropertyChanged(self, value):
        self.PropertyChanged -= value

    def OnPropertyChanged(self, propertyName):
        if self.PropertyChanged is not None:
            self._propertyChangedCaller(self, PropertyChangedEventArgs(propertyName))
    
    def declare_notifiable(self, *symbols):
        for symbol in symbols:
            self.define_notifiable_property(symbol)

#     def define_notifiable_property(self, symbol):
#         dnp = """
# import sys
# sys.path.append(__file__)
# @notify_property
# def {0}(self):
#     return self._{0}
#
# @{0}.setter
# def {0}(self, value):
#     self._{0} = value
# """.format(symbol)
#         d = globals()
#         exec dnp.strip() in d
#         setattr(self.__class__, symbol, d[symbol])

    

class ViewModel(NotifyPropertyChangedBase):
    
    def __init__(self):
        super(ViewModel, self).__init__()
        # must be string to two-way binding work correctly
        #self.define_notifiable_property("size")
        self.size = '10'
        
        
    @notify_property
    def size(self):
        return self._size

    @size.setter
    def size(self, value):
        self._size = value
        print 'Size changed to %r' % self.size


class TestWPF(object):

    def __init__(self):
        self._vm = ViewModel()
        self.root = XamlReader.Parse(XAML_str)
        self.DataPanel.DataContext = self._vm
        self.Button.Click += self.OnClick
        
    def OnClick(self, sender, event):
        # must be string to two-way binding work correctly
        self._vm.size = '10'

    def __getattr__(self, name):
        # provides easy access to XAML elements (e.g. self.Button)
        return self.root.FindName(name)


if __name__ == '__main__':
    tw = TestWPF()
    app = Application()
    app.Run(tw.root)

