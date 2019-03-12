# -*- coding: utf-8 -*-
#
# source: http://gui-at.blogspot.de/2009/11/inotifypropertychanged-in-ironpython.html

from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs
import pyevent


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


