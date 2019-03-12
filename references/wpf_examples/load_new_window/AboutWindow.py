import wpf

from System.Windows import Window

class AboutWindow(Window):
    def __init__(selfAbout):        
        wpf.LoadComponent(selfAbout, 'AboutWindow.xaml')