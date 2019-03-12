# -*- coding: utf-8 -*-
# http://www.ironpython.info/index.php?title=WPF_Example

# Reference the WPF assemblies
import clr
clr.AddReferenceByName("PresentationFramework, Version=3.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35")
clr.AddReferenceByName("PresentationCore, Version=3.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35")
import System.Windows

# Initialization Constants
Window = System.Windows.Window
Application = System.Windows.Application
Button = System.Windows.Controls.Button
StackPanel = System.Windows.Controls.StackPanel
Label = System.Windows.Controls.Label
Thickness = System.Windows.Thickness
DropShadowBitmapEffect = System.Windows.Media.Effects.DropShadowBitmapEffect


# Create window
my_window = Window()
my_window.Title = 'Welcome to IronPython'

# Create StackPanel to Layout UI elements 
my_stack = StackPanel()
my_stack.Margin = Thickness(15)
my_window.Content = my_stack

# Create Button and add a Button Click event handler
my_button = Button()
my_button.Content = 'Push Me'
my_button.FontSize = 24
my_button.BitmapEffect = DropShadowBitmapEffect()

def clicker(sender, args):

   # Create new label
   my_message = Label()
   my_message.FontSize = 48
   my_message.Content = 'Welcome to IronPython!'

   # Add label into stack panel of controls
   my_stack.Children.Add (my_message)

my_button.Click += clicker

my_stack.Children.Add (my_button)

# Run application
my_app = Application()
my_app.Run (my_window)