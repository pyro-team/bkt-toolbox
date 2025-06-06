# -*- coding: utf-8 -*-
'''
Created on 07.07.2016

@author: rdebeerst
'''


#
# Tutorial: write functions for BKT
# =================================


#
# Create and load a module
# ------------------------
# 
# Create a python-file in /modules, e.g tutorial.py
# Link the python-file in the configuration. Add the following entry to the
# config.txt:
#       module = modules.tutorial
#


#
# Write functionality
# ------------------- 
#
# There are to restrictions for writing functions for BKT, you can us
# functions, classes, modules as if you would normally do.
# 
# The following example-function will equalize the height of shapes
def equal_height(shapes):
    height = shapes[0].height
    for shape in shapes:
        shape.height = height

# If you're not familiar with the VBA object model, you have the folowing
# options
#   * Online documentation: https://msdn.microsoft.com/en-us/library/office/jj162978.aspx
#   * Object catalogue in Visual Basic Editor, accessible from developer tab
#   * BKT-Console (see below)
#


# 
# Create the ribbon-frontend
# --------------------------
#
# To create a ribbon-frontend in office addins (without BKT), you have to
# define an XML specifying the ui-elements following the CustomUI documentation:
#    https://msdn.microsoft.com/en-us/library/dd909370(v=office.12).aspx
# 
# For the definition of a new Button on the Ribbon, we will need a Group
# in a Tab where we can but our Button in:
#    <tab label="Tutorial Tab">
#      <group label="My first group">
#        <button label="equal height" onAction="..." />
#      </group>
#    </tab>
#
# Within BKT, you won't need to write this XML-code. Instead you
# have the following options:
#   (1) add annotations to your code
#   (2) specify the ui separately
#


#
# UI definition by annotations (option 1)
# ---------------------------------------
#
# To use annotations you have to write your code in FeatureContainer-classes.
# You create one of these classes for every container-element (e.g. Tab, Group).
# The example above is achieved through:

import bkt

@bkt.group(label="My first group")
class MyFirstBKTGroup(bkt.FeatureContainer):
    
    @bkt.arg_shapes
    @bkt.button(label="equal height")
    def equal_height(self, shapes):
        height = shapes[0].height
        for shape in shapes:
            shape.height = height

@bkt.powerpoint
@bkt.tab(label="tutorial-1")
class MyFirstBKTTab(bkt.FeatureContainer):
    first_group = bkt.use(MyFirstBKTGroup)

#
# Explanations:
#   * Your features are written as methods in FeatureContainer classes
#   * @bkt.button specifies a button with label-attribute
#   * The method equal_height will be called on the click-event of this button
#   * @bkt.arg_shapes specifies that the method will get the selected shapes as parameter
#   * @bkt.group specifies a group with label-attribute
#   * Inside MyFirstBKTTab, bkt.use specifies children-elements
#   * @bkt.tab specifies a tab with label-attribute
#   * @bkt.powerpoint adds this tab to the PowerPoint configuration
#


#
# UI definition through ribbon classes (option 2)
# -----------------------------------------------
#
# The definition of the UI via ribbon classes separates UI definition from code.
# This gives you some more flexibility for the UI creation but also needs some
# more typing. Within this option, there are no restrictions on how you
# organize your bkt-functions.
#
# To specify the UI, you will simply define ribbon objects which 1-1 represent
# the XML code above.
#
# For this example, we will use the equal_height method again. The UI is then
# defined by:

my_second_tab = bkt.ribbon.Tab(
    # Attributes are given as parameters:
    label='tutorial-2',
    # Sub-elements are specified through the children-parameter
    children = [
        bkt.ribbon.Group(
            label='My first group',
            children = [
                bkt.ribbon.Button(
                    label='equal height',
                    # Now the Button gets a callback for the click-event 
                    # (called on_action) and refer to the function
                    # defined above
                    on_action=bkt.Callback(equal_height)
                )
            ]
        )
    ]
)

# Finally, we add the newly defined tab to the PowerPoint-configuration
bkt.powerpoint.add_tab(my_second_tab)


#
# There's more than just buttons
# ------------------------------
#
# For all of the CustomUI elements, there are corresponding decorators or 
# python classes. Following python notation, the python classes are in
# CamelCase:
#   * Tab, Group, Button, ToggleButton, EditBox, ...
# and the corresponding decorators in underscore_case:
#   * tab, group, button, toggle_button, edit_box, ...
#
# To get an overview, see the CustomUI documentation
#    https://msdn.microsoft.com/en-us/library/dd909370(v=office.12).aspx
# 
# Parameters added to the python classes (or decorators) become XML-Attributes 
# for the ribbon UI. Hereby, the pythonic underscore notation is transformed
# into camelCase, e.g.
button = bkt.ribbon.Button(label='my first button', image_mso='ShapeHeight')
# defines the following XML:
#   <button label='my first button' imageMso='ShapeHeight'>
#
# WARNING: Typos within the parameters are not evaluated or corrected. The will 
# lead to an invalid XML and none of the BKT-Frontends will show up in the
# ribbon.
#
#
# To see all the possible CustomUI elements and attributes in action, see the
# examples in the demo-Module demo/demo_customui.py
 

#
# Callbacks with parameters
# -------------------------
#
# Some of the ribbon-callbacks have predefined parameters, e.g. the on_changed
# callback has a value-attribute giving a string of the input that was typed
# into the textbox. Functions, to be used in those callbacks, have to have
# these parameter names:

def on_change_function(value):
    bkt.message('edit-box changed: new text=%s' % value)

textbox = bkt.ribbon.EditBox(
    label='my first editbox',
    size_string='xxxxxxxx',
    on_change=bkt.Callback(on_change_function)
)

# Other callbacks with parameters are: on_toggle_action (pressed), 
# on_action_indexed (selected_item, index), get_item_*** (index)
# For examples of all callbacks see demo/demo_customui.py


#
# Callback context parameters
# ---------------------------
#
# In order to manipulate an office document within your function,
# you will need access to the VBA object model. 
#
# Context-parameters are passed to the callback-functions by BKT
# using one of the following options:
# 
# (1) Using annotations
# If you use annotations for the UI-definiton (see above), you simply
# add annotations for the arguments your function want to retreive.
# E.g. with @bkt.arg_shapes your function will get the selected shapes
# as list in the 'shapes'-parameter.
#
#
# (2) Configuring parameters directly
# If you specify the frontend using ribbon-classes, you configure the 
# context-parameters by passing more arguments to the Callback-constructur:
on_action_callback = bkt.Callback(equal_height, shapes=True)
# This callback will also be called passing the selected shapes as argument
#
#
# (2a) Using variable convention
# You can use also use the automatic option, where variable-names in the 
# parameters of your function are recocnised.
# This is done in the definition of my_second_tab used above:
callback = bkt.Callback(equal_height)
# The shapes-parameter is automatically found and configured acordingly.
#
#
# The following parameters/flags are available:
#   * application: the complete office-application context
#   * context: a context object, including the application, addin, config, current control, etc.
# and for PowerPoint
#   * slides: list of selected slides
#   * shapes: list of selected shapes ()
#   * slide: the seletected slide, min=max=1
#   * shape: the seletected shape, min=max=1
#   * require_text (flag): ensures that the callback is only called if the shape has a text frame
# The corresponding decorators are: @bkt.arg_context, @bkt.arg_slides, etc.
#


#
# Special BKT-Controls
# --------------------
#
# BKT defines some special frontend-elements, not defined by the default 
# CustomUI, e.g.
#   * SpinnerBox: horizontal group of EditBox, Decrement-Button, Increment-Button
#   * RoundingSpinnerBox: SpinnerBox with functionality for increment/decrement
#   * ColorGallery: Gallery-UI with color buttons, analogous to the built-in PowerPoint color gallery
# These special controls use the default CustomUI-Elements under the hood.
#
# Examples for their usage and configurations can be found here: 
#    demo/demo_bkt.py
#


#
# Context menus and contextual tabs
# ---------------------------------
#
# Besides ribbon-Tabs, CustomUI also allows the definition of context menus,
# contextual tabs and repurposed commands.
# These elements can also be defined in BKT by using
#   * bkt.powerpoint.add_context_menu(menu-mso-id, context-menu-element)
#   * bkt.powerpoint.add_contextual_tab(mso-id, tab-element)
#   * bkt.powerpoint.add_command(command)
#
# Examples for context-menu and contextual-tab can be found in the toolbox, 
# e.g. the Harvey-Moon implementation: modules/toolbox/harbey.py


#
# BKT console
# -----------
# 
# BKT offers a console to manipulate office objects directly and explore the
# VBA object model. You can access the console via the developer tab.
#
# Usage:
#   * Write python code in the input-box
#   * Execute the code by pressend ctrl-enter
#   * Access the office object model via: context.app.###
#   * Explore objects using tab completion
#


