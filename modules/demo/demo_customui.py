import bkt


# define some default callbacks for the demo-tab
action_callback = bkt.Callback(
    lambda current_control: bkt.message('current_control clicked: label=%s,\nid=%s' % (current_control['label'], current_control['id'])),
    current_control=True)

toggle_callback = bkt.Callback(
    lambda pressed, current_control: bkt.message('toggle-button clicked: label=%s,\nid=%s\n\npressed-state after click=%s' % (current_control['label'], current_control['id'], pressed)),
    current_control=True)

change_callback = bkt.Callback(
    lambda value, current_control: bkt.message('edit-box changed: label=%s,\nid=%s\n\nnew text=%s\n(text will be reset by get_text-Callback)' % (current_control['label'], current_control['id'], value)),
    current_control=True)

get_text_callback = bkt.Callback(lambda : 'default text')

get_content_callback = bkt.callbacks.Callback(
    lambda : '<menu xmlns="http://schemas.microsoft.com/office/2006/01/customui"><button id="button1" label="Button 1" /><button id="button2" label="Button 2" /><button id="button3" label="Button 3" /></menu>')

action_indexed_callback = bkt.Callback(
    lambda selected_item, index, current_control: bkt.message('current_control clicked: label=%s,\nid=%s\n\nindex=%s, selected_item_id=%s' % (current_control['label'], current_control['id'], index, selected_item)),
    current_control=True)



bkt.powerpoint.add_tab(
    bkt.ribbon.Tab(
        label="Demo MS-CustomUI",
        children = [
            bkt.ribbon.Group(
                label="MSOffice Controls",
                children=[
                    bkt.mso.control.Paste(size="large"),
                    bkt.mso.control.Copy,
                    #mso-control-clone
                    #mso-button
                    #mso-menu
                ]
            ),
            bkt.ribbon.Group(
                label="Buttons",
                
                 # group-image is shown if group collapses (on small window-widths)
                image_mso="Bold",
                
                children=[
                    bkt.ribbon.Button(label="large button", size="large", on_action=action_callback),
                    bkt.ribbon.Button(label="normal button", size="normal", on_action=action_callback),
                    bkt.ribbon.ButtonGroup(
                        children=[
                            # children must be unsized
                            bkt.ribbon.Button(label="button group 1", on_action=action_callback),
                            bkt.ribbon.Button(label="2", on_action=action_callback),
                            bkt.ribbon.Button(label="3", on_action=action_callback)
                        ]
                    ),
                    bkt.ribbon.Button(label="button w keytip", keytip="K", on_action=action_callback, 
                        supertip="press alt to access commands via keystrokes. on this tab, this button can be accessed with 'K'"),
                    bkt.ribbon.ToggleButton(label="toggle button", on_toggle_action=toggle_callback),
                    bkt.ribbon.SplitButton(
                        children=[
                            # button/toggle-button and menu
                            bkt.ribbon.Button(label="split-button", on_action=action_callback),
                            bkt.ribbon.Menu(children=[
                                # buttons/controls inside menus must not have 'size'-attribute
                                bkt.ribbon.Button(label="button 1", on_action=action_callback),
                                bkt.ribbon.Button(label="button 2", on_action=action_callback)
                            ])
                        ]
                    ),
                    bkt.ribbon.SplitButton(
                        size="large",
                        children=[
                            bkt.ribbon.ToggleButton(label="toggle in large split-button", on_toggle_action=toggle_callback),
                            bkt.ribbon.Menu(children=[
                                bkt.ribbon.Button(label="button 1", on_action=action_callback),
                                bkt.ribbon.Button(label="button 2", on_action=action_callback)
                            ])
                        ]
                    )
                ]
            ),
            bkt.ribbon.Group(
                label="Text",
                children=[
                    bkt.ribbon.LabelControl(label="Label"),
                    bkt.ribbon.Button(label="pause pointer on this button", screentip="the button's screentip", supertip="the button's supertip", on_action=action_callback),
                    bkt.ribbon.Box(boxStyle="horizontal",
                        children=[
                            bkt.ribbon.LabelControl(label="button w/o label:"),
                            bkt.ribbon.Button(label="this label is only visible in tooltip", show_label=False, image_mso="HappyFace", on_action=action_callback)
                        ]
                    )
                ]
            ),
            bkt.ribbon.Group(
                label="Edit-Controls",
                children=[
                    bkt.ribbon.EditBox(label="Edit Box", on_change=change_callback, get_text=get_text_callback),
                    bkt.ribbon.CheckBox(label="Check box", on_toggle_action=toggle_callback),
                    bkt.ribbon.LabelControl(label=" "),  # this empty label forces a column break
                    bkt.ribbon.ComboBox(label="Combo Box", on_change=change_callback, get_text=get_text_callback, maxLength=10, children=[
                        bkt.ribbon.Item(label="Item 1"),
                        bkt.ribbon.Item(label="Item 2"),
                        bkt.ribbon.Item(label="Item 3")
                    ] ), 
                    bkt.ribbon.DropDown(label="Drop-down", on_action_indexed=action_indexed_callback, children=[
                        bkt.ribbon.Item(label="Item 1"),
                        bkt.ribbon.Item(label="Item 2"),
                        bkt.ribbon.Item(label="Item 3"),
                        bkt.ribbon.Button(label="button 1", on_action=action_callback),
                        bkt.ribbon.Button(label="button 2", on_action=action_callback)
                    ] )
                ]
            ),
            bkt.ribbon.Group(
                label="Menus",
                children=[
                    bkt.ribbon.Menu(
                        label="Menu",
                        children=[
                            bkt.ribbon.Button(label="button", on_action=action_callback),
                            bkt.ribbon.ToggleButton(label="toggle button"),
                            bkt.ribbon.CheckBox(label="check box")
                        ]
                    ),
                    bkt.ribbon.Menu(
                        label="seperators",
                        children=[
                            bkt.ribbon.MenuSeparator(title="separator as menu title"),
                            bkt.ribbon.Button(label="button 1", on_action=action_callback),
                            bkt.ribbon.Button(label="button 2", on_action=action_callback),
                            bkt.ribbon.MenuSeparator(),
                            bkt.ribbon.Button(label="button 3 (second button group)", image_mso="HappyFace", on_action=action_callback),
                            bkt.ribbon.Button(label="button 4", on_action=action_callback),
                            bkt.ribbon.MenuSeparator(title="seperator with title"),
                            bkt.ribbon.Button(label="button 5 (third button group)", on_action=action_callback),
                            bkt.ribbon.Button(label="button 6", on_action=action_callback)
                        ]
                    ),
                    bkt.ribbon.DynamicMenu(
                        label="dynamic", 
                        get_content=get_content_callback
                    ),
                    bkt.ribbon.Menu(
                        label="large buttons",
                        item_size="large",
                        children=[
                            bkt.ribbon.Button(label="button", on_action=action_callback),
                            bkt.ribbon.ToggleButton(label="toggle button", on_toggle_action=toggle_callback),
                            bkt.ribbon.CheckBox(label="check box", on_toggle_action=toggle_callback)
                        ]
                    ),
                    bkt.ribbon.Gallery(
                        label="Gallery",
                        on_action_indexed=action_indexed_callback, 
                        children=[
                            bkt.ribbon.Item(label="item 1"),
                            bkt.ribbon.Item(label="item 2"),
                            bkt.ribbon.Item(label="item 3"),
                            bkt.ribbon.Item(label="item 4"),
                            bkt.ribbon.Item(label="item 5")
                        ]
                    )
                ]
            ),
            bkt.ribbon.Group(
                label="Dialog Launcher",
                children=[
                    bkt.ribbon.DialogBoxLauncher(on_action=action_callback, label="Dialog-Box Button")
                ]
            ),
            bkt.ribbon.Group(
                label="Boxes",
                children=[
                    bkt.ribbon.Box(boxStyle="horizontal",
                        children=[
                            bkt.ribbon.Box(boxStyle="vertical",
                                children=[
                                    bkt.ribbon.Button(label="b1", on_action=action_callback),
                                    bkt.ribbon.Button(label="b2", on_action=action_callback)
                                ]
                            ),
                            bkt.ribbon.Button(label="b3", on_action=action_callback),
                            bkt.ribbon.Box(boxStyle="vertical",
                                children=[
                                    bkt.ribbon.Button(label="b4", on_action=action_callback),
                                    bkt.ribbon.Button(label="b5", on_action=action_callback)
                                ]
                            )
                        ]
                    )
                ]
            ),
            # bkt.ribbon.Group(
            #     label="Separator",
            #     children=[
            #         bkt.ribbon.Button(label="Button 1", size="large", imageMso="HappyFace", on_action=action_callback),
            #         # Separator not working in PowerPoint 2010
            #         bkt.ribbon.Separator(),
            #         bkt.ribbon.Button(label="Button 2", size="large", imageMso="HappyFace", on_action=action_callback)
            #     ]
            # ),
            
            
            bkt.ribbon.Group(
                label="Demo of all callbacks",
                children=[
                    bkt.ribbon.Button(
                        image_mso="HappyFace",
                        
                        # get_description
                        #   applies to: button checkBox dynamicMenu gallery menu toggleButton
                        get_description = bkt.Callback(lambda: "my description"),
                        
                        # get_enabled
                        #   applies to: button checkBox comboBox dropDown dynamicMenu editBox gallery labelControl menu splitButton toggleButton
                        get_enabled     = bkt.Callback(lambda: True),
                        
                        # get_image
                        #   must return image of type System.Drawing.Bitmap
                        #   applies to: group button comboBox dropDown dynamicMenu editBox gallery menu toggleButton
                        #   mutually exclusive: get_image, image, image_mso
                        #get_image       = bkt.Callback(lambda: None),
                        
                        # get_keytip
                        #   applies to: tab group button checkBox comboBox dropDown dynamicMenu editBox gallery menu splitButton toggleButton
                        get_keytip      = bkt.Callback(lambda: 'T'),
                        
                        # get_label
                        #   applies to: tab group button checkBox comboBox dropDown dynamicMenu editBox gallery labelControl menu splitButton toggleButton
                        get_label       = bkt.Callback(lambda: 'my label'),
                        
                        # get_screentip
                        #   applies to: group button checkBox comboBox dropDown dynamicMenu editBox gallery labelControl menu toggleButton
                        get_screentip   = bkt.Callback(lambda: "button's screentip"),
                        
                        # get_show_image
                        #   applies to: button comboBox dropDown dynamicMenu editBox gallery menu toggleButton
                        get_show_image  = bkt.Callback(lambda: True),
                        
                        # get_show_label
                        #   automatically true, if size is large
                        #   applies to: button comboBox dropDown dynamicMenu editBox gallery labelControl menu splitButton toggleButton
                        get_show_label  = bkt.Callback(lambda: False),

                        # get_size
                        #   allowed return values: normal, large
                        #   applies to: button dynamicMenu gallery menu splitButton toggleButton
                        get_size        = bkt.Callback(lambda: "normal"),
                        
                        # get_supertip
                        #   applies to: group button checkBox comboBox dropDown dynamicMenu editBox gallery labelControl menu separator splitButton toggleButton
                        get_supertip    = bkt.Callback(lambda: "button's supertip"),
                        
                        # get_visible
                        #   applies to: tab group box button buttonGroup checkBox comboBox dropDown dynamicMenu editBox gallery labelControl menu separator splitButton toggleButton
                        get_visible     = bkt.Callback(lambda: True),
                        
                        # on_action
                        #   applies to: button
                        on_action       = bkt.Callback(lambda: bkt.message('message'))
                    ),
                    
                    
                    bkt.ribbon.ToggleButton(
                        label="toggle button",
                        supertip="this toggle button will always be pressed",
                        
                        # get_pressed
                        #   applies to: checkBox toggleButton
                        get_pressed     = bkt.Callback(lambda: True),
                        
                        # on_action
                        #   applies to: checkBox toggleButton
                        on_toggle_action = bkt.Callback(
                            lambda pressed: bkt.message('toggle-button clicked, pressed-state=%s' % (pressed)),
                            bkt.CallbackTypes.on_toggle_action),
                        
                        # via customui.LoadImage, the image is automatically loaded from /resources/images
                        image = "Test32"
                        
                    ),
                    
                    
                    bkt.ribbon.DynamicMenu(
                        label = "dynamic menu",
                        
                        # get_content
                        #   applies to: dynamicMenu
                        get_content = bkt.callbacks.Callback(
                            lambda : '<menu xmlns="http://schemas.microsoft.com/office/2006/01/customui"><button id="button1" label="Button 1" /><button id="button2" label="Button 2" /><button id="button3" label="Button 3" /></menu>',
                            bkt.callbacks.CallbackTypes.get_content)
                    ),
                    
                    
                    bkt.ribbon.Gallery(
                        label="gallery",
                        rows=3,
                        
                        # get_item_count
                        #   applies to: comboBox dropDown gallery
                        get_item_count          = bkt.callbacks.Callback(lambda: 10),
                        
                        # get_selected_item_id
                        #   applies to: dropDown gallery
                        #   mutually exclusive: get_selected_item_id, get_selected_item_index
                        get_selected_item_id    = bkt.callbacks.Callback(lambda: 'item-id-6'),
                        
                        # get_selected_item_index
                        #   applies to: dropDown gallery
                        #   mutually exclusive: get_selected_item_id, get_selected_item_index
                        #get_selected_item_index = bkt.callbacks.Callback(lambda: 5),
                        
                        
                        # get_item_id
                        #   applies to: comboBox dropDown gallery
                        get_item_id             = bkt.callbacks.Callback(lambda index: 'item-id-%s' % (index)),
                        
                        # get_item_image
                        #   must return image of type System.Drawing.Bitmap
                        #   applies to: comboBox dropDown gallery
                        get_item_image          = bkt.callbacks.Callback(lambda index: None),
                        
                        # get_item_label
                        #   applies to: comboBox dropDown gallery
                        get_item_label          = bkt.callbacks.Callback(lambda index: 'item %s' % (index)),
                        
                        # get_item_screentip
                        #   applies to: comboBox dropDown gallery
                        get_item_screentip      = bkt.callbacks.Callback(lambda index: 'item %s\'s screentip'),
                        
                        # get_item_supertip
                        #   applies to: comboBox dropDown gallery
                        get_item_supertip       = bkt.callbacks.Callback(lambda index: 'item %s\'s supertip'),
                        
                        
                        # get_item_height
                        #   applies to: gallery
                        get_item_height         = bkt.callbacks.Callback(lambda: 32),
                        
                        # get_item_width
                        #   applies to: gallery
                        get_item_width          = bkt.callbacks.Callback(lambda: 32),
                        
                        # on_action_indexed
                        #   applies to: dropDown gallery
                        on_action_indexed       = bkt.Callback(lambda selected_item, index: bkt.message('gallery item %s clicked' % (index)))
                    ),
                    
                    
                    bkt.ribbon.EditBox(
                        label="edit box",
                        
                        # get_text
                        #   applies to: comboBox editBox
                        get_text = bkt.Callback(lambda: 'default text'),
                        
                        # on_change
                        #   applies to: comboBox editBox
                        on_change = bkt.Callback(
                            lambda value: bkt.message('text changed, value=%s' % (value)),
                            bkt.CallbackTypes.on_change)
                    ),
                    
                    
                    bkt.ribbon.Menu(
                        label="menu",
                        children= [
                            bkt.ribbon.Button(label="a button"),
                            bkt.ribbon.MenuSeparator(
                                
                                # get_title
                                #   applies to: menuSeparator
                                get_title = bkt.Callback(lambda: 'separator title'),
                            ),
                            bkt.ribbon.Button(label="another button")
                        ]
                    )
                    
                    # the callbacks 'loadImage' and 'onLoad' are only on customUI
                ]
            ),
        ]
        
        
        
        
        # TODO: quick-access-toolbar
        # <ribbon><qat><documentControls><button>...</button></documentControls></qat></ribbon>
        # <officeMenu>
        # <sharedControls>
        
        # TODO: repurposed commands
        # <customUI><commands><command>...</command></commands></customUI>
        
        # TODO: contextual tabs
        # <ribbon><contextualTabs><tabSet><tab>...</tab></tabSet></contextualTabs></ribbon>
        
        # TODO: Menu with title / splitButton with title
        # <menu> 
        
    )
        
)

