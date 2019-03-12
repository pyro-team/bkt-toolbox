import bkt

# define some default callbacks for the demo-tab
action_callback = bkt.Callback(
    lambda current_control: bkt.helpers.message('current_control clicked: label=%s,\nid=%s' % (current_control['label'], current_control['id'])),
    current_control=True)



backstage_control1 = bkt.ribbon.Tab(
        label="BKT Demo 1",
        title="BKT Demo 1 of backstage area",
        columnWidthPercent="30",
        insertAfterMso="TabInfo",
        children=[
            bkt.ribbon.FirstColumn(children=[
                bkt.ribbon.Group(label="Test Group 1", children=[
                    bkt.ribbon.PrimaryItem(children=[
                        bkt.ribbon.Menu(
                            label="Primary Item Menu",
                            image_mso="HappyFace",
                            children=[
                                bkt.ribbon.MenuGroup(
                                    label="Menu group large",
                                    item_size="large",
                                    children=[
                                        bkt.ribbon.Button(
                                            label="Test",
                                            description="Lorem ipsum",
                                            image_mso="HappyFace",
                                            on_action=action_callback,
                                        ),
                                        bkt.ribbon.Button(
                                            label="Test",
                                            description="Lorem ipsum",
                                            image_mso="HappyFace",
                                            on_action=action_callback,
                                        )
                                    ]
                                ),
                                bkt.ribbon.MenuGroup(
                                    label="Menu group small",
                                    children=[
                                        bkt.ribbon.Button(
                                            label="Test",
                                            image_mso="HappyFace",
                                            on_action=action_callback,
                                        ),
                                        bkt.ribbon.Button(
                                            label="Test",
                                            image_mso="HappyFace",
                                            on_action=action_callback,
                                        )
                                    ]
                                )
                            ]
                        )
                    ]),
                    bkt.ribbon.TopItems(children=[
                        bkt.ribbon.Label(label="Label 1 Lorem Ipsum"),
                        bkt.ribbon.LayoutContainer(layoutChildren="horizontal", children=[
                            bkt.ribbon.ImageControl(image_mso="ColorGray"),
                            bkt.ribbon.Label(label="Label 2 with bullet Lorem Ipsum"),
                        ]),
                        bkt.ribbon.LayoutContainer(layoutChildren="horizontal", children=[
                            bkt.ribbon.ImageControl(image_mso="OutlineDemote"),
                            bkt.ribbon.Label(label="Label 3 with bullet Lorem Ipsum"),
                        ]),
                        bkt.ribbon.LayoutContainer(layoutChildren="horizontal", children=[
                            bkt.ribbon.ImageControl(image_mso="NextArrow"),
                            bkt.ribbon.Label(label="Label 3 with bullet Lorem Ipsum"),
                        ]),
                    ]),
                ]),
                bkt.ribbon.Group(label="Test Group 2", children=[
                    bkt.ribbon.PrimaryItem(children=[
                        bkt.ribbon.Button(
                            label="Primary button and Close",
                            image_mso="HappyFace",
                            on_action=action_callback,
                            is_definitive=True,
                        ),
                    ]),
                    bkt.ribbon.TopItems(children=[
                        bkt.ribbon.Button(
                            label="Top item button",
                            image_mso="HappyFace",
                            on_action=action_callback,
                        ),
                        bkt.ribbon.Button(
                            label="Top item button",
                            image_mso="HappyFace",
                            on_action=action_callback,
                        ),
                        bkt.ribbon.Button(
                            label="Top item button",
                            image_mso="HappyFace",
                            on_action=action_callback,
                        ),
                        bkt.ribbon.Button(
                            label="Top item button",
                            image_mso="HappyFace",
                            on_action=action_callback,
                        ),
                    ]),
                    bkt.ribbon.BottomItems(children=[
                        bkt.ribbon.Button(
                            label="Bottom item button",
                            image_mso="HappyFace",
                            on_action=action_callback,
                        ),
                        bkt.ribbon.Button(
                            label="Bottom item button",
                            image_mso="HappyFace",
                            on_action=action_callback,
                        ),
                        bkt.ribbon.Button(
                            label="Bottom item button",
                            image_mso="HappyFace",
                            on_action=action_callback,
                        ),
                    ])
                ]),
            ]),
            bkt.ribbon.SecondColumn(children=[
                bkt.ribbon.Group(label="Test Group 3 in second column", children=[
                    bkt.ribbon.TopItems(children=[
                        bkt.ribbon.Button(
                            label="Button with is_definitive to close backstage",
                            image_mso="HappyFace",
                            on_action=action_callback,
                            is_definitive=True,
                        ),
                        bkt.ribbon.EditBox(
                            label="Editbox",
                        ),
                        bkt.ribbon.ComboBox(
                            label="Combobox",
                            children=[
                                bkt.ribbon.Item(label="One"),
                                bkt.ribbon.Item(label="Two"),
                                bkt.ribbon.Item(label="Three"),
                            ]
                        ),
                        bkt.ribbon.GroupBox(
                            label="Group Box",
                            children=[
                                bkt.ribbon.CheckBox(label="Check Box 1"),
                                bkt.ribbon.CheckBox(label="Check Box 2"),
                                bkt.ribbon.RadioGroup(label="Radio Group", children=[
                                    bkt.ribbon.RadioButton(label="Radio Button 1"),
                                    bkt.ribbon.RadioButton(label="Radio Button 2"),
                                ]),
                            ]
                        ),
                        bkt.ribbon.LayoutContainer(
                            layoutChildren="horizontal",
                            children=[
                                bkt.ribbon.Button(
                                    label="Button left",
                                    image_mso="HappyFace",
                                    on_action=action_callback,
                                ),
                                bkt.ribbon.Button(
                                    label="Bottom middle",
                                    image_mso="HappyFace",
                                    on_action=action_callback,
                                ),
                                bkt.ribbon.Button(
                                    label="Bottom right",
                                    image_mso="HappyFace",
                                    on_action=action_callback,
                                ),
                            ]
                        )
                    ])
                ]),
            ])
        ]
    )


backstage_control2 = bkt.ribbon.Tab(
        label="BKT Demo 2",
        title="BKT Demo 2 of backstage area",
        insertAfterMso="TabInfo",
        firstColumnMinWidth="500",
        firstColumnMaxWidth="500",
        children=[
            bkt.ribbon.FirstColumn(children=[
                bkt.ribbon.TaskFormGroup(label="Task Form Group Label", children=[
                    bkt.ribbon.Category(label="Category Label 1", children=[

                        bkt.ribbon.Task(label="Task 1", description="Lorem ipsum", image_mso="HappyFace", children=[
                            bkt.ribbon.Group(label="Test Group 1", children=[
                                bkt.ribbon.TopItems(children=[
                                    bkt.ribbon.Button(
                                        label="Test1",
                                        image_mso="HappyFace",
                                        on_action=action_callback,
                                    ),
                                    bkt.ribbon.Button(
                                        label="Test2",
                                        image_mso="HappyFace",
                                        on_action=action_callback,
                                    )
                                ]),
                            ])
                        ]),
                        bkt.ribbon.Task(label="Task 2", description="Lorem ipsum", image_mso="HappyFace", children=[
                            bkt.ribbon.Group(label="Test Group 2", children=[
                                bkt.ribbon.TopItems(children=[
                                    bkt.ribbon.Button(
                                        label="Test3",
                                        image_mso="HappyFace",
                                        on_action=action_callback,
                                    ),
                                    bkt.ribbon.Button(
                                        label="Test4",
                                        image_mso="HappyFace",
                                        on_action=action_callback,
                                    )
                                ]),
                            ])
                        ]),

                    ]),

                    bkt.ribbon.Category(label="Category Label 2", children=[

                        bkt.ribbon.Task(label="Task 3", description="Lorem ipsum", image_mso="HappyFace", children=[
                            bkt.ribbon.Group(label="Test Group 3", children=[
                                bkt.ribbon.TopItems(children=[
                                    bkt.ribbon.Button(
                                        label="Test5",
                                        image_mso="HappyFace",
                                        on_action=action_callback,
                                    ),
                                    bkt.ribbon.Button(
                                        label="Test6",
                                        image_mso="HappyFace",
                                        on_action=action_callback,
                                    )
                                ]),
                            ])
                        ]),
                        bkt.ribbon.Task(label="Task 4", description="Lorem ipsum", image_mso="HappyFace", children=[
                            bkt.ribbon.Group(label="Test Group 4", children=[
                                bkt.ribbon.TopItems(children=[
                                    bkt.ribbon.Button(
                                        label="Test7",
                                        image_mso="HappyFace",
                                        on_action=action_callback,
                                    ),
                                    bkt.ribbon.Button(
                                        label="Test8",
                                        image_mso="HappyFace",
                                        on_action=action_callback,
                                    )
                                ]),
                            ])
                        ]),

                    ]),
                ])
            ])
        ]
    )



bkt.powerpoint.add_backstage_control(backstage_control1)
bkt.powerpoint.add_backstage_control(backstage_control2)

bkt.excel.add_backstage_control(backstage_control1)
bkt.excel.add_backstage_control(backstage_control2)

bkt.word.add_backstage_control(backstage_control1)
bkt.word.add_backstage_control(backstage_control2)