# -*- coding: utf-8 -*-

import bkt
import bkt.taskpane
#import logging


# ====================
# = Define callbacks =
# ====================



action_callback = bkt.Callback(
    lambda current_control: bkt.message('control clicked: id=%s' % (current_control.id)),
    current_control=True)

action_callback_header = bkt.Callback(
    lambda current_control: bkt.message('control clicked: header=%s id=%s' % (current_control['header'], current_control.id)),
    current_control=True)

action_indexed_callback_header = bkt.Callback(
    lambda selected_item, index, current_control: bkt.message('control clicked: header=%s,\nid=%s\n\nindex=%s, selected_item_id=%s' % (current_control['header'], current_control.id, index, selected_item)),
    current_control=True)

toggle_callback = bkt.Callback(
    lambda pressed, current_control: bkt.message('toggle-button clicked: id=%s\n\npressed-state after click=%s' % (current_control.id, pressed)),
    current_control=True)

toggle_callback_header = bkt.Callback(
    lambda pressed, current_control: bkt.message('toggle-button clicked: header=%s, id=%s\n\npressed/checked-state after click=%s' % (current_control['header'], current_control.id, pressed)),
    current_control=True)

change_callback = bkt.Callback(
    lambda value, current_control: bkt.message('text changed: id=%s\n\nnew text=%s' % (current_control.id, value)),
    current_control=True)
change_callback_header = bkt.Callback(
    lambda value, current_control: bkt.message('text changed: header=%s, id=%s\n\nnew text=%s' % (current_control['header'], current_control.id, value)),
    current_control=True)
value_change_callback_header = bkt.Callback(
    lambda value, current_control, old_value, new_value: bkt.message('value changed: header=%s, id=%s\n\nold value=%s, new value=%s, current value=%s' % (current_control['header'], current_control.id, old_value, new_value, value)),
    current_control=True)


rgb_color_change_callback = bkt.Callback(
    lambda color, current_control: bkt.message('color changed: id=%s\n\nnew color rgb=%s' % (current_control.id, color)),
    current_control=True
)


wpf_callback = bkt.Callback(
    lambda current_control: bkt.message('control event: id=%s' % (current_control.id)),
    current_control=True)


# =========================
# = Resources and Styling =
# =========================

stack_panel_style = bkt.taskpane.XmlPart(
"""<StackPanel.Resources>
    <!-- <BitmapImage x:Key="settings" UriSource="S:\\Tooling\\Toolbox-git\\bkt-framework\\resources\\images\\settings.png"/> // -->
    <!-- Relativer Pfad z.B. <BitmapImage x:Key="settings" UriSource="pack://siteoforigin:,,,/settings.png"/>
         zeigt auf C:\Program Files (x86)\Microsoft Office\Office15\settings.png //-->
    
    
    <SolidColorBrush x:Key="BKTDivider" Color="#1F000000" po:Freeze="True" /> <!-- 12% -->
    <SolidColorBrush x:Key="GroupLabel" Color="#AA000000" po:Freeze="True" />

    <SolidColorBrush x:Key="Ribbon_MouseOverColor"       Color="#FFFBDDD3" />
    <SolidColorBrush x:Key="Ribbon_CheckedColor"         Color="#FFFBC2A5" />
    <SolidColorBrush x:Key="Ribbon_FocusedColor"         Color="#FFFBC2A5" />
    <SolidColorBrush x:Key="Ribbon_PressedColor"         Color="#FFF3AB89" />
    <SolidColorBrush x:Key="Ribbon_MouseOverBorderColor" Color="#FFF3AB89" />
    <SolidColorBrush x:Key="Ribbon_CheckedBorderColor"   Color="#FFF3AB89" />
    <SolidColorBrush x:Key="Ribbon_FocusedBorderColor"   Color="#FFF3AB89" />
    <SolidColorBrush x:Key="Ribbon_PressedBorderColor"   Color="#FFF3AB89" />
    <SolidColorBrush x:Key="Ribbon_BorderColor"          Color="#FFCCCCCC" />

    <Style x:Key="RibbonToolbar">
      <Style.Resources>
          <Style TargetType="{x:Type r:RibbonButton}">
              <Setter Property="Control.Padding" Value="5,4,5,4" />
              <Setter Property="MouseOverBackground" Value="{StaticResource Ribbon_MouseOverColor}"/>
              <Setter Property="FocusedBackground" Value="{StaticResource Ribbon_FocusedColor}" />
              <Setter Property="PressedBackground" Value="{StaticResource Ribbon_PressedBorderColor}" />
          </Style>
          <Style TargetType="{x:Type r:RibbonMenuButton}">
              <Setter Property="Control.Padding" Value="5,4,5,4" />
              <Setter Property="MouseOverBackground" Value="{StaticResource Ribbon_MouseOverColor}" />
              <Setter Property="MouseOverBorderBrush" Value="{StaticResource Ribbon_MouseOverColor}" />
              <Setter Property="FocusedBackground" Value="{StaticResource Ribbon_MouseOverColor}" />
              <Setter Property="PressedBackground" Value="{StaticResource Ribbon_PressedColor}" />
          </Style>
          <Style TargetType="{x:Type r:RibbonSplitButton}">
              <Setter Property="Control.Padding" Value="5,4,5,4" />
              <Setter Property="MouseOverBackground" Value="{StaticResource Ribbon_MouseOverColor}" />
              <Setter Property="MouseOverBorderBrush" Value="{StaticResource Ribbon_MouseOverColor}" />
              <Setter Property="FocusedBackground" Value="{StaticResource Ribbon_MouseOverColor}" />
              <Setter Property="PressedBackground" Value="{StaticResource Ribbon_PressedColor}" />
              <Setter Property="CheckedBackground" Value="{StaticResource Ribbon_CheckedColor}" />
              <Setter Property="CheckedBorderBrush" Value="{StaticResource Ribbon_CheckedBorderColor}" />
          </Style>
          <Style TargetType="{x:Type r:RibbonToggleButton}">
              <Setter Property="Control.Padding" Value="5,4,5,4" />
              <Setter Property="MouseOverBackground" Value="{StaticResource Ribbon_MouseOverColor}" />
              <Setter Property="FocusedBackground" Value="{StaticResource Ribbon_MouseOverColor}" />
              <Setter Property="PressedBackground" Value="{StaticResource Ribbon_PressedColor}" />
              <Setter Property="PressedBorderBrush" Value="{StaticResource Ribbon_PressedColor}" />
              <Setter Property="CheckedBackground" Value="{StaticResource Ribbon_PressedColor}" />
              <Setter Property="CheckedBorderBrush" Value="{x:Null}" />
          </Style>
          <Style TargetType="{x:Type r:RibbonMenuItem}">
              <Setter Property="MouseOverBackground" Value="{StaticResource Ribbon_MouseOverColor}" />
              <Setter Property="MouseOverBorderBrush" Value="{StaticResource Ribbon_MouseOverColor}" />
              <Setter Property="PressedBackground" Value="{StaticResource Ribbon_PressedColor}" />
          </Style>
          <Style TargetType="{x:Type r:RibbonGalleryItem}">
              <Setter Property="MouseOverBackground" Value="{StaticResource Ribbon_MouseOverColor}" />
              <Setter Property="PressedBackground" Value="{StaticResource Ribbon_PressedColor}" />
              <Setter Property="CheckedBackground" Value="{StaticResource Ribbon_CheckedColor}" />
              <Setter Property="CheckedBorderBrush" Value="{StaticResource Ribbon_CheckedBorderColor}" />
          </Style>
          <Style TargetType="{x:Type r:RibbonCheckBox}">
              <Setter Property="MouseOverBorderBrush" Value="{StaticResource Ribbon_CheckedBorderColor}" />
              <Setter Property="MouseOverBackground" Value="{StaticResource Ribbon_CheckedColor}" />
              <Setter Property="CheckedBorderBrush" Value="{StaticResource Ribbon_CheckedBorderColor}" />
              <Setter Property="CheckedBackground" Value="{StaticResource Ribbon_CheckedColor}" />
              <Setter Property="PressedBorderBrush" Value="{StaticResource Ribbon_PressedBorderColor}" />
              <Setter Property="PressedBackground" Value="{StaticResource Ribbon_PressedColor}" />
              <Setter Property="FocusedBorderBrush" Value="{StaticResource Ribbon_FocusedBorderColor}" />
              <Setter Property="FocusedBackground" Value="{StaticResource Ribbon_FocusedColor}" />
          </Style>
          <Style TargetType="{x:Type r:RibbonTextBox}">
              <Setter Property="Control.BorderBrush" Value="#22000000" />
              <Setter Property="MouseOverBorderBrush" Value="{StaticResource Ribbon_MouseOverBorderColor}" />
              <Setter Property="FocusedBorderBrush" Value="{StaticResource Ribbon_FocusedBorderColor}" />
          </Style>
          <Style TargetType="{x:Type r:RibbonComboBox}">
              <Setter Property="Control.Padding" Value="5,4,5,4" />
              <Setter Property="Control.BorderBrush" Value="#22000000" />
              <Setter Property="MouseOverBorderBrush" Value="{StaticResource Ribbon_MouseOverBorderColor}" />
              <Setter Property="FocusedBorderBrush" Value="{StaticResource Ribbon_FocusedBorderColor}" />
              <Setter Property="PressedBackground" Value="{StaticResource Ribbon_PressedColor}" />
          </Style>
          <Style TargetType="{x:Type r:RibbonSeparator}">
              <Setter Property="Control.BorderBrush" Value="#44000000" />
              <Setter Property="Control.Background" Value="#44000000" />
              <Setter Property="Control.BorderThickness" Value="4" />
          </Style>
      </Style.Resources>
    </Style>
</StackPanel.Resources>""")

bkt.powerpoint.add_taskpane_control(stack_panel_style)





# =========================
# = Ribbon-Style Controls =
# =========================



# ribbon_style_expander = bkt.taskpane.Wpf.Expander(
#     Header="Windows Ribbon",
#     #IsExpanded="True",
#     children = [
#         bkt.taskpane.Wpf.StackPanel(
#             Orientation="Vertical",
#             Margin="0,5,0,5",
#             children = [
#                 bkt.taskpane.Wpf.WrapPanel(
#                     Orientation="Horizontal",
#                     Margin="0,5,0,5",
#                     Style="{StaticResource RibbonToolbar}",
#                     children = [
#                         bkt.taskpane.Ribbon.RibbonButton(Label="Button A"),
#                         bkt.taskpane.Ribbon.RibbonButton(Label="Button B"),
#                         bkt.taskpane.Ribbon.RibbonToggleButton(Label="Toggle")
#                 ]),
#                 bkt.taskpane.Wpf.WrapPanel(
#                     Orientation="Horizontal",
#                     Margin="0,5,0,5",
#                     Style="{StaticResource RibbonToolbar}",
#                     children = [
#                         bkt.taskpane.Ribbon.RibbonMenuButton(
#                             Label="Menu",
#                             children=[
#                                 bkt.taskpane.Ribbon.RibbonMenuItem(Header="Button 1"),
#                                 bkt.taskpane.Ribbon.RibbonMenuItem(Header="Button 2"),
#                                 bkt.taskpane.Ribbon.RibbonMenuItem(Header="Button 3"),
#                                 bkt.taskpane.Ribbon.RibbonSeparator(),
#                                 bkt.taskpane.Ribbon.RibbonMenuItem(Header="Button 4"),
#                         ]),
#                         bkt.taskpane.Ribbon.RibbonSplitButton(
#                             Label="Split Button",
#                             children=[
#                                 bkt.taskpane.Ribbon.RibbonMenuItem(Header="Menu Item #1"),
#                                 bkt.taskpane.Ribbon.RibbonMenuItem(Header="Menu Item #2"),
#                                 bkt.taskpane.Ribbon.RibbonMenuItem(Header="Menu Item #3"),
#                                 bkt.taskpane.Ribbon.RibbonMenuItem(Header="Menu Item #4"),
#                                 bkt.taskpane.Ribbon.RibbonMenuItem(Header="Menu Item #5")
#                         ]),
#                 ]),
#                 bkt.taskpane.Wpf.WrapPanel(
#                     Orientation="Horizontal",
#                     Margin="0,5,0,5",
#                     Style="{StaticResource RibbonToolbar}",
#                     children=[
#                         bkt.taskpane.Ribbon.RibbonCheckBox(Label="check me"),
#                         bkt.taskpane.Ribbon.RibbonTextBox(Label="Foo", Text="Bar"),
#                         bkt.taskpane.Ribbon.RibbonComboBox(
#                             Label="Combo",
#                             IsEditable="False",
#                             children = [
#                                 bkt.taskpane.Ribbon.RibbonGallery(
#                                     SelectedValue="Einfuegen",
#                                     SelectedValuePath="Content",
#                                     MaxColumnCount="1",
#                                     children=[
#                                         bkt.taskpane.Ribbon.RibbonGalleryCategory(children=[
#                                             bkt.taskpane.Ribbon.RibbonGalleryItem(Content="Einfuegen"),
#                                             bkt.taskpane.Ribbon.RibbonGalleryItem(Content="Inhalte einfuegen ...")
#                                         ])
#                                 ])
#                         ]),
#                         bkt.taskpane.Ribbon.RibbonComboBox(
#                             Label="Edit-Combo",
#                             IsEditable="True",
#                             children = [
#                                 bkt.taskpane.Ribbon.RibbonMenuItem(Header="Menu Item #1"),
#                                 bkt.taskpane.Ribbon.RibbonMenuItem(Header="Menu Item #2"),
#                                 bkt.taskpane.Ribbon.RibbonMenuItem(Header="Menu Item #3"),
#                         ])
#                 ])
#         ])
# ])
#
# bkt.powerpoint.add_taskpane_control(ribbon_style_expander)
#


            



# =======================
# = Aligned Input Boxes =
# =======================

grid_expander = bkt.taskpane.Wpf.Expander(
    Header="Grid-aligned controls",
    IsExpanded="False",
    children = [
        bkt.taskpane.Wpf.StackPanel(
            Orientation="Vertical",
            Margin="0,5,0,5",
            Style="{StaticResource RibbonToolbar}",
            children = [
                bkt.taskpane.Wpf.Grid(
                    children=[
                        bkt.taskpane.XmlPart(
                            """<Grid.Resources>
                                <Style TargetType="{x:Type TextBox}">
                                    <Setter Property="Margin" Value="0,0,0,4" />
                                </Style>
                            </Grid.Resources>"""),
                        bkt.taskpane.XmlPart(
                            """<Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>"""),
                        bkt.taskpane.XmlPart(
                            """<Grid.RowDefinitions>
                                <RowDefinition Height="*"/>
                                <RowDefinition Height="*"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>"""),
                        bkt.taskpane.Wpf.Label(   attributes={"Grid.Column": "0"}, Padding="5,3,5,1", Content="Text"),
                        bkt.taskpane.Ribbon.RibbonTextBox( attributes={"Grid.Column":"1"},  Padding="0,2,0,2", Text="Text"),
                        bkt.taskpane.Wpf.Label(   attributes={"Grid.Column":"0", "Grid.Row":"1"}, Padding="5,3,5,1", Content="long label"),
                        bkt.taskpane.Ribbon.RibbonTextBox( attributes={"Grid.Column":"1", "Grid.Row":"1"}, Padding="0,2,0,2", Text="Text"),
                        bkt.taskpane.Wpf.Label(   attributes={"Grid.Column":"0", "Grid.Row":"2"}, Padding="5,3,5,1", Content="very long label"),
                        bkt.taskpane.Ribbon.RibbonTextBox( attributes={"Grid.Column":"1", "Grid.Row":"2"}, Padding="0,2,0,2", Text="Text")
                ])
        ])
])

bkt.powerpoint.add_taskpane_control(grid_expander)





# ==================
# = Demo Callbacks =
# ==================

    

# action_expander = bkt.taskpane.Wpf.Expander(
#     Header="Demo actions",
#     children=[
#         bkt.taskpane.Wpf.WrapPanel(
#             Orientation="Horizontal",
#             Margin="0,5,0,5",
#             children=[
#                 bkt.taskpane.Wpf.Button(
#                     Content="on action",
#                     on_action = action_callback
#                 ),
#                 bkt.taskpane.Wpf.Button(
#                     Content="wpf click",
#                     wpf_event = wpf_callback
#                 ),
#                 bkt.taskpane.Wpf.ToggleButton(
#                     Content="toggle action",
#                     on_toggle_action = toggle_callback
#                 )
#         ])
# ])
#
# bkt.powerpoint.add_taskpane_control(action_expander)



# ==========
# = Groups =
# ==========

# group_expander = bkt.taskpane.Wpf.Expander(
#     Header="Group separators",
#     children=[
#         bkt.taskpane.Wpf.StackPanel(
#             Orientation="Vertical",
#             Margin="0,5,0,0",
#             children=[
#                 bkt.taskpane.Wpf.Button(Content="Dummy Button"),
#                 bkt.taskpane.GroupSeparator(label="Group Centered Style"),
#                 bkt.taskpane.Wpf.Button(Content="Dummy Button"),
#                 bkt.taskpane.GroupSeparator(Label="Group Centered Style 2"),
#                 bkt.taskpane.Wpf.Button(Content="Dummy Button"),
#                 bkt.taskpane.GroupSeparator(),
#                 bkt.taskpane.Wpf.Button(Content="Dummy Button"),
#                 bkt.taskpane.Group(label="a group with subcontrols", children=[
#                     bkt.taskpane.Wpf.WrapPanel(
#                         Orientation="Horizontal",
#                         Margin="0,5,0,5",
#                         children=[bkt.taskpane.Wpf.Button(Content="#1"), bkt.taskpane.Wpf.Button(Content="#2"), bkt.taskpane.Wpf.Button(Content="#3")
#                     ])
#                 ]),
#         ])
# ])
#
# bkt.powerpoint.add_taskpane_control(group_expander)





# ============================================
# = Basic Wpf Controls with full flexibility =
# ============================================

# expander = bkt.taskpane.Wpf.Expander(
#     Header="Full flexibility",
#     children=[
#         bkt.taskpane.XmlPart(
#         """<StackPanel Orientation="Vertical">
#         <Label>Any Wpf controls can be placed here</Label>
#         <WrapPanel Orientation="Horizontal" Margin="0,5,0,5" Style="{StaticResource RibbonToolbar}">
#             <r:RibbonButton Label="Button A" SmallImageSource="{StaticResource settings}" Padding="5,2,5,2" Height="28"/>
#             <r:RibbonButton Label="Button B" />
#             <r:RibbonToggleButton Label="Toggle" />
#         </WrapPanel>
#         <WrapPanel Orientation="Horizontal" Margin="0,5,0,5" Style="{StaticResource RibbonToolbar}">
#             <r:RibbonButton Label="Button A" LargeImageSource="{StaticResource settings}" Padding="5,0,5,0" Height="68"/>
#             <r:RibbonButton Label="Button B" />
#             <r:RibbonToggleButton Label="Toggle" />
#         </WrapPanel>
#         </StackPanel>
#         """
#         )
# ])
#
# bkt.powerpoint.add_taskpane_control(expander)




# =================
# = Fluent Ribbon =
# =================

expander = bkt.taskpane.Wpf.Expander(header="Fluent Ribbon",
    children=[
        bkt.taskpane.Wpf.StackPanel(
            Orientation="Vertical",
            children=[
                bkt.taskpane.FluentRibbon.Button(header="Button #1", size="middle"), #, click="TaskPane_Wpf_Event"),
                bkt.taskpane.FluentRibbon.Button(header="#2", size="middle"),
                bkt.taskpane.FluentRibbon.Button(header="#3", size="middle")
        ])
])
bkt.powerpoint.add_taskpane_control(expander)

# as in ribbon: Expander / Buttons
expander = bkt.taskpane.Expander(auto_wrap=True, header="Expander with single Button-Group",
    children=[
        bkt.taskpane.Button(header="my"     , on_action = action_callback_header),
        bkt.taskpane.Button(header="first"  , on_action = action_callback_header),
        bkt.taskpane.Button(header="buttons", on_action = action_callback_header)
])
bkt.powerpoint.add_taskpane_control(expander)

# as in ribbon: Expander / Group / Buttons
expander = bkt.taskpane.Expander(auto_stack=True, header="Expander with multiple groups",
    children=[
        bkt.taskpane.Group(auto_wrap=True, label="first group",
            children=[
                bkt.taskpane.Button(header="my"     , on_action = action_callback_header),
                bkt.taskpane.Button(header="first"  , on_action = action_callback_header),
                bkt.taskpane.Button(header="buttons", on_action = action_callback_header)
        ]),
        bkt.taskpane.Group(auto_wrap=True, label="second group",
            children=[
                bkt.taskpane.Button(header="some"   , wpf_event = wpf_callback),
                bkt.taskpane.Button(header="more"   , wpf_event = wpf_callback),
                bkt.taskpane.Button(header="buttons", wpf_event = wpf_callback)
        ])
])
bkt.powerpoint.add_taskpane_control(expander)

# as in ribbon: Expander / Group / Buttons
expander = bkt.taskpane.Expander(auto_stack=True, header="Fluent Ribbon control examples",
    children=[
        # BUTTONS
        bkt.taskpane.Group(auto_wrap=True, label="buttons",
            children=[
                bkt.taskpane.Button(header="my", image="settings", 
                    on_action = action_callback_header,
                    # Screentip
                    screentip="This is a ScreenTip.",
                    screentip_image="settings",
                    screentip_title="Button #1 Screentip Title",
                    help_topic="Help for ScreenTip"),
                bkt.taskpane.Button(header="first", image="settings", on_action = action_callback_header),
                bkt.taskpane.Button(header="buttons", image="settings", is_enabled=False,
                    on_action = action_callback_header,
                    # Screentip
                    screentip="This is another ScreenTip.",
                    screentip_image="settings",
                    screentip_title="Button #3 Screentip Title",
                    disable_reason="Lorem ipsum dolor sit amet.",
                    help_topic="Help for ScreenTip")
        ]),
        # TOGGLE BUTTONS 
        bkt.taskpane.Group(auto_wrap=True, label="toggles",
            children=[
                bkt.taskpane.FluentRibbon.ToggleButton(header="toggle me", size="middle", image="settings", on_toggle_action=toggle_callback_header),
                bkt.taskpane.FluentRibbon.CheckBox(header="check me", is_checked=True, image="settings", on_toggle_action=toggle_callback_header)
        ]),
        bkt.taskpane.Group(auto_wrap=True, show_separator=False,
            children=[
                bkt.taskpane.FluentRibbon.RadioButton(header="radio 1", group_name="test_radio_group", image="settings", on_toggle_action=toggle_callback_header, is_checked=True),
                bkt.taskpane.FluentRibbon.RadioButton(header="radio 2", group_name="test_radio_group", image="settings", on_toggle_action=toggle_callback_header)
        ]),
        bkt.taskpane.Group(auto_wrap=True, show_separator=False,
            children=[
                bkt.taskpane.FluentRibbon.ToggleButton(header="Toggle #1", group_name="test_toggle_group", size="Middle", image="settings", on_toggle_action=toggle_callback_header, is_checked=True),
                bkt.taskpane.FluentRibbon.ToggleButton(header="#2",        group_name="test_toggle_group", size="Middle", image="settings", on_toggle_action=toggle_callback_header),
                bkt.taskpane.FluentRibbon.ToggleButton(header="#3",        group_name="test_toggle_group", size="Middle", image="settings", on_toggle_action=toggle_callback_header)
        ]),
        # MENUS
        bkt.taskpane.Group(auto_wrap=True, label="Menus",
            children=[
                bkt.taskpane.FluentRibbon.DropDownButton(header="Simple Menu", size="Middle", image="settings", children=[
                    bkt.taskpane.FluentRibbon.MenuItem(header="Item 1", on_action=action_callback_header),
                    bkt.taskpane.FluentRibbon.MenuItem(header="Item 2", on_action=action_callback_header),
                    bkt.taskpane.Wpf.Separator(),
                    bkt.taskpane.FluentRibbon.MenuItem(header="Item 3", children=[
                        bkt.taskpane.FluentRibbon.MenuItem(header="Item 1", on_action=action_callback_header),
                        bkt.taskpane.FluentRibbon.MenuItem(header="Item 2", on_action=action_callback_header)
                    ]),
                    bkt.taskpane.FluentRibbon.MenuItem(header="Item 4", is_splited=True, on_action=action_callback_header, children=[
                        bkt.taskpane.FluentRibbon.MenuItem(header="Item 1", on_action=action_callback_header),
                        bkt.taskpane.FluentRibbon.MenuItem(header="Item 2", on_action=action_callback_header)
                    ]),
                    bkt.taskpane.FluentRibbon.MenuItem(header="Item 5", on_action=action_callback_header),
                    bkt.taskpane.FluentRibbon.MenuItem(header="Check me", is_checkable=True, is_checked=True, on_toggle_action=toggle_callback_header)
                ]),
                bkt.taskpane.FluentRibbon.SplitButton(header="Split Button", size="Middle", image="settings", on_action=action_callback_header, children=[
                    bkt.taskpane.FluentRibbon.MenuItem(header="Item 1", on_action=action_callback_header),
                    bkt.taskpane.FluentRibbon.MenuItem(header="Item 2", on_action=action_callback_header),
                    bkt.taskpane.FluentRibbon.MenuItem(header="Item 3", on_action=action_callback_header),
                ])
        ]),
        # INPUT
        bkt.taskpane.Group(auto_wrap=True, label="Input",
            children=[
                bkt.taskpane.FluentRibbon.TextBox(header="Text", max_length=5, input_width=70, text="default text", image="settings", on_change=change_callback_header),
                bkt.taskpane.FluentRibbon.ComboBox(header="ComboBox", is_read_only=True, selected_index='0', image="settings",
                    
                    # events: use action_indexed or on_change to either get selected-index or text-value
                    on_action_indexed=action_indexed_callback_header,
                    on_change=change_callback_header,
                    
                    # add button-menu to ComboBox
                    menu = bkt.taskpane.FluentRibbon.RibbonMenu(children=[
                        bkt.taskpane.FluentRibbon.MenuItem(header="Menu Item 1", on_action=action_callback_header),
                        bkt.taskpane.Wpf.Separator(),
                        bkt.taskpane.FluentRibbon.MenuItem(header="Menu Item 2", on_action=action_callback_header),
                        bkt.taskpane.FluentRibbon.MenuItem(header="Menu Item 3", on_action=action_callback_header),
                    ]),
                    
                    # defined Combo-Items
                    children=[
                        bkt.taskpane.Wpf.ComboBoxItem(content="Combo Item #1"),
                        bkt.taskpane.Wpf.ComboBoxItem(content="Combo Item #2"),
                        bkt.taskpane.Wpf.ComboBoxItem(content="Combo Item #3"),
                        bkt.taskpane.Wpf.ComboBoxItem(content="Combo Item #4"),
                    ]
                ),
                bkt.taskpane.FluentRibbon.ComboBox(header="ComboBox editable", is_read_only=False, selected_index='0', image="settings", 
                    on_change=change_callback_header,
                    children=[
                        bkt.taskpane.Wpf.ComboBoxItem(content="Combo Item #1"),
                        bkt.taskpane.Wpf.ComboBoxItem(content="Combo Item #2"),
                        bkt.taskpane.Wpf.ComboBoxItem(content="Combo Item #3"),
                        bkt.taskpane.Wpf.ComboBoxItem(content="Combo Item #4")
                    ]
                ),
                bkt.taskpane.FluentRibbon.Spinner(header="Spinner", input_width=100, image="settings", on_change=value_change_callback_header),
                bkt.taskpane.FluentRibbon.Spinner(id="spin2", header="Spinner", input_width=100, image="settings", on_change=value_change_callback_header),
                bkt.taskpane.FluentRibbon.Spinner(header="Spinner (%)", input_width=100, image="settings", on_change=value_change_callback_header, minimum='-1', maximum='1', increment=0.05, format='P0'),
                bkt.taskpane.Wpf.DatePicker(on_change=change_callback)
        ]),
        # GALLERY
        bkt.taskpane.Group(auto_wrap=True, label="Galleries", children=[
            bkt.taskpane.FluentRibbon.DropDownButton(header="Gallery Menu", size="Middle", image="settings", children=[
                bkt.taskpane.FluentRibbon.Gallery(item_width=20, item_height=20, on_change=change_callback, children=[
                    bkt.taskpane.Wpf.TextBlock(text="1"),
                    bkt.taskpane.Wpf.TextBlock(text="2"),
                    bkt.taskpane.Wpf.TextBlock(text="3")
                ]),
                bkt.taskpane.FluentRibbon.MenuItem(header="Menu Item 1", on_action=action_callback_header),
                bkt.taskpane.FluentRibbon.MenuItem(header="Menu Item 2", on_action=action_callback_header)
            ]),
            bkt.taskpane.FluentRibbon.DropDownButton(header="Gallery Menu (with groups)", size="Middle", image="settings", children=[
                bkt.taskpane.FluentRibbon.Gallery(item_width=20, item_height=20, group_by="Tag", on_change=change_callback, children=[
                    bkt.taskpane.Wpf.TextBlock(tag="Group 1", text="1"),
                    bkt.taskpane.Wpf.TextBlock(tag="Group 1", text="2"),
                    bkt.taskpane.Wpf.TextBlock(tag="Group 1", text="3"),
                    bkt.taskpane.Wpf.TextBlock(tag="Group 1", text="4"),
                    bkt.taskpane.Wpf.TextBlock(tag="Group 2", text="5"),
                    bkt.taskpane.Wpf.TextBlock(tag="Group 2", text="6"),
                    bkt.taskpane.Wpf.TextBlock(tag="Group 2", text="7")
                ]),
                bkt.taskpane.FluentRibbon.MenuItem(header="Menu Item 1", on_action=action_callback_header),
                bkt.taskpane.FluentRibbon.MenuItem(header="Menu Item 2", on_action=action_callback_header)
            ]),
            
            # COLOR GALLERIES
            bkt.taskpane.FluentRibbon.DropDownButton(header="Standard colors", size="Middle", image="settings", children=[
                bkt.taskpane.FluentRibbon.ColorGallery(mode="StandardColors", on_rgb_color_change=rgb_color_change_callback)
            ]),
            bkt.taskpane.FluentRibbon.DropDownButton(header="Highlight colors", size="Middle", image="settings", children=[
                bkt.taskpane.FluentRibbon.ColorGallery(mode="HighlightColors", IsAutomaticColorButtonVisible=False, on_rgb_color_change=rgb_color_change_callback)
            ]),
            bkt.taskpane.FluentRibbon.DropDownButton(header="Theme colors", size="Middle", image="settings", children=[
                bkt.taskpane.FluentRibbon.ColorGallery(mode="ThemeColors", StandardColorGridRows="3", Columns="10", ThemeColorGridRows="5", IsNoColorButtonVisible="True", on_rgb_color_change=rgb_color_change_callback)
            ]),
            
            # INLINE GALLERY
            bkt.taskpane.FluentRibbon.InRibbonGallery(header="Inline gallery", item_width=40, item_height=40, group_by="Tag", image="settings", 
                max_items_in_row=4, min_items_in_row=2, min_items_in_drop_down_row=3, on_change=change_callback_header, children=[
                    bkt.taskpane.Wpf.TextBlock(tag="Group 1", text="1"),
                    bkt.taskpane.Wpf.TextBlock(tag="Group 1", text="2"),
                    bkt.taskpane.Wpf.TextBlock(tag="Group 1", text="3"),
                    bkt.taskpane.Wpf.TextBlock(tag="Group 1", text="4"),
                    bkt.taskpane.Wpf.TextBlock(tag="Group 2", text="A"),
                    bkt.taskpane.Wpf.TextBlock(tag="Group 2", text="B"),
                    bkt.taskpane.Wpf.TextBlock(tag="Group 2", text="C")
            ])
            
        ])

])
bkt.powerpoint.add_taskpane_control(expander)


# ======================
# = Fluent Ribbon Xaml =
# ======================
#
# import os
#
# xaml_filename=os.path.join(os.path.dirname(os.path.realpath(__file__)), "demo_fluent_ribbon.xaml")
# xml_string = ""
# for line in open(xaml_filename, "r"):
#     xml_string += line + "\n"
#
# expander = bkt.taskpane.Wpf.Expander(
#     Header="Fluent Ribbon (2)",
#     children=[bkt.taskpane.XmlPart(xml_string)]
# )
#
# bkt.powerpoint.add_taskpane_control(expander)
#

