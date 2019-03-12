import bkt
import unittest


XMLNS = ' xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"'
XMLNS_FR = ' xmlns="urn:fluent-ribbon"'


class TaskpaneBaseObjectTest(unittest.TestCase):
    
    
    def test_XamlPropertyElement(self):
        bkt.taskpane.TaskPaneControl.no_id = True
        
        # default XamlPropertyElement
        b = bkt.taskpane.XamlPropertyElement()
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<NotSpecified.NotSpecified' + XMLNS + ' />')
        #self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<NotSpecified.Resources xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" />')

        # specifying property name failed
        b = bkt.taskpane.XamlPropertyElement(property_name="PropertyName")
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<NotSpecified.PropertyName' + XMLNS + ' />')

        # specifying type name failed
        b = bkt.taskpane.XamlPropertyElement(type_name="TypeName")
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<TypeName.NotSpecified' + XMLNS + ' />')

        # specifying type name and property name failed
        b = bkt.taskpane.XamlPropertyElement(type_name="TypeName", property_name="PropertyName")
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<TypeName.PropertyName' + XMLNS + ' />')

        # specifying type name at xml-generation failed
        b = bkt.taskpane.XamlPropertyElement(property_name="PropertyName")
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml("TypeName")), u'<TypeName.PropertyName' + XMLNS + ' />')
        
        bkt.taskpane.TaskPaneControl.no_id = False
    
    
    def test_XamlPropertyElement_fixed_type(self):
        bkt.taskpane.TaskPaneControl.no_id = True
        myclass = type("myclassname", (bkt.taskpane.XamlPropertyElement,), {'_type_name': 'FixedTypeName'})
        
        # Definition of XamlPropertyElement with fixed type name failed
        b = myclass()
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<FixedTypeName.NotSpecified' + XMLNS + ' />')

        # specifying property name failed
        b = myclass(property_name="PropertyName")
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<FixedTypeName.PropertyName' + XMLNS + ' />')

        # type name should be overwritable
        b = myclass(type_name="TypeName")
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<TypeName.NotSpecified' + XMLNS + ' />')

        # type name and property name should be overwritable
        b = myclass(type_name="TypeName", property_name="PropertyName")
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<TypeName.PropertyName' + XMLNS + ' />')
        
        bkt.taskpane.TaskPaneControl.no_id = False
    
    
    def test_XamlPropertyElement_fixed_property(self):
        bkt.taskpane.TaskPaneControl.no_id = True
        
        myclass = type("myclassname", (bkt.taskpane.XamlPropertyElement,), {'_property_name': 'FixedPropertyName'})
        
        # Definition of XamlPropertyElement with fixed property name failed
        b = myclass()
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<NotSpecified.FixedPropertyName' + XMLNS + ' />')

        # property name should be overwritable
        b = myclass(property_name="PropertyName")
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<NotSpecified.PropertyName' + XMLNS + ' />')

        # specifying type name failed
        b = myclass(type_name="TypeName")
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<TypeName.FixedPropertyName' + XMLNS + ' />')
        
        bkt.taskpane.TaskPaneControl.no_id = False
    
    
    def test_XamlPropertyElements(self):
        self.maxDiff = None
        bkt.taskpane.TaskPaneControl.no_id = True

        # simple usage of XamlPropertyElementGenerator failed
        b = bkt.taskpane.XamlPropertyElements.Resources()
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<NotSpecified.Resources' + XMLNS + ' />')
        
        # specification of type name failed
        b = bkt.taskpane.XamlPropertyElements.Resources("Button")
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<Button.Resources' + XMLNS + ' />')

        # specification of type name failed
        b = bkt.taskpane.XamlPropertyElements.Resources(type_name="Button")
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<Button.Resources' + XMLNS + ' />')

        # property name should be overwritable
        b = bkt.taskpane.XamlPropertyElements.Resources(type_name="Button", property_name="Overwritten")
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<Button.Overwritten' + XMLNS + ' />')
        
        bkt.taskpane.TaskPaneControl.no_id = False
    
    
    def test_XamlPropertyElement_Attribute(self):
        bkt.taskpane.TaskPaneControl.no_id = True

        # usage of XamlPropertyElement as attribute failed
        b = bkt.taskpane.TaskPaneControl(resources=bkt.taskpane.XamlPropertyElement(property_name="PropertyName"))
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<TaskPaneControl' + XMLNS + '>\r\n  <TaskPaneControl.PropertyName />\r\n</TaskPaneControl>')
        
        # usage of XamlPropertyElementGenerator as attribute failed
        b = bkt.taskpane.TaskPaneControl(resources=bkt.taskpane.XamlPropertyElements.Resources())
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<TaskPaneControl' + XMLNS + '>\r\n  <TaskPaneControl.Resources />\r\n</TaskPaneControl>')

        # usage of XamlPropertyElementGenerator with other xml-namespace failed
        b = bkt.taskpane.FluentRibbon.Button(resources=bkt.taskpane.XamlPropertyElements.Resources())
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<Button' + XMLNS_FR + '>\r\n  <Button.Resources />\r\n</Button>')
        
        # type name should not be overwritable
        b = bkt.taskpane.TaskPaneControl(resources=bkt.taskpane.XamlPropertyElement(type_name="TypeName", property_name="PropertyName"))
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<TaskPaneControl' + XMLNS + '>\r\n  <TaskPaneControl.PropertyName />\r\n</TaskPaneControl>')
        
        # type name should not be overwritable
        b = bkt.taskpane.TaskPaneControl(resources=bkt.taskpane.XamlPropertyElements.Resources(type_name="TypeName"))
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<TaskPaneControl' + XMLNS + '>\r\n  <TaskPaneControl.Resources />\r\n</TaskPaneControl>')
        
        bkt.taskpane.TaskPaneControl.no_id = False
    
    
    def test_XamlPropertyElement_Child(self):
        bkt.taskpane.TaskPaneControl.no_id = True

        # usage of XamlPropertyElement as child-element faild
        b = bkt.taskpane.TaskPaneControl(children=[bkt.taskpane.XamlPropertyElement(type_name="TypeName", property_name="PropertyName")])
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<TaskPaneControl' + XMLNS + '>\r\n  <TypeName.PropertyName />\r\n</TaskPaneControl>')
        
        # usage of XamlPropertyElementGenerator as child-element failed
        b = bkt.taskpane.TaskPaneControl(children=[bkt.taskpane.XamlPropertyElements.Resources("TypeName")])
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<TaskPaneControl' + XMLNS + '>\r\n  <TypeName.Resources />\r\n</TaskPaneControl>')
        
        # no default type_name if child definition is used
        # type name should have no fallback if child definition is used
        b = bkt.taskpane.TaskPaneControl(children=[bkt.taskpane.XamlPropertyElement(property_name="PropertyName")])
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<TaskPaneControl' + XMLNS + '>\r\n  <NotSpecified.PropertyName />\r\n</TaskPaneControl>')
        
        # type name should have no fallback if child definition is used
        b = bkt.taskpane.TaskPaneControl(children=[bkt.taskpane.XamlPropertyElements.Resources()])
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<TaskPaneControl' + XMLNS + '>\r\n  <NotSpecified.Resources />\r\n</TaskPaneControl>')
        
        bkt.taskpane.TaskPaneControl.no_id = False
    
    
    def test_WPFRibbon(self):
        bkt.taskpane.TaskPaneControl.no_id = True

        # definition of RibbonButton failed
        b = bkt.taskpane.Ribbon.RibbonButton()
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<RibbonButton xmlns="clr-namespace:System.Windows.Controls.Ribbon;assembly=System.Windows.Controls.Ribbon" />')

        bkt.taskpane.TaskPaneControl.no_id = False


    def test_FluentRibbon(self):
        bkt.taskpane.TaskPaneControl.no_id = True
        bkt.taskpane.FluentRibbonControl.no_id = True

        # definition of FluentRibbon-Button failed
        b = bkt.taskpane.FluentRibbon.Button()
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<Button xmlns="urn:fluent-ribbon" />')

        bkt.taskpane.FluentRibbonControl.no_id = False
    
    
    def test_FluentRibbon_ScreenTip(self):
        bkt.taskpane.TaskPaneControl.no_id = True
        bkt.taskpane.FluentRibbonControl.no_id = True
        
        # Button with simple tooltip attribute failed
        b = bkt.taskpane.FluentRibbon.Button(tool_tip="Tooltip text")
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<Button ToolTip="Tooltip text"' + XMLNS_FR + ' />')
        
        # Definition ToolTip-Property-Element failed
        b = bkt.taskpane.ToolTip("Button")
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<Button.ToolTip' + XMLNS_FR + ' />')
        
        # Definition of Screentip-Element failed
        b = bkt.taskpane.FluentRibbon.ScreenTip(text="screentip text")
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<ScreenTip Text="screentip text"' + XMLNS_FR + ' />')

        # Definition of Screentip-Element failed
        b = bkt.taskpane.FluentRibbon.ScreenTip(text="screentip text", title="screentip title", disable_reason="This button is diabled because ...", help_topic="Info for additional help")
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<ScreenTip DisableReason="This button is diabled because ..." HelpTopic="Info for additional help" Text="screentip text" Title="screentip title"' + XMLNS_FR + ' />')
        
        # Screentip-attribute should be parsed to ScreenTip-object
        b = bkt.taskpane.FluentRibbon.Button(screentip="Screentip text")
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<Button' + XMLNS_FR + '>\r\n  <Button.ToolTip>\r\n    <ScreenTip IsRibbonAligned="False" Text="Screentip text" />\r\n  </Button.ToolTip>\r\n</Button>')

        # Screentip-attribute should be parsed to ScreenTip-object
        b = bkt.taskpane.FluentRibbon.Button(screentip="Screentip text", screentip_title="Title", disable_reason="This button is diabled because ...", help_topic="Info for additional help")
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<Button' + XMLNS_FR + '>\r\n  <Button.ToolTip>\r\n    <ScreenTip DisableReason="This button is diabled because ..." HelpTopic="Info for additional help" IsRibbonAligned="False" Text="Screentip text" Title="Title" />\r\n  </Button.ToolTip>\r\n</Button>')
        
        # Screentip definition should overwrite tooltip
        b = bkt.taskpane.FluentRibbon.Button(tool_tip="Tooltip text", screentip="Screentip text")
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<Button' + XMLNS_FR + '>\r\n  <Button.ToolTip>\r\n    <ScreenTip IsRibbonAligned="False" Text="Screentip text" />\r\n  </Button.ToolTip>\r\n</Button>')
        
        # Definition of screentip through tooltip-attribute failed
        b = bkt.taskpane.FluentRibbon.Button(tool_tip=bkt.taskpane.FluentRibbon.ScreenTip(text="Screentip text"))
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<Button' + XMLNS_FR + '>\r\n  <Button.ToolTip>\r\n    <ScreenTip Text="Screentip text" />\r\n  </Button.ToolTip>\r\n</Button>')


        bkt.taskpane.FluentRibbonControl.no_id = False
    
    
    def test_FluentRibbon_Image(self):
        bkt.taskpane.TaskPaneControl.no_id = True
        bkt.taskpane.FluentRibbonControl.no_id = True
        
        b = bkt.taskpane.FluentRibbon.Button(image="test_image")
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<Button Icon="{StaticResource test_image}"' + XMLNS_FR + ' />')
        

    def test_ExpanderStackPanel(self):
        bkt.taskpane.TaskPaneControl.no_id = True
        bkt.taskpane.FluentRibbonControl.no_id = True
        
        b = bkt.taskpane.Expander(auto_stack=True, children=[bkt.taskpane.Wpf.Button()])
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<Expander'+XMLNS+'>\r\n  <StackPanel Orientation="Vertical">\r\n    <Button />\r\n  </StackPanel>\r\n</Expander>')

        b = bkt.taskpane.Expander(auto_wrap=True, children=[bkt.taskpane.Wpf.Button()])
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<Expander'+XMLNS+'>\r\n  <WrapPanel Orientation="Horizontal">\r\n    <Button />\r\n  </WrapPanel>\r\n</Expander>')
        
        b = bkt.taskpane.Expander(auto_stack=True, header="Test Header", children=[bkt.taskpane.Wpf.Button()])
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<Expander Header="Test Header"'+XMLNS+'>\r\n  <StackPanel Orientation="Vertical">\r\n    <Button />\r\n  </StackPanel>\r\n</Expander>')
    
    
    def test_Group(self):
        bkt.taskpane.TaskPaneControl.no_id = True
        bkt.taskpane.FluentRibbonControl.no_id = True
        self.maxDiff = None
        
        b = bkt.taskpane.Group(auto_wrap=True, children=[bkt.taskpane.Wpf.Button()])
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<StackPanel Orientation="Vertical"'+XMLNS+'>\r\n  <Grid Margin="0,10,0,5">\r\n    <Grid.ColumnDefinitions>\r\n      <ColumnDefinition Width="*" />\r\n    </Grid.ColumnDefinitions>\r\n    <Border Height="1" Background="{StaticResource BKTDivider}" HorizontalAlignment="Stretch" SnapsToDevicePixels="True" Margin="7,3,10,3" />\r\n  </Grid>\r\n  <WrapPanel Orientation="Horizontal">\r\n    <Button />\r\n  </WrapPanel>\r\n</StackPanel>')

        b = bkt.taskpane.Group(auto_wrap=True, show_separator=False, children=[bkt.taskpane.Wpf.Button()])
        self.assertEqual(bkt.xml.WpfXMLFactory.to_string(b.wpf_xml()), u'<StackPanel Orientation="Vertical"'+XMLNS+'>\r\n  <WrapPanel Orientation="Horizontal">\r\n    <Button />\r\n  </WrapPanel>\r\n</StackPanel>')

