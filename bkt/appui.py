# -*- coding: utf-8 -*-
'''
Provider of app-specific ribbon UIs

Created on 11.11.2019
@author: rdebeerst
'''

from __future__ import absolute_import

import logging

from collections import OrderedDict, defaultdict

import bkt.helpers as _h #used to access config and resources

import bkt.xml as mod_xml #creating customui xml
import bkt.ribbon as mod_ribbon #ribbon controls

import bkt.taskpane
import bkt.ui #bkt.ui is not required here, but if it is not loaded anywhere, its not available in feature folders. This happended to me when I remove dev module from config.


class CustomRibbonUI(object):
    '''
    Manage CustomUI-extensions for a ribbon.
    https://msdn.microsoft.com/en-us/library/dd926139(v=office.12).aspx
    '''
    def __init__(self, ribbon_id, short_id=None):
        # if ribbon_id in self.registry:
        #     raise ValueError('duplicate ribbon id %r' % ribbon_id)
        # self.registry[ribbon_id] = self
        self.ribbon_id = ribbon_id
        self.short_id = short_id
        self.tabs = OrderedDict()
        self.contextual_tabs = defaultdict(list)
        self.context_menus = {}
        self.commands = {}
        self.backstage_controls = []

        self.lazy_replacements = {}
        self.lazy_extensions = defaultdict(list)
    
    
    def add_tab(self, tab, extend=False):
        ''' add tab to ribbon-ui '''
        # self.tabs.append(tab)
        # modules that use annotations (dev and christoph) do not have an id
        tab_id = getattr(tab, "id", "bkt_tab_" + str(len(self.tabs)+1))

        if tab_id in self.tabs:
            if extend:
                self.add_groups_to_tab(tab_id, tab.children)
            else:
                raise ValueError('duplicate tab id %r' % tab_id)
        else:
            self.tabs[tab_id] = tab
            #logging.debug('added tab with id %r', tab_id)

        return tab_id

    def add_contextual_tab(self, id_mso, tab):
        ''' add contextual tab '''
        self.contextual_tabs[id_mso].append(tab)
        # if id_mso in self.contextual_tabs:
        #     self.contextual_tabs[id_mso].append(tab)
        # else:
        #     self.contextual_tabs[id_mso] = [tab]
        
    def __call__(self, cls):
        ''' act as a class decorator for tabs, e.g. @bkt.powerpoint ''' 
        self.add_tab(cls)
        return cls

    def add_groups_to_tab(self, tab_id, children):
        if not tab_id in self.tabs:
            raise ValueError('no tab with id %r' % tab_id)
        try:
            self.tabs[tab_id].children.extend(children)
            #self.tabs[tab_id].children += children
        except:
            raise ValueError('error adding groups to tab with id %r' % tab_id)
    
    def add_context_menu(self, menu):
        ''' add context menu to ribbon-ui '''
        if menu['id_mso'] is None:
            raise ValueError('context menus need idMso')
        
        if menu['id_mso'] in self.context_menus:
            # falls gleiche Id mehrmals auftaucht, wird einfach aneinandergehaegt
            for child in menu.children:
                self.context_menus[menu['id_mso']].children.append(child)
        else:
            self.context_menus[menu['id_mso']] = menu
    
    def add_repurposed_command(self, command):
        ''' add repurposed command to ribbon-ui '''
        if not isinstance(command, mod_ribbon.Command):
            raise ValueError("control must be type Command or repurpose_id must be defined")

        if command['id_mso'] is None:
            raise ValueError('repurposed command has no idMso')
        # if command['id_mso'] in self.commands:
        #     # FIXME: chain the commands
        #     pass
        else:
            self.commands[command['id_mso']] = command
    
    def add_backstage_control(self, control):
        ''' add control to backstage area '''
        self.backstage_controls.append(control)
    
    def add_lazy_replacement(self, id, control):
        ''' add control that replaces existing control during customui loading '''
        self.lazy_replacements[id] = control
    
    def add_lazy_extension(self, id, controls):
        ''' add list of controls that extend children of existing control during customui loading '''
        self.lazy_extensions[id].extend(controls)
    
    




class AppUI(object):
    '''
    Register UI-Extensions for an office app: ribbon, task pane.
    Instance also acts as an CustomRibbonUI-instance
    '''
    
    # registry = {}
    #
    # @classmethod
    # def get_app_ui_for_app(cls, app_name, *args, **kwargs):
    #     ''' create/return a 'singleton'-AppUI-instance for the given app-name '''
    #     if app_name in cls.registry:
    #         return cls.registry[app_name]
    #     else:
    #         return cls(app_name, *args, **kwargs)
    
    
    def __init__(self, ribbon_ids=[], short_ids=[]):
        ''' initialize AppUI for given app-name with CustomRibbonUIs as given by ribbon_ids '''
        
        # init CustomRibbonUIs
        self.custom_ribbon_uis = OrderedDict()
        for ribbon_id in ribbon_ids:
            self.custom_ribbon_uis[ribbon_id] = CustomRibbonUI(ribbon_id)
        self.default_custom_ribbon_id = ribbon_ids[0]
        
        # init other UI
        self.taskpane_controls = []
        
        self.base_taskpane_control = None
        self.base_customui_controls = {}
    
    
    def __getattr__(self, name):
        ''' fallback to give direct access to default ribbon ui, e.g. add_tab, add_context_menu, etc.  '''
        return getattr(self.custom_ribbon_uis[self.default_custom_ribbon_id], name)
    
    
    def __call__(self, cls):
        ''' act as a class decorator for tabs, e.g. @bkt.powerpoint '''
        self.custom_ribbon_uis[self.default_custom_ribbon_id].add_tab(cls)
        return cls

    
    
    def add_taskpane_control(self, control):
        ''' add control to task pane panel '''
        self.taskpane_controls.append(control)
    
    
    # =======================
    # = base control access =
    # =======================

    def get_customui_control(self, ribbon_id=None):
        ribbon_id = ribbon_id or self.default_custom_ribbon_id
        
        try:
            return self.base_customui_controls[ribbon_id]
        except KeyError:
            customui_control = self.create_customui_control(ribbon_id)
            self.base_customui_controls[ribbon_id] = customui_control
            return customui_control
        
    
    def get_taskpane_control(self):
        if self.base_taskpane_control:
            return self.base_taskpane_control
        else:
            taskpane_control = self.create_taskpane_control()
            self.base_taskpane_control = taskpane_control
            return taskpane_control
    
    def get_customui(self, ribbon_id=None):
        xml = mod_xml.RibbonXMLFactory.to_string(self.get_customui_control(ribbon_id=ribbon_id).xml())
        return ('<!-- ribbon_id: %s -->\r\n' % self.ribbon_id) + xml
    
    def get_taskpane_ui(self):
        taskpane_control = self.get_taskpane_control()
        if taskpane_control:
            return mod_xml.WpfXMLFactory.to_string(taskpane_control.wpf_xml())
        else:
            return None
    
    
    
    # ====================
    # = control creation =
    # ====================
    
    def create_control(self, element, ribbon_id=None):
        '''
        Create RibbonControl for element, where element is an annotated class or a RibbonControl-instance.
        In case element is an annotated class, the RibbonControl is created using ControlFactory.
        In case element is already an RibbonControl-instance, create_control is applied on its children.
        '''
        
        ribbon_id = ribbon_id or self.default_custom_ribbon_id
        
        if isinstance(element, mod_ribbon.RibbonControl):
            # lazy replacement and extension of controls
            element_id = element.id
            if element_id in self.custom_ribbon_uis[ribbon_id].lazy_replacements:
                element = self.custom_ribbon_uis[ribbon_id].lazy_replacements[element_id]
                logging.debug("create_control: element with id %s replaced by element with id %s", element_id, element.id)
            if element_id in self.custom_ribbon_uis[ribbon_id].lazy_extensions:
                element.children.extend( self.custom_ribbon_uis[ribbon_id].lazy_extensions[element_id] )
                logging.debug("create_control: element with id %s extended", element_id)
            
            element.children = [self.create_control(c, ribbon_id=ribbon_id) for c in element.children ]
            return element
        
        elif bkt.config.enable_legacy_syntax or False:
            from bkt.annotation import ContainerUsage #@deprecated
            from bkt.factory import ControlFactory #@deprecated
            
            if isinstance(element, ContainerUsage):
                logging.debug("create_control for ContainerUsage: %s", element.container)
                return ControlFactory(element.container, ribbon_info=None).create_control()
            
            else:
                logging.warning("FeatureContainer used where instance of ContainerUsage was expected: %s", element)
                return ControlFactory(element, ribbon_info=None).create_control()
        
        else:
            logging.warning("create_control for element %s skipped", element)
    
    
    
    
    def create_customui_control(self, ribbon_id=None):
        '''
        Creates the base-control (<customUI>) with Ribbon-Tabs, ContextMenus and Commands from the given ribbon_id.
        '''
        
        ribbon_id = ribbon_id or self.default_custom_ribbon_id
        
        # Context-Menus
        context_menus = []
        if len(self.custom_ribbon_uis[ribbon_id].context_menus) > 0:
            context_menus = [mod_ribbon.ContextMenus(
                children = [ self.create_control(c, ribbon_id=ribbon_id)  for c in self.custom_ribbon_uis[ribbon_id].context_menus.values()]
            )]
        
        # Commands
        commands = []
        if len(self.custom_ribbon_uis[ribbon_id].commands) > 0:
            commands = [mod_ribbon.CommandList(
                children = [self.create_control(c, ribbon_id=ribbon_id)  for c in self.custom_ribbon_uis[ribbon_id].commands.values()]
            )]
        
        # Contextual-Tabsets
        contextual_tabs = []
        # if len(self.custom_ribbon_uis[ribbon_id].contextual_tabsets) > 0:
        #     contextual_tabs = [mod_ribbon.ContextualTabs(
        #         children = [self.create_control(t, ribbon_id)  for  t in self.custom_ribbon_uis[ribbon_id].contextual_tabsets]
        #     )]
        if len(self.custom_ribbon_uis[ribbon_id].contextual_tabs) > 0:
            contextual_tabs = [mod_ribbon.ContextualTabs(
                children = [
                    mod_ribbon.TabSet(
                        id_mso=id_mso,
                        children=[
                            self.create_control(tab, ribbon_id)
                            for tab in tablist
                        ]
                    )
                    for id_mso, tablist in self.custom_ribbon_uis[ribbon_id].contextual_tabs.iteritems()
                ]
            )]
        
        # Quick Access Toolbar
        quick_access_toolbar = []
        # TODO: implement customization of quick_access_toolbar
        # quick_access_toolbar = [mod_ribbon.Qat(
        #     children = [
        #             # u.a für PowerPoint
        #             mod_ribbon.SharedControls(
        #                 children=[
        #                     mod_ribbon.Control(idQ='bkt-settings', image='settings')
        #                 ]
        #             ),
        #     ]
        # )]
        
        # Backstage Controls
        backstage_controls = []
        if len(self.custom_ribbon_uis[ribbon_id].backstage_controls) > 0:
            backstage_controls = [mod_ribbon.Backstage(
                children = [self.create_control(c, ribbon_id) for c in self.custom_ribbon_uis[ribbon_id].backstage_controls]
            )]
        
        # Tabs
        tabs = []
        if len(self.custom_ribbon_uis[ribbon_id].tabs) > 0:
            tabs = [
                mod_ribbon.Tabs(
                    children = [self.create_control(tab, ribbon_id)  for _, tab in self.custom_ribbon_uis[ribbon_id].tabs.iteritems()]
                )
            ]

        # Ribbon
        ribbon = mod_ribbon.Ribbon(start_from_scratch=False,
            children = quick_access_toolbar + tabs + contextual_tabs
        )

        
        
        # build-up CustomUI
        customUI = mod_ribbon.CustomUI(
            onLoad='PythonOnRibbonLoad', loadImage='PythonLoadImage',
            children = [ ribbon ] + backstage_controls + context_menus + commands
        )
        
        return customUI
    
    
    
    def create_taskpane_control(self):
        ''' create the master task pane control for the addin '''
        
        taskpane_control = None
        taskpane_controls = []
        if len(self.taskpane_controls) > 0:
            taskpane_controls = [self.create_control(c, ribbon_id=None)  for c in self.taskpane_controls]
        
            # ScrollViewer / StackPanel
            stack_panel = bkt.taskpane.Wpf.StackPanel(
                Margin="0",
                Orientation="Vertical",
                children = [
                    ctrl
                    for ctrl in taskpane_controls
                ]
            )
            image_resources = {
                image_name: _h.Resources.images.locate(image_name)
                for image_name in stack_panel.collect_image_resources()
            }
            logging.debug('image resources: %s', image_resources)
            taskpane_control = bkt.taskpane.BaseScrollViewer(
                image_resources = image_resources,
                children = [stack_panel]
            )
        
        return taskpane_control
    
    
    
    


class AppUIPowerPoint(AppUI):
    '''
    Register UI-Extensions for PowerPoint: ribbon, task pane, context dialogs
    '''
    def __init__(self, ribbon_ids=[], short_ids=[]):
        super(AppUIPowerPoint, self).__init__(ribbon_ids=ribbon_ids, short_ids=short_ids)
        
        from bkt.contextdialogs import ContextDialogs
        
        self.use_contextdialogs = not _h.config.ppt_use_contextdialogs is False
        self.context_dialogs = ContextDialogs()






class AppUIs(object):
    '''
    Provide access to specific AppUI-instances for office applications
    '''
    
    registry = {}
    
    app_ui_classes = {
        'Microsoft PowerPoint': AppUIPowerPoint
    }
    
    ribbon_ids = {
        'Microsoft PowerPoint': ['Microsoft.PowerPoint.Presentation', 'Microsoft.Mso.IMLayerUI'], #'Microsoft.Mso.IMLayerUI' is the overlay UI of contacts in the comments taskpane which calls get_custom_ui with this ribbon id
        'Microsoft Excel':      ['Microsoft.Excel.Workbook'],
        'Microsoft Visio':      ['Microsoft.Visio.Drawing'],
        'Microsoft Word':       ['Microsoft.Word.Document'],
        'Outlook':              [
            'Microsoft.Outlook.Explorer',
            'Microsoft.OMS.MMS.Compose',
            'Microsoft.OMS.MMS.Read',
            'Microsoft.OMS.SMS.Compose',
            'Microsoft.OMS.SMS.Read',
            'Microsoft.Outlook.Appointment',
            'Microsoft.Outlook.Contact',
            'Microsoft.Outlook.DistributionList',
            'Microsoft.Outlook.Journal',
            'Microsoft.Outlook.Mail.Compose',
            'Microsoft.Outlook.Mail.Read',
            'Microsoft.Outlook.MeetingRequest.Read',
            'Microsoft.Outlook.MeetingRequest.Send',
            'Microsoft.Outlook.Post.Compose',
            'Microsoft.Outlook.Post.Read',
            'Microsoft.Outlook.Report',
            'Microsoft.Outlook.Resend',
            'Microsoft.Outlook.Response.Compose',
            'Microsoft.Outlook.Response.CounterPropose',
            'Microsoft.Outlook.Response.Read',
            'Microsoft.Outlook.RSS',
            'Microsoft.Outlook.Sharing.Compose',
            'Microsoft.Outlook.Sharing.Read',
            'Microsoft.Outlook.Task'
        ]
    }
    
    
    
    # @property
    # @classmethod
    # def PowerPoint(cls):
    #     return cls.get_app_ui("Microsoft PowerPoint")


    @classmethod
    def get_app_ui(cls, app_name):
        try:
            return cls.registry[app_name]
        except KeyError:
            instance = cls.create_app_ui(app_name)
            cls.registry[app_name] = instance
            return instance
    
    
    @classmethod
    def create_app_ui(cls, app_name):
        ''' create AppUI-instance for given app name '''
        # get AppUI-subclass for app name
        app_ui_class = cls.app_ui_classes.get(app_name, AppUI)
        # create instance
        return app_ui_class(ribbon_ids = cls.ribbon_ids.get(app_name, ["Default"]))

        


excel        = AppUIs.get_app_ui('Microsoft Excel')
outlook      = AppUIs.get_app_ui('Outlook') #NOTE: Its not Microsoft Outlook!
powerpoint   = AppUIs.get_app_ui('Microsoft PowerPoint')
visio        = AppUIs.get_app_ui('Microsoft Visio')
word         = AppUIs.get_app_ui('Microsoft Word')

