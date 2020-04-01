# -*- coding: utf-8 -*-
'''
Created on 26.02.2020

@author: fstallmann
'''

from __future__ import absolute_import

import logging
import json
import io
import os.path

from collections import OrderedDict

import bkt
import modules.settings as settings



class DevGroup(object):
    log_level = None
    
    @staticmethod
    def show_console(context):
        import bkt.console as co
        co.console.Visible = True
        co.console.scroll_down()
        co.console.BringToFront()  # @UndefinedVariable
        co.console._globals['context'] = context

    @staticmethod
    def reload_bkt(context):
        settings.BKTReload.reload_bkt(context)
        # import bkt.console
        # try:
        #     addin = context.app.COMAddIns["BKT.AddIn"]
        #     addin.Connect = False
        #     addin.Connect = True
        # except Exception, e:
        #     bkt.console.show_message(str(e))
    
    @staticmethod
    def show_config(context):
        import bkt.console
        def _iter_lines():
            cfg = dict(context.config.items("BKT"))
            for k in sorted(cfg):
                yield "{:30} = {}".format(k, str(getattr(context.config, k)))
            yield ''
        
        bkt.console.show_message('\r\n'.join(_iter_lines()))
    
    @staticmethod
    def show_settings(context):
        import bkt.console
        def _iter_lines():
            for k in sorted(context.settings):
                yield "{:35} = {}".format(k, context.settings.get(k, "ERROR"))
            yield ''
        
        bkt.console.show_message('\r\n'.join(_iter_lines()))

    @staticmethod
    def show_ribbon_xml(python_addin, ribbon_id):
        import bkt.console
        bkt.console.show_message(python_addin.get_custom_ui(ribbon_id))
        
    @staticmethod
    def toggle_show_exception(pressed):
        bkt.config.set_smart("show_exception", pressed)
        
    @staticmethod
    def toggle_log_write_file(pressed):
        bkt.config.set_smart("log_write_file", pressed)
        
    @staticmethod
    def toggle_legacy_syntax(pressed):
        bkt.config.set_smart("enable_legacy_syntax", pressed)
    
    @staticmethod
    def change_log_level(pressed, current_control):
        bkt.config.set_smart("log_level", current_control["tag"])
    
    @classmethod
    def get_log_level(cls, current_control):
        if cls.log_level is None:
            logger = logging.getLogger()
            cls.log_level = logging.getLevelName(logger.level)
        return cls.log_level == current_control["tag"]


class AllControls(object):
    types_include = [bkt.ribbon.Button, bkt.ribbon.ToggleButton, bkt.ribbon.Gallery, bkt.ribbon.Menu, bkt.ribbon.DynamicMenu, bkt.ribbon.EditBox, bkt.ribbon.ComboBox, bkt.ribbon.SpinnerBox, bkt.ribbon.MSOControl]
    types_exclude = [bkt.ribbon.DialogBoxLauncher]

    types_haschildren = [bkt.ribbon.Menu, bkt.ribbon.SplitButton, bkt.ribbon.Box, bkt.ribbon.Gallery, bkt.ribbon.PrimaryItem, bkt.ribbon.MenuGroup]
    types_haslabel    = [bkt.ribbon.Menu, bkt.ribbon.Gallery, bkt.ribbon.DynamicMenu]

    cm_descriptions = {
        'ContextMenuSpell': 'Kontextmenü für rot unterstrichene Wörter',
        'ContextMenuShape': 'Kontextmenü bei Auswahl eines einzelnen Shapes',
        'ContextMenuTextEdit': 'Kontextmenü innerhalb eines Textfelds oder bei selektiertem Text',
        'ContextMenuFrame': 'Kontextmenü für leere Stelle auf der Folie',
        'ContextMenuPicture': 'Kontextmenü für Grafiken und Bilder',
        'ContextMenuObjectsGroup': 'Kontextmenü bei Auswahl mehrerer Shapes oder Objekte',
        'ContextMenuThumbnail': 'Kontextmenü für die Folien-Vorschau im rechten Panel',
        'ContextMenuShapeConnector': 'Kontextmenü für einen einzelnen Verbinder',
        'ContextMenuShapeFreeform': 'Kontextmenü für sog. Freeform-Shapes, also Shape mit beliebiger selbst erstellter Form'
    }

    def __init__(self, context):
        self.context      = context
        self.python_addin = context.python_addin
        self.all_controls = []
    
    def run(self):
        self.add_all_standard_tabs()
        self.add_all_contextual_tabs()
        self.add_all_backstage_controls()
        self.add_all_context_menus()

        self.write_json()
        self.write_markdown()

    def _getattr(self, object, *args, **kwargs):
        for arg in args:
            try:
                return object[arg]
            except:
                pass
        else:
            try:
                return kwargs["default"]
            except KeyError:
                return None

    def _get_control_dict(self, control, submenu=None):
        control_dict = OrderedDict()
        if isinstance(control, bkt.ribbon.ContextMenu):
            id_mso = self._getattr(control, "id_mso")
            control_dict['id']          = id_mso
            control_dict['image']       = id_mso
            control_dict['name']        = id_mso
            control_dict['description'] = self.cm_descriptions.get(id_mso)
            control_dict['is_standard'] = True

        elif isinstance(control, bkt.ribbon.MSOControl):
            id_mso = self._getattr(control, "id_mso")
            control_dict['id']          = id_mso
            control_dict['image']       = id_mso
            control_dict['name']        = self.context.app.commandbars.GetLabelMso(id_mso)
            control_dict['description'] = self.context.app.commandbars.GetSupertipMso(id_mso)
            control_dict['is_standard'] = True
        
        else:
            control_dict['id']          = self._getattr(control, "id")
            control_dict['image']       = self._getattr(control, "image", "image_mso")
            control_dict['name']        = self._getattr(control, 'label', 'screentip')
            control_dict['description'] = self._getattr(control, 'supertip', 'description')
            control_dict['is_standard'] = False
        
        control_dict['type']            = type(control).__name__
        
        if submenu is not None:
            control_dict['submenu']     = " > ".join(submenu)
        return control_dict
    
    def _add_group_child_control(self, list_obj, control, submenu):
        if any(isinstance(control, t) for t in self.types_include) and type(control) not in self.types_exclude:
            c_name = self._getattr(control, 'label', 'screentip')
            #skip controls where label AND screentip are not given, i.e. callback functions
            if c_name or isinstance(control, bkt.ribbon.MSOControl):
                list_obj.append( self._get_control_dict(control, submenu) )

        if any(isinstance(control, t) for t in self.types_haschildren):
            self._iterate_over_group_children(list_obj, control, submenu)
    
    def _iterate_over_group_children(self, list_obj, control, submenu):
        if any(isinstance(control, t) for t in self.types_haslabel):
            new_submenu = self._getattr(control, 'label', 'screentip')
            if not new_submenu:
                logging.warning("missing label for id "+self._getattr(control, 'id'))
            else:
                submenu = submenu + [new_submenu]
            
        if isinstance(control, bkt.ribbon.SpinnerBox):
            list_obj.append( self._get_control_dict(control.txt_box, submenu) )
            
            if control.image_element:
                self._add_group_child_control(list_obj, control.image_element, submenu)

        else:
            for child_control in control.children:
                self._add_group_child_control(list_obj, child_control, submenu)

    def add_all_standard_tabs(self):
        #standard tabs
        for tab_id, tab in self.python_addin.app_ui.tabs.iteritems():
            tab_label = self._getattr(tab, "label")
            #if no label is given try getting standard idmso label
            if tab_label is None:
                try:
                    tab_label = self.context.app.CommandBars.GetLabelMso(tab_id)
                except:
                    continue
            tab_control = OrderedDict()
            tab_control["id"]       = tab_id
            tab_control["type"]     = "tab"
            tab_control["name"]     = tab_label
            tab_control["children"] = []
            for group in tab.children:
                try:
                    group_control = self._get_control_dict(group)
                    group_control["children"] = []
                    tab_control["children"].append(group_control)
                    self._iterate_over_group_children(group_control["children"], group, [])
                except:
                    pass
            self.all_controls.append(tab_control)

    def add_all_contextual_tabs(self):
        #contextual_tabs
        for tab_id, tablist in self.python_addin.app_ui.contextual_tabs.iteritems():
            for tab in tablist:
                tab_control = OrderedDict()
                tab_control["id"]       = tab_id
                tab_control["type"]     = "contextual_tab"
                tab_control["name"]    = self.context.app.CommandBars.GetLabelMso(tab_id)
                tab_control["children"] = []
                for group in tab.children:
                    try:
                        group_control = self._get_control_dict(group)
                        group_control["children"] = []
                        tab_control["children"].append(group_control)
                        self._iterate_over_group_children(group_control["children"], group, [])
                    except:
                        bkt.helpers.exception_as_message()
                self.all_controls.append(tab_control)

    def add_all_backstage_controls(self):
        #backstage controls
        #NOTE: this solution does not cover the full possibilities of backstage, but it is sufficient as of now
        for tab in self.python_addin.app_ui.backstage_controls:
            tab_control = OrderedDict()
            tab_control["id"]       = self._getattr(tab, "id")
            tab_control["type"]     = "backstage"
            tab_control["name"]     = self._getattr(tab, "label")
            tab_control["children"] = []
            for cols in tab.children:
                for group in cols.children:
                    try:
                        group_control = self._get_control_dict(group)
                        group_control["children"] = []
                        tab_control["children"].append(group_control)
                        self._iterate_over_group_children(group_control["children"], group, [])
                    except:
                        pass
            self.all_controls.append(tab_control)
    
    def add_all_context_menus(self):
        context_menus = OrderedDict()
        context_menus["id"]       = "ContextMenu"
        context_menus["type"]     = "context_menu"
        context_menus["name"]     = "Kontextmenüs"
        context_menus["children"] = []
        #context menu controls
        for _, contextmenu in self.python_addin.app_ui.context_menus.iteritems():
            try:
                cm_control = self._get_control_dict(contextmenu)
                cm_control["children"] = []
                context_menus["children"].append(cm_control)
                self._iterate_over_group_children(cm_control["children"], contextmenu, [])
            except:
                pass
        self.all_controls.append(context_menus)

    def write_json(self):
        file = os.path.join(os.path.dirname(__file__), "all_controls.json")
        with io.open(file, 'w', encoding='utf-8') as json_file:
            # bkt.console.show_message(json.dumps(all_controls, ensure_ascii=False))
            json.dump(self.all_controls, json_file, ensure_ascii=False, indent=2)
    
    def write_markdown(self):
        file = os.path.join(os.path.dirname(__file__), "all_controls.md")
        with io.open(file, 'w', encoding='utf-8') as md_file:
            for parent in self.all_controls:
                if len(parent["children"]) == 0:
                    #tab without groups, e.g. FormatTab (contextual tab which only has visible callback but no children)
                    continue
                md_file.write("## {}\n\n".format(parent["name"]))
                for group in parent["children"]:
                    md_file.write('### {name}\n\n'.format(**group))
                    if group["description"]:
                        md_file.write('{description}\n\n'.format(**group))
                    if parent["type"] != "context_menu":
                        md_file.write('<img src="documentation/groups/{id}.png">\n\n'.format(**group))
                    if len(group["children"]) > 0:
                        md_file.write("| {:50} | {:50} |\n".format("Name", "Beschreibung"))
                        md_file.write("| {:-<50} | {:-<50} |\n".format("-", "-"))
                        for control in group["children"]:
                            if not control["name"] and not control["description"]:
                                continue
                            if control["description"]:
                                description = control["description"].replace("\n", "<br>")
                            else:
                                description = ""
                            name = control["name"]
                            if control["is_standard"]:
                                name = u"*{}*".format(name)
                            if control["submenu"]:
                                name = u"{} > {}".format(control["submenu"], name)
                            md_file.write(u"| {:50} | {:50} |\n".format(name, description))
                        md_file.write("\n")
                    md_file.write("\n")
                md_file.write("\n\n\n")
    
    @staticmethod
    def generate_overview(context):
        controls = AllControls(context)
        controls.run()
        bkt.helpers.message("Files successfully created")



common_group = bkt.ribbon.Group(
    id="bkt_common_dev_group",
    label="Common development tools",
    image_mso="ControlsGallery",
    children=[
        bkt.ribbon.Button(
            label="Interactive Console",
            size="large",
            image_mso="WatchWindow",
            on_action=bkt.Callback(DevGroup.show_console, context=True, transaction=False),
        ),
        bkt.ribbon.Button(
            label="Reload BKT",
            size="large",
            image_mso="AccessRefreshAllLists",
            on_action=bkt.Callback(DevGroup.reload_bkt, context=True, transaction=False),
        ),
        bkt.ribbon.Separator(),
        bkt.ribbon.Menu(
            label="Config",
            size="large",
            image_mso="WebPageComponent",
            children=[
                bkt.ribbon.Button(
                    label="Open BKT folder",
                    image_mso="Folder",
                    on_action=bkt.Callback(settings.BKTInfos.open_folder, transaction=False)
                ),
                bkt.ribbon.Button(
                    label="Open Cache folder",
                    image_mso="Folder",
                    on_action=bkt.Callback(lambda: settings.BKTInfos.open_folder(bkt.helpers.get_cache_folder()), transaction=False)
                ),
                bkt.ribbon.Button(
                    label="Open Settings folder",
                    image_mso="Folder",
                    on_action=bkt.Callback(lambda: settings.BKTInfos.open_folder(bkt.helpers.get_settings_folder()), transaction=False)
                ),
                bkt.ribbon.Button(
                    label="Open Favorites folder",
                    image_mso="Folder",
                    on_action=bkt.Callback(lambda: settings.BKTInfos.open_folder(bkt.helpers.get_fav_folder()), transaction=False)
                ),
                bkt.ribbon.MenuSeparator(),
                bkt.ribbon.Button(
                    label="Open config.txt",
                    image_mso="NewNotepadTool",
                    on_action=bkt.Callback(settings.BKTInfos.open_config, transaction=False)
                ),
                bkt.ribbon.Button(
                    label="Show config",
                    image_mso="WebPageComponent",
                    on_action=bkt.Callback(DevGroup.show_config, context=True, transaction=False),
                ),
                bkt.ribbon.MenuSeparator(),
                bkt.ribbon.Button(
                    label="Show app settings",
                    image_mso="WebPartProperties",
                    on_action=bkt.Callback(DevGroup.show_settings, context=True, transaction=False),
                ),
                bkt.ribbon.MenuSeparator(),
                bkt.ribbon.ToggleButton(
                    label="Show exceptions on/off",
                    get_pressed=bkt.Callback(lambda: bkt.config.show_exception or False),
                    on_toggle_action=bkt.Callback(DevGroup.toggle_show_exception, transaction=False)
                ),
                bkt.ribbon.ToggleButton(
                    label="Write log file on/off",
                    get_pressed=bkt.Callback(lambda: bkt.config.log_write_file or False),
                    on_toggle_action=bkt.Callback(DevGroup.toggle_log_write_file, transaction=False)
                ),
                bkt.ribbon.ToggleButton(
                    label="Legacy Syntax on/off",
                    get_pressed=bkt.Callback(lambda: bkt.config.enable_legacy_syntax or False),
                    on_toggle_action=bkt.Callback(DevGroup.toggle_legacy_syntax, transaction=False)
                ),
                bkt.ribbon.Menu(
                    label="Change log-level",
                    children=[
                        bkt.ribbon.ToggleButton(
                            label="DEBUG",
                            tag="DEBUG",
                            get_pressed=bkt.Callback(DevGroup.get_log_level, current_control=True),
                            on_toggle_action=bkt.Callback(DevGroup.change_log_level, current_control=True, transaction=False)
                        ),
                        bkt.ribbon.ToggleButton(
                            label="INFO",
                            tag="INFO",
                            get_pressed=bkt.Callback(DevGroup.get_log_level, current_control=True),
                            on_toggle_action=bkt.Callback(DevGroup.change_log_level, current_control=True, transaction=False)
                        ),
                        bkt.ribbon.ToggleButton(
                            label="WARNING",
                            tag="WARNING",
                            get_pressed=bkt.Callback(DevGroup.get_log_level, current_control=True),
                            on_toggle_action=bkt.Callback(DevGroup.change_log_level, current_control=True, transaction=False)
                        ),
                        bkt.ribbon.ToggleButton(
                            label="ERROR",
                            tag="ERROR",
                            get_pressed=bkt.Callback(DevGroup.get_log_level, current_control=True),
                            on_toggle_action=bkt.Callback(DevGroup.change_log_level, current_control=True, transaction=False)
                        ),
                        bkt.ribbon.ToggleButton(
                            label="CRITICAL",
                            tag="CRITICAL",
                            get_pressed=bkt.Callback(DevGroup.get_log_level, current_control=True),
                            on_toggle_action=bkt.Callback(DevGroup.change_log_level, current_control=True, transaction=False)
                        ),
                    ]
                )
            ]
        ),
        bkt.ribbon.Button(
            label="Show Ribbon XML",
            size="large",
            image="xml",
            on_action=bkt.Callback(DevGroup.show_ribbon_xml, python_addin=True, ribbon_id=True, transaction=False),
        ),
        bkt.ribbon.Button(
            label="Generate overview",
            size="large",
            image_mso="CreateMap",
            on_action=bkt.Callback(AllControls.generate_overview, context=True, transaction=False),
        ),
        #TODO: create new feature folder, clear all caches
        #ICONS: ControlsPane
    ]
)


class ImageMso(object):
    search_limit = 100

    search_engine = None
    search_term = ""
    search_results = None

    # re1 = re.compile(r'(.)([A-Z][a-z]+)')
    # re2 = re.compile(r'([a-z0-9])([A-Z])')

    # @classmethod
    # def camel_to_keywords(cls, name):
    #     name = name.replace("_", "")
    #     name = re.sub(cls.re1, r'\1 \2', name)
    #     return re.sub(cls.re2, r'\1 \2', name).lower()

    @classmethod
    def load_json(cls, search_engine):
        import os.path
        import io
        import json

        search_writer = search_engine.writer()
        file = os.path.join(os.path.dirname(os.path.realpath(__file__)), "imagemso2.json")
        with io.open(file, 'r', encoding='utf-8') as json_file:
            icons = json.load(json_file, object_pairs_hook=OrderedDict)
            for icon, keywords in icons.iteritems():
                # keywords = cls.camel_to_keywords(icon)
                search_writer.add_document(module="imagemso", name=icon, keywords=keywords)
            search_writer.commit()

        # file2 = os.path.join(os.path.dirname(os.path.realpath(__file__)), "imagemso2.json")
        # with io.open(file2, 'w', encoding='utf-8') as json_file:
        #     json.dump(new_imagemso, json_file, ensure_ascii=False, indent=2)

    @classmethod
    def get_search_engine(cls):
        if cls.search_engine:
            return cls.search_engine

        from bkt.search import get_search_engine
        cls.search_engine = get_search_engine("imagemso")
        # initialize search index on first use of engine
        cls.load_json(cls.search_engine)
        return cls.search_engine
    
    @classmethod
    def _perform_search(cls):
        if len(cls.search_term) > 0:
            cls.search_results = None
            engine = cls.get_search_engine()
            with engine.searcher() as searcher:
                    cls.search_results = searcher.search(cls.search_term)
        else:
            cls.search_results = None

    @classmethod
    def set_search_term(cls, value):
        cls.search_term = value
        cls._perform_search()

    @classmethod
    def get_search_term(cls):
        return cls.search_term
    
    @classmethod
    def get_symbol_galleries(cls):
        if not cls.search_results or len(cls.search_results) == 0:
            image_msos = [
                bkt.ribbon.Button(
                    label="Keine Ergebnisse für »{}«".format(cls.search_term),
                    enabled=False
                )
            ]
        
        else:
            image_msos = [
                bkt.ribbon.MenuSeparator(title="{} Ergebnisse für »{}«".format(len(cls.search_results), cls.search_term))
            ]
            if len(cls.search_results) > cls.search_limit:
                image_msos.append(
                    bkt.ribbon.Button(label="Zeige die ersten {} Ergebnisse".format(cls.search_limit), enabled=False)
                )
            for icon in cls.search_results.limit(cls.search_limit):
                image_msos.append(
                    bkt.ribbon.Button(
                        label=icon.name,
                        supertip=icon.keywords,
                        image_mso=icon.name,
                        on_action=bkt.Callback(cls.copy_to_clipboard)
                    )
                )
        
        return bkt.ribbon.Menu(
                    xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                    id=None,
                    children=image_msos
                )

    @classmethod
    def get_enabled_results(cls):
        return cls.search_results is not None
    
    @classmethod
    def get_results_label(cls):
        if cls.search_results is not None:
            return "{} Icons".format(len(cls.search_results))
        else:
            return "Ergebnis"

    @classmethod
    def copy_to_clipboard(cls, current_control):
        import bkt.dotnet as dotnet
        Forms = dotnet.import_forms() #required to read clipboard
        Forms.Clipboard.SetText( current_control['label'] )



iconsearch_group = bkt.ribbon.Group(
    id="bkt_common_dev_iconsearch_group",
    label="ImageMso",
    image_mso='SearchUI',
    children=[
        bkt.ribbon.Label(
            label="Suchwort:",
        ),
        bkt.ribbon.EditBox(
            label="Suchwort",
            show_label=False,
            sizeString = '#########',
            get_text = bkt.Callback(ImageMso.get_search_term),
            on_change = bkt.Callback(ImageMso.set_search_term),
            supertip="Suchwort eingeben und ENTER klicken",
        ),
        bkt.ribbon.DynamicMenu(
            get_label=bkt.Callback(ImageMso.get_results_label),
            get_content=bkt.Callback(ImageMso.get_symbol_galleries),
            get_enabled=bkt.Callback(ImageMso.get_enabled_results),
            screentip="Suchergebnisse",
            supertip="Zeigt die Suchergebnisse der Icon-Suche nach dem gewünschten Suchwort an",
        ),
    ]
)


common_groups = [common_group, iconsearch_group]