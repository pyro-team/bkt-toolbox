# -*- coding: utf-8 -*-
'''
Created on 29.08.2019

@author: fstallmann
'''



import bkt

# import toolbox modules with ui
from . import arrange
from . import info
from . import language
from . import text
from . import fontawesome
from . import shape_adjustments
from . import stateshapes
from . import slides
from . import shapes as mod_shapes
from . import shape_selection


class ToolboxUi(object):
    _instance = None

    default_settings = {
        "tab_name": "Toolbox", #Tab Name

        "clipboard_group": 1, #page no 1
        "slides_group": 1, #page no 1

        "default_group_font": 0, #page no 1
        "default_group_paragraph": 0, #page no 1
        "compact_font_group": 1, #off
        "compact_paragraph_group": 1, #off

        "shapes_group": 1, #page no 1
        "styles_group": 1, #page no 1

        "size_group": 1, #page no 1
        "pos_size_group": 2, #page no 2

        "arrange_mini_group": 0, #off
        "arrange_group": 1, #page no 1
        "arrange_adv_group": 2, #page no 2
        "arrange_adv_easy_group": 2, #page no 2
        "arrange_euclid_group": 2, #page no 2
        "arrange_dist_rota_group": 1, #page no 1

        "text_padding_group": 1, #page no 1
        "text_par_group": 1, #page no 1
        "text_parindent_group": 2, #page no 2

        "format_group": 2, #page no 2
        "split_group": 0, #off
        "language_group": 0, #off
        "adjustments_group": 1, #page no 1
        "stateshape_group": 2, #page no 2
        "iconsearch_group": 2, #page no 2
    }

    @classmethod
    def get_instance(cls):
        assert isinstance(cls._instance, cls), "Toolbox UI not created"
        return cls._instance

    def __init__(self, theme_settings, total_pages):
        assert ToolboxUi._instance is None, "Toolbox UI already created!"

        #default theme settings
        self.theme_settings = self.default_settings
        self.theme_settings.update(theme_settings)
        #personal theme settings
        self.toolboxui_settings = self.theme_settings.copy()
        self.toolboxui_settings.update(bkt.settings.get("toolboxui.settings", {}))

        self.toolboxui_pages = [list() for _ in range(total_pages+1)] #0=no page/ will be discarded

        ToolboxUi._instance = self
    
    def get_all_keys(self):
        return list(self.theme_settings.keys())

    def get_theme_setting(self, key):
        return self.theme_settings[key]
    
    def get_setting(self, key):
        return self.toolboxui_settings[key]

    def set_setting(self, key, value):
        settings = bkt.settings.get("toolboxui.settings", {})
        if value == self.theme_settings[key]:
            try:
                del settings[key]
                bkt.settings["toolboxui.settings"] = settings
            except:
                pass
        else:
            settings[key] = value
            bkt.settings["toolboxui.settings"] = settings
    
    def reset_to_defaults(self):
        try:
            del bkt.settings["toolboxui.settings"]
        except:
            pass



    ### get list of toolbox pages based on settings
    def render_pages(self):
        self._render_pages_defaults()

        ### optional components
        self._render_pages_fontpar()
        self._render_pages_shapestyle()
        self._render_pages_possize()
        self._render_pages_arrange()
        self._render_pages_text()
        self._render_pages_others()

        ### standard groups at the end
        self.toolboxui_pages[1].append(info.info_group)

    def _render_pages_defaults(self):
        ### standard definitions of page 1 and 2
        # self.toolboxui_pages[1].extend([
        #     shape_selection.clipboard_group,
        #     slides.slides_group,
        # ])
        self.toolboxui_pages[self.toolboxui_settings["clipboard_group"]].append(shape_selection.clipboard_group)
        self.toolboxui_pages[self.toolboxui_settings["slides_group"]].append(slides.slides_group)
    
    def _render_pages_fontpar(self):
        #position size
        self.toolboxui_pages[self.toolboxui_settings["default_group_font"]].append(bkt.mso.group.GroupFont)
        self.toolboxui_pages[self.toolboxui_settings["default_group_paragraph"]].append(bkt.mso.group.GroupParagraph)
        self.toolboxui_pages[self.toolboxui_settings["compact_font_group"]].append(text.compact_font_group)
        self.toolboxui_pages[self.toolboxui_settings["compact_paragraph_group"]].append(text.compact_paragraph_group)
    
    def _render_pages_shapestyle(self):
        #position size
        self.toolboxui_pages[self.toolboxui_settings["shapes_group"]].append(mod_shapes.shapes_group)
        self.toolboxui_pages[self.toolboxui_settings["styles_group"]].append(mod_shapes.styles_group)
    
    def _render_pages_possize(self):
        #position size
        self.toolboxui_pages[self.toolboxui_settings["size_group"]].append(mod_shapes.size_group)
        self.toolboxui_pages[self.toolboxui_settings["pos_size_group"]].append(mod_shapes.pos_size_group)

    def _render_pages_arrange(self):
        #arrangement/alignment
        self.toolboxui_pages[self.toolboxui_settings["arrange_group"]].append(arrange.arrange_group)
        self.toolboxui_pages[self.toolboxui_settings["arrange_mini_group"]].append(arrange.arrange_advanced_small_group)

        self.toolboxui_pages[self.toolboxui_settings["arrange_adv_group"]].append(arrange.arrange_advanced_group)
        self.toolboxui_pages[self.toolboxui_settings["arrange_adv_easy_group"]].append(arrange.arrange_adv_easy_group)

        self.toolboxui_pages[self.toolboxui_settings["arrange_euclid_group"]].append(arrange.euclid_angle_group)
        self.toolboxui_pages[self.toolboxui_settings["arrange_dist_rota_group"]].append(arrange.distance_rotation_group)

    def _render_pages_text(self):
        #text settings
        self.toolboxui_pages[self.toolboxui_settings["text_padding_group"]].append(text.innenabstand_gruppe)
        self.toolboxui_pages[self.toolboxui_settings["text_par_group"]].append(text.paragraph_group)
        self.toolboxui_pages[self.toolboxui_settings["text_parindent_group"]].append(text.paragraph_indent_group)

    def _render_pages_others(self):
        #others
        self.toolboxui_pages[self.toolboxui_settings["format_group"]].append(mod_shapes.format_group)
        self.toolboxui_pages[self.toolboxui_settings["split_group"]].append(mod_shapes.split_shapes_group)
        self.toolboxui_pages[self.toolboxui_settings["language_group"]].append(language.sprachen_gruppe)
        self.toolboxui_pages[self.toolboxui_settings["adjustments_group"]].append(shape_adjustments.adjustments_group)
        self.toolboxui_pages[self.toolboxui_settings["stateshape_group"]].append(stateshapes.stateshape_gruppe)
        self.toolboxui_pages[self.toolboxui_settings["iconsearch_group"]].append(fontawesome.fontsearch_gruppe)


    def render_contextmenus(self):
        # define context-menus and context-tabs
        from .contextmenus import common


    def get_page(self, index):
        return self.toolboxui_pages[index]
    
    def get_page_name(self, index):
        total = len(self.toolboxui_pages)-1
        return "{} {}/{}".format(self.toolboxui_settings["tab_name"], index, total)
    

    def show_settings_editor(self, context):
        from .dialogs.toolbox_ui import ToolboxUiWindow
        ToolboxUiWindow.create_and_show_dialog(self, context)
