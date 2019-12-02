# -*- coding: utf-8 -*-

import logging
import System
import traceback

# wpf basics
from . import dotnet
wpf = dotnet.import_wpf()

# for getting coordinates for rotated shapes
from bkt.library.algorithms import get_bounding_nodes

# for Primitives.Popup
from System.Windows import Controls

import importlib


BKT_CONTEXTDIALOG_TAGKEY = 'BKT_CONTEXTDIALOG'



class ContextDialog(object):
    '''
    Represents a single context-dialog.
    A context dialog is a window (popup-window), show in context of a specific selection
    (e.g. shape with specific tag).
    '''
    
    
    def __init__(self, id, module=None, window_class=None):
        ''' constructor '''
        self.id = id
        self.module_name = module
        self.module = None
        self.window_class = window_class
    
    def show_dialog_at_shape_position(self, shape, context):
        ''' create window for the context dialog at show it at shape's position '''
        wnd = self.create_window(context)
        return self.show_window_at_shape_position(wnd, context.app.ActiveWindow, shape)
        
    def show_window_at_shape_position(self, dialog_window, active_window, shape):
        ''' show the given window at shape's position '''
        if isinstance(dialog_window, Controls.Primitives.Popup):
            # this is a popup window, popup doesnt need scaling factor
            left, top = DialogHelpers.get_dialog_positon_from_shape(active_window, shape, consider_scaling=False) 
            # set position and show dialog
            dialog_window.PlacementRectangle = System.Windows.Rect(left, top, 1, 1)
            dialog_window.IsOpen = True
            # Popup is automatically a child window
            
        else:
            # normal window
            # set position
            # left, top = DialogHelpers.get_dialog_positon_from_shape(active_window, shape, consider_scaling=True)
            # dialog_window.Top=top
            # dialog_window.Left=left
            left, top = DialogHelpers.get_dialog_positon_from_shape(active_window, shape, consider_scaling=False)
            dialog_window.SetDevicePosition(left, top)
            # make dialog a child window
            # System.Windows.Interop.WindowInteropHelper(dialog_window).Owner = DialogHelpers.get_main_window_handle()
            # show as non-blocking dialog
            dialog_window.Show()
            # put focus back on office window
            # active_window.Activate() #NOTE: do not use this: BktWindow-Popup is not stealing focus anymore. if active_window has a modal dialog or messagebox, this causes problems.
            
        return dialog_window
            
        
    def create_window(self, context):
        ''' create window for the context dialog, without showing it '''
        logging.debug('ContextDialog.create_window')
        try:
            if self.window_class:
                return self.window_class(context)
            elif self.module_name:
                self.import_module()
                return self.module.create_window(context)
            
        except:
            logging.error(traceback.format_exc())
    
    # def show(self, parent_window_handle, left, top, context):
    #     ''' show the context dialog from the corresponding module '''
    #     logging.debug('ContextDialog.show')
    #     try:
    #         if self.window_class:
    #             return self.show_window(parent_window_handle, left, top, context)
    #         elif self.module_name:
    #             self.import_module()
    #             return self.module.show(parent_window_handle, left, top, context)
    #
    #     except:
    #         logging.error(traceback.format_exc())
    #
    # def show_window(self, parent_window_handle, left, top, context):
    #     ''' show the context dialog using the corresponding window-class '''
    #     logging.debug("ContextDialog.show_window")
    #     if not self.window_class:
    #         return
    #
    #     # create window class
    #     wnd = self.window_class(context)
    #
    #     if isinstance(wnd, Controls.Primitives.Popup):
    #         pass
    #     else:
    #         # make the window a child window
    #         System.Windows.Interop.WindowInteropHelper(wnd).Owner = parent_window_handle
    #         wnd.Top=top
    #         wnd.Left=left
    #         # show as non-blocking dialog
    #         wnd.Show()
    #
    #     return wnd
        
        
    
    def import_module(self):
        '''
        equivalent to: import <<module_name>>
        will not reload if module was already loaded
        '''
        if not self.module:
            logging.debug('ContextDialog.import_module importing %s' % self.module_name)
            #do an import equivalent to:  import <<module_name>>
            #FIXME: use importlib.import_module
            self.module = importlib.import_module(self.module_name)
            # self.module = __import__(self.module_name, globals(), locals(), [], -1)
        
        

class ContextDialogs(object):
    '''
    Register and manage several ContextDialog-instances.
    Provides methods to show or hide dialogs in different situations, considering:
      - left/right click
      - selection-type
      - shape-type
      - shape-tag
    Current implementation assumes PowerPoint-context
    '''
    
        
    def __init__(self):
        ''' constructor '''
        self.dialogs = {}
        self.active_dialog = None
        
        self.drag_started = False
        self.key_is_down  = False
        self.showing_dialog_for_shape = False
        
        self.addin = None #c-addin
        self.context = None
        
    def register(self, id, module):
        ''' register a context dialog '''
        logging.debug('ContextDialogs.register: id=%s' % id)
        self.dialogs[id] = ContextDialog(id,module)
    
    def register_dialog(self, context_dialog):
        ''' register a context dialog from context-dialog-object '''
        logging.debug('ContextDialogs.register_dialog: id=%s' % context_dialog.id)
        self.dialogs[context_dialog.id] = context_dialog
    
    def unregister(self, id):
        ''' unregister a context dialog '''
        logging.debug('ContextDialogs.unregister: id=%s' % id)
        try:
            del self.dialogs[id]
        except IndexError:
            pass

    def re_show_shape_dialogs(self):
        ''' re-show context dialogs for current context '''
        logging.debug('ContextDialogs.re_show_shape_dialogs')

        try:
            if not self.context or self.showing_dialog_for_shape:
                return #if context not defined or dialog already visible, skip re-show
            self.show_shape_dialog_for_selection(self.context.selection, self.context)
        except:
            logging.error(traceback.format_exc())


    def show_shape_dialog_for_selection(self, selection, context):
        ''' show a context dialog for selected shape if exactly one shape is selected '''
        logging.debug('ContextDialogs.show_shape_dialog_for_selection')
        
        try:
            #save addin from context to (un)hook mouse/key events
            if not self.addin:
                self.context = context
                self.addin = context.addin
            # selection type
            # 0 = ppSelectionNone
            # 1 = ppSelectionSlide
            # 2 = ppSelectionShape
            # 3 = ppSelectionText
            if not self.drag_started and selection.type == 2:
                shapes = list(iter(selection.ShapeRange))
            
                if len(shapes) == 1:
                    self.show_shape_dialog_for_shape(shapes[0], context)
                elif len(shapes) > 1:
                    self.show_shape_dialog_for_shapes(shapes, context)
                else:
                    self.close_active_dialog()
            
            else:
                self.close_active_dialog()
        except:
            logging.error(traceback.format_exc())
    
    
    def hide_on_window_deactivate(self):
        logging.debug('ContextDialogs.hide_on_window_deactivate')
        try:
            self.close_active_dialog()
        except:
            logging.error(traceback.format_exc())
    
    
    def show_shape_dialog_for_shapes(self, shapes, context):
        if shapes[0].Tags(BKT_CONTEXTDIALOG_TAGKEY) != "" and len({shape.Tags(BKT_CONTEXTDIALOG_TAGKEY) for shape in shapes}) == 1:
            # all shapes have same dialog
            self.show_shape_dialog_for_shape(shapes[-1], context)
        else:
            self.show_master_shape_dialog(shapes, context)
    
    
    def show_shape_dialog_for_shape(self, shape, context):
        ''' create and show a context dialog for the given shape, depending on the shape's settings '''
        logging.debug('ContextDialogs.show_shape_dialog_for_shape')
        
        try:
            ### close active dialog
            self.close_active_dialog()
        except:
            logging.error(traceback.format_exc())

        try:
            ### check shape tag and show suitable dialog
            logging.debug('ContextDialogs.show_shape_dialog_for_shape check tag')
            
            if shape.Tags(BKT_CONTEXTDIALOG_TAGKEY) == '':
                return
            elif not shape.Tags(BKT_CONTEXTDIALOG_TAGKEY) in self.dialogs:
                logging.warning('No dialog registerd for given key: %s' % shape.Tags(BKT_CONTEXTDIALOG_TAGKEY))
                return
            else:
                ctx_dialog = self.dialogs.get(shape.Tags(BKT_CONTEXTDIALOG_TAGKEY)) or None
            
            if not ctx_dialog:
                return
            
            self.active_dialog = ctx_dialog.show_dialog_at_shape_position(shape, context)
            # logging.debug('ContextDialogs.show_shape_dialog_for_shape reactivate window')
            # context.app.ActiveWindow.Activate()
            self.showing_dialog_for_shape = True
            DialogHelpers.hook_events(self)
            
        except:
            logging.error(traceback.format_exc())
    
    def show_master_shape_dialog(self, shapes, context):
        ''' create and show a context dialog for the given shape, depending on the shape's settings '''
        logging.debug('ContextDialogs.show_master_shape_dialog')
        
        try:
            ### close active dialog
            self.close_active_dialog()
        except:
            logging.error(traceback.format_exc())
            
        try:
            ### check shape tag and show suitable dialog
            logging.debug('ContextDialogs.show_master_shape_dialog check tag')
            
            ctx_dialog = self.dialogs.get("MASTER") or None
            
            if not ctx_dialog:
                return
            
            master_shape = ctx_dialog.get_master_shape(shapes)
            if not master_shape:
                return

            self.active_dialog = ctx_dialog.show_dialog_at_shape_position(master_shape, context)
            # logging.debug('ContextDialogs.show_master_shape_dialog reactivate window')
            # context.app.ActiveWindow.Activate()
            self.showing_dialog_for_shape = True
            DialogHelpers.hook_events(self)
            
        except:
            logging.error(traceback.format_exc())
    
    def close_active_dialog(self):
        ''' close the latest active dialog if it still exists '''
        logging.debug('ContextDialogs.close_active_dialog')
        try:
            if self.active_dialog:
                if isinstance(self.active_dialog, Controls.Primitives.Popup):
                    self.active_dialog.IsOpen = False
                else:
                    self.active_dialog.Close()
                self.active_dialog = None
                self.showing_dialog_for_shape = False
                if self.addin:
                    DialogHelpers.unhook_events(self)
            
        except:
            logging.error(traceback.format_exc())
        
    
    def mouse_down(self, sender, e):
        ''' object sender, MouseEventExtArgs e) '''
        logging.debug("ContextDialogs.mouse_down")
        if self.showing_dialog_for_shape and self.active_dialog:
            if not self.active_dialog.IsMouseOver:
                self.close_active_dialog()

    def mouse_up(self, sender, e):
        ''' object sender, MouseEventExtArgs e) '''
        logging.debug("ContextDialogs.mouse_up")
        if not self.drag_started:
            if self.context and DialogHelpers.coordinates_within_slideview_window(e.X, e.Y, self.context):
                self.re_show_shape_dialogs()

    # def mouse_move(self, sender, e):
    #     ''' object sender, MouseEventExtArgs e) '''
    #     if self.showing_dialog_for_shape:
    #         if self.drag_started:
    #             logging.debug("ContextDialogs.mouse_move/dragging")
    #             self.close_active_dialog() #FIXME: if you drag a rectangle to select multiple shapes, afterwars popup immediatly closes

    def mouse_drag_start(self, sender, e):
        logging.debug("ContextDialogs.mouse_drag_start")
        self.drag_started = True
        if self.showing_dialog_for_shape and self.active_dialog:
            self.close_active_dialog()

    def mouse_drag_end(self, sender, e):
        logging.debug("ContextDialogs.mouse_drag_end")
        self.drag_started = False

    def key_down(self, sender, e):
        ''' object sender, MouseEventExtArgs e) '''
        logging.debug("ContextDialogs.key_down")
        if self.showing_dialog_for_shape and not self.key_is_down and self.active_dialog:
            self.close_active_dialog()
        self.key_is_down = True
        
    def key_up(self, sender, e):
        ''' object sender, MouseEventExtArgs e) '''
        logging.debug("ContextDialogs.key_up")
        self.key_is_down = False
        # if self.showing_dialog_for_shape and self.active_dialog:
        #     self.close_active_dialog()



class DialogHelpers(object):
    hooked = False

    @classmethod
    def hook_events(cls, ctx_dialogs):
        if not cls.hooked:
            cls.hooked = True
        # ctx_dialogs.addin.HookEvents()

    @classmethod
    def unhook_events(cls, ctx_dialogs):
        cls.hooked = False
        # ctx_dialogs.addin.UnhookEvents()

    
    @classmethod
    def get_dialog_positon_from_shape(cls, active_window, shape, consider_scaling=True):
        ''' get position at which context dialog of given shape should be shown  '''
        logging.debug('DialogHelpers.get_dialog_positon_from_shape')
        try:
            #window = cls.get_window_from_shape(shape)
            #window = context.app.ActiveWindow
            
            # consider rotated shapes
            if shape.rotation != 0:
                nodes = get_bounding_nodes(shape)
                # shp_x = min( p[0] for p in nodes )
                # shp_y = max( p[1] for p in nodes )
                # use point at lowest corner
                # nodes.sort(key=lambda p: (-p[1], p[0]))
                p_index = min(range(4), key=lambda i: (-nodes[i][1], nodes[i][0])) #range(4)=range(len(nodes))
                shp_x, shp_y = nodes[p_index]
            else:
                shp_x, shp_y = shape.left, shape.top+shape.height
            
            if consider_scaling:
                scaling_factor = cls.dpi_scaling_factor()
                # offset -7 for shadow radius
                # FIXME: window shadow should be outside of client window, see
                # https://marcin.floryan.pl/blog/2010/08/wpf-drop-shadow-with-windows-dwm-api
                left = active_window.PointsToScreenPixelsX(shp_x - 7 ) / scaling_factor
                top  = active_window.PointsToScreenPixelsY(shp_y - 7 + 6) / scaling_factor
                # left = Forms.Control.MousePosition.X/ scaling_factor
                # top = Forms.Control.MousePosition.Y/ scaling_factor
            else:
                left = active_window.PointsToScreenPixelsX(shp_x)
                top  = active_window.PointsToScreenPixelsY(shp_y + 4)

            return left, top
        except:
            logging.error(traceback.format_exc())
    
    
    @staticmethod
    def get_window_from_shape(shape):
        ''' returns window of the shape's presentation '''
        try:
            return shape.parent.parent.Windows.item(1)
        except:
            logging.error(traceback.format_exc())
    
    @staticmethod
    def coordinates_within_slideview_window(x, y, context):
        ''' returns true if mouse coordinates are within slide '''
        try:
            #FIXME: would be better to check actual slide view window instead of slide, because this doesnt work with zoomed in slides
            active_window = context.app.ActiveWindow

            ### METHOD 1: calculate coordinates of slide and compare
            page_setup = context.app.ActivePresentation.PageSetup
            l,t = active_window.PointsToScreenPixelsX(0), active_window.PointsToScreenPixelsY(0)
            w,h = active_window.PointsToScreenPixelsX(page_setup.SlideWidth), active_window.PointsToScreenPixelsY(page_setup.SlideHeight)
            return x>l and y>t and x<l+w and y<t+h

            ### METHOD 2: try to find a shape under cursor
            #FIXME: shape selection frame is actually bigger than shape
            # x,y = active_window.PointsToScreenPixelsX(x), active_window.PointsToScreenPixelsY(y)
            # shape = active_window.RangeFromPoint(x,y)
            # return shape is not None
        except:
            return False


    # FIXME:
    # https://dzimchuk.net/best-way-to-get-dpi-value-in-wpf/
    @staticmethod
    def dpi_scaling_factor():
        ''' read dpi scaling factor from registry '''
        try:
            from bkt import dotnet
            Win32 = dotnet.import_win32()
            RegistryHive = Win32.RegistryHive
            RegistryView = Win32.RegistryView
            RegistryKey = Win32.RegistryKey
            hkcu = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Default)
            dpi = hkcu.OpenSubKey("Control Panel").OpenSubKey("Desktop").OpenSubKey("WindowMetrics").GetValue("AppliedDPI")
            return dpi/96.
        except:
            logging.error(traceback.format_exc())
            return 1.

            
    @staticmethod
    def get_main_window_handle():
        ''' returns main window hwnd handle of current process '''
        try:
            return System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle
        except:
            logging.error(traceback.format_exc())
    
    

