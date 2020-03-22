# -*- coding: utf-8 -*-
'''
Handling of app-specific events, pre-defined callbacks and resources

Created on 11.07.2014
@authors: cschmitt, rdebeerst
'''

from __future__ import absolute_import

import traceback
import logging

import shelve #for resources cache
import os.path #for resources cache
import time #for cache invalidation and event throttling

import bkt.helpers as _h
from bkt.context import InappropriateContextError
from bkt.callbacks import Callback, CallbackTypes




class AppEventType(object):
    def __init__(self, event_name=None, custom=False, office=False):
        self.event_name = event_name
        self.custom = custom
        self.office = office
        self.registered_methods = []

    def set_attribute(self, attr):
        self.event_name = attr

    def register(self, method):
        logging.debug("Method registered for event: %r" % self.event_name)

        if isinstance(method, Callback) and method.callback_type is None:
            method.callback_type = CallbackTypes.bkt_event
        
        self.registered_methods.append(method)
        return self

    def unregister(self, method):
        logging.debug("Method unregistered for event: %r" % self.event_name)
        self.registered_methods.remove(method)
        return self
    
    # def fire(self):
    #     logging.debug("Fired event: %r" % self.event_name)
    #     for method in self.registered_methods:
    #         method()

    def __iter__(self):
        for method in self.registered_methods:
            yield method
    
    def __repr__(self):
        return "<AppEventType name=%s methods=%s>" % (self.event_name, len(self.registered_methods))

    __iadd__ = register
    __isub__ = unregister
    # __call__ = fire


class AppEventsRegister(object):
    ''' Static class to handle registration of methods to BKT and Office events '''

    def __init__(self):
        self._event_types = {}

    def __getattr__(self, attr):
        try:
            return self._event_types[attr]
        except:
            # define custom event
            custom = AppEventType(custom=True,event_name=attr)
            # save custom event-type. On the second access the same object will be returned
            self._event_types[attr] = custom
            return custom
    
    def __setattr__(self, attr, value):
        if isinstance(value, AppEventType):
            value.set_attribute(attr)
            self._event_types[attr] = value
        else:
            super(AppEventsRegister, self).__setattr__(attr, value)


AppEvents = AppEventsRegister()

#BKT events
AppEvents.bkt_load       = AppEventType()
AppEvents.bkt_unload     = AppEventType()
AppEvents.bkt_invalidate = AppEventType()

#Generic office events
AppEvents.window_activate   = AppEventType(office=True) #keyword args: ppt: presentation, window; xls: workbook, windows, visio: window
AppEvents.window_deactivate = AppEventType(office=True) #keyword args: ppt: presentation, window; xls: workbook, windows, visio: NOT AVAILABLE
AppEvents.selection_changed = AppEventType(office=True) #keyword args: - #NOTE: sheet selection in excel, window selection in ppt, selection or cell in visio

#PPT-specific events
AppEvents.slide_selection_changed  = AppEventType(office=True) #keyword args: -
AppEvents.after_shapesize_changed  = AppEventType(office=True) #keyword args: -
AppEvents.after_new_presentation   = AppEventType(office=True) #keyword args: presentation
AppEvents.after_presentation_open  = AppEventType(office=True) #keyword args: presentation
AppEvents.presentation_close       = AppEventType(office=True) #keyword args: presentation



class AppCallbacks(object):
    ''' Abstract class of application-level callbacks used in addin '''
        
    def invoke_callback(self, callback, *args, **kwargs):
        pass
    
    def undo_start(self, callback):
        pass
    
    def undo_end(self, callback):
        pass
    
    def bind_app_events(self):
        pass
        
    def unbind_app_events(self):
        pass
    



class AppCallbacksBase(AppCallbacks):
    
    def __init__(self, addin, app_ui, appcontext, appevents):
        self.addin = addin
        self.app_ui = app_ui
        self.context = appcontext
        self.events = appevents
        
        self.cache = {}
        # cache_cb_types = [bkt.callbacks.CallbackTypes.get_visible, bkt.callbacks.CallbackTypes.get_enabled]
        self.cache_timeout = 0.5
        self.cache_last_refresh = 0
    
    def destroy(self):
        self.unbind_app_events()
        
        self.addin = None
        self.app_ui = None
        self.context = None
        self.events = None
        self.cache = {}


    # ==========
    # = events =
    # ==========

    def fire_event(self, event, **kwargs):
        logging.debug("Event triggered: %r" % event)
        for method in event:
            try:
                if isinstance(method, Callback):
                    self.invoke_callback(self.context, method, **kwargs)
                else:
                    method(**kwargs)
            except:
                logging.error("Error triggering event method")
                logging.error(traceback.format_exc())



    # ================
    # = invalidation =
    # ================
    
    def invalidate(self):
        #clearing of caches is done in addin.invalidate_ribbon()
        #do invalidation
        self.addin.invalidate_ribbon()
    
    def refresh_cache(self, force=False):
        #global cache timeout will prevent that manual invalidates are not working properly
        if force or time.time() - self.cache_last_refresh > self.cache_timeout:
            self.cache = {}
            self.cache_last_refresh = time.time()
            return True
        return False
    
    
    # =======================
    # = callback invocation =
    # =======================
    
    def invoke_callback(self, context, callback, *args, **kwargs):
        #logging.debug("AppCallbacksBase.invoke_callback")
        #kwargs = {}
        do_cache = False
        # if callback.callback_type in self.cache_cb_types and not self.refresh_cache():
        if callback.callback_type.cacheable and callback.invocation_context.cache and not self.refresh_cache():
            # cache_key = repr([callback.method.__name__] + kwargs.keys()) #TESTME: add invocation context to key?
            # cache_key = callback.method.__name__ #only method name not sufficient if same name is used in different classes
            cache_key = str(callback.method) #TESTME: is method string representation sufficient as key? add callback type?
            do_cache = True
            try:
                logging.debug("trying cache for %r" % cache_key)
                return self.cache[cache_key]
                # if time.time() - self.cache[cache_key][1] < self.cache_timeout:
                #     return self.cache[cache_key][0]
            except KeyError:
                logging.debug("no cache for %r" % cache_key)

        for i, arg in enumerate(args):
            kwargs[callback.callback_type.pos_args[i]] = arg
        
        if callback.invocation_context is not None:
            try:
                ctx_args = context.resolve_arguments(callback.invocation_context)
                kwargs.update(ctx_args)
            except InappropriateContextError:
                #traceback.print_exc()
                logging.debug("InappropriateContextError")
                #TESTME: also cache InappropriateContextError, so return value "None"
                if do_cache:
                    self.cache[cache_key] = None
                return
        
        self.undo_start(callback)
        logging.debug("AppCallbacksBase.invoke_callback: run callback method\nkwargs=%s" % kwargs)
        return_value = callback.method(**kwargs)
        self.undo_end(callback)
        
        if do_cache:
            self.cache[cache_key] = return_value
            # self.cache[cache_key] = (return_value, time.time())
        
        return return_value
    



class AppCallbacksExcel(AppCallbacksBase):

    # ============================
    # = general callbacks/events =
    # ============================
    
    # def undo_start(self, callback):
    #     self.context.app.Interactive = False
    #     self.context.app.DisplayAlerts = False
    #     self.context.app.ScreenUpdating = False

    # def undo_end(self, callback):
    #     self.context.app.Interactive = True
    #     self.context.app.DisplayAlerts = True
    #     self.context.app.ScreenUpdating = True
    
    
    # ==============================
    # = binding application events =
    # ==============================
    
    def bind_app_events(self):
        # def dump(*args):
        #     print args
        
        if self.context is None:
            print('WARNING: no addin context available, no application events registered')
            return

        app = self.context.app
        if app is None:
            # for tests without application object
            print('WARNING: no application available, no application events registered')
            return

        #set application for various excel helper functions
        import bkt.library.excel.helpers as xllib
        xllib.set_application(app)

        app.SheetSelectionChange += self.sheet_selection_changed
        app.WindowActivate += self.window_activate
        app.WindowDeactivate += self.window_deactivate
    
    
    def unbind_app_events(self):
        from System.Runtime.InteropServices.Marshal import ReleaseComObject
        #import bkt.ui
        #bkt.console.show_message("%s.on_destroy()" % type(self).__name__)

        if self.context:
            app = self.context.app
            if app is None:
                # for tests without application object
                return
            app.SheetSelectionChange -= self.sheet_selection_changed
            app.WindowActivate -= self.window_activate
            app.WindowDeactivate -= self.window_deactivate
            #logging.debug("app events deregistered")

            #Very important line: will fix the problem that Excel process keeps running after closing the app (only in async mode)
            ReleaseComObject(app)

        #self.addin = None
    
    
    # ==================================
    # = application-specific callbacks =
    # ==================================
    
    def sheet_selection_changed(self, sheet, target):
        logging.debug("app event sheet_selection_changed")

        self.fire_event(self.events.selection_changed)
        self.invalidate()

    def window_activate(self, workbook, window):
        logging.debug("app event window_activate")

        self.fire_event(self.events.window_activate, workbook=workbook, window=window)
        self.invalidate()

    def window_deactivate(self, workbook, window):
        logging.debug("app event window_deactivate")

        self.fire_event(self.events.window_deactivate, workbook=workbook, window=window)
        self.invalidate()




class AppCallbacksVisio(AppCallbacksBase):
    ''' Handler for Visio application-level callbacks '''
    
    def __init__(self, *args, **kwargs):
        super(AppCallbacksVisio, self).__init__(*args, **kwargs)
        import bkt.library.visio as mod_visio
        self.mod_visio = mod_visio
        self.scope_id = 0
    
    # ============================
    # = general callbacks/events =
    # ============================
    
    def undo_start(self, callback):
        # print callback.method.__name__ + " / " + str(callback.callback_type.transactional)
        #NOTE: duplicate open of beginscope will completely disable undo in visio, therefore need to check scope_id=0!
        if self.scope_id == 0 and callback.callback_type.transactional and callback.method.__name__ not in ["undo", "redo", "reload_bkt"]:
            self.scope_id = self.context.app.BeginUndoScope("BKT Operation")
        # logging.warning(self.scope_id)

    def undo_end(self, callback):
        if self.scope_id > 0 and callback.callback_type.transactional and self.context:
            self.context.app.EndUndoScope(self.scope_id, True)
            self.scope_id = 0
        # logging.warning(self.scope_id)
    
    
    # ==============================
    # = binding application events =
    # ==============================

    def bind_app_events(self):
        
        if self.context is None:
            print('WARNING: no addin context available, no application events registered')
            return

        app = self.context.app
        if app is None:
            # for tests without application object
            print('WARNING: no application available, no application events registered')
            return
        
        app.SelectionChanged += self.selection_changed
        app.CellChanged += self.cell_changed
        app.WindowActivated += self.window_activated


    def unbind_app_events(self):
        self.context.app.SelectionChanged -= self.selection_changed
        self.context.app.CellChanged -= self.cell_changed
        self.context.app.WindowActivated -= self.window_activated

    def selection_changed(self, window):
        logging.debug("app event selection_changed")

        self.fire_event(self.events.selection_changed)
        self.invalidate()

    def cell_changed(self, cell):
        logging.debug("app event cell_changed")

        self.fire_event(self.events.selection_changed)
        self.invalidate()

    def window_activated(self, window):
        logging.debug("app event window_activated")

        self.fire_event(self.events.window_activate, window=window)
        self.invalidate()


class AppCallbacksPowerPoint(AppCallbacksBase):
    ''' Handler for PowerPoint application-level callbacks '''
    
    # ============================
    # = general callbacks/events =
    # ============================
    
    def undo_start(self, callback):
        if callback.callback_type.transactional:
            self.context.app.StartNewUndoEntry()
            # self.context.app.ActiveWindow.Presentation.SetUndoText("BKT Operation")
    
    
    def invalidate(self):
        try:
            self.app_ui.context_dialogs.close_active_dialog()
        except:
            logging.error(traceback.format_exc())

        super(AppCallbacksPowerPoint, self).invalidate()

    # ==============================
    # = binding application events =
    # ==============================
    
    def bind_app_events(self):
        # def dump(*args):
        #     print args
        
        if self.context is None:
            print('WARNING: no addin context available, no application events registered')
            return
        
        app = self.context.app
        if app is None:
            # for tests without application object
            print('WARNING: no application available, no application events registered')
            return
        
        # register invalidation callbacks
        app.WindowActivate += self.window_activate
        #app.WindowActivate += self.invalidate
        #app.WindowActivate += dump

        app.WindowDeactivate += self.window_deactivate
        #app.WindowDeactivate += self.invalidate
        #app.WindowDeactivate += dump
        
        app.SlideSelectionChanged += self.slide_selection_changed
        #app.SlideSelectionChanged += self.invalidate
        #app.SlideSelectionChanged += dump
        
        # app.WindowSelectionChange += self.window_selection_changed
        #app.WindowSelectionChange += self.invalidate
        #app.WindowSelectionChange += dump
        
        #event available in PPT2013
        if float(app.Version) >= 15.0:
            app.AfterShapeSizeChange += self.after_shape_size_changed
            #app.AfterShapeSizeChange += self.invalidate
            #app.AfterShapeSizeChange += dump
        
    
        app.AfterPresentationOpen  += self.after_presentation_open
        app.AfterNewPresentation   += self.after_new_presentation
        app.PresentationCloseFinal += self.presentation_close
        
        #print 'PPT events registered'
        
        
    def unbind_app_events(self):
        # from System.Runtime.InteropServices.Marshal import ReleaseComObject

        if self.context:
            app = self.context.app
            if app is None:
                # for tests without application object
                return
            app.SlideSelectionChanged -= self.slide_selection_changed 
            app.WindowActivate -= self.window_activate 
            # app.WindowSelectionChange -= self.window_selection_changed
            #event available in PPT2013
            if float(app.Version) >= 15.0:
                app.AfterShapeSizeChange -= self.after_shape_size_changed
    
            app.AfterPresentationOpen  -= self.after_presentation_open
            app.AfterNewPresentation   -= self.after_new_presentation
            app.PresentationCloseFinal -= self.presentation_close
            # NOTE: ReleaseComObject is necessary for Excel, but in Powerpoint it seems to lead to crashes after Powerpoint is closed
            # ReleaseComObject(app)

        if self.app_ui:
            if self.app_ui.context_dialogs:
                logging.debug("AppCallbacksPowerPoint.on_destroy: close active context dialog")
                self.app_ui.context_dialogs.close_active_dialog()
        
    
    # ==================================
    # = application-specific callbacks =
    # ==================================
    
    def slide_selection_changed(self, sld_range):
        logging.debug("app event slide_selection_changed")

        self.fire_event(self.events.slide_selection_changed)
        self.invalidate()
    
    def window_activate(self, pres, wnd):
        logging.debug("app event window_activate")

        self.fire_event(self.events.window_activate, presentation=pres, window=wnd)
        self.invalidate()
    
    def window_deactivate(self, pres, wnd):
        logging.debug("app event window_deactivate")

        self.fire_event(self.events.window_deactivate, presentation=pres, window=wnd)
    
    last_time_after_shape_size_changed = 0
    def after_shape_size_changed(self, shape):
        logging.debug("app event after_shape_size_changed")
        # If multiple shapes are resized together, the event fires individually for each shape.
        # restrict to a few invalidations per second
        if time.time() - self.last_time_after_shape_size_changed > 0.1:
            self.fire_event(self.events.after_shapesize_changed)
            self.invalidate()
            self.last_time_after_shape_size_changed = time.time()
    
    # last_time_window_selection_changed_in_text = 0
    # NOTE: This event is binded in C#-Addin for performance reasons!
    def window_selection_changed(self, selection):
        logging.debug("app event window_selection_changed")
        
        self.fire_event(self.events.selection_changed)
        try:
            self.invalidate()
        except:
            logging.error(traceback.format_exc())

        try:
            if self.app_ui.use_contextdialogs:
                self.app_ui.context_dialogs.show_shape_dialog_for_selection(selection, self.context)
        except:
            logging.error(traceback.format_exc())
        
        # # 0 = ppSelectionNone
        # # 1 = ppSelectionSlide
        # # 2 = ppSelectionShape
        # # 3 = ppSelectionText
        # if selection.type == 3:
        #     # selection in text, event raised a lot during keyboard input
        #     # restrict to a few invalidations per second
        #     #logging.debug("text selection")
        #     #logging.debug("last_time_window_selection_changed_in_text " + str(self.last_time_window_selection_changed_in_text))
        #     if time.time() - self.last_time_window_selection_changed_in_text > 0.5:
        #         self.invalidate()
        #         self.last_time_window_selection_changed_in_text = time.time()
            
        # else:
        #     #logging.debug("other selection")
        #     self.last_time_window_selection_changed_in_text = 0
        #     try:
        #         self.invalidate()
        #     except:
        #         logging.error(traceback.format_exc())
            
        #     try:
        #         self.app_ui.context_dialogs.show_shape_dialog_for_selection(selection, self.context)
        #     except:
        #         logging.error(traceback.format_exc())
    
    def after_presentation_open(self, pres):
        logging.debug("app event after_presentation_open")
        self.fire_event(self.events.after_presentation_open, presentation=pres)
    
    def after_new_presentation(self, pres):
        logging.debug("app event after_new_presentation")
        self.fire_event(self.events.after_new_presentation, presentation=pres)
    
    def presentation_close(self, pres):
        logging.debug("app event presentation_close")
        self.fire_event(self.events.presentation_close, presentation=pres)





class AppCallbacksFactory(object):
    '''
    Provide access to specific AppCallbacks-instances for office applications
    '''
    
    registry = {}
    
    app_callbacks_classes = {
        'Microsoft Excel':      AppCallbacksExcel,
        'Microsoft PowerPoint': AppCallbacksPowerPoint,
        'Microsoft Visio':      AppCallbacksVisio
    }
    
    # @property
    # @classmethod
    # def PowerPoint(cls):
    #     return cls.get_app_ui("Microsoft PowerPoint")
    
    @classmethod
    def get_app_callbacks(cls, app_name):
        if app_name in cls.registry:
            return cls.registry[app_name]
        else:
            instance = cls.create_app_callbacks(app_name)
            cls.registry[app_name] = instance
            return instance
    
    
    @classmethod
    def create_app_callbacks(cls, app_name, *args, **kwargs):
        ''' create AppCallbacks-instance for given app name '''
        # get AppUI-subclass for app name
        app_callbacks_class = cls.app_callbacks_classes.get(app_name, AppCallbacksBase)
        # create instance
        return app_callbacks_class(*args, **kwargs)

        







class Resources(object):
    ''' Encapsulated path resolution for file resources (such as images) '''
    root_folders = []
    images = None
    
    def __init__(self, category, suffix):
        self.category = category
        self.suffix = suffix

        cache_file = os.path.join( _h.get_cache_folder(), "resources.%s.cache"%category )
        
        try:
            self._cache = shelve.open(cache_file, protocol=2)
        except:
            logging.error("Loading resource cache failed")
            logging.debug(traceback.format_exc())
            
    def locate(self, name):
        try:
            return self._cache[name]
        except KeyError:
            logging.info("Locate resource: %s"%name)
            for root_folder in self.root_folders:
                path = os.path.join(root_folder, self.category, name + '.' + self.suffix)
                if os.path.exists(path):
                    self._cache[name] = path
                    self._cache.sync() #sync after each change as .close() is never called
                    return path
            return None
        except:
            logging.error("Unknown error reading from resource cache")
            logging.debug(traceback.format_exc())
            return None
    
    @staticmethod
    def bootstrap():
        package_dir = os.path.dirname(__file__)
        Resources.root_folders = [ os.path.normpath(os.path.join(package_dir,'..','resources')) ]
        Resources.images = Resources("images", "png")
    
Resources.bootstrap()




# ========================
# = Predefined callbacks =
# ========================

def get_enabled_ppt_shapes_or_text_selected(selection):
    # print "callback get_enabled_ppt_shapes_or_text_selected"
    return (selection.Type == 2 or selection.Type == 3)

def get_enabled_ppt_selection_contains_textframe(selection):
    # print "callback get_enabled_ppt_selection_contains_textframe"
    try:
        if not selection.Type in [2,3]:
            #neither text nor shapes selected
            return False
        elif selection.HasChildShapeRange:
            #selection within a group
            if selection.ChildShapeRange.HasTextFrame == 0:
                #none of the shapes has a TextFrame (otherwise HasTextFrame is -1 or -2)
                #note: a group cannot contain a table or "subgroups", no need to check for this
                return False
            else:
                return True
        elif selection.ShapeRange.HasTextFrame in [-2, -1] or selection.ShapeRange.HasTable in [-2, -1]:
            #at least one shape has a textframe (otherwise HasTextFrame is 0) or at least one table is selected (otherwise HasTable is 0)
            return True
        else:
            #shape selection may contain a group or SmartArt
            # (SmartArts are very strange objects, HasSmartArt is only impemented for single selection, GroupItems of SmartArt can be iterated but GroupItems.Range is invalid request)
            for shape in selection.ShapeRange:
                if shape.Type == 24: #msoSmartArt
                    return True
                if shape.Type == 6: #msoGroup
                    if shape.GroupItems.Range(None).HasTextFrame in [-2, -1]: #Range(None) return a ShapeRange object with all shapes in the group
                        return True
            return False #no TextFrames found
    except:
        #Any failure (e.g. ShapeRange fails when notes field is entered)
        return False

def get_enabled_ppt_shapes_min2_selected(selection):
    # print "callback get_enabled_ppt_shapes_min2_selected"
    try:
        if selection.Type != 2 and selection.Type != 3:
            return False

        if selection.HasChildShapeRange:
            return selection.ChildShapeRange.Count >= 2
        else:
            return selection.ShapeRange.Count >= 2
    except:
        return False


# ppt_shapes_or_text_selected      = Callback(get_enabled_ppt_shapes_or_text_selected)
ppt_shapes_or_text_selected      = "GetEnabled_Ppt_ShapesOrText"
# ppt_selection_contains_textframe = Callback(get_enabled_ppt_selection_contains_textframe)
ppt_selection_contains_textframe = "GetEnabled_Ppt_ContainsTextFrame"

# ppt_shapes_min2_selected         = Callback(get_enabled_ppt_shapes_min2_selected)
ppt_shapes_min2_selected         = "GetEnabled_Ppt_Shapes_MinTwo"
# ppt_shapes_min2_selected         = Callback(lambda shapes: True, shapes=True, shapes_min=2)

ppt_shapes_exactly1_selected         = "GetEnabled_Ppt_Shapes_ExactOne"
ppt_shapes_exactly2_selected         = "GetEnabled_Ppt_Shapes_ExactTwo"