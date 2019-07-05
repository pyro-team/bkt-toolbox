# -*- coding: utf-8 -*-
'''
Created on 13.11.2014

@authors: cschmitt, rdebeerst
'''

import sys
import traceback

import bkt
# import bkt.helpers as _h

_h = bkt.helpers
# linq = bkt.dotnet.import_linq() #this import doesn't seem to be required...
Bitmap = bkt.dotnet.import_drawing().Bitmap

import time
import logging
import os.path

import imp #for importing feature folder
import importlib #for importing traditional modules

from collections import OrderedDict #for import cache
import shelve #for import cache

#from helpers import config


# ======================
# = Initialize Logging =
# ======================

#FIXME: gleiche Log-Datei wie im .Net-Addin verwenden. Verwendung von bkt-debug.log führt noch zu Fehlern (Verlust von log-Text), da die Logger nicht Zeilenweise schreiben. Alternativ logging über C#-Addin-Klasse durchführen
if bkt.config.log_write_file == 'true' or (type(bkt.config.log_write_file) == bool and bkt.config.log_write_file):
    log_level = logging.WARNING
    try:
        log_level = getattr(logging, bkt.config.log_level or 'WARNING')
    except:
        pass
    
    logfile = os.path.join(os.path.dirname(os.path.realpath(__file__)), "..", "bkt-debug-py.log")
    filehandler = logging.FileHandler(logfile, 'w', 'utf-8')
    filehandler.setLevel(log_level)
    filehandler.setFormatter(logging.Formatter(u'%(asctime)s %(levelname)s: %(message)s'))
    logger = logging.getLogger()
    logger.setLevel(log_level)
    logger.addHandler(filehandler)

    # logging.basicConfig(
    #     filename=os.path.join(os.path.dirname(os.path.realpath(__file__)), "..", "bkt-debug-py.log"), 
    #     filemode='w',
    #     format='%(asctime)s %(levelname)s: %(message)s', 
    #     level=log_level
    # )



class CallbackManager(object):
    '''
    Manage all addin-callbacks and run callbacks for control-ids.
    '''
    
    def __init__(self):
        self.callback_resolution = {}
        self.ribbon_controls = {}
        self.context = None
    
    
    def init_callbacks_from_control(self, control):
        ''' initialize all callbacks from given control and its children '''
        logging.debug('CallbackManager: init_callbacks for %s ' % control)
        
        if (not control):
            return
        
        controls = set()
        for callback in control.collect_callbacks():
            cb_key = (callback.callback_type, callback.control.id, callback.control.id_tag)
            #logging.debug('callback key (cb_key): ' + str(cb_key))
            #logging.debug('defined callback: control %s / callback-type %s' % (callback.control, callback.callback_type))
            self.callback_resolution[cb_key] = callback
            #logging.debug('method: ' + str(callback))
            
            if not callback.control in controls:
                controls.add(callback.control)
                self.ribbon_controls[callback.control.id] = callback.control
    
    
    def resolve_callback(self, callback_type, control):
        ''' obtain callback from callback_type and control '''        
        # FIXME: detect WPF or Ribbon Control
        try:
            control_id = control.Id
        except:
            try:
                control_id = control.Name
            except:
                control_id = None
        
        my_control = self.ribbon_controls.get(control_id)
        if my_control is None:
            #print("callback %s unresolved for id %s (unknown control)" % (callback_type, control.Id))
            return (None, None)
        
        cb_key = (callback_type, my_control.id, my_control.id_tag)
        # logging.debug("resolve_callback by tuple: %s, %s, %s" % cb_key)
        return my_control, self.callback_resolution.get(cb_key)







def add_callbacks(cls):
    ''' decorator to add callback-functions to all callbacks defined in CallbackTypes '''
    
    def create_cb_method(callback):
        
        def addin_callback(self, control, *args):
            try:
                return self._callback(callback, control, *args)
            except:
                traceback.print_exc()
                logging.warning("error on callback. callback=" + str(callback) + ", control.id=" + str(control.id))
                try:
                    if bkt.config.show_exception:
                        # show exception only, if multiple errors do not occur within a second
                        # --> breaks exception-messages during ribbon-invalidate
                        if time.time() - self.last_exception_time < 1:
                            return
                        _h.exception_as_message()
                        self.last_exception_time = time.time()
                except:
                    pass
                    
        addin_callback.__name__ = callback.python_name
        return addin_callback
    
    for name, callback in bkt.callbacks.CallbackTypes.callback_map().items():
        if callback.custom:
            continue
        setattr(cls, name, create_cb_method(callback))
        
    return cls



@add_callbacks
class AddIn(object):
    '''
    Python counterpart of the IDTExtensibility2 and IRibbonExtensibility implementation.
    This class is responsible for the creation of the ribbon XML data as well as the dispatching
    of IRibbonExtensibility callbacks and their resolution to the corresponding (Python) targets.
    '''
    
    
    
    def __init__(self):
        ''' only set empty attributes here, actual work is done in on_create '''
        logging.info('\n=================================\n===== New AddIn initialized =====\n=================================')
        

        self.fallback_map = {bkt.callbacks.CallbackTypes.get_enabled: self.fallback_get_enabled}
        self.last_exception = None
        self.last_exception_time = 0
        self.app_callbacks = None
        self.reset()
    
    
    def reset(self):
        self.created = False
        self.context = None
        self.callback_manager = CallbackManager()
        
        self.events = None
        if self.app_callbacks:
            self.app_callbacks.unbind_app_events()
        self.app_callbacks=None
        self.app_ui = None
        
    
    
    def fallback_get_enabled(self, control):
        if control.default_callback is None or control.default_callback_control is None:
            return True
        
        #logging.debug('fallback_get_enabled: called for control id=%s' % (control.id))
        
        cb_control = control.default_callback_control
        cb_key = (cb_control.default_callback, cb_control.id, cb_control.id_tag)
        #FIXME
        callback = self.callback_manager.callback_resolution.get(cb_key)
        
        if callback is None:
            #print('fallback_get_enabled: no target found for %s, resolved with %s' % (control, cb_control))
            return True
        
        if callback.invocation_context is None:
            #print('fallback_get_enabled: target %s for %s, resolved with %s, has no invocation context' % (callback, control, cb_control))
            return True
        
        try:
            #logging.debug("fallback_get_enabled: try resolve_arguments")
            self.context.resolve_arguments(callback.invocation_context)
            return True
        except bkt.context.InappropriateContextError:
            #logging.debug("fallback_get_enabled: InappropriateContextError")
            return False
    
    def _callback(self, callback_type, control, *args, **kwargs):
        logging.debug("invoke callback for control. callback=" + str(callback_type) + ", control.id=" + str(control.id))
        #logging.debug("invoke callback for control. callback=" + str(callback_type) + ", control=" + str(control))
        my_control, callback = self.callback_manager.resolve_callback(callback_type, control)
        return_value = None
        
        if my_control is None:
            logging.warning("could not process callback. no control for control %s, event type %s" % (control, callback_type))
            if callback_type == bkt.callbacks.CallbackTypes.get_enabled:
                return True
            else:
                return
        
        if callback is None:
            #Do not show hundreds of warnings due to get_enabled
            if callback_type == bkt.callbacks.CallbackTypes.get_enabled:
                logging.debug("could not process callback. no callback of type %s for control %s. trying fallback" %  (callback_type, control))
            else:
                logging.warning("could not process callback. no callback of type %s for control %s. trying fallback" %  (callback_type, control))
            #logging.warning("could not process callback. no callback of type %s for control %s. trying fallback" %  (callback_type, control))
            fallback = self.fallback_map.get(callback_type)
            
            if fallback is None:
                #logging.debug("callback_type %s unresolved for id %s (no callback registered for control id)" % (callback_type, control.id))
                return
            
            #logging.debug("trying fallback")
            return fallback(my_control)
        
        #logging.debug("invoking callback: %s --- args=%s --- kwargs=%s" % (callback, args, kwargs))
        try:
            self.context.current_control = my_control
            self.context.customui_control = control
            #kwargs.update(self.context.resolve_callback.resolve_arguments(callback.invocation_context))
            #return_value= self.app_callbacks.invoke_callback(callback, *args, **kwargs)
            return_value= self.context.invoke_callback(callback, *args, **kwargs)
            logging.debug("return value=%s" % return_value)
            
            if callback.callback_type == bkt.callbacks.CallbackTypes.get_content:
                # get_content return ribbon-menu-object
                # initialize callbacks and return xml
                menu = return_value
                if isinstance(menu, bkt.ribbon.Menu):
                    self.callback_manager.init_callbacks_from_control(menu)
                    return_value = menu.xml_string()
                else:
                    logging.warning("Unexpected return-type in callback for get_content: got %s, expected %s" % (type(menu), bkt.ribbon.Menu))
                    return_value = str(menu)
                
                logging.debug("get_content returned:\n %s" % (return_value))
            
        except:
            logging.error("invoke callback failed for\ncallback-type=" + str(callback_type) + "\ncallback=" + str(callback))
            logging.debug(traceback.format_exc())
            try:
                if bkt.config.show_exception:
                    # show exception only, if multiple errors do not occur within a second
                    # --> breaks exception-messages during ribbon-invalidate
                    if time.time() - self.last_exception_time < 1:
                        return
                    _h.exception_as_message()
                    self.last_exception_time = time.time()
            except:
                pass
        
        finally:
            # this causes multiple invalidations
            # don't do this for enabled/visible/etc events
            # if callback_type in [bkt.callbacks.CallbackTypes.on_action,
            #     bkt.callbacks.CallbackTypes.on_action_indexed,
            #     bkt.callbacks.CallbackTypes.on_action_repurposed,
            #     bkt.callbacks.CallbackTypes.on_toggle_action,
            #     bkt.callbacks.CallbackTypes.on_change]:
            if callback_type.transactional:
                self.invalidate_ribbon()
        
        return return_value
            
    #@deprecated
    # def get_enabled_ppt_shapes_or_text_selected(self, control):
    #     return (self.context.app.ActiveWindow.selection.Type == 2 or self.context.app.ActiveWindow.selection.Type == 3)
    
    
    
    def task_pane(self, sender, eventargs):
        logging.debug("---------- event: %s ---------- type / name: %s / %s ----------" % (eventargs.RoutedEvent, type(eventargs.Source), eventargs.Source.Name))
        #logging.debug("task pane invoked callback. event-type=%s, sender name/type=%s/%s" % (eventargs.RoutedEvent, eventargs.Source.Name, eventargs.Source))
        # logging.debug("routed event details: type=%s" % type(eventargs.RoutedEvent))
        # logging.debug("routed event details: name=%s" % eventargs.RoutedEvent.Name)
        # logging.debug("routed event details: handler-type=%s" % eventargs.RoutedEvent.HandlerType)
        # logging.debug("routed event details: owner-type=%s" % eventargs.RoutedEvent.OwnerType)
        # logging.debug("routed source details: type=%s" % type(eventargs.Source))
        # logging.debug("routed source details: get-type=%s" % eventargs.Source.GetType())
        # logging.warning("routed event details: owner-type==ButtonBase %s" % (eventargs.RoutedEvent.OwnerType == controls.Primitives.ButtonBase))
        #_h.message("Hello World from Python!\nYou just clicked a task pane control!\n\nyou clicked: %s\nevent-type: %s" % (eventargs.Source, eventargs.RoutedEvent))
        
        try:        
            # 1) handle general wpf-event
            self._callback(bkt.callbacks.CallbackTypes.wpf_event, eventargs.Source)
            # TODO: check whether event was handled
            
            # 2) reroute by EventType
            # FIXME: define routings in taskpane.py
            # TODO: invalidation und Nutzung get_pressed, get_enabled, get_text, ...
            logging.debug("Start RoutedEvents-mapping: RoutedEvent-Name=%s, Source-Type=%s" % (str(eventargs.RoutedEvent.Name), str(eventargs.Source.GetType())))
            
            if str(eventargs.RoutedEvent.Name)=='LostFocus':
                if str(eventargs.Source.GetType()) in ['Fluent.TextBox', 'Fluent.Spinner']:
                    ### TEXTBOX LOST FOCUS EVENT
                    logging.debug("map RoutedEvent to CallbackType.on_change: text=%s" % eventargs.Source.Text)
                    self._callback(bkt.callbacks.CallbackTypes.on_change, eventargs.Source, eventargs.Source.Text)
                    
                elif str(eventargs.Source.GetType()) in ['Fluent.ComboBox']:
                    if not eventargs.Source.IsReadOnly:
                        ### EDITABLE COMBOBOX LOST FOCUS EVENT
                        logging.debug("map RoutedEvent to CallbackType.on_change: text=%s" % eventargs.Source.Text)
                        self._callback(bkt.callbacks.CallbackTypes.on_change, eventargs.Source, eventargs.Source.Text)
                    
            elif str(eventargs.RoutedEvent.Name)=='KeyDown':
                if str(eventargs.Source.GetType()) in ['Fluent.TextBox', 'Fluent.Spinner']:
                    ### TEXTBOX ENTER FIRED
                    logging.debug("map RoutedEvent to CallbackType.on_change: text=%s" % eventargs.Source.Text)
                    self._callback(bkt.callbacks.CallbackTypes.on_change, eventargs.Source, eventargs.Source.Text)
                    
                elif str(eventargs.Source.GetType()) in ['Fluent.ComboBox']:
                    if not eventargs.Source.IsReadOnly:
                        ### EDITABLE COMBOBOX ENTER FIRED
                        logging.debug("map RoutedEvent to CallbackType.on_change: text=%s" % eventargs.Source.Text)
                        self._callback(bkt.callbacks.CallbackTypes.on_change, eventargs.Source, eventargs.Source.Text)
            
            elif str(eventargs.RoutedEvent.Name)=='Click':
                if str(eventargs.Source.GetType()) in ['Fluent.MenuItem']:
                    if eventargs.Source.IsCheckable == True:
                        ### MENU ITEM TOGGLE EVENT
                        logging.debug("map RoutedEvent to CallbackType.on_toggle_action: pressed/checked=%s" % eventargs.Source.IsChecked)
                        self._callback(bkt.callbacks.CallbackTypes.on_toggle_action, eventargs.Source, eventargs.Source.IsChecked)
                        
                    else:
                        ### MENU ITEM CLICK EVENT
                        logging.debug("map RoutedEvent to CallbackType.on_action")
                        self._callback(bkt.callbacks.CallbackTypes.on_action, eventargs.Source)
                        
                elif str(eventargs.Source.GetType()) in ['Fluent.Spinner']:
                    ### SPINNER BUTTON CLICK
                    logging.debug("map RoutedEvent to CallbackType.on_change")
                    self._callback(bkt.callbacks.CallbackTypes.on_change, eventargs.Source, eventargs.Source.Value)
                    #self._callback(bkt.callbacks.CallbackTypes.on_action, eventargs.Source, eventargs.Source.Value)
                    
                elif str(eventargs.Source.GetType()) in ['System.Windows.Controls.Primitives.ToggleButton', 'Fluent.ToggleButton', 'Fluent.CheckBox', 'Fluent.RadioButton']:
                    ### TOGGLE EVENT
                    logging.debug("map RoutedEvent to CallbackType.on_toggle_action: pressed/checked=%s" % eventargs.Source.IsChecked)
                    self._callback(bkt.callbacks.CallbackTypes.on_toggle_action, eventargs.Source, eventargs.Source.IsChecked)
                    
                else:
                    ### OTHER CLICK EVENT
                    logging.debug("map RoutedEvent to CallbackType.on_action")
                    self._callback(bkt.callbacks.CallbackTypes.on_action, eventargs.Source)
            
            # FIXME
            # elif str(eventargs.RoutedEvent.Name)=='ValueChanged':
            #     logging.debug("map RoutedEvent to CallbackType.on_change: value=?" )
            
            elif str(eventargs.RoutedEvent.Name)=='SelectionChanged':
                
                if str(eventargs.Source.GetType()) in ['Fluent.Gallery', 'Fluent.InRibbonGallery']:
                    ### GALLERY SELECTION CHANGE EVENT
                    # assume eventargs.Source.SelectedValue is TextBlock
                    logging.debug("map RoutedEvent to on_change: value=%s" % eventargs.Source.SelectedValue.Text)
                    self._callback(bkt.callbacks.CallbackTypes.on_change, eventargs.Source, eventargs.Source.SelectedValue.Text)
                
                elif str(eventargs.Source.GetType()) in ['Fluent.ComboBox']:
                    ### COMBOBOX ACTION INDEXD EVENT / SELECTION CHANGE EVENT
                    logging.debug("map RoutedEvent to on_action_indexed: index=%s" % eventargs.Source.SelectedIndex)
                    self._callback(bkt.callbacks.CallbackTypes.on_action_indexed, eventargs.Source, eventargs.Source.SelectedIndex, eventargs.Source.SelectedIndex)
                    
                    logging.debug("map RoutedEvent to on_change: value=%s" % str(eventargs.Source.SelectedValue))
                    self._callback(bkt.callbacks.CallbackTypes.on_change, eventargs.Source, str(eventargs.Source.SelectedValue.Content))
            
            
            elif str(eventargs.RoutedEvent.Name)=='SelectedDateChanged':
                ### DATE PICKER CHANGE EVENT
                logging.debug("map RoutedEvent to CallbackType.on_change: value=%s" % eventargs.Source.SelectedDate)
                self._callback(bkt.callbacks.CallbackTypes.on_change, eventargs.Source, eventargs.Source.SelectedDate)
            
            
            elif str(eventargs.RoutedEvent.Name)=='SelectedColorChanged':
                ### SELECTED COLOR CHANGE EVENT
                logging.debug("map RoutedEvent to on_rgb_color_change: value=%s" % eventargs.Source.SelectedColor)
                self._callback(bkt.callbacks.CallbackTypes.on_rgb_color_change, eventargs.Source, color=eventargs.Source.SelectedColor)
                logging.debug("done")
                
            else:
                pass

            logging.debug("----------")
        
        except:
            logging.debug(traceback.format_exc())
            traceback.print_exc()
            try:
                if bkt.config.show_exception:
                    # show exception only, if multiple errors do not occur within a second
                    # --> breaks exception-messages during ribbon-invalidate
                    if time.time() - self.last_exception_time < 1:
                        return
                    _h.exception_as_message()
                    self.last_exception_time = time.time()
            except:
                pass
    
    
    
    def task_pane_value_changed(self, sender, eventargs):
        logging.debug("---------- event: ValueCanged ---------- type / name: %s ----------" % (sender.Name))
        
        try:        
            # eventargs RoutedPropertyChangedEventArgs<double>
            
            logging.debug('value changed, old-value=%s, new-value=%s' % (eventargs.OldValue, eventargs.NewValue) )
            
            logging.debug("map RoutedEvent to CallbackType.on_change")
            self._callback(bkt.callbacks.CallbackTypes.on_change, sender, eventargs.NewValue, old_value=eventargs.OldValue, new_value=eventargs.NewValue)
        
            logging.debug("----------")
        
        except:
            traceback.print_exc()
            try:
                if bkt.config.show_exception:
                    # show exception only, if multiple errors do not occur within a second
                    # --> breaks exception-messages during ribbon-invalidate
                    if time.time() - self.last_exception_time < 1:
                        return
                    _h.exception_as_message()
                    self.last_exception_time = time.time()
            except:
                pass
    
    
    # ===============
    # = mouse event =
    # ===============
    
    def mouse_down(self, sender, e):
        ''' object sender, MouseEventExtArgs e) '''
        if self.app_ui:
            if hasattr(self.app_ui, 'context_dialogs'):
                self.app_ui.context_dialogs.mouse_down(sender, e)
        
    def mouse_up(self, sender, e):
        ''' object sender, MouseEventExtArgs e) '''
        if self.app_ui:
            if hasattr(self.app_ui, 'context_dialogs'):
                self.app_ui.context_dialogs.mouse_up(sender, e)

    # def mouse_move(self, sender, e):
    #     ''' object sender, MouseEventExtArgs e) '''
    #     if self.app_ui:
    #         if hasattr(self.app_ui, 'context_dialogs'):
    #             self.app_ui.context_dialogs.mouse_move(sender, e)

    def mouse_drag_start(self, sender, e):
        ''' object sender, MouseEventExtArgs e) '''
        if self.app_ui:
            if hasattr(self.app_ui, 'context_dialogs'):
                self.app_ui.context_dialogs.mouse_drag_start(sender, e)

    def mouse_drag_end(self, sender, e):
        ''' object sender, MouseEventExtArgs e) '''
        if self.app_ui:
            if hasattr(self.app_ui, 'context_dialogs'):
                self.app_ui.context_dialogs.mouse_drag_end(sender, e)
        
    def key_down(self, sender, e):
        ''' object sender, KeyEventArgs e) '''
        if self.app_ui:
            if hasattr(self.app_ui, 'context_dialogs'):
                self.app_ui.context_dialogs.key_down(sender, e)
    
    def key_up(self, sender, e):
        ''' object sender, KeyEventArgs e) '''
        if self.app_ui:
            if hasattr(self.app_ui, 'context_dialogs'):
                self.app_ui.context_dialogs.key_up(sender, e)
    
    
    # ===============================
    # = app events binded in dotnet =
    # ===============================

    def ppt_selection_changed(self, selection):
        self.app_callbacks.window_selection_changed(selection)
    
    




    def invalidate_ribbon(self):
        if self.context:
            # can be false after dev-button did an addin-reconnect
            if self.context.ribbon:
                # FIXME: calling of ribbon.Invalidate should be done in apps.AppCallbacksBase.invalidate
                # FIXME: then call app_callbacks.invalidate here
                # reset caches to ensure proper invalidate
                self.app_callbacks.refresh_cache(True)
                self.context.refresh_cache(True)
                # print('invalidating ribbon')
                self.app_callbacks.fire_event(self.events.bkt_invalidate)
                self.context.ribbon.Invalidate()
                # reset caches for immediate interaction after invalidate
                self.app_callbacks.refresh_cache(True)
                self.context.refresh_cache(True)
            
    def on_destroy(self):
        self.app_callbacks.fire_event(self.events.bkt_unload)
        _h.settings.close() #save settings to database
        self.reset()
    
    def on_create(self, dotnet_context):
        logging.debug('on_create')
        if self.created:
            # TODO: discuss whether multiple calls to on_create are a relevant use case
            return
        self.created = True
        
        #open app-specific settings database
        _h.settings.open(dotnet_context.hostAppName)

        # wrap dotnet-context and add self as python_addin
        self.context = bkt.context.AppContext.create_app_context(dotnet_context.hostAppName, dotnet_context, python_addin=self)
        #self.dotnet_context.python_addin = self
        
        # extend PYTHONPATH
        for path in bkt.config.pythonpath or []:
            sys.path.append(path)
        
        # load modules listed in configuration
        for module in bkt.config.modules or []:
            if not module in sys.modules:
                logging.info('importing module: %s' % module)
                try:
                    importlib.import_module(module)
                    # __import__(module)
                except:
                    logging.error('failed to load %s' % module)
                    _h.message('failed to load %s' % module)
                    _h.message(traceback.format_exc())
                    #_h.exception_as_message('failed to load %s' % module)
        

        CACHE_VERSION = "20190603"
        cache_file = os.path.join( _h.get_cache_folder(), "%s.import.cache" % self.context.app_name )
        import_cache = shelve.open(cache_file, protocol=2)

        # STRUCTURE OF IMPORT CACHE #
        #############################
        # cache.app             name of app, e.g. PowerPoint
        # cache.time            timestamp of cache creation
        # cache.version         constant CACHE_VERSION to indicate new cache structure in new releases
        # cache.hash            hashed version of string representation of feature_folder list to invalidate cache
        # sys.path              set of folders to be added to sys.path (using set for deduplication)
        # resources.path        set of folders to be added to resources paths (for png icons)
        # inits.features        OrderedDict with module name as key and dict as value with infos (e.g. name, folder) for each feature_folder with __bkt_init__
        # inits.legacy          list of tuples with (module_name, folder) for all feature_folder without __bkt_init__

        try:
            # raise KeyError("STOP CACHE")
            if import_cache['cache.version'] != CACHE_VERSION or import_cache['cache.hash'] != hash(str(bkt.config.feature_folders)):
                logging.warning('invalid cache, fallback to normal import process')
                raise KeyError("Invalid cache version")

            logging.info('start import from cache')
            sys.path.extend(import_cache['sys.path'])
            bkt.apps.Resources.root_folders.extend(import_cache['resources.path'])

            for module_name, feature in import_cache['inits.features'].items():
                # feature = import_cache['feature.'+module_name]
                logging.info('importing bkt feature: %s' % feature['name'])
                try:
                    # import module as package, acts like 'import module_name'
                    #f, path, description = imp.find_module(module_name, base_folder)
                    imp.load_module(module_name, None, feature['folder'], ('', '', imp.PKG_DIRECTORY))
                    # run bkt_init
                    module = imp.load_source(module_name + '.__bkt_init__' , os.path.join(feature['folder'], "__bkt_init__.py"))

                    if feature['use_constructor']:
                        logging.debug("loading bkt feature %s using constructor method" % feature['name'])
                        module.BktFeature.contructor()
                    
                except:
                    logging.error('failed to load feature %s from cache' % module_name)
                    logging.error(traceback.format_exc())
                    _h.message('failed to load feature %s from cache' % module_name)
                    if bkt.config.show_exception:
                        _h.exception_as_message()
                    #TODO: remove cache on error?

            for module_name, folder in import_cache['inits.legacy']:
                logging.info('importing legacy feature-folder: %s' % folder)
                try:
                    imp.load_source(module_name, os.path.join(folder, "__init__.py"))
                except:
                    logging.error('failed to load legacy feature %s from cache' % module_name)
                    logging.error(traceback.format_exc())
                    _h.message('failed to load legacy feature %s from cache' % module_name)
                    if bkt.config.show_exception:
                        _h.exception_as_message()
                    #TODO: remove cache on error?
        
        # load and renew cache:
        except KeyError:
            logging.info('start import and build new cache')
            import_cache.clear()
            
            #do not write directly to cache as cache cannot know which values within a list change
            _c_sys_paths = set()
            _c_resources_paths = set()
            _c_bkt_inits = OrderedDict()
            _c_legacy_inits = []

            _known_conflicts = [] #list of features that cannot be loaded due to conflicting feature

            # load modules from feature-folders and fill cache on-the-fly
            for folder in bkt.config.feature_folders:
                logging.info('importing feature-folder: %s' % folder)
                
                base_folder = os.path.realpath(os.path.join(folder, ".."))
                module_name = os.path.basename(folder)
                init_filename = os.path.join(folder, "__bkt_init__.py")
                
                # initialize resource-folders from feature-folders
                res_path = os.path.join(folder,'resources')
                if os.path.exists(res_path):
                    bkt.apps.Resources.root_folders.append(res_path)
                    _c_resources_paths.add(res_path)

                # load using bkt_init.py file
                if os.path.isfile(init_filename):
                    try:
                        sys.path.append(base_folder)
                        _c_sys_paths.add(base_folder)
                        # import module as package, acts like 'import module_name'
                        #f, path, description = imp.find_module(module_name, base_folder)
                        imp.load_module(module_name, None, folder, ('', '', imp.PKG_DIRECTORY))
                        # run bkt_init
                        module = imp.load_source(module_name + '.__bkt_init__' , init_filename)
                        logging.debug('module: %s' % module)
                        
                        try:
                            #only load if relevant for active app
                            if self.context.app_name not in module.BktFeature.relevant_apps:
                                logging.info("bkt feature %s not relevant for current app" % module.BktFeature.name)
                                continue
                            #only load if dependencies are loaded
                            dependencies = getattr(module.BktFeature, "dependencies", [])
                            if not all(d in _c_bkt_inits for d in dependencies):
                                logging.warning("bkt feature %s is missing dependencies" % module.BktFeature.name)
                                continue
                            #only load if there are no conflicting features loaded
                            conflicts = getattr(module.BktFeature, "conflicts", [])
                            if module_name in _known_conflicts or any(m in _c_bkt_inits for m in conflicts):
                                logging.warning("bkt feature %s has a conflict" % module.BktFeature.name)
                                continue

                            logging.debug("loading bkt feature %s using constructor method" % module.BktFeature.name)
                            module.BktFeature.contructor()
                            
                            #add to cache only if relevant for this app and loading successful
                            _c_bkt_inits[module_name] = {
                                'name': module.BktFeature.name,
                                'folder': folder,
                                'use_constructor': True
                            }
                            #add conflicting modules
                            _known_conflicts.extend(conflicts)
                        except AttributeError:
                            #legacy bkt_init, load for all apps, always
                            logging.debug("bkt feature %s not using new loading mechanism" % module_name)
                            _c_bkt_inits[module_name] = {
                                'name': module_name,
                                'folder': folder,
                                'use_constructor': False
                            }
                    except:
                        logging.error('failed to load feature-folder %s' % folder)
                        logging.error(traceback.format_exc())
                        _h.message('failed to load feature-folder %s' % folder)
                        if bkt.config.show_exception:
                            _h.exception_as_message()
                        #TODO: Offer user to remove feature folder from config on error

                # backwards compatibility: load module from init.py
                elif os.path.isfile(os.path.join(folder, "__init__.py")):
                    try:
                        sys.path.append(folder)
                        imp.load_source(module_name, os.path.join(folder, "__init__.py"))
                        
                        _c_sys_paths.add(folder)
                        _c_legacy_inits.append((module_name, folder))
                    except:
                        logging.error('failed to load legacy feature-folder %s' % folder)
                        logging.error(traceback.format_exc())
                        _h.message('failed to load legacy feature-folder %s' % folder)
                        if bkt.config.show_exception:
                            _h.exception_as_message()
                        #TODO: Offer user to remove feature folder from config on error
            
            #save to cache
            import_cache['cache.app']       = self.context.app_name
            import_cache['cache.time']      = time.time()
            import_cache['cache.version']   = CACHE_VERSION
            import_cache['cache.hash']      = hash(str(bkt.config.feature_folders))
            import_cache['sys.path']        = _c_sys_paths
            import_cache['resources.path']  = _c_resources_paths
            import_cache['inits.features']  = _c_bkt_inits
            import_cache['inits.legacy']    = _c_legacy_inits
        finally:
            import_cache.close()


        
        #### initialize AppUI, AppCallbacks
        try:
            logging.debug('initialize classes for app: %s' % self.context.host_app_name)
            # retrieve AppUI-instance
            self.app_ui = bkt.appui.AppUIs.get_app_ui(self.context.host_app_name)
            self.events = bkt.apps.AppEvents
            # create ApplicationCallback-instance
            self.app_callbacks = bkt.apps.AppCallbacksFactory.create_app_callbacks(
                self.context.host_app_name,
                addin = self,
                app_ui = self.app_ui,
                appcontext = self.context,
                appevents = self.events
            )
            self.context.app_callbacks = self.app_callbacks
            self.context.app_ui = self.app_ui
            
        except:
            logging.critical("initialize app-classes failed")
            logging.debug(traceback.format_exc())
            _h.message("initialize app-classes failed")
        
        
            

        ### bind callbacks to app-sepcific events
        try:
            logging.debug('bind application events')
            self.app_callbacks.bind_app_events()
        except:
            logging.critical("binding of callbacks to application events failed")
            logging.debug(traceback.format_exc())
            _h.message("binding of callbacks to application events failed")
        
        
        logging.debug('on_create done ')
        #_h.message('on_create done')
    
    
    
    
    def on_ribbon_load(self, ribbon):
        ''' IRibbonUI ribbon'''
        self.app_callbacks.fire_event(self.events.bkt_load)
    
    def load_image(self, image_name):
        path = bkt.apps.Resources.images.locate(image_name)  #@UndefinedVariable
        if path is None:
            return
        return Bitmap.FromFile(path)
            
    def get_custom_ui(self, ribbon_id):
        try:
            logging.info('Retrieve CustomUI for ribbon: %s' % ribbon_id)
            
            if self.app_ui is None:
                return None
            
            ### initialize UI-callbacks
            
            try:
                # init ribbon-callbacks
                self.callback_manager.init_callbacks_from_control(self.app_ui.get_customui_control(ribbon_id))
                # init taskpane-callbacks
                self.callback_manager.init_callbacks_from_control(self.app_ui.get_taskpane_control())
            except:
                logging.critical("initialization of ui-callbacks failed")
                logging.debug(traceback.format_exc())
                _h.message("initialization of ui-callbacks failed")
            
            
            ### retrieve ribbon ui
            custom_ui = self.app_ui.get_customui(ribbon_id)
            return custom_ui
            
        except:
            #traceback.print_exc()
            logging.critical('get_custom_ui failed!')
            logging.debug(traceback.format_exc())
            _h.message(traceback.format_exc())
            #_h.exception_as_message()
    
    
    def get_custom_taskpane_ui(self):
        try:
            logging.debug('Retrieve UI for taskpane')
            
            if self.app_ui is None:
                return None
            else:
                logging.debug(self.app_ui.get_taskpane_ui())
                return self.app_ui.get_taskpane_ui()
        except:
            #traceback.print_exc()
            logging.critical('get_custom_taskpane_ui failed!')
            _h.message(traceback.format_exc())
            #_h.exception_as_message()

        



