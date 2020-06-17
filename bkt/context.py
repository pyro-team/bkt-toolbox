# -*- coding: utf-8 -*-
'''
Resolve app-specific parameters and resolve arguments of callbacks

Created on 11.11.2019
@author: rdebeerst
'''

from __future__ import absolute_import

import logging
import time #required for cache

# from System.Runtime.InteropServices import Marshal

import bkt.helpers as _h #for providing config and settings
from bkt.library.comrelease import AutoReleasingComObject


class InappropriateContextError(Exception):
    pass


class AppContext(object):
    app_name = 'Uknown'
    
    @classmethod
    def create_app_context(cls, app_name, *args, **kwargs):
        ''' provide AppContext-instance, wich can be specialiced specific apps '''
        
        # app-specific classes
        app_classes = [AppContextPowerPoint, AppContextExcel, AppContextVisio]
        # compare app-name in app-specific classes
        for app_class in app_classes:
            if app_name == app_class.app_name:
                return app_class(*args, **kwargs)
        
        # default
        return AppContext(*args, **kwargs)
    
    
    def __init__(self, dotnet_context, python_addin=None, config=None, app_callbacks=None, app_ui=None, settings=None):
        ''' Initialize AppContext '''
        # attributes with direct access
        self.dotnet_context = dotnet_context
        self.python_addin = python_addin
        self.app_callbacks = app_callbacks
        self.app_ui = app_ui
        # no direct access to config, fallback to global config-instance
        self._config = config
        self._settings = settings
        # uninitialized values
        self.customui_control = None
        self.current_control = None
        # caching of arguments
        self.cache = {}
        self.cache_timeout = 0.5
        self.cache_last_refresh = 0
        
        # references with auto-release-effect
        self.app = AutoReleasingComObject(dotnet_context.app, release_self=False)
    
    
    def release_com_references(self):
        logging.debug("Context.release_com_references")
        self.app.dispose()
    
    
    def refresh_cache(self, force=False):
        #global cache timeout will prevent that manual invalidates are not working properly
        if force or time.time() - self.cache_last_refresh > self.cache_timeout:
            logging.debug("Context.refresh_cache, force=%s", force)
            self.cache = {}
            self.cache_last_refresh = time.time()
            return True
        return False
        
    
    @property
    def config(self):
        ''' return local config (from initialization) or global config instance '''
        if self._config:
            return self._config
        else:
            return _h.config
    
    @property
    def settings(self):
        ''' return local settings (from initialization) or global settings instance '''
        if self._settings:
            return self._settings
        else:
            return _h.settings
    
    # ===========================================
    # = convenience properties for .Net-Context =
    # ===========================================
    
    # @property
    # def app(self):
    #     return self.dotnet_context.app
    
    @property
    def addin(self):
        return self.dotnet_context.addin
    
    @property
    def debug(self):
        return self.dotnet_context.debug
    
    @property
    def ribbon(self):
        return self.dotnet_context.ribbon

    @property
    def host_app_name(self):
        return self.dotnet_context.hostAppName
    
    
    # =====================
    # = argument resolver =
    # =====================
    
    def resolve_generic_arguments(self, nfo):
        ''' resolve generic context-arguments defined through nfo '''
        if nfo is None:
            return {}
        
        args = {}
        
        # GENERAL ACCESS
        if nfo.context:
            args['context'] = self
        
        if nfo.python_addin:
            args['python_addin'] = self.python_addin
        
        # FIXME
        # if nfo.app_events:
        #     args['app_events'] = self.app_events
        
        # if nfo.app_ui:
        #     args['app_ui'] = self.app_ui
        
        # .Net-PROPERTIES
        if nfo.application:
            args['application'] = self.app
        
        
        # OTHERS
        
        # FIXME: 
        if nfo.ribbon_id:
            args['ribbon_id'] = ""#self.addin_customui.ribbon_id
        
        if nfo.customui_control:
            args['customui_control'] = self.customui_control

        if nfo.current_control:
            args['current_control'] = self.current_control
    
        return args
    
    
    def resolve_arguments(self, nfo):
        ''' resolve arguments defined through nfo in the current context '''
        
        # refresh argument cache
        self.refresh_cache()

        # resolve generic arguments
        args = self.resolve_generic_arguments(nfo)
        
        # resolve application-specific arguments
        args.update( self.resolve_app_arguments(nfo) )
        
        return args
    
    def resolve_app_arguments(self, nfo):
        ''' resolve app-specific arguments defined through nfo in the current context '''
        
        return {}
        
        
    def fail(self):
        ''' throws InappropriateContextError, for usage in argument-resolver-methods '''
        raise InappropriateContextError
    
    
    # ========================================================
    # = convenience methods for instances in current context =
    # ========================================================
    
    #def invoke_callback(self, *args, **kwargs):
    def invoke_callback(self, callback, *args, **kwargs):
        ''' convenience-method to invoke a callback form the AppEvents-instance in the current context '''
        # resolve generic arguments
        kwargs.update(self.resolve_generic_arguments(callback.invocation_context))
        # application-specific arguments should be resolved by invoke_callback
        return_value = self.app_callbacks.invoke_callback(self, callback, *args, **kwargs)
        # release com objects
        # logging.debug("Context.invoke_callback: request com release after callback %s", callback.method)
        # self.release_com_references()
        return return_value
    
    
    
    # =================================================
    # = complete set of properties for tab-completion =
    # =================================================
    
    def __dir__(self):
        ''' extend the dir of this class by the members of the wrapped dotnet_context.
            needed for tab-completion in the console
        '''
        #res = set(self.__dict__) | set(type(self).__dict__) | set(dir(self.dotnet_context)) 
        res = set(self.__dict__) | set(type(self).__dict__)
        return sorted(res)







# =================
# =  POWER POINT  =
# =================

class AppContextPowerPoint(AppContext):
    app_name = 'Microsoft PowerPoint'
    
    def __init__(self, *args, **kwargs):
        
        super(AppContextPowerPoint, self).__init__(*args, **kwargs)
        
        from bkt.library.powerpoint import wrap_shapes
        self.wrap_shapes = wrap_shapes
    
    @property
    def shapes(self):
        ''' gives list-access to app.ActiveWindow.Selection.ShapeRange / ChildShapeRange '''
        # ShapeRange accessible if shape or text selected
        selection = self.selection
        if selection.Type != 2 and selection.Type != 3:
            return []
        
        if selection.HasChildShapeRange:
            # shape selection inside grouped shapes
            shapes = list(iter(selection.ChildShapeRange))
        else:
            shapes = list(iter(selection.ShapeRange))
        
        
        return shapes

    @property
    def shape(self):
        return self.shapes[0]
    
    @property
    def shapes_wrapped(self):
        self.wrap_shapes(self.shapes)

    @property
    def slides(self):
        ''' gives list-access to app.ActiveWindow.Selection.SlideRange '''
        try:
            slides = list(iter(self.selection.SlideRange))
        except EnvironmentError:
            #fallback for Invalid request.  SlideRange cannot be constructed from a Master.
            return [self.app.ActiveWindow.View.Slide]
        return slides

    @property
    def slide(self):
        try:
            return self.slides[0]
        except EnvironmentError:
            #fallback for Invalid request.  SlideRange cannot be constructed from a Master.
            return self.app.ActiveWindow.View.Slide
    
    @property
    def selection(self):
        try:
            return self.app.ActiveWindow.Selection
        except:
            # fails, if ActiveWindow is not available
            self.fail()

    @property
    def presentation(self):
        try:
            # return self.app.ActiveWindow.Presentation #fails, if ActiveWindow is not available (e.g. in slideshow mode), so better to use ActivePresentation
            return self.app.ActivePresentation
        except:
            self.fail()
    
    
    def resolve_app_arguments(self, nfo):
        '''
        resolves the arguments of a target function from the addin context and a context information
        object
        '''
        
        if not self.app:
            self.fail()

        args = {}
        
        if nfo.shapes or nfo.shape:
            # fail here, if selection is not available
            selection = self.selection
            
            # ShapeRange accessible if shape or text selected
            if selection.Type != 2 and selection.Type != 3:
                self.fail()
            
            try:
                shapes = self.cache['shapes']
                shapes_textframes = self.cache['shapes_textframes']
            except KeyError:
                try:
                    if selection.HasChildShapeRange:
                        # shape selection inside grouped shapes
                        self.cache['shapes'] = shapes = list(iter(selection.ChildShapeRange))
                        self.cache['shapes_textframes'] = shapes_textframes = selection.ChildShapeRange.HasTextFrame
                    else:
                        self.cache['shapes'] = shapes = list(iter(selection.ShapeRange))
                        self.cache['shapes_textframes'] = shapes_textframes = selection.ShapeRange.HasTextFrame
                except:
                    shapes = None
            
            if not shapes:
                self.fail()
            check_limits(shapes, nfo.shapes_min, nfo.shapes_max)
            
            if nfo.require_text:
                if shapes_textframes != -1:
                    self.fail()
                # for shape in shapes:
                #     if not shape.HasTextFrame:
                #         self.fail()

            if nfo.wrap_shapes:
                try:
                    shapes = self.cache['wrapped_shapes']
                except KeyError:
                    self.cache['wrapped_shapes'] = shapes = self.wrap_shapes(shapes)

            if nfo.shape:
                if len(shapes) != 1:
                    self.fail()
                args['shape'] = shapes[0]
            else:
                args['shapes'] = shapes
        
        if nfo.slides or nfo.slide:
            # fail here, if selection is not available
            selection = self.selection
            
            # SlideRange accessible if slides, shapes or text selected
            try:
                slides = self.cache['slides']
            except KeyError:
                try:
                    self.cache['slides'] = slides = list(iter(selection.SlideRange))
                except EnvironmentError:
                    #fallback to slide in view, e.g. Invalid request.  SlideRange cannot be constructed from a Master.
                    try:
                        self.cache['slides'] = slides = [self.app.ActiveWindow.View.Slide]
                    except:
                        self.fail()
                
            if not slides:
                self.fail()
            check_limits(slides, nfo.slides_min, nfo.slides_max)
            
            if nfo.slide:
                if len(slides) != 1:
                    self.fail()
                args['slide'] = slides[0]
            else:
                args['slides'] = slides
            
        if nfo.presentation:
            args['presentation'] = self.presentation

        if nfo.selection:
            args['selection'] = self.selection
            
        return args
    





# ===========
# =  EXCEL  =
# ===========


class AppContextExcel(AppContext):
    app_name = 'Microsoft Excel'
    
    def resolve_app_arguments(self, nfo):
        # #TODO: check if this is helpful:
        # if not self.app.Ready:
        #     self.fail()

        # #Dirty hack to disable menu in edit mode / doesnt work as invalidate is not called
        # try:
        #     if not self.app.CommandBars.GetEnabledMso("MergeCellsAcross"):
        #         logging.warning("edit mode, menu disabled")
        #         self.fail()
        # except:
        #     logging.warning("could not get enabled state of menu")

        args = {}

        if nfo.workbook:
            workbook = self.app.ActiveWorkbook
            if not workbook:
                self.fail()
            args['workbook'] = workbook

        if nfo.sheet:
            sheet = self.app.ActiveSheet
            if not sheet or (nfo.require_worksheet and sheet.Type != -4167): #Worksheet
                self.fail()
            args['sheet'] = sheet

        if nfo.sheets:
            try:
                args['sheets'] = list(iter(self.app.ActiveWorkbook.Sheets))
            except:
                self.fail()

        if nfo.selected_sheets:
            try:
                args['selected_sheets'] = list(iter(self.app.ActiveWindow.SelectedSheets))
            except:
                self.fail()
    

        if nfo.cell:
            try:
                cell = self.app.ActiveWindow.ActiveCell
            except:
                cell = None
            if not cell:
                self.fail()
            args['cell'] = cell

        if nfo.selection or nfo.cells or nfo.areas:
            try:
                #this fails if Selection is not a RangeSelection (e.g. shape, comment, etc.)
                #self.app.ActiveWindow.Selection.Cells -> problem: SheetSelectionChange does not trigger on selection of shapes
                selection = self.app.ActiveWindow.RangeSelection
            except:
                self.fail()

            if nfo.selection:
                args['selection'] = selection

            if nfo.cells:
                try:
                    #cells = list(iter(self.app.ActiveWindow.RangeSelection.Cells))  # crashs excel when too many cells are selected
                    cells = iter(selection.Cells)
                except:
                    cells = None
                if not cells:
                    self.fail()
                args['cells'] = cells

            if nfo.areas:
                try:
                    areas = list(iter(selection.Areas)) # TODO: test if this might cause performance issues if too many areas selected (like 'cells' argument)
                except:
                    areas = None
                if not areas:
                    self.fail()
                check_limits(areas, nfo.areas_min, nfo.areas_max)
                args['areas'] = areas
    
        return args






# ===========
# =  VISIO  =
# ===========

class AppContextVisio(AppContext):
    app_name = 'Microsoft Visio'
    
    def __init__(self, *args, **kwargs):
        
        super(AppContextVisio, self).__init__(*args, **kwargs)
        
        import bkt.library.visio as mod_visio
        self.mod_visio = mod_visio
        
    
    def resolve_app_arguments(self, nfo):
        if self.app.ActiveWindow is None:
            self.fail()

        args = {}
        if nfo.selection:
            args['selection'] = self.app.ActiveWindow.Selection

        if nfo.page:
            try:
                page = self.mod_visio.VisioPage(self.app.ActivePage)
            except:
                self.fail()
            args['page'] = page

        if nfo.shapes or nfo.shape:
            try:
                shapes  = [self.mod_visio.VisioShape(s) for s  in self.app.ActiveWindow.Selection]
            except:
                self.fail()

            check_limits(shapes, nfo.shapes_min, nfo.shapes_max)

            if nfo.shape:
                if len(shapes) != 1:
                    self.fail()
                args['shape'] = shapes[0]
            else:
                args['shapes'] = shapes
        
        if nfo.page_shapes:
            try:
                page_shapes = [self.mod_visio.VisioShape(s) for s in self.app.ActiveWindow.Page.Shapes]
            except:
                self.fail()
            args['page_shapes'] = page_shapes
        
        return args
        # except Exception, e:
        #     traceback.print_exc()
        #     raise InappropriateContextError(*e.args)




# ===========
# = HELPERS =
# ===========

def check_limits(collection, min_size=None, max_size=None):
    if min_size is not None and len(collection) < min_size:
        raise InappropriateContextError
    if max_size is not None and len(collection) > max_size:
        raise InappropriateContextError
