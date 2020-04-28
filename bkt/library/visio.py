# -*- coding: utf-8 -*-
'''
Created on 03.01.2013

@author: cschmitt
'''
import re
import traceback
import weakref
import os.path
import logging

from System.Runtime.InteropServices import COMException

from bkt import dotnet
Visio = dotnet.import_visio()

Units =  Visio.VisUnitCodes #@UndefinedVariable
Sections = Visio.VisSectionIndices #@UndefinedVariable
Rows = Visio.VisRowIndices #@UndefinedVariable
RowTags = Visio.VisRowTags #@UndefinedVariable
Cells = Visio.VisCellIndices #@UndefinedVariable

V_SCALE = 25.4
V_SCALE_I = 1.0/V_SCALE

U_MM = Units.visMillimeters
U_PT = Units.visPoints
U_RAD = Units.visRadians
U_NONE = Units.visNoCast
U_PERC = Units.visPercent
U_COLOR = Units.visUnitsColor

T_CONSTANT = 1
T_COLOR = 2
T_FORMULA = 4

def inch2mm(*args):
    if len(args) == 1:
        return args[0]*V_SCALE
    return tuple(v*V_SCALE for v in args)

def mm2inch(*args):
    if len(args) == 1:
        return args[0]*V_SCALE_I
    return tuple(v*V_SCALE_I for v in args)

def get_palette_color(doc,index):
    col = doc.Colors[index]
    return (int(col.Red), int(col.Green), int(col.Blue))

def is_integer(s):
    try:
        int(s)
        return True
    except ValueError:
        return False

def rgb_to_form(rgb, wrapper):
    return "RGB(%d,%d,%d)" % rgb

def form_to_rgb(form, wrapper):
    #print form
    if is_integer(form) and int(form) < 24:
        return get_palette_color(wrapper.shape.Document,int(form))
        
    gs = re.findall('[0-9]+',form)
    while len(gs) < 3:
        gs.append(0)
    return tuple([int(v) for v in gs[0:3]])

C_FONT = (T_CONSTANT, U_NONE,
          (Sections.visSectionCharacter,
           Rows.visRowCharacter,
           Cells.visCharacterFont)) #@UndefinedVariable

C_FONTSIZE = (T_CONSTANT, U_PT,
              (Sections.visSectionCharacter,
               Rows.visRowCharacter,
               Cells.visCharacterSize)) #@UndefinedVariable

C_FONTSIZE_MM = (T_CONSTANT, U_MM,
              (Sections.visSectionCharacter,
               Rows.visRowCharacter,
               Cells.visCharacterSize)) #@UndefinedVariable

C_FONTSTYLE = (T_CONSTANT, U_NONE,
               (Sections.visSectionCharacter,
                Rows.visRowCharacter,
                Cells.visCharacterStyle)) #@UndefinedVariable

C_TEXTCOLOR = (T_COLOR, (form_to_rgb, rgb_to_form),
               (Sections.visSectionCharacter,
                Rows.visRowCharacter,
                Cells.visCharacterColor)) #@UndefinedVariable

C_PAGEHEIGHT = (T_CONSTANT, U_MM,
                (Sections.visSectionObject,
                 Rows.visRowPage,
                 Cells.visPageHeight)) #@UndefinedVariable

C_PAGEWIDTH = (T_CONSTANT, U_MM,
                (Sections.visSectionObject,
                 Rows.visRowPage,
                 Cells.visPageWidth)) #@UndefinedVariable

C_HEIGHT = (T_CONSTANT, U_MM,
                (Sections.visSectionObject,
                 Rows.visRowXFormOut,
                 Cells.visXFormHeight)) #@UndefinedVariable

C_WIDTH = (T_CONSTANT, U_MM,
                (Sections.visSectionObject,
                 Rows.visRowXFormOut,
                 Cells.visXFormWidth)) #@UndefinedVariable

C_PINX = (T_CONSTANT, U_MM,
                (Sections.visSectionObject,
                 Rows.visRowXFormOut,
                 Cells.visXFormPinX)) #@UndefinedVariable

C_PINY = (T_CONSTANT, U_MM,
                (Sections.visSectionObject,
                 Rows.visRowXFormOut,
                 Cells.visXFormPinY)) #@UndefinedVariable

C_LPINX = (T_CONSTANT, U_MM,
                (Sections.visSectionObject,
                 Rows.visRowXFormOut,
                 Cells.visXFormLocPinX)) #@UndefinedVariable

C_LPINY = (T_CONSTANT, U_MM,
                (Sections.visSectionObject,
                 Rows.visRowXFormOut,
                 Cells.visXFormLocPinY)) #@UndefinedVariable

C_LPINX_FORMULA = (T_FORMULA, U_NONE,
                (Sections.visSectionObject,
                 Rows.visRowXFormOut,
                 Cells.visXFormLocPinX)) #@UndefinedVariable

C_LPINY_FORMULA = (T_FORMULA, U_NONE,
                (Sections.visSectionObject,
                 Rows.visRowXFormOut,
                 Cells.visXFormLocPinY)) #@UndefinedVariable

C_ANGLE = (T_CONSTANT, U_RAD,
                (Sections.visSectionObject,
                 Rows.visRowXFormOut,
                 Cells.visXFormAngle)) #@UndefinedVariable

C_SHAPECOLOR = (T_COLOR, (form_to_rgb, rgb_to_form),
               (Sections.visSectionObject,
                Rows.visRowFill,
                Cells.visFillForegnd)) #@UndefinedVariable

C_FILLPATTERN = (T_CONSTANT, U_NONE,
               (Sections.visSectionObject,
                Rows.visRowFill,
                Cells.visFillPattern)) #@UndefinedVariable

C_COLORTRANS = (T_CONSTANT, U_PERC,
               (Sections.visSectionObject,
                Rows.visRowFill,
                Cells.visFillForegndTrans)) #@UndefinedVariable

C_LINEWIDTH = (T_CONSTANT, U_PT,
               (Sections.visSectionObject,
                Rows.visRowLine,
                Cells.visLineWeight)) #@UndefinedVariable

C_LINEROUNDING = (T_CONSTANT, U_MM,
               (Sections.visSectionObject,
                Rows.visRowLine,
                Cells.visLineRounding)) #@UndefinedVariable

C_LINEPATTERN = (T_CONSTANT, U_NONE,
               (Sections.visSectionObject,
                Rows.visRowLine,
                Cells.visLinePattern)) #@UndefinedVariable

C_LINECOLOR = (T_COLOR, (form_to_rgb, rgb_to_form),
               (Sections.visSectionObject,
                Rows.visRowLine,
                Cells.visLineColor)) #@UndefinedVariable

#C_LINEROUND = (T_COLOR, (form_to_rgb, rgb_to_form),
#               (Sections.visSectionObject,
#                Rows.visRowLine,
#                Cells.visLineRounding)) #@UndefinedVariable

C_TOPMARGIN = (T_CONSTANT, U_MM,
               (Sections.visSectionObject,
                Rows.visRowText,
                Cells.visTxtBlkTopMargin)) #@UndefinedVariable

C_BOTTOMMARGIN = (T_CONSTANT, U_MM,
               (Sections.visSectionObject,
                Rows.visRowText,
                Cells.visTxtBlkBottomMargin)) #@UndefinedVariable

C_LEFTMARGIN = (T_CONSTANT, U_MM,
               (Sections.visSectionObject,
                Rows.visRowText,
                Cells.visTxtBlkLeftMargin)) #@UndefinedVariable

C_RIGHTMARGIN = (T_CONSTANT, U_MM,
               (Sections.visSectionObject,
                Rows.visRowText,
                Cells.visTxtBlkRightMargin)) #@UndefinedVariable

C_LINESPACING = (T_CONSTANT, U_PERC,
                  (Sections.visSectionParagraph,
                   Rows.visRowParagraph,
                   Cells.visSpaceLine))

C_VALIGN = (T_CONSTANT, U_NONE,
               (Sections.visSectionObject,
                Rows.visRowText,
                Cells.visTxtBlkVerticalAlign)) #@UndefinedVariable

C_BEGINX = (T_CONSTANT, U_MM,
               (Sections.visSectionObject,
                Rows.visRowXForm1D,
                Cells.vis1DBeginX)) #@UndefinedVariable

C_ENDX = (T_CONSTANT, U_MM,
               (Sections.visSectionObject,
                Rows.visRowXForm1D,
                Cells.vis1DEndX)) #@UndefinedVariable

C_BEGINY = (T_CONSTANT, U_MM,
               (Sections.visSectionObject,
                Rows.visRowXForm1D,
                Cells.vis1DBeginY)) #@UndefinedVariable

C_ENDY = (T_CONSTANT, U_MM,
               (Sections.visSectionObject,
                Rows.visRowXForm1D,
                Cells.vis1DEndY)) #@UndefinedVariable

C_TXT_ANGLE = (T_CONSTANT, U_RAD,
               (Sections.visSectionObject,
               Rows.visRowTextXForm,
               Cells.visXFormAngle))

C_TXT_WIDTH = (T_CONSTANT, U_MM,
               (Sections.visSectionObject,
               Rows.visRowTextXForm,
               Cells.visXFormWidth))

C_TXT_HEIGHT = (T_CONSTANT, U_MM,
               (Sections.visSectionObject,
               Rows.visRowTextXForm,
               Cells.visXFormHeight))

C_TXT_BACK =   (T_CONSTANT, U_NONE,
                 (Sections.visSectionObject,
                  Rows.visRowText,
                  Cells.visTxtBlkBkgnd))

C_END_ARROW = (T_CONSTANT, U_NONE,
               (Sections.visSectionObject,
               Rows.visRowLine,
               Cells.visLineEndArrow))

C_BEGIN_ARROW = (T_CONSTANT, U_NONE,
               (Sections.visSectionObject,
               Rows.visRowLine,
               Cells.visLineBeginArrow))

C_ROUTE_STYLE = (T_CONSTANT, U_NONE,
                (Sections.visSectionObject,
                 Rows.visRowShapeLayout,
                 Cells.visSLORouteStyle))
 
SHAPE_ATT_MAP = {
                 "width": C_WIDTH,
                 "height": C_HEIGHT,
                 "pagewidth": C_PAGEWIDTH,
                 "pageheight": C_PAGEHEIGHT,
                 "angle": C_ANGLE,
                 "font": C_FONT,
                 "fontsize": C_FONTSIZE,
                 "fontsize_mm": C_FONTSIZE_MM,
                 "fontstyle": C_FONTSTYLE,
                 "textcolor": C_TEXTCOLOR,
                 "color": C_SHAPECOLOR,
                 "fillpattern": C_FILLPATTERN,
                 "linecolor": C_LINECOLOR,
                 "linewidth": C_LINEWIDTH,
                 "stroke": C_LINEWIDTH,
                 "linerounding": C_LINEROUNDING,
                 "linestyle": C_LINEPATTERN,
                 "linepattern": C_LINEPATTERN,
                 "pinx": C_PINX,
                 "piny": C_PINY,
                 "x": C_PINX,
                 "y": C_PINY,
                 "localpinx": C_LPINX,
                 "localpiny": C_LPINY,
                 "localpinx_formula": C_LPINX_FORMULA,
                 "localpiny_formula": C_LPINY_FORMULA,
                 "tmargin": C_TOPMARGIN,
                 "bmargin": C_BOTTOMMARGIN,
                 "rmargin": C_RIGHTMARGIN,
                 "lmargin": C_LEFTMARGIN,
                 "lineround": C_LINEROUNDING,
                 "colortrans": C_COLORTRANS,
                 "beginx" : C_BEGINX,
                 "endx" : C_ENDX,
                 "beginy" : C_BEGINY,
                 "endy" : C_ENDY,
                 "txt_angle" : C_TXT_ANGLE,
                 "txt_width" : C_TXT_WIDTH,
                 "txt_height" : C_TXT_HEIGHT,
                 "txt_back" : C_TXT_BACK,
                 "txt_valign": C_VALIGN,
                 "txt_color": C_TEXTCOLOR,
                 "txt_linespace": C_LINESPACING,
                 "begin_arrow" : C_BEGIN_ARROW,
                 "end_arrow" : C_END_ARROW,
                 "route_style": C_ROUTE_STYLE,
                 }

#def unwrap_vs(f):
#    def _f(*args):
#        nargs = []
#        for arg in args:
#            if isinstance(arg, VisioShape):
#                arg = arg.shape
#            nargs.append(arg)
#        return f(*nargs)
#    return _f

class Wrapper(object):
    def __eq__(self,other):
        return self._raw == other._raw
    
    def __ne__(self,other):
        return not self.__eq__(other)
    
    def __hash__(self):
        return hash(self._raw)
        

def wrapper(name):
    def decorator(cls):
        def get_raw(self):
            return getattr(self,name)
        def set_raw(self,value):
            setattr(self,name,value)
        setattr(cls,'_raw',property(get_raw,set_raw))
        return cls
    return decorator

def unwrap(obj):
    if isinstance(obj, Wrapper):
        return obj._raw
    return obj

class WrapperUnavilableException(Exception):
    pass

class GlobalCache(object):
    def __init__(self):
        self.cache = weakref.WeakValueDictionary()
        
    def register(self,obj,wrapper):
        self.cache[obj] = wrapper
        
    def __getitem__(self,obj):
        wrapper = self.cache.get(obj)
        if wrapper is None:
            raise WrapperUnavilableException(obj)
        return wrapper
        
_global_cache = GlobalCache() 

class Rectangle(object):
    def __init__(self,x,y,width,height):
        self.x = x
        self.y = y
        self.width = width
        self.height = height
    
    @property
    def x1(self):
        return self.x + self.width
    
    @property
    def y1(self):
        return self.y + self.height
    
    def is_inside(self,rect):
        x0,y0 = rect.x, rect.y
        x1,y1 = rect.x1, rect.y1
        return (x0 <= self.x <= x1 and x0 <= self.x1 <= x1 and
                y0 <= self.y <= y1 and y0 <= self.y1 <= y1)

class static(object):
    @staticmethod
    def get_window(doc):
        app = doc.Application
        for window in app.Windows:
            if window.Document == doc:
                return window
        raise ValueError('no window found')

    @staticmethod
    def activate_page(page):
        page = unwrap(page)
        window = static.get_window(page.Document)
        window.Activate()
        window.Page = page
        
    @staticmethod
    def ensure_active(page):
        page = unwrap(page)
        if page == page.Application.ActivePage:
            return
        else:
            static.activate_page(page)
    
    @staticmethod
    def get_section_row(shape,section,name):
        shape = unwrap(shape)
        fq = section + '.' + name
        if not shape.CellExistsU(fq,0):
            raise KeyError(name,fq)
        cell = shape.CellsU(fq)
        return cell.ContainingRow
    
    @staticmethod
    def get_connection_row(shape,name):
        return static.get_section_row(shape,'Connections',name)
    
    @staticmethod
    def get_user_row(shape,name):
        return static.get_section_row(shape,'User',name)
    
    @staticmethod
    def get_control_row(shape,name):
        return static.get_section_row(shape,'Controls',name)
    
    @staticmethod
    def fix_pin(shape, left=0, bottom=0):
        old_x = shape._left
        old_y = shape._bottom
        shape.localpinx = left
        shape.localpiny = bottom
        shape._left = old_x
        shape._bottom = old_y

    @staticmethod
    def fix_pin_destructive(shape, left=0, bottom=0):
        shape.localpinx = left
        shape.localpiny = bottom

class ItemAsAttribute(object):
    def __getattr__(self,attr):
        if attr.startswith('__'):
            raise AttributeError(attr)
        try:
            return self.__getitem__(attr)
        except KeyError:
            raise AttributeError(attr)

class ShapeContext(ItemAsAttribute):
    def __init__(self,shape):
        self.shape = shape

class ConnectionRow(object):
    def __init__(self,shape,row):
        self.shape = shape
        self.row = row

    @property
    def xcell(self):
        return self.row.CellU(Cells.visX)
    
    @property
    def ycell(self):
        return self.row.CellU(Cells.visY)
    
    def _get_target(self,obj,target):
        if isinstance(obj, ConnectionRow):
            row = obj
        elif isinstance(obj, VisioShape):
            row = obj.points[target]
        else:
            raise NotImplementedError(obj)
        return row
    
    def glue_to(self,obj,target=None):
        row = self._get_target(obj, target)
        self.xcell.GlueTo(row.xcell)
        self.ycell.GlueTo(row.ycell)
        
    def connect_to(self,obj,connector=None,target=None):
        trow = self._get_target(obj, target)
        if connector is None:
            connector = self.shape.visio.app.ConnectorToolDataObject

        cs = self.shape.page.drop(connector)
        cs.cells.beginx.GlueTo(self.xcell)
        cs.cells.beginy.GlueTo(self.ycell)
        cs.cells.endx.GlueTo(trow.xcell)
        cs.cells.endy.GlueTo(trow.ycell)
        return cs

class ConnectionsContext(ShapeContext):
    def __getitem__(self,row_name):
        row = static.get_connection_row(self.shape, row_name)
        return ConnectionRow(self.shape,row)
    
class CellsContext(ShapeContext):
    def __getitem__(self,cell_name):
        try:
            _,_,src = SHAPE_ATT_MAP[cell_name]
            return self.shape.shape.CellsSRC(*src)
        except KeyError:
            return self._get_fallback(cell_name)
        
    def _get_fallback(self,cell_name):
        raw_shape = self.shape.shape
        if not raw_shape.CellExistsU(cell_name,0):
            raise KeyError(cell_name)
        return raw_shape.CellsU(cell_name)

class MasterDropCallable(object):
    def __init__(self,context,master):
        self.context = context
        self.master = master
    
    def __call__(self,x=0,y=0,fix_pin=False):
        return self.context(self.master,x,y,fix_pin)

class DropContext(ItemAsAttribute):
    def __init__(self,visio_page):
        self.visio_page = visio_page
        
    def __call__(self,obj,x=0,y=0,fix_pin=False):
        if isinstance(obj, basestring):
            drop_obj = self.visio_page.visio.masters[obj]
        elif isinstance(obj, VisioShape):
            drop_obj = obj.shape
        else:
            drop_obj = obj
        
        if drop_obj is None:
            raise ValueError('%r resolved to None' % obj)
        
        xi,yi = mm2inch(x,y)
        shape = VisioShape(self.visio_page.page.Drop(drop_obj,xi,yi))
        if fix_pin:
            static.fix_pin_destructive(shape)
        return shape
    
    def __getitem__(self,name):
        master = self.visio_page.visio.masters[name]
        return MasterDropCallable(self, master)

def _create_getter(name):
    def getter(self):
        access_type, spec, src = SHAPE_ATT_MAP[name]
        if access_type == T_COLOR:
            conv, _ = spec
            return conv(self.shape.CellsSRC(*src).ResultStr(U_COLOR),self)
        elif access_type == T_FORMULA:
            return self.shape.CellsSRC(*src).FormulaU
        else:
            return self.shape.CellsSRC(*src).Result(spec)
    return getter

def _create_setter(name):
    def setter(self,value):
        access_type, spec, src = SHAPE_ATT_MAP[name]
        if access_type == T_COLOR:
            _, conv = spec
            self.shape.CellsSRC(*src).FormulaForceU = conv(value,self) 
        elif access_type == T_FORMULA:
            self.shape.CellsSRC(*src).FormulaU = value
        else:
            self.shape.CellsSRC(*src).Result[spec] = value
    return setter

def check_angle(method):
    def _method(self,*args,**kwargs):
        s = self
        if s.angle != 0:
            #raise ValueError('shape has angle != 0: angle=%r, sw=%r' % (s.angle,s))
            logging.warning('shape has angle != 0: angle=%r, sw=%r. Using Bounding-Box.' % (s.angle,s))
        return method(self,*args,**kwargs)
    return _method


class PropertyAccessor(ShapeContext):
    def __getitem__(self, attr):
        try:
            cell = unwrap(self.shape).CellsU["Prop." + attr]
            return cell.ResultStrU(0)
        except COMException:
            raise KeyError(attr)

@wrapper("shape")    
class VisioShape(Wrapper):
    def __init__(self, shape):
        self.shape = shape
        self.points = ConnectionsContext(self)
        self.cells = CellsContext(self)
        self.shape_data = PropertyAccessor(self)
        
    @property
    def page(self):
        page = self.shape.ContainingPage
        try:
            return _global_cache[page]
        except WrapperUnavilableException:
            return VisioPage(page)

    @property
    def visio(self):
        return self.page.visio

    ### Special getter/setter _x, _y, _width, _height that consider 1D-Shapes (connectors) ###

    def get_x_save(self):
        if self.shape.OneD == 0:
            return self.x
        else:
            return self.beginx
    
    def set_x_save(self,x):
        if self.shape.OneD == 0:
            self.x = x
        else:
            self.endx = self.endx + (x-self.beginx)
            self.beginx = x

    def get_y_save(self):
        if self.shape.OneD == 0:
            return self.y
        else:
            return self.beginy
    
    def set_y_save(self,y):
        if self.shape.OneD == 0:
            self.y = y
        else:
            self.endy = self.endy + (y-self.beginy)
            self.beginy = y

    def get_width_save(self):
        if self.shape.OneD == 0:
            return self.width
        else:
            return self.endx - self.beginx
    
    def set_width_save(self,width):
        if self.shape.OneD == 0:
            self.width = width
        else:
            self.endx = self.beginx + width

    def get_height_save(self):
        if self.shape.OneD == 0:
            return self.height
        else:
            return self.endy - self.beginy
    
    def set_height_save(self,height):
        if self.shape.OneD == 0:
            self.height = height
        else:
            self.endy = self.beginy + height
        
    _x      = property(get_x_save,set_x_save)
    _y      = property(get_y_save,set_y_save)
    _width  = property(get_width_save,set_width_save)
    _height = property(get_height_save,set_height_save)
    del get_x_save
    del set_x_save
    del get_y_save
    del set_y_save
    del get_width_save
    del set_width_save
    del get_height_save
    del set_height_save

    ### Special getter/setter _left and _bottom that always return the left bottom corner ###

    # @check_angle
    def get_x_real(self):
        if self.shape.OneD == -1:
            return min(self.beginx, self.endx)
        elif self.angle != 0:
            bb = self.bounding_box
            return bb.x
        else:
            return self.x - self.localpinx

    # @check_angle
    def set_x_real(self,x):
        if self.shape.OneD == -1:
            if self.beginx < self.endx:
                self.endx = x + self._width
                self.beginx = x
            else:
                self.beginx = x - self._width
                self.endx = x
        elif self.angle != 0:
            bb = self.bounding_box
            self.x = x + (self.x - bb.x)
        else:
            self.x = x + self.localpinx
        
    # @check_angle
    def get_y_real(self):
        if self.shape.OneD == -1:
            return min(self.beginy, self.endy)
        elif self.angle != 0:
            bb = self.bounding_box
            return bb.y
        else:
            return self.y - self.localpiny
        
    # @check_angle
    def set_y_real(self,y):
        if self.shape.OneD == -1:
            if self.beginy < self.endy:
                self.endy = y + self._height
                self.beginy = y
            else:
                self.beginy = y - self._height
                self.endy = y
        elif self.angle != 0:
            bb = self.bounding_box
            self.y = y + (self.y - bb.y)
        else:
            self.y = y + self.localpiny
        
    _left   = property(get_x_real,set_x_real)
    _bottom = property(get_y_real,set_y_real)
    del get_x_real
    del set_x_real
    del get_y_real
    del set_y_real
    
    ### Bounding box calculation ###

    @property
    def bounding_box(self):
        out = [clr.Reference[float]() for _ in range(4)]
        self.shape.BoundingBox(8192+4, *out) #visBBoxDrawingCoords + visBBoxExtents
        # Round due to small errors in converting numbers
        dblLeft = round(inch2mm(out[0].Value),5)
        dblBottom = round(inch2mm(out[1].Value),5)
        dblRight = round(inch2mm(out[2].Value),5)
        dblTop = round(inch2mm(out[3].Value),5)
        return Rectangle(dblLeft, dblBottom, dblRight-dblLeft, dblTop-dblBottom)
        #return Rectangle(self._x,self._y,self.width,self.height)
    
    @property
    def type(self):
        return int(self.shape.Type)
    
    def get_text(self):
        return self.shape.Text
    
    def set_text(self,value):
        self.shape.Text = value
        
    text = property(get_text,set_text)
    del get_text
    del set_text
    
    @property
    def shapes(self):
        return [VisioShape(c) for c in self.shape.Shapes]
    
    def group(self,*shapes):
        return self.page.group(self,*shapes)
    
    def ungroup(self):
        s = self.shape
        if s.Type != 2:
            raise ValueError('shape type must be visTypeGroup (2)')
        contained = self.shapes
        s.Ungroup()
        return contained
            
    def attach_to(self,anchor,parent,target):
        self.points[anchor].glue_to(parent.points[target])
        
    def connect_dynamic(self,other,connector=None):
        if connector is None:
            connector = self.visio.app.ConnectorToolDataObject
        cs = self.page.drop(connector)
        cs.cells.beginx.GlueTo(self.cells.x)
        cs.cells.beginy.GlueTo(self.cells.y)
        cs.cells.endx.GlueTo(other.cells.x)
        cs.cells.endy.GlueTo(other.cells.y)
            
    def __repr__(self):
        try:
            master = self.shape.Master
            if master:
                master = master.Name
            return '<ShapeWrapper: shape_id=%r, name=%r, master=%r>' % (self.shape.ID,self.shape.name,master)
        except:
            traceback.print_exc()
            return object.__repr__(self)

def decorate_wrapper(cls):
    for name in SHAPE_ATT_MAP:
        fget = _create_getter(name)
        fset = _create_setter(name)
        setattr(cls,name,property(fget,fset))

decorate_wrapper(VisioShape)
del decorate_wrapper

def draw_method(name):
    def draw(self,x,y,w,h):
        method = getattr(self.page,name)
        shape = method(*mm2inch(x,y,x+w,y+h))
        return VisioShape(shape)
    return draw


def page_property(name):
    def getter(self):
        return getattr(self.wrapped,name)
    def setter(self,value):
        setattr(self.wrapped,name,value)
    return property(getter,setter)

@wrapper("page")    
class VisioPage(Wrapper):
    def __init__(self,page):
        self.page = page
        self.wrapped = VisioShape(page.PageSheet)
        self.drop = DropContext(self)
        _global_cache.register(page,self)
        
    drawRect = draw_method('drawRectangle')

    width = page_property('pagewidth')
    height = page_property('pageheight')
    
    def group_all(self):
        self.group(*(s for s in self.page.Shapes))
    
    def group(self,*shapes):
        static.ensure_active(self)
        window = self.page.Application.ActiveWindow
        window.DeselectAll()
        for shape in shapes:
            window.Select(unwrap(shape),2)
        return VisioShape(window.Selection.Group())
    
    @property
    def visio(self):
        app = self.page.Application
        try:
            return _global_cache[app]
        except WrapperUnavilableException:
            print 'Warning: Wrapper for current application not found, references to loaded stencils lost.'
            return VisioWrapper(app)
    
    @property
    def doc(self):
        return self.page.Document
    
    @property    
    def fontmap(self):
        if self.visio.fontmap is None:
            self.visio.create_fontmap(self.doc)
        return self.visio.fontmap

@wrapper("app")    
class VisioWrapper(Wrapper):
    def __init__(self,app=None,visible=False):
        if app is None:
            if not visible:
                app = Visio.InvisibleAppClass()
            else:
                app = Visio.ApplicationClass()
        self.app = app
        self.masters = {}
        self.stencils = {}
        self.fontmap = None
        _global_cache.register(app,self)
        
    def create_fontmap(self,doc):
        self.fontmap = dict((font.Name,font.ID) for font in doc.Fonts)
        
    def new_page(self,name=None,doc=None):
        if doc is None:
            doc = self.app.Documents.Add('')
            page = doc.Pages(1)
        else:
            page = doc.Pages.Add()
        if name is not None:
            page.Name = name
        page.AutoSize = False
        return VisioPage(page)
    
    def load_stencil(self,stencil_path):
        stencil_path = os.path.abspath(stencil_path)
        stencil_doc = self.app.Documents.Add(stencil_path)
        masters = {}
        for m in stencil_doc.Masters:
            masters[m.Name] = m
        self.stencils[stencil_path] = stencil_doc
        self.masters.update(masters)
        return stencil_doc,masters
    
    def close_stencils(self):
        while self._stencils_open():
            for window in [w for w in self.app.Windows]:
                if window.Type == 2:
                    print 'closing %s' % window.Caption
                    window.Close()
                    
    def _stencils_open(self):
        for window in self.app.Windows:
            if window.Type == 2:
                return True
        return False
    
class VisioContext(object):
    def __init__(self,visible=False,page_name=None,page_size=None):
        self.visio = VisioWrapper(visible=visible)
        self.page = self.visio.new_page(page_name)
        if page_size is not None:
            w,h = page_size
            self.page.width = w
            self.page.height = h
        
def page(*args,**kwargs):
    return VisioContext(*args,**kwargs).page