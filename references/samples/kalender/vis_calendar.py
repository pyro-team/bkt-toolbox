# -*- coding: utf-8 -*-
'''
Created on 14.11.2012

@author: 802300
'''
import os.path
import datetime

import work_calendar
#from vgen.visio import VisioCanvas #, ShapeWrapper
from bkt.library.visio import VisioWrapper #, ShapeWrapper

VISIO_VISIBLE = False
class Geometry(object):
    '''
    Defines the geometry of the calendar.
    base_width is the edge length of the squares that represent the days of the calendar.
    base_width MUST equal the size of the corresponding Visio shape. 
    '''
    base_width = 10.0
    padding_factor = 1.1
    base_padded = base_width*padding_factor
    padding = base_padded-base_width

def list_year(year):
    lyear = [datetime.date(year,1,1)]
    delta = datetime.timedelta(days=1)
    while True:
        next_day = lyear[-1] + delta
        if next_day.year == year:
            lyear.append(next_day)
        else:
            break
    return lyear

def inc_month(y,m):
    if m < 0 or m > 12:
        raise ValueError
    if m == 12:
        return y+1,1
    else:
        return y,m+1

def dec_month(y,m):
    if m < 0 or m > 12:
        raise ValueError
    if m == 1:
        return y-1,12
    else:
        return y,m-1

def last_day(y,m):
    d = 28
    delta = datetime.timedelta(days=1)
    while True:
        date = datetime.date(y,m,d)+delta
        if date.month != m:
            return d
        d = date.day
    raise AssertionError

class MultipleFlags(Exception):
    pass

class FlaggedDate(object):
    def __init__(self,date):
        self.date = date
        self.flags = {}
        
    def setflag(self,flag):
        self.flags[flag] = True
        
    def hasflag(self,flag):
        return self.flags[flag]
    
    def removeflag(self,flag):
        if not flag in self.flags:
            return
        del self.flags[flag]
        
    def __getattr__(self,flag):
        if flag not in self.flags:
            return False
        return self.flags[flag]
    
    def __repr__(self):
        return '<%s %s>' % (self.date,'|'.join(sorted(self.flags)))

class BaiscCalendarModel(object):
    def __init__(self,year,data):
        self.year = year
        self.year_data = data
        for d in list_year(year):
            if not d in data:
                data[d] = FlaggedDate(d)
    
    def sorted_days(self):
        return [self.year_data[k] for k in sorted(self.year_data) if k.year == self.year]

    @property
    def calendars(self):
        bank = work_calendar.Calendar(work_calendar.BANK)
        target = work_calendar.TargetCalendar()
        hessen = work_calendar.HessenCalendar()
        return (bank,target,hessen)

    def flag_days(self):
        bank, target, hessen = self.calendars
        flag = self._flag
        
        self.flag_custom()

        for day in self.year_data.itervalues():
            if bank.is_workday(day.date):
                flag(day,'workday')
            if target.is_workday(day.date):
                flag(day,'target_workday')
            if hessen.is_workday(day.date):
                flag(day,'hessen_workday')
        
    def _flag(self,day,flag):
        if isinstance(day, FlaggedDate):
            day = day.date
        if not day in self.year_data:
            self.year_data[day] = FlaggedDate(day)
        self.year_data[day].setflag(flag)

tc_white = (255,255,255)
tc_gray = (128,128,128)
tc_black = (0,0,0)
tc_pale_red = (192,80,70)
                
class FlagViz(object):
    def __init__(self,color,txt_color,prio=0):
        self.color = color
        self.txt_color = txt_color
        self.prio = prio

class RowLayoutInformation(object):
    def __init__(self,year,index,month,days,elements):
        self.year = year
        self.month = month
        self.index = index
        self.days = days
        self.elements = elements
    
    def remove_positions(self,positions):
        shift = Geometry.base_width
        positions = list(reversed(sorted(positions)))
        
        def shift_positions():
            for i in range(len(positions)):
                positions[i] -= shift
        
        while positions:
            rem = positions.pop()
            print 'deleting position %s' % rem
            shift_positions()
            new_days = [d for d in self.days if d.x != rem]
            
            #if len(self.days) == len(new_days):
            #    raise KeyError('no day found at position %s found' % rem)
            
            self.days = new_days
            self.shift_left(rem)
        
    def shift_left(self,position):
        for elem in self.days+self.elements:
            if elem.x < position:
                continue
            elem.x -= Geometry.base_width
        
class LayoutElement(object):
    def __init__(self,callback,x,y):
        self.callback = callback
        self.x = x
        self.y = y
    
    def draw(self,visualizer):
        return self.callback(visualizer,self.x,self.y)

class DayLayoutInformation(object):
    def __init__(self,day,x,y):
        self.day = day
        self.x = x
        self.y = y
    
class CalendarRow(object):
    month_shape_width = 2.5*Geometry.base_width

    def __init__(self,year,month):
        self.year = year
        self.month = month
        self.days = []

    def get_width(self):
        raise NotImplementedError

    def get_height(self):
        return Geometry.base_padded
    
    def draw_month(self,visualizer,x,y,year=None,month=None):
        if year is None:
            year = self.year
        if month is None:
            month = self.month
        rect = visualizer.vc.drawRect(x,y,self.month_shape_width,Geometry.base_width)
        rect.text = '%02d/%04d' % (month,year)
        rect.font = visualizer.arial
        if year%2 == 0:
            rect.color = 200,200,200
        return rect

    def choose_days(self,flagged_days):
        # flagged_days should be sorted
        self.days = [d for d in flagged_days if d.date.year == self.year and d.date.month == self.month]
    
class LinearMonth(CalendarRow):
        
    def do_layout(self,xoff,yoff):
        self.y = yoff
        size = Geometry.base_width
        
        def day_layout(day_,x=0,y=0):
            x += xoff
            y += yoff
            return DayLayoutInformation(day_,x,y)
        
        day_layout_list = [day_layout(day,size*(day.date.day-1)) for day in self.days]
        draw_month = LayoutElement(self.draw_month, 31.5*size, yoff)
        
        return day_layout_list,[draw_month]
        
    def get_width(self):
        return 31.5*Geometry.base_width + self.month_shape_width

def wrapped_month_factory(**params):
    def factory(year,month):
        return WrappedMonth(year,month,**params)
    return factory

class WrappedMonth(CalendarRow):
    def __init__(self,year,month,split_at=16,alignment='right',draw_separator=True,split_spacing=None):
        CalendarRow.__init__(self,year,month)

        if split_spacing is None:
            split_spacing = 1.5*Geometry.base_width
        
        self.split_at = split_at
        self.split_spacing = split_spacing

        self.alignment = alignment
        self.draw_separator = draw_separator
        
    def choose_days(self,flagged_days):
        ya,ma = self.year,self.month
        yp,mp = dec_month(ya, ma)
        
        def accept(date):
            yc,mc,dc = date.year,date.month,date.day
            return ((yc,mc) == (ya,ma) and dc < self.split_at) or ((yc,mc) == (yp,mp) and dc >= self.split_at)
        
        self.days = [d for d in flagged_days if accept(d.date)]
        
    def get_alignment_offsets(self):
        size = Geometry.base_width
        year,month = self.year,self.month
        last_previous = last_day(*dec_month(year, month))
        
        if self.alignment == 'independent':
            al = 0
            ar = (31-self.split_at+1)*size + self.split_spacing
            lm = (last_previous-self.split_at+1)*size
        
        elif self.alignment == 'left':
            al = 0
            lm = (last_previous-self.split_at+1)*size
            ar = lm + self.split_spacing
        
        elif self.alignment == 'right':
            ar = (31-self.split_at+1)*size + self.split_spacing
            lm = ar - self.split_spacing
            al = lm - (last_previous-self.split_at+1)*size
        
        else:
            raise AssertionError
        
        return (al,ar,lm)
            
            
    def iter_layout(self,xoff,yoff):
        size = Geometry.base_width
        
        xtra = self.month_shape_width + 0.5*size
        left_align, right_align, _ = self.get_alignment_offsets()
        
        run1 = [day for day in self.days if day.date.month != self.month]
        for day in run1:
            dn = day.date.day
            x = xtra + (dn-self.split_at)*size + left_align
            yield day,x
            
        run2 = [day for day in self.days if day.date.month == self.month]
        for day in run2:
            dn = day.date.day
            x = xtra + (dn-1)*size + right_align
            yield day,x
        
    def do_layout(self,xoff,yoff):
        ####### callbacks
        ### left month shape
        def draw_left(viz,x,y):
            py,pm = dec_month(self.year, self.month)
            return self.draw_month(viz, x, y, py, pm)
        
        cb_left = LayoutElement(draw_left, 0, yoff)
        
        ### right month shape
        cb_right_x = 32*Geometry.base_width + self.month_shape_width + self.split_spacing
        cb_right = LayoutElement(self.draw_month, cb_right_x, yoff)

        ### separator
        def draw_separator(viz,x,y):
            if not self.draw_separator:
                return None
            sep = viz.drop_nofix('pointing_triangle')
            sep._y = y
            sep.x = x
            return sep

        _,_,left_max = self.get_alignment_offsets()
        xtra = self.month_shape_width + 0.5*Geometry.base_width
        sep_x = xtra + left_max + 0.5*self.split_spacing
        cb_sep = LayoutElement(draw_separator,sep_x,yoff)
        
        ####### layout information for days
        days = [DayLayoutInformation(day, x, yoff) for day, x in self.iter_layout(xoff, yoff)]
        
        return days,[cb_left,cb_right,cb_sep]
        
    def get_width(self):
        return 32*Geometry.base_width + 2*self.month_shape_width + self.split_spacing

class DefaultColumnFilter(object):
    def get_deletions(self,columns):
        exclude = set(['hessen_workday','workday','target_workday'])
        mask = []
        for column in columns:
            flags = self.get_flags(column) - exclude
            if len(flags) == 0:
                mask.append(True)
            else:
                mask.append(False)
        return list(self.get_start(mask)) + list(self.get_end(mask))
                
    def get_start(self,mask):
        for i in range(len(mask)-1):
            if not mask[i+1]:
                break
            yield i

    def get_end(self,mask):
        n = len(mask)
        for i in range(n-1):
            pos = n-i-1
            if not mask[pos-1]:
                break
            yield pos
    
    def get_flags(self,column):
        flags = set()
        for day in column:
            flags.update(day.flags)
        return flags

class CustomDecorationContext(object):
    def __init__(self,day,shape,flag,viz,overlay_flag=None,overlay_viz=None):
        self.day = day
        self.shape = shape
        self.flag = flag
        self.viz = viz
        self.overlay_flag = overlay_flag
        self.overlay_viz = overlay_viz

class Visualizer(object):
    workday_viz = FlagViz(tc_white, tc_black)
    non_workday_viz = FlagViz(tc_white, tc_gray)
    holyday_viz = FlagViz(tc_white, tc_pale_red)

    def __init__(self,days,color_flags):
        self.days = days
        self.color_flags = color_flags

        months = set((d.date.year,d.date.month) for d in days)
        self.months = dict((m,i+1) for i,m in enumerate(sorted(months)))
        self.num_months = len(self.months)

        self.source_dir = os.path.dirname(os.path.abspath(__file__))

    def all_color_flags(self,day):
        return sorted(((flag,self.color_flags[flag]) for flag in sorted(self.color_flags) if getattr(day,flag)),key=lambda f : f[1].prio)
    
    def unique_color_flag(self,day):
        flags = [flag for flag in self.color_flags if getattr(day,flag)]
        if len(flags) == 1:
            f = flags[0]
            return f ,self.color_flags[f]
        if len(flags) == 0:
            return None, self.get_default_viz(day)
        else:
            raise MultipleFlags(*flags)
        
    def get_default_viz(self,day):
        if day.workday:
            return self.workday_viz
        else:
            if day.date.weekday() in (5,6):
                return self.non_workday_viz
            else:
                return self.holyday_viz
            
    def drop_nofix(self,master):
        sw = self.vc.drop(master,fix_pin=False)
        return sw
    
    def drop_at(self,master,x=0,y=0):
        sw = self.vc.drop(master,fix_pin=False)
        sw._x = x
        sw._y = y
        return sw
    
    def render_day(self,day_layout):
        day = day_layout.day
        drop = self.drop_nofix

        rect = self.drop_at('back',day_layout.x,day_layout.y)
        try:
            color_flag, color_flag_viz = self.unique_color_flag(day)
            overlay_flag, overlay_flag_viz = None, None
        except MultipleFlags, e:
            color_flag, color_flag_viz, overlay_flag, overlay_flag_viz = self.resolve_color_conflict(day,e.args,rect)

        rect.color = color_flag_viz.color
        
        txt = drop('number')
        txt.attach_to('anchor', rect, 'day')
        
        txt.text = str(day.date.day)
        txt.txt_color = color_flag_viz.txt_color
        
        wd = day.date.weekday()
        if wd in (5,6):
            dn = drop('wide_text')
            dn.attach_to('anchor', rect, 'top_dot')
            if wd == 5:
                dn.text = 'Sa'
                dn.txt_color = color_flag_viz.txt_color
            elif wd == 6:
                dn.text = 'So'
                dn.txt_color = tc_pale_red
        
        try:
            context = CustomDecorationContext(day, rect, color_flag, color_flag_viz, overlay_flag, overlay_flag_viz)
            self.custom_decoration2(context)
        except (NotImplementedError, AttributeError):
            self.custom_decoration(day,rect,color_flag_viz)

        frame = drop('frame')
        frame.attach_to('anchor', rect, 'origin')
            
    def custom_decoration(self,day,rect,viz):
        pass
            
    def resolve_color_conflict(self,day,flags,rect):
        fv = self.all_color_flags(day)
        if len(fv) == 2:
            (f1,fv1), (f2,fv2) = fv
            rect.color = fv1.color
            
            half = self.drop_nofix('half_back')
            half.color = fv2.color
            half.attach_to('anchor', rect, 'origin')
            return f1, fv1, f2, fv2
        else:
            print 'warning: unresolved color conflict for %r' % day
            return self.get_default_viz(day)

    def vis_new(self,name,style=None,**style_params):
        if style is None or style == 'linear':
            row_constructor = LinearMonth
        elif style == 'wrapped':
            row_constructor = wrapped_month_factory(**style_params)
        else:
            raise ValueError
        
        layout_data = []
        y_offset = 0
        width = 0
        
        for i,(year,month) in enumerate(reversed(sorted(self.months))):
            print 'processing month %04d-%02d, index %d' % (year,month,i)
            row = row_constructor(year,month)
            row.choose_days(self.days)
            days, elements = row.do_layout(0, y_offset)
            layout_data.append(RowLayoutInformation(year, month, i, days, elements))

            y_offset += row.get_height()
            width = max(width,row.get_width())
            
        height = y_offset
        self.dimension = (width,height)
        self.layout_data = layout_data
        
        self.filter_columns(DefaultColumnFilter())
        
        self.render(name)
        
    def get_all_day_layouts(self):
        res = []
        for row in self.layout_data:
            res.extend(row.days)
        return res
    
    def filter_columns(self,column_filter):
        day_layouts = self.get_all_day_layouts()
        col_positions = sorted(set([l.x for l in day_layouts]))
        
        columns = []
        by_index = {}
        for index,position in enumerate(col_positions):
            by_index[index] = position
            column = [l.day for l in day_layouts if l.x == position]
            columns.append(column)
            
        deletion_indices = column_filter.get_deletions(columns)
        deletion_positions = [by_index[i] for i in deletion_indices]
        
        print 'column indices to be deleted: %r' % deletion_indices
        print 'column positions to be deleted: %r' % deletion_positions
                
        for row in self.layout_data:
            row.remove_positions(deletion_positions)
        
    def render(self,name):
        visio = VisioWrapper()
        self.vc = visio.new_page(name)
        p = self.vc
        p.visio.app.Visible = VISIO_VISIBLE
        p.visio.load_stencil(os.path.join(self.source_dir,'stencil.vss'))
        p.width, p.height = self.dimension

        self.arial = p.fontmap['Arial']
        
        try:
            for row in self.layout_data:
                for day in row.days:
                    self.render_day(day)
                for elem in row.elements:
                    elem.draw(self)
        finally:
            self.vc.visio.app.Visible = True
            self.vc.visio.close_stencils()

def render_cal(vis_cls,model_cls,start='01.01.2014',end='31.12.2016',color_flags=None,style=None,style_params=None,no_vis=False,**kwargs):
    def parse_date(ds):
        d,m,y = (int(p) for p in ds.split('.'))
        return datetime.date(y,m,d)
    
    first = parse_date(start)
    last = parse_date(end)
    y0 = first.year
    y1 = last.year
    
    name = model_cls.__name__
    years = list(range(y0-1,y1+2))
    
    data = {}
    
    for year in years:
        f = model_cls(year,data,**kwargs)
        f.flag_days()
        
        flags_ = set()
        for d in f.sorted_days():
            flags_.update(d.flags)
    
    #first = datetime.date(y0,m0,1)
    #last = datetime.date(y1,m1,last_day(y1, m1))
    days = [data[d] for d in sorted(data) if first <= d <= last]
    
    if no_vis:
        return
    
    if style_params is None:
        style_params = {}
    v = vis_cls(days,color_flags)
    v.vis_new(name,style,**style_params)
