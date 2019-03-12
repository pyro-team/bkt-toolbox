# -*- coding: utf-8 -*-
'''
Created on 18.12.2012

@author: 802300
'''

from colors import *
from vis_calendar import *

class MonthContainer(object):
    def __init__(self,year,month):
        self.year = year
        self.month = month
        self.days = {}
    
    def flag(self,day,flag):
        if not flag in self.days:
            self.days[flag] = set()
        self.days[flag].add(day)
    
    def search(self,flag,unique=True):
        days = list(self.days.get(flag) or [])
        if unique:
            if len(days) != 1:
                raise ValueError
            return days[0]
        return days
    
    def __getattr__(self,attr):
        return self.search(attr)

class Mahnlauf(BaiscCalendarModel):
    due = 1
    mahnlauf = 4
    
    def _flag(self, day, flag):
        self.month_container.flag(day,flag)
        BaiscCalendarModel._flag(self, day, flag)
    
    def flag_custom(self):
        self.months = {}
        for month in range(1,13):
            self.month_container = MonthContainer(self.year, month)
            self.flag_month(month)
            self.months[(self.year,month)] = self.month_container
            
    def einreichung(self,beginn_vorlauf):
        bank,target,_ = self.calendars
        current = target.previous_workday(beginn_vorlauf)
        delta = datetime.timedelta(days=1)
        while not bank.is_workday(current):
            current -= delta
        return current

    def flag_month(self,month):
        due_main = self.flag_due(month)
        mahnlauf, mahnschreiben = self.flag_mahnlauf(due_main)
        z_sub = self.terminate_zwischeneinzug(due_main,mahnlauf,mahnschreiben)
        self.flag_reaction(mahnschreiben, z_sub)
        self.flag_zwischeneinzug(z_sub)
            
    def flag_due(self,month):
        _,target,_ = self.calendars
        flag = self._flag
        
        due_main = target.next_workday2(datetime.date(self.year,month,self.due))
        flag(due_main,'due_main')
        flag(due_main,'due')
        for d in (1,2,5):
            flag(due_main,'d%d' % d)
        
        vorlauf = {}
        for v in (1,2,5):
            v_day = target.relative_workday(due_main,-v)
            flag(v_day,'vorlauf')
            flag(v_day,'v%d' % v)
            
            s = self.einreichung(v_day)
            flag(s,'submission')
            flag(s,'s%d' % v)

            vorlauf[v] = v_day
        return due_main
    
    def flag_mahnlauf(self,due_main):
        bank,_,_ = self.calendars
        flag = self._flag
        
        mahnlauf = bank.relative_workday(due_main,self.mahnlauf)
        flag(mahnlauf,'mahnlauf')
        
        mahnschreiben = bank.relative_workday(mahnlauf,2)
        flag(mahnschreiben,'mahnschreiben')
        return mahnlauf, mahnschreiben
    
    def flag_reaction(self,mahnschreiben,z_sub):
        bank,_,_ = self.calendars
        flag = self._flag

        reaktion = mahnschreiben
        while reaktion < z_sub:
            flag(reaktion,'reaktion')
            reaktion = bank.next_workday(reaktion)
    
    def flag_zwischeneinzug(self,submission):
        return self.flag_zwischeneinzug_due_flex(submission)
    
    def flag_zwischeneinzug_due_flex(self,submission):
        bank,target,_ = self.calendars
        flag = self._flag

        z_sub = submission
        flag(z_sub,'z_sub')
        
        for d in (1,2,5):
            flag(z_sub,'z_s%d' % d)
            z_v1 = bank.relative_workday(z_sub,d)
            flag(z_v1,'z_v%d' % d)
            flag(z_v1,'z_vorlauf')
            
            z_due = target.next_workday(z_v1)
            flag(z_due,'z_due')
            flag(z_due,'z_d%d' % d)
            
        return z_sub

class MahnlaufZRelativ(Mahnlauf):
    zwischeneinzug = 3
    
    def terminate_zwischeneinzug(self,due_main,mahnlauf,mahnschreiben):
        bank,_,_ = self.calendars
        return bank.relative_workday(mahnschreiben,3)
        
class MahnlaufZSub15(Mahnlauf):
    def terminate_zwischeneinzug(self, due_main, mahnlauf, mahnschreiben):
        d = due_main
        if self.due == 1:
            sub = datetime.date(d.year,d.month,15)
        else:
            y, m = d.year,d.month
            y, m = inc_month(y, m)
            sub = datetime.date(y,m,1)
        bank,_,_ = self.calendars
        return bank.next_workday2(sub)

        
class MahnlaufZDCor1_15(Mahnlauf):
    def terminate_zwischeneinzug(self, due_main, mahnlauf, mahnschreiben):
        d = due_main
        bank,target,_ = self.calendars
        if self.due == 1:
            cor1due_raw = datetime.date(d.year,d.month,15)
        else:
            raise NotImplementedError
        cor1due = target.next_workday2(cor1due_raw)
        vorlauf = target.previous_workday(cor1due)
        submission = bank.previous_workday(vorlauf)
        return submission


#c_einzug = (192,0,0)
#c_einzug_last = (255,86,30)
#c_sepa = (255,192,0)
#c_mahn = (75,172,198)
c_zwischen_vorlauf = (191,191,191)
#c_zwischen_last = (96,96,96)
c_zwischen = (64,64,64)
#c_normal = (255,255,0)
#
#c_cycle_dark = (32,88,103)

fv_due = FlagViz(c_cycle_dark,tc_white,prio=6)
fv_vorlauf = FlagViz(c_cycle_light,tc_black,prio=5)
fv_mahn = FlagViz(c_cycle_alt,tc_black,prio=1)
fv_zw_due = FlagViz(c_due,tc_white,prio=9)
fv_zw_vorlauf = FlagViz(c_vorlauf,tc_black,prio=8)
fv_zw_sub = FlagViz(c_due_later,tc_white,prio=7)
fv_sub = FlagViz(c_cycle,tc_white,prio=4)
fv_zustellung = FlagViz(c_mahn_light,tc_black,prio=2)
fv_reaktion = FlagViz(c_mahn_extra_light,tc_black,prio=3)

color_flags = {'due_main' : fv_zw_due,
               'vorlauf' : fv_zw_vorlauf,
               'submission' : fv_zw_sub,
               'mahnlauf' : fv_mahn,
               'mahnschreiben' : fv_zustellung,
               'reaktion' : fv_reaktion,
               'z_vorlauf' : fv_vorlauf,
               'z_due' : fv_due,
               'z_sub' : fv_sub,
               #'submission_zwischen' : fv_zw_sub
               }

class MahnlaufVisualizer(Visualizer):
    def custom_decoration2(self, c):
        day, viz, rect = c.day, c.viz, c.shape
        print day
        drop = self.vc.drop
        
        def drop_dot(text=''):
            dot = drop('dot')
            dot.text = text
            dot.color = viz.txt_color
            dot.txt_color = viz.color
            return dot
        
        def num_dots(dproperty):
            count = 0
            for d in (1,2,5):
                if getattr(day,dproperty+str(d)):
                    count += 1
            return count
        
        def first_dot(dproperty):
            for d in (1,2,5):
                if getattr(day,dproperty+str(d)):
                    return str(d)
        
        def draw_dots(dproperty,dviz=None,targets=None):
            if dviz is None:
                dviz = viz
            #dproperty = 's'
            if targets is None:
                targets = ['origin','base_center','base_right']
            dots = ['5','2','1']
            for dot_text in dots:
                if not getattr(day,dproperty+dot_text):
                    continue
                dot = drop('dot')
                dot.text = dot_text
                dot.color = dviz.txt_color
                dot.txt_color = dviz.color
                dot.attach_to('anchor', rect, targets.pop())

        if (not day.workday) and day.target_workday:
            corner = drop('corner')
            corner.attach_to('anchor', rect, 'corner')
            corner.color = tc_black

        if day.vorlauf:
            draw_dots('v',self.color_flags['vorlauf'])
        
        if day.z_vorlauf:
            draw_dots('z_v',self.color_flags['z_vorlauf'])

        if day.submission:
            if num_dots('s') > 1:
                dot = drop_dot('+')
            else:
                dot = drop_dot(first_dot('s'))
            dot.attach_to('anchor', rect, 'left')

class MahnlaufVisualizer2(Visualizer):
    def custom_decoration2(self, c):
        day, viz, rect = c.day, c.viz, c.shape
        print day
        drop = self.vc.drop
        
        def drop_dot(text=''):
            dot = drop('dot')
            dot.text = text
            dot.color = viz.txt_color
            dot.txt_color = viz.color
            return dot
        
        def num_dots(dproperty):
            count = 0
            for d in (1,2,5):
                if getattr(day,dproperty+str(d)):
                    count += 1
            return count
        
        def first_dot(dproperty):
            for d in (1,2,5):
                if getattr(day,dproperty+str(d)):
                    return str(d)
        
        def draw_dot(dproperty,cflag):
            n = num_dots(dproperty)
            if n == 0:
                return
            elif n > 1:
                text = '+'
            else:
                text = first_dot(dproperty)
            
            if cflag == c.overlay_flag:
                dviz = c.overlay_viz
                target = 'right'
            else:
                dviz = c.viz
                target = 'left'
                
            dot = drop('dot')
            dot.text = text
            dot.color = dviz.txt_color
            dot.txt_color = dviz.color
            dot.attach_to('anchor', rect, target)
            
            
        if (not day.workday) and day.target_workday:
            corner = drop('corner')
            corner.attach_to('anchor', rect, 'corner')
            corner.color = tc_black
        
        if day.submission:
            draw_dot('s','submission')

        if day.vorlauf:
            draw_dot('v','vorlauf')

        if day.z_vorlauf:
            draw_dot('z_v','z_vorlauf')

        if day.z_sub:
            draw_dot('z_s','z_sub')

        if day.z_due:
            draw_dot('z_d','z_due')
            
class HeadlessSimulation(object):
    def __init__(self,start=2014,end=2018):
        self.years = range(start,end+1)
        self.years_raw = range(start-1,end+2)
        self.model_cls = MahnlaufZRelativ
    
    def sorted_months(self):
        return [t[1] for t in sorted(self.months.iteritems())]
    
    def foo(self):
        data = {}
        months = {}
        for year in self.years_raw:
            f = self.model_cls(year,data)
            f.flag_days()
            months.update(f.months)
            #flags_ = set()
            #for d in f.sorted_days():
            #    flags_.update(d.flags)
        self.months = dict(((y,m),c) for ((y,m),c) in months.iteritems() if y in self.years)
        print len(self.months)
        for m in self.sorted_months():
            print m.z_s5-m.mahnschreiben

def main_sim():
    HeadlessSimulation().foo()

def main_vis():
    #import vis_calendar
    #vis_calendar.VISIO_VISIBLE = True
    render_cal(MahnlaufVisualizer2, MahnlaufZSub15,
               #2013, 2015, m0=12, m1=1,
               start = '01.12.2013', end = '31.01.2015',
               color_flags=color_flags, style='wrapped',
               style_params=dict(alignment='independent',split_at=20))

if __name__ == '__main__':
    main_vis()
        