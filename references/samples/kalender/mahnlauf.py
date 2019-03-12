# -*- coding: utf-8 -*-
'''
Created on 29.11.2012

@author: 802300
'''

from vis_calendar import *

class Mahnlauf(BaiscCalendarModel):
    flag_latest = True
    cal = work_calendar.Calendar(work_calendar.BANK)
    target = work_calendar.TargetCalendar()
    
    def __init__(self, year, data, einzug=1):
        BaiscCalendarModel.__init__(self, year, data)
        if einzug not in (1,15):
            raise ValueError
        self.einzug = einzug
    
    def respect_latest(self,zwischen):
        return self.flag_latest
    
    def flag_blocked(self,due,cor1sub,d5sub=None):
        delta = datetime.timedelta(days=1)
        
        if d5sub is not None:
            day = d5sub
            while day < cor1sub:
                self._flag(day, 'partially_blocked')
                day += delta
        
        day = cor1sub
        end = self.cal.relative_workday(due, 3)
        while day <= end:
            self._flag(day, 'blocked')
            day += delta

    def flag_pblocked(self,zwischen,esub):
        delta = datetime.timedelta(days=1)
        day = esub
        while day <= zwischen:
            self._flag(day, 'potentially_blocked')
            day += delta
    
    def flag_custom(self):
        cal,target,_ = self.calendars
        flag = self._flag
        
        for month in range(1,13):
            u0 = cal.ultimo(self.year, month, 0)
            u1 = cal.ultimo(self.year, month, 1)
            for u in (u0,u1):
                flag(u,'mahnlauf_dummy')
                
            d_einzug = self.einzug

            einzug = target.next_workday2(datetime.date(self.year,month,d_einzug))
            flag(einzug,'due_main')
            flag(einzug,'due')
            
            zwischen = self.tag_zwischeneinzug(month, d_einzug, einzug)
            flag(zwischen,'due_zwischen')
            flag(einzug,'due')

            mahn_de = self.tag_mahnlauf(month, d_einzug, einzug)
            flag(mahn_de,'mahnlauf')
            
            cor1sub = None
            d5sub = None
            
            for due,flag_prefix,ist_zwischen in ((einzug,'main',False),(zwischen,'zwischen',True)):
                for delay in (1,2,5):
                    if not self.respect_delay(delay, ist_zwischen):
                        continue
                    sub = target.latest_submission(due, delay)
                    while not cal.is_workday(sub):
                        sub = cal.previous_workday(sub)
                    flag(sub,'submission')
                    flag(sub, 's%d' % delay)
                    flag(sub, 'submission_%s' % flag_prefix)

                    if not ist_zwischen:
                        if delay == 1:
                            cor1sub = sub
                        elif delay == 5:
                            d5sub = sub
                    else:
                        esub = sub
            
            if cor1sub is None:
                raise AssertionError
            
            self.flag_blocked(einzug, cor1sub, d5sub)
            self.flag_pblocked(zwischen, esub)
                    
                
class MahnlaufIst(Mahnlauf):
    flag_latest = False
    
    def tag_mahnlauf(self,monat,einzugstermin,due_date):
        return self.cal.relative_workday(due_date, 4)
        
    def tag_zwischeneinzug(self,monat,einzugstermin,due_date):
        if einzugstermin not in (1,15):
            raise AssertionError
        if einzugstermin == 1:
            d_zw = 15
        else:
            d_zw = 1
        return self.cal.next_workday2(datetime.date(self.year,monat,d_zw))
        #return d_zw
        
    def respect_delay(self,delay,zwischen):
        return False
        if zwischen:
            return delay in (1,2)
        return True
                
class MahnlaufSEPANaiv(Mahnlauf):
    def tag_mahnlauf(self,monat,einzugstermin,due_date):
        return self.cal.relative_workday(due_date, 4)
        
    def tag_zwischeneinzug(self,monat,einzugstermin,due_date):
        year = self.year
        if einzugstermin not in (1,15):
            raise AssertionError
        if einzugstermin == 1:
            d_zw = 15
        else:
            d_zw = 1
            year, monat = inc_month(year, monat)
        return self.cal.next_workday2(datetime.date(year,monat,d_zw))
        
    def respect_delay(self,delay,zwischen):
        return True

class MahnlaufSEPAKx(Mahnlauf):
    def tag_mahnlauf(self,monat,einzugstermin,due_date):
        return self.cal.relative_workday(due_date, 4)
        
    def tag_zwischeneinzug(self,monat,einzugstermin,due_date):
        year = self.year
        if einzugstermin not in (1,15):
            raise AssertionError
        if einzugstermin == 1:
            d_zw = self.e15_d_zw
        else:
            d_zw = self.e01_d_zw
            year, monat = inc_month(year, monat)
        return self.cal.next_workday2(datetime.date(year,monat,d_zw))
        #return d_zw
        
    def respect_delay(self,delay,zwischen):
        if zwischen:
            return delay in (1,2)
        return True

class MahnlaufSEPAK15(Mahnlauf):
    e01_d_zw = 1
    e15_d_zw = 15
    
class MahnlaufSEPAK18(MahnlaufSEPAKx):
    e01_d_zw = 1
    e15_d_zw = 18
    
class MahnlaufSEPAD11_Mod(Mahnlauf):
    mahn = 4
    mahn_rel = 7
    
    def tag_mahnlauf(self,monat,einzugstermin,due_date):
        return self.cal.relative_workday(due_date, self.mahn)
        
    def tag_zwischeneinzug(self,month,einzugstermin,due_date):
        year = self.year
        ze = self.cal.relative_workday(due_date, self.mahn+self.mahn_rel)
        if ze.month == month and einzugstermin == 15:
            year, month = inc_month(year, month)
            fn = datetime.date(year,month,1)
            ze_new = self.cal.next_workday2(fn)
            print '%r is in same month as due date, new date: %r, chosen relative to %r' % (ze,ze_new,fn)
            return ze_new
        return ze
        
    def respect_delay(self,delay,zwischen):
        if zwischen:
            return delay in (1,2)
        return True
    
class MahnlaufSEPAD10_Mod2(Mahnlauf):
    mahn = 4
    mahn_rel = 6
    
    def respect_latest(self, zwischen):
        return False
    
    def tag_mahnlauf(self,monat,einzugstermin,due_date):
        return self.cal.relative_workday(due_date, self.mahn)
        
    def tag_zwischeneinzug(self,month,einzugstermin,due_date):
        year = self.year
        ze = self.cal.relative_workday(due_date, self.mahn+self.mahn_rel)
        if ze.month == month and einzugstermin == 15:
            year, month = inc_month(year, month)
            fn = datetime.date(year,month,1)
            ze_new = self.cal.next_workday2(fn)
            print '%r is in same month as due date, new date: %r, chosen relative to %r' % (ze,ze_new,fn)
            return ze_new
        return ze
        
    def respect_delay(self,delay,zwischen):
        if zwischen:
            return delay in (1,2,5)
        return True

c_einzug = (192,0,0)
c_einzug_last = (255,86,30)
c_sepa = (255,192,0)
c_mahn = (75,172,198)
c_zwischen_sepa = (191,191,191)
c_zwischen_last = (96,96,96)
c_zwischen = (64,64,64)
c_normal = (255,255,0)

fv_due = FlagViz(c_einzug,tc_white)
fv_sub = FlagViz(c_sepa,tc_black)
fv_mahn = FlagViz(c_mahn,tc_white)
fv_zw_due = FlagViz(c_zwischen,tc_white)
fv_zw_sub = FlagViz(c_zwischen_sepa,tc_black)

color_flags = {'due_main' : fv_due,
               'submission_main' : fv_sub,
               'mahnlauf' : fv_mahn,
               'due_zwischen' : fv_zw_due,
               'submission_zwischen' : fv_zw_sub
               }

   
class MahnlaufVisualisierung(Visualizer):        
    def custom_decoration(self,day,rect,viz):
        print day
        drop = self.vc.drop
        
        def draw_dots(dproperty):
            targets = ['origin','base_center','base_right']
            dots = ['5','2','1']
            for dot_text in dots:
                if not getattr(day,dproperty+dot_text):
                    continue
                dot = drop('dot')
                dot.text = dot_text
                dot.color = viz.txt_color
                dot.txt_color = viz.color
                dot.attach_to('anchor', rect, targets.pop())
        
        draw_dots('s')
        
        if (not day.workday) and day.target_workday:
            dot = drop('corner')
            dot.color = viz.txt_color
            dot.attach_to('anchor', rect, 'corner')
            
        if day.blocked:
            cross = drop('blocked')
            cross.color = viz.txt_color
            cross.attach_to('anchor', rect, 'right')
            
        if day.partially_blocked or day.potentially_blocked:
            cross = drop('warning')
            cross.color = viz.txt_color
            cross.attach_to('anchor', rect, 'right')
            
    #def resolve_color_conflict(self,day,flags,rect):
    #    raise MultipleFlags

def main():
    render_cal(MahnlaufVisualisierung,MahnlaufSEPAD10_Mod2,
               color_flags=color_flags,einzug=15)
    return
    #for layer in (1,2):
    #    #render_cal(MahnlaufIst,layer=layer)
    #    #render_cal(MahnlaufSEPANaiv,layer=layer)
    #    #render_cal(MahnlaufSEPAK18,layer=layer)
    #    render_cal(MahnlaufSEPAD10_Mod2)
    #render_cal(MahnlaufSEPANaiv,layer=2)
    #render_cal(MahnlaufSEPAD11_Mod,layer=1)

if __name__ == '__main__':
    main()