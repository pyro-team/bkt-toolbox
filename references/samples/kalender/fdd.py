# -*- coding: utf-8 -*-
'''
Created on 29.11.2012

@author: 802300
'''

import datetime
from colors import *
from vis_calendar import *
            
class FDDEinzugNeu(BaiscCalendarModel):
    tight = False
    
    def beginn_vorlauf(self,due,vorlauf,bank,target):
        einreichung_target = target.relative_workday(due,offset=-vorlauf)
        return einreichung_target
    
    def flag_custom(self):
        for month in range(1,13):
            self.flag_month(month)
            
    def flag_month(self,month):
        bank,target,hessen = self.calendars
        flag = self._flag
        
        print '%04d-%02d' % (self.year,month)
        h16 = hessen.workday16(self.year, month)
        flag(h16,'cycle_fdd')
        if h16.month != month:
            raise AssertionError
        
        cycle_tb = bank.next_workday(h16)
        
        if self.tight:
            submission = cycle_tb
            flag(h16,'cycle_tb')
            flag(cycle_tb,'cycle_tb_con')
        else:
            submission = bank.next_workday(cycle_tb)
            flag(cycle_tb,'cycle_tb')

        flag(submission,'submission')
        
        start_vorlauf = target.next_workday(submission)
        
        dy,dm = inc_month(self.year,month)
        due = target.next_workday2(datetime.date(dy,dm,1))
        
        flag(due,'due')
        flag(due,'due_main')
        
        for d in (1,2,5):
            vorlauf = self.beginn_vorlauf(due, d, bank, target)
            if vorlauf < start_vorlauf:
                vorlauf = start_vorlauf
                actual_due = target.earliest_due_date(start_vorlauf,d)
                print u'Fälligkeit kann mit D-%d nicht eingehalten werden' % d
                print u'    %s statt %s' % (actual_due,due)
                flag(actual_due,'due_later')
            else:
                actual_due = due

            flag(actual_due,'d%d' % d)
            flag(actual_due,'due')
            flag(vorlauf,'vorlauf')
            flag(vorlauf,'v%d' % d)
            

#c_cycle = (75,172,198)
#c_cycle_light = (139,202,218)
#c_due = (192,0,0)
#c_due_later = (255,86,30)
#c_sub = (255,192,0)


cflags = {'cycle_fdd':FlagViz(c_cycle_alt, tc_white),
          'cycle_tb':FlagViz(c_cycle, tc_white, prio=1),
          'cycle_tb_con':FlagViz(c_cycle, tc_white, prio=2),
          'submission_pre':FlagViz(c_cycle_dark, tc_white,prio=3),
          'submission':FlagViz(c_cycle_dark, tc_white,prio=4),
          'vorlauf':FlagViz(c_sub_alt, tc_black,prio=5),
          'due_main':FlagViz(c_due, tc_white),
          'due_later':FlagViz(c_due_later, tc_white),
          }

class FDDVisualizer(Visualizer):
    def custom_decoration(self, day, rect, viz):
        print day
        drop = self.vc.drop
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

        if day.due:
            draw_dots('d')
        if day.vorlauf:
            draw_dots('v',self.color_flags['vorlauf'])
        if (not day.workday) and day.target_workday:
            corner = drop('corner')
            corner.attach_to('anchor', rect, 'corner')
            corner.color = tc_black
                
            
def main():
    #import vis_calendar
    #vis_calendar.VISIO_VISIBLE = True
    render_cal(FDDVisualizer, FDDEinzugNeu,
               #2013, 2015, m0=12, m1=1,
               start='01.12.2013',end='31.07.2014',
               color_flags=cflags, style='wrapped',
               style_params=dict(alignment='independent'))

if __name__ == '__main__':
    main()
