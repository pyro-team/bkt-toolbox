# -*- coding: utf-8 -*-
'''
Created on 29.11.2012

@author: 802300
'''

import datetime

from vis_calendar import *
from colors import *

import vis_calendar
from work_calendar import Calendar, CARD_PROCESS
#from mahnlauf import c_einzug, c_mahn

class CPEinzug(BaiscCalendarModel):
    def flag_custom(self):
        cal,target,hessen = self.calendars
        flag = self._flag
        
        for month in range(1,13):
            h16 = hessen.workday16(self.year, month)
            flag(h16,'cycle')
            due = hessen.relative_workday(h16, 5)
            flag(due,'due')
            
class CPEinzugNeu(BaiscCalendarModel):
    tight = True
    #def einreichung(self,due,vorlauf,bank,target):
    #    einreichung_target = target.relative_workday(due,offset=-vorlauf)
    #    if not bank.is_workday(einreichung_target):
    #        einreichung_new = bank.previous_workday(einreichung_target)
    #        print u'Spätester TARGET-Einreichungstag (%s) für Fälligkeit am %s mit D-%d ist kein Bankarbeitstag, stattdessen Einreichung am %s' % (einreichung_target,due,vorlauf,einreichung_new)
    #        return einreichung_target,einreichung_new
    #    return einreichung_target,einreichung_target
    
    def beginn_vorlauf(self,due,vorlauf,bank,target):
        einreichung_target = target.relative_workday(due,offset=-vorlauf)
        return einreichung_target
    
    
    def flag_custom(self):
        bank,target,_ = self.calendars
        cp = Calendar(CARD_PROCESS)
        flag = self._flag
        
        for month in range(1,13):
            print '%04d-%02d' % (self.year,month)
            #h16 = hessen.workday16(self.year, month)
            
            k22 = datetime.date(self.year,month,22)
            if not cp.is_workday(k22):
                k22 = cp.previous_workday(k22)
            
#            tb_due = datetime.date(self.year,inc_month(self.year, month),1)
#            if not target.is_workday(tb_due):
#                tb_due = target.next_workday(tb_due) 
            
            flag(k22,'cycle_cp')
            
            if k22.month != month:
                raise AssertionError
            
            pb_due = bank.relative_workday(k22,3)
            pb_vorlauf = target.previous_workday(pb_due)
            
            if bank.is_workday(pb_vorlauf):
                cycle_tb = pb_vorlauf
            else:
                cycle_tb = pb_due
                #raise AssertionError(self.year,month)
                
            flag(cycle_tb,'cycle_tb')
            flag(pb_due,'pb_due')
            flag(pb_vorlauf,'pb_vorlauf')
            flag(pb_vorlauf,'pb_vorlauf1')
            
            dy,dm = inc_month(self.year,month)
            tb_due = target.next_workday2(datetime.date(dy,dm,1))
            flag(tb_due,'tb_due')
            flag(tb_due,'tb_due_main')
            
            submission = bank.next_workday(cycle_tb)
            flag(submission,'tb_submission')

            if self.tight:
                min_vorlauf = submission
            else:
                min_vorlauf = bank.next_workday(submission)
            
            for d in (1,2,5):
                vorlauf = self.beginn_vorlauf(tb_due, d, bank, target)
                if vorlauf < min_vorlauf:
                    new_due = target.earliest_due_date(min_vorlauf,d)
                    vorlauf = min_vorlauf
                    print u'Fälligkeit kann mit D-%d nich eingehalten werden' % d
                    print u'    %s statt %s' % (new_due,tb_due)
                    
                    #sub = submission
                    flag(new_due,'tb_due')
                    flag(new_due,'tb_due_later')
                    flag(new_due,'d%d' % d)
                else:
                    flag(tb_due,'d%d' % d)
                    
                flag(vorlauf,'vorlauf')
                flag(vorlauf,'v%s' % d) 


cflags = {'cycle_cp':FlagViz(c_cycle_alt, tc_black, prio=1),
          'pb_vorlauf':FlagViz(c_cycle_light, tc_white, prio=2),
          'cycle_tb':FlagViz(c_cycle, tc_white, prio=3),
          'vorlauf':FlagViz(c_sub_alt, tc_black, prio=4),
          'tb_due_main':FlagViz(c_due, tc_white),
          'tb_due_later':FlagViz(c_due_later, tc_white),
          'tb_submission':FlagViz(c_cycle_dark, tc_white, prio=1),
          }

class CPVisualizer(Visualizer):
    def custom_decoration(self, day, rect, viz):
        print day
        drop = self.drop_nofix
        #drop = self.vc.drop
        
        if (not day.workday) and day.target_workday:
            corner = drop('corner')
            corner.color = tc_black
            corner.attach_to('anchor', rect, 'corner')
        
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

        if day.vorlauf:
            draw_dots('v',self.color_flags['vorlauf'])
        if day.tb_due:
            draw_dots('d')
        if day.pb_vorlauf:
            draw_dots('pb_vorlauf',targets=['left'])
        if day.pb_due:
            letter = drop('dot')
            letter.attach_to('anchor', rect, 'left')
            letter.text = 'F'
            letter.color = viz.txt_color
            letter.txt_color = viz.color 
            
def main():
    vis_calendar.VISIO_VISIBLE = True
    render_cal(CPVisualizer, CPEinzugNeu,
               start='20.12.2013',end='19.06.2014',
               color_flags=cflags, style='wrapped',
               style_params=dict(alignment='independent',split_at=20))
    
    
if __name__ == '__main__':
    main()
