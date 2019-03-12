# -*- coding: utf-8 -*-
'''
Created on 22.10.2012

@author: 802300
'''

import datetime

_easter_sunday = ["24.04.2011",
                  "08.04.2012",
                  "31.03.2013",
                  "20.04.2014",
                  "05.04.2015",
                  "27.03.2016",
                  "16.04.2017",
                  "01.04.2018",
                  "21.04.2019",
                  "12.04.2020",
                  "04.04.2021",
                  "17.04.2022",
                  "09.04.2023",
                  "31.03.2024",
                  "20.04.2025",
                  "05.04.2026",
                  "28.03.2027",
                  "16.04.2028",
                  "01.04.2029",
                  "21.04.2030",
                  "13.04.2031"]

def iter_easter():
    for s in _easter_sunday:
        d,m,y = (int(c) for c in s.split('.'))
        yield y,m,d

class FixedHoliday(object):
    def __init__(self,month,day):
        self.month = month
        self.day = day
        
    def get_day(self,year):
        return datetime.date(year,self.month,self.day)
    
NEW_YEAR = FixedHoliday(1, 1)
MAY_DAY = FixedHoliday(5, 1)
GERMAN_REUNIFICATION = FixedHoliday(10,3)
CHRISTMAS_EVE = FixedHoliday(12,24)
CHRISTMAS_1 = FixedHoliday(12,25)
CHRISTMAS_2 = FixedHoliday(12,26)
NEW_YEARS_EVE = FixedHoliday(12,31)

class EasterRelativeHoliday(object):
    easter_sunday = dict((year,datetime.date(year,month,day)) for year,month,day in iter_easter())
    #print easter_sunday
    
    def __init__(self,delta):
        self.delta = datetime.timedelta(days=delta)
    
    def get_day(self,year):
        return self.easter_sunday[year] + self.delta

GOOD_FRIDAY = EasterRelativeHoliday(-2)
EASTER_MONDAY = EasterRelativeHoliday(1)
ASCENSION_THURSDAY = EasterRelativeHoliday(39)
PENTECOAST_MONDAY = EasterRelativeHoliday(50)
CORPUS_CHRISTI = EasterRelativeHoliday(60)

TARGET = [NEW_YEAR,
          GOOD_FRIDAY,
          EASTER_MONDAY,
          MAY_DAY,
          CHRISTMAS_1,
          CHRISTMAS_2,
          ]

HESSEN = TARGET + [ASCENSION_THURSDAY,
                   PENTECOAST_MONDAY,
                   CORPUS_CHRISTI,
                   GERMAN_REUNIFICATION
                   ]

BUND = TARGET + [ASCENSION_THURSDAY,
                   PENTECOAST_MONDAY,
                   GERMAN_REUNIFICATION
                   ]

BANK = BUND + [CHRISTMAS_EVE, NEW_YEARS_EVE]
CARD_PROCESS = BUND + [NEW_YEARS_EVE]

class DateRange(object):
    def __init__(self,first,last):
        self.first = first
        self.last = last
        self.delta = datetime.timedelta(days=1)
    
    def __iter__(self):
        current = self.first
        while current <= self.last:
            yield current
            current = current+self.delta
            
class Calendar(object):
    def __init__(self,holidays):
        self.holidays = holidays
        self.holiday_by_year = {}
        
    def cache_holidays(self,year):
        hd = set(h.get_day(year) for h in self.holidays)
        self.holiday_by_year[year] = hd
        
    def is_holiday(self,date):
        if not date.year in self.holiday_by_year:
            self.cache_holidays(date.year)
        return date in self.holiday_by_year[date.year]
    
    def is_workday(self,date):
        return date.weekday() <= 4 and not self.is_holiday(date)
    
    def next_workday(self,date):
        delta = datetime.timedelta(days=1)
        current = date
        while True:
            current = current + delta
            if self.is_workday(current):
                return current
            
    def previous_workday(self,date):
        delta = datetime.timedelta(days=-1)
        current = date
        while True:
            current = current + delta
            if self.is_workday(current):
                return current
    
    def _days(self,year,month):
        delta = datetime.timedelta(days=1)
        days = [datetime.date(year,month,1)]
        while True:
            next_day = days[-1] + delta
            if next_day.month == month:
                days.append(next_day)
            else:
                break
        return days
    
    def _workdays(self,year,month):
        return [d for d in self._days(year,month) if self.is_workday(d)]
    
    def ultimo(self,year,month,offset=0):
        work_days = self._workdays(year, month)
        return work_days[-(1+offset)]
            
    def first(self,year,month,offset=0):
        work_days = self._workdays(year, month)
        return work_days[offset]
    
    def next_workday2(self,date):
        if self.is_workday(date):
            return date
        else:
            return self.next_workday(date)
            
    def relative_workday(self,date,offset=1):
        if offset > 0:
            method = self.next_workday
        else:
            offset *= -1
            method = self.previous_workday
        
        current = date
        while offset > 0:
            current = method(current)
            offset -= 1
        return current
        
    def iter_workdays(self,date_range):
        for date in date_range:
            if self.is_workday(date):
                yield date
                
class TargetCalendar(Calendar):
    def __init__(self):
        Calendar.__init__(self,TARGET)

    def actual_due(self,due_date):
        if self.is_workday(due_date):
            actual_due = due_date
        else:
            actual_due = self.next_workday(due_date)
        return actual_due

    def latest_submission(self,due_date,delay=1):
        if not self.is_workday(due_date):
            raise ValueError
        return self.relative_workday(self.actual_due(due_date), -delay)
    
    def earliest_due_date(self,submission_date,delay=1):
        if not self.is_workday(submission_date):
            raise ValueError
        return self.relative_workday(submission_date, delay)
           
class HessenCalendar(Calendar):
    def __init__(self):
        Calendar.__init__(self,HESSEN)
    
    def workday16(self,year,month):
        current = datetime.date(year,month,1)
        delta = datetime.timedelta(days=1)
        workdays = 0
        while True:
            if self.is_workday(current):
                workdays += 1
                if workdays == 16:
                    if current.month != month:
                        raise AssertionError
                    return current
            current = current + delta

def count_following_target_days_till_end_of_month(date):
    tc = TargetCalendar()
    current = date
    count = 0
    while current.month == date.month:
        current = tc.next_workday(current)
        count += 1
    return count-1
    

def main():
    ''' Test some foo '''
    cal = Calendar(BANK)
    print cal.ultimo(2012,03)
    return
    
    #em = EasterRelativeHoliday(1)
    #print em.get_day(2012)
    print count_following_target_days_till_end_of_month(datetime.date(2012,10,26))
    print count_following_target_days_till_end_of_month(datetime.date(2012,10,27))
    print count_following_target_days_till_end_of_month(datetime.date(2012,10,28))
    print count_following_target_days_till_end_of_month(datetime.date(2012,10,29))
    print count_following_target_days_till_end_of_month(datetime.date(2012,10,30))
    print count_following_target_days_till_end_of_month(datetime.date(2012,10,31))
    
    print '----------'

    tc = TargetCalendar()
    print tc.latest_submission(datetime.date(2012,04,6), 2)
    print tc.latest_submission(datetime.date(2012,04,7), 2)
    print tc.latest_submission(datetime.date(2012,04,8), 2)
    print tc.latest_submission(datetime.date(2012,04,9), 2)
    print tc.latest_submission(datetime.date(2012,04,10), 2)

    print '----------'
    for year in range(2012,2023):
        for month in range(1,13):
            for day in (1,15):
                due = datetime.date(year,month,day)
                actual_due = tc.actual_due(due)
                sub = tc.latest_submission(due, 2)
                last_change = tc.previous_workday(sub)
                delta = (due-last_change).days
                new_due = tc.actual_due(tc.earliest_due_date(sub, 5))
                due_delta = (new_due-actual_due).days
                print 'due=%s, actual_due=%s, submission=%s, last_change=%s (%d CD before due date), new_due=%s, actual_due_delta=%d CD, new_due_day_of_month=%d' % (due,actual_due,sub,last_change,delta,new_due,due_delta,new_due.day)
    
    print '----------'
    hc = HessenCalendar()
    for year in range(2012,2023):
        for month in range(1,13):
            date = hc.workday16(year, month)
            rem_tt = count_following_target_days_till_end_of_month(date)
            print (date,rem_tt)

if __name__ == '__main__':
    main()

