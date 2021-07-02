
import datetime
from datetime import datetime as dtime
from datetime import time
from dateutil.relativedelta import relativedelta

def timerange(PERIOD):
    # Ok PD - yesterday
    # Ok PW - previous week
    # Ok PM - previous month
    # Ok PQ - previous quarter
    # Ok PY - previous calendar year
    # Ok YTD year to date
    # Ok WTD current week to date
    # Ok TODAY today
    # mm/dd/yyyy  - a specific date 
    # ALL - since the beginning 
    # P90 - previous 90days

    # 0 is Monday 4 is Friday

    # .strftime('%Y-%m-%d')

    now = dtime.now()
    one_day = datetime.timedelta(days=1)
    one_week = datetime.timedelta(days=7)
    one_month = relativedelta(months=+1)
    ninety_days = datetime.timedelta(days=90)

    inrange = False

    if PERIOD=='PW':
      now = now - one_week
      while (now.weekday()!=6):
        now = now - one_day
      START_DATE = dtime.combine(now, time.min)
      while (now.weekday()!=5):
        now = now + one_day
      END_DATE = dtime.combine(now, time.max) 
      periodFormat = 'YYYY-MM/DD DY'
    elif PERIOD=='PM':
      current_month = now.month
      while (True):
        #print("{} - {}".format(now, now.weekday()))
        if not inrange and now.month!=current_month:
          END_DATE = dtime.combine(now, time.max) 
          inrange = True
        elif inrange and now.day==1:
          START_DATE = dtime.combine(now, time.min)
          break
        now = now - one_day
        periodFormat = 'YYYY-MM/DD DY'
    elif PERIOD=='PQ': 
      now = now - one_month
      while (now.month%3)>0:
        now = now - one_month    
      qtrend_month = now.month
      while (now.month==qtrend_month):
        now = now + one_day
      now = now - one_day
      END_DATE = dtime.combine(now, time.max)              
      now = now - one_month        
      while (now.month%3)>0:
        now = now - one_month    
      last_qtrend_month = now.month
      while (now.month==last_qtrend_month):
        now = now + one_day
      START_DATE = dtime.combine(now, time.min)
      periodFormat = 'YYYY-MM/DD WW'
    elif PERIOD=='YTD':
      END_DATE = now
      START_DATE = dtime.combine(now.replace(month=1,day=1), time.min)
      periodFormat = 'YYYY-MM'
    elif PERIOD=='WTD':
      END_DATE = now
      while(now.weekday()!=6):
        now = now - one_day
      START_DATE = dtime.combine(now, time.min)
      periodFormat = 'YYYY-MM/DD DY'
    elif PERIOD=='TODAY':
      END_DATE = now
      START_DATE = dtime.combine(now, time.min)
      periodFormat = 'HH24'
    elif PERIOD=='PD':
      now = now - one_day
      START_DATE = dtime.combine(now, time.min)
      END_DATE = dtime.combine(now, time.max)
      periodFormat = 'HH24'
    elif PERIOD=='PY':
      lastyear = now.year - 1
      START_DATE = dtime.combine(now.replace(year=lastyear, month=1, day=1), time.min)
      END_DATE = dtime.combine(now.replace(year=lastyear, month=12, day=31), time.max)
      periodFormat = 'YYYY-MM'
    elif PERIOD=='SI':
      START_DATE = dtime.combine(now.replace(year=2018, month=1, day=1), time.min)
      END_DATE = now
      periodFormat = 'YYYY-MM'
    elif PERIOD=='P90':
      END_DATE = dtime.combine(now, time.max)
      now = now - ninety_days
      START_DATE = dtime.combine(now, time.min)
      periodFormat = 'YYYY-MM WW'
    else:
      now = datetime.datetime.strptime(PERIOD, '%m/%d/%Y')
      START_DATE = dtime.combine(now, time.min)
      END_DATE = dtime.combine(now, time.max)
      periodFormat = 'HH24'
    
    #print("{} - {}".format(START_DATE, END_DATE))
    #print("Done!")
    return [START_DATE, END_DATE, periodFormat]

class Timeranger:
  def __init__(self):
        self.type = input("What type of piano? ")
        self.height = input("What height (in feet)? ")
        self.price = input("How much did it cost? ")
        self.age = input("How old is it (in years)? ")

