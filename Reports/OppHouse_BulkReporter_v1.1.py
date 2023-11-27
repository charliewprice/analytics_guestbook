
# coding: utf-8

# In[1]:


import ReportMailer as rm


# # Install a conda package in the current Jupyter kernel
# import sys
# !{sys.executable} -m pip install xlsxwriter

# # Reporting Period #
# * PD - yesterday
# * PW - previous week
# * PM - previous month
# * PQ - previous quarter
# * PY - previous calendar year
# * YTD year to date
# * WTD current week to date
# * TODAY today
# * P90 - previous 90days

# In[2]:


acronyms = [ "PD", "PW", "PM", "PQ", "PY", "YTD", "WTD", "TODAY", "P90"]
friendly = [ "Prior Day", "Prior Week", "Prior Month", "Prior Quarter", "Prior Year", "Year to Date", "Week to Date", "Today", "Previous 90 days"]

def toPeriodFriendly(INTERVAL):
    for n in range(0, len(acronyms)):
      if INTERVAL==acronyms[n]:
         return friendly[n]
    return "Unknown Period"


# class StopExecution(Exception):
#     def _render_traceback_(self):
#         pass
# 
# raise StopExecution

# def exit(): raise StopExecution

# In[3]:


import os
import pandas as pd

path = "/home/gbadmin/jupy-notebooks/Reports/inqueue"
dir_list = os.listdir(path)

if len(dir_list)>0:
  filename = "inqueue/{}".format(dir_list[0])
  print(filename)

  rep = pd.read_json(filename, lines=True)

  for _ , row in rep.iterrows():    
    _INTERVAL = row['interval']
    _TYPE = row['type']
    if (_TYPE=='ALL'):
       _APPUSER  = row['appuser']
       print(_INTERVAL)
       print(_TYPE)
       # delete the file
       os.remove(filename)
    else:
      print("Non-Rollup type")
      exit()
       
else:
  print("NOTHING TO DO")
  exit()



# #_INTERVAL = "PY"
# _INTERVAL = 'YYYYMM202307'

# In[4]:


from pandas import ExcelWriter
from pandas import ExcelFile
import xlsxwriter
import datetime

def createSpreadsheetAndMailIt(collections, reportname, recipients, subject, body):
  now = datetime.datetime.now().strftime('%Y-%m-%d_%H%M%S')
  filename = reportname + "-" + now + ".xlsx"
  #writer = ExcelWriter(filename)
  writer = ExcelWriter(filename, engine='xlsxwriter')
  workbook = writer.book  
  print("Writing dataframe to Excel file {0}".format(filename))
  for collection in collections:
    print("Writing {} to spreadsheet".format(collection["name"]))
    collection["dataframe"].to_excel(writer, sheet_name=collection["name"], index=False)
    worksheet = writer.sheets[collection["name"]]
    if 'colwidths' in collection:
      colwidths = collection['colwidths']
      print("colwidths={}".format(colwidths))
      for n in range(0, len(colwidths)):
        col = str(chr(65 + n))
        worksheet.set_column('{}:{}'.format(col,col), colwidths[n])
    else:
      worksheet.set_column('A:A', 30)
  writer.save()
  for recipient in recipients: 
    rm.mailer(recipient, subject, body, filename)
  print("Excel/Email Done!")


# In[5]:


import psycopg2 as pg
import pandas.io.sql as psql
import pandas as pd
import GuestbookDbConnect as gdb

conn = gdb.guestbookDbConnect()


# In[6]:


import ReportTimerange as rtr

if _INTERVAL.startswith("YYYYMM"):
  yyyymm = _INTERVAL[6:12]
  print(yyyymm)
  yyyy = int(yyyymm[0:4])
  mm = int(yyyymm[4:6])
  print("{} {}".format(yyyy,mm))
  trange = rtr.mmyyyy(mm,yyyy)
else:
  trange = rtr.timerange(_INTERVAL)
print(trange)

collections = []
summaries = []

collection = {}
collection["name"] = "Outline"

df = pd.DataFrame({"Report Tab":[               "Summary",               "Distinct Clients",               "Services",               "Housing",               "Incarcerations",               "Gender",               "Ethnicity",               "Veterans",               "New Clients"
                                 ], \
                   "Description":[ \
              "summary totals for the report period", \
              "a count of individuals served by sub-period within the period.", \
              "breakdown by service type and sub-period within the period.", \
              "breakdown by housing conditions and sub-period within the period.", \
              "clients reporting incarceration in previous 90 day period.", \
              "breakdown by gender and sub-period within the period.", \
              "breakdown by ethnicity and sub-period within the period.", \
              "breakdown by veteran status and sub-period within the period.", \
              "a list of clients with first visit occurring in the report period."
                                 ]})

collection["dataframe"] = df 
collection["colwidths"] = [30, 100]
collections.append(collection)

df.head(1000)


# ## Total distinct clients in report period ##

# In[7]:


summary = {}
summary["name"] = "Total distinct clients in report period"

query = "SELECT min(timestamp)AS start, max(timestamp) AS end,          COUNT(DISTINCT person_id) AS clients FROM guestbook_personsnapshot          WHERE timestamp BETWEEN '{}' AND '{}'".format(trange[0], trange[1])

print(query)

data = pd.read_sql(query, conn)

summary["count"] = data.iloc[0,2]
summaries.append(summary)
print(summaries)

data.head(100)


# # Distinct clients by sub-interval within the report period #
# * Statistics focused on the breadth of individuals that are served during the time period.

# In[8]:


collection = {}
collection["name"] = "Distinct Clients"

query = "SELECT to_char(timestamp,'{}') AS period,           COUNT(DISTINCT person.idperson) AS clients          FROM guestbook_personsnapshot snapshot          JOIN guestbook_person person ON person.idperson=snapshot.person_id          WHERE timestamp BETWEEN '{}' AND '{}'          GROUP BY period".format(trange[2], trange[0], trange[1])
print(query)

data = pd.read_sql(query, conn)
#data.head(1000)

persons = data.pivot_table('clients', index=['period']).fillna(0).astype(int).reset_index('period')

collection["dataframe"] = persons 
collection["colwidths"] = [30, 20]
collections.append(collection)
persons.head(1000)


# ## Total Services delivered in report period ##

# In[9]:


query = "SELECT          left(service.name, strpos(service.name, '/') - 1) AS servicename, COUNT(*) AS total FROM guestbook_personsnapshot snapshot         JOIN guestbook_personservicerequest servicerequest ON servicerequest.connection_id=snapshot.idsnapshot          JOIN guestbook_service service ON service.idservice=servicerequest.service_id          WHERE timestamp BETWEEN '{}' AND '{}'  AND service.points<=0          GROUP BY servicename          ORDER BY total desc".format(trange[0], trange[1])

print(query)

data = pd.read_sql(query, conn)


for n in range(0, data.shape[0]):  
  summary = {}
  summary["name"]  = data.iloc[n,0] + " services delivered."
  summary["count"] = data.iloc[n,1]
  summaries.append(summary)
print(summaries)
data.head(1000)


# # Services delivered by sub-interval within the report period #

# In[10]:


collection = {}
collection["name"] = "Services"

#query = "SELECT to_char(timestamp,'{}') AS period, \
#         service.name AS servicename, COUNT(*) AS total FROM guestbook_personsnapshot snapshot\
#         JOIN guestbook_personservicerequest servicerequest ON servicerequest.connection_id=snapshot.idsnapshot \
#         JOIN guestbook_service service ON service.idservice=servicerequest.service_id \
#         WHERE timestamp BETWEEN '{}' AND '{}' \
#         GROUP BY period, servicename".format(trange[2], trange[0], trange[1])

query = "SELECT to_char(timestamp,'{}') AS period,          left(service.name, strpos(service.name, '/') - 1) AS servicename, COUNT(*) AS total FROM guestbook_personsnapshot snapshot         JOIN guestbook_personservicerequest servicerequest ON servicerequest.connection_id=snapshot.idsnapshot          JOIN guestbook_service service ON service.idservice=servicerequest.service_id          WHERE timestamp BETWEEN '{}' AND '{}' AND service.points<=0          GROUP BY period, servicename".format(trange[2], trange[0], trange[1])



print(query)

data = pd.read_sql(query, conn)

services = data.pivot_table('total', index=['period'], columns='servicename').fillna(0).astype(int).reset_index('period')
collection["dataframe"] = services 
collection["colwidths"] = [30, 20, 20, 20, 20, 20, 20, 20]
collections.append(collection)

services.head(1000)


# # Housing totals within the report period#

# In[11]:


query = "SELECT           left(housing.name, strpos(housing.name, '(') - 1)  AS hresponse, COUNT(*) AS total FROM guestbook_personsnapshot snapshot         JOIN guestbook_personsurvey survey ON survey.connection_id=snapshot.idsnapshot          JOIN guestbook_housingresponse housing ON housing.idhousing=survey.object_id          WHERE timestamp BETWEEN '{}' AND '{}' AND prompt_id=8          GROUP BY hresponse".format(trange[0], trange[1])

print(query)

data = pd.read_sql(query, conn)

for n in range(0, data.shape[0]):  
  summary = {}
  summary["name"]  = " Days spent in " + data.iloc[n,0] 
  summary["count"] = data.iloc[n,1]
  summaries.append(summary)
print(summaries)

data.head(1000)


# # Housing by City within the report period#

# In[12]:


query = "SELECT           city.name AS cresponse, COUNT(*) AS total FROM guestbook_personsnapshot snapshot         JOIN guestbook_personsurvey survey ON survey.connection_id=snapshot.idsnapshot          JOIN guestbook_cityresponse city ON city.idcity=survey.object_id          WHERE timestamp BETWEEN '{}' AND '{}' AND prompt_id=53          GROUP BY cresponse ORDER BY total desc".format(trange[0], trange[1])

print(query)

data = pd.read_sql(query, conn)

for n in range(0, data.shape[0]):  
  summary = {}
  summary["name"]  = " Nights spent in " + data.iloc[n,0] 
  summary["count"] = data.iloc[n,1]
  summaries.append(summary)
print(summaries)

data.head(1000)


# # Housing by sub-interval within the report period#
# * Answers to the prompt 'Where did you spend last night?'

# In[13]:


collection = {}
collection["name"] = "Housing"

#query = "SELECT to_char(timestamp,'{}') AS period, \
#         housing.name AS hresponse, COUNT(*) AS total FROM guestbook_personsnapshot snapshot\
#         JOIN guestbook_personsurvey survey ON survey.connection_id=snapshot.idsnapshot \
#         JOIN guestbook_housingresponse housing ON housing.idhousing=survey.object_id \
#         WHERE timestamp BETWEEN '{}' AND '{}' AND prompt_id=8 \
#         GROUP BY period, hresponse".format(trange[2], trange[0], trange[1])

query = "SELECT to_char(timestamp,'{}') AS period,          left(housing.name, strpos(housing.name, '(') - 1) AS hresponse, COUNT(*) AS total FROM guestbook_personsnapshot snapshot         JOIN guestbook_personsurvey survey ON survey.connection_id=snapshot.idsnapshot          JOIN guestbook_housingresponse housing ON housing.idhousing=survey.object_id          WHERE timestamp BETWEEN '{}' AND '{}' AND prompt_id=8          GROUP BY period, hresponse".format(trange[2], trange[0], trange[1])

print(query)

data = pd.read_sql(query, conn)

housing = data.pivot_table('total', index=['period'], columns='hresponse').fillna(0).astype(int).reset_index('period')
#housing = housing.fillna(0)
collection["dataframe"] = housing  
collection["colwidths"] = [30, 30, 30, 30, 30, 30]
collections.append(collection)
housing.head(1000)


# # Housing by City and sub-interval within the report period#

# In[14]:


collection = {}
collection["name"] = "Housing by City"

#query = "SELECT to_char(timestamp,'{}') AS period, \
#         housing.name AS hresponse, COUNT(*) AS total FROM guestbook_personsnapshot snapshot\
#         JOIN guestbook_personsurvey survey ON survey.connection_id=snapshot.idsnapshot \
#         JOIN guestbook_housingresponse housing ON housing.idhousing=survey.object_id \
#         WHERE timestamp BETWEEN '{}' AND '{}' AND prompt_id=8 \
#         GROUP BY period, hresponse".format(trange[2], trange[0], trange[1])

query = "SELECT to_char(timestamp,'{}') AS period,          city.name AS cresponse, COUNT(*) AS total FROM guestbook_personsnapshot snapshot         JOIN guestbook_personsurvey survey ON survey.connection_id=snapshot.idsnapshot          JOIN guestbook_cityresponse city ON city.idcity=survey.object_id          WHERE timestamp BETWEEN '{}' AND '{}' AND prompt_id=53          GROUP BY period, cresponse".format(trange[2], trange[0], trange[1])

print(query)

data = pd.read_sql(query, conn)

city = data.pivot_table('total', index=['period'], columns='cresponse').fillna(0).astype(int).reset_index('period')
#housing = housing.fillna(0)
collection["dataframe"] = city  
collection["colwidths"] = [30, 30, 30, 30]
collections.append(collection)
city.head(1000)


# # Clients (by name) Reporting Unsheltered Housing #
# * more than once in the report period

# In[15]:


collection = {}
collection["name"] = "Unsheltered_by_Name"

unsheltered = 5
threshold = 1

#query = "SELECT to_char(timestamp,'{}') AS period, \
#         housing.name AS hresponse, COUNT(*) AS total FROM guestbook_personsnapshot snapshot\
#         JOIN guestbook_personsurvey survey ON survey.connection_id=snapshot.idsnapshot \
#         JOIN guestbook_housingresponse housing ON housing.idhousing=survey.object_id \
#         WHERE timestamp BETWEEN '{}' AND '{}' AND prompt_id=8 \
#         GROUP BY period, hresponse".format(trange[2], trange[0], trange[1])

query = "SELECT person.firstname, person.lastname, person.aliasname,          COUNT(*) AS unsheltered_nights FROM guestbook_personsnapshot snapshot         JOIN guestbook_person person ON person.idperson=snapshot.person_id          JOIN guestbook_personsurvey survey ON survey.connection_id=snapshot.idsnapshot          JOIN guestbook_housingresponse housing ON housing.idhousing=survey.object_id          WHERE timestamp BETWEEN '{}' AND '{}' AND prompt_id=8 AND object_id={}          GROUP BY person.idperson ORDER BY COUNT(*) desc".format(trange[0], trange[1], unsheltered)

print(query)

data = pd.read_sql(query, conn)
data.head(100)

#housing = data.pivot_table('total', index=['period'], columns='hresponse').fillna(0).astype(int).reset_index('period')
collection["dataframe"] = data 
collection["colwidths"] = [30, 30, 30, 30, 30, 30]
collections.append(collection)
data.head(1000)


# # Clients Reporting Incarceration in last 90 days #

# In[16]:


collection = {}
collection["name"] = "Incarcerations"


#query = "SELECT to_char(timestamp,'{}') AS period, \
#         housing.name AS hresponse, COUNT(*) AS total FROM guestbook_personsnapshot snapshot\
#         JOIN guestbook_personsurvey survey ON survey.connection_id=snapshot.idsnapshot \
#         JOIN guestbook_housingresponse housing ON housing.idhousing=survey.object_id \
#         WHERE timestamp BETWEEN '{}' AND '{}' AND prompt_id=8 \
#         GROUP BY period, hresponse".format(trange[2], trange[0], trange[1])

query = "SELECT person.idperson, person.firstname, person.lastname         FROM guestbook_personsnapshot snapshot         JOIN guestbook_person person ON person.idperson=snapshot.person_id          JOIN guestbook_personsurvey survey ON survey.connection_id=snapshot.idsnapshot          WHERE survey.object_id=2 AND timestamp BETWEEN '{}' AND '{}' AND prompt_id=4          GROUP BY person.idperson, person.firstname, person.lastname          ORDER BY person.idperson".format(trange[0], trange[1])

print(query)

data = pd.read_sql(query, conn)

#housing = data.pivot_table('total', index=['period'], columns='hresponse').fillna(0).astype(int).reset_index('period')
collection["dataframe"] = data 
collection["colwidths"] = [30, 30, 30]
collections.append(collection)
data.head(1000)


# # Gender totals in the report period #

# In[17]:


query = "SELECT           gender.name AS response, COUNT(DISTINCT person.idperson) AS total FROM guestbook_personsnapshot snapshot         JOIN guestbook_person person ON person.idperson=snapshot.person_id          JOIN guestbook_genderresponse gender ON gender.idgender=person.gender_id          WHERE timestamp BETWEEN '{}' AND '{}'          GROUP BY response".format(trange[0], trange[1])

print(query)
#summaries = []

data = pd.read_sql(query, conn)

for n in range(0, data.shape[0]):  
  summary = {}
  summary["name"]  = "Gender -" + data.iloc[n,0] 
  summary["count"] = data.iloc[n,1]
  summaries.append(summary)
print(summaries)

data.head(100)


# # Gender by sub-interval within the report period #

# In[18]:


collection = {}
collection["name"] = "Gender"

query = "SELECT to_char(timestamp,'{}') AS period,          gender.name AS response, COUNT(DISTINCT person.idperson) AS total FROM guestbook_personsnapshot snapshot         JOIN guestbook_person person ON person.idperson=snapshot.person_id          JOIN guestbook_genderresponse gender ON gender.idgender=person.gender_id          WHERE timestamp BETWEEN '{}' AND '{}'          GROUP BY period, response".format(trange[2], trange[0], trange[1])

print(query)

data = pd.read_sql(query, conn)

gender = data.pivot_table('total', index=['period'], columns='response').fillna(0).astype(int).reset_index('period')
collection["dataframe"] = gender
collection["colwidths"] = [30, 20, 20, 20]
#collection["colwidths"] = []
collections.append(collection)
gender.head(1000)


# # Ethnicity totals within the report period #

# In[19]:


query = "SELECT           ethnicity.name AS response, COUNT(DISTINCT person.idperson) AS total FROM guestbook_personsnapshot snapshot         JOIN guestbook_person person ON person.idperson=snapshot.person_id          JOIN guestbook_ethnicityresponse ethnicity ON ethnicity.idethnicity=person.ethnicity_id          WHERE timestamp BETWEEN '{}' AND '{}'          GROUP BY response".format(trange[0], trange[1])

print(query)
#summaries = []

data = pd.read_sql(query, conn)
for n in range(0, data.shape[0]):  
  summary = {}
  summary["name"]  = "Ethnicity -" + data.iloc[n,0] 
  summary["count"] = data.iloc[n,1]
  summaries.append(summary)
print(summaries)
data.head(1000)


# # Ethnicity by sub-interval within the report period #

# In[20]:


collection = {}
collection["name"] = "Ethnicity"

query = "SELECT to_char(timestamp,'{}') AS period,          ethnicity.name AS response, COUNT(DISTINCT person.idperson) AS total FROM guestbook_personsnapshot snapshot         JOIN guestbook_person person ON person.idperson=snapshot.person_id          JOIN guestbook_ethnicityresponse ethnicity ON ethnicity.idethnicity=person.ethnicity_id          WHERE timestamp BETWEEN '{}' AND '{}'          GROUP BY period, response".format(trange[2], trange[0], trange[1])

print(query)

data = pd.read_sql(query, conn)

ethnicity = data.pivot_table('total', index=['period'], columns='response').fillna(0).astype(int).reset_index('period')
collection["dataframe"] = ethnicity 
collection["colwidths"] = [30, 20, 20, 20, 20, 20, 20, 20]
collections.append(collection)
ethnicity.head(1000)


# # Veteran status totals within the report period#

# In[21]:


query = "SELECT           yesno.name AS response, COUNT(DISTINCT person.idperson) AS total FROM guestbook_personsnapshot snapshot         JOIN guestbook_person person ON person.idperson=snapshot.person_id          JOIN guestbook_yesnoresponse yesno ON yesno.idyesno=person.veteran_id          WHERE timestamp BETWEEN '{}' AND '{}'          GROUP BY response          ORDER BY response desc".format(trange[0], trange[1])

print(query)
#summaries = []

data = pd.read_sql(query, conn)

for n in range(0, data.shape[0]):  
  summary = {}
  if data.iloc[n,0]=="Yes":
    summary["name"] = "Veterans"
  elif data.iloc[n,0]=="No":
    summary["name"] = "Non-veterans"
  else:
    summary["name"] = "Veteran status Unknown"
  summary["count"] = data.iloc[n,1]
  summaries.append(summary)
print(summaries)

data.head(1000)


# # Veteran Status by sub-interval within the report period#

# In[22]:


collection = {}
collection["name"] = "Veterans"

query = "SELECT to_char(timestamp,'{}') AS period,          yesno.name AS response, COUNT(DISTINCT person.idperson) AS total FROM guestbook_personsnapshot snapshot         JOIN guestbook_person person ON person.idperson=snapshot.person_id          JOIN guestbook_yesnoresponse yesno ON yesno.idyesno=person.veteran_id          WHERE timestamp BETWEEN '{}' AND '{}'          GROUP BY period, response".format(trange[2], trange[0], trange[1])

print(query)

data = pd.read_sql(query, conn)
#data.head(1000)



veterans  = data.pivot_table('total', index=['period'], columns='response').fillna(0).astype(int).reset_index('period')
veterans.columns = ['period', 'non-veteran', 'unknown','veteran']
collection["dataframe"] = veterans 
collection["colwidths"] = [30, 30, 30, 30]
collections.append(collection)
veterans.head(1000)


# # New Clients within the report period#
# * Clients first appearing in the report interval

# In[23]:


collection = {}
collection["name"] = "New Clients"

query = "SELECT firstname, lastname, aliasname, timelinestarttime::date AS startdate FROM guestbook_person          WHERE timelinestarttime BETWEEN '{}' AND '{}'          ORDER BY timelinestarttime asc".format(trange[0], trange[1])

print(query)

data = pd.read_sql(query, conn)
collection["dataframe"] = data
collection["colwidths"] = [30, 30, 30, 30]
collections.append(collection)

data.head(5000)


# # Summary Counts #

# In[24]:


collection = {}
summ = pd.DataFrame.from_dict(summaries)
collection["name"] = "Summary"
collection["dataframe"] = summ
collection["colwidths"] = [30, 60]
collections.insert(1, collection)
#collections.append(collection)
summ.head(1000)


# In[25]:


_REPORTNAME = _INTERVAL
_SUBJECT = "Opportunity House - Reports for {}".format(toPeriodFriendly(_INTERVAL))
_BODY    = "Spreadsheet (attached) with reports for period {} through {}.".format(trange[0], trange[1])
#_EMAIL_RECIPIENT = ['cprice9739@carolina.rr.com', 'jahood1@yahoo.com']
#_EMAIL_RECIPIENT = ['cprice9739@carolina.rr.com', 'pastor@opphouse.net']
#_EMAIL_RECIPIENT = ['cprice9739@carolina.rr.com']

query = "SELECT id, email FROM guestbook_appuser WHERE id={}".format(_APPUSER)

#print(query)
data = pd.read_sql(query, conn)
_EMAIL_RECIPIENT = {data.iloc[0,1]}
#_EMAIL_RECIPIENT = ['holobox@gmail.com']
print(">>{}<<".format(_EMAIL_RECIPIENT))

createSpreadsheetAndMailIt(collections, _REPORTNAME, _EMAIL_RECIPIENT, _SUBJECT, _BODY)

