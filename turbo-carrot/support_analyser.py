'''
Created on 8 Feb 2016

@author: gjermund.vingerhagen
'''

import matplotlib.pyplot as plt
import matplotlib
import numpy as np

import sqlite3 as lite
import pandas as pd

DATABASENAME = 'N:\\Gjermund\\database_support\\support.db'
#DATABASENAME = 'C:\\python\\database\\support.db'

con = lite.connect(DATABASENAME)

cur = con.cursor()    
cur.execute("SELECT * FROM Support")
rows = cur.fetchall()        
con.close()

li =[]
for item in rows:
    li.append(item)
    
df = pd.DataFrame(li,columns=['id','starttime','duration','regtime','name','company','phone','product','serial','problem','solution','solved','forwarded','followup','user'])
df['starttime'] = pd.to_datetime(df['starttime'],format="%Y-%m-%d %H:%M:%S.%f")
df = df[df['starttime']>pd.to_datetime('2016-02-01 00:00:00',format="%Y-%m-%d %H:%M:%S")]

# Plot barchart with day of time / half hour increments
times = pd.DatetimeIndex(df['starttime'])
grouped = df.groupby(times.hour)

ti = dict()
to = dict()
for hour,group in grouped:
    minutes = pd.DatetimeIndex(group['starttime'])
    grouped2 = group.groupby(minutes.minute)
    ti[hour] = 0
    ti[hour+0.5] = 0
    to[hour] = 0
    to[hour+0.5] = 0
    for min,group2 in grouped2:
        if min < 30:
            ti[hour] += len(group2[group2['user']=='GV'])
            to[hour] += len(group2[group2['user']=='CO'])
        elif min <60:
            ti[hour+0.5] += len(group2[group2['user']=='GV'])
            to[hour+0.5] += len(group2[group2['user']=='CO'])
print(ti)
a = list(ti.keys())
b = list(ti.values())
c = list(to.values())
print(b)

plt.bar(a,b,width=0.5,color='r')
plt.bar(a,c,width=0.5,bottom=b)
plt.show()

# Plot barchart with enquiries pr week day. 
times = pd.DatetimeIndex(df['starttime'])
grouped = df.groupby(times.dayofweek)

ti = dict()
to = dict()
for day,group in grouped:
    ti[day] = len(group[group['user']=='GV'])
    to[day] = len(group[group['user']=='CO'])


a = list(ti.keys())
b = list(ti.values())
c = list(to.values())


plt.bar(a,b,width=1,color='r')
plt.bar(a,c,width=1,bottom=b)
plt.show()
  
# Plot barchart with enquiries pr week day. 
times = pd.DatetimeIndex(df['starttime'])
grouped = df.groupby([times.year,times.dayofyear])

ti = list()
tix = list()
to = list()
tr = list()
for yearweek,group in grouped:
    ti.append((yearweek[0]-2016)*365+yearweek[1])
    tix.append((yearweek[0]-2016)*365+yearweek[1]+0.4)
    to.append(len(group[group['user']=='GV']))
    tr.append(len(group[group['user']=='CO']))


plt.bar(ti,to,width=0.9,color='r')
plt.bar(ti,tr,width=0.9,bottom=to)
plt.show()