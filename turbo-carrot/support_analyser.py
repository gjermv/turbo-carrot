'''
Created on 8 Feb 2016

@author: gjermund.vingerhagen
'''

import matplotlib.pyplot as plt
import matplotlib
import numpy as np

import sqlite3 as lite
import pandas as pd


#DATABASENAME = 'N:\\Gjermund\\database_support\\support.db'
DATABASENAME = 'C:\\python\\database\\support.db'

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


# Plot barchart with day of time / half hour increments
times = pd.DatetimeIndex(df['starttime'])
grouped = df.groupby(times.hour)

ti = dict()
for hour,group in grouped:
    minutes = pd.DatetimeIndex(group['starttime'])
    grouped2 = group.groupby(minutes.minute)
    ti[hour] = 0
    ti[hour+0.5] = 0
    for min,group2 in grouped2:
        if min < 30:
            ti[hour] += len(group2)
        elif min <60:
            ti[hour+0.5] += len(group2)
print(ti)
a = list(ti.keys())
b = list(ti.values())
print(b)

plt.bar(a,b,width=0.5)
plt.show()

# Plot barchart with enquiries pr week day. 
times = pd.DatetimeIndex(df['starttime'])
grouped = df.groupby(times.dayofweek)

ti = dict()
for day,group in grouped:
    ti[day] = len(group)

print(ti)
a = list(ti.keys())
b = list(ti.values())
print(b)

plt.bar(a,b,width=1)
plt.show()
  

