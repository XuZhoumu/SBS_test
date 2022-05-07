# -*- coding: utf-8 -*-
"""
Created on Wed 2018/8/1

@author: Yunfu.ou
"""

import matplotlib.pyplot as plt
import numpy as np
from scipy.integrate import simps
import csv
import pandas as pd
from pandas import Series,DataFrame
import openpyxl


plt.close()

## Define variables
n = 13
sizedata = pd.read_excel('尺寸数据123.xlsx',usecols=[1,2],skiprows=(0))
sizedata.columns = Series(['thickness','width'])
width = sizedata['width']
thickness = sizedata['thickness']

Fsbs_list=[]
plt.figure()
plt.xlabel('Displacement [mm]')
plt.ylabel('Force [N]')

#get sheet_name
#alldata = pd.read_excel('原始数据.xlsx')
#sheetname = pd.series(pd.sheet_name(alldate))
for var in range(1,n+1):
     b = width[var-1]
     h = thickness[var-1]
     df = pd.read_excel('新-原始数据.xlsx', sheet_name=(var-1),usecols=[2, 14], skiprows=1)
     df.columns = Series(['force','disp'])

     # Split data in displacement and force and time
  
     disptot = df['disp']
     forcetot = df['force']


     ## Cleanup raw data (test data made at room 615 don't need this step.)
     forcetotcor = forcetot
     disptotcor = disptot
     
     # Remove negative force
     for index in range(len(forcetot)):
         if forcetot[index] < 0:
             forcetotcor[index] = 0
             
     Pm = max(forcetotcor)
     Fsbs = 0.75 * Pm / (b * h)
     Fsbs_list.append(Fsbs)

     # Plot raw data
     plt.plot(disptotcor, forcetotcor/b)

     print("Fsbs {} =".format(var), Fsbs)

     # Plain text
     filename = '{}.txt'.format(var)
     f = open(filename, 'w')
     f.write('{0:>10s} {1:>10s}'.format('Disp.', 'Force/Width'))
     for disp, force in zip(disptotcor, forcetotcor/b):
         f.write('\n{0:10.6f} {1:10.6f}'.format(disp, force))
     f.close()
plt.title('Specimen {}'.format(var))
plt.show()
filename = 'Fsbs.csv'
f = open(filename, 'w')
for i in Fsbs_list:
    f.write('\n {0:10.6f}'.format(i))
f.close()
