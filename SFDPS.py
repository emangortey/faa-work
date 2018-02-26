# -*- coding: utf-8 -*-
"""
Created on Mon Jan 29 11:02:35 2018

@author: gdard3
"""

import pandas as pd
import numpy as np
import xlrd as xl
from xlwt import Workbook

#xlrd
#open workbook
sample = xl.open_workbook('C:\\Users\\gdard3\\Georgia Institute of Technology\\Fischer, Olivia J - DataFusionGC1718\\FAA\\DataAnalysisWork\\SFDPS Data cleaning\\All SFDPS Excel Files\\FLTR_SFDPS_2017042123.xlsx')
#print(sample.sheet_names())
#open the 1st sheet
sample1 = sample.sheet_by_name('Sheet1')
info = sample1.col_values(1)

#create new workbook
final = Workbook()
results = final.add_sheet('Results', cell_overwrite_ok = True)
line1 = results.row(0)

#compute the number of flights
flights = []
for k in range (0, len(info)-1):
    if info[k] == 'FDPSMsg xmlns="us:gov:dot:faa:atm:enroute:entities:flightdata">':
        flights.append(k+1)
flights.append(len(info))
#print(len(flights)-1)

#first loop only for the first flight        
head = []
foot = []        

for i in range (flights[0]+1,flights[1]-1):
    position = info[i].find('>')
    if position != -1:
        front = info[i][0:position]
        back = info[i][position+1:len(info[i])]
        head.append(front)
        if back == '':
            foot.append('//')
        else:
            foot.append(back)



#general loop to add all flights in the Excel file  
for l in range (1, len(flights)-1):
    for x in range (flights[l]+1,flights[l+1]-1):
        position = info[x].find('>')
        if position != -1:
            front = info[x][0:position]
            back = info[x][position+1:len(info[x])]
            if front in head:
                pos = head.index(front)
                if back =='':
                    results.row(l+1).write(pos, '//')
                else: 
                    results.row(l+1).write(pos, back)
            else: 
                head.append(front)
                pos = head.index(front)
                if back =='':
                    results.row(l+1).write(pos, '//')
                else: 
                    results.row(l+1).write(pos, back)
                foot.append('')
                
        
#write in the excel file the two first lines             
for m in range (0, len(head)):
    results.row(0).write(m,head[m])
    results.row(1).write(m,foot[m])


#RequiredProp = final.add_sheet('RequiredProperties', cell_overwrite_ok = True)
#sample1 = sample.sheet_by_name('Sheet1')
#info = sample1.col_values(1)
#
##create new workbook
#final = Workbook()
#results = final.add_sheet('Results', cell_overwrite_ok = True)
#line1 = results.row(0)

#column1=results.col_values(0)
#line1=results.lin_values(0)

            
               
final.save('C:\\Users\\gdard3\\Georgia Institute of Technology\\Fischer, Olivia J - DataFusionGC1718\\FAA\\DataAnalysisWork\\SFDPS Data cleaning\\All SFDPS results\\results23.xls')
    