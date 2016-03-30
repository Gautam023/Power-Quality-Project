# -*- coding: utf-8 -*-
"""
Created on Wed Mar 30 12:07:36 2016

@author: Gautam
"""
import xlrd
import matplotlib.pyplot as plt
import numpy as np
import math

#Desigin Filter
def build_filter(Q,n,Vs,QF):
    w = 2*3.141*50
    C = Q/(n*w*(Vs^2))
    L = 1/(n*n*w*w*C)
    R = (n*w*L)/QF
    return C,L,R

#Checking for Harmonic Resonance
def check_harmonic(C,L,n):
    if 50 == 1/(2*3.141*n*math.sqrt(C*L)):
        return True
    else:
        return False

#Opeing Excel File to Read data        
book1 = xlrd.open_workbook("C:\Users\gautam\Desktop\PQ-Data-2system.xlsx")
book2 = xlrd.open_workbook("C:\Users\gautam\Desktop\PQ-Data-4system.xlsx")
book3 = xlrd.open_workbook("C:\Users\gautam\Desktop\PQ-Data-6system.xlsx")
book4 = xlrd.open_workbook("C:\Users\gautam\Desktop\PQ-Data-9system.xlsx")

first_sheet_1 = book1.sheet_by_index(0)
first_sheet_2 = book2.sheet_by_index(0)
first_sheet_3 = book3.sheet_by_index(0)
first_sheet_4 = book4.sheet_by_index(0)

#Collecting Data
IHD3_data1 = first_sheet_1.col_values(colx=3,start_rowx=1,end_rowx=157)
IHD3_data2 = first_sheet_2.col_values(colx=3,start_rowx=1,end_rowx=157)
IHD3_data3 = first_sheet_3.col_values(colx=3,start_rowx=1,end_rowx=157)
IHD3_data4 = first_sheet_4.col_values(colx=3,start_rowx=1,end_rowx=157)


reactive_power = first_sheet_4.col_values(colx=9,start_rowx=1,end_rowx=157)
Reactive_power = (reactive_power)
avg_reactive_power = np.mean(Reactive_power)
max_reactive_power = max(reactive_power)
filter_values_3 = build_filter(max_reactive_power,3,221,50)
filter_values_5 = build_filter(max_reactive_power,5,221,50)
harmonic_resonance_3 = check_harmonic(filter_values_3[0],filter_values_3[1],3)
harmonic_resonance_5 = check_harmonic(filter_values_5[0],filter_values_5[1],5)

#Logging
log = open("C:\Users\gautam\Desktop\PQ Datalog.txt","a")
log.write("Max Reactive Power: "+str(max_reactive_power))
log.write("\nFilter values for 3rd harmonic component: "+str(filter_values_3))
log.write("\nFilter values for 5th harmonic component: "+str(filter_values_5))
log.write("\nHarmonic Resonance for 3rd harmonic filter: "+str(harmonic_resonance_3))
log.write("\nHarmonic Resonance for 5th harmonic filter: "+str(harmonic_resonance_5))
log.write("\n------------------------------------------------------------------------\n")
log.close()

x = np.arange(0.0,156.0,1)
fig1 = plt.figure(num=1, figsize=(12, 12), dpi=80, facecolor='w', edgecolor='k')
ax   = fig1.add_subplot(111)
ax.set_xlabel('Time')
ax.set_ylabel('IHD3')
ax.set_title('IHD3 for different loads')
plt.plot(x,IHD3_data1,'b-',label="2system")
plt.plot(x,IHD3_data2,'g-',label="4system")
plt.plot(x,IHD3_data3,'r-',label="6system")
plt.plot(x,IHD3_data4,'y-',label="9system")
plt.legend()
plt.show()
