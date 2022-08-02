# -*- coding: utf-8 -*-
import sys
import xlrd
import numpy as np
import matplotlib.pyplot as plt
# import pandas as pd
# import matplotlib.pyplot as plt
# Load xls file
data = xlrd.open_workbook('107_Xitun2.xls')
#print(data.sheets()[0].row_values(1)[0])
for n in range(len(data.sheet_names())):
    table = data.sheets()[n]
    pmData = []
    coData = []
    so2Data = []
    tempPMData = []
    tempCOData = []
    tempSO2Data = []
    
    for i in range(table.nrows):
        #print('Page {}: '.format(n), end='')
        if i != 0:
            lastCount = table.row_values(i)[0] #當前這筆日期
            #print("last" + str(lastCount))
            preCount =  table.row_values(i-1)[0] #上一筆日期
            #print("pre" + str(preCount))
            #if lastCount != preCount
            if lastCount[5:7] != preCount[5:7] or i == 7300:
                #print("last" + str(lastCount))
                #print("pre" + str(preCount))
                pmTotal = 0.00000
                # if i == 1802:
                #     pmTotal += 44
                # elif i == 2402:
                #     pmTotal += 53
                
                coTotal = 0.00000
                so2Total = 0.00000

                pmAvg = 0.00000
                coAvg = 0.00000
                so2Avg = 0.00000
                
                #print(tempPMData)
                #print(tempCOData)
                #print(tempSO2Data)

                for j in range(len(tempPMData)):
                    pmTotal += tempPMData[j]
                    if j == len(tempPMData)-1:
                        pmAvg = pmTotal/len(tempPMData)
                #print("AVGG:PMM:" + str(pmAvg))
                pmData.append(pmAvg) #儲存
                tempPMData = [] #清除
                #print("")

                for j in range(len(tempCOData)):
                    coTotal += tempCOData[j]
                    if j == len(tempCOData)-1:
                        coAvg = coTotal/len(tempCOData)
                #print("AVGG:COO:" + str(coAvg))
                coData.append(coAvg) #儲存
                tempCOData = [] #清除
                #print("")

                for j in range(len(tempSO2Data)):
                    so2Total += tempSO2Data[j]
                    if j == len(tempSO2Data)-1:
                        so2Avg = so2Total/len(tempSO2Data)
                #print("AVGG:SO2:" + str(so2Avg))
                so2Data.append(so2Avg) #儲存
                tempSO2Data = [] #清除
                #print("")
        #print(str(table.row_values(i)[15]).find("#"))
        if str(table.row_values(i)[12]).find("#") != -1 or str(table.row_values(i)[12]).find("NR") != -1 or str(table.row_values(i)[12]).find("x") != -1:
            #print("#############################")
            continue
        # else:
        #     print("SSSSSSSS")
        if table.row_values(i)[2] == "PM2.5":
            # print("PM2.5 DATA:")
            # print(i)
            # print(table.row_values(i)[0]) 
            # print(str(table.row_values(i)[12]))
            if(i == 5131): #1271 (3) 44 2311 (4) 53
                continue
            tempPMData.append(float(table.row_values(i)[12]))     
        elif table.row_values(i)[2] == "CO":
            if i == 5123: #空值
                continue
            # print("CO DATA:")
            # print(i)
            # print(table.row_values(i)[0])  
            # print(table.row_values(i)[12])
            tempCOData.append(float(table.row_values(i)[12]))  
        elif table.row_values(i)[2] == "SO2":
            if i == 5135: #空值
                continue
            # print(i)
            # print("SO2 DATA:")
            # print(table.row_values(i)[0]) 
            # print(table.row_values(i)[12])
            tempSO2Data.append(float(table.row_values(i)[12]))
pmData.pop(0)
coData.pop(0)
so2Data.pop(0)
print("PM2.5 1~12月平均資料")
for n in range(12):
    print(str(n+1)+"月份:"+str(pmData[n]))
print("")
print("CO 1~12月平均資料")
for n in range(12):
    print(str(n+1)+"月份:"+str(coData[n]))
print("")
print("SO2 1~12月平均資料")
for n in range(12):
    print(str(n+1)+"月份:"+str(so2Data[n]))

mouth = ['01','02','03','04','05','06','07','08','09','10','11','12']

x=np.arange(20,350)
y=np.arange(0,30)
l1=plt.plot(mouth,pmData,'r--',label='PM2.5')
l2=plt.plot(mouth,coData,'g--',label='CO')
l3=plt.plot(mouth,so2Data,'b--',label='SO2')
plt.plot(mouth,pmData,'ro-',mouth,coData,'g+-',mouth,so2Data,'b^-')
#plt.title('2018 西屯空汙107_XiTun2.xls  (每日9點)')
plt.xlabel('month')
plt.ylabel('data')
plt.legend()
plt.show()
