import numpy as np
import pandas as pd
import xlwings as xw
import csv
import json
import time
import random
# reading data form excel
def transformData(filePath):
	wb = xw.Book(filePath)
	RouteInfo = wb.sheets[0]
	OrderInfo = wb.sheets[1]
	order = OrderInfo.range("A1").expand().options(pd.DataFrame).value
	route = RouteInfo.range("A1").expand().options(pd.DataFrame).value

	order['Tax Percentage'][order['Journey Type'] == 'Domestic'] = 0
	route['Cost'] = route[route.columns[6:11]].sum(axis=1)
	route['Time'] = np.ceil(route[route.columns[13:17]].sum(axis=1)/24)
	route = route[list(route.columns[0:4])+ ['Fixed Freight Cost', 'Time', \
	                                     'Cost', 'Warehouse Cost', 'Travel Mode', 'Transit Duty', 'Risk Value'] + list(route.columns[-9:-2])]	
	route = pd.melt(route, id_vars=route.columns[0:11], value_vars=route.columns[-7:] \
                         , var_name='Weekday', value_name='Feasibility')

	route['Weekday'] = route['Weekday'].replace({'Monday': 1, 'Tuesday': 2, 'Wednesday': 3, \
	                                                 'Thursday': 4, 'Friday': 5, 'Saturday': 6, 'Sunday': 7})
	return order, route

def set_param(order, route):

        bigM = 100000
        route = route[route['Feasibility']==1]
        route = route.reset_index()

        portSet = set(route['Source'])| set(route['Destination'])
        portIndex = dict(zip(range(len(portSet)), portSet))
        indexPort = dict(zip(portIndex.values(), portIndex.keys()))
        kStartPort = np.array(order["Ship From"])
        kStartPort = np.array(order["Ship From"].replace(indexPort))
        kEndPort = np.array(order['Ship To'].replace(indexPort))
        portSpace = len(portSet)
        maxDate = np.max(order['Required Delivery Date'])
        minDate = np.min(order['Order Date'])
        dateSpace = (maxDate-minDate).days
        goods = order.shape[0]
        goodsName = np.array(order['Commodity'])
        kOrderDate = np.array((order['Order Date'] - minDate).dt.days)
        kDDL = np.array((order['Required Delivery Date'] - minDate).dt.days)

        startWeekday = minDate.weekday()+1
        weekday = np.mod((np.arange(dateSpace) + startWeekday), 7)
        weekday[weekday == 0] = 7

        weekdayDateList = {i: [] for i in range(1, 8)}
        for i in range(len(weekday)):
                weekdayDateList[weekday[i]].append(i)
        for i in weekdayDateList:
                weekdayDateList[i] = json.dumps(weekdayDateList[i])
        source = list(route['Source'].replace(indexPort))
        destination = list(route['Destination'].replace(indexPort))
        DateList = list(route['Weekday'].replace(weekdayDateList).apply(json.loads))
        kVol = np.array(order['Volume'])
        tranTime,cVol, tranCost, tranFixCost, tDuty, riskValue, riskPer = {},{},{}, {}, {}, {}, {}
        for i in range(portSpace):
         for j in range(portSpace):
          cVol[i,j] = 0.1
          tDuty[i,j] = bigM
          riskValue[i,j] = bigM
          riskPer[i,j] = bigM
          for t in range(dateSpace):
                tranTime[i,j,t] = bigM
                tranCost[i,j,t] = bigM
                tranFixCost[i,j,t] = bigM
        for i in range(route.shape[0]):
         cVol[source[i], destination[i]] = route['Container Size'][i]
         tDuty[source[i], destination[i]] = route['Transit Duty'][i]
         riskValue[source[i], destination[i]] = route['Risk Value'][i]
         riskPer[source[i], destination[i]] = (route['Risk Value'][i]+route['Risk Value'][i]*0.4)
         for t in DateList[i]:
               tranTime[source[i], destination[i], t] = route['Time'][i]
               tranCost[source[i], destination[i], t] = route['Cost'][i]
               tranFixCost[source[i], destination[i], t] = route['Fixed Freight Cost'][i]

        route['Warehouse Cost'][route['Warehouse Cost'].isnull()] = bigM
        whCost = route[['Source', 'Warehouse Cost']].drop_duplicates()
        whCost['index'] = whCost['Source'].replace(indexPort)
        whCost = np.array(whCost.sort_values(by='index')['Warehouse Cost'])
        kValue = np.array(order['Order Value'])
        taxPct = np.array(order['Tax Percentage'])
        return portSpace, dateSpace, goods, kStartPort, kEndPort, portIndex, tranTime, kOrderDate,\
               cVol, kVol, kDDL, tranCost, tranFixCost, whCost, taxPct,kValue,tDuty, goodsName, minDate, indexPort, riskValue, riskPer

# Cost functions
def tCost(Y, Z, C, FC):
    cost = sum(Y[i,j,t]*C[i,j,t]+Z[i,j,t]*FC[i,j,t] for i in range(portSpace) for j in range(portSpace) for t in range(dateSpace))
    return cost
def wCost(X, tranTime, whCost, kVol):
   a = sum(sum(sum(t*X[i,j,t,k] for j in range(portSpace) for t in range(dateSpace))*kVol[k] for k in range(goods))*whCost[i] for i in range(portSpace))
   b = sum(sum(sum((t + tranTime[i,j,t])*X[i,j,t,k] for i in range(portSpace) for t in range(dateSpace))*kVol[k] for k in range(goods) if kEndPort[k] != j)\
            *whCost[j] for j in range(portSpace))
   return  a-b
def taxCost(X,taxPct,kValue,tDuty ):
   a = sum(taxPct[k]*kValue[k] for k in range(goods))
   b = sum(sum(X[i,j,t,k]*kValue[k] for t in range(dateSpace) for k in range(goods))*tDuty[i,j] for i in range(portSpace) for j in range(portSpace))
   return  a+b

order, route = transformData("data.xlsx")
portSpace, dateSpace, goods, kStartPort, kEndPort, portIndex, tranTime, kOrderDate, cVol, kVol, \
kDDL, tranCost, tranFixCost, whCost, taxPct,kValue,tDuty, goodsName, minDate, indexPort, riskValue, riskPer = set_param(order,route)

# adjust variables
Gama = 1
costLimit = 230716
start = time.time()

#optimization model
from gurobipy import*
m = Model()

X = m.addVars(portSpace, portSpace, dateSpace, goods, vtype ='b', name = 'x')
Y = m.addVars(portSpace, portSpace, dateSpace, vtype ='i', name = 'y')
Z = m.addVars(portSpace, portSpace, dateSpace, vtype ='b', name = 'z')
W = m.addVar(vtype ='c', name = 'W')
Mu = m.addVar(vtype ='c', name = 'Mu')
m.update()

m.setObjective( W, GRB.MINIMIZE)

#The constraint numbers and names are the same as equation numbers in the text.
#constr 27
m.addConstr(quicksum(Y[i,j,t]*riskValue[i,j] for i in range(portSpace) for j in range(portSpace) for t in range(dateSpace))+ Gama*Mu <= W)
#constr 28
for i in range(portSpace):
   for j in range(portSpace):
         for t in range(portSpace):
               m.addConstr(Y[i,j,t]*riskPer[i,j] <= Mu)
#constr 29
m.addConstr(tCost(Y, Z,tranCost, tranFixCost )+wCost(X, tranTime, whCost, kVol)+taxCost(X,taxPct,kValue,tDuty ) <= costLimit)

#constr 14,15
for k in range(goods):
        m.addConstr(quicksum(X[kStartPort[k],j,t,k] for j in range(portSpace) for t in range(dateSpace)) == 1)
        m.addConstr(quicksum(X[i,kEndPort[k],t,k] for i in range(portSpace) for t in range(dateSpace)) == 1)

#constr 16,17
for k in range(goods):
        m.addConstr(quicksum(X[i,kStartPort[k],t,k] for i in range(portSpace) for t in range(dateSpace)) == 0)
        m.addConstr(quicksum(X[kEndPort[k],j,t,k] for j in range(portSpace) for t in range(dateSpace)) == 0)
 
#constr 18
for k in range(goods):
    for j in range(portSpace):
            if j != kStartPort[k] and j != kEndPort[k]:
                m.addConstr(quicksum(X[i,j,t,k] for i in range(portSpace) for t in range(dateSpace))
                               -quicksum(X[j,i,t,k] for i in range(portSpace) for t in range(dateSpace)) == 0)
#constr 19,20
for k in range(goods):
    for i in range(portSpace):
            m.addConstr(quicksum(X[i,j,t,k] for t in range(dateSpace) for j in range(portSpace)) <= 1)
for k in range(goods):
    for j in range(portSpace):
            m.addConstr(quicksum(X[i,j,t,k] for t in range(dateSpace) for i in range(portSpace)) <= 1)        
#constr 21, 22      
for k in range(goods):
    for i in range(portSpace):
            if i != kStartPort[k] and i != kEndPort[k]:
                m.addConstr((quicksum(t*X[i,j,t,k] for j in range(portSpace) for t in range(dateSpace))
                -quicksum(t*X[j,i,t,k] + tranTime[j,i,t]*X[j,i,t,k] for j in range(portSpace) for t in range(dateSpace))) >= 0)            
for k in range(goods):
     m.addConstr((quicksum(t*X[kStartPort[k],j,t,k] for j in range(portSpace) for t in range(dateSpace))
        -quicksum(t*X[j,kStartPort[k],t,k] + tranTime[j,kStartPort[k],t]*X[j,kStartPort[k],t,k]
                  for j in range(portSpace) for t in range(dateSpace))) >= kOrderDate[k])
#constr 23,24
for i in range(portSpace):
 for j in range(portSpace):
  for t in range(dateSpace):
          m.addConstr(cVol[i,j]*Y[i,j,t] - quicksum(X[i,j,t,k]*kVol[k] for k in range(goods))>= 0)
          m.addConstr(Z[i,j,t] - quicksum(X[i,j,t,k] for k in range(goods))*(10**-5) >= 0)
#constr 25
for k in range(goods):
        m.addConstr(quicksum(t*X[i,kEndPort[k],t,k] + tranTime[i,kEndPort[k],t]*X[i,kEndPort[k],t,k]\
                             for i in range(portSpace) for t in range(dateSpace)) <= kDDL[k] )
m.optimize()
end = time.time()
print('Solution time', end - start)
print ('Risk Value :', m.objVal)

# Optimal transportation cost      
tCost = sum(Y[i,j,t].x*tranCost[i,j,t]+Z[i,j,t].x*tranFixCost[i,j,t] \
                                    for i in range(portSpace) for j in range(portSpace) for t in range(dateSpace))
#Optimal warehouse cost
wa = sum(sum(sum(t*X[i,j,t,k].x for j in range(portSpace) for t in range(dateSpace))*kVol[k] for k in range(goods))*whCost[i] for i in range(portSpace))
      
wb = sum(sum(sum((t + tranTime[i,j,t])*X[i,j,t,k].x for i in range(portSpace) for t in range(dateSpace))*kVol[k] for k in range(goods)if kEndPort[k] != j)\
            *whCost[j] for j in range(portSpace))
#Optimal tax cost
ta = sum(taxPct[k]*kValue[k] for k in range(goods))
tb = sum(sum(X[i,j,t,k].x*kValue[k] for t in range(dateSpace) for k in range(goods))*tDuty[i,j] for i in range(portSpace) for j in range(portSpace))

print ('Transportation Cost  :', tCost)
print ('Werhouse Cost  :', wa-wb)
print ('Tax Cost  :', ta+tb)
print ('sumCost   :', tCost + wa-wb +ta+tb)

def xValues(x):
   temp, temp2 = [], []
   for i in range(portSpace):
      for j in range(portSpace):
         for t in range(dateSpace):
            for k in range(goods):
                if  x[i,j,t,k].x > 0:
                        temp2 = [portIndex[i],portIndex[j], (minDate + pd.to_timedelta(t, unit='days')).date().isoformat(), k]
                        temp.append(temp2)
                        temp2 = []
   temp = sorted(temp, key = lambda x: x[:][2])
   temp = sorted(temp, key = lambda x: x[:][3])
   return temp 

solution_ = {}
for i in range(goods):
   solution_['goods-' + str(i + 1)] = list(filter(lambda x: x[:][3] == i, xValues(X)))

   
# write to txt
def txt_solution(route, order):
    travelMode = dict(zip(zip(route['Source'], route['Destination']), route['Travel Mode']))
    txt = "Solution"
    txt += "\nNumber of goods: " + str(order['Commodity'].count())
    txt += "\nTotal cost: " + str(m.objVal)
    for i in range(order.shape[0]):
        txt += "\n------------------------------------"
        txt += "\nGoods-" + str(i + 1)
        txt += "\nOrder date: " + pd.to_datetime(order['Order Date']).iloc[i].date().isoformat()
        txt += "\nRoute:"
        solution = solution_['goods-' + str(i + 1)]
        route_txt = ''
        a = 1
        for j in solution:
            route_txt += "\n(" + str(a) + ")Date: " + j[2]
            route_txt += "  From: " + j[0]
            route_txt += "  To: " + j[1]
            route_txt += "  By: " + travelMode[(j[0], j[1])]
            a += 1
        txt += route_txt
    return txt

txt = txt_solution(route, order)
with open("Solution.txt", "w") as text_file:
       text_file.write(txt)



