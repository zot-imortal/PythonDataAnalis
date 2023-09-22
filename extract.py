import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import plotly.express as px
from tabulate import tabulate
exel_file = 'Aggregation results.xlsx'

exel_file2 = 'Wärmepumpen1.xlsx'
exel_file3 = 'Emob.csv'
require_cols = [0, 2,3, 5,6, 8,9]
df = pd.read_excel(exel_file, sheet_name='Sheet1', usecols = require_cols)
#require_cols2 = [0, 1]
df2 = pd.read_excel(exel_file2, sheet_name='Wärmepumpen')
require_cols3 = [0, 2]
df3 = pd.read_csv(exel_file3, usecols = require_cols)
""" 
max_el = df['PV'].astype(float).max()
listOfIndexes = df.index[df['PV'] == max_el].tolist()/
timeEl = str(df.at[listOfIndexes[0], 'Time'])


dict = {}
firstElTime = str(df['Time'][0])[0:10]
print(firstElTime)
lastElTime = str(df['Time'][len(df['Time'])-1])[0:10]
print(lastElTime)
dict[firstElTime] = ''

while firstElTime != lastElTime:
    date = firstElTime
    list = []
    nextEl = ''
    
    for index, row in df.iterrows():
        if str(row['Time'])[0:10] == firstElTime:           
            list.append(float(row['PV']))
        else: 
            if str(row['Time'])[0:10] in dict.keys():
                next
            else:
                firstElTime = str(row['Time'])[0:10]
                break
    dict[date] = sum(list)
    
for key,values in dict.items():
    print(key, str(round(values, 2)) + ' kW')

MaxKey = max(dict, key=dict.get)
print("The DAY with max production, kW:\n", MaxKey)

print('\nMax produced Ppeak: ' + str(max_el) + ' kW')
print('Date              : ' + timeEl[0:10])

 """
df_filtered = df[df['Time'].dt.strftime('%Y-%m-%d') == '2023-01-31']

df_filtered2 = df2[df2['AGGREGATION'].dt.strftime('%Y-%m-%d') == '2023-01-31']




df_filtered2['WP'] = df_filtered2.iloc[:,2:7].astype(float).sum(1)
df_filtered['WP'] = df_filtered2[['WP']].values
df_filtered['WP + Cons1'] = df_filtered[['Cons1', 'WP']].astype(float).sum(1)
print(tabulate(df_filtered, headers='keys', tablefmt='psql'))
fig1 = px.line(df_filtered, x='Time', y=['PV1', 'Cons1', 'WP + Cons1'])
fig1.show()




df4 = pd.read_excel(exel_file2, sheet_name='Wärmepumpen')
df4['AGGREGATION'] = pd.to_datetime(df4['AGGREGATION'], format='%Y-%m-%d %H:%M')

start_date = '2023-01-31 12:50'
end_date = '2023-02-01 21:00'
mask = (df4['AGGREGATION'] > start_date) & (df4['AGGREGATION'] <= end_date)
df4 = df4.loc[mask]
print(tabulate(df4, headers='keys', tablefmt='psql'))