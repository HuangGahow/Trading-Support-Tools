# -*- coding: utf-8 -*-
"""
Created on Thu Nov 10 09:11:29 2022

@author: Huangjiahao
"""



import pandas as pd
import numpy as np
import os
np.set_printoptions(suppress=True)

computedate = input("please input date in yyyymmdd:")

path1 =r'\\cicc.group\dfs\DEPT\FID\WY&YJ\0000报备文件\FIRT结算通知\FIRT_结算通知-'+computedate+'.xlsm'

path2 =  r'\\cicc.group\dfs\DEPT\FID\WY&YJ\A bond簿记表-NAFMII SAC ISDA\北向国债期货连接端'+ computedate +'.xlsx'

settlement = pd.read_excel(io = path1,sheet_name='互换连接端结算',header=None )
settlement = settlement.dropna(axis=0,how='any')
settlement.drop([0,1,8,9],axis=1,inplace = True) 
settlement.rename({2:'Date',3:'ReferenceNumber',4:'UnwindNumber',5:'RemainNumber', \
                   6:'pt',7:'p0',10:'PNL',11:'NOA',12:'支付方',13:'SOAR',14:'NOAR'},axis=1, inplace = True)
settlement['UnwindNumber'],settlement['RemainNumber']=settlement['UnwindNumber'].apply(lambda x: round(x,1))*100, settlement['RemainNumber'].apply(lambda x: round(x,1))*100
settlement['SOAR'],settlement['NOAR']=settlement['SOAR'].apply(lambda x: round(x,1)), settlement['NOAR'].apply(lambda x: round(x,1))


uwfuture = pd.read_excel(path2, sheet_name="国债期货结算")
#delete the rows include missing value
uwfuture = uwfuture.dropna(axis=0, subset=['交易编号'])
Date = uwfuture['平仓日期'].iloc[-1]
uwfuture = uwfuture.loc[uwfuture['平仓日期'] == Date]
uwfuture['剩余名义本金（不含价格）']=uwfuture['剩余名义本金（不含价格）'].apply(lambda x: round(x,1))
uwfuture['结算名义本金（不含价格）']=uwfuture['结算名义本金（不含价格）'].apply(lambda x: round(x,1))
#python科学计数法，1.21+e8 =1209999.999
uwfuture.drop(['Unnamed: 0','Unnamed: 3','Unnamed: 10','Unnamed: 11'],axis = 1, inplace = True)

result = uwfuture.merge(settlement, left_on='交易编号', right_on='ReferenceNumber',how='left')


result['DateVeri'] = (result['平仓日期']==result['Date'])
result['UnwindNumberVeri'] = (result['结算名义本金（不含价格）']==result['UnwindNumber'])
result['RemainNumberVeri'] = (result['剩余名义本金（不含价格）']==result['RemainNumber'])
result['P0Veri'] = (result['Unnamed: 4']==result['p0'])
result['PtVeri'] = (result['Unnamed: 5']==result['pt'])
result['PNLVeri'] = (result['结算金额']==result['PNL'])
result['PayerVeri']=(result['支付方_x']==result['支付方_y'])
result['SOARVeri'] = (result['结算名义本金（报备）']==result['SOAR'])
result['NOARVeri'] = (result['剩余名义本金（报备）']==result['NOAR'])

report = result.loc[:,['交易编号','DateVeri', 'UnwindNumberVeri','RemainNumberVeri', 'P0Veri','PtVeri','PNLVeri','PayerVeri','SOARVeri', 'NOARVeri']]

folder = os.getcwd() + '\\Reports\\'
if not os.path.exists(folder):
    os.makedirs(folder) 
report.to_excel(folder + computedate +'-结算通知_Report.xlsx',sheet_name='Report',index=False,header=True)
print("Done")

