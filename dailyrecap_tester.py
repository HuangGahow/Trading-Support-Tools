# -*- coding: utf-8 -*-
"""
Created on Mon Nov 14 10:16:45 2022

@author: Huangjiahao
"""

import pandas as pd
import numpy as np

np.set_printoptions(suppress=True)
"""
支付方，开仓手数
"""
# In[1]


path1 = input("please input your dailyrecap_接口备案's path:")

path2 = input("please input your 北向国债期货连接端's path:")


folder = os.getcwd() + '\\Reports\\'
if not os.path.exists(folder):
    os.makedirs(folder)
    

recapunwind = pd.read_excel(io = path1,sheet_name='互换结算',header=0)


recapunwind.drop(recapunwind.iloc[:,[0,1,2,8]],axis=1,inplace = True) 
recapunwind = recapunwind.dropna(axis=0,how='any')
recapunwind['提前终止标的资产数量']=recapunwind['提前终止标的资产数量'].apply(lambda x: round(x,1))*100
recapunwind['剩余标的资产数量']=recapunwind['剩余标的资产数量'].apply(lambda x: round(x,1))*100


uwfuture = pd.read_excel(path2, sheet_name="国债期货结算")
uwfuture = uwfuture.dropna(axis=0, subset=['交易编号'])
Date = uwfuture['平仓日期'].iloc[-1]
uwfuture = uwfuture.loc[uwfuture['平仓日期'] == Date]
uwfuture['剩余名义本金（不含价格）']=uwfuture['剩余名义本金（不含价格）'].apply(lambda x: round(x,1))
uwfuture['结算名义本金（不含价格）']=uwfuture['结算名义本金（不含价格）'].apply(lambda x: round(x,1))

uwfuture.drop(['Unnamed: 0','Unnamed: 3','Unnamed: 10','Unnamed: 11'],axis = 1, inplace = True)
result = uwfuture.merge(recapunwind, left_on='交易编号', right_on='交易编号',how='left')

result['UnwindDateVeri'] =(result['平仓日期']== result['结算日期'])

result['PtVeri'] =(result['Unnamed: 5']== result['期末价格'])
result['UNAVeri'] =(result['结算名义本金（不含价格）']== result['提前终止标的资产数量'])
result['ONApostVeri'] =(result['剩余名义本金（不含价格）']== result['剩余标的资产数量'])


result['EAVeri']=(result['结算金额']==result['结算金额（元）'])
result['PayerVeri']=(result['支付方_x']==result['支付方_y'])
result['UnwindARVeri']=(result['结算名义本金（报备）']==result['本次平仓'+'\n'+'名义本金CNY'].apply(lambda x: round(x,1)))
result['RemainARVeri']=(result['剩余名义本金（报备）']==result['剩余'+'\n'+'名义本金CNY'].apply(lambda x: round(x,1)))

report = result.loc[:,['交易编号','UnwindDateVeri', \
                       'PtVeri','UNAVeri', 'ONApostVeri',\
                           'EAVeri','PayerVeri','UnwindARVeri','RemainARVeri']]
    
    
    

report.to_excel(folder + 'dailyrecap_Report1.xlsx',sheet_name='互换结算Report',index=False,header=True)
print("Done")



# In[2]
recapswap = pd.read_excel(io = path1,sheet_name='收益互换',header=1)

recapswap.drop(recapswap.iloc[:,[1,2,3,11,12,13,14,15,16]],axis=1,inplace = True)
recapswap = recapswap.dropna(axis=0,how='any') 


bond = pd.read_excel(path2, sheet_name="Bond")
#delete the rows include missing value
bond = bond.dropna(axis=0, subset=['交易确认书编号'])
Date = bond['起始日'].iloc[-1]
bond = bond.loc[bond['起始日'] == Date]
bond.drop(["协议", "方向" ,"客户名称","状态","系统编号","交易达成日", "结算金额（客户方向）", \
           "新备案账号", "标的小类" ,"备注", "Unnamed: 28" ],axis = 1, inplace = True)
    
resultx = bond.merge(recapswap, left_on='交易确认书编号', right_on='交易确认书编号（双方约定）',how='left')
resultx['TradeDateVeri'] =(resultx['起始日_x']== resultx['起始日_y'])
resultx['P0Veri'] =(resultx['期初价格']== resultx['期初价格(计价币种)'])
resultx['ShareNumberVeri'] =(resultx['新开手数']== resultx['标的数量（商品是手数、债券是张数）'])
resultx['B/SVeri'] =(resultx['B/S_x']== resultx['B/S_y'])
resultx['ValuationDateVeri'] =(resultx['到期日_x']== resultx['到期日_y'])
resultx['ContractVeri']=(resultx['标的_x']==resultx['标的_y'])
resultx['NotionalAmountVeri']=(resultx['名义本金CNY']==resultx['名义本金额(人民币)'])

reportx = resultx.loc[:,['交易确认书编号','TradeDateVeri', \
                       'P0Veri','ShareNumberVeri','B/SVeri','ValuationDateVeri', 'NotionalAmountVeri',\
                           'ContractVeri']]

reportx.to_excel(folder +'dailyrecap_Report2.xlsx',sheet_name='收益互换Report',index=False,header=True)
print("Done")
