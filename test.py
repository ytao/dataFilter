import pandas as p
import os
import sys
import numpy as np

type='dq'

x=p.read_excel('./source/actuals/'+type+'.xlsx')
a_length=len(x['备件名称'])

contract_path='./source/contracts/'+type+'/'

c_files=os.listdir(contract_path)
m=0

for cf in c_files:
    mf = contract_path + cf
    s=p.read_excel(mf,header=0,sheetname='Sheet1')
    s = s[['名称','型号','单价']]
    s['数量']=0
    list_name=s['名称'].tolist()
    list_model=s['型号'].tolist()
    for i in range(0,a_length-1):
        a_name=x['备件名称'][i]
        a_model=x['型号'][i]
        # print(a_name + ':\t' + str(a_model))
        if a_name in list_name and a_model in list_model:
            mindex=list_name.index(a_name)
            m_num=int(s['数量'][mindex])
            # print('m_num=' + str(m_num))
            m_num += x['数量'][i]
            # s['数量'][mindex] = m_num
            s.at[mindex,'数量']=m_num
    s.to_excel('./out/'+cf)
