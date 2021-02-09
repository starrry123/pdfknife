#ATO tax rate For FY 2021-2021 

tax_rate=[[180000,0.45,51667],
          [120000,0.37,29467],
          [45000,0.325,5092],
          [18200,0.19,0]]

def cal_tax(m):
    if m>tax_rate[0][0]:
        return tax_rate[0][2]+tax_rate[0][1]*(m-tax_rate[0][0])
    elif m<18201:
        return 0
    else:
        tax_rate.pop(0)
        return cal_tax(m)
income=0
tax_data=[]
income_data=[]
k=200
for i in range(0,k):
    income+=1000
    income_data.append(income)
    tax_data.append(cal_tax(income))
    tax_rate=[[180000,0.45,51667],
          [120000,0.37,29467],
          [45000,0.325,5092],
          [18200,0.19,0]]

import plotly.express as px
fig = px.line(x=income_data,y=tax_data,title='ATO tax rate')
fig.show()
