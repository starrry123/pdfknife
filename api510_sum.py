import re
f=open(r'api510_cn.txt','r',encoding='utf8')

name=re.compile('Name:(.+)', re.UNICODE)
add1=re.compile('Address:(.*)', re.UNICODE)
add2=re.compile('City/State/PostalCode/Country:(.*)', re.UNICODE)
pcode=re.compile('Phone Country Code:(.*)', re.UNICODE)
phone=re.compile('Phone:(.*)', re.UNICODE)   
cert=re.compile('Certification[s]?:(.*)', re.UNICODE)
ind=re.compile('Industry Type:(.*)', re.UNICODE)
matchers=(name, add1, add2,pcode,phone,cert,ind)
data=[]
b=[]

line=f.read()

for matcher in matchers:
    m=matcher.findall(line)
    m2=[re.sub('\t','',x) for x in m]
    m3=[re.sub(',',' ',x) for x in m2]
    data.append(m3)
   
b=list(zip(data[0],data[1],data[2],data[3],data[4],data[5],data[6]))
import csv
with open('api_test.csv','w',newline='',encoding='utf8') as file:
    file.write("\uFEFF")
    writer=csv.writer(file,delimiter=';')
    writer.writerows(b)

