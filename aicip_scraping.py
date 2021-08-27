import requests,re,itertools
from bs4 import BeautifulSoup
from prettytable import PrettyTable
    
def scrape_data(url):
        print('processing...',url)
        page = requests.get(url)
        soup = BeautifulSoup(page.content, 'html.parser')
        mydivs = soup.findAll("div", {"class": "table-inspector_inner"})
        data=[]
        pretty=PrettyTable()
        pretty.field_names=['Name','Registration No','Qualification','Date Valid to', 'Status', 'State','Modules','Email','Mobile No','Suburb']
        for div in mydivs:
                name=div.find('h5')
                data.clear()
                data.append(name.text)
                for tr in div.findAll('tr'):
                        for td in tr.findAll('td'):
                                v=td.text.strip().split(':')
                                if len(v)==1 and ':' in td.text:
                                        data.append('')
                                elif len(v)==2:
                                        data.append(v[1].strip())
                pretty.add_row(data)
        print (pretty)
        
for i in range(1,33):
        if i==1:
                url='https://aicip.org.au/inspector-find/'
        else:
                url='https://aicip.org.au/inspector-find/page/'+str(i)+'/'
        scrape_data(url)
