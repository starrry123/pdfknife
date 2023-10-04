import requests,re,itertools
from bs4 import BeautifulSoup
import pandas as pd  # Import pandas

url='https://aicip.org.au/inspector-find/'

def get_page_num(url):
    print('processing...', url)
    page = requests.get(url)
    soup = BeautifulSoup(page.content, 'html.parser')
    numeric_links = [link for link in soup.find('ul', class_='page-numbers').find_all('a') if link.text.isdigit()]
    last_page_number = int(numeric_links[-1].text)
    return last_page_number

def scrape_data(url):
    print('processing...', url)
    page = requests.get(url)
    soup = BeautifulSoup(page.content, 'html.parser')
    mydivs = soup.findAll("div", {"class": "table-inspector_inner"})
    records = []

    for div in mydivs:
        data = {}
        name = div.find('h5').text
        data['Name'] = name
        for tr in div.findAll('tr'):
            td = tr.find('td')
            if td:
                key, value = map(str.strip, td.text.split(':', 1))
                data[key] = value

        records.append(data)

    return records

# Create an empty list to store all the records
columns = ['Name', 'Registration No', 'Qualification', 'Date Valid to', 'Status', 'State',
           'Modules', 'Email', 'Mobile No', 'Suburb']
all_records = []

for i in range(1, get_page_num(url)+1):
    if i == 1:
        url = 'https://aicip.org.au/inspector-find/'
    else:
        url = 'https://aicip.org.au/inspector-find/page/' + str(i) + '/'
    all_records.extend(scrape_data(url))

# Create a DataFrame from the list of records
df = pd.DataFrame(all_records, columns=columns)

# Remove duplicates based on the 'Name' column
df.drop_duplicates(subset='Name', keep='last', inplace=True)

# Save the DataFrame to an Excel file
excel_filename = 'aicip_inspectors_data.xlsx'
df.to_excel(excel_filename, index=False)

print(f'Data saved to {excel_filename}')
