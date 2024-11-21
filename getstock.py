import requests
from bs4 import BeautifulSoup
import json
import pandas as pd

mystocks = ['ARCO', 'AYI' ,'FUTU', 'TAL', 'STLA', 'SATS']
stockdata = []

def getData(symbol):
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36'}
    url = f'https://finance.yahoo.com/quote/{symbol}/'
    r = requests.get(url, headers=headers)
    soup = BeautifulSoup(r.text, 'html.parser')
    stock = {
    'name' : symbol,
    'price' : soup.find('div', {'class': 'container yf-1tejb6'}).find_all('span')[0].text,
    'change' : soup.find('div', {'class': 'container yf-1tejb6'}).find_all('span')[1].text,
    }
    return stock

for item in mystocks:
    stockdata.append(getData(item))
    print('Getting: ', item)
    
with open('stockdata.json', 'w') as f:
    json.dump(stockdata, f)


df = pd.DataFrame(stockdata)
df.set_index('name', inplace = True)
df.to_excel('stockoutput.xlsx')
print("Fin.")
