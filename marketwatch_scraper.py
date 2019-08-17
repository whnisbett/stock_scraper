import openpyxl as xl
import requests
from bs4 import BeautifulSoup

#code getFunctions for all individual attributes

def getCurrentPPS(soup):
    price_div = soup.find('h3', class_ = 'intraday__price')
    curr_price = price_div.find('bg-quote',class_='value').contents[0]
    curr_price = float(curr_price.replace(',',''))
    return curr_price

def getPercentChange(soup):
    close_div = soup.find_all('div', class_='intraday__close')[0]
    per_change = close_div.find_all('td',class_='table__cell not-fixed positive')[1].contents[0][:-1]
    per_change = float(per_change.replace(',',''))
    return per_change

def getSector(soup):
    sector_div = soup.find('div', class_='intraday__sector')
    sector = sector_div.find('span',class_='label').contents[0]
    return sector

def getSectorChange(soup):
    sector_div = soup.find('div', class_ = 'intraday__sector')
    sector_change_div = sector_div.find('span',class_='change--percent positive')
    if len(sector_change_div) == 0:
        sector_change_div = sector_div.find('span',class_='change--percent negative')
    sector_change = sector_change_div.find('span').contents[0][:-1]
    sector_change = float(sector_change.replace(',',''))
    return sector_change

def getStockAttributes(ticker):
    #get the source code, then call individual get functions for each attribute, passing the source code to each one
    url = 'https://www.marketwatch.com/investing/stock/' + ticker.lower()
    page = requests.get(url,'html.parser')
    if page.status_code != 200:
        raise Exception('Failed to retrieve URL contents')
    else:
        attributes={}
        soup = BeautifulSoup(page.content,'html.parser')

        attributes['Current PPS'] = getCurrentPPS(soup)
        attributes['% Change'] = getPercentChange(soup)
        #attributes['Sector'] = getSector(soup)
        attributes['Sector % Change'] = getSectorChange(soup)
        # attributes['5 Day Performance'] = get5Day(soup)
        # attributes['1 Month Performance'] = get1Month(soup)
        # attributes['3 Month Performance'] = get3Month(soup)
        # attributes['1 Year Performance'] = get1Year(soup)
        # attributes['YTD Performance'] = getYTD(soup)
        # attributes['52 Week Range'] = get52WeekRange(soup)
        # attributes['P/E Ratio'] = getPE_Ratio(soup)
        # attributes['EPS'] = getEPS(soup)
        # attributes['Market Cap'] = getMarketCap(soup)
        # attributes['% of Float Shorted'] = getPercentShorted(soup)
        return attributes

wb_path = '/Users/billy/Personal/Finances/Financials.xlsx'
wb = xl.load_workbook(wb_path)
sheet = wb['Stocks']
ticker_col = 1
first_ticker = sheet.min_row + 1
last_ticker = sheet.max_row
attributes_cell_keys = {'Current PPS':'B',
                        '% Change':'C',
                        'Sector':'D',
                        'Sector % Change':'E',
                        '5 Day Performance': 'F',
                        '1 Month Performance': 'G',
                        '3 Month Performance': 'H',
                        '52 Week Range':'I',
                        'P/E Ratio':'J',
                        'EPS':'K',
                        'Market Cap':'L',
                        '% of Float Shorted':'M'}

for r in range(first_ticker,last_ticker + 1):
    ticker = sheet.cell(row = r,column = ticker_col).value
    attributes = getStockAttributes(ticker)
    for key in attributes.keys():
        value = attributes[key]
        column = attributes_cell_keys[key]
        cell = column + str(r)
        print(cell)
        sheet[cell] = value