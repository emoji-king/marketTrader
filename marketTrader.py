# marketTrader.py v1.0
# WIP; still using paper trades. Also, I apologise for some sloppiness down in the code.
import config
import json
import os
import requests
import time
import xlsxwriter
import xlwings as xw
import xlwings.constants
from bs4 import BeautifulSoup
from datetime import datetime, time as Time
from holidays import US as holidaysUS
from pyautogui import alert
from pytz import timezone
from sklearn.neural_network import MLPRegressor
from sys import exit
from yahoo_fin import stock_info as si



'''
Settings for the Program;
num_of_stocks determines how many stocks will be searched for;
If notify is True, it will notify you when a trade is made;
If check_close is true, the program will close itself if the market is closed(if False, may not work correctly!);
Make sure you have a config.py file in the local directory that contains an API key(ex in README)
'''
num_of_stocks = 10
notify = True
check_close = False


def afterHours(now=None):
    # checks if the market is open(only used in autonomous trading bot)
    tz = timezone('US/Eastern')
    us_holidays = holidaysUS()
    if not now:
        now = datetime.now(tz)
    openTime = Time(hour=9 , minute=30 , second=0)
    closeTime = Time(hour=16 , minute=0 , second=0)
    # If a holiday
    if now.strftime('%Y-%m-%d') in us_holidays:
        return True
    # If before 0930 or after 1600
    if (now.time() < openTime) or (now.time() > closeTime):
        return True
    # If it's a weekend
    if now.date().weekday() > 4:
        return True
    return False

def getTopStocks():
    global stock_list
    data = requests.get('https://finance.yahoo.com/screener/predefined/day_gainers')
    if data.status_code != 200:
        print(f'Error! Status code was {data.status_code}.')
    html_data = BeautifulSoup(data.content , 'html.parser')
    temp = list(html_data.children)[1]
    temp = list(temp.children)[1]
    temp = list(temp.children)[2]
    html = str(list(temp.children)[0])
    stock_list = html[html.find('"pageCategory":"YFINANCE:') + 25:html.find('","fallbackCategory":')].split(',')[0:num_of_stocks]   #list indexed starts with meta data object, so add 1 to index because index[0] is removed
    print(f':Top {len(stock_list)} stocks: {stock_list}')


def getStockData(stock):
    global stock_list
    if afterHours() is True and check_close is True:
        return False
    if stock is 'All':
        stock_list = stock_list
    elif type(stock) is list:
        stock_list = stock
    else:
        stock_list = []
        stock_list.append(stock)
    for stock in stock_list:
        try:
            resp = requests.get('https://www.alphavantage.co/query?', {'function':'TIME_SERIES_INTRADAY','symbol':stock,'interval': '1min', 'apikey': config.API_KEY})
            if 'Error Message' in resp.text:
                print(f':Removed {stock} from list.')
                stock_list.remove(stock)
        except:
            stock_list.remove(stock)
        time.sleep(5)
    # stores data for each stock in an excel file
    stockDict = {}
    for symbol in stock_list:
        if os.path.exists(f'{symbol}.xlsx'):
            os.remove(f'{symbol}.xlsx')
        with open(f'{symbol}.xlsx', 'w') as w:
            w.write('')
        # gets stock data
        data = {
            'function': 'TIME_SERIES_INTRADAY' ,
            'symbol': symbol ,
            'interval': '1min' ,
            'outputsize': 'full' ,
            'datatype': 'json' ,
            'apikey': config.API_KEY
        }
        response = requests.get('https://www.alphavantage.co/query?', params=data)
        try:
            dummy = response.json()['Time Series (1min)']
        except Exception as exc:
            print(f':While getting stock data for {symbol}, got error from AlphaVantage... Waiting 15 seconds...')
            with open('error.txt', 'a') as a:
                a.write(f'Error: {exc}\njson data: {response.json()}')
            time.sleep(15)
            getStockData(stock_list[stock_list.index(symbol): len(stock_list)])
            break
        data = response.json()['Time Series (1min)']
        # transfers data to excel
        stock_finances = []
        workbook = xlsxwriter.Workbook(f'{symbol}.xlsx')
        worksheet = workbook.add_worksheet('stock data')
        worksheet.write(0 , 1 , "open")
        worksheet.write(0 , 2 , 'high')
        worksheet.write(0 , 3 , 'low')
        worksheet.write(0 , 4 , 'close')
        worksheet.write(0 , 5 , 'volume')
        d_count = 1
        for day in data.keys():
            stock_data = data[day]
            stock_finances.append(stock_data)
            worksheet.write(d_count , 0 , str(day))
            worksheet.write(d_count , 1 , float(stock_data['1. open']))
            worksheet.write(d_count , 2 , float(stock_data['2. high']))
            worksheet.write(d_count , 3 , float(stock_data['3. low']))
            worksheet.write(d_count , 4 , float(stock_data['4. close']))
            worksheet.write(d_count , 5 , float(stock_data['5. volume']))
            d_count += 1
        # insert formulas
        i , maxRow = 1 , worksheet.dim_rowmax + 1
        num_format = workbook.add_format({'num_format': '0.00;0.00'})
        percent_format = workbook.add_format({'num_format': '0.0000%'})
        # avg per day
        worksheet.write(0 , 7 , 'Day Avg.')
        while i < maxRow:
            worksheet.write(i , 7 , f'=(B{i + 1}+E{i + 1})/2' , num_format)
            i += 1
        worksheet.write(maxRow - 2 , 8 , 'Avg(Total)')
        worksheet.write(maxRow - 1 , 8 , f'=AVERAGE(H2:H{maxRow})' , num_format)
        # total percent change
        worksheet.write(maxRow - 2 , 10 , 'Total %change')
        worksheet.write(maxRow - 1 , 10 , f'=((E2-B101)/B101)*100' , num_format)
        workbook.close()
        # gets formula results to compile a dict
        wb = xw.Book(f'{symbol}.xlsx')
        percentChange = wb.sheets['stock data'].range(f'K{maxRow}').value
        totalAvg = wb.sheets['stock data'].range(f'I{maxRow}').value
        wb.app.quit()
        stockDict[symbol] = f'[{percentChange}, {totalAvg}]'
        with open("stockData.json" , "a") as append:
            json.dump(stockDict , append)
        # displays formula results
        if percentChange != None:
            print(f':{symbol} total change is {str(percentChange)[0:str(percentChange).find(".") + 2]}%')
        else:
            print(f':{symbol} encountered an error during parsing... Removing from list.')
            if os.path.exists(f'{symbol}.xlsx'):
                os.remove(f'{symbol}.xlsx')
            stock_list.remove(symbol)


def neuralNetPredition(symbol):
    print(f':Starting neural network training for {symbol}...')
    wb = xw.Book(f'{symbol}.xlsx')
    # gets value of maxRow
    RR2 = wb.sheets['stock data'].api.Cells.Find(What="*" ,
                                                 After=wb.sheets['stock data'].api.Cells(1 , 1) ,
                                                 LookAt=xlwings.constants.LookAt.xlPart ,
                                                 LookIn=xlwings.constants.FindLookIn.xlFormulas ,
                                                 SearchOrder=xlwings.constants.SearchOrder.xlByRows ,
                                                 SearchDirection=xlwings.constants.SearchDirection.xlPrevious ,
                                                 MatchCase=False)
    maxRow = RR2.Row
    if maxRow << 500:
        pass
    x_list , y_list = [] , []
    row = 3
    list = wb.sheets['stock data'].range(f'B{row}:E{row}').value
    cellRange = range(row , maxRow)
    # creates x and y data sets
    for row in cellRange:
        y = wb.sheets['stock data'].range(f'E{row - 1}').value
        values = wb.sheets['stock data'].range(f'B{row}:E{row}').value
        x_list.append(values)
        y_list.append(y)
    # runs neural network model
    row_to_predict = 2
    cells = f'B{row_to_predict}:E{row_to_predict}'
    clf = MLPRegressor(solver='lbfgs' , alpha=1e-5 , random_state=1)
    clf.fit(x_list , y_list)
    guess = clf.predict([wb.sheets['stock data'].range(cells).value])
    print(f':Model prediciton for {symbol}: {guess}')
    with open('log.txt' , 'a') as aF:
        aF.write(
            f":{symbol}'s machine learning data: guess={guess};\npredicitonX_value={wb.sheets['stock data'].range(cells).value};\nmaxRows={maxRow};\n\n")
    prevClose = wb.sheets['stock data'].range(f'E{row_to_predict}').value
    percentChange = ((guess-prevClose)/prevClose)*100
    wb.app.quit()
    resultList = [float(guess), float(percentChange), prevClose]
    return resultList


def activeTrader(symbol):
    global stock_list
    global prediciton
    print(f'[activeTrader]: Initializing stock {symbol}')
    investment = False
    while True:
        if getStockData([symbol]) is False and check_close is True:
            print(':The market is closed...')
            exit()
        prediction = neuralNetPredition(symbol)
        cur_price = si.get_live_price(symbol)
        #should i invest
        if investment is False:
            if prediction[1] > 0.021:
                print(f'[activeTrader]: Investing in {symbol}; estimated gain is {prediction[1]}; entry price is {cur_price}.')
                with open('trade_log.txt', 'a') as w:
                    w.write(f'Invested at {cur_price}. Predicitons: {prediction};\n')
                investment = True
            else:
                print(f'[activeTrader]: Not investing in {symbol}; estimated change is {prediction[1]}; current price is {cur_price}.')
        #should i sell?
        elif investment is True:
            if prediction[1] < .0115:
                print(f'[activeTrader]: Selling {symbol}; estimated change is {prediction[1]}; exit price is {cur_price}.')
                with open('trade_log.txt', 'a') as w:
                    w.write(f'Sold at {cur_price}. Predicitons: {prediction};\n')
                if notify is True:
                    alert(f'{symbol}: Sold at {cur_price}. Predicitons: {prediction};')
                investment = False
            else:
                print(f'[activeTrader]: Estimated gain for {symbol} is {prediction[1]}; keeping investment; current price is {cur_price}.')


def main():
    action = input(f'1. Get top {num_of_stocks} stocks\n2. Get stock data\n3. Train and run neural network prediction\n4. Run autonomous trade bot\n? ')
    if action == '1':
        getTopStocks()
        main()
    elif action == '2':
        getStockData('All')
        main()
    elif action == '3':
        symbol = input(':Run neural network on what stock? Type in symbol or leave blank to run stock_list.\n? ').upper()
        if symbol == '':
            if len(stock_list) != 0:
                for symbol in stock_list:
                    neuralNetPredition(symbol)
            else:
                getTopStocks()
                for symbol in stock_list:
                    neuralNetPredition(symbol)
        else:
            getStockData(symbol)
            neuralNetPredition(symbol)
    elif action == '4':
        print(':Refreshing stock_list...')
        getTopStocks()
        if getStockData('All') is False and check_close is True:
            print(':The market is closed...')
            exit()
        stockPredictions = []
        for symbol in stock_list:
            results = neuralNetPredition(symbol)
            stockPredictions.append([symbol, results[0], results[1], results[2]])
        preferredStock = ['None', 0]
        for resultSet in stockPredictions:
            if resultSet[2] > preferredStock[1]:
                preferredStock[0] = resultSet[0]
                preferredStock[1] = resultSet[2]
        if preferredStock[0] != 'None':
            print(f':Neural Network preferred stock: {preferredStock[0]}; Estimated Gain: {preferredStock[1]}')
            print(f': Starting Active Trader with preferred stock {[preferredStock[0]]}...')
            stock_list.remove(preferredStock[0])
            for stock in stock_list:
                try:
                    os.remove(f'{stock}.xlsx')
                except:
                    pass
            activeTrader(preferredStock[0])
        else:
            print(f':Neural Network did not determine preferred stock... Run again.')
            main()



if __name__ == '__main__':
    main()