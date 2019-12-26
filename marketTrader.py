# marketTrader.py v1.1
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
from os import listdir
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
num_of_stocks = 5
notify = True
check_close = True


def clean():
    print(':Deleted old .xlsx files.')
    for file_name in listdir(os.path.dirname(os.path.abspath(__file__))):
        if file_name.endswith('.xlsx'):
            os.remove(os.path.dirname(os.path.abspath(__file__)) + '/' + file_name)


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
        print('Error! Status code was {}.'.format(data.status_code))
    html_data = BeautifulSoup(data.content , 'html.parser')
    temp = list(html_data.children)[1]
    temp = list(temp.children)[1]
    temp = list(temp.children)[2]
    html = str(list(temp.children)[0])
    stock_list = html[html.find('"pageCategory":"YFINANCE:') + 25:html.find('","fallbackCategory":')].split(',')[0:num_of_stocks]   #list indexed starts with meta data object, so add 1 to index because index[0] is removed
    print(':Top {} stocks: {}.'.format(num_of_stocks, stock_list))


def getStockData(stock):
    global stock_list
    if afterHours() is True and check_close is True:
        print(':The market is closed...')
        return False
    if stock == 'All':
        stock_list = stock_list
    elif type(stock) is list:
        stock_list = stock
    else:
        stock_list = []
        stock_list.append(stock)
    stockDict = {}
    for symbol in stock_list:
        if os.path.exists('{}.xlsx'.format(symbol)):
            os.remove('{}.xlsx'.format(symbol))
        time.sleep(3)
        with open('{}.xlsx'.format(symbol), 'w') as w:
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
        try:
            response = requests.get('https://www.alphavantage.co/query?', params=data)
            if 'Error Message' in response.text:
                print(':Removed {} from list.'.format(symbol))
                stock_list.remove(stock)
                pass
        except:
            stock_list.remove(stock)
            pass
        try:
            dummy = response.json()['Time Series (1min)']
        except Exception as exc:
            print(':While getting stock data for {}, got error from AlphaVantage... Waiting 15 seconds...'.format(symbol))
            with open('error.txt', 'a') as a:
                a.write('Error: {}\njson data: {}'.format(exc, response.json))
            time.sleep(15)
            getStockData(stock_list[stock_list.index(symbol): len(stock_list)])
            break
        data = response.json()['Time Series (1min)']
        # transfers data to excel
        stock_finances = []
        workbook = xlsxwriter.Workbook('{}.xlsx'.format(symbol))
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
            worksheet.write(i , 7 , '=(B{}+E{})/2'.format(i+1, i+1) , num_format)
            i += 1
        worksheet.write(maxRow - 2 , 8 , 'Avg(Total)')
        worksheet.write(maxRow - 1 , 8 , '=AVERAGE(H2:H{})'.format(maxRow) , num_format)
        # total percent change
        worksheet.write(maxRow - 2 , 10 , 'Total %change')
        worksheet.write(maxRow - 1 , 10 , '=((E2-B101)/B101)*100', num_format)
        workbook.close()
        # gets formula results to compile a dict
        wb = xw.Book('{}.xlsx'.format(symbol))
        percentChange = wb.sheets['stock data'].range('K{}'.format(maxRow)).value
        totalAvg = wb.sheets['stock data'].range('I{}'.format(maxRow)).value
        wb.app.quit()
        stockDict[symbol] = '[{}, {}]'.format(percentChange, totalAvg)
        with open("stockData.json" , "a") as append:
            json.dump(stockDict , append)
        # displays formula results
        if percentChange != None:
            print(':{} total change is {}%'.format(symbol, str(percentChange)[0:str(percentChange).find(".") + 2]))
        else:
            print(':{} encountered an error during parsing... Removing from list.'.format(symbol))
            if os.path.exists('{}.xlsx'):
                os.remove('{}.xlsx'.format(symbol))
            stock_list.remove(symbol)


def neuralNetPredition(symbol):
    print(':Starting neural network training for {}...'.format(symbol))
    wb = xw.Book('{}.xlsx'.format(symbol))
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
    row = 3 #default is 3; you can increase for testing
    list = wb.sheets['stock data'].range('B{}:E{}'.format(row, row)).value
    cellRange = range(row , maxRow)
    # creates x and y data sets
    for row in cellRange:
        y = wb.sheets['stock data'].range('E{}'.format(row - 1)).value
        values = wb.sheets['stock data'].range('B{}:E{}'.format(row, row)).value
        x_list.append(values)
        y_list.append(y)
    # runs neural network model
    row_to_predict = row-1   #2
    cells = 'B{}:E{}'.format(row_to_predict, row_to_predict)
    clf = MLPRegressor(solver='lbfgs' , alpha=1e-5 , random_state=1)
    clf.fit(x_list , y_list)
    guess = clf.predict([wb.sheets['stock data'].range(cells).value])
    print(':Model prediciton for {}: {}'.format(symbol, guess))
    with open('log.txt' , 'a') as aF:
        aF.write(":{}'s machine learning data: guess={};\npredicitonX_value={};\nmaxRows={};\n\n".format(symbol, guess, wb.sheets['stock data'].range(cells).value), maxRow)
    prevClose = wb.sheets['stock data'].range('E{}'.format(row_to_predict)).value
    percentChange = ((guess-prevClose)/prevClose)*100
    wb.app.quit()
    resultList = [float(guess), float(percentChange), prevClose]
    return resultList


def activeTrader(symbol):
    global stock_list
    global prediciton
    print('[activeTrader]: Initializing stock {}'.format(symbol))
    investment = False
    investPrice = None
    while True:
        if getStockData([symbol]) is False and check_close is True:
            print(':The market is closed...')
            exit()
        prediction = neuralNetPredition(symbol)
        cur_price = si.get_live_price(symbol)
        #should i invest
        if investment is False:
            if prediction[1] > 0.021:      # 0.021
                print('[activeTrader]: Investing in {}; estimated gain is {}%; entry price is {}.'.format(symbol, prediction[1], cur_price))
                with open('trade_log.txt', 'a') as w:
                    w.write('Invested at {}. Predicitons: {};\n'.format(cur_price, prediciton))
                investPrice = cur_price
                investment = True
            else:
                print('[activeTrader]: Not investing in {}; estimated change is {}%; current price is {}.'.format(symbol, prediction[1], cur_price))
        #should i sell?
        elif investment is True:
            if prediction[1] < 0.0115:       #0.0115
                print('[activeTrader]: Selling {}; estimated change is {}%; exit price is {}.'.format(symbol, prediction[1], cur_price))
                with open('trade_log.txt', 'a') as w:
                    w.write('Sold at {}. Original investment: {};\n'.format(cur_price, investPrice))
                if notify is True:
                    alert('{}: Sold at {}/original investment: {}; Predicitons: {};'.format(symbol, cur_price, investPrice, prediction))
                investment = False
            else:
                print('[activeTrader]: Keeping investment; estimated gain for {} is {}%; current price is {};'.format(symbol, prediction[1], cur_price))


def main():
    action = input('1. Get top {} stocks\n2. Get stock data\n3. Train and run neural network prediction\n4. Run autonomous trade bot\n? '.format(num_of_stocks))
    if action.lower() == 'clean':
        clean()
    if action == '1':
        getTopStocks()
    elif action == '2':
        getStockData('All')
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
            print('[Debug]: {} > {}'.format(resultSet[2], preferredStock[1]))
            if resultSet[2] > preferredStock[1]:
                preferredStock[0] = resultSet[0]
                preferredStock[1] = resultSet[2]
        if preferredStock[0] != 'None':
            print(':Neural Network preferred stock: {}; Estimated Gain: {}'.format(preferredStock[0]), preferredStock[1])
            print(': Starting Active Trader with preferred stock {}...'.format([preferredStock[0]]))
            stock_list.remove(preferredStock[0])
            clean()
            activeTrader(preferredStock[0])
        else:
            print(':Neural Network did not determine preferred stock... Run again.')
    main()



if __name__ == '__main__':
    main()

