# This repository is a work in progress!
I'm testing the program with paper trades right now trying to improve the machine learning and data limits within the code. I provided some test results at the bottom of this readme in the tests section. I'm hoping to have the neural network live trading out of the box. I'm thinking it could be finished by the end of January 2020. Thanks!
# marketTrader v1.0
A python script that will: gather and pull data from top gaining stocks, calculate which stock has the highest chance of profit, and using machine learning to buy and sell throughout a typical market day. (WIP)

This is not a module. Clone this repository to run it.
## Dependencies/Requirements
Install all required modules with
```pip install -r requirements.txt```

Required modules are:
- beautifulsoup4==4.8.1
- holidays==0.9.12
- sklearn==0.0
- PyAutoGUI==0.9.38
- pytz==2018.7
- XlsxWriter==1.1.8
- xlwings==0.16.4
- yahoo-fin==0.8.4

The program does use xlwings, which opens Excel files(.xlsx) while they are being read. If you don't have Microsoft Excel installed or your subscription has expired, **the program may not work**!
## AlphaVantage
This repository uses AlphaVantage's API to request data lists on stocks. All you need for this program to work is a free API key from their site, [(https://www.alphavantage.co/)].
# Use
The marketTrader.py file is standalone; it only requires a config.py file with your AlphaVantage API key. I do recommend putting it in a folder because it will produce .xlsx files and .json logs.
## Setup
Before use, you need to create a *config.py* file. This is what your *config.py* file should look like:
```
API_KEY = '01234567890'
```
Just replace the string with your API token(which you can get from [(https://www.alphavantage.co/)]) and put the *config.py* file in the **same directory** as *marketTrader.py*. *Note: Make sure you select the free plan! It is all you need to run this program.*
## Startup
When the program is opened, the menu below is displayed:
```
1. Get top 5 stocks
2. Get stock data
3. Train and run neural network prediction
4. Run autonomous trade bot
```
Type the corresponding number selection to start the program. What each option does is explained below:
### 1. Get Top 5 Stocks:
This option uses Yahoo! Finance to get the top 10 stocks from the top gainers page(10 is changable). It puts them into a list(*stock_list*). This list is used throughout the rest of the program. *Note: some options, like option 4, will run this automatically.*
### 2. Get Stock Data:
This option sends a request to AlphaVantage, an API service, to get data for each stock in the stock list. 
Stock data recieved:
- timestamp/date
- open
- high
- low
- close
- volume(not used currently)
The stock data is then organized and put into its own Excel file. All Excel files are named uniformly({symbol}.xlsx).
### 3. Train and run neural network prediction:
This option implements machine learning to predict the future price of a stock based on gathered data. The neural network used is *MLPRegressor* from the module *scikit-learn*. It outputs a list which includes estimated future open and percent change(previous close to predicted open).
### 4. Run autonomous trade bot
This option utilizes the entire program and loops specific functions to automate trading on the stock market. It will even tell you when the market has closed!

*Note: This program is still in development. Until I finish testing and debugging it, I won't implement the real stock trading. 
it is currently only using market data to paper trade.*
# How does marketTrader work?
This section is specifically detailing the automated trade bot. 

The bot starts by using Yahoo! Finance to get the top 10 stocks from the top gainers page. It then gathers stock data on each of the stocks in the list. It runs all of the stocks through the neural network and uses its predictions to determine which stock has the most potential at that time(it uses this stock for the rest of the program despite changes in other stock's data). It then will create a loop that: gathers stock data, runs neural network, and calculates if it should invest or sell. It uses the predicitons percent change to determine if it should invest or sell. Currently, if the percentage is greater than positive 2.1% predicted change it will invest; if the percentage is less than positive 1.15% predicted gains, it will sell. The program will not reinvest if it already has made an investment until it sells that investment. Likewise, it will not sell an investment that hasn't been made yet.
# Tests
This all seems to work in theory, but does it actually work?

Below is a log file from the current round of testing. This log was over the time of about 45 minutes of runnning.
```
Invested at 16.1200008392334. Predicitons: [16.19724042913184, 0.16846276519379505, 16.17];
Sold at 16.149999618530273. Predicitons: [16.160008592736567, 5.317287479409647e-05, 16.16];
Invested at 16.149999618530273. Predicitons: [16.12506220171521, 0.031403236446703835, 16.12];
Sold at 16.139999389648438. Predicitons: [16.126069142303155, -0.024369855529099562, 16.13];
Invested at 16.149999618530273. Predicitons: [16.227193222959606, 0.10606553337202489, 16.21];
Sold at 16.190000534057617. Predicitons: [16.168246129109814, -0.010846449537337132, 16.17];
```
*Note: In the predicitons list, the first number is the predicted next open, percent change, and the last value is the current price.*

So far, the model and calculations are doing reasonably well; they are turning a profit and are doing okay at predicitong trends. I'll keep you posted