#! python3
#script to update excel sheet daily with mining profits from pools

#need to download bs4, openpyxl, selenium, and download/add geckodriver to path
import os, time, openpyxl, bs4, requests
from selenium import webdriver

#Pulls curent BTC price from coinbase.com
def btcScrape():
    print('Getting current BTC price')
    url = 'https://api.coinbase.com/v2/prices/BTC-USD/spot'
    #Requests .json data from coinbase
    res = requests.get(url)
    coinbaseData = res.json()
    #Takes BTC amount from response.json
    BTC = coinbaseData['data']['amount']
    return BTC

def duplicateCheck(rowNum, payDate, sheet):
    if (payDate == sheet.cell(row=(rowNum-1), column=1).value):
        replace = input('Current pay time matches last recorded at ' + str(payDate) + ', possible duplicate. \nWould you like to continue? (y/n)').lower()
        if (replace == 'y'):
            return True
        elif (replace == 'n'):
            return False
        else:
            print ('Invalid input, please try again')
            duplicateCheck(rowNum,payDate,sheet)

def profitUpdate(sheet, date, pool, paid, btcVal):
    rowNum = 6
    #checks if cell has value, and iterates until it finds empty row
    while True:
        if (sheet.cell(row=rowNum, column=3).value != None):
            rowNum += 1
            continue
        else:
            break
    while True:
        duplicateCheck(rowNum, date, sheet)
        #update BTC value
        sheet.cell(row=2, column=3).value = float(btcVal)
        #time paid was added
        sheet.cell(row=rowNum, column=1).value = date
        #Pool
        sheet.cell(row=rowNum, column=2).value = pool
        #total paid value
        sheet.cell(row=rowNum, column=3).value = float(paid)
        #price per coin
        sheet.cell(row=rowNum, column=4).value = float(btcVal)
        #coin since previous day
        sheet.cell(row=rowNum, column=5).value = '=$C' + str(rowNum) + '/($A' + str(rowNum) + '-$A' + str(rowNum - 1) + ')'
        #dollar per day
        sheet.cell(row=rowNum, column=6).value = '=$E' + str(rowNum) + '*$D' + str(rowNum)
        #BTC per rig per day
        sheet.cell(row=rowNum, column=7).value = '=$E' + str(rowNum) + '/$G$2'
        #dollar per rig per day
        sheet.cell(row=rowNum, column=8).value = '=$F' + str(rowNum) + '/$G$2'
        #save new additions
        wb.save("MiningProfit.xlsx")
        break

def zpoolScrape():
    #Sets pool value, gets BTC price, and starts count for sheet page
    currentPool = "zpool"
    btcPrice = btcScrape()
    count = 1
    #Iterates through address dictionary
    for cryptoAdd in address.items():
        url = 'https://www.zpool.ca/?address=' + cryptoAdd[1]
        #Uses headless firefox to get html data for date and paid amount
        driver.get(url)
        print('Getting 24hr profit data for ' + cryptoAdd[0])
        #stalls script to allow page to finish javascript for html scrape
        time.sleep(10)
        #html data from zpool
        data = driver.page_source
        #uncomment section if needing to look at html data
        # print(data)
        # htmlData = open('htmldata.txt', 'a')
        # htmlData.write(data)
        # htmlData.close()
        print('Analyzing zpool data')
        #retrieves payment data from zpool
        html = bs4.BeautifulSoup(data, 'html.parser')
        bList = html.select('#main_wallet_results tbody b')
        amountPaid  = bList[-1].getText()
        #checks paid amount and if not number, returns no payment
        try:
            float(amountPaid)
            #Retrieves date time of payment if payment occured
            bspanList = html.select('#main_wallet_results tbody b span')
            #trims seconds off date time to allow excel date formatting
            dateTime = bspanList[-1].get('title')[:-3]
            sheetTitle = wb['BTC' + str(count)]
            #checks that payment occured
            if (amountPaid != 0):
                print('Writing data to workbook')
                profitUpdate(sheetTitle, dateTime, currentPool, amountPaid, btcPrice)
        except:
            print('No payment recorded')
        #changes count to open next tab
        count += 1

#must change directory to location of excel workbook
os.chdir('D:\\')
#BTC address
address = {'Miner0-3':'3DfY3BE5w72x7AdrU8o3NvjEzqLvxq1As3', 'Miner4':'3G89tWUHn6VdnaBxoMRTd4WrokJnmVHGfx'}
print('Starting Firefox')
#Starts a headless firefox program
os.environ['MOZ_HEADLESS'] = '1'
driver = webdriver.Firefox(log_path = ".\\nul")
print('Starting Excel')
#opens workbook for data writing
wb = openpyxl.load_workbook("MiningProfit.xlsx")
zpoolScrape()
#closes headless firefox
driver.quit()
print('Complete')
input('Press enter to exit')
