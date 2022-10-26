from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from time import sleep
import pandas as pd
from bs4 import BeautifulSoup
import openpyxl

#Enable options to make chromedriver work in the background. Not be in the front.
#chrome_options = Options()
#chrome_options.add_argument("--headless")

#Establish google chrome as the default web browser to use when calling driver
driver = webdriver.Chrome(r'C:\Users\1263654\Desktop\chromedriver.exe')#,options=chrome_options)

IndexTickerList=pd.read_excel(r'F:\401K\1. 401K & ESOP\User - Karsten\401k-ESOP-Canada\Fund Performance Review\Scraping\Ticker lists\2022.10.26 Tickers.xlsx',sheet_name='IndexNTarget', engine='openpyxl')
csv_outputpath=r'F:\401K\1. 401K & ESOP\User - Karsten\401k-ESOP-Canada\Fund Performance Review\Scraping'

#BondTickerList.head()
TickerList = IndexTickerList.Ticker.unique()

#This is creating the dataframe and naming the columns outside of the loop so that information can be appended to it later.
DataCompile = []

for tick in TickerList:
    url_MS = r'https://www.morningstar.com/funds/xnas/'+(tick) +r'/quote'
    url_TD = r'https://research.tdameritrade.com/grid/public/mutualfunds/profile/performanceBuffer.asp?symbol='+(tick)

    # paths to metrics on the quote page
    NE = '//*[@id="__layout"]/div/div/div[2]/div[3]/div/main/div[2]/div/div/div/section[1]/div[1]/div[4]/div[2]/span/span'

    #path to metrics on the risk page
    r2 = '//*[@id="__layout"]/div/div/div[2]/div[3]/div/main/div[2]/div/div/div[1]/sal-components/div/sal-components-funds-risk/div/div/div/div/div/div[2]/div[1]/div[2]/div/div/div[2]/div/div[1]/table/tbody/tr[3]/td[2]/span'
    sharpe = '//*[@id="__layout"]/div/div/div[2]/div[3]/div/main/div[2]/div/div/div[1]/sal-components/div/sal-components-funds-risk/div/div/div/div/div/div[2]/div[1]/div[2]/div/div/div[2]/div/div[1]/table/tbody/tr[4]/td[2]/span'
    stdev ='//*[@id="__layout"]/div/div/div[2]/div[3]/div/main/div[2]/div/div/div[1]/sal-components/div/sal-components-funds-risk/div/div/div/div/div/div[2]/div[1]/div[2]/div/div/div[2]/div/div[1]/table/tbody/tr[5]/td[2]/span'
    MS_date ='//*[@id="__layout"]/div/div/div[2]/div[3]/div/main/div[2]/div/div/div[1]/sal-components/div/sal-components-funds-risk/div/div/div/div/div/div[2]/div[1]/div[2]/div/div/div[2]/div/div[2]/span[2]/span[3]'

     #Create loop to try ticker/fund 3 times. Failures should print ticker and attempt name. 
     #   If failing 3 times on the same ticker, proceed to next. If successfully able to grab all information, proceed to next ticker. 
    attempts = 0
    while attempts < 3:
        attempts +=1
        #Everything inside of this try is to pull information on one ticker. If succesful, loop will break and proceed to the next ticker. 
        try:
            driver.get(url_MS)
            net_expense = driver.find_element_by_xpath(NE).get_attribute('innerHTML')

            #Click on the risk tab.
            driver.find_element_by_xpath('//*[@id="fund__tab-risk"]/a/span').click()
            sleep(4) # wait for 2 seconds

            # Now that you are on Risk page, grab beta, rsquare, sharpe, the upside, and downside for three years. 
            r2_3 = driver.find_element_by_xpath(r2).get_attribute('innerHTML') 
            sharpe_3 = driver.find_element_by_xpath(sharpe).get_attribute('innerHTML') 
            stdev_3 = driver.find_element_by_xpath(stdev).get_attribute('innerHTML') 
            MS_date_3 = driver.find_element_by_xpath(MS_date).get_attribute('innerHTML') 

            #navigate to 5 year section of risk
            driver.find_element_by_xpath('//*[@id="for5Year"]').click() 
            sleep(3)
            r2_5 = driver.find_element_by_xpath(r2).get_attribute('innerHTML')
            sharpe_5 = driver.find_element_by_xpath(sharpe).get_attribute('innerHTML')
            stdev_5 = driver.find_element_by_xpath(stdev).get_attribute('innerHTML') 

            driver.find_element_by_xpath('//*[@id="for10Year"]').click()
            sleep(4)
            r2_10 = driver.find_element_by_xpath(r2).get_attribute('innerHTML')
            sharpe_10 = driver.find_element_by_xpath(sharpe).get_attribute('innerHTML')
            stdev_10 = driver.find_element_by_xpath(stdev).get_attribute('innerHTML') 

            driver.get(url_TD)
            sleep(4)
            return_1 = driver.find_element_by_xpath('//*[@id="table-trailingTotalReturnsTable"]/tbody/tr[4]/td[2]').get_attribute('innerHTML')
            return_3 = driver.find_element_by_xpath('//*[@id="table-trailingTotalReturnsTable"]/tbody/tr[5]/td[2]').get_attribute('innerHTML')
            return_5 = driver.find_element_by_xpath('//*[@id="table-trailingTotalReturnsTable"]/tbody/tr[6]/td[2]').get_attribute('innerHTML')
            return_10 = driver.find_element_by_xpath('//*[@id="table-trailingTotalReturnsTable"]/tbody/tr[7]/td[2]').get_attribute('innerHTML')

            # remove blank spaces from return numbers and clean up dirty sections from others.
            return_1 = ''.join(return_1.split('<span class="positive">'))
            return_1 = ''.join(return_1.split('%</span>'))
            return_1 = ''.join(return_1.split('<span class="negative">'))
            return_3 = ''.join(return_3.split('<span class="positive">'))
            return_3 = ''.join(return_3.split('<span class="negative">'))
            return_3 = ''.join(return_3.split('%</span>'))
            return_5 = ''.join(return_5.split('<span class="positive">'))
            return_5 = ''.join(return_5.split('%</span>'))
            return_5 = ''.join(return_5.split('<span class="negative">'))
            return_10 = ''.join(return_10.split('<span class="positive">'))
            return_10 = ''.join(return_10.split('%</span>'))
            return_10 = ''.join(return_10.split('<span class="negative">'))
            
            # Organize and name rows for data.
            metric_dict_data = {
                'Ticker': tick,
                'Net Expense Ratio': net_expense,
                'MS Date data as of': MS_date_3,
                '3 Year Rsquare': r2_3,
                '3 Year Sharpe Ratio': sharpe_3,
                '3 Year StandardDeviation':stdev_3,
                '1 Year return': return_1,
                '3 Year return': return_3,

                ## 5 year
                '5 Year Rsquare': r2_5,
                '5 Year Sharpe Ratio': sharpe_5,
                '5 Year StandardDeviation':stdev_5,
                '5 Year return': return_5,

                ## 10 year
                '10 Year Rsquare': r2_10,
                '10 Year Sharpe Ratio': sharpe_10,
                '10 Year StandardDeviation':stdev_10,
                '10 Year return': return_10
                }

            #Append to the overarching dataframe outside of this loop.
            DataCompile.append(metric_dict_data)
            #If able to get through all points on the ticker and append, then break from this loop and move to next ticker
            break
        #If unable to grab information, print the ticker and which attempt this was. Then try again for this ticker.
        #If it has been 3 tries, it will proceed to the next ticker.
        except:
            print('failed' + tick + ' ' + str(attempts))

print ('finished pulling data')

#Convert the data into a dictionary and then into a dataframe so we can export to csv.
Data_dict = dict(zip(TickerList[0:250], DataCompile))
Data_df = pd.DataFrame.from_dict(Data_dict)

#Before spitting out data, it is currently horizontal. Transpose it so it is vertical (up and down)
Data_df_V = Data_df.transpose()
print(Data_df_V)

#Spit out data to csv
Data_df_V.to_csv(csv_outputpath + '\\IndexNTargetDataFile4.csv')
