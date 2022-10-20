from selenium import webdriver
from time import sleep
import pandas as pd
from bs4 import BeautifulSoup
import openpyxl

#Establish google chrome as the default web browser to use when calling driver
driver = webdriver.Chrome(r'C:\Users\1263654\Desktop\chromedriver.exe')

BondTickerList=pd.read_excel(r'F:\401K\1. 401K & ESOP\User - Karsten\401k-ESOP-Canada\Fund Performance Review\Scraping\bond_tickers2.xlsx',sheet_name='Tickers', engine='openpyxl')
csv_outputpath=r'F:\401K\1. 401K & ESOP\User - Karsten\401k-ESOP-Canada\Fund Performance Review\Scraping\2022.09.30'

#BondTickerList.head()
TickerList = BondTickerList.Ticker.unique()

#This is creating the dataframe and naming the columns outside of the loop so that information can be appended to it later.
DataCompile = []

for tick in TickerList:
        url_MS = r'https://www.morningstar.com/funds/xnas/'+(tick) +r'/quote'
        url_TD = r'https://research.tdameritrade.com/grid/public/mutualfunds/profile/performanceBuffer.asp?symbol='+(tick)

        # paths to metrics on the quote page
        NE = '//*[@id="__layout"]/div/div/div[2]/div[3]/div/main/div[2]/div/div/div/section[1]/div[1]/div[4]/div[2]/span/span'
        category = '//*[@id="__layout"]/div/div/div[2]/div[3]/div/main/div[2]/div/div/div/section[1]/div[1]/div[7]/div[2]/span'
        TTMyield = '//*[@id="__layout"]/div/div/div[2]/div[3]/div/main/div[2]/div/div/div/section[1]/div[1]/div[11]/div[2]/span'
        duration = '//*[@id="__layout"]/div/div/div[2]/div[3]/div/main/div[2]/div/div/div/section[1]/div[1]/div[12]/div[2]/span'

        #path to metrics on the risk page
        beta = '//*[@id="__layout"]/div/div/div[2]/div[3]/div/main/div[2]/div/div/div[1]/sal-components/div/sal-components-funds-risk/div/div/div/div/div/div[2]/div[1]/div[2]/div/div/div[2]/div/div[1]/table/tbody/tr[2]/td[2]/span'
        r2 = '//*[@id="__layout"]/div/div/div[2]/div[3]/div/main/div[2]/div/div/div[1]/sal-components/div/sal-components-funds-risk/div/div/div/div/div/div[2]/div[1]/div[2]/div/div/div[2]/div/div[1]/table/tbody/tr[3]/td[2]/span'
        sharpe = '//*[@id="__layout"]/div/div/div[2]/div[3]/div/main/div[2]/div/div/div[1]/sal-components/div/sal-components-funds-risk/div/div/div/div/div/div[2]/div[1]/div[2]/div/div/div[2]/div/div[1]/table/tbody/tr[4]/td[2]/span'
        upside = '//*[@id="__layout"]/div/div/div[2]/div[3]/div/main/div[2]/div/div/div[1]/sal-components/div/sal-components-funds-risk/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div[2]/div/div[1]/table[1]/tbody[1]/tr[1]/td[2]'
        downside = '//*[@id="__layout"]/div/div/div[2]/div[3]/div/main/div[2]/div/div/div[1]/sal-components/div/sal-components-funds-risk/div/div/div/div/div/div[2]/div[2]/div[2]/div/div/div[2]/div/div[1]/table[1]/tbody[1]/tr[2]/td[2]'

        driver.get(url_MS)
        net_expense = driver.find_element_by_xpath(NE).get_attribute('innerHTML')
        TickCategory = driver.find_element_by_xpath(category).get_attribute('innerHTML')
        TTMyield = driver.find_element_by_xpath(TTMyield).get_attribute('innerHTML')
        duration = driver.find_element_by_xpath(duration).get_attribute('innerHTML')

        #Click on the risk tab.
        driver.find_element_by_xpath('//*[@id="fund__tab-risk"]/a/span').click()
        sleep(3) # wait for 2 seconds

        # Now that you are on Risk page, grab beta, rsquare, sharpe, the upside, and downside for three years. 
        beta_3 = driver.find_element_by_xpath(beta).get_attribute('innerHTML') 
        r2_3 = driver.find_element_by_xpath(r2).get_attribute('innerHTML') 
        sharpe_3 = driver.find_element_by_xpath(sharpe).get_attribute('innerHTML') 
        upside_3 = driver.find_element_by_xpath(upside).get_attribute('innerHTML') 
        downside_3 = driver.find_element_by_xpath(downside).get_attribute('innerHTML') 

        #navigate to 5 year section of risk
        driver.find_element_by_xpath('//*[@id="for5Year"]').click() 
        sleep(3)
        r2_5 = driver.find_element_by_xpath(r2).get_attribute('innerhtml')
        sharpe_5 = driver.find_element_by_xpath(sharpe).get_attribute('innerhtml')
        upside_5 = driver.find_element_by_xpath(upside).get_attribute('innerhtml')
        downside_5 = driver.find_element_by_xpath(downside).get_attribute('innerhtml')

        driver.find_element_by_xpath('//*[@id="for10Year"]').click()
        sleep(4)
        r2_10 = driver.find_element_by_xpath(r2).get_attribute('innerHTML')
        sharpe_10 = driver.find_element_by_xpath(sharpe).get_attribute('innerHTML')
        upside_10 = driver.find_element_by_xpath(upside).get_attribute('innerHTML')
        downside_10 = driver.find_element_by_xpath(downside).get_attribute('innerHTML')

        driver.get(url_TD)
        sleep(4)
        return_1 = driver.find_element_by_xpath('//*[@id="table-trailingTotalReturnsTable"]/tbody/tr[4]/td[2]').get_attribute('innerHTML')
        return_3 = driver.find_element_by_xpath('//*[@id="table-trailingTotalReturnsTable"]/tbody/tr[5]/td[2]').get_attribute('innerHTML')
        return_5 = driver.find_element_by_xpath('//*[@id="table-trailingTotalReturnsTable"]/tbody/tr[6]/td[2]').get_attribute('innerHTML')
        return_10 = driver.find_element_by_xpath('//*[@id="table-trailingTotalReturnsTable"]/tbody/tr[7]/td[2]').get_attribute('innerHTML')

        # remove blank spaces from return numbers and clean up dirty sections from others.
        duration = ''.join(duration.split('<span class="mdc-data-point mdc-data-point--number" data-v-7ba8d775="" data-v-8645dbb6="">'))
        duration = ''.join(duration.split('<span class="mdc-data-point mdc-data-point--number" data-v-23f1d76c="" data-v-7eac76f0="">'))
        duration = ''.join(duration.split('</span>'))
        duration = ''.join(duration.split('\n\t\t\t\t\t\t\t\t\n\t\t\t\t\t\t\t'))
        duration = ''.join(duration.split('years'))
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

        metric_dict_data = {
        # Organize and name rows for data.
            'Ticker': tick,
            'Net Expense Ratio': net_expense,
            'Category': TickCategory,
            'TTM yield': TTMyield,
            'Effective duration': duration,
            '3 year Beta': beta_3,
            '3 Year Rsquare': r2_3,
            '3 Year Sharpe Ratio': sharpe_3,
            '3 Year Upside': upside_3,
            '3 Year downside': downside_3,
            '1 Year return': return_1,
            '3 Year return': return_3,

            ## 5 year
            '5 Year Rsquare': r2_5,
            '5 Year Sharpe Ratio': sharpe_5,
            '5 Year Upside': upside_5,
            '5 Year downside': downside_5,
            '5 Year return': return_5,

            ## 10 year
            '10 Year Rsquare': r2_10,
            '10 Year Sharpe Ratio': sharpe_10,
            '10 Year Upside': upside_10,
            '10 Year downside': downside_10,
            '10 Year return': return_10
            }

        #Append to the overarching dataframe outside of this loop.
        DataCompile.append(metric_dict_data)

#print (DataCompile)

Data_dict = dict(zip(TickerList[0:250], DataCompile))
Data_df = pd.DataFrame.from_dict(Data_dict)
Data_df.to_csv(csv_outputpath + '\\BondsDataFile.csv')

#metric_dict.to_csv(r'C:\Users\1263654\Desktop\BondTrial.csv')
