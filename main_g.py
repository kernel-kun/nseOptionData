#####################################################
#       Execute this before Running the Script:     #
#       > pip install -r requirements.txt           #
#####################################################


# Importing required Libraries/Modules
from http import server
from nsepython import *         # Includes pandas, scipy, datetime, json, math, os, requests, logging by default
import gspread
import os

# Configure Logging Info & Warn/Error to console
logging.basicConfig(format='%(asctime)s - %(message)s', datefmt='%d-%b-%y %H:%M:%S', level=logging.INFO)

# Suppress/Ignore pandas 'frame.append' warnings
import warnings
warnings.filterwarnings("ignore", message="The frame.append method is deprecated")

def printSep(noOfNewLines: int = 1) -> None:
    print("-" * 80, end='')
    print("\n" * noOfNewLines, end='')


startTime = time.time()
printSep()
logging.info("Script started")

##############
#   CHECKS   #
##############

# Check if nsepython works
# print(indices)

try:
    # Check if IP isn't blocked by NSE
    headers = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:84.0) Gecko/20100101 Firefox/84.0'}
    main_url = "https://www.nseindia.com/"
    response = requests.get(main_url, headers=headers, timeout=10)
    if response.status_code == 200:
        ##############################
        #       NSEPython Logic:     #
        ##############################

        #------------- FETCHING LATEST DATA -------------#
        #--- FROM "https://www.nseindia.com/option-chain" ---#

        # oi_data    : pandas dataframe
        # ltp        : last traded price
        # crontime   : Data’s Updation Time as per NSE Server when fetched
        oi_data, ltp, crontime = oi_chain_builder("NIFTY","latest","full")

        # print(ltp)

        # Log Script run time and NSE Update time to console
        logging.info(f"Current Date&Time of Scraped data is: {crontime}")

        #----------- Basic Functions -----------#
        # 1. Get No. of Row and Cols in oi_data
        # print(oi_data.shape)
        # 2. Get list of all col names in oi_data
        # print(oi_data.columns)

        # 3. Filer out the cols that we want
        oi_data_filtered = oi_data[["CALLS_Bid Price", "CALLS_Ask Price", "Strike Price", "PUTS_Bid Price", "PUTS_Ask Price"]]

        # 4. Export dataframe to xlsx file
        # oi_data_filtered.to_excel("test.xlsx")

        #---------- Print the complete dataframe ----------#
        # pd.set_option('display.max_columns', None)
        # pd.set_option('display.max_rows', None)
        # print(oi_data)
        # print(oi_data_filtered)


        ############################
        #       gspread Logic:     #
        ############################

        try:
            gs_credFile = json.loads( os.environ["CRED"] )
            gc = gspread.service_account_from_dict(gs_credFile)

            try:
                # Establish the connection
                wb = gc.open( os.environ["WB"] )      # wb  - workbook

                #----------- gspread Functions -----------#
                # 1. Print list all available Worksheets
                def listWorksheets():
                    print(wb.worksheets())

                # 2. Select a Worksheet
                wks = wb.worksheet( os.environ["WKS"] )     # wks - worksheet

                # 3. Get current Worksheet URL
                def currWorksheetURL():
                    print(wks.url)

                # 4. Print all values of Worksheet
                # print(wks.get_values())

                # 5. Completely clear worksheet
                wks.clear()

                # 6. Update worksheet values from dataframe
                wks.update([oi_data_filtered.columns.values.tolist()] + oi_data_filtered.values.tolist())
                wks.update('H1:H2', [['Last Update'], [crontime]])

            except:
                logging.error("Error while accessing Worksheet")

        except:
            logging.error("Error while accessing credentials from os.environ")

        # Checking if Google Sheet got updated successfully
        if wks.acell('H2').value == crontime:
            logging.info("Script ran successfully")
        else:
            logging.error("It seems Worksheet didn't get updated")

    else:
        logging.error(f"Error Code - {response.status_code}")

except:
    logging.error("IP Address is blocked, please use Proxy/VPN instead")

endTime = time.time()
print(f"Time taken to execute: {(endTime - startTime):.5f} sec")
printSep(2)
