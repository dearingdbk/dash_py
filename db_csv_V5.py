
'''
You Need to install 4-5 programs before you are able to run this script.
1. Python 3.4 or higher
2. PIP - if using python 3.4 pip should already be installed you can verify by typing "pip -V" into cmd prompt.
3. Pandas - goto cmd prompt type in: "pip install pandas" without quotes.
4. SQlite3 - necessary to create and operate on tabular data, and perform SQL like queries on data.
5. openpyxl - goto cmd prompt type in: "pip install openpyxl" without quotes.
6. Sublime Text - optional but makes editing and running scripts easier and may alleviate errors with file permissions.
'''

import sqlite3
from datetime import datetime
from dateutil.parser import parse
import pandas as pd
from pandas import DataFrame
import tkinter as tk
from tkinter import simpledialog
import re
from openpyxl import load_workbook
#import sys
#from xlutils.copy import copy
#from xlrd import open_workbook


USER_INP = ""


conn = None;
try:
    conn = sqlite3.connect(':memory:') # This allows the database to run in RAM, with no requirement to create a file.
    #conn = sqlite3.connect('dash_delivers.db')  # You can create a new database by changing the name within the quotes.
    print(sqlite3.version)
except Error as e:
    print(e)



c = conn.cursor() # The database will be saved in the location where your 'py' file is saved

# Create table - DRIVERS from dash_output.csv
c.execute('''CREATE TABLE IF NOT EXISTS DRIVERS
             ([generated_id] INTEGER PRIMARY KEY, [Source.Name] text, [Order ID] text, [Type] text, [Outcome] text, [First Name] text,
              [Last Name] text, [Address] text, [Subtotal] DECIMAL(13,2), [Delivery fee] DECIMAL(13,2), [Tip amount - gross] DECIMAL(13,2), [Total taxes] DECIMAL(13,2), 
              [Total] DECIMAL(13,2), [Payment Method] text, [Fulfillment date (YYYY-MM-DD)] date, [Fulfillment time] text, [Confirmed date (YYYY-MM-DD)] text,
              [Confirmed time] text, [Placed date (YYYY-MM-DD)] text, [Placed time (hh:mm)] text )''')

c.execute('''CREATE TABLE IF NOT EXISTS DASH_DATA
             ([generated_id] INTEGER PRIMARY KEY, [restaurant_name] text, [sales_total] DECIMAL(13,2), [pickup_total] DECIMAL(13,2), [delivery_debit_total] DECIMAL(13,2),
              [delivery_cash_total] DECIMAL(13,2), [delivery_fee_debit] DECIMAL(13,2), [delivery_fee_cash] DECIMAL(13,2))''')             
conn.commit()

read_drivers = pd.read_csv (r'dash_output.csv')
read_drivers.to_sql('DRIVERS', conn, if_exists='replace', index = False) # Insert the values from the csv file into the table 'DRIVERS'


# When reading the csv:
# - Place 'r' before the path string to read any special characters, such as '\'
# - Don't forget to put the file name at the end of the path + '.csv'
# - Before running the code, make sure that the column names in the CSV files match with the column names in the tables created and in the query below
# - If needed make sure that all the columns are in a TEXT format



def export_to_sheets():
	# set file path
	#filepath="/home/ubuntu/demo.xlsx"
	# load demo.xlsx 
	wb=load_workbook('template_settlement.xlsx')
	# get Sheet
	source=wb['Sheet1']
	# copy sheet
	#target=wb.copy_worksheet(source)
	# save workbook
	wb.save('new_some_document.xlsx')
	# done
	return

#######################################################################################
# get_date_range():                                                                   #
# Function to prompt user for required date range to append it to output files as     #
# required.                                                                           #
#######################################################################################
def get_date_range():
	ROOT = tk.Tk()
	ROOT.withdraw()
	global USER_INP
	USER_INP = simpledialog.askstring(title="Date Range",
                                  	  prompt="Input the date range to append to the end of each file name:\"Aug 12 - Aug 21 2019\"")
	USER_INP = re.sub('[^A-Za-z0-9\\-]+', '', USER_INP)
	return

#######################################################################################
# total_sales(name):                                                                  #
# Function to pull required data from created SQL database, and store it in the newly #
# created table.                                                                      #
#######################################################################################
def total_sales(name):
   # Pull total With Taxes
   c.execute('''
	SELECT SUM(DRIVERS.[Total]) 
	FROM DRIVERS 
	WHERE DRIVERS.[Outcome] == \'accepted\' AND DRIVERS .[Source.Name] == (?)
	GROUP BY DRIVERS.[Source.Name]''', (name,))
   fred = c.fetchall()
   if fred:
   		for item in fred[0]:
   			total = float(item)
   else:
    	total = 0.0
   # END Pull total
   # Pull pickup total with taxes 
   c.execute('''
	SELECT SUM(DRIVERS.[Total]) 
	FROM DRIVERS 
	WHERE DRIVERS.[Outcome] == \'accepted\' AND DRIVERS .[Source.Name] == (?) AND DRIVERS.[Type] == \'pickup\'
	GROUP BY DRIVERS.[Source.Name]''', (name,))
   fred = c.fetchall()
   if fred:
   		for item in fred[0]:
   			pickup_total = float(item)
   else:
    	pickup_total = 0.0
    # END Pull Pickup Total

    # Pull delivery debit total with taxes
   c.execute('''
	SELECT SUM(DRIVERS.[Total]) 
	FROM DRIVERS 
	WHERE DRIVERS.[Outcome] == \'accepted\' AND DRIVERS .[Source.Name] == (?) AND DRIVERS.[Type] == \'delivery\' AND DRIVERS.[Payment Method] == \'CARD\'
	GROUP BY DRIVERS.[Source.Name]''', (name,))
   fred = c.fetchall()
   if fred:
   		for item in fred[0]:
   			delivery_debit_total = float(item)
   else:
    	delivery_debit_total = 0.0    
    # END pull delivery debit total

    # Pull delivery cash total with taxes
   c.execute('''
	SELECT SUM(DRIVERS.[Total]) 
	FROM DRIVERS 
	WHERE DRIVERS.[Outcome] == \'accepted\' AND DRIVERS .[Source.Name] == (?) AND DRIVERS.[Type] == \'delivery\' AND DRIVERS.[Payment Method] == \'CASH\'
	GROUP BY DRIVERS.[Source.Name]''', (name,))
   fred = c.fetchall()
   if fred:
   		for item in fred[0]:
   			delivery_cash_total = float(item)
   else:
    	delivery_cash_total = 0.0
    # END pull delivery cash total
    
    # Pull delivery fee total (debit)
   c.execute('''
	SELECT SUM(DRIVERS.[Delivery fee]) 
	FROM DRIVERS 
	WHERE DRIVERS.[Outcome] == \'accepted\' AND DRIVERS .[Source.Name] == (?) AND DRIVERS.[Type] == \'delivery\' AND DRIVERS.[Payment Method] == \'CARD\'
	GROUP BY DRIVERS.[Source.Name]''', (name,))
   fred = c.fetchall()
   if fred:
   		for item in fred[0]:
   			delivery_fee_debit = float(item)
   else:
    	delivery_fee_debit = 0.0
    # END pull delivery fee total

   # Pull delivery fee total (cash)
   c.execute('''
	SELECT SUM(DRIVERS.[Delivery fee]) 
	FROM DRIVERS 
	WHERE DRIVERS.[Outcome] == \'accepted\' AND DRIVERS .[Source.Name] == (?) AND DRIVERS.[Type] == \'delivery\' AND DRIVERS.[Payment Method] == \'CASH\'
	GROUP BY DRIVERS.[Source.Name]''', (name,))
   fred = c.fetchall()
   if fred:
   		for item in fred[0]:
   			delivery_fee_cash = float(item)
   else:
    	delivery_fee_cash = 0.0
    # END pull delivery fee total (cash)



   c.execute('''INSERT INTO DASH_DATA (restaurant_name,sales_total, pickup_total, delivery_debit_total, delivery_cash_total, delivery_fee_debit, delivery_fee_cash) VALUES ((?),(?),(?),(?),(?),(?),(?))''',
    (name, total, pickup_total, delivery_debit_total, delivery_cash_total, delivery_fee_debit, delivery_fee_cash))
   
   return
   ###############################################################
   # End of Function to pull out data                            #
   ###############################################################
 

# Select all Restaurants in the .csv
c.execute('''
	SELECT DRIVERS.[Source.Name]
	FROM DRIVERS
	GROUP BY DRIVERS.[Source.Name]
		 ''')

restaurant_run = c.fetchall()

for restaurant in restaurant_run:
	total_sales(restaurant[0])


c.execute('''
	SELECT [restaurant_name], [sales_total], [pickup_total], [delivery_debit_total], [delivery_cash_total], [delivery_fee_debit], [delivery_fee_cash]
	FROM DASH_DATA 
		 ''')

#df = DataFrame(c.fetchall())
df = DataFrame(c.fetchall(), columns=['Source.Name', 'Subtotal', 'Pickup Total', 'Delivery Total (Debit)', 'Delivery Total (Cash)', 'Delivery Fee (Debit)', 'Delivery Fee (Cash)'])
print (df) 


#get_date_range()
export_to_sheets()
#df.to_sql('DRIVERS', conn, if_exists='append', index = False) # Insert the values from the INSERT QUERY into the table 'DAILY_STATUS'

try:
	export_csv = df.to_csv (r'export_list.csv', index = None, header=True) # Uncomment this syntax if you wish to export the results to CSV. Make sure to adjust the path name
except PermissionError:
	print("export_list.csv is open, cannot save output.")
# Don't forget to add '.csv' at the end of the path (as well as r at the beg to address special characters)


#export_to_sheets()

c.execute('''
	DROP TABLE IF EXISTS DRIVERS
		''')
c.execute('''
	DROP TABLE IF EXISTS DASH_DATA
		''')


conn.close()
