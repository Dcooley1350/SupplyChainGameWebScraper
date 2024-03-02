import os
import json
import mechanize
from bs4 import BeautifulSoup
from http.cookiejar import CookieJar
import pandas
import datetime

# Get the current date and time
now = datetime.datetime.now()
# Print the current date and time in a specific format
print(now.strftime("%m-%d-%Y %H:%M:%S"))

# Set up the browser, authenticate
cj = CookieJar()
br = mechanize.Browser()
br.set_cookiejar(cj)
url = os.getenv('URL', "http://op.responsive.net/sc/ba727/entry.html")
br.open(url)
br.select_form(nr=0)
br.form['id'] = os.getenv('USERNAME')
br.form['password'] = os.getenv('PASSWORD')
br.submit()

STANDARD_URLS = {
    "Shipments": {
        "url": "https://op.responsive.net/SupplyChain/SCPlotk?submit=plot+shipments&data=SHIP1SEG1",
    },
    "Cash": {
        "url": "https://op.responsive.net/SupplyChain/SCPlotk?submit=plot+cash+balance&data=BALANCE",

    },
    "Demand": {
        "url": "https://op.responsive.net/SupplyChain/SCPlotk?submit=plot+demand&data=DEMAND1",

    },
    "Lost Demand": {
        "url": "https://op.responsive.net/SupplyChain/SCPlotk?submit=plot+lost+demand&data=LOST1",

    },
}

# Scrape nice, law abiding, upstanding citizen, standard style data
STANDARD_DATA = {}
STANDARD_HEADERS = STANDARD_URLS.keys()
for key, value in STANDARD_URLS.items():
    soup = BeautifulSoup(br.open(value['url']), "lxml")
    data = soup.find_all("script")[6].string
    data = data.split("\n")[4].split("'")[5].split()

    for i, v in enumerate(data):
        if i % 2 == 0:
            # if key doesn't exist, create it
            data_key = int(float(v))
            if not STANDARD_DATA.get(data_key):
                STANDARD_DATA[data_key] = []
            STANDARD_DATA[data_key].append(float(data[i+1]))

# Create a dataframe standard data
standard_df = pandas.DataFrame.from_dict(STANDARD_DATA, orient="index")
standard_df.columns = STANDARD_HEADERS
standard_df.sort_index(inplace=True)

# Scrape friggin weird Inventory 4 column table with duplicate days(keys) and empty values (blegh)
INVENTORY_DATA = {}
INVENTORY_HEADERS = ["Warehouse", "Mail", "Truck"]
soup = BeautifulSoup(br.open("https://op.responsive.net/SupplyChain/SCPlotk?submit=plot+inventory&data=WH1"), "lxml")
data = soup.find_all("script")[6].string
datasets = []
datasets.append(data.split("\n")[4].split("'")[5].split())
datasets.append(data.split("\n")[5].split("'")[5].split())
datasets.append(data.split("\n")[6].split("'")[5].split())

column_count = 0
for dataset in datasets:
    previous_key = None
    for i, v in enumerate(dataset):
        if i % 2 == 0:
            # format the key to 7 decimal places
            key = float(format(float(v), '.7f'))
            # If this key equals the previous, we round it up on digit so we have unique keys (hacky, but it works **shrug**)
            if key == previous_key:
                key += 0.00001
            # if key doesn't exist, create it
            if not INVENTORY_DATA.get(key):
                INVENTORY_DATA[key] = []
                # if the column count is less than what it should be, add fills
            while len(INVENTORY_DATA[key]) < column_count:
                INVENTORY_DATA[key].append(None)
            # set value to
            INVENTORY_DATA[key].append(float(dataset[i+1]))
        previous_key = key
    column_count += 1

# Backfill any rows that are not full
for key, value in INVENTORY_DATA.items():
    while len(value) < column_count:
        value.append(None)

# Create a dataframe from inventory data
inventory_df = pandas.DataFrame.from_dict(INVENTORY_DATA, orient="index")
inventory_df.columns = INVENTORY_HEADERS
inventory_df.sort_index(inplace=True)

# Scrape mildly weird WIP 2 column table with duplicate days(keys)(small blegh)
WIP_DATA = {}
WIP_HEADERS = ["WIP"]
soup = BeautifulSoup(br.open("https://op.responsive.net/SupplyChain/SCPlotk?submit=plot+wip&data=WIP1"), "lxml")
data = soup.find_all("script")[6].string
data = data.split("\n")[4].split("'")[5].split()

previous_key = None
for i, v in enumerate(data):
    if i % 2 == 0:
        # cast key to float and round it to 3 digits
        key = round(float(v), 3)
        # if key doesn't exist, create it
        if not WIP_DATA.get(key):
            WIP_DATA[key] = []
        else:
            key += 0.001
            WIP_DATA[key] = []
        # set value to
        WIP_DATA[key].append(float(data[i+1]))
    previous_key = key

# Create a dataframe from WIP data
wip_df = pandas.DataFrame.from_dict(WIP_DATA, orient="index")
wip_df.columns = WIP_HEADERS
wip_df.sort_index(inplace=True)

# Set up the Excel writer based on existence of the file
kwargs = {}
path = os.getenv('OUTPUT_FILE', 'output.xlsx')
if os.path.exists(path):
    kwargs['mode'] = 'a'
    kwargs['if_sheet_exists'] ='overlay'
else:
    kwargs['mode'] = 'w'
writer = pandas.ExcelWriter(path, **kwargs)

# Write dataframes to Excel sheets and save
standard_df.to_excel(writer, sheet_name='standard_data')
inventory_df.to_excel(writer, sheet_name='inventory_data')
wip_df.to_excel(writer, sheet_name='wip_data')

writer._save()
print(f"Data saved to Excel file: {path}")
