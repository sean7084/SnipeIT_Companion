# Import required libraries
import requests
import pandas as pd
import json
import logging
import datetime
from SnipeIT_AssetTag import *
from api_key import *

# Date, time
now = datetime.datetime.now()
now_date = now.strftime("%Y-%m-%d")
now_time = now.strftime("%H-%M-%S")

# Logging module configuration
logging.basicConfig(
    filename=f"X:/Project/AMS_Asset_Export/Logs/app_{now_date}_{now_time}.log",
    encoding="utf-8",
    filemode="a",
    format="{asctime} - {levelname} - {message}",
    style="{",
    datefmt="%Y-%m-%d %H:%M",
    )

# Credential for connection
headers = {
    "accept": "application/json",
    "Authorization": f"Bearer {api_key}",
    "content-type": "application/json"
    }

# Request information from AMS
url = "http://192.168.31.153/api/v1/hardware"
response = requests.get(url, headers=headers)
response_str = response.text
#print(response_str)

# Convert response to the DataFrame
response_dict = json.loads(response_str)
df = []
df = pd.json_normalize(response_dict["rows"])

# Add the Kering Asset Number (Hengji and MCM are excluded)
"""
For the purpose of testing
snipe_it_asset_id = 2
kering_asset_tag = "abcd-12345678910"
"""

## Create var
id_list = []

## Upadating SnipeIT's database
for i in range(0,len(df['id'])):
    # If statement to fileter out the entries that have not assigned a asset tag and have a assigned location & brand
    if df['custom_fields.Kering Asset Tag.value'][i] == "" and df['custom_fields.Kering Location.field'][i] != "" \
        and df['company.id'][i] != "":
        
        # Creating the variable
        kering_asset_tag = ""

        # List to categorize brand that inside/outside of Kering, and brands that all/part of assets are tagged based on SN.
        Not_Kering_Brands = [1,5]
        Kering_Brands = [4,6,7,8,9,10,11,12,13,14,15,16,17,18]
        
        # List of categories that asset tag is related to SN or not
        Categories_w_SN = [10,11,12,15,16,17,18,21,22]
        Categories_wo_SN = [13,19,23]

        # Assign the tag
        if df['company.id'][i] in Not_Kering_Brands:
            continue
        elif df['company.id'][i] in Kering_Brands and df['category.id'][i] in Categories_w_SN:
            kering_asset_tag = KeringAssetTagBySN(df = df, i = i)
        elif df['company.id'][i] in Kering_Brands and df['category.id'][i] in Categories_wo_SN:
            kering_asset_tag = KeringAssetTagDedicated(df = df, i = i)
        else:
            logging.error("Unexpected company.id or category.id in SnipeIT")
    
        # Update on SnipeIT using api
        snipe_it_asset_id = df['id'][i] # Define the SnipeIT ID
        url = f"http://192.168.31.153/api/v1/hardware/{snipe_it_asset_id}"
        payload = {"_snipeit_kering_asset_tag_2":f"{kering_asset_tag}",
                   "status_id":4}
        response = requests.put(url, json=payload, headers=headers)

        # Store the SnipeIT ID in a list
        id_list.append(snipe_it_asset_id)
    
# Export asset list for record and label printing
## Request updated information from AMS
url = "http://192.168.31.153/api/v1/hardware"
response = requests.get(url, headers=headers)
response_str = response.text

## Convert updated response to the DataFrame
response_dict = json.loads(response_str)
df = []
df = pd.json_normalize(response_dict["rows"])

## Filter out the rows
df_modified_rows = df.loc[df['id'].isin(id_list)]
#print(df_modified_rows.columns)


df_sc = df_modified_rows[['company.name',
                        'assigned_to',
                        'custom_fields.Kering Location.value',
                        'purchase_cost',
                        'manufacturer.name',
                        'model.name',
                        'custom_fields.Kering Asset Tag.value',
                        'asset_tag',
                        'company.name']] ## Select columns
# FUTHER MOD: CURRENCY, TAX

## Export as "Asset List + datetime.xlsx"

datatoexcel = pd.ExcelWriter("X:/Project/AMS_Asset_Export/Asset List_"+now_date+"_"+now_time+".xlsx")
df_sc.to_excel(datatoexcel)
datatoexcel.close()

# Put script in a single executable file
# cd HengJi\AMS_Companion
# pyinstaller -F -n SnipeIT_Companion_v1.2 SnipeIT_Main.py