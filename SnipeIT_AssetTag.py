# Import required libraries
import pandas as pd
import logging
import numpy
import datetime

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


## Function KeringAssetTagBySN: Used to generate kering asset ID based on region, brand, category, location, and SN.
def KeringAssetTagBySN(df, i):
    
    # Creating variables
    kering_asset_tag = "" # Kering asset tag
    KeringBrandForPy = "" # 2nd character of Kering asset tag
    KeringCategoryForPy = "" # 3rd character of Kering asset tag
    KeringLocForPy = "" # 4th character of Kering asset tag

    # Define the brand        
    if df['company.id'][i] == 4: #Kering
        KeringBrandForPy = "K"
    elif df['company.id'][i] == 6: #BV
        KeringBrandForPy = "V"
    elif df['company.id'][i] == 7: #KeringEyeware
        KeringBrandForPy = "E"
    elif df['company.id'][i] == 8: #YSL
        KeringBrandForPy = "S"
    elif df['company.id'][i] == 9: #GG
        KeringBrandForPy = "G"
    elif df['company.id'][i] == 10: #BAL
        KeringBrandForPy = "B"
    elif df['company.id'][i] == 11: #POM
        KeringBrandForPy = "P"
    elif df['company.id'][i] == 12: #BOU
        KeringBrandForPy = "O"
    elif df['company.id'][i] == 13: #AMQ
        KeringBrandForPy = "M"
    elif df['company.id'][i] == 14: #GucciTimepieces
        KeringBrandForPy = "G"
    elif df['company.id'][i] == 15: #QEE
        KeringBrandForPy = "Q"
    elif df['company.id'][i] == 17: #BRI
        KeringBrandForPy = "R"
    else:   # Log exception
        logging.error("Unexpected brand name in KeringAssetTagBySN")
        
    # Define the category
    if df['category.id'][i] == 10: #Laptop
        KeringCategoryForPy = "L"
    elif df['category.id'][i] == 11: #Desktop
        KeringCategoryForPy = "D"
    elif df['category.id'][i] == 12: #PCScanner
        KeringCategoryForPy = "S"
    elif df['category.id'][i] == 15: #Printer
        KeringCategoryForPy = "P"
    elif df['category.id'][i] == 16: #iPad
        KeringCategoryForPy = "A"
    elif df['category.id'][i] == 17: #iPod
        KeringCategoryForPy = "O"
    elif df['category.id'][i] == 18: #iPhone
        KeringCategoryForPy = "P"
    elif df['category.id'][i] == 21: #Mac
        KeringCategoryForPy = "M"
    elif df['category.id'][i] == 22: #Tablet
        KeringCategoryForPy = "T"
    else:
        logging.error("Unexpected device category in KeringAssetTagBySN")

    # Define the location
    if df['custom_fields.Kering Location.value'][i] == "办公室":
        KeringLocForPy = "O"
    elif df['custom_fields.Kering Location.value'][i] == "店铺":
        KeringLocForPy = "R"
    elif df['custom_fields.Kering Location.value'][i] == "仓库":
        KeringLocForPy = "W"
    elif df['custom_fields.Kering Location.value'][i] == "其他":
        KeringLocForPy = "T"
    else:
        logging.error("Unexpected location in KeringAssetTagBySN")

    # Retrieve serial number
    SnipeITAssetTag = df['asset_tag'][i] # Quote the Serial No in AMS, which is named as Asset Tag

    # Concat the kering asset ID
    kering_asset_tag = f"C{KeringBrandForPy}{KeringCategoryForPy}{KeringLocForPy}-{SnipeITAssetTag}"

    # Return the result
    return kering_asset_tag

## Function KeringAssetTagDedicated: Used to generate kering asset ID based on region, brand, and category.
def KeringAssetTagDedicated(df, i):
    
    # Creating variables
    KeringBrandForPy = ""
    KeringCountryForPy = ""
    KeringCategoryForPy = ""
    kering_asset_tag_prefix = ""
    
    # Define the brand        
    if df['company.id'][i] == 4: #Kering
        KeringBrandForPy = "P"
    elif df['company.id'][i] == 6: #BV
        KeringBrandForPy = "B"
    elif df['company.id'][i] == 7: #KeringEyeware
        KeringBrandForPy = "KEYE"
    elif df['company.id'][i] == 8: #YSL
        KeringBrandForPy = "Y"
    elif df['company.id'][i] == 9: #GG
        KeringBrandForPy = "G"
    elif df['company.id'][i] == 10 and df['category.id'][i] == 19: #BAL Monitor
        KeringBrandForPy = "BAL"
    elif df['company.id'][i] == 10 and df['category.id'][i] in [13,23]: #BAL LC&LP
        KeringBrandForPy = "BA"
    elif df['company.id'][i] == 11: #POM
        KeringBrandForPy = "PO"
    elif df['company.id'][i] == 12: #BOU
        KeringBrandForPy = "BOU"
    elif df['company.id'][i] == 13: #AMQ
        KeringBrandForPy = "A"
    elif df['company.id'][i] == 14: #GucciTimepieces
        KeringBrandForPy = "G"
    elif df['company.id'][i] == 15: #QEE
        KeringBrandForPy = "Q"
    elif df['company.id'][i] == 17: #BRI
        KeringBrandForPy = "BRI"
    elif df['company.id'][i] == 18: #LGI
        KeringBrandForPy = "L"
    else:   # Log exception
        logging.error("Unexpected brand in KeringAssetTagDedicated")

    # Define the country
    if df['company.id'][i] in [6,10,13]: #BV,BAL,AMQ
        KeringCountryForPy = "CN"
    elif df['company.id'][i] == 17 and df['category.id'][i] == 19: #BRI Monitor
        KeringCountryForPy = "CN"
    elif df['company.id'][i] == 17 and df['category.id'][i] == 19: #BRI LC&LP
        KeringCountryForPy = "N"
    else:   # Log exception
        logging.error("Unexpected country in KeringAssetTagDedicated")

        
    # Define the category
    if df['category.id'][i] == 13: #LineaPro
        KeringCategoryForPy = "LP"
    elif df['category.id'][i] == 19: #Monitor
        KeringCategoryForPy = "LCD"
    elif df['category.id'][i] == 23: #LineaPro Charger
        KeringCategoryForPy = "LC"
    else:
        logging.error("Unexpected device category in KeringAssetTagDedicated")
    
    # Combine the brand, country, and category
    kering_asset_tag_prefix = f"{KeringBrandForPy}{KeringCountryForPy}{KeringCategoryForPy}"

    # Retrieve the latest numeric value from
    filtered_df = df[df['custom_fields.Kering Asset Tag.value'].str.startswith(kering_asset_tag_prefix)].copy()  # Step 1: Filter rows that start with the given prefix
    filtered_df['numeric_part'] = filtered_df['custom_fields.Kering Asset Tag.value'].str.extract(r'(\d{4})$').astype(int) # Step 2: Extract the numeric part
    max_numeric_value = filtered_df['numeric_part'].max() # Step 3: Find the maximum numeric part
    next_numeric_value = max_numeric_value + 1 # Step 4: Increment the numeric part by 1

    # Define numeric values of the brand_category that have not been used
    if numpy.isnan(next_numeric_value) == True:
        if df['category.id'][i] == 19 and df['company.id'][i] == 12: # BRI monitor starts at 50
           next_numeric_value = '50'
        elif df['category.id'][i] == 19 and df['company.id'][i] ==  17: # BOU monitor starts at 90
            next_numeric_value = '90'
        elif df['category.id'][i] == 19 and df['company.id'][i] ==  4: # Kering's monitor starts with 420
            next_numeric_value = '420'
        elif df['category.id'][i] in [13,23] and df['company.id'][i] ==  10: # BAL's LP & LC starts with 220
            next_numeric_value = '220'
        else:
            logging.error("Unexpected nan in KeringAssetTagDedicated")

    kering_asset_tag = f"{kering_asset_tag_prefix}{str(next_numeric_value).zfill(4)}" # Step 5: Construct the new ID
    
    # Return the result
    return kering_asset_tag