import time
import json
import os
from datetime import datetime
import pandas as pd
import sys
import xlsxwriter
import warnings
import re
import sys
import shutil
import traceback
import math
import hrequests as requests
warnings.filterwarnings('ignore')

def get_inputs():
 
    print('Processing The Settings Sheet ...')
    print('-'*75)
    # assuming the inputs to be in the same script directory
    path = os.path.join(os.getcwd(), 'settings.xlsx')

    if not os.path.isfile(path):
        print('Error: Missing the settings file "settings.xlsx"')
        input('Press any key to exit')
        sys.exit(1)
    try:
        settings = {}
        df = pd.read_excel(path)
        for col in df.columns:
            df[col] = df[col].astype(str)

        for ind in df.index:
            settings[df.iloc[ind, 0].lower().replace(' ', '-')] = df.iloc[ind, 1].lower()

    except Exception as err:
        print('Error: Failed to process the settings sheet')
        print(err)
        input('Press any key to exit')
        sys.exit(1)

    return settings

def initialize_outputs(settings):

    # removing the previous output file
    stamp = datetime.now().strftime("%d_%m_%Y_%H_%M")
    path = os.path.join(os.getcwd(), 'data', stamp)
    if os.path.exists(path):
        shutil.rmtree(path)
    os.makedirs(path)

    outputs = {}
    for key, value in settings.items():
        if value == 'yes':
            file = f'{key.title()}_{stamp}.csv'
            output = os.path.join(path, file)
            # workbook = xlsxwriter.Workbook(output)
            # workbook.add_worksheet()
            # workbook.close()
            outputs[key] = output

    return outputs

def scrape_prods(outputs, settings):
    for category, status in settings.items():
        if status != 'yes': continue
        name = category.replace("-", ' ').title()
        print(f'Scraping Category: {name}')
        print('-'*75)
        df = pd.DataFrame()
        url = f"https://portal.spray.com/en-us/categories/{category}?all=true"
        for _ in range(10):
            try:
                response = requests.get(url)
                time.sleep(0.5)
                if response.status_code == 200:
                    break
            except:
                print(f'Failed to load the category page: {url}')
                print(traceback.format_exc())
                time.sleep(10)

        nprods = int(re.findall(r'Found\s([\d,]+)\sproducts', response.text)[0].replace(',', ''))
        print(f"Number of products: {nprods}")
        #print('-'*75)
        npages = math.ceil(nprods/10)
        iprod = 0
        for page in range(1, npages+1):
            pageUrl = url + f"&page={page}"
            for _ in range(10):
                try:
                    response = requests.get(pageUrl)
                    time.sleep(0.5)
                    if response.status_code == 200:
                        searchData = json.loads(re.findall(
                        '<script id="__NEXT_DATA__" type="application/json">(.*?)</script>', response.text)[0])
                        break
                except:
                    print(f'Failed to load the search page: {pageUrl}')
                    print(traceback.format_exc())
                    time.sleep(10)
        
            prods = searchData["props"]["pageProps"]["finderData"]["facetedSearchProductViewModels"]
            for prod in prods:
                prodUrl = "https://portal.spray.com/en-us" + prod["product"]["url"]
                if prodUrl == "https://portal.spray.com/en-us/products/aab707-1-4-al-15":
                    debug = True
                iprod += 1
                print(f'Scraping Product {iprod}/{nprods}')
                for _ in range(10):
                    try:
                        response = requests.get(prodUrl)
                        time.sleep(1)
                        if response.status_code == 200:
                            prodData = json.loads(re.findall(
                            '<script id="__NEXT_DATA__" type="application/json">(.*?)</script>', response.text)[0])
                            break
                    except:
                        print(f'Failed to load the product page: {prodUrl}')
                        print(traceback.format_exc())
                        time.sleep(10)

                prodData = prodData["props"]["pageProps"]
                prodDetails = {}
                try:
                    prodDetails["ProductName"] = prodData["product"]["name"]["en"] + ', ' + prodData["product"]["number"]
                except:
                    pass

                try:
                    prodDetails["ProductLink"] = prodUrl
                except:
                    pass

                try:
                    prodDetails["ProductId"] = prodData["product"]["id"]
                except:
                    pass

                try:
                    prodDetails["ModelId"] = prodData["product"]["modelId"]
                except:
                    pass

                try:
                    prodDetails["ProductCode"] = prodData["product"]["number"]
                except:
                    pass

                try:
                    prodDetails["Audience"] = prodData["product"]["audience"]
                except:
                    pass

                try:
                    prodDetails["Description"] = prodData["product"]["description"]["en"].replace("*", "")
                except:
                    pass

                try:
                    n = 1
                    for resource in  prodData["product"]["resources"]:
                        if resource["type"] not in prodDetails:
                            prodDetails[resource["type"] + 'Link'] = resource["url"]
                        else:
                            n += 1
                            prodDetails[resource["type"] + str(n) + "Link"] = resource["url"]
                except:
                    pass

                try:
                    for option in  prodData["product"]["options"]:
                        optionName = option["typeCode"]
                        variants = option["variants"]
                        for variant in variants:
                            if variant["productNumber"] == prodDetails["ProductCode"]:
                                prodDetails[optionName] = variant["displays"][0]["value"]["en"]
                                break
                except:
                    pass

                try:
                    for attr in prodData["product"]["attributes"]:
                        try:
                            prodDetails[attr["typeCode"]] = attr["value"]
                            if attr["unitSymbol"]:
                                prodDetails[attr["typeCode"]] = str(prodDetails[attr["typeCode"]]) + ' ' + attr["unitSymbol"]

                            if "conditions" in attr and attr["conditions"]:
                                try:
                                    prodDetails[attr["typeCode"]] = str(prodDetails[attr["typeCode"]]) + " @ " + str(attr["conditions"][0]["value"]) + " " + attr["conditions"][0]["unitSymbol"]
                                except:
                                    prodDetails[attr["typeCode"]] = str(prodDetails[attr["typeCode"]]) + " @ " + str(attr["conditions"][0]["value"])
                        except:
                            pass
                except:
                    pass

                try:
                    for type in  prodData["attributeTypes"]:
                        if type["description"]:
                            prodDetails[type["code"] + "Description"] = type["description"]["en"]

                except:
                    pass
                
                # Updating keys format
                prodDetails = {convert_key_format(k): v for k, v in prodDetails.items()}
                df = pd.concat([df, pd.DataFrame([prodDetails.copy()])], ignore_index=True)
                         
        if df.shape[0] > 0:
            path = outputs[category]
            # drop irelevant columns
            cols = ["Answers To Questions", "Materials", "Materials Description", "Product Type Description"]
            for col in cols:
                if col in df.columns:
                    df.drop(col, axis=1, inplace=True)

            df = df.rename(columns={
                "Shipping Estimate": "Estimated Ready to Ship",
                "Product Link": "Product URL",
                "Image Link":"Product Image",
                "Description":"General Description",
                })

            df[["Estimated Ready to Ship", "Spray Pattern"]] = df[["Estimated Ready to Ship", "Spray Pattern"]].applymap(convert_key_format)
            df["Material Composition"] = df["Material Composition"].apply(lambda x:x.replace("_", " "))
            # Reorder the DataFrame
            orderedCols = ["Product Name", "Product URL", "Product Code", "Model", "Product Image", "Estimated Ready to Ship", "Product Bulletin Link", "Catalog Detail Link", "Interactive Model Link", "Product Description", "Capacity Size", "Capacity Size Description", "Inlet Connection Gender", "Inlet Connection Gender Description", "Inlet Connection Size", "Inlet Connection Size Description", "Inlet Connection Type", "Inlet Connection Type Description", "Inlet Connection Thread Type", "Inlet Connection Thread Type Description"]
            existingCols = [col for col in orderedCols if col in df.columns]
            remainingCols = [col for col in df.columns if col not in existingCols]
            newCols = existingCols + remainingCols          
            df = df[newCols]

            df.to_csv(path, index=False, encoding="windows-1252")      

def convert_key_format(key):
    # Return NaN as is if the value is NaN
    if pd.isna(key):
        return key
    # Add a space before each uppercase letter, except the first one
    try:
        return ''.join([' ' + char if (char.isupper() or char.isnumeric()) and i != 0 else char for i, char in enumerate(key)])
    except:
        return key

if __name__ == '__main__':

    start = time.time()
    settings = get_inputs()
    outputs = initialize_outputs(settings)
    try:
        scrape_prods(outputs, settings)
    except Exception as err: 
        print(f'Error: {err}')

    print('-'*75)
    time_mins = round(((time.time() - start)/60), 2)
    hrs = round(time_mins/60, 2)
    input(f'Scraping is completed successfully in {time_mins} mins ({hrs} hours). Press any key to exit.')