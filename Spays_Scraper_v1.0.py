import time
import json
import os
from datetime import datetime
import pandas as pd
import sys
import xlsxwriter
import warnings
import re
import ast
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
    industries = {
        "https://portal.spray.com/en-us/categories/food-and-beverage?all=true":"Food & Beverage",
        "https://portal.spray.com/en-us/categories/pulp-paper-and-tissue?all=true":"Pulp, Paper & Tissue",
        "https://portal.spray.com/en-us/categories/steel-and-metals?all=true":"Steel & Metals",
        "https://portal.spray.com/en-us/categories/car-wash?all=true":"Car Wash",
    }
    for category, status in settings.items():
        if status != 'yes': continue
        name = category.replace("-", ' ').title()
        print(f'Scraping Category: {name}')
        print('-'*75)
        df = pd.DataFrame()
        if category != "industries":
            urls = [f"https://portal.spray.com/en-us/categories/{category}?all=true"]
        else:
            urls = ["https://portal.spray.com/en-us/categories/food-and-beverage?all=true", "https://portal.spray.com/en-us/categories/pulp-paper-and-tissue?all=true", "https://portal.spray.com/en-us/categories/steel-and-metals?all=true", "https://portal.spray.com/en-us/categories/car-wash?all=true"]

        for url in urls:
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

            if category == "industries":
                print(f"Industry: {industries[url]}")

            print(f"Number of products: {nprods}")
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
                    #prodUrl = "https://portal.spray.com/en-us/products/63225e-1-2-40-3a" 
                    iprod += 1
                    print(f'Scraping Product {iprod}/{nprods}')
                    for _ in range(10):
                        try:
                            response = requests.get(prodUrl)
                            time.sleep(0.5)
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

                    if category == "industries":
                        prodDetails["Industry"] = industries[url]

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
                        counter = {}
                        for resource in  prodData["product"]["resources"]:
                            if "variation" in resource and resource["variation"] != "Invariant":
                                name = resource["type"] + resource["variation"]
                            else:
                                name = resource["type"]
                            if name + 'Link' not in prodDetails:
                                prodDetails[name + 'Link'] = resource["url"]
                            else:
                                if name + 'Link' not in counter:
                                    counter[name + 'Link'] = 2
                                else:
                                    counter[name + 'Link'] += 1

                                prodDetails[name + str(counter[name + 'Link']) + "Link"] = resource["url"]
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
                                # Initialize the base value without any suffix
                                base_key = attr["typeCode"]
                                # Loop over each variation within the "displays" list if it exists
                                if "displays" in attr and attr["displays"]:
                                    for display in attr["displays"]:
                                        # Determine suffix based on the "variation" key
                                        if "variation" in display and display["variation"] != "Invariant":
                                            suffix = display["variation"]
                                        else:
                                            suffix = ""

                                        # Construct the full key with suffix and retrieve the value
                                        key = base_key + suffix
                                        try:
                                            value = display["value"]["en"]
                                        except:
                                            value = display["value"]

                                        # Check for a unit symbol and add it to the value if it exists
                                        if "unitSymbol" in display and display["unitSymbol"]:
                                            value = f"{value} {display['unitSymbol']}"

                                        # Check for conditions and append if they exist
                                        if "conditions" in attr and attr["conditions"]:
                                            try:
                                                condition = attr["conditions"][0]
                                                if "displays" in condition and condition["displays"]:
                                                    for conditionDisplay in condition["displays"]:
                                                        if conditionDisplay["variation"] == display["variation"]:
                                                            try:
                                                                condition_value = conditionDisplay["value"]["en"]
                                                            except:
                                                                condition_value = conditionDisplay["value"]
                                                            condition_unit = conditionDisplay.get("unitSymbol", "")
                                                            if condition_unit:
                                                                value += f" @ {condition_value} {condition_unit}"
                                                            else:
                                                                value += f" @ {condition_value}"
                                                        elif conditionDisplay["variation"] != display["variation"] and (display["variation"] == "Invariant" or conditionDisplay["variation"] == "Invariant"):
                                                            try:
                                                                condition_value = conditionDisplay["value"]["en"]
                                                            except:
                                                                condition_value = conditionDisplay["value"]
                                                            condition_unit = conditionDisplay.get("unitSymbol", "")
                                                            if condition_unit:
                                                                value += f" @ {condition_value} {condition_unit}"
                                                            else:
                                                                value += f" @ {condition_value}"
                                            except:
                                                pass

                                        # Assign the final value to prodDetails with the constructed key
                                        if key not in prodDetails:
                                            prodDetails[key] = value
                                        elif key in prodDetails and prodDetails[key] != value:
                                            prodDetails[key] += ', ' + value


                                else:
                                    # If no variations are found, simply store the default value
                                    prodDetails[base_key] = attr["value"]
                                    if "unitSymbol" in attr:
                                        prodDetails[base_key] += f" {attr['unitSymbol']}"
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
                    if "Model" not in prodDetails:
                        try:
                            prodDetails["Model"] = re.findall(r'<button type="button" class="ms-Link root-\d+">(.*?)</button>', response.text)[0]
                        except:
                            pass

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
                "Image 2 Link":"Product Image 2",
                "Image 3 Link":"Product Image 3",
                "Image 4 Link":"Product Image 4",
                "Image 5 Link":"Product Image 5",
                "Image 6 Link":"Product Image 6",
                "Image 7 Link":"Product Image 7",
                "Image 8 Link":"Product Image 8",
                "Image 9 Link":"Product Image 9",
                "Description":"General Description",
                "Drop Size D M Z":"Drop Size (Sauter Mean Diameter)",
                })

            try:
                if "Estimated Ready to Ship" in df.columns:
                    df["Estimated Ready to Ship"] = df["Estimated Ready to Ship"].apply(convert_key_format)
                if "Spray Pattern" in df.columns:
                    df["Spray Pattern"] = df["Spray Pattern"].apply(convert_key_format)
                if "Material Composition" in df.columns:
                    df["Material Composition"] = df["Material Composition"].apply(lambda x:x.replace("_", " ") if isinstance(x, str) else x)

                # update range columns format
                for col in df.columns:
                    if "Range" in col and "Description" not in col:
                        df[col] = df[col].apply(convert_range_string)
            except Exception as err:
                print(f"Error: Failed to process the df columns")
                print(err)

            # Reorder the DataFrame
            orderedCols = ["Product Name", "Product URL", "Product Code", "Product Image", "Product Image 2", "Product Image 3", "Product Image 4", "Product Image 5", "Product Image 6", "Product Image 7", "Product Image 8", "Product Image 9", "Estimated Ready to Ship", "Industry", "Product Bulletin Link", "Product Bulletin 2 Link", "Product Bulletin 3 Link", "Product Bulletin 4 Link", "Product Bulletin 5 Link", "Product Bulletin 6 Link", "Product Bulletin 7 Link", "Product Bulletin 8 Link", "Product Bulletin 9 Link", "Catalog Detail Link", "Catalog Detail Metric Link", "Catalog Detail Us Link", "Interactive Model Link", "Video Link", "Video 2 Link", "Video 3 Link", "Video 4 Link", "Video 5 Link", "Video 6 Link", "Video 7 Link", "Video 8 Link", "Video 9 Link", "Flow Image Link", "Flow Image 2 Link", "Flow Image 3 Link", "Flow Image 4 Link", "Flow Image 5 Link", "Flow Image 6 Link", "Flow Image 7 Link", "Flow Image 8 Link", "Flow Image 9 Link", "Case Study Link", "Case Study 2 Link", "Case Study 3 Link", "Case Study 4 Link", "Case Study 5 Link", "Case Study 6 Link", "Case Study 7 Link", "Case Study 8 Link", "Case Study 9 Link", "Data Sheet Link", "Data Sheet Metric Link", "Data Sheet Us Link", "General Description", "Body Type", "Air Cap Component", "Fluid Cap Component", "Capacity Size", "Capacity Size Description", "Inlet Connection Gender", "Inlet Connection Gender Description", "Outlet Connection Gender", "Inlet Connection Size", "Inlet Connection Size Description", "Outlet Connection Size", "Inlet Connection Type", "Inlet Connection Type Description", "Outlet Connection Type", "Outlet Connection Type Description", "Liquid Flow Rate at Rated Pressure", "Liquid Flow Rate at Rated Pressure Description", "Cap Hex Size", "Height Us", "Height Metric", "Length Us", "Length Metric", "Length Description", "Width Us", "Width Metric", "Maximum Air Pressure", "Maximum Flow", "Maximum Operating Speed", "Maximum Pressure", "Spray Tips", "Voltage", "Nozzle Count", "Inlet Connection Thread Type", "Inlet Connection Thread Type Description", "Outlet Connection Thread Type", "Material Code", "Material Composition", "Material Composition Description", "Model", "Spray Angle at Rated Pressure", "Spray Angle at Rated Pressure Description", "Tip Type", "Tip Type Description", "Design Feature", "Design Feature Description", "Setup Mix Type", "Setup Type", "Product Type", "Relative Drop Size Group", "Relative Drop Size Group Description", "Spray Angle Range", "Spray Angle Range Description", "Spray Angle Us", "Spray Angle Metric", "Spray Angle Description", "Air Cap Part Number", "Fluid Cap Part Number", "Compatible Needle Size", "Operating Pressure Range Metric", "Operating Pressure Range Us", "Brand", "Brand Description", "Spray Angle Catalog Code", "Spray Angle Catalog Code Description", "Approximate Free Passage Diameter Us", "Approximate Free Passage Diameter Metric", "Equivalent Orifice Diameter Us", "Equivalent Orifice Diameter Metric", "Equivalent Orifice Diameter Description", "Color", "Body Sales Part Number", "Tip Sales Part Number", "Body Hex Size", "Impact Group", "A Dimension Metric", "A Dimension Us", "B Dimension Metric", "B Dimension Us", "C Dimension Metric", "C Dimension Us", "D Dimension Metric", "D Dimension Us", "E Dimension Metric", "E Dimension Us", "Liquid Flow Rate Range Metric", "Liquid Flow Rate Range Us", "Liquid Flow Rate Range Description", "Liquid Pressure Range Us", "Liquid Pressure Range Metric", "Liquid Pressure Range Description", "Rated Pressure Us", "Rated Pressure Metric", "Rated Pressure Description", "Relative Drop Size Range", "Relative Drop Size Range Description", "Maximum Free Passage", "Maximum Recommended Tank Diameter Metric", "Maximum Recommended Tank Diameter Us", "Maximum Recommended Tank Diameter Description", "Maximum Temperature Metric", "Maximum Temperature Us", "Mounting Points", "Minimum Tank Opening Metric", "Minimum Tank Opening Us", "Operating Principle", "Recommended Strainer Mesh", "Spray Coverage", "Spray Coverage Description", "Tank Mounting Options", "Tank Mounting Options Description", "Spray Pattern", "Spray Pattern Description", "Weight Us", "Weight Metric", "Air Flow Rate Us", "Air Flow Rate Metric", "Price Type", "Audience", "Marketing Score", "Marketing Score Description", "Sales Score", "Sales Score Description", "Business Score", "Business Score Description"]
            try:
                existingCols = [col for col in orderedCols if col in df.columns]
                remainingCols = [col for col in df.columns if col not in existingCols]
                newCols = existingCols + remainingCols          
                df = df[newCols]
                df = sort_description_columns(df, orderedCols)
            except Exception as err:
                print(f"Error: Failed to reorder the df columns")
                print(err)
                
            df.to_csv(path, index=False, encoding="windows-1252")      

def convert_range_string(input_str):
    if pd.isna(input_str):
        return input_str
    # Extract the dictionary part and the unit using regex
    match = re.match(r"\{(.+?)\}\s*(\S+)", input_str)
    if not match:
        return input_str  # Return as-is if the format is unexpected
    
    dict_str, unit = match.groups()
    
    # Safely parse the dictionary string
    try:
        range_dict = ast.literal_eval("{" + dict_str + "}")
        minimum = range_dict.get('minimum')
        maximum = range_dict.get('maximum')
        
        # Format the output as "min - max unit"
        return f"{minimum} - {maximum} {unit}"
    except (ValueError, SyntaxError):
        return input_str  # Return as-is if there's an error in parsing

# Sort function
def sort_description_columns(df, orderedCols):
    columns = df.columns.tolist()
    reordered_columns = []
    descreption_columns = {}
    for col in columns:
        if col.endswith(" Description"):
            base_name = col.replace(" Description", "")
            descreption_columns[base_name] = col  # Map base name to its description column

    # Separate base and "Description" columns
    for col in columns:
        if col in orderedCols:
            reordered_columns.append(col)
        elif col.endswith(" Description"):
            continue
        else:
            # Add the base column to reordered list
            reordered_columns.append(col)

            # Check if there are any description columns matching this base column with optional " Us" or " Metric" suffix
            base_name = col.replace(" Us", "").replace(" Metric", "")
            if base_name in descreption_columns:
                # Ensure that the description is added only once, after the last occurrence of the base column variants
                if descreption_columns[base_name] not in reordered_columns:
                    reordered_columns.append(descreption_columns[base_name])
                else:
                    reordered_columns.remove(descreption_columns[base_name])
                    reordered_columns.append(descreption_columns[base_name])

    # Return the DataFrame with reordered columns
    return df[reordered_columns]

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