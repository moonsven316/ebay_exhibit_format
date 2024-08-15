import requests
import pandas as pd
import time
import csv
import random
import re
import os
from google.cloud import translate_v2 as translate
from urllib.parse import urlparse
from forex_python.converter import CurrencyRates

os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = "credential.json"

data_format = {
    "url":"",
    "price":"",
    "shipping":"",
    "title":"",
    "description":"",
    "category":"",
    "size":"",
    "brand":"",
    "color":"",
    "status":"",
    "sellerUrl":"",
    "likes":"",
    "comments":"",
    "pageViews":"",
    "image1":"",
    "image2":"",
    "image3":"",
    "image4":"",
    "image5":"",
    "image6":"",
    "image7":"",
    "image8":"",
    "image9":"",
    "image10":"",
    "image11":"",
    "image12":"",
    "image13":"",
    "image14":"",
    "image15":"",
    "image16":"",
    "image17":"",
    "image18":"",
    "image19":"",
    "image20":"",
    "Title":"",
    "Description":"",
    "MPN":"",
    "REF":"",
    "EPID":"",
    "UPC":"",
    "Brand":"Seiko",
    "Model":"",
    "Movement":"",
    "Beats per Hour":"",
    "Jewels":"",
    "Style":"",
    "Case Material":"",
    "Water Resistance":"",
    "Watch Shape":"",
    "Case Thickness":"",
    "Case Width":"",
    "Band Width":"",
    "sum_size":"",
    "specific_keyword":"",
    "ebay_html":""
}
file_path = ['2024-07-11T09_59_51_MERCARI.xlsx', 
             '2024-06-25T10_14_39_PAYPAY.xlsx', 
             '2024-07-15T11_53_49_RAKUMA.xlsx']
url_text = "data/URL.txt"
with open(url_text, 'r') as file:
    content = file.read()
remove_url_list = content.split("\n")
replacements_df = pd.read_csv("data/キーワードの置換.csv", sep=',', header=0, index_col=None, usecols=['置換前', '置換後'])
remove_product = pd.read_csv("data/出品の可否判定-2.csv", sep=',', header=0, index_col=None, usecols=['削除対象キーワード','ブランド1', 'ブランド2', '出品者URL'])
database = pd.read_csv("data/deta base.csv", sep=',', header=0, index_col=None, usecols=['MPN', 'Ref', 'EPID', 'Title', 'Brand', 'Model', 'Movement', 'Beats per Hour', 'Jewels', 'Style', 'Case Material', 'Water Resistance', 'Watch Shape', 'Case Thickness', 'Case Width', 'Band Width'])
with open("data/動作取得キーワード.txt", 'r') as file:
    specific_keyword = file.read().splitlines()  
description_replace = pd.read_csv("data/商品状態データテーブル.csv", sep=',', header=0, index_col=None, usecols=['状態', 'キーワード'])
with open("data/HTMLテンプレート.txt", 'r', encoding='utf-8') as file:
    ebay_html = file.read()
with open("data/削除キーワード.txt", 'r') as file:
    remove_keywords = file.read().splitlines()  
   
def first_step(row):
    print(111111111111)
    if row["URL"] in remove_url_list:
        return
    if row["いいねの数"] == 0:
        return
    data_format["url"] = row["URL"] if pd.notna(row["URL"]) else ""
    data_format["price"] = row["価格"] if pd.notna(row["価格"]) else ""
    try:
        data_format["shipping"] = row["配送料の負担"] if pd.notna(row["配送料の負担"]) else ""
    except KeyError:
        data_format["shipping"] = ""
    try:
        data_format["title"] = row["タイトル"] if pd.notna(row["タイトル"]) else ""
    except KeyError:
        data_format["title"] = row["商品タイトル"] if pd.notna(row["商品タイトル"]) else ""
    try:
        data_format["description"] = row["商品の説明"] if pd.notna(row["商品の説明"]) else ""
    except KeyError:
        data_format["description"] = row["商品説明"] if pd.notna(row["商品説明"]) else ""
    try:
        data_format["category"] = row["カテゴリ"] if pd.notna(row["カテゴリ"]) else ""
    except KeyError:
        data_format["category"] = row["カテゴリー"] if pd.notna(row["カテゴリー"]) else ""
    try:
        data_format["size"] = row["商品のサイズ"] if pd.notna(row["商品のサイズ"]) else ""
    except KeyError:
        try:
            data_format["size"] = row["サイズ（S/M/L）"] if pd.notna(row["サイズ（S/M/L）"]) else ""
        except KeyError:
            data_format["size"] = row["サイズ"] if pd.notna(row["サイズ"]) else ""
    data_format["brand"] = row["ブランド"] if pd.notna(row["ブランド"]) else ""
    try:
        data_format["color"] = row["カラー"] if pd.notna(row["カラー"]) else ""
    except KeyError:
        try:
            data_format["color"] = row["色"] if pd.notna(row["色"]) else ""
        except KeyError:
            data_format["color"] = ""
    data_format["status"] = row["商品の状態"] if pd.notna(row["商品の状態"]) else ""
    try:
        data_format["sellerUrl"] = row["出品者URL"] if pd.notna(row["出品者URL"]) else ""
    except KeyError:
        data_format["sellerUrl"] = row["出品者のURL"] if pd.notna(row["出品者のURL"]) else ""
    try:
        data_format["likes"] = row["いいねの数"] if pd.notna(row["いいねの数"]) else ""
    except KeyError:
        data_format['likes'] = row["いいね数"] if pd.notna(row["いいね数"]) else ""
    try:
        data_format["comments"] = row["コメントの数"] if pd.notna(row["コメントの数"]) else ""
    except KeyError:
        data_format["comments"] = ""
    try:
        data_format["pageViews"] = row["ページビューの数"] if pd.notna(row["ページビューの数"]) else ""
    except KeyError:
        data_format["pageViews"] = ""
    data_format["image1"] = row["画像1"] if pd.notna(row["画像1"]) else ""
    data_format["image2"] = row["画像2"] if pd.notna(row["画像2"]) else ""
    data_format["image3"] = row["画像3"] if pd.notna(row["画像3"]) else ""
    data_format["image4"] = row["画像4"] if pd.notna(row["画像4"]) else ""
    data_format["image5"] = row["画像5"] if pd.notna(row["画像5"]) else ""
    data_format["image6"] = row["画像6"] if pd.notna(row["画像6"]) else ""
    data_format["image7"] = row["画像7"] if pd.notna(row["画像7"]) else ""
    data_format["image8"] = row["画像8"] if pd.notna(row["画像8"]) else ""
    data_format["image9"] = row["画像9"] if pd.notna(row["画像9"]) else ""
    data_format["image10"] = row["画像10"] if pd.notna(row["画像10"]) else ""
    second_step(data_format)
    
def second_step(data_format):
    print(222222)
    def replace_keywords(text, replacements_df):
        for index, row in replacements_df.iterrows():
            old_keyword = row['置換前']
            if pd.isna(row['置換後']):
                new_keyword = ""
            else:
                new_keyword = row['置換後']
            text = text.replace(str(old_keyword), str(new_keyword))
        return text
    data_format['title'] = replace_keywords(data_format['title'], replacements_df)
    data_format['description'] = replace_keywords(data_format['description'], replacements_df)
    fourth_step(data_format)

def fourth_step(data_format):
    print(44444)
    keywords = remove_product["削除対象キーワード"]
    if any(pd.notna(keyword) and (keyword in data_format["title"] or keyword in data_format["description"]) for keyword in keywords):
        return
    
    if all(pd.notna(brand) and brand not in data_format["title"] for brand in remove_product["ブランド1"]):
        return
    
    if all(pd.notna(brand) and brand not in data_format["brand"] for brand in remove_product["ブランド2"]):
        return
    
    if any(pd.notna(url) and url in data_format["sellerUrl"] for url in remove_product["出品者URL"]):
        return
    
    fifth_step(data_format)
    
def fifth_step(data_format):
    print(555555)
    def get_brand():
        for brand in database["Brand"]:
            if pd.isna(brand):
                continue
            if brand in data_format["title"]:
                return brand
        return "Seiko"
    
    def get_model():
        for model in database["Model"]:
            if pd.isna(model):
                continue
            if model in data_format["title"]:
                return model
        return ""
    
    def get_move():
        for move in database["Movement"]:
            if pd.isna(move):
                continue
            if move in data_format["title"]:
                return move
        return ""
    
    def get_beats():
        for beats in database["Beats per Hour"]:
            if pd.isna(beats):
                continue
            if beats in data_format["title"]:
                return beats
        return ""
    
    def get_jewel():
        for jewel in database["Jewels"]:
            if pd.isna(jewel):
                continue
            if jewel in data_format["title"]:
                return jewel
        return ""
    
    def get_style():
        for style in database["Style"]:
            if pd.isna(style):
                continue
            if style in data_format["title"]:
                return style
        return ""

    # ----------------------------------------
    titleMpns = []
    descMpns = []
    
    for index, row in database.iterrows():
        if pd.isna(row["MPN"]):
            continue
        if row["MPN"] in data_format["title"]:
            titleMpns.append(row["MPN"])
        if row["MPN"] in data_format["description"]:
            descMpns.append(row["MPN"])
            
    if len(titleMpns):
        if len(titleMpns) == 1:
            data_format["MPN"] = titleMpns[0]
    else:
        if len(descMpns):
            if len(descMpns) == 1:
                data_format["MPN"] = descMpns[0]
    # ----------------------------------------            
    titleRefs = []
    descRefs = []
    
    for index, row in database.iterrows():
        if pd.isna(row["Ref"]):
            continue
        if pd.isna(row["MPN"]):
            continue
        if row["MPN"] == data_format["MPN"]:
            data_format["REF"] = row["Ref"]
            break
        
    if data_format["REF"] == "":
        for index, row in database.iterrows():
            if pd.isna(row["Ref"]):
                continue
            if row["Ref"] in data_format["title"]:
                titleRefs.append(row["Ref"])
            if row["Ref"] in data_format["description"]:
                descRefs.append(row["Ref"])
            
        if len(titleRefs):
            if len(titleRefs) == 1:
                data_format["REF"] = titleRefs[0]
        else:
            if len(descRefs):
                if len(descRefs) == 1:
                    data_format["REF"] = descRefs[0]
    # ----------------------------------------
    for index, row in database.iterrows():
        if pd.isna(row["EPID"]):
            continue
        if row["MPN"] == data_format["MPN"]:
            data_format["EPID"] = row["EPID"]
        if row["Ref"] == data_format["REF"]:
            data_format["EPID"] = row["EPID"]
    if data_format["EPID"] != "":
        pass
    # -----------------------------------------
    for index, row in database.iterrows():
        if pd.isna(row["Brand"]):
            continue
        if data_format["MPN"] == "" and data_format["REF"] == "":
            data_format["Brand"] = get_brand()
            data_format["Model"] = get_model()
            data_format["Movement"] = get_move()
            data_format["Beats per Hour"] = get_beats()
            data_format["Jewels"] = get_jewel()
            data_format["Style"] = get_style()
        else:
            if data_format["MPN"] != "":
                if data_format["MPN"] == row["MPN"]:
                    data_format["Brand"] = row["Brand"]
                    data_format["Model"] = row["Model"] if not pd.isna(row["Model"]) else ""
                    data_format["Movement"] = row["Movement"] if not pd.isna(row["Movement"]) else ""
                    data_format["Beats per Hour"] = row["Beats per Hour"] if not pd.isna(row["Beats per Hour"]) else ""
                    data_format["Jewels"] = row["Jewels"] if not pd.isna(row["Jewels"]) else ""
                    data_format["Style"] = row["Style"] if not pd.isna(row["Style"]) else ""
                    data_format["Case Material"] = row["Case Material"] if not pd.isna(row["Case Material"]) else ""
                    data_format["Water Resistance"] = row["Water Resistance"] if not pd.isna(row["Water Resistance"]) else ""
                    data_format["Watch Shape"] = row["Watch Shape"] if not pd.isna(row["Watch Shape"]) else ""
                    data_format["Case Thickness"] = row["Case Thickness"] if not pd.isna(row["Case Thickness"]) else ""
                    data_format["Case Width"] = row["Case Width"] if not pd.isna(row["Case Width"]) else ""
                    data_format["Band Width"] = row["Band Width"] if not pd.isna(row["Band Width"]) else ""
                    break
            if data_format["REF"] != "":       
                if data_format["REF"] == row["Ref"]:
                    data_format["Brand"] = row["Ref"]
                    data_format["Model"] = row["Model"] if not pd.isna(row["Model"]) else ""
                    data_format["Movement"] = row["Movement"] if not pd.isna(row["Movement"]) else ""
                    data_format["Beats per Hour"] = row["Beats per Hour"] if not pd.isna(row["Beats per Hour"]) else ""
                    data_format["Jewels"] = row["Jewels"] if not pd.isna(row["Jewels"]) else ""
                    data_format["Style"] = row["Style"] if not pd.isna(row["Style"]) else ""
                    data_format["Case Material"] = row["Case Material"] if not pd.isna(row["Case Material"]) else ""
                    data_format["Water Resistance"] = row["Water Resistance"] if not pd.isna(row["Water Resistance"]) else ""
                    data_format["Watch Shape"] = row["Watch Shape"] if not pd.isna(row["Watch Shape"]) else ""
                    data_format["Case Thickness"] = row["Case Thickness"] if not pd.isna(row["Case Thickness"]) else ""
                    data_format["Case Width"] = row["Case Width"] if not pd.isna(row["Case Width"]) else ""
                    data_format["Band Width"] = row["Band Width"] if not pd.isna(row["Band Width"]) else ""
                    break

    sixth_step(data_format)

def sixth_step(data_format):
    print(66666)
    def remove_duplicates_and_limit_length(input_string, max_length=73):
        words = input_string.split(" ")
        seen = set()
        unique_words = []
        
        for word in words:
            normalized_word = word.lower()
            if normalized_word not in seen:
                seen.add(normalized_word)
                unique_words.append(word)
        
        result = " ".join(unique_words)
        
        if len(result) > max_length:
            result = result[:max_length].rsplit(' ', 1)[0]

        return result
    
    def translate_text(text, target_language):
        print('translate')
        client = translate.Client()
        result = client.translate(text, target_language=target_language)
        return result['translatedText']
    
    def make_title(data_format):
        print('maketitle')
        product_info = f"{data_format['Brand']} {data_format['Model']} {data_format['MPN']} {data_format['REF']} {data_format['Movement']} {data_format['Beats per Hour']} {data_format['Jewels']} {data_format['Style']}"

        title = data_format['title']
        for info in [data_format["Brand"], data_format["Model"], data_format['MPN'], data_format['REF'], data_format['Movement'], data_format['Beats per Hour'], data_format['Jewels'], data_format['Style']]:
            title = title.replace(info, '').strip()
            
        target_language = "en"
        translated_title = translate_text(title, target_language)
        for remove_key in remove_keywords:
            if remove_key in translated_title:
                translated_title.replace(remove_key, "")

        combined_title = f"{product_info} {translated_title}"

        combined_title = remove_duplicates_and_limit_length(combined_title)

        random_number = str(random.randint(0, 99))
        final_title = f"{combined_title} ({random_number})"
        finaly_title = final_title.title()
        return finaly_title
    
    random_number = random.randint(10, 99)
    if data_format["MPN"] == "" and data_format["REF"] == "":
        data_format["Title"] = make_title(data_format)
    else:
        for index, row in database.iterrows():
            if data_format["MPN"] == row["MPN"] or data_format["REF"] == row["Ref"]:
                if data_format["MPN"] == row["MPN"]:
                    data_format["Title"] = f'{row["Title"]}({random_number})'
                    break
                if data_format["REF"] == row["Ref"]:
                    data_format["Title"] = f'{row["Title"]}({random_number})'
                    break
        data_format["Title"] = make_title(data_format)
                
    seventh_step(data_format)
    
def seventh_step(data_format):
    print(77777)
    data_format["specific_keyword"] = ""
    def get_number(text):
        number_pattern = re.compile(r'\d+(?:\.\d+)?')
        match = number_pattern.search(text)
        if match:
            return match.group()
        else:
            return None
    def extract_number(text, keywords, units):
        keyword_pattern = '|'.join(map(re.escape, keywords))
        unit_pattern = '|'.join(map(re.escape, units))
        
        pattern = re.compile(rf"({keyword_pattern})\s*(\d+(?:\.\d+)?)(?:\s*({unit_pattern}))?", re.IGNORECASE)
        
        matches = pattern.findall(text)
        
        results = [f"{match[0]}{match[1]}{match[2] or ''}" for match in matches]
    
        return results if results else None
    
    keyword = ["腕周", "日差", "直径", "縦", "ラグ幅"]
    unit = ["cm", "mm", "CM", "MM", "分", "秒"]

    size_result = extract_number(data_format["title"], keyword, unit)
    if size_result == None:
        size_result = extract_number(data_format["description"], keyword, unit)
    
    if size_result is not None:
        wristSize = ""
        dailyRate = ""
        Diameter = ""
        Length = ""
        lugWidth = ""
        for size in size_result:
            if "腕周" in size:
                if "CM" in size:
                    arm_size = float(get_number(size))
                    if 10 <= arm_size <= 30:
                        wristSize = f"{arm_size}cm"
                else:
                    arm_size_pattern = re.compile(r"腕周\s*(\d+(?:\.\d+)?)\s*(MM)?", re.IGNORECASE)
                    arm_size_matches = arm_size_pattern.findall(size)
                    if arm_size_matches:
                        arm_size_numbers = float(arm_size_matches[0][0]) * 0.1
                        if 10 <= arm_size_numbers <= 30:
                            wristSize = f"{arm_size_numbers}cm"

            elif "日差" in size:
                if "秒" in size:
                    second = float(get_number(size))
                    if 1 <= second <= 600:
                        dailyRate = f"{second} Seconds"
                else:
                    second_pattern = re.compile(r"日差\s*(\d+(?:\.\d+)?)\s*(分)?", re.IGNORECASE)
                    second_matches = second_pattern.findall(size)
                    if second_matches:
                        second_numbers = float(second_matches[0][0]) * 60
                        if 1 <= second_numbers <= 600:
                            dailyRate = f"{second_numbers} Seconds"

            elif "直径" in size:
                if "MM" in size:
                    diameter = float(get_number(size))
                    if 10 <= diameter <= 50:
                        Diameter = f"{diameter}mm"
                else:
                    diameter_pattern = re.compile(r"直径\s*(\d+(?:\.\d+)?)\s*(CM)?", re.IGNORECASE)
                    diameter_matches = diameter_pattern.findall(size)
                    if diameter_matches:
                        diameter_numbers = float(diameter_matches[0][0]) * 10
                        if 10 <= diameter_numbers <= 50:
                            Diameter = f"{diameter_numbers}mm"

            elif "縦" in size:
                if "MM" in size:
                    vertical = float(get_number(size))
                    if 10 <= vertical <= 50:
                        Length = f"{vertical}mm"
                else:
                    vertical_pattern = re.compile(r"縦\s*(\d+(?:\.\d+)?)\s*(CM)?", re.IGNORECASE)
                    vertical_matches = vertical_pattern.findall(size)
                    if vertical_matches:
                        vertical_numbers = float(vertical_matches[0][0]) * 10
                        if 10 <= vertical_numbers <= 50:
                            Length = f"{vertical_numbers}mm"

            elif "ラグ幅" in size:
                if "MM" in size:
                    lug_width = float(get_number(size))
                    if 10 <= lug_width <= 26:
                        lugWidth = f"{lug_width}mm"
                else:
                    lug_width_pattern = re.compile(r"ラグ幅\s*(\d+(?:\.\d+)?)\s*(CM)?", re.IGNORECASE)
                    lug_width_matches = lug_width_pattern.findall(size)
                    if lug_width_matches:
                        lug_width_numbers = float(lug_width_matches[0][0]) * 10
                        if 10 <= lug_width_numbers <= 26:
                            lugWidth = f"{lug_width_numbers}mm"
                            
        data_format["sum_size"] = f"Wrist size:{wristSize}<br>Daily rate:{dailyRate}<br>Diameter:{Diameter}<br>Length:{Length}<br>Lug width:{lugWidth}<br>"
    for key in specific_keyword:
        if key in data_format["title"]:
            data_format["specific_keyword"] += f"{key}<br>"
        elif key in data_format["description"]:
            data_format["specific_keyword"] += f"{key}<br>"
    
    for index, row in description_replace.iterrows():
        if row["状態"] in data_format["status"]:
            data_format["status"] = data_format["status"].replace(str(row["状態"]), str(row["キーワード"]))
            
    data_format["Description"] = f"{data_format["status"]} {data_format["specific_keyword"]} {data_format["sum_size"]}"
    
    data_format["ebay_html"] = ebay_html
    data_format["ebay_html"] = data_format["ebay_html"].replace("[商品名]", data_format["Title"])
    data_format["ebay_html"] = data_format["ebay_html"].replace("[商品説明]", data_format["Description"])

    eighth_step(data_format)
    
def eighth_step(data_format):
    print(888888)
    urls = []
    def process_url(url, prefix):
        parsed_url = urlparse(url)
        path = parsed_url.path
        
        file_name = path.split('/')[-1].split('?')[0]
        
        processed_url = f"{prefix}{file_name}"
        
        return processed_url

    def create_item_photo_urls(urls, prefix):
        processed_urls = [process_url(url, prefix) for url in urls]
        return '|'.join(processed_urls)
    for i in range(1, 11):
        if data_format[f"image{i}"] == "":
            continue
        else:
            urls.append(data_format[f"image{i}"])

    prefix = "http://export.daa.jp/20240618/"
    item_photo_url = create_item_photo_urls(urls, prefix)
    
    if "新品、未使用" in data_format["status"]:
        status = 1000
    else:
        status = 3000
    
    price = data_format["price"]
    if "着払い(購入者負担)" in data_format["shipping"]:
        domestic_shipping = 1000
    else:
        domestic_shipping = 0
    international_shipping = 2933
    
    try:
        response = requests.get("https://api.exchangerate-api.com/v4/latest/USD")
        response.raise_for_status()
        data = response.json()
        rate = data['rates']['JPY']
    except requests.exceptions.RequestException as e:
        print(f"Request error: {e}")
    except ValueError as e:
        print(f"JSON decode error: {e}")
    
    total_cost_jpy = price + domestic_shipping + international_shipping
    if total_cost_jpy > 100000:
        variable = 0.73
    elif total_cost_jpy > 40000:
        variable = 0.69
    elif total_cost_jpy > 15000:
        variable = 0.66
    else:
        variable = 0.64
    start_price = total_cost_jpy / rate / variable
    
    if status == 1000:
        shipping_profile_name = "new-watch-FE"
    elif status == 3000:
        shipping_profile_name = "used-watch-FE"
        
    if "https://jp.mercari.com" in data_format["url"]:
        url = data_format["url"].replace(str("https://jp.mercari.com/item/m"), str("#m@"))
        sku_label = url + "{a}" + str(price) + "{b}" + str(int(start_price * 0.95)) + "{c}" + shipping_profile_name
    elif "https://paypayfleamarket.yahoo.co.jp" in data_format["url"]:
        url = data_format["url"].replace(str("https://paypayfleamarket.yahoo.co.jp/item"), str("#p@"))
        sku_label = url + "{a}" + str(price) + "{b}" + str(int(start_price * 0.95)) + "{c}" + shipping_profile_name
    elif "https://item.fril.jp" in data_format["url"]:
        url = data_format["url"].replace(str("https://item.fril.jp/"), str("#r@"))
        sku_label = url + "{c}" + shipping_profile_name
        
    if data_format["MPN"] != "":
        reference_number = data_format["MPN"]
    else:
        reference_number = ""
    if reference_number == "":
        reference_number = data_format["REF"]
    
    data = {
        "Action": "SiteID=US|Country=JP|Currency=USD|Version=1193",
        "Category ID": 31387,
        "Title": data_format["Title"],
        "Condition ID": status,
        "Item photo URL": item_photo_url,
        "Description": data_format["Description"],
        "Format": "FixedPrice",
        "Duration": "GTC",
        "Start price": start_price,
        "Quantity": 1,
        "Location": "Japan",
        "Shipping profile name": shipping_profile_name,
        "Return profile name": "returns1",
        "Payment profile name": "eBay Payments:Immediate pay",
        "Custom label (SKU)": sku_label,
        "C:Brand": data_format["Brand"],
        "C:Department": "Men",
        "C:Type": "Wristwatch",
        "C:Model": data_format["Model"],
        "C:Reference Number": reference_number,
        "P:EPID": data_format["EPID"],
        "P:UPC": data_format["UPC"],
        "P:EAN": "",
        "C:Style": data_format["Style"],
        "C:Movement": data_format["Movement"],
        "C:Case Size": data_format["Case Width"],
        "C:Case Material": data_format["Case Material"],
        "C:Water Resistance": data_format["Water Resistance"],
        "C:Watch Shape": data_format["Watch Shape"],
        "C:Band Width": data_format["Band Width"],
        "C:Case Thickness": data_format["Case Thickness"]
    }
    print(data["Title"])
    out = pd.DataFrame([data])
    out.to_csv('output.csv', mode='a', header=not pd.io.common.file_exists('output.csv'), index=False, encoding='utf-8-sig')
    
    with open("data/URL.txt", "a") as file:
        file.write(data_format['url'] + "\n")
        
    for i in range(1, 11):
        if data_format[f"image{i}"] != "":
            with open("data/画像URL.txt", "a") as file:
                file.write(data_format[f"image{i}"] + "\n")

def main():
    for file in file_path:
        df = pd.read_excel(file)
        for index, row in df.iterrows():
            first_step(row)

if __name__ == '__main__':
    main()

