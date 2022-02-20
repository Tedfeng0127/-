# 抓取momo購物網行動版網頁上所有冰箱資料
# 產品連結、產品名稱、價格、品牌名稱、類型(Ex:雙門上下門、右開、變頻)、容量、效能


import requests
from bs4 import BeautifulSoup
import pymysql


def get_all_href(start_page, end_page):
    """
    抓取所有商品的連結網址(需要先看一下搜尋結果總共有幾頁)
    回傳一個包含所有連結網址的列表
    start_page:起始頁碼
    end_page:結束頁碼
    """
    returnlist = []
    headers = {"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/97.0.4692.71 Safari/537.36"}
    for page in range(start_page, end_page + 1):
        # 這個網址是momo購物網行動版網頁搜尋"冰箱"之後頁面的網址
        url = "https://m.momoshop.com.tw/search.momo?searchKeyword=%E5%86%B0%E7%AE%B1&couponSeq=&searchType=1&cateLevel=-1&curPage={}&cateCode=&cateName=&maxPage=18&minPage=1&_advCp=N&_advFirst=N&_advFreeze=N&_advSuperstore=N&_advTvShop=N&_advTomorrow=N&_advNAM=N&_advStock=N&_advPrefere=N&_advThreeHours=N&_advPriceS=&_advPriceE=&_brandNameList=&_brandNoList=&ent=b&_imgSH=fourCardType&specialGoodsType=&_isFuzzy=0&_spAttL=&_mAttL=&_sAttL=&_noAttL=&topMAttL=&topSAttL=&topNoAttL=&hotKeyType=0".format(page)
        response = requests.get(url, headers = headers)
        soup1 = BeautifulSoup(response.text, "html.parser")
        soup2 = soup1.find_all("li", class_ = "goodsItemLi")
        hrefs = [tag.find("a").get("href") for tag in soup2]
        returnlist += hrefs
    return returnlist

def get_product_info(href):
    """
    輸入冰箱的產品頁面連結
    輸出[產品連結、產品名稱、價格、品牌名稱、類型、容量、效能]
    """
    headers = {"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/97.0.4692.71 Safari/537.36"}
    base_url = "https://m.momoshop.com.tw/"
    product_url = base_url + href  # 產品連結
    response = requests.get(product_url, headers = headers)
    soup = BeautifulSoup(response.text, "html.parser")
    product_name = soup.find("p", class_ = "fprdTitle").text  # 產品名稱
    product_price = soup.find("p", class_ = "priceTxtArea").find("b").text  # 價格
    
    # 如果找不到品牌名稱，就直接回傳。因為沒有品牌名稱代表當前頁面不是冰箱的產品頁面，可以直接跳過這個商品。
    if not soup.find("td", class_ = "brandNameMode"):
        return
    else:
        brand_name = soup.find("td", class_ = "brandNameMode").find("a").get("title")  # 品牌名稱
    
    soup1 = soup.find("div", class_ = "attributesArea").find_all("tr")
    soup1 = [s for s in soup1 if s.find("th")]
    attributes = [tag.find("th").text for tag in soup1]
    # 如果商品頁面內沒有這些資料就寫"無"(有些商品頁面內沒有這幾項資訊)
    if "類型" not in attributes:
        category = "無"
    if "容量" not in attributes:
        volume = "無"
    if "效能" not in attributes:
        efficiency = "無"
    for i in soup1:
        if i.find("th").text == "類型":
            content = i.find_all("li")
            if len(content) > 1:
                category = "、".join([c.text for c in content])  # 類型
            else:
                category = content[0].text

        if i.find("th").text == "容量":
            volume = i.find("li").text  # 容量

        if i.find("th").text == "效能":
            efficiency = i.find("li").text  # 效能

    return [product_url, product_name, product_price, brand_name, category, volume, efficiency]


hrefs = get_all_href(1, 18)
# print(len(hrefs))
all_product_infos = []
for href in hrefs:
    all_product_infos.append(get_product_info(href))

# 把沒有抓到冰箱資訊的那幾筆資料去除
all_product_infos = [product_info for product_info in all_product_infos if product_info]
# print(len(all_product_infos))

# 把價格的資料型態從字串(str)轉成數字(int)，並把價格資料裡面的逗號去掉
for i in all_product_infos:
    i[2] = int(i[2].replace(",", ""))


# 和資料庫連線，並創建資料表(table)

# 資料庫設定(refrigerator_database是已經在資料庫中創建好的database(schema))
db_settings = {
    "host":"127.0.0.1",
    "port":3306,
    "user":"root",
    "password":"123456",
    "db":"refrigerator_database",
    "charset":"utf8"
}

# 建立connect物件，用來讓python可以操作MySQL資料庫
connect = pymysql.connect(**db_settings)

with connect.cursor() as cursor:
    # SQL指令
    sql = """
    create table refrigerator_data
    (product_url varchar(255),
    product_name varchar(255),
    price int(10),
    brand_name varchar(20),
    category varchar(20),
    volume varchar(10),
    efficiency char(8));
    """
    
    cursor.execute(sql)
    connect.commit()
connect.close()


# 把資料傳入資料庫
connect = pymysql.connect(**db_settings)

with connect.cursor() as cursor:
    sql = """
    insert into refrigerator_data(product_url, product_name, price, brand_name, category, volume, efficiency)
    values(%s, %s, %s, %s, %s, %s, %s);
    """
    
    for data in all_product_infos:
        cursor.execute(sql, tuple(data))
    connect.commit()
connect.close()