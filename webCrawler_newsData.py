# 抓取 2017/01/01 - 2020/10/01 鉅亨網台股新聞的資訊
# 標題、發佈時間、新聞類別、關鍵字、內文

import requests
from bs4 import BeautifulSoup
import json
import time
import datetime
import pandas as pd


def get_last_page_number(url):
    """
    拿到特定時間區段內所有新聞原始碼的總頁面數
    input:網址
    output:總頁數的值
    """
    headers = {"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
                "AppleWebKit/537.36 (KHTML, like Gecko)"
                "Chrome/86.0.4240.75 Safari/537.36"}
    response = requests.get(url, headers)
    srccode1 = json.loads(response.text)
    return srccode1["items"]["last_page"]

def get_html(url):
    """
    抓取某時間區間內一個頁面的原始碼
    input:包含30篇新聞的一個原始碼頁面(利用for迴圈替換頁數，去疊代所有新聞原始碼的頁面)
    output:動態網頁的json檔
    """
    headers = {"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
                "AppleWebKit/537.36 (KHTML, like Gecko)"
                "Chrome/86.0.4240.75 Safari/537.36"}
    response = requests.get(url, headers)
    response.raise_for_status
    response.encoding = response.apparent_encoding
    return response.text

def get_json(text):
    """
    吃json檔，然後轉成字典型態
    """
    items = json.loads(text)
    return items

def get_newsId(items):
    """
    抓取新聞Id跟這個Id對應到的新聞標題
    input:一個頁面的json檔(已經轉成字典)
    output:字典(key=新聞Id，value=新聞標題)
    """
    returndict = {}
    news = items["items"]["data"]
    for i in news:
        returndict[i["newsId"]] = i["title"]
    return returndict

def crawl(newsId):
    """
    替換url內的新聞Id進入各篇新聞頁面抓取資料
    input:包含新聞Id跟標題的字典
    output:很多篇新聞組成的列表
    """
    returnlist = []
    for Id in newsId:
        newsdict = {}
        url = "https://news.cnyes.com/news/id/%s?exp=a" % (Id)
        response = requests.get(url)
        soup = BeautifulSoup(response.text, "html.parser")
        time = soup.find("time").get_text()
        keywords = []
        if soup.find("nav", class_ = "_1qS9 _2Zhy"):
            keywords_list = soup.find("nav", class_ = "_1qS9 _2Zhy").find_all("span", class_ = "_1E-R")
            for keyword in keywords_list:
                keywords.append(keyword.get_text())
        else:
            keywords = ["None"]
        content = soup.find("div", class_ = "_2E8y").get_text()
        category = soup.find_all("span", itemprop = "item")[-1].get_text()
        newsdict["title"] = newsId[Id]
        newsdict["time"] = time
        newsdict["category"] = category
        newsdict["keywords"] = str(keywords)
        newsdict["content"] = content
        returnlist.append(newsdict)
    return returnlist

def save_json(many_news, count):
    """
    把抓到的新聞內容存入字典，再存入列表，並寫入json檔
    利用參數count區別不同的新聞檔案
    """
    with open(r"project\news\news%s.json" % (count), "w", encoding = "utf8") as f:
        json.dump(many_news, f, ensure_ascii = False)

def save_csv(many_news, count):
    """
    把抓到的新聞內容存進DataFrame，再存成csv檔
    利用參數count區別不同新聞檔案
    """
    df = pd.DataFrame(many_news)
    df.to_csv(r"project\news\news%s.csv" % (count), index = False)

def crawl_news(start, timespan):
    """
    利用前面的函式完成新聞資料的抓取
    start:開始爬取的時間點(Ex: "2017 Jan 01 00:00:00")
    timespan:要爬取的月份數
    """
    # start = "2017 Oct 15 00:00:00"                                             # 資料區間開始時間
    start_time = int(time.mktime(time.strptime(start, "%Y %b %d %H:%M:%S")))   # 轉成秒數
    # end = "2020 Oct 15 00:00:00"                                               # 資料區間結束時間
    # end_time = int(time.mktime(time.strptime(end, "%Y %b %d %H:%M:%S")))       # 轉成秒數
    step = 2592000                                                             # 一次抓取範圍的秒數(30天的秒數)

    count = 10
    while count <= timespan:
        try:
            many_news = []
            a, b = start_time, start_time + step
            base_url = "https://api.cnyes.com/media/api/v1/newslist/category/tw_stock?startAt=%s&endAt=%s&limit=30" % (a, b)
            last_page = get_last_page_number(base_url)
            for page in range(1, last_page + 1):
                url = base_url + "&page=%s" % (page)
                text = get_html(url)
                items = get_json(text)
                newsId = get_newsId(items)
                news = crawl(newsId)
                many_news += news
            # save_json(many_news, count)  # 存json檔
            save_csv(many_news, count)     # 存csv檔
            start_time += step
            print(count)
            count += 1
        except Exception as e:
            print(e)
            break

crawl_news("2019 Oct 01 00:00:00", 12)