# 抓取所有設計師資料
# 設計師名稱、任職髮廊、髮廊位置、評論則數、設計師作品數、設計師追蹤數


import requests
from bs4 import BeautifulSoup
import json
import pandas as pd


def get_designers_info(text):
    """
    輸入網頁原始碼(json檔的格式)，把它轉換成字典型態，然後從裡面提取需要的資訊，
    輸出包含多個[設計師名稱、任職髮廊、髮廊位置、評論則數、設計師作品數、設計師追蹤者]的列表
    """
    returnlist = []
    designers_dict = json.loads(text)["data"]["userlist"]
    for d in designers_dict:
        designer_info = [d["name"], d["wpInfo"]["wp_name"], d["wpInfo"]["wp_address"], int(d["professionInfo"]["reviewNum"]), int(d["num_post"]), int(d["num_followers"])]
        returnlist.append(designer_info)
    return returnlist


url = "https://style-map.com/api/user/filter"
headers = {"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/97.0.4692.71 Safari/537.36"}
start, end = 0, 100
all_designers_info = []
while end <= 1400:
    payload = {"start":start, "end":end}
    response = requests.post(url, data = payload, headers = headers)
    all_designers_info += get_designers_info(response.text)
    start += 100
    end += 100
# print(len(all_designers_info))

# 把資料整理成DataFrame格式
df = pd.DataFrame(all_designers_info)
df.columns = ["設計師名稱", "任職髮廊", "髮廊位置", "評論則數", "設計師作品數", "設計師追蹤數"]
# print(df)

# 輸出成excel檔
df.to_excel("設計師資料.xlsx", index = False)