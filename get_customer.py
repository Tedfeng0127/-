import requests          # 用來發出請求的模組
import json              # 操作json檔案的模組，可以把字典型態(dict)的資料轉成json檔案

# 原理概論：
# 我們(client)發出request給華南銀行的伺服器(server)，
# 華南銀行的server會辨別我們的request(用我們輸入的url來辨別)，透過我們給的參數(uuid、api_key等等)來辨識我們需要什麼回應(response)，
# response就是伺服器回傳給我們的資料，只是有各種不同形式，有時候是網頁，有時候是pdf，
# 這裡的response是json檔案。
# (json檔案是網路傳輸資料的一種標準格式，因為資料規則簡單，需要的容量少，所以常常用來傳遞大量資料。)


# 發出request，透過API直接得到資料。
response = requests.get(url = 'https://www.fintechersapi.com/bank/huanan/getUUIDs?api_key=cef5e50c-c6df-46c0-86a8-69fa2fdc0fe1', headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'
                                 'AppleWebKit/537.36 (KHTML, like Gecko)'
                                 'Chrome/75.0.3770.100 Safari/537.36'})
# 上面的response大概長這樣
# {
#   "uuid_list": [
#     "2bef6593954b43bcb6bc93260e4ec25f",
#     "b66410d666b946989c4d42423041edbd",
#     "eb28013699de437f97bf7d7195abe2b9",
#     "e8e8a73a6a244c42880128a9e0bdb060",
#     "2b9ddc16bd0048c9935528ed9c14a697",
#     "1827e418bbe746a18c6a2701ba1fbc16",
#     "0747308e74ca4aa3bc4a4d498d52c557",
#     "fc05790459ca442db697d191f4bc245f",
#     "9a2c250157354020a0f31de622f15713",
#     "184f70b3fb054157b7e797f1e7f68035"
#   ]
# }


                                
# uuid_list，一次10筆
# 用json.loads把response.text的內容轉換成python內的dict資料型態
# response.text是把回應轉化成text(可以被轉化成python的檔案)
response = json.loads(response.text)        
# print(response)
# print(type(response))

# 銀行客戶帳戶API範例，根據這個結構進行解析
# {
#   "trans_record": [
#     {
#       "account_num": "160345191470",
#       "trans_date": "2015-10-01",
#       "amount": 2000,
#       "balance": 4282299,
#       "trans_channel": "臨櫃",
#       "id": 25827,
#       "trans_type": "台幣轉帳"
#     }
#   ],
#   "uuid": "833ecc53b7064974a0abaeb986da9c7a"
# }
while True:
    i = 1        # 記數
    for uuid in response['uuid_list']:
        # 保險資料取得
        resp_insurance = requests.get(url = 'https://www.fintechersapi.com/bank/huanan/insurance/record?uuid={}&api_key=cef5e50c-c6df-46c0-86a8-69fa2fdc0fe1'.format(uuid), headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'            
                            'AppleWebKit/537.36 (KHTML, like Gecko)'
                            'Chrome/75.0.3770.100 Safari/537.36'})
        # input(resp_insurance.encoding)，得到ISO-8859-1編碼
        # 處理編碼問題
        resp_insurance = resp_insurance.text.encode('ISO-8859-1').decode('utf-8')
        
        # 客戶基本資料取得
        resp_client = requests.get(url = 'https://www.fintechersapi.com/bank/huanan/digitalfin/customers?uuid={}&api_key=cef5e50c-c6df-46c0-86a8-69fa2fdc0fe1'.format(uuid), headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'            
                            'AppleWebKit/537.36 (KHTML, like Gecko)'
                            'Chrome/75.0.3770.100 Safari/537.36'})
        # 客戶近10筆交易資料取得
        resp_recent_10 = requests.get(url = 'https://www.fintechersapi.com/bank/huanan/digitalfin/account_records?uuid={}&api_key=cef5e50c-c6df-46c0-86a8-69fa2fdc0fe1'.format(uuid), headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'
                                    'AppleWebKit/537.36 (KHTML, like Gecko)'
                                    'Chrome/75.0.3770.100 Safari/537.36'})                                  
        print('第{}筆資料'.format(i))
        i += 1
        print(f'uuid: {uuid}')
        # json轉成dict
        resp_insurance = json.loads(resp_insurance)
        resp_client = json.loads(resp_client.text)
        resp_recent_10 = json.loads(resp_recent_10.text)
        
        print('保險資料如下: \n', resp_insurance)                 # 直接印出dict型態
        for column in resp_insurance['insurance_record'][0]:     # 分解
            print(column, resp_insurance['insurance_record'][0][column], sep = '\t')
        
        print('客戶資料如下: \n', resp_client)                    # dict型態
        for column in resp_client:                               # 分解
            print(column, resp_client[column], sep = '\t')

        print('客戶近10筆交易資料如下: \n', resp_recent_10)
        for column in resp_recent_10['trans_record'][0]:
            print(column, resp_recent_10['trans_record'][0][column], sep = '\t')

        answer = input('下一筆請按enter，離開則隨便輸入再按enter: ')
        if answer == '':
            continue
        else:
            break

    leave_or_not_text = input('按enter繼續下個十筆，其於輸入離開程式')
    if leave_or_not_text == '':
        continue
    else:
        break