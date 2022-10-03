import requests
import re
import gspread
from datetime import datetime
import pytz
import math
import time

currency_dic = {"GBP": "英镑", "HKD": "港币", "USD": "美元", "CHF": "瑞士法郎", "SGD": "新加坡元", "PKR": "巴基斯坦卢比", "SEK": "瑞典克朗", "DKK": "丹麦克朗", "NOK": "挪威克朗", "JPY": "日元", "CAD": "加拿大元", "AUD": "澳大利亚元", "MYR": "林吉特", "EUR": "欧元", "RUB": "卢布", "MOP": "澳门元", "THB": "泰国铢", "NZD": "新西兰元", "ZAR": "南非兰特", "KZT": "哈萨克斯坦 坚戈", "KRW": "韩元"}

def get_rate():
    source = requests.get("http://www.icbc.com.cn/ICBCDynamicSite/Optimize/Quotation/QuotationListIframe.aspx").text

    currency_rate = {}

    for currency in currency_dic:
        currency_rate[currency] = re.findall(r'(?<=width="14%">).*?(?=</td>)', re.search(r'(?<=' + currency_dic[currency] + '\(' + currency + '\)</td><td class="tdCommonTableData" align="right" valign="middle").*(?=</td>)', source).group(0))

    return currency_rate

def current_time():
    now = datetime.now(pytz.timezone('Asia/Shanghai'))
    return now.strftime("%m/%d/%Y, %H:%M:%S")

def write_rate(rate):
    gc = gspread.oauth()
    sh = gc.open("Foreign Currency Control Sheet")
    ws = sh.worksheet("ICBC Buy Rate")

    row = 2

    while ws.cell(row, 1).value != None:
        try:
            write_content = rate[ws.cell(row, 1).value][0]
        except:
            write_content = "NOT FOUND"
        ws.update_cell(row, 2, write_content)
        row += 1
    
    ws.update_cell(1, 5, current_time())

def sleep(mins):
    now = time.time()
    time.sleep((math.floor(now / 60 / mins) + 1) * 60 * mins - now)

def main():
    while True:
        rate = get_rate()
        write_rate(rate)
        sleep(5)

if __name__ == "__main__":
    main()