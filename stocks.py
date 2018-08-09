# -*- coding: utf-8 -*-
from flask import Flask
from MyStocks import *
from pyWord import *
from myEmail import *
import json

app = Flask(__name__)


@app.route('/getdata')
def get_stocks():
    try:
        print('starting request data ...')
        stocks_hk = ['00688', '02009', '01313', '03323', '00914']
        results = {}
        t = MyStocks()
        # stocks_hk = ['00688']
        company = {}
        for stock_hk in stocks_hk:
            res = t.request_stocks(stock_hk)
            result_t = t.get_needs(res)
            name = 'H' + stock_hk
            company[name] = result_t
        # results['company'] = company

        h = HangQing()
        res_h1 = h.request1(0)
        result_h1 = h.get_needs(res_h1, 1)
        results['szzs'] = result_h1

        res_h2 = h.request1(1)
        result_h2 = h.get_needs(res_h2, 1)
        results['szcz'] = result_h2

        res_h3 = h.request2()
        result_h3 = h.get_needs(res_h3, 2)
        results['hszs'] = result_h3

        stocks_a = ['sh601992', 'sz000401', 'sz000856', 'sh600585', 'sz000002']
        for stock_a in stocks_a:
            res = h.request1(-1, stock_a)
            result_h = h.get_needs_a(res)
            company[stock_a] = result_h
        results['company'] = company

        result_json = json.dumps(results, ensure_ascii=False)
        print('request data complete')
        print(result_json)

        myword = MyWord()
        filepath = myword.generate(result_json)

        email = MyEmail()
        info = email.send_email(filepath)
        return info
    except:
        return '文件生成失败，自己复制粘贴去吧'


@app.route('/')
def get():
    result = {
    "szzs":{
        "nowpri":"2873.5938",
        "increase":"-8.6316",
        "increPer":"-0.30"
    },
    "szcz":{
        "nowpri":"9295.930",
        "increase":"-56.321",
        "increPer":"-0.60"
    },
    "hszs":{
        "nowpri":"28804.281",
        "increase":"23.140",
        "increPer":"0.080"
    },
    "company":{
        "H00688":{
            "code":"00688",
            "stockName":"中国海外发展",
            "currentPrice":"25.100",
            "changeAmount":"+0.150",
            "priceChangeRatio":"+0.60%",
            "close":"24.950",
            "open":"24.950",
            "maxPrice":"25.200",
            "minPrice":"24.550",
            "volume":"1085.30万",
            "amount":"2.71亿",
            "marketCapitalization":"2750亿"
        },
        "H02009":{
            "code":"02009",
            "stockName":"金隅集团",
            "currentPrice":"3.060",
            "changeAmount":"+0.130",
            "priceChangeRatio":"+4.44%",
            "close":"2.930",
            "open":"2.940",
            "maxPrice":"3.060",
            "minPrice":"2.910",
            "volume":"2759.05万",
            "amount":"8280.47万",
            "marketCapitalization":"327亿"
        },
        "H01313":{
            "code":"01313",
            "stockName":"华润水泥控股",
            "currentPrice":"9.200",
            "changeAmount":"+0.310",
            "priceChangeRatio":"+3.49%",
            "close":"8.890",
            "open":"8.890",
            "maxPrice":"9.230",
            "minPrice":"8.760",
            "volume":"2868.92万",
            "amount":"2.61亿",
            "marketCapitalization":"642亿"
        },
        "H03323":{
            "code":"03323",
            "stockName":"中国建材",
            "currentPrice":"8.550",
            "changeAmount":"+0.340",
            "priceChangeRatio":"+4.14%",
            "close":"8.210",
            "open":"8.210",
            "maxPrice":"8.600",
            "minPrice":"8.150",
            "volume":"4496.89万",
            "amount":"3.79亿",
            "marketCapitalization":"721亿"
        },
        "H00914":{
            "code":"00914",
            "stockName":"海螺水泥",
            "currentPrice":"49.100",
            "changeAmount":"+1.050",
            "priceChangeRatio":"+2.19%",
            "close":"48.050",
            "open":"47.800",
            "maxPrice":"49.400",
            "minPrice":"47.500",
            "volume":"1413.78万",
            "amount":"6.87亿",
            "marketCapitalization":"2602亿"
        },
        "sh601992":{
            "code":"sh601992",
            "stockName":"金隅集团",
            "currentPrice":"3.730",
            "changeAmount":"-0.010",
            "priceChangeRatio":"-0.27",
            "close":"3.740",
            "open":"3.740",
            "maxPrice":"3.770",
            "minPrice":"3.720",
            "volume":"344542",
            "amount":"129309735.000"
        },
        "sz000401":{
            "code":"sz000401",
            "stockName":"冀东水泥",
            "currentPrice":"11.170",
            "changeAmount":"-0.23",
            "priceChangeRatio":"-2.02",
            "close":"11.400",
            "open":"11.480",
            "maxPrice":"11.500",
            "minPrice":"11.160",
            "volume":"419440",
            "amount":"472984986.900"
        },
        "sz000856":{
            "code":"sz000856",
            "stockName":"冀东装备",
            "currentPrice":"13.280",
            "changeAmount":"-0.38",
            "priceChangeRatio":"-2.78",
            "close":"13.660",
            "open":"13.490",
            "maxPrice":"13.640",
            "minPrice":"13.230",
            "volume":"80787",
            "amount":"108004228.100"
        },
        "sh600585":{
            "code":"sh600585",
            "stockName":"海螺水泥",
            "currentPrice":"37.680",
            "changeAmount":"0.430",
            "priceChangeRatio":"1.15",
            "close":"37.250",
            "open":"37.400",
            "maxPrice":"37.960",
            "minPrice":"36.920",
            "volume":"332010",
            "amount":"1245699431.000"
        },
        "sz000002":{
            "code":"sz000002",
            "stockName":"万 科Ａ",
            "currentPrice":"23.320",
            "changeAmount":"-0.39",
            "priceChangeRatio":"-1.64",
            "close":"23.710",
            "open":"23.550",
            "maxPrice":"23.810",
            "minPrice":"23.240",
            "volume":"322827",
            "amount":"757327314.910"
        }
    }
}
    result_json = json.dumps(result, ensure_ascii=False)
    myword = MyWord()
    filepath = myword.generate(result_json)
    print("filepath:")
    print(filepath)
    email = MyEmail()
    info = email.send_email(filepath)
    return info


