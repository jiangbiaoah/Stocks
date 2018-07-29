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
        for stock in stocks_hk:
            res = t.request_stocks(stock)
            result_t = t.get_needs(res)
            name = 'H' + stock
            company[name] = result_t
        results['company'] = company

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
    result = {"hszs": {"increPer": "-0.480", "nowpri": "28781.139", "increase": "-139.760"}, "company": {
        "H01313": {"currentPrice": "8.890", "volume": "2165.53万", "changeAmount": "+0.140",
                   "priceChangeRatio": "+1.60%", "close": "8.750", "code": "01313", "amount": "1.92亿", "open": "8.880",
                   "minPrice": "8.760", "stockName": "华润水泥控股", "marketCapitalization": "621亿", "maxPrice": "9.050"},
        "H00688": {"currentPrice": "24.950", "volume": "1594.79万", "changeAmount": "-0.050",
                   "priceChangeRatio": "-0.20%", "close": "25.000", "code": "00688", "amount": "3.98亿",
                   "open": "25.050", "minPrice": "24.750", "stockName": "中国海外发展", "marketCapitalization": "2734亿",
                   "maxPrice": "25.350"},
        "H00914": {"currentPrice": "48.050", "volume": "1232.76万", "changeAmount": "-0.850",
                   "priceChangeRatio": "-1.74%", "close": "48.900", "code": "00914", "amount": "5.98亿",
                   "open": "49.250", "minPrice": "47.850", "stockName": "海螺水泥", "marketCapitalization": "2546亿",
                   "maxPrice": "49.700"},
        "H03323": {"currentPrice": "8.210", "volume": "2368.57万", "changeAmount": "-0.040",
                   "priceChangeRatio": "-0.48%", "close": "8.250", "code": "03323", "amount": "1.95亿", "open": "8.320",
                   "minPrice": "8.090", "stockName": "中国建材", "marketCapitalization": "692亿", "maxPrice": "8.480"},
        "H02009": {"currentPrice": "2.930", "volume": "4364.95万", "changeAmount": "-0.010",
                   "priceChangeRatio": "-0.34%", "close": "2.940", "code": "02009", "amount": "1.28亿", "open": "2.940",
                   "minPrice": "2.900", "stockName": "金隅集团", "marketCapitalization": "313亿", "maxPrice": "2.970"}},
              "szzs": {"increPer": "-0.74", "nowpri": "2882.2254", "increase": "-21.4213"},
              "szcz": {"increPer": "-1.18", "nowpri": "9352.251", "increase": "-111.509"}}
    result_json = json.dumps(result, ensure_ascii=False)
    myword = MyWord()
    filepath = myword.generate(result_json)
    print("filepath:")
    print(filepath)
    email = MyEmail()
    info = email.send_email(filepath)
    return info
