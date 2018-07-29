# -*- coding: utf-8 -*-
import json
import re
from urllib.parse import urlencode
from urllib.request import urlopen


# ----------------------------------
# HK:00688,02009,01313,03323,00914
# 600026,000002,600585,000856,601992,000401
# 601992,000401,000856,600585,000002
# 02009,00914,03323,01313,00688
# ----------------------------------
class MyStocks:
    appkey = "9808007c691a6f944c6edeaad2a829db"

    # 股票查询
    def request_stocks(self, stock='', m="GET"):
        url = "http://op.juhe.cn/onebox/stock/query"
        params = {
            "key": self.appkey,  # 应用APPKEY(应用详细页查询)
            "dtype": "",  # 返回数据的格式,xml或json，默认json
            "stock": stock,  # 股票名称
        }
        params = urlencode(params)
        if m == "GET":
            f = urlopen("%s?%s" % (url, params))
        else:
            f = urlopen(url, params)

        content = f.read()
        # print("content:"+content)
        res = json.loads(content.decode('utf-8'))
        if res:
            error_code = res["error_code"]
            if error_code == 0:
                # 成功请求
                # print(res["result"])
                print("Request Successful")
            else:
                print("%s:%s" % (res["error_code"], res["reason"]))
        else:
            print("request api error")
        return res

    # 从请求结果中提取需要的数据
    def get_needs(self, res):
        if res["error_code"] != 0:
            return
        result = {}
        pattern = re.compile(r'\d+')
        code = pattern.findall(res["result"]["title"])[0]
        result['code'] =code     # 股票代码*
        result['stockName'] = res["result"]["stockName"]    # 股票名称*
        result['currentPrice'] = res["result"]["currentPrice"]  # 最新价*
        result['changeAmount'] =res["result"]["changeAmount"]   # 涨跌额*
        result['priceChangeRatio'] =res["result"]["priceChangeRatio"]   # 涨跌幅*
        result['close'] = res["result"]["close"]   # 昨收
        result['open'] = res["result"]["open"]     # 今开
        result['maxPrice'] = res["result"]["maxPrice"] # 最高
        result['minPrice'] = res["result"]["minPrice"] # 最低
        result['volume'] = res["result"]["volume"]     # 成交量
        result['amount'] =res["result"]["amount"]      # 成交额*
        result['marketCapitalization'] =str(round(float(res["result"]["marketCapitalization"])/100000000)) + "亿"  # 市值*
        # self.result['turnOverRate'] = res["result"]["turnOverRate"] # 换手率*
        return result


class HangQing():
    appkey = "bb79e2c25150abd9b4cf5a0b9a806185"

    # 1.沪深股市
    def request1(self, type=0, m="GET"):
        url = "http://web.juhe.cn:8080/finance/stock/hs"
        params = {
            "gid": "",  # 股票编号，上海股市以sh开头，深圳股市以sz开头如：sh601009
            "key": self.appkey,  # APP Key
            "type": type,  # 0代表上证指数，1代表深证指数

        }
        params = urlencode(params)
        if m == "GET":
            f = urlopen("%s?%s" % (url, params))
        else:
            f = urlopen(url, params)

        content = f.read()
        res = json.loads(content.decode('utf-8'))
        if res:
            error_code = res["error_code"]
            if error_code == 0:
                # 成功请求
                print("Request Successful")
            else:
                print("%s:%s" % (res["error_code"], res["reason"]))
        else:
            print("request api error")
        return res

    # 香港股市
    def request2(self, m="GET"):
        url = "http://web.juhe.cn:8080/finance/stock/hk"
        params = {
            "num": "00001",  # 股票代码，如：00001 为“长江实业”股票代码
            "key": self.appkey,  # APP Key

        }
        params = urlencode(params)
        if m == "GET":
            f = urlopen("%s?%s" % (url, params))
        else:
            f = urlopen(url, params)

        content = f.read()
        res = json.loads(content.decode('utf-8'))
        if res:
            error_code = res["error_code"]
            if error_code == 0:
                # 成功请求
                print("Request Successful")
            else:
                print("%s:%s" % (res["error_code"], res["reason"]))
        else:
            print("request api error")
        return res

    def get_needs(self, res, type=1):
        if res["error_code"] != 0:
            return
        result = {}
        if type == 1:   # 上证指数/深证成指
            result['nowpri'] = res['result']['nowpri']  # 深证成指/上证指数
            result['increase'] = res['result']['increase']  # 涨跌额
            result['increPer'] = res['result']['increPer']  # 涨跌幅
        elif type == 2:           # 恒生指数
            result['nowpri'] = res['result'][0]['hengsheng_data']['lastestpri']  # 恒生指数
            result['increase'] = res['result'][0]['hengsheng_data']['uppic']  # 涨跌额
            result['increPer'] = res['result'][0]['hengsheng_data']['limit']  # 涨跌幅

        return result


def remaindecimal(string, num=2):
    result = str(round(float(string), num))
    return result
