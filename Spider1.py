import urllib.error
import requests
import ssl
import re
from bs4 import BeautifulSoup
from urllib import request
import xlwt

class SpiderTest:
    def __init__(self, baseUrl, itemInfs):
        self.baseUrl = baseUrl
        self.itemInfs = itemInfs
        self.dataList = []

    def gethtml(self,url):  # 网页请求
        head = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.74 Safari/537.36 Edg/99.0.1150.55"
        }
        html = ""
        context = ssl._create_unverified_context()
        try:
            # print(head)
            # print('2')
            response = requests.get(url, headers=head)
            # print('1')
            text = response.status_code
            # print("status:{}".format(text))
            req = request.Request(url, headers=head)
            response2 = request.urlopen(req, context=context)
            html = response2.read().decode("utf-8")
        except urllib.error.URLError as e:
            if hasattr(e, "code"):
                print(e.code)
            if hasattr(e, "reason"):
                print(e.reason)
        return html

    def parseData(self, item):
        data = []
        res = re.findall(eval(self.itemInfs["name"]), item)
        data.append(res)
        # print(res)
        res = re.findall(eval(self.itemInfs["score"]), item)
        data.append(res)
        # print(res)
        res = re.findall(eval(self.itemInfs["time"]), item)
        data.append(res)
        # print(res)
        res = re.findall(eval(self.itemInfs["usefulNum"]), item)
        data.append(res)
        # print(res)
        res = re.findall(eval(self.itemInfs["comment"]), item)
        data.append(res)
        # print(res)
        return data

    def getData(self):
        for i in range(0, 50):
            url = self.baseUrl + str(i * 20) + "&limit=20&status=P&sort=new_score"#我也不知道怎么回事，这样保持不动就可以跑起来，没有问题的话请不要修改它
            html = self.gethtml(url)
#            file = open('douban250.html', 'rb')
#            html = file.read()
            bs = BeautifulSoup(html, "html.parser")
            for item in bs.find_all('div', class_='comment'):
                data = self.parseData(str(item))
                self.dataList.append(data)


    def saveXls(self,fname):
        workbook = xlwt.Workbook(encoding="utf-8")
        worksheet = workbook.add_sheet('sheet1')
        keys = self.itemInfs.keys()
        for i,key in enumerate(keys):
            worksheet.write(0, i, key)
        for i, film in enumerate(self.dataList):
            for j, item in enumerate(film):
                worksheet.write(i+1, j, item)
        workbook.save(fname)

def main():
    itemInfs = {
        "name": '''re.compile(r'<a class="" href=.*>(.*)</a>')''',
        "score": '''re.compile(r'<span class="allstar.* rating" title="(.*)"></span>')''',
        "time": '''re.compile(r'<span class="comment-time" title="(.*)">')''',
        "usefulNum": '''re.compile(r'<span class="votes vote-count">(.*?)</span>')''',
        "comment": '''re.compile(r'<span class="short">(.*)</span>')'''
    }
    test = SpiderTest("https://movie.douban.com/subject/1959877/comments?start=", itemInfs)
    test.getData()
    # print(test.dataList)
    print(test.dataList)
    test.saveXls("text.xls")
    return

main()