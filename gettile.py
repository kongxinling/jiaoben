#检查是否域名能够访问，如果能则返回网站标题，保存在title中。
import requests
from bs4 import BeautifulSoup
from urllib.request import urlopen
import sys,socket
import io
import sys
from lxml import html
import urllib
from urllib import request
import traceback
from urllib import error




def gettitle(url):

    requests.packages.urllib3.disable_warnings()
    req = request.Request(url)
    try:
        re = request.urlopen(req)
        html = urlopen(url)
        # 解析返回包的内容
        #捕获异常，目标标签在网页中缺失
        try:
            soup = BeautifulSoup(html.read(), 'lxml')
            title = soup.title.text
            tfw = open("title.txt", "a")
            tfw.write(str(soup.title.text) + "\n")
            tfw.close()
            ufw = open("url.txt", "a")
            ufw.write(str(re.url) + "\n")
            ufw.close()
            # 要加close不然无法写入
        except AttributeError as e:
            print(url + " " + "no title")
            efw=open("eception.txt","a")
            efw.write(url+" no tile"+"\n")
    except error.HTTPError as e:
        print(e.code)
        efw = open("eception.txt", "a")
        efw.write(url + " "+str(e.code) + "\n")
    except error.URLError as e:
        print(e.reason)
        efw = open("eception.txt", "a")
        efw.write(url + " "+str(e.reason) + "\n")
f=open("http.txt")
#将要检查的域名放入http.txt文档中
for line in f.readlines():
    line = line.strip('\n')
    url = "http://" + line
    gettitle(url)



















