import os

import requests
import re
from lxml import etree
import pandas as pd

headers = {
    'Host': 'news.tju.edu.cn',
    'Accept-Encoding': 'gzip, deflate',
    'Accept-Language': 'zh-CN,zh;q=0.9,en-US;q=0.8,en;q=0.7,zh-TW;q=0.6',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.108 Safari/537.36'
}
base_url = "http://lanunion.cqu.edu.cn/repair/admin"

def login(username, passwd):
    login_url = base_url + "/index.php/user/loginpost"
    login_value = {
        "username": username,
        "password": passwd
    }
    res = s.post(login_url, headers=headers, data=login_value)
    html = etree.HTML(res.content.decode('utf-8'))
    name = html.xpath('//*[@id="logo"]/div[2]//text()')
    print("欢迎, " + name[1])


def get_integration(begintime, endtime):
    time_submit_url = base_url + "/index.php/credit/setcondition"
    integration_value = {
        "time_begin": begintime,
        "time_end": endtime,
        "username": '',
        "issum": '1'
    }
    res = s.post(time_submit_url, headers=headers,data=integration_value)
    html = etree.HTML(res.content.decode('utf-8'))
    total_pages = html.xpath('//div[@class="jr_pager_indent"]//text()')
    total_pages = int(re.findall('\d+', str(total_pages))[-1])
    # print(total_pages)
    integration_changes = []
    for i in range(1, total_pages+1):
        # print(i)
        integration_url = base_url + "/index.php/credit/index/" + str(i)
        res = s.get(integration_url, headers=headers)
        html = etree.HTML(res.content.decode('utf-8'))

        stuId = html.xpath('//*[@id="credit"]/tbody/tr/td[2]/text()')
        stuName = html.xpath('//*[@id="credit"]/tbody/tr/td[3]/text()')
        changes = html.xpath('//*[@id="credit"]/tbody/tr/td[4]/text()')

        for i, item in enumerate(stuId):
            integration_changes.append({
                "学号": str(stuId[i]),
                "姓名": str(stuName[i]),
                "积分变化": str(changes[i])
            })
    # print(integration_changes)
    return integration_changes

if __name__ == '__main__':
    s = requests.Session()
    username = input("请输入学号: ")
    passwd = input("请输入密码: ")
    login(username, passwd)
    begintime = input("请输入查询开始时间(如'2019-11-01'): ")
    endtime = input("请输入查询截止时间(如'2019-11-30'): ")
    integration_changes = get_integration(begintime, endtime)

    result = []
    for i in range(0, len(integration_changes)):
        result.append(list(integration_changes[i].values()))

    temp = begintime.split("-")
    begintime = ""
    for i in range(len(temp)):
        begintime += temp[i]
    temp = endtime.split("-")
    endtime = ""
    for i in range(len(temp)):
        endtime += temp[i]

    xlsx_name = begintime + "-" + endtime + "积分变动.xlsx"

    df1 = pd.DataFrame(result)
    df1.to_excel(xlsx_name)
    currentPath = os.getcwd()
    print(xlsx_name + ", 已保存至" + currentPath)