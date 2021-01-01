import requests
from bs4 import BeautifulSoup
import os
import xlwt
import re


class Get_HKCPA():
    def __init__(self):
        self.url = r'https://www.hkicpa.org.hk/en/Membership/Find-a-CPA/Membership-List'
        self.params = {
            'pg': 1,
            'fnsn': '',
            'gn': '',
            'mneq': '',
            'fneq': '',
            'mbst': '',
            'sdhd': '',
            'pceq': ''
        }

    def get_params(self, page):
        params = {
            'pg': page,
            'fnsn': '',
            'gn': '',
            'mneq': '',
            'fneq': '',
            'mbst': '',
            'sdhd': '',
            'pceq': ''
        }
        return params

    def write2excel(self, member_list, path, filename):
        if os.path.exists(path):
            os.chdir(path)
        else:
            os.mkdir(path)
            os.chdir(path)

        title = ["Name", "Membership No.", "Practising Member", "Name on the PC", "PC No.", "SD (Insolvency) Holder"]
        book = xlwt.Workbook()  # 创建一个excel对象
        sheet = book.add_sheet('Sheet1', cell_overwrite_ok=True)  # 添加一个sheet页
        for i in range(len(title)):  # 循环列
            sheet.write(0, i, title[i])  # 将title数组中的字段写入到0行i列中
        t = 0
        for item in member_list:
            sheet.write(1 + t, 0, item['Name'])
            sheet.write(1 + t, 1, item['Membership No.'])
            sheet.write(1 + t, 2, item['Practising Member'])
            sheet.write(1 + t, 3, item['Name on the PC'])
            sheet.write(1 + t, 4, item['PC No.'])
            sheet.write(1 + t, 5, item['SD (Insolvency) Holder'])

            t += 1
        book.save(filename + '.xls')  # 保存excel

    def get_pages(self):
        num = []
        html = requests.get(url=self.url, params=self.params, verify=False)
        soup = BeautifulSoup(html.text, 'html.parser')
        pages = soup.find_all('a', href=re.compile('pg=\d+&fnsn=&gn=&mneq=&fneq=&mbst=&sdhd=&pceq='))
        for page in pages:
            page = page.string
            if page:
                page = page.replace(' ', '')
                page = int(page.replace('\n', ''))
                num.append(page)
        return max(num)

    def main(self, path, filename, allpages,startpage,endpage):
        requests.packages.urllib3.disable_warnings()
        if allpages == 'yes':
            print('\n查找总页数......')
            page_value = self.get_pages()
            startpagenum=1
            print('\n总共有：', page_value, '页')
        else:
            page_value = endpage
            startpagenum=startpage
            print('\n总共有：', endpage-startpagenum+1, '页')
        print('---------------------')
        member_list = []
        for page in range(startpagenum, page_value + 1):
            print('正在下载第', page, '页......')
            params = self.get_params(page)
            html = requests.get(url=self.url, params=params, verify=False)
            soup = BeautifulSoup(html.text, 'html.parser')
            table = soup.find_all('table', class_='responsive')
            for idx, tr in enumerate(table[0].find_all('tr')):
                if idx != 0:
                    tds = tr.find_all('td')
                    member_list.append({
                        'Name': tds[0].contents[0].replace('\n', '').strip(),
                        'Membership No.': tds[1].contents[0],
                        'Practising Member': tds[2].contents[0],
                        'Name on the PC': tds[3].contents[0],
                        'PC No.': tds[4].contents[0],
                        'SD (Insolvency) Holder': tds[5].contents[0]
                    })
        self.write2excel(member_list, path, filename)
        print('\n-----已成功下载------')


if __name__ == '__main__':
    spider = Get_HKCPA()
    filename = 'HKCPA_Memberslist'
    path = str(input("请输入保存路径(默认：c:\\hkcpa)："))
    if path == '' :
       path = r'c:\\hkcpa'
    allpages = str(input("全部下载请输入‘yes’："))
    if allpages != 'yes':
        startpage = int(input('请输入下载起始页：'))
        endpage = int(input('请输入下载结束页：'))
    else:
        allpages = 'yes'
        startpage = 1
        endpage = 2
    spider.main(path,filename,allpages,startpage,endpage)
    # page = 'yes'，下载全部；否则按 page 下载
