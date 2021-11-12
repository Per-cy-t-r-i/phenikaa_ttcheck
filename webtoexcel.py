import requests
from bs4 import BeautifulSoup
from html2excel import ExcelParser


LOGIN_URL="https://qldt.phenikaa-uni.edu.vn/Login.aspx"
WANTED_URL="https://qldt.phenikaa-uni.edu.vn/wfrmDangKyLopTinChiB3.aspx"
EXCEL_OUT = "tkbtheoweb.xlsx"
test_file= 'test.html'
acc = 'account.txt'

with open(acc, 'r') as f:
    username=f.readline().replace("\n","")
    password=f.readline()
    f.close()


s=requests.Session()
r=s.get(LOGIN_URL)

soup=BeautifulSoup(r.content,'html.parser')

VIEWSTATE=soup.find(id="__VIEWSTATE")['value']
EVENTVALIDATION=soup.find(id="__EVENTVALIDATION")['value']
VIEWSTATEGENERATOR=soup.find(id="__VIEWSTATEGENERATOR")['value']

login_data={"__VIEWSTATE":VIEWSTATE,
"txtusername":username,
"txtpassword":password,
"__VIEWSTATE":VIEWSTATE,
"__EVENTVALIDATION":EVENTVALIDATION,
"__VIEWSTATEGENERATOR":VIEWSTATEGENERATOR,
"btnDangNhap":""
}

s.post(LOGIN_URL, data=login_data)
open_page = s.get(WANTED_URL)
if open_page.url == LOGIN_URL:
    print('Login failed')
else:
    soup = BeautifulSoup(open_page.content, 'html.parser')

    table = soup.find(id='grdViewLopDangKy')

    html_content = table.prettify()
    with open(test_file, 'w', encoding = 'utf-8') as f:
        f.write(html_content)
        f.close()
    parser = ExcelParser(test_file)
    parser.to_excel(EXCEL_OUT)
