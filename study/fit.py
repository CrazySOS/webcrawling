from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from datetime import date
import openpyxl
import time


# 크롤링할 사이트 및 로딩 대기
driver = webdriver.Chrome("./chromedriver")
driver.get("https://www.naver.com/")
time.sleep(2)

# 검색창 선택, 검색어 선택, 엔터치는 것, 로딩 대기
elem = driver.find_element_by_name("query")
elem.send_keys("Fed")
elem.send_keys(Keys.RETURN)
time.sleep(2)

# 검색 결과가 나온 창에서 최신순(a herf) 를 선택
driver.find_element_by_partial_link_text('최신순').click()
# 또는 driver.find_element_by_link_text('Send InMail').click() 도 사용 가능
time.sleep(2)

# 뉴스 타이틀 수집
res=[]
news_titles = driver.find_elements_by_css_selector(".news_tit")

for i in news_titles:
    (title, link) = i.text, i.get_attribute('href')
    # print(title, link)
    res.append({'title':title, 'link':link}) 
# # 뉴스 하이퍼링크 수집
# for i in news_titles:
#     href = i.get_attribute('href')
#     print(href)


print (res)

wb = openpyxl.load_workbook("new2202.xlsx")
todaynow = date.today()
ws_want = wb.create_sheet(todaynow.strftime('%Y-%m-%d'), index=0)
ws = wb.active
wb.save("new2202.xlsx")

fieldnames = ['title', 'link']
ws.append(['title', 'link'])

for x in res:
    values = (x[k] for k in fieldnames)

    ws.append(values)

wb.save("new2202.xlsx")