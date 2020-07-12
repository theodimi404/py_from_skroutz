from urllib.parse import urlparse
import bs4
import re
import requests
from selenium.webdriver import Chrome
from selenium.webdriver.chrome.options import Options
from openpyxl import Workbook
from openpyxl.styles import Font


def google_search(product):
    url = f'https://www.google.com/search?q=skroutz{product}'
    result = []
    res = requests.get(url)
    if res.status_code == 200:
        soup = bs4.BeautifulSoup(res.text, 'html.parser')
        for link in soup.find_all('a'):
            k = link.get('href')
            try:
                m = re.search("(?P<url>https?://[^\s]+)", k)
                n = m.group(0)
                rul = n.split('&')[0]
                if urlparse(rul).netloc == 'www.skroutz.gr':
                    result.append(rul)
                else:
                    continue
            except:
                continue
        if result:
            response = result[0]
        else:
            response = 'No song lyrics'
    else:
        response = 'Error 404'

    return response


url = google_search('Crucial 4GB DDR4-2400MHz')

print(url)

opts = Options()
opts.set_headless()
assert opts.headless  # Operating in headless mode
browser = Chrome(options=opts)
browser.get(url)

html_source = browser.page_source

soup = bs4.BeautifulSoup(html_source, 'html.parser')
# print(soup)

prices = soup.findAll("a", {"class": "js-product-link product-link content-placeholder"})
names = soup.findAll("a", {"class": "js-product-link content-placeholder"})
description = soup.find("div", {"class": "simple-description js-description-html"})
rating = soup.find("div", {"class": "rating-average cf"})
number_of_users_rating = soup.find("div", {"class": "actual-rating "})

wb = Workbook()

sheet = wb['Sheet']


max_width = 0
for i in range(len(prices)):
    sheet.cell(row=i + 2, column=1).value = prices[i].text
    sheet.cell(row=i + 2, column=2).value = names[i].text
    sheet.cell(row=i + 2, column=2).hyperlink = 'https://www.skroutz.gr' + names[i]['href']
    if len(names[i].text) > max_width:
        sheet.column_dimensions['B'].width = len(names[i].text)
        max_width = len(names[i].text)
try:
    sheet.cell(row=2, column=3).value = rating.text[0:3]
    sheet.cell(row=2, column=4).value = number_of_users_rating.text

except:
    sheet.cell(row=2, column=3).value = 'No ratings'


sheet.cell(row=4, column=3).value = 'Min Price'
sheet.cell(row=4, column=3).font = Font(bold=True)
sheet.cell(row=4, column=4).value = 'Max Price'
sheet.cell(row=4, column=4).font = Font(bold=True)
sheet.cell(row=5, column=3).value = sheet.cell(row=2, column=1).value
sheet.cell(row=5, column=4).value = sheet.cell(row=i+2, column=1).value


sheet.cell(row=1, column=1).value = 'Prices'
sheet.cell(row=1, column=2).value = 'Names'
sheet.cell(row=1, column=3).value = 'Rating'
sheet.cell(row=1, column=4).value = '#Voters'

sheet.cell(row=1, column=1).font = Font(bold=True)
sheet.cell(row=1, column=2).font = Font(bold=True)
sheet.cell(row=1, column=3).font = Font(bold=True)
sheet.cell(row=1, column=4).font = Font(bold=True)

specs = soup.findAll("div", {"class": "spec-details"})
k = 1
max_i = 0
max_l = 0
for j in specs:
    title = j.find("h3", {"class": ""})
    dt = j.findAll("dt", {"class": ""})
    span = j.findAll("span", {"class": ""})
    print(dt)
    try:
        sheet.cell(row=k, column=5).value = title.text
        sheet.cell(row=k, column=5).font = Font(bold=True)
        k = k + 1
    except:
        print("")
    for i, l in zip(dt, span):
        try:
            sheet.cell(row=k, column=5).value = i.text
            sheet.cell(row=k, column=6).value = l.text
            k = k + 1
            if len(i.text) > max_i:
                sheet.column_dimensions['E'].width = len(i.text)
                max_i = len(i.text)
            if len(l.text) > max_l:
                sheet.column_dimensions['F'].width = len(l.text)
                max_l = len(l.text)
        except:
            print("")

wb.save('document.xlsx')
