from lxml import html
import requests
import xlwt
import urllib
import lxml.html


connection = urllib.urlopen('http://nrfbigshow.nrf.com/exhibitors')

dom =  lxml.html.fromstring(connection.read())
links = []

for link in dom.xpath('//a/@href'): # select the url in href for all a tags(links)
    if "company" in link:
        links.append(link)
n = 0
companyNames = []
companyLocations = []
companyURLS = []

for link in links:
    page = requests.get('http://nrfbigshow.nrf.com' + str(link))
    tree = html.fromstring(page.content)
    location = tree.xpath('//div[@class="company_contact_city_state"]/text()')
    companyName = tree.xpath('//div[@class="company_name_mobile_only"]/text()')

    newConnection = urllib.urlopen('http://nrfbigshow.nrf.com' + str(link))
    newDom =  lxml.html.fromstring(newConnection.read())
    try:
        for url in newDom.xpath('//a/@href'): # select the url in href for all a tags(links)
            if companyName[0].decode('utf-8').lower().split(" ", 1)[0] in url.decode('utf-8').lower():
                companyNames.append(companyName[0])
                try:
                    companyLocations.append(location[0])
                except IndexError:
                    companyLocations.append(" ")
                companyURLS.append(url)
                print companyName[0]
                n = n + 1
    except UnicodeEncodeError:
        companyLocations.append(" ")
        companyNames.append(" ")
        companyURLS.append(" ")

book = xlwt.Workbook()
sheet1 = book.add_sheet('Sheet 1')


sheet1.write(0,0, "Company Name")
sheet1.write(0,1, "Company Location")
sheet1.write(0,2, "Company URL")

i = 0

while i < n:
    print i
    sheet1.write(i+1, 0, companyNames[i])
    sheet1.write(i+1, 1, companyLocations[i])
    sheet1.write(i+1, 2, companyURLS[i])
    i = i + 1

book.save('companyDetails.xls')
