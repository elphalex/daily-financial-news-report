# -*- coding: utf-8 -*-
"""
Created on Tue Jun 23 14:04:02 2020

This script generates a daily financial report including news and market data, and sends it via email.
"""

import os
import random
from reportlab.lib.units import cm
from reportlab.lib.colors import HexColor
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
import requests
import pandas as pd
from bs4 import BeautifulSoup
from pytrends.request import TrendReq
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, PageBreak
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.styles import ParagraphStyle
from PyPDF2 import PdfFileWriter, PdfFileReader
import csv

bingsafecount = 0

headings = [
    '<a name="1"/>Bloomberg Finacial News', '<a name="44"/>Bloomberg Commodity News', 
    '<a name="45"/>Bloomberg Rates News', '<a name="46"/>The Motley Fool News', 
    '<a name="47"/>Forbes business News', '<a name="2"/>World Breaking News', 
    '<a name="3"/>Trending topics and Related news', '<a name="4"/>Market Data:', 
    '<a name="5"/>Forex market: Major currency pairs','<a name="6"/>Forex market: Minor currency pairs',
    '<a name="7"/>Forex market: Exotic currency pairs','<a name="48"/>Commodities',
    '<a name="8"/>Cryptocurrency: Market cap','<a name="9"/>Cryptocurrency: Gainers',
    '<a name="10"/>Cryptocurrency: Losers','<a name="11"/>Stock Indexes: Major',
    '<a name="12"/>Stock market: USA market cap','<a name="13"/>Stock market: USA gainers',
    '<a name="14"/>Stock market: USA losers','<a name="15"/>Stock market: USA most active',
    '<a name="16"/>Stock market: UK market cap','<a name="17"/>Stock market: UK gainers',
    '<a name="18"/>Stock market: UK losers','<a name="19"/>Stock market: UK most active',
    '<a name="20"/>Stock market: Netherlands market cap','<a name="21"/>Stock market: Netherlands gainers',
    '<a name="22"/>Stock market: Netherlands losers','<a name="23"/>Stock market: Netherlands most active',
    '<a name="24"/>Stock market: Germany market cap','<a name="25"/>Stock market: Germany gainers',
    '<a name="26"/>Stock market: Germany losers','<a name="27"/>Stock market: Germany most active',
    '<a name="28"/>Stock market: Poland market cap','<a name="29"/>Stock market: Poland gainers',
    '<a name="30"/>Stock market: Poland losers','<a name="31"/>Stock market: Poland most active',
    '<a name="32"/>Stock market: Japan market cap','<a name="33"/>Stock market: Japan gainers',
    '<a name="34"/>Stock market: Japan losers','<a name="35"/>Stock market: Japan most active',
    '<a name="36"/>Stock market: China market cap','<a name="37"/>Stock market: China gainers',
    '<a name="38"/>Stock market: China losers','<a name="39"/>Stock market: China most active',
    '<a name="40"/>Stock market: South Africa market cap','<a name="41"/>Stock market: South Africa gainers',
    '<a name="42"/>Stock market: South Africa losers','<a name="43"/>Stock market: South Africa most active'
]

def create_report():
    aliceblue = HexColor("#1496BB")
    coral = HexColor("#F58B4C")
    daisy = HexColor("#D3B53D")
    indigo = HexColor("#3C6478")
    ruby = HexColor("#CD594A")
    green = HexColor("#A3B86C")
    darkeralice = HexColor('#107896')
    lightblue = HexColor('#add8e6')

    report_name = datetime.today().strftime('%Y-%m-%d') + " Report.pdf"

    pdf = SimpleDocTemplate('draft' + report_name)
    flowobj = []
    styles = getSampleStyleSheet()
    flowobj.append(Paragraph('Table Of Contents', style = styles['Heading1']))

    style = ParagraphStyle(name='Normal', alignment=0, fontSize=10)
    style1 = ParagraphStyle(name='Normal', alignment=1, fontSize=10)

    tocheadings = [
        '<a href="#1" color="black">Bloomberg Finacial News</a>', '<a href="#44" color="black">Bloomberg Commodity News</a>',
        '<a href="#45" color="black">Bloomberg Rates News</a>', '<a href="#46" color="black">The Motley Fool News</a>', 
        '<a href="#47" color="black">Forbes business News</a>', '<a href="#2" color="black">World Breaking News</a>', 
        '<a href="#3" color="black">Trending topics and Related news</a>', '<a href="#4" color="black">Market Data:</a>',
        '<a href="#5" color="black">Forex market: Major currency pairs</a>', '<a href="#6" color="black">Forex market: Minor currency pairs</a>', 
        '<a href="#7" color="black">Forex market: Exotic currency pairs</a>', '<a href="#48" color="black">Commodities</a>', 
        '<a href="#8" color="black">Cryptocurrency: Market cap</a>', '<a href="#9" color="black">Cryptocurrency: Gainers</a>', 
        '<a href="#10" color="black">Cryptocurrency: Losers</a>', '<a href="#11" color="black">Stock Indexes: Major</a>', 
        '<a href="#12" color="black">Stock market: USA market cap</a>', '<a href="#13" color="black">Stock market: USA gainers</a>', 
        '<a href="#14" color="black">Stock market: USA losers</a>', '<a href="#15" color="black">Stock market: USA most active</a>', 
        '<a href="#16" color="black">Stock market: UK market cap</a>', '<a href="#17" color="black">Stock market: UK gainers</a>', 
        '<a href="#18" color="black">Stock market: UK losers</a>', '<a href="#19" color="black">Stock market: UK most active</a>', 
        '<a href="#20" color="black">Stock market: Netherlands market cap</a>', '<a href="#21" color="black">Stock market: Netherlands gainers</a>', 
        '<a href="#22" color="black">Stock market: Netherlands losers</a>', '<a href="#23" color="black">Stock market: Netherlands most active</a>', 
        '<a href="#24" color="black">Stock market: Germany market cap</a>', '<a href="#25" color="black">Stock market: Germany gainers</a>', 
        '<a href="#26" color="black">Stock market: Germany losers</a>', '<a href="#27" color="black">Stock market: Germany most active</a>', 
        '<a href="#28" color="black">Stock market: Poland market cap</a>', '<a href="#29" color="black">Stock market: Poland gainers</a>', 
        '<a href="#30" color="black">Stock market: Poland losers</a>', '<a href="#31" color="black">Stock market: Poland most active</a>', 
        '<a href="#32" color="black">Stock market: Japan market cap</a>', '<a href="#33" color="black">Stock market: Japan gainers</a>', 
        '<a href="#34" color="black">Stock market: Japan losers</a>', '<a href="#35" color="black">Stock market: Japan most active</a>', 
        '<a href="#36" color="black">Stock market: China market cap</a>', '<a href="#37" color="black">Stock market: China gainers</a>', 
        '<a href="#38" color="black">Stock market: China losers</a>', '<a href="#39" color="black">Stock market: China most active</a>', 
        '<a href="#40" color="black">Stock market: South Africa market cap</a>', '<a href="#41" color="black">Stock market: South Africa gainers</a>', 
        '<a href="#42" color="black">Stock market: South Africa losers</a>', '<a href="#43" color="black">Stock market: South Africa most active</a>'
    ]

    tableofcontents = []
    num = 0
    for row in tocheadings:
        l = []
        if num > 7:
            l.append(Paragraph('     8.' + str(num-7), style=style1))
            l.append(Paragraph('                    ' + row, style=style))
        else: 
            l.append(Paragraph(str(num + 1), style=style1))
            l.append(Paragraph(row, style=style))
        tableofcontents.append(l)
        num = num+1

    toc = Table(tableofcontents, colWidths=[4*cm,16*cm], rowHeights=0.475*cm)
    toc.setStyle(TableStyle([('INNERGRID', (0,0), (-1,-1), 0.25, coral),
                             ('BOX', (0,0), (-1,-1), 3.5, colors.black),
                             ('BACKGROUND', (0, 0), (-1, -1), coral)
                            ]))
    flowobj.append(toc)
    flowobj.append(PageBreak())
    bloom = bloom_news()
    breaking = breaking_news()
    motley = motleyfool_news()
    forbes = forbesbusiness_news()
    trends = google_trends()
    marketdata = market_data()

    # Bloomberg financial news table
    heading(flowobj, 0)
    t = Table(bloom[0], colWidths=[10*cm,10*cm])
    t.setStyle(TableStyle([('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),
                           ('BOX', (0,0), (-1,-1), 3.5, colors.black),
                           ('BACKGROUND', (0, 0), (-1, -1), aliceblue)
                          ]))
    flowobj.append(t)
    flowobj.append(PageBreak())

    # Bloomberg commodity news table
    heading(flowobj, 1)
    t = Table(bloom[1], colWidths=[10*cm,10*cm])
    t.setStyle(TableStyle([('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),
                           ('BOX', (0,0), (-1,-1), 3.5, colors.black),
                           ('BACKGROUND', (0, 0), (-1, -1), aliceblue)
                          ]))
    flowobj.append(t)
    flowobj.append(PageBreak())

    # Bloomberg rates news table
    heading(flowobj, 2)
    t = Table(bloom[2], colWidths=[10*cm,10*cm])
    t.setStyle(TableStyle([('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),
                           ('BOX', (0,0), (-1,-1), 3.5, colors.black),
                           ('BACKGROUND', (0, 0), (-1, -1), aliceblue)
                          ]))
    flowobj.append(t)
    flowobj.append(PageBreak())

    # The Motley Fool news table
    heading(flowobj, 3)
    t = Table(motley, colWidths=[10*cm,10*cm])
    t.setStyle(TableStyle([('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),
                           ('BOX', (0,0), (-1,-1), 3.5, colors.black),
                           ('BACKGROUND', (0, 0), (-1, -1), green)
                          ]))
    flowobj.append(t)
    flowobj.append(PageBreak())

    # Forbes business news table
    heading(flowobj, 4)
    t = Table(forbes, colWidths=[6*cm,10*cm,4*cm])
    t.setStyle(TableStyle([('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),
                           ('BOX', (0,0), (-1,-1), 3.5, colors.black),
                           ('BACKGROUND', (0, 0), (-1, -1), coral)
                          ]))
    flowobj.append(t)
    flowobj.append(PageBreak())

    # Breaking news Table
    heading(flowobj, 5)
    t = Table(breaking, colWidths=[10*cm,10*cm])
    t.setStyle(TableStyle([('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),
                           ('BOX', (0,0), (-1,-1), 3.5, colors.black),
                           ('BACKGROUND', (0, 0), (-1, -1), ruby)
                          ]))
    flowobj.append(t)
    flowobj.append(PageBreak())

    # Trends table
    heading(flowobj, 6)
    t = Table(trends, colWidths=[4*cm,8*cm,8*cm])
    t.setStyle(TableStyle([('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),
                           ('BOX', (0,0), (-1,-1), 3.5, colors.black),
                           ('BACKGROUND', (0, 0), (-1, -1), aliceblue)
                          ]))

    flowobj.append(t)
    flowobj.append(PageBreak())

    keys = [
        'forexmajor','forexminor','forexexotic','commodity','cryptotop','cryptogainers','cryptolosers','index',
        'ustop','usgainers','uslosers','usactive',
        'uktop','ukgainers','uklosers','ukactive',
        'nltop','nlgainers','nllosers','nlactive',
        'gertop','gergainers','gerlosers','geractive',
        'potop','pogainers','polosers','poactive',
        'japtop','japgainers','japlosers','japactive',
        'chitop','chigainers','chilosers','chiactive',
        'satop','sagainers','salosers','saactive'
    ]
    forexindexcol = [4.5*cm,2*cm,2.5*cm,2.5*cm,2.5*cm,2.5*cm,3.5*cm]
    stockcol = [5*cm,1.25*cm,1.5*cm,1.25*cm,1.5*cm,2*cm,2*cm,1*cm,1.5*cm,3*cm]
    cryptotopcol = [4*cm,3*cm,3*cm,2*cm,2*cm,2*cm,2*cm]
    cryptocol = [4*cm,4*cm,4*cm,4*cm,4*cm]
    commoditycol = [5*cm,5*cm,5*cm,5*cm]
    x = 0

    heading(flowobj, 7)
    for key in keys: 
        col = []
        if (x >= 0 and x<=2) or (x==7): col = forexindexcol
        if x==3: col = commoditycol
        if (x == 5) or (x==6): col = cryptocol
        if (x == 4): col = cryptotopcol
        if (x > 7): col = stockcol
        x = x + 1
        t = Table(marketdata[key], colWidths=col)
        t.setStyle(TableStyle([('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),
                               ('BOX', (0,0), (-1,-1), 3.5, colors.black),('FONTSIZE', (0, 0), (-1, -1), 8),
                               ('BACKGROUND', (0, 0), (-1, -1), lightblue)
                              ]))

        if x > 3: heading(flowobj, x + 7)
        else: heading(flowobj, x + 7)
        flowobj.append(t)
        flowobj.append(PageBreak())

    pdf.build(flowobj)
    report_name = encrypt(report_name)
    email_recievers = ["email@example.com"]
    for email in email_recievers:
        send_mail(report_name, email)
    print(report_name + " created and email delivered")
    
def heading(flowobj, x):
    styles = getSampleStyleSheet()
    if (x > 7 and x <43) or (x==11): flowobj.append(Paragraph(headings[x],style = styles['Heading2']))
    else: flowobj.append(Paragraph(headings[x],style = styles['Heading1']))
    
def page_setup(c):
    c.setFillColor(HexColor("#312828"))
    path = c.beginPath()
    path.moveTo(0*cm,0*cm)
    path.lineTo(0*cm,29.7*cm)
    path.lineTo(21*cm,29.7*cm)
    path.lineTo(21*cm,0*cm)
    c.drawPath(path, True, True)
    c.setFillColor(HexColor("#00688B"))
    path = c.beginPath()
    path.moveTo(0.25*cm,0.25*cm)
    path.lineTo(0.25*cm,29.45*cm)
    path.lineTo(20.75*cm,29.45*cm)
    path.lineTo(20.75*cm,0.25*cm)
    c.drawPath(path, True, True)
 
def bloom_news():
    url = "https://bloomberg-market-and-financial-news.p.rapidapi.com/stories/list"
    querystring_currency = {"template":"CURRENCY","id":"usdjpy"}
    querystring_commodity = {"template":"COMMODITY","id":"usdjpy"}
    querystring_rate = {"template":"RATE","id":"usdjpy"}
    headers = {
        'x-rapidapi-host': "bloomberg-market-and-financial-news.p.rapidapi.com",
        'x-rapidapi-key': "YOUR_API_KEY"
    }
    response_currency = requests.request("GET", url, headers=headers, params=querystring_currency)
    response_commodity = requests.request("GET", url, headers=headers, params=querystring_commodity)
    response_rate = requests.request("GET", url, headers=headers, params=querystring_rate)
    styles = getSampleStyleSheet()
    style1 = ParagraphStyle(name='Normal', alignment=1, fontSize=10)
    currencydf = pd.DataFrame.from_dict(response_currency.json())
    currencylist = [['News Title', 'News Link']]
    for row in currencydf.iterrows():
        l = []
        l.append(Paragraph(row[1][0]["title"], style=styles["Normal"]))
        l.append(Paragraph('<link href="' + row[1][0]["shortURL"] + '" color="blue">' + 'Click here to view article' + '</link>', style=style1))
        currencylist.append(l)
   
    commoditydf = pd.DataFrame.from_dict(response_commodity.json())
    commoditylist = [['News Title', 'News Link']]
    for row in commoditydf.iterrows():
        l = []
        l.append(Paragraph(row[1][0]["title"], style=styles["Normal"]))
        l.append(Paragraph('<link href="' + row[1][0]["shortURL"] + '" color="blue">' + 'Click here to view article' + '</link>', style=style1))
        commoditylist.append(l)
    
    ratedf = pd.DataFrame.from_dict(response_rate.json())
    ratelist = [['News Title', 'News Link']]
    for row in ratedf.iterrows():
        l = []
        l.append(Paragraph(row[1][0]["title"], style=styles["Normal"]))
        l.append(Paragraph('<link href="' + row[1][0]["shortURL"] + '" color="blue">' + 'Click here to view article' + '</link>', style=style1))
        ratelist.append(l)
        
    return (currencylist, commoditylist, ratelist)    

def motleyfool_news():
    url = 'https://www.fool.com'
    res = requests.get(url)
    soup = BeautifulSoup(res.content,'lxml')   
    news = soup.find('section', class_ = 'hp-news-panel')  
    data = [['News Title', 'Link']]
    styles = getSampleStyleSheet()
    style1 = ParagraphStyle(name='Normal', alignment=1, fontSize=10)
    
    for x in news.findChildren("article" , recursive=False):
        l = []
        if x.find('h2', class_='hp-news-panel-article-header'):
            l.append(Paragraph(x.find('h2', class_='hp-news-panel-article-header').text, style=styles["Normal"]))
        else: 
            l.append(Paragraph(x.find('a').text, style=styles["Normal"]))
        l.append(Paragraph('<link href="' + url + x.find('a').get('href') + '" color="blue">' + 'Click here to view article' + '</link>', style=style1))
        data.append(l)
    return data

def forbesbusiness_news():
    headers = {
        'authority': 'www.forbes.com',
        'cache-control': 'max-age=0',
        'upgrade-insecure-requests': '1',
        'user-agent': 'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.89 Mobile Safari/537.36',
        'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'sec-fetch-site': 'none',
        'sec-fetch-mode': 'navigate',
        'sec-fetch-user': '?1',
        'sec-fetch-dest': 'document',
        'accept-language': 'en-US,en;q=0.9',
        'cookie': 'client_id=4f023ae61ab633d8e7e2410a838a6ef93b8; notice_preferences=2:1a8b5228dd7ff0717196863a5d28ce6c; notice_gdpr_prefs=0,1,2:1a8b5228dd7ff0717196863a5d28ce6c; _cb_ls=1; _cb=DV5nM5D9GHBDCzpjkB; _ga=GA1.2.611301459.1592235761; __tbc=%7Bjzx%7DIFcj-ZhxuNCMjI4-mDfH1HGM-3PFKcN8Miwl1Jhx9eZNEmuQGlLmxXFL-9qM-F_OBO51AtKdJ3qgOfi3P9vM0qBHA3PyvmasSB5xaCbWibdU2meZrLoZ92gJ8xiw07mk3E9l5ifC0NcYbET3aSZxuA; xbc=%7Bjzx%7DGUDHEU3rvhv6-gySw5OY32YdbGDIZI_hJ7AHN4OvkbydVClZ3QNjNrlQVyHGl3ynSJzzGsKf0w3VfH3le6pYqMAfTQAzgDTJbUHa-cJS7p3ITwLt3PmPKKvsIVyFnHji; __gads=ID=a5ac1829fa387f90:T=1592235777:S=ALNI_MZfOqlh-TglrQCWbFNtjcjgFfkMGQ; _fbp=fb.1.1592235779290.60238202; __qca=P0-59264648-1592235777617; xdibx=N4Ig-mBGAeDGCuAnRIBcoAOGAuBnNAjAKwCcATGQMxEDsAHHQAyUkA0IGAbrAHbaHtc-VMXJVaDZmw6dcvfiPaIkAGzQgAFtmwZcqAPT6A7iYB0AMwD2iSAFNcp2JYC2-3AEts9.c.cBrW0sALwBDHncw.TJGaP1GADZ9Yn0eWyNYENxsFVsAWnhwrwATXNwQnNyQrERLTnLK3PMVdwxcy3Nc7A08p3cQdhVVdTdPb18A4LCIniiYxjjE5NT0zOy8gtGSsoqqjBq6lQamlraOrp7Ldx5SkIBPXFy9219bRFyckIBzeDz-kBU8IRSBRqPQmCwAL7sCAwJ6cNCgIp3YQAbVEIIkTAALDQALpQ8BQaC2Ti2PjCUDRBKUSgIkDw9AgWACEAKNHA8T0EiMEiUfGCOnM1CMdhs.kgFCMoUi1loFHioqCtAysUEoWgaWiuX4gnRMhYxiMOkMjUstnozl0ch0IjiilM5Va1DygmS03Cp0u9iKqWO2XO8Xqh0e.0uiEEmFwdw-kAhLHRLEkIodWyQEKwXJYrHxeK5SBEG2Z2A0EJFSCUMsJDMW0F0GhY4ggCFAA__; _chartbeat2=.1592235763483.1592235790833.1.CZMImKDrkr6iBB9QkcCHzJBoDWn8ZI.3',
    }
    styles = getSampleStyleSheet()
    style1 = ParagraphStyle(name='Normal', alignment=1, fontSize=10)
    response = requests.get('https://www.forbes.com/business/', headers=headers)
    soup = BeautifulSoup(response.text, 'lxml')
    data = [['Title','Description', 'Link']]
    for head in soup.find_all("a", {"class": "headlink"}):
        l = []
        l.append(Paragraph(head.text, style=styles["Normal"]))
        l.append('Description not available')
        l.append(Paragraph('<link href="' + head.get('href') + '" color="blue">' + 'Click here to view article' + '</link>', style=style1))
        data.append(l)
    
    for article in soup.find("div", {"class": "chansec-stream__items"}).find_all("article", {"class": "stream-item et-promoblock-removeable-item et-promoblock-star-item"}):
        l = []
        l.append(Paragraph(article.find("a", {"class": "stream-item__title"}).text, style=styles["Normal"]))
        l.append(Paragraph(article.find("div", {"class": "stream-item__description"}).text, style=styles["Normal"]))
        l.append(Paragraph('<link href="' + article.find("a", {"class": "stream-item__title"}).get('href') + '" color="blue">' + 'Click here to view article' + '</link>', style=style1))
        data.append(l)
        
    return data

def breaking_news():
    url = "https://bing-news-search1.p.rapidapi.com/news/trendingtopics"
    querystring = {"cc":"us","setLang":"EN","textFormat":"Raw","safeSearch":"Off"}
    headers = {
        'x-rapidapi-host': "bing-news-search1.p.rapidapi.com",
        'x-rapidapi-key': "YOUR_API_KEY",
        'x-bingapis-sdk': "true"
    }
    response = requests.request("GET", url, headers=headers, params=querystring)
    styles = getSampleStyleSheet()
    style1 = ParagraphStyle(name='Normal', alignment=1, fontSize=10)
    df = pd.DataFrame.from_dict(response.json())
    news = [['News Name', 'News Link']]
    for row in df.iterrows():
        l = []
        l.append(Paragraph(row[1][1]["name"], style=styles["Normal"]))
        l.append(Paragraph('<link href="' + row[1][1]["webSearchUrl"] + '" color="blue">' + 'Click here to view articles' + '</link>', style=style1))
        news.append(l)
    return(news)
    
def market_data():
    url = "https://www.tradingview.com/markets/stocks-"
    countrydict = {
        "usa": "usa/market-movers-", "uk": "united-kingdom/market-movers-", 
        "nl": "netherlands/market-movers-", "ger": "germany/market-movers-", 
        "po": "poland/market-movers-", "jap": "japan/market-movers-", 
        "chi": "china/market-movers-", "sa": "south-africa/market-movers-"
    }
    urldict = {"cap": "large-cap/", "gainers": "gainers/", "losers": "losers/", "active": "active/"}
    topindexurl = "https://www.tradingview.com/markets/indices/quotes-major/"
    cryptourl = "https://www.tradingview.com/markets/cryptocurrencies/prices-all/"
    cryptomoversurl = "https://coinmarketcap.com/gainers-losers/"
    forexurlmajor = "https://www.tradingview.com/markets/currencies/rates-major/"
    forexurlminor = "https://www.tradingview.com/markets/currencies/rates-minor/"
    forexurlexotic = "https://www.tradingview.com/markets/currencies/rates-exotic/"
    commodity = 'https://tradingeconomics.com/forecast/commodity'

    marketdict = {
        'ustop': wst_tradingview(url + countrydict['usa'] + urldict['cap'], 0),
        'usgainers': wst_tradingview(url + countrydict['usa'] + urldict['gainers'], 0),
        'uslosers': wst_tradingview(url + countrydict['usa'] + urldict['losers'], 0),
        'usactive': wst_tradingview(url + countrydict['usa'] + urldict['active'], 0),
        'uktop': wst_tradingview(url + countrydict['uk'] + urldict['cap'], 0),
        'ukgainers': wst_tradingview(url + countrydict['uk'] + urldict['gainers'], 0),
        'uklosers': wst_tradingview(url + countrydict['uk'] + urldict['losers'], 0),
        'ukactive': wst_tradingview(url + countrydict['uk'] + urldict['active'], 0),
        'nltop': wst_tradingview(url + countrydict['nl'] + urldict['cap'], 0),
        'nlgainers': wst_tradingview(url + countrydict['nl'] + urldict['gainers'], 0),
        'nllosers': wst_tradingview(url + countrydict['nl'] + urldict['losers'], 0),
        'nlactive': wst_tradingview(url + countrydict['nl'] + urldict['active'], 0),
        'gertop': wst_tradingview(url + countrydict['ger'] + urldict['cap'], 0),
        'gergainers': wst_tradingview(url + countrydict['ger'] + urldict['gainers'], 0),
        'gerlosers': wst_tradingview(url + countrydict['ger'] + urldict['losers'], 0),
        'geractive': wst_tradingview(url + countrydict['ger'] + urldict['active'], 0),
        'potop': wst_tradingview(url + countrydict['po'] + urldict['cap'], 0),
        'pogainers': wst_tradingview(url + countrydict['po'] + urldict['gainers'], 0),
        'polosers': wst_tradingview(url + countrydict['po'] + urldict['losers'], 0),
        'poactive': wst_tradingview(url + countrydict['po'] + urldict['active'], 0),
        'japtop': wst_tradingview(url + countrydict['jap'] + urldict['cap'], 0),
        'japgainers': wst_tradingview(url + countrydict['jap'] + urldict['gainers'], 0),
        'japlosers': wst_tradingview(url + countrydict['jap'] + urldict['losers'], 0),
        'japactive': wst_tradingview(url + countrydict['jap'] + urldict['active'], 0),
        'chitop': wst_tradingview(url + countrydict['chi'] + urldict['cap'], 0),
        'chigainers': wst_tradingview(url + countrydict['chi'] + urldict['gainers'], 0),
        'chilosers': wst_tradingview(url + countrydict['chi'] + urldict['losers'], 0),
        'chiactive': wst_tradingview(url + countrydict['chi'] + urldict['active'], 0),
        'satop': wst_tradingview(url + countrydict['sa'] + urldict['cap'], 0),
        'sagainers': wst_tradingview(url + countrydict['sa'] + urldict['gainers'], 0),
        'salosers': wst_tradingview(url + countrydict['sa'] + urldict['losers'], 0),
        'saactive': wst_tradingview(url + countrydict['sa'] + urldict['active'], 0),
        'index': wst_tradingview(topindexurl, 1),
        'cryptotop': wst_tradingview(cryptourl, 2),
        'cryptogainers': wst_tradingview(cryptomoversurl, 5),
        'cryptolosers': wst_tradingview(cryptomoversurl, 6),
        'forexmajor': wst_tradingview(forexurlmajor, 4),
        'forexminor': wst_tradingview(forexurlminor, 4),
        'forexexotic': wst_tradingview(forexurlexotic, 4),
        'commodity': wst_tradingview(commodity, 7),
    }
  
    return(marketdict)
    
def wst_tradingview(Url, code):
    while True:
        try:
            res = requests.get(Url)
            soup = BeautifulSoup(res.content,'lxml')
            table = soup.find_all('table')[0]  
            data = pd.read_html(str(table))
            url = 'https://www.google.com/search?q='
            style = ParagraphStyle(
                name='Normal',
                fontSize=7,
                textColor=colors.black
            )
            style1 = ParagraphStyle(
                name='Normal',
                fontSize=7,
                textColor=colors.green
            )
            style2 = ParagraphStyle(
                name='Normal',
                fontSize=7,
                textColor=colors.red
            )
           
            if code == 0:  # stock data
                stocklist = [['STOCK NAME', 'LAST', 'CHG%', 'CHG', 'RATING', 'VOL', 'MKTCAP', 'P/E', 'EPS', 'SECTOR']]
                for row in data[0].iterrows():
                    sty = style1
                    if '-' in str(row[1][2]): sty = style2
                    l = []
                    l.append(Paragraph('<link href="' + url + str(row[1][0]) + ' price' + '" color="black">' + row[1][0] + '</link>', style=style))
                    l.append(Paragraph(str(row[1][1]), style=style))
                    l.append(Paragraph(str(row[1][2]), style=sty))
                    l.append(Paragraph(str(row[1][3]), style=sty))
                    l.append(Paragraph(str(row[1][4]), style=style))
                    l.append(Paragraph(str(row[1][5]), style=style))
                    l.append(Paragraph(str(row[1][6]), style=style))
                    l.append(Paragraph(str(row[1][7]), style=style))
                    l.append(Paragraph(str(row[1][8]), style=style))
                    l.append(Paragraph(str(row[1][10]), style=style))
                    stocklist.append(l)
                data = stocklist
                  
            if code == 1:  # indexes
                indexlist = [['INDEX NAME', 'LAST', 'CHG%', 'CHG', 'HIGH', 'LOW', 'RATING']]
                for row in data[0].iterrows():
                    sty = style1
                    if '-' in row[1][2]: sty = style2
                    l = []
                    l.append(Paragraph('<link href="' + url + row[1][0] + ' price' + '" color="black">' + row[1][0] + '</link>', style=style))
                    l.append(Paragraph(str(row[1][1]), style=style))
                    l.append(Paragraph(row[1][2], style=sty))
                    l.append(Paragraph(str(row[1][3]), style=sty))
                    l.append(Paragraph(str(row[1][4]), style=style))
                    l.append(Paragraph(str(row[1][5]), style=style))
                    l.append(Paragraph(row[1][6], style=style))
                    indexlist.append(l)
                data = indexlist
               
            if code == 2:  # crypto 
                data = data[0]
                data = data[:50]
                cryptolist = [['COIN NAME', 'MKTCAP', 'FD MKTCAP', 'LAST', 'AVL COINS', 'TTL COINS', 'VOLUME', 'CHG%']]
                for row in data.iterrows():
                    sty = style1
                    if '-' in row[1][7]: sty = style2
                    l = []
                    l.append(Paragraph('<link href="' + url + row[1][0] + ' price' + '" color="black">' + row[1][0] + '</link>', style=style))
                    l.append(Paragraph(row[1][1], style=style))
                    l.append(Paragraph(row[1][2], style=style))
                    l.append(Paragraph(str(row[1][3]), style=style))
                    l.append(Paragraph(row[1][4], style=style))
                    l.append(Paragraph(row[1][5], style=style))
                    l.append(Paragraph(row[1][6], style=style))
                    l.append(Paragraph(row[1][7], style=sty))
                    cryptolist.append(l)
                data = cryptolist
               
            if code == 3:  # not using 3
                pass

            if code == 4:  # forex
                data = data[0].drop(labels=['Unnamed: 4','Unnamed: 5'], axis=1)
                forexlist = [['Ticker', 'Last','Chg%','Chg','High', 'Low', 'Rating']]
                for row in data.iterrows():
                    sty = style1
                    if '-' in row[1][2]: sty = style2
                    l = []
                    l.append(Paragraph('<link href="' + url + str(row[1][0]) + ' price' + '" color="black">' + str(row[1][0]) + '</link>', style=style))
                    l.append(Paragraph(str(row[1][1]), style=style))
                    l.append(Paragraph(str(row[1][2]), style=sty))
                    l.append(Paragraph(str(row[1][3]), style=sty))
                    l.append(Paragraph(str(row[1][4]), style=style))
                    l.append(Paragraph(str(row[1][5]), style=style))
                    l.append(Paragraph(str(row[1][6]), style=style))
                    forexlist.append(l)
                data = forexlist
               
            if code == 5:  # crypto gainers
                table = soup.find_all('table')[0]  
                data = pd.read_html(str(table))
                data = data[0]
                data['24h'] = data['24h'].apply(lambda x: '+' + str(x))  
                cryptolist = [['Rank','Coin','Price','Volume 24h', '24h']]
                for row in data.iterrows():
                    sty = style1
                    if '-' in row[1][4]: sty = style2
                    l = []
                    l.append(Paragraph(str(row[1][0]), style=style))
                    l.append(Paragraph('<link href="' + str(url + row[1][1]) + ' price' + '" color="black">' + str(row[1][1]) + '</link>', style=style))
                    l.append(Paragraph(str(row[1][2]), style=style))
                    l.append(Paragraph(str(row[1][3]), style=style))
                    l.append(Paragraph(str(row[1][4]), style=sty))
                    cryptolist.append(l)
                data = cryptolist
               
            if code == 6:  # crypto losers
                table = soup.find_all('table')[1]  
                data = pd.read_html(str(table))
                data = data[0]
                data['24h'] = data['24h'].apply(lambda x: '-' + str(x))       
                cryptolist = [['Rank','Coin','Price','Volume 24h', '24h']]
                for row in data.iterrows():
                    sty = style1
                    if '-' in row[1][4]: sty = style2
                    l = []
                    l.append(Paragraph(str(row[1][0]), style=style))
                    l.append(Paragraph('<link href="' + str(url + row[1][1]) + ' price' + '" color="black">' + str(row[1][1]) + '</link>', style=style))
                    l.append(Paragraph(str(row[1][2]), style=style))
                    l.append(Paragraph(str(row[1][3]), style=style))
                    l.append(Paragraph(str(row[1][4]), style=sty))
                    cryptolist.append(l)
                data = cryptolist
           
            if code == 7:  # commodities
                res = requests.get('https://tradingeconomics.com/forecast/commodity')
                soup = BeautifulSoup(res.content,'lxml').find(class_ = "col-lg-8 col-md-9")
                tables = soup.find_all('table')
                commoditylist = [['COMMODITY NAME', 'PRICE', 'CHG', 'CHG%']]
                for table in tables:
                    data = pd.read_html(str(table))
                    for row in data[0].iterrows():
                        sty = style1
                        chg = str(row[1][3])
                        if '-' in str(row[1][4]):
                            sty = style2
                            chg = '-' + chg
                        l = []
                        l.append(Paragraph('<link href="' + url + row[1][1] + ' price' + '" color="black">' + row[1][1] + '</link>', style=style))
                        l.append(Paragraph(str(row[1][2]), style=style))
                        l.append(Paragraph(str(chg), style=sty))
                        l.append(Paragraph(str(row[1][4]), style=sty))
                        commoditylist.append(l)
                data = commoditylist
           return data
           break
        except ValueError:
            print('Failed to scrape: '+ Url)
            return [['error', 'error']]

def google_trends():
    pytrend = TrendReq()  
    df = pytrend.trending_searches()
    data = df[0]
    styles = getSampleStyleSheet()
    style1 = ParagraphStyle(name='Normal', alignment=1, fontSize=10)
    search = [['Trending Topic', 'Related News', 'News Url']]
    for row in data:
        l = []
        news = bing_news_search(row)
        l.append(Paragraph(row, style=styles["Normal"]))
        l.append(Paragraph(news[0], style=styles["Normal"]))
        l.append(Paragraph('<link href="' + news[1] + '" color="blue">' + 'Click here to view articles related to this topic   ' + '</link>', style=style1))
        search.append(l)
    return(search)

def bing_news_search(query):
    global bingsafecount
    bingsafecount += 1
    if bingsafecount >= 33:
        return ('out of requests', 'out of requests')
    url = "https://microsoft-azure-bing-news-search-v1.p.rapidapi.com/search"
    querystring = {"mkt":"en-US","q":query}
    headers = {
        'x-rapidapi-host': "microsoft-azure-bing-news-search-v1.p.rapidapi.com",
        'x-rapidapi-key': "YOUR_API_KEY"
    }
    response = requests.request("GET", url, headers=headers, params=querystring)
    return (response.json()['value'][0]['name'], response.json()['value'][0]['url'])

def encrypt(filename):
    pdf = open('draft'+filename, "rb") 
    pdf_writer = PdfFileWriter()
    pdf_reader = PdfFileReader(pdf)
    for page in range(pdf_reader.getNumPages()):
        pdf_writer.addPage(pdf_reader.getPage(page))
    pdf_writer.encrypt(user_pwd='minx', owner_pwd=None, use_128bit=True)
    with open(filename, 'wb') as file:
        pdf_writer.write(file)
    return (filename)

def send_mail(file, reciever): 
    from_email = "your_email@gmail.com"
    to_email = reciever  
    msg = MIMEMultipart() 
    msg['From'] = from_email 
    msg['To'] = to_email 
    msg['Subject'] =  ' Report for '+datetime.today().strftime('%A %d %B %Y')
    body = 'no contents'
    msg.attach(MIMEText(body, 'plain')) 
    filename = file
    attachment = open(filename, "rb") 
    p = MIMEBase('application', 'octet-stream') 
    p.set_payload((attachment).read()) 
    encoders.encode_base64(p) 
    p.add_header('Content-Disposition', "attachment; filename= %s" % filename) 
    msg.attach(p) 
    s = smtplib.SMTP('smtp.gmail.com', 587) 
    s.starttls() 
    s.login(from_email, "your_password") 
    text = msg.as_string() 
    s.sendmail(from_email, to_email, text) 
    s.quit() 

def main():
    create_report()
  
if __name__ == "__main__":
    main()


