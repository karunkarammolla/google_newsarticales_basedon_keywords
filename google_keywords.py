import pandas as pd
import time
import re
import datetime
import dateutil
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from bs4 import BeautifulSoup
import time
from htmldate import find_date

import pandas as pd
import os
import sqlite3
from tqdm import tqdm
import feedparser as fp
import json
import newspaper
from newspaper import Article
from datetime import datetime
from datetime import date

import requests
from bs4 import BeautifulSoup


def extract_data(source_url, insert_data_in_table):

    try:

        first_article = Article(url=source_url)
        first_article.download()
        first_article.parse()

        try:
            published_date = find_date(source_url)
        except:
            published_date = ""

        title = first_article.title
        content = first_article.text

        authors = ",".join(first_article.authors)

        try:
            image = first_article.top_image
        except:
            image = ""

        try:
            video = ",".join(first_article.movies)
        except:
            video = ""

        first_article.nlp()
        summary = first_article.summary
        keywords = ",".join(first_article.keywords)
        print(published_date, authors, title)
        insert_data_in_table.append((
            published_date,
            authors,
            title,
            content,
            image,
            video,
            summary,
            keywords,
            source_url,
            company,
            concern
        ))

    except Exception as e:
        print(e)


def get_articles_link(search_word, company):

    options = Options()
    driver = webdriver.Chrome(r'C:\\Desktop\\mahesh_aryan\\chromedriver_win32 (2)\\chromedriver', options=options)

    SEARCH_URL = "https://www.google.com/search?q="+str(search_word)+"&tbm=nws"
    print("searchurl", SEARCH_URL)
    driver.get(SEARCH_URL)
    time.sleep(10)

    Search_Word_List = []
    soup = BeautifulSoup(driver.page_source, 'lxml')
    news = soup.find('div', attrs={'id': 'rso'}).find_all('a', attrs={'style': 'text-decoration:none;display:block'})
    for i in news:
        try:
            print(i['href'])
            Search_Word_List.append(i['href'])
        except Exception as e:
            pass


    Links = driver.find_elements_by_class_name('fl')
    pages = []

    for i in Links:
        pages.append(i.get_attribute('href'))


    for number in pages:
        print("page number:",number)
        try:
            driver.get(number)
            time.sleep(2)
            page_soup = BeautifulSoup(driver.page_source, 'lxml')
            news = page_soup.find('div', attrs={'id': 'rso'}).find_all('a', attrs={'style': 'text-decoration:none;display:block'})
            for i in news:
                try:
                    Search_Word_List.append(i['href'])
                except Exception as e:
                    pass

        except:
            pass


    insert_data_in_table = []

    for links in Search_Word_List:
        extract_data(links,insert_data_in_table)

    insert_data_in_table = pd.DataFrame(insert_data_in_table,
                                        columns=["published_date",
                                                "authors",
                                                "title",
                                                "content",
                                                "image",
                                                "video",
                                                "summary",
                                                "keywords",
                                                "source_url",
                                                 "Company",
                                                 "Keyword"])


    search_word = search_word.replace(" ","")
    search_word = search_word + ".xlsx"


    insert_data_in_table.drop_duplicates(subset=insert_data_in_table.columns, keep=False, inplace=True)

    nameFile = os.getcwd() + '\\' + 'output' + '\\' + search_word
    writer = pd.ExcelWriter(nameFile, engine='xlsxwriter',options={'strings_to_urls': False})
    insert_data_in_table.to_excel(writer, index=False)
    writer.save()

    driver.close()


Company_Word = ["Exxon"]
concern_keywords = ['LGBT']

# concern_keywords = ['LGBT', 'Gender diversity']#,'Gender Pay','Gender Equality','Transgender']

for company in Company_Word:
    for concern in concern_keywords:
        google_search_keyword = str(company)+" "+str(concern)
        print("Searching for " + str(google_search_keyword))
        get_articles_link(google_search_keyword, company)

# Merging all csv files
path = os.getcwd() + '\\' + 'output'
output_path = os.getcwd() + '\\' + 'merged_output'
outfname = 'Google_Keywords_output.xlsx'
output = output_path + outfname
main_df = pd.DataFrame()
for root, dirs, files in os.walk(path):
    for fname in files:
        try:
            print(fname)
            df = pd.DataFrame()
            df = pd.read_excel(path+fname)
            main_df = main_df.append(df)
        except:
            pass

main_df.to_excel(output, index=False, encoding="utf_8_sig")