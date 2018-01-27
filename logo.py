# -*- coding: utf-8 -*-
"""
Created on Tue Aug  8 09:49:56 2017

@author: HoWingWong
"""
import pyodbc
import urllib
import pandas as pd
import numpy as np
from bs4 import BeautifulSoup
import requests
import re


LogoOrganization = pd.read_excel('LogoOrganization.xlsx')

def findlogo():
    for i in range(len(LogoOrganization.organizationid)):
        try:
            urllib.request.urlretrieve("https://logo.clearbit.com/" + str(LogoOrganization["website url"][i]), str(LogoOrganization.organizationid[i]) + ".jpg")
        except:
            print("Couldn't find " +str([i])+ ' logo')


def searching(search):
    page = requests.get("http://www.google.com/search?q="+search)
    soup = BeautifulSoup(page.content,"html5lib" )
    links = soup.find_all("a",href=re.compile("(?<=/url\?q=)(htt.*://.*)"))
    urls = [re.split(":(?=http)",link["href"].replace("/url?q=",""))[0] for link in links]
    return [url for url in urls if 'webcache' not in url]

def findweb():
    for i in range(len(LogoOrganization.organizationid)):
        try:
            LogoOrganization.loc[LogoOrganization.organizationname == LogoOrganization.organizationname[i],['website url']] = searching(LogoOrganization.organizationname[i])[0].split('/')[2]
        except:
            print(LogoOrganization.organizationname[i] + ' Error')

def nFile(data, nFilename):
    '''
    This function is to export Pandas DataFrame into xlsx file
    '''
    writer = pd.ExcelWriter(nFilename + '.xlsx', engine = 'xlsxwriter')  
    data.to_excel(writer, sheet_name = 'sheet1')  
    writer.save()
#print (searching(LogoOrganization.organizationname[i])[0].split('/')[2])
#url?q=https: