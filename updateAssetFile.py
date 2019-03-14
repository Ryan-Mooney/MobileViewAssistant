from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from xlrd import open_workbook
from xlutils.copy import copy
import random, time, re, xlwt, datetime, os, xlutils, xlwt

def getNewAssetList():
    wb=open_workbook(r'D:\Users\Ryan\Documents\Programming\MobileViewAssistant\DefaultAssetList.xlsx')
    assetList={}

    #Allows for multiple sheets of different asset lists
    for sheet in wb.sheets():
        number_of_rows=sheet.nrows
        number_of_columns = sheet.ncols

        #Makes dictionary of asset numbers and descriptors
        for row in range(1, number_of_rows):
            Attributes={}
            for col in range(1, number_of_columns):
                Attributes[sheet.cell(0, col).value]=sheet.cell(row, col).value
                assetList[(sheet.cell(row, 0).value)]=Attributes


    wb2=open_workbook(r'D:\Users\Ryan\Documents\Programming\MobileViewAssistant\Updated Mercy Asset List.xls')
    assetList2={}

    #Allows for multiple sheets of different asset lists
    for sheet in wb2.sheets():
        number_of_rows=sheet.nrows
        number_of_columns = sheet.ncols

        #Makes dictionary of asset numbers and descriptors
        for row in range(1, number_of_rows):
            Attributes={}
            for col in range(1, number_of_columns):
                Attributes[sheet.cell(0, col).value]=sheet.cell(row, col).value
                assetList2[(sheet.cell(row, 0).value)]=Attributes
                
    #Updates asset list
    for asset in assetList:
        if asset in assetList2:
            assetList[asset]['CEID']=assetList2[asset]['CEID']
            assetList[asset]['Serial Number']=assetList2[asset]['Serial Number']
        else:
            assetList[asset]['CEID']=''
            assetList[asset]['Serial Number']=''

    #Exports to new excel workbook
    book = xlwt.Workbook(encoding="utf-8")
    styleHeader=xlwt.easyxf('font: name Times New Roman, height 280, bold on')
    styleSubheader=xlwt.easyxf('font: name Times New Roman, height 240, bold on')
    styleNormal=xlwt.easyxf('font: name Times New Roman, height 240')
    styleSpecial=xlwt.easyxf('font: name Times New Roman, height 240, bold on; pattern: pattern solid, fore_colour light_green')
    styleLowBattery=xlwt.easyxf('font: name Times New Roman, height 240, colour red, bold on; pattern: pattern solid, fore_colour light_orange')

    #Initializes first row for the Asset Info sheet
    sheet1 = book.add_sheet("Asset Info")
    sheet1.write(0, 0, "Asset", styleHeader)
    sheet1.write(0, 1, "Type", styleHeader)
    sheet1.write(0, 2, "PM Month", styleHeader)
    sheet1.write(0, 3, "PM Month Number", styleHeader)
    sheet1.write(0, 4, "CEID", styleHeader)
    sheet1.write(0, 5, "Serial Number", styleHeader)
    i=1

    for asset in assetList.keys():
        if asset !=' ' or asset !="'":
            sheet1.write(i, 0, asset, styleNormal)
            sheet1.write(i, 1, assetList[asset]['Type'], styleNormal)
            sheet1.write(i, 2, assetList[asset]['PM Month'], styleNormal)
            sheet1.write(i, 3, assetList[asset]['PM Month Number'], styleNormal)
            sheet1.write(i, 4, assetList[asset]['CEID'], styleNormal)
            sheet1.write(i, 5, assetList[asset]['Serial Number'], styleNormal)
        i=i+1

    cwd=os.getcwd()
    file=cwd+'\\NewDefaultAssetList.xls'
    print(file)
    book.save(cwd+'\\NewDefaultAssetList.xls')
    os.startfile(file)
    return(assetList, assetList2)

assetList, assetList2=getNewAssetList()
