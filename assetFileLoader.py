from xlrd import open_workbook

def assetListCreator(assetfile):
    #Loads excel spreadsheet and determines how many descriptors are being used
    #Assumes that first column is asset numbers
    wb=open_workbook(assetfile)
    assetList={}
    #Allows for multiple sheets of different asset lists
    for sheet in wb.sheets():
        number_of_rows=sheet.nrows
        number_of_columns = sheet.ncols

        #Makes dictionary of asset numbers and descriptors
        for row in range(1, number_of_rows):
            assetList[int(sheet.cell(row, 0).value)]={}
            for col in range(1, number_of_columns):
                assetList[int(sheet.cell(row, 0).value)][sheet.cell(0, col).value]=sheet.cell(row, col).value
    return(assetList)

#x=assetFileLoader(r'D:\Users\Ryan\Documents\Programming\MobileViewAssistant\DefaultAssetList.xlsx')
