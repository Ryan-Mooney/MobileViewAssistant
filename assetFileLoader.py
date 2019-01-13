from xlrd import open_workbook

def assetListCreator(assetfile, trial_type):
    #Loads excel spreadsheet and determines how many descriptors are being used
    #Assumes that first column is asset numbers
    wb=open_workbook(assetfile)
    assetList={}

    #Determines which assets to add to assetList
    if trial_type[0:8]=="PM Month":
        pm_month=trial_type[10:]
    #Allows for multiple sheets of different asset lists
    for sheet in wb.sheets():
        number_of_rows=sheet.nrows
        number_of_columns = sheet.ncols

        #Makes dictionary of asset numbers and descriptors
        for row in range(1, number_of_rows):
            Attributes={}
            for col in range(1, number_of_columns):
                Attributes[sheet.cell(0, col).value]=sheet.cell(row, col).value
            if trial_type=='All Assets' or Attributes['PM Month']==pm_month:
                assetList[int(sheet.cell(row, 0).value)]=Attributes
    return(assetList)

#x=assetFileLoader(r'D:\Users\Ryan\Documents\Programming\MobileViewAssistant\DefaultAssetList.xlsx')
