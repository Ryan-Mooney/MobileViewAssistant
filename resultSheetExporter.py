from assetFileLoader import *
from MobileViewAssetFinder import *
from dbManagement import *
import xlwt, datetime, os

#Exports Data to Excel Report using the assetList
def exportToExcel(assetList, trial, floor_counter, floor_list):
    book = xlwt.Workbook(encoding="utf-8")
    styleHeader=xlwt.easyxf('font: name Times New Roman, height 280, bold on')
    styleSubheader=xlwt.easyxf('font: name Times New Roman, height 240, bold on')
    styleNormal=xlwt.easyxf('font: name Times New Roman, height 240')
    styleSpecial=xlwt.easyxf('font: name Times New Roman, height 240, bold on; pattern: pattern solid, fore_colour green')

    #Initializes first row for the Asset Info sheet
    sheet1 = book.add_sheet("Asset Info")
    sheet1.write(0, 0, "Asset", styleHeader)
    sheet1.write(0, 1, "Location", styleHeader)
    sheet1.write(0, 2, "Type", styleHeader)
    sheet1.write(0, 3, "PM Month", styleHeader)
    sheet1.col(0).width, sheet1.col(1).width, sheet1.col(2).width,  sheet1.col(3).width = 5200, 5200, 6240, 5200
    i=1
    
    for asset in sorted(assetList.keys()):
        sheet1.write(i, 0, asset, styleNormal)
        sheet1.write(i, 1, assetList[asset]['Location'], styleNormal)
        if assetList[asset]['Type']:
            sheet1.write(i, 2, assetList[asset]['Type'], styleNormal)
        if assetList[asset]['PM Month']:
            sheet1.write(i, 3, assetList[asset]['PM Month'], styleNormal)
        i+=1

    #sheet1=autoAdjustColWidth(sheet1)
    
    #Initializes first row for the Location Summary sheet
    sheet2 = book.add_sheet("Location Summary")
    sheet2.write(0, 0, "Location", styleHeader)
    sheet2.write(0, 1, "Type", styleHeader)
    sheet2.write(0, 2, "Number", styleHeader)
    #sheet2.write(0, 3, "Change", styleHeader)
    sheet2.col(0).width, sheet2.col(1).width, sheet2.col(2).width = 5200, 6240, 5200
    i=1
    total=0
    
    for location in sorted(floor_counter.keys()):
        sheet2.write(i, 0, location, styleSubheader)
        for assetType in sorted(floor_counter[location].keys()):
            sheet2.write(i, 1, assetType, styleNormal)
            sheet2.write(i, 2, floor_counter[location][assetType], styleNormal)
            total=total+floor_counter[location][assetType]
            i=i+1
            #Add the change data here when ready
        sheet2.write(i, 1, 'Total', styleSpecial)
        sheet2.write(i, 2, total, styleSpecial)
        total=0
        i+=1

    #sheet2=autoAdjustColWidth(sheet2)
    
    #Initializes first row for the Assets by Location sheet
    sheet3 = book.add_sheet("Assets by Location")
    sheet3.write(0, 0, "Location", styleHeader)
    sheet3.write(0, 1, "Asset", styleHeader)
    sheet3.write(0, 2, "PM Month", styleHeader)
    sheet3.col(0).width, sheet3.col(1).width, sheet3.col(2).width = 3200, 3000, 5200
    i=1

    for location in sorted(floor_list.keys()):
        sheet3.write(i, 0, location, styleSubheader)
        for asset in sorted(floor_list[location]):
            sheet3.write(i, 1, asset, styleNormal)
            if assetList[asset]['PM Month']:
                sheet3.write(i, 2, assetList[asset]['PM Month'], styleNormal)
            i+=1

    #sheet3=autoAdjustColWidth(sheet3)
    
    #Creates name for File
    cwd=os.getcwd()
    name='Asset Results'+' - Trial '+str(trial)+' - '+datetime.datetime.today().strftime('%m-%d-%Y')+'.xls'
    try:
        os.mkdir(cwd+'\\Trial Results\\')
    except:
        pass
    book.save(cwd+'\\Trial Results\\'+name)
    name=cwd+'\\Trial Results\\'+name
    #'./Trial Results/'+
    return(name)

def autoAdjustColWidth(worksheet):
    for col in worksheet.columns:
     max_length = 0
     column = col[0].column # Get the column name
     for cell in col:
         try: # Necessary to avoid error on empty cells
             if len(str(cell.value)) > max_length:
                 max_length = len(cell.value)
         except:
             pass
     adjusted_width = (max_length + 2) * 1.2
     worksheet.column_dimensions[column].width = adjusted_width
     return(worksheet)

#Used for test purposes
##username='adsf'
##password='asdf'
##assetfile='./AssetListTest.xlsx'
##assetList, floor_counter, floor_list=get_asset_locations(username, password, assetListCreator(assetfile))
##connection=connect()
##assetList, trial=assign_trial_number(assetList, connection)
##
##exportToExcel(assetList, trial, floor_counter, floor_list)
