import xlwt
import xlrd
import sys
from xlutils.copy import copy
import datetime
import os
import xlwt
 
#Creating styles

GreenStyle = xlwt.easyxf('align: horiz center, vertical center; font: name Calibri, height 220, color green;borders: left thin, right thin, top thin, bottom thin;')
RedStyle = xlwt.easyxf('align: horiz center, vertical center; font: name Calibri, height 220, color red;borders: left thin, right thin, top thin, bottom thin;')
NormalStyle = xlwt.easyxf('align: horiz center, vertical center; font: name Calibri, height 220; borders: left thin, right thin, top thin, bottom thin;')
TitlePassedColumnStyle = xlwt.easyxf('pattern: pattern solid, fore_colour light_yellow; font: name Calibri, height 220, color green, bold 1; align: horiz center, vertical bottom; borders: left thin, right thin, top thin, bottom thin;')
TitleDateColumnStyle = xlwt.easyxf('pattern: pattern solid, fore_colour light_yellow; font: name Calibri, height 220, color black, bold 1; align: horiz center, vertical bottom; borders: left thin, right thin, top thin, bottom thin;')
TitleDateColumnStyle.num_format_str = "DD-MMM-YY"
TitleNoTestPassedColumnStyle = xlwt.easyxf('pattern: pattern solid, fore_colour light_yellow; font: name Calibri, height 220, color green, bold 1; align: horiz center, vertical bottom; borders: left thin, right thin, top thin, bottom thin;')
TitleFailedColumnStyle = xlwt.easyxf('pattern: pattern solid, fore_colour light_yellow; font: name Calibri, height 220, color red, bold 1; align: horiz center, vertical bottom; borders: left thin, right thin, top thin, bottom thin;')
with xlrd.open_workbook("/Users/tusharchopra/Desktop/Tushar/Development/Saion/Saion_TestCase_Template.xls", formatting_info=True, on_demand=True) as readFileObject:
    
    r_sheet         = readFileObject.sheet_by_index(0)
    writeFileObect  = copy(readFileObject)
    w_sheet         = writeFileObect.get_sheet(0)

#update header
    excelCoumns = "F15:F29"
    todaysDate = datetime.datetime.now().date()
    w_sheet.write(7, 3, todaysDate, TitleDateColumnStyle)
    w_sheet.write(8, 3, xlwt.Formula('COUNTIF('+excelCoumns+',"Passed")'), TitlePassedColumnStyle)
    w_sheet.write(9, 3, xlwt.Formula('COUNTIF('+excelCoumns+',"Failed")'), TitleFailedColumnStyle)
# Test Case 1

    if find("1470676922948.png") and find("1470677389901.png") and find("1470677398488.png"):

        w_sheet.write(14, 4, "Yes", NormalStyle)
        w_sheet.write(14, 5, "Passed", GreenStyle)

    else:
        w_sheet.write(14, 5, "Failed", RedStyle)    
 
# Test Case 2

    click("1470676929869.png")
    if find("1470677421089.png"):
        w_sheet.write(15, 4, "Yes", NormalStyle)
        w_sheet.write(15, 5, "Passed", GreenStyle)

    else:
        w_sheet.write(15, 5, "Failed", RedStyle)    
    wait(2)
    click("1470676987473.png")
    wait(1)
    
# Test Case 3

    click("1470677136350.png")
    wait(1)
    if find("1470677449388.png"):
        w_sheet.write(16, 4, "Yes", NormalStyle)
        w_sheet.write(16, 5, "Passed", GreenStyle)

    else:
        w_sheet.write(16, 5, "Failed", RedStyle) 
        
    click("1470676185188.png")
    wait(1)
    
# Test Case 4

    click("1470677334068.png")
    if find("1470677508307.png") and find("1470677550935.png"):    
        w_sheet.write(18, 4, "Yes", NormalStyle)
        w_sheet.write(18, 5, "Passed", GreenStyle)

    else:
        w_sheet.write(18, 5, "Failed", RedStyle) 

# Test Case 5

    click("1470677593156.png")
    find("1470677768311.png")
    click("1470677780475.png")
    if find("1470677990494.png"):
        w_sheet.write(19, 4, "Yes", NormalStyle)
        w_sheet.write(19, 5, "Passed", GreenStyle)

    else:
        w_sheet.write(19, 5, "Failed", RedStyle) 

# Test Case 6

    click("1470678012234.png")
    wait(3)
    if find("1470675956992.png") and find("1470676044605.png") and find("1470676053131.png")and find("1470676070640.png") and find("1470676076438.png"):
       w_sheet.write(20, 4, "Yes", NormalStyle)
       w_sheet.write(20, 5, "Passed", GreenStyle)

    else:
        w_sheet.write(20, 5, "Failed", RedStyle)     
        
# Test Case 7

    if find("1470675956992.png") and find("1470676044605.png") and find("1470676053131.png")and find("1470676070640.png") and find("1470676076438.png"):
        w_sheet.write(22, 4, "Yes", NormalStyle)
        w_sheet.write(22, 5, "Passed", GreenStyle)

    else:
        w_sheet.write(22, 5, "Failed") 
        
# Test Case 8
    if find("1470675980427.png") and find("1470675988231.png"):
        w_sheet.write(23, 4, "Yes", NormalStyle)
        w_sheet.write(23, 5, "Passed", GreenStyle)

    else:
        w_sheet.write(23, 5, "Failed", RedStyle) 

# Test Case 9

    click("1470676098167.png")
    wait(1)
    if find("1470677449388.png"):
        w_sheet.write(24, 4, "Yes", NormalStyle)
        w_sheet.write(24, 5, "Passed", GreenStyle)

    else:
        w_sheet.write(24, 5, "Failed", RedStyle) 
    click("1470676185188.png")
    wait(1)
    
# Test Case 10

    click("1470676220551.png")
    wait(1)
    if find("1470764275066.png"):
        w_sheet.write(25, 4, "Yes", NormalStyle)
        w_sheet.write(25, 5, "Passed", GreenStyle)

    else:
        w_sheet.write(25, 5, "Failed", RedStyle) 
    
    click("1470676185188.png")
    wait(1)
    
# Test Case 11

    click("1470676318709.png")
    wait(5)
    if find("1470764421742.png"):
         w_sheet.write(26, 4, "Yes", NormalStyle)
         w_sheet.write(26, 5, "Passed", GreenStyle)

    else:
         w_sheet.write(26, 5, "Failed", RedStyle) 
    
        
    click("1470676185188.png")
    wait(1)
    
# Test Case 12

    click("1470676392372.png")
    wait(1)
    if find("1470764517723.png"):
        w_sheet.write(27, 4, "Yes", NormalStyle)
        w_sheet.write(27, 5, "Passed", GreenStyle)

    else:
        w_sheet.write(27, 5, "Failed", RedStyle)  
    click("1470676185188.png")
    
    # Test Case 13
    
    click("1470676871384.png")
    wait(1)
    if find("1470676922948.png") and find("1470677389901.png") and find("1470677398488.png"):

        w_sheet.write(28, 4, "Yes", NormalStyle)
        w_sheet.write(28, 5, "Passed", GreenStyle)

    else:
        w_sheet.write(28, 5, "Failed", RedStyle)
    writeFileObect.save("/Users/tusharchopra/Desktop/Saion_TestCase_Template1.xls")
    exit()













