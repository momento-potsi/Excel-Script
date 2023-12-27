
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font

from Styling import DEFAULT_ALIGNMENT, DEFAULT_BORDER, DEFAULT_FILL, DEFAULT_FONT, StyleConfig, formatWrite, ColorPalette
from SheetData import SheetDataEntry, SheetDataEnum, ExcelSheetData

import SystemInfo


class StylingUnitTest(object):
    def colorStyling(self): # Test Passed
        style = StyleConfig (
            currentFont = DEFAULT_FONT, 
            currentBorder = DEFAULT_BORDER, 
            currentFill = DEFAULT_FILL, 
            currentAlignment = DEFAULT_ALIGNMENT
        )

        colorList = [
            ColorPalette.WHITE, 
            ColorPalette.YELLOW_HIGHLIGHT, 
            ColorPalette.BRIGHT_GREEN, 
            ColorPalette.PALE_GREEN, 
            ColorPalette.PALE_BLUE, 
            ColorPalette.GRAY,
            ColorPalette.PALE_YELLOW, 
            ColorPalette.PALE_ORANGE,
            ColorPalette.PALE_RED
        ]
        
        wb = Workbook()
        wb.save(SystemInfo.CURRENT_PATH + "sample.xlsx") # to clear file

        for i in range(len(colorList)):    
            testColor: str = colorList[i].value[0]
            # fills in background color
            style.currentFill = PatternFill (fill_type = "solid", start_color = testColor, end_color = '00FFFFFF')

            wb = load_workbook(SystemInfo.CURRENT_PATH + "sample.xlsx")   
            formatWrite(wb, style, 'A' + str(i + 1), str(colorList[i].name))
            wb.save(SystemInfo.CURRENT_PATH + "sample.xlsx")
            print("[Unit Test]: Styling " + colorList[i].value[0])

        """ Output => {
                [Testing]...
                [Unit Test]: Styling 00FFFFFF
                [Unit Test]: Styling FFFFF200
                [Unit Test]: Styling FF72BF44
                [Unit Test]: Styling FFCCFFCC
                [Unit Test]: Styling FFADC5E7
                [Unit Test]: Styling FF808080
                [Unit Test]: Styling FFFFF9AE
                [Unit Test]: Styling FFF9A870
                [Unit Test]: Styling FFF37B70
                Done!
            }        
        """
# End StylingUnitTest
            

class DataEntryTest(object):

    def entryInstanceTest():
        cellType = SheetDataEnum.Cell
        
        pass
# Todo: add commenting, refactor for performance, clean up
