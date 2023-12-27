import openpyxl

# create a new workbook and select the active worksheet
workbook = openpyxl.Workbook()
worksheet = workbook.active

# populate some sample data    
worksheet["A1"] = "Fruit"
worksheet["B1"] = "Color"
worksheet["A2"] = "Apple"
worksheet["B2"] = "Red"
worksheet["A3"] = "Banana"
worksheet["B3"] = "Yellow"
worksheet["A4"] = "Coconut"
worksheet["B4"] = "Brown"

# define a table style
mediumStyle = openpyxl.worksheet.table.TableStyleInfo(name='TableStyleMedium2',
                                                      showRowStripes=True)
# create a table
table = openpyxl.worksheet.table.Table(ref='A1:B4',
                                       displayName='FruitColors',
                                       tableStyleInfo=mediumStyle)
# add the table to the worksheet
worksheet.add_table(table)

# save the workbook file
workbook.save('/home/tosin/samplefruit.xlsx')