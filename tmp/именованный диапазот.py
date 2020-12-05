# Мне удалось это решить. Возможно, следующее будет полезно для кого-то еще, 
#кто ищет доступ к значениям каждой ячейки в определенном имени или именованном диапазоне, используя openpyxl.



import openpyxl

wb = openpyxl.load_workbook('пиф_id.xlsx') 

# Этот вариант подходит только для поименнованных списков 
#(Для поименнованных таблиц не подходит. Этот метод их не видит.).

#getting the address 
address = list(wb.defined_names['ПИФ'].destinations)

#removing the $ from the address
for sheetname, cellAddress in address:
    cellAddress = cellAddress.replace('$','')

#looping through each cell address, extracting it from the tuple and printing it out     
worksheet = wb[sheetname]
for i in range(0,len(worksheet[cellAddress])):
    for item in worksheet[cellAddress][i]:
        print item.value`

#========================================
# Использование поименнованных таблиц:

ws = wb[sheet_with_tables]    
for tbl in ws._tables:
    print(tbl)
    print(" : " + tbl.displayName)
    print("   -  name = " + tbl.name)
    print("   -  type = " + (tbl.tableType if isinstance(tbl.tableType, str) else 'n/a'))
    print("   - range = " + tbl.ref)
    print("   - #cols = %d" % len(tbl.tableColumns))
    for col in tbl.tableColumns:
        print("     : " + col.name)