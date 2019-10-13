import openpyxl

g12b= openpyxl.load_workbook('G12.xlsx')
"""for i in g12b.sheetnames:  
    print(g12b)"""
g12s = g12b['BIO AGRI M1']
a4= g12s['A10']
print(a4.internal_value)
g12b.close()