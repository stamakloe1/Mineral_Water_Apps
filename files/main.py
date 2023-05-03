import itertools
import pandas as pd
import numpy as np
#Author: Samuel Tamakloe
from openpyxl import *
from sales_class import Store_data,root

#Calculate
def calculate():
    df = pd.read_excel("omft.xlsx")
    df['Total Sales'] = None
    df['Total  Debt'] = None
    df['Total Depot'] = None 
    df['chpMoney'] = None
    df['Total Amount'] = None
    df['Returns'] = None
    
    index_bgsale = df.columns.get_loc('Bag Sale')
    index_price = df.columns.get_loc('Price')
    index_sales = df.columns.get_loc('Total Sales')
    index_dpbag = df.columns.get_loc('Depo bag')
    index_tdepo = df.columns.get_loc('Total Depot')
    index_tdbt = df.columns.get_loc('Total Debt')
    index_chmny = df.columns.get_loc('chpMoney')
    index_chbg = df.columns.get_loc('chpBag')    
    index_chmny = df.columns.get_loc('chpMoney')
    index_ttamnt = df.columns.get_loc('Total Amount')     
    index_dtbag = df.columns.get_loc('Debt bag')
    index_rtbag = df.columns.get_loc('Returns')
    index_gas = df.columns.get_loc('Fuel')
    index_exp = df.columns.get_loc('Expenses')
    index_nbags = df.columns.get_loc('NO.Bags')
    
    
    for row in range(0, len(df)):
        df.iat[row, index_rtbag] = df.iat[row,
                                           index_nbags] - df.iat[row, index_bgsale]- df.iat[row, 
                                                                                            index_dtbag] - df.iat[row, index_dpbag]         
        df.iat[row, index_sales] = df.iat[row,
                                         index_bgsale] * df.iat[row, index_price] 
        df.iat[row, index_tdbt] = df.iat[row,
                                           index_dtbag] * df.iat[row, index_price]         
        df.iat[row, index_tdepo] = df.iat[row,
                                          index_dpbag] * df.iat[row, index_price]               
        df.iat[row, index_chmny] = df.iat[row,
                                           index_chbg] * df.iat[row, index_price]
        df.iat[row, index_ttamnt] = df.iat[row,
                                           index_sales] - df.iat[row, index_tdepo]- df.iat[row, 
                                                                                           index_gas] - df.iat[row, index_exp]        
      
    #lst_row = ['Total'] + list(df.sum())[1:]
    #df2 = pd.DataFrame(data=[lst_row], columns=df.columns)
    #df = df.append(df2, ignore_index=True)
    #df.groupby('Date','NO.Bags','Bag Sale','Depo bag','Debt bag','chpBag','Bonus','Returns',
                #'Price', 'Total Sales','Total Debt','Total Depot','Fuel','chpMoney','Expenses','Total Amount').sum()
    #df.sum()
    
    
    #df.drop(["Total Debt"], axis = 1, inplace = True)
    #df = df.iloc[17:]
        
    df.to_excel("omft.xlsx", index=False)            

def main():
    try:
        wb = load_workbook('omft.xlsx')
        sheet = wb.active
        obj1 = Store_data(sheet)
        obj1.layout()
        root.mainloop()
        obj1.excel()
        obj1.insert()
        wb.save('omft.xlsx')
        calculate()
        try:
            #calculate()
            wb = load_workbook('omft.xlsx')
            sheet = wb.active   
            obj1 = Store_data(sheet)
            obj1.excel()
            #worksheet = sheet.getWorksheets().get(0)
            #cell = worksheet.getCells()
            #cells.groupRows(0,7,True)
            ##cells.groupColumns(1,16,True)
            sheet.delete_cols(18)
            wb.save('omft.xlsx')
        except:
            pass
            #calculate()
        
        
    except PermissionError:
        print('Permission denied : excel.xlsx file is already opened for other use')
        
if __name__ == "__main__":
    main()
else:
    pass