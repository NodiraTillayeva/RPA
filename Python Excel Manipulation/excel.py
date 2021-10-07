import numpy as np
import pandas as pd
import xlsxwriter

# Get our input data
init_data=pd.read_excel('C:/emails/currencies.xlsx',sheet_name='Sheet1')

# Initialize the excel output file
excel_file_path='C:/emails/currencies.xlsx'
workbook=xlsxwriter.Workbook(excel_file_path)
worksheet=workbook.add_worksheet()
date_format=workbook.add_format({'num_format': 'dd/mm/yy'})

total_rows=len(init_data.index)

for i,col_name in enumerate(init_data.columns):
    worksheet.write(0,i,col_name,date_format)
    if(i==0):
        worksheet.write_column(1,i,init_data[col_name])
    else:
        worksheet.write_column(1, i, init_data[col_name])

for i in range(total_rows):
    sheet_no='Line'+str(i+1)
    worksheet.write_url((i+1), 27, f"internal:'{sheet_no}'!A1", string='Graph'+str(i+1))

for i in range(total_rows):
    worksheet_name='Line'+str(i+1)
    worksheet=workbook.add_worksheet(worksheet_name)
    chart=workbook.add_chart({'type':'scatter','subtype':'straight'})
    sheet_range='=Sheet1!$C$'+str(i+2)+':$AA$'+str(i+2)
    chart.add_series({'categories':'=Sheet1!$C$1:$AA$1',
                          'values':sheet_range,
                          'name':'Currency Change'
                        })
    chart.set_size({'width': 700, 'height': 400})
    chart.set_x_axis({'name': 'Dates'})
    chart.set_y_axis({'name': 'Currency Value'}) 
    chart.set_title({'name': 'Currency Change'})
    worksheet.insert_chart('A1',chart)    
                 



workbook.close()