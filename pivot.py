import pandas as pd
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
import sys
from pprint import pprint
from datetime import datetime
import xlsxwriter
# from openpyxl import load_workbook

def set_style(writer, name, data):
    add = 0
    if name == 'SUMM-ABS Status':
        add = 2
    workbook = writer.book
    worksheet = writer.sheets[name]
    # set border for total grid 
    worksheet.set_column('A:B', 20)
    border_fmt = workbook.add_format({'bottom':1, 'top':1, 'left':1, 'right':1})
    worksheet.conditional_format(xlsxwriter.utility.xl_range(0, 0, len(data)+add, len(data.columns)+1), {'type': 'no_errors', 'format': border_fmt})
    back_color = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#000000'})
    # color_range = "K1:K{}".format(len(data)+3)
    # set color for first header row
    row = add + 1
    color_range = "A1:K{}".format(row, row)
    worksheet.conditional_format(color_range, {'type': 'no_errors','format': back_color})
    # set color for last last row
    row = len(data)+add + 1
    color_range = "A{}:K{}".format(row, row)
    worksheet.conditional_format(color_range, {'type': 'no_errors','format': back_color})
    cell_format = workbook.add_format()
    # set bold for last last row
    cell_format.set_bold(True)
    worksheet.conditional_format(color_range, {'type': 'no_errors','format': cell_format})
    # set bold for last column
    color_range = "K1:K{}".format(len(data)+3)
    worksheet.conditional_format(color_range, {'type': 'no_errors','format': cell_format})

    # Set bod for rows with sub total
    cnt = add + 1
    for index, row in data.iterrows():
        cnt += 1
        # Check if it is sub total row
        if len(index[1]) == 0:
            color_range = "C{}:K{}".format(cnt, cnt)
            cell_format = workbook.add_format()
            cell_format.set_bold(True)
            worksheet.conditional_format(color_range, {'type': 'no_errors','format': cell_format})
            

if __name__ == '__main__':
    # from pandas import ExcelWriter
    pd.set_option('display.max_rows', 10000)
    pd.set_option('display.max_colwidth', 300)

    prefix = 'UBILLS TEST'
    infile = prefix + '.txt'
    # Read txt file before process
    df1 = pd.read_csv(infile, sep=';')
    # Replace space with '_' so that we can indicate with variable 
    df1.columns = df1.columns.str.replace(' ', '_')
    # Convert input data to numeric values
    df1['DAYS'] = pd.to_numeric(df1['DAYS'],errors='coerce')
    df1['><30_DAYS'] = np.where(df1['DAYS']>30,'GREATER THAN 30 DAYS','LESS THAN 30 DAYS')
    # Create writer for output
    outfile = prefix + '.xlsx'
    writer = pd.ExcelWriter(outfile, engine ='xlsxwriter')
    # Save only Data sheet with input
    df1.to_excel(writer, sheet_name='Data')
    # Replace null with blank
    df1['ABS_STATS'].fillna('(blank)', inplace=True)
    df1['CHECKS'].fillna('(blank)', inplace=True)
    # Create Sum Pivot
    sum_bar_status = pd.pivot_table(df1, values='ACCOUNT_NO', 
        index=['BAR_STATUS', 'ACCT_TYPE'], 
        columns='CHECKS', 
        aggfunc="count",
        margins=True)

    # Save Last row named "All"
    last_row = sum_bar_status.tail(1)
    # Remove this last field
    sum_bar_status.drop(sum_bar_status.tail(1).index,inplace=True)
    # Get Sum of sub group
    grp_bar_status = sum_bar_status.groupby('BAR_STATUS').sum().assign(ACCT_TYPE='').set_index('ACCT_TYPE', append=True)
    # Append this Sum of subgroup to original data
    res_bar_status = sum_bar_status.append(grp_bar_status).sort_index()
    # Append saved "All" record
    res_bar_status = res_bar_status.append(last_row)


    # Duplication calculation for average
    sum_abs_status = pd.pivot_table(df1, values='ACCOUNT_NO', 
        index=['BAR_STATUS', 'ACCT_TYPE'], 
        columns=['><30_DAYS', 'ABS_STATS'], 
        aggfunc="count",
        margins=True)

    last_row = sum_abs_status.tail(1)
    sum_abs_status.drop(sum_abs_status.tail(1).index,inplace=True)
    grp_abs_status = sum_abs_status.groupby('BAR_STATUS').sum().assign(ACCT_TYPE='').set_index('ACCT_TYPE', append=True)
    res_abs_status = sum_abs_status.append(grp_abs_status).sort_index()
    res_abs_status = res_abs_status.append(last_row)

    # result = pd.concat([
    #     d.append(d.sum().rename((k, 'Total')))
    #     for k, d in sum_bar_status.groupby(level=0)
    # ]).append(sum_bar_status.sum().rename(('Grand', 'Total')))
    # result = pd.concat([
    #     d.append(d.sum().rename((k, 'Total')))
    #     for k, d in sum_bar_status.groupby(level=0)
    # ])

    # Save calculated sum and average data to each new sheet
    res_abs_status.to_excel(writer, sheet_name='SUMM-ABS Status')
    res_bar_status.to_excel(writer, sheet_name='SUMM-BAR Status')

    set_style(writer, 'SUMM-ABS Status', res_abs_status)
    set_style(writer, 'SUMM-BAR Status', res_bar_status)

    writer.save()

