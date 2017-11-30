from openpyxl import load_workbook

wb = load_workbook(filename='accelerate.xlsx')
print ('All availabel sheets in workbook are:  ', wb.sheetnames)
ws = wb['Sheet']

year_count = 0
month_count = 0
week_count = 0

prod_year_count = 0
prod_month_count = 0
prod_week_count = 0

for row in range(1, ws.max_row + 1):
    if str(ws['A' + str(row)].value) == 'Year':
        year_count += ws['F' + str(row)].value or 0
        time_A = ws['B' + str(row)].value or 0
        time_C = ws['D' + str(row)].value or 0
        prod_year_count += time_A + time_C

    if str(ws['A' + str(row)].value) == 'Month':
        month_count += ws['F' + str(row)].value or 0
        time_A = ws['B' + str(row)].value or 0
        time_C = ws['D' + str(row)].value or 0
        prod_month_count += time_A + time_C

    if str(ws['A' + str(row)].value) == 'Week':
        week_count += ws['F' + str(row)].value or 0
        time_A = ws['B' + str(row)].value or 0
        time_C = ws['D' + str(row)].value or 0
        prod_week_count += time_A + time_C

for row in range(1, ws.max_row + 1):

    if str(ws['A' + str(row)].value) == 'Year Count':
        ws['B' + str(row)] = year_count
    if str(ws['A' + str(row)].value) == 'Month Count':
        ws['B' + str(row)] = month_count
    if str(ws['A' + str(row)].value) == 'Week Count':
        ws['B' + str(row)] = week_count

    if str(ws['A' + str(row)].value) == 'prod-year':
        ws['B' + str(row)] = year_count / prod_year_count
    if str(ws['A' + str(row)].value) == 'prod-month':
        ws['B' + str(row)] = year_count / prod_month_count
    if str(ws['A' + str(row)].value) == 'prod-week':
        ws['B' + str(row)] = year_count / prod_week_count

wb.save(filename='accelerate.xlsx')
