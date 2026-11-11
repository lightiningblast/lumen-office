import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

wb = openpyxl.Workbook()
ws = wb.active
ws.title = 'Sales'
ws.append(['Region', 'Q1', 'Q2', 'Growth'])
for i, (r, a, b) in enumerate([('North', 1200, 1500), ('South', 900, 850), ('East', 2000.5, 2600), ('West', 700, 1100)], start=2):
    ws.append([r, a, b, f'=(C{i}-B{i})/B{i}'])
    ws[f'D{i}'].number_format = '0.0%'
ws.append(['Total', '=SUM(B2:B5)', '=SUM(C2:C5)', '=AVERAGE(D2:D5)'])
for c in 'ABCD':
    ws[c + '1'].fill = PatternFill('solid', fgColor='217346')
    ws[c + '1'].font = Font(bold=True, color='FFFFFF')
    ws[c + '1'].alignment = Alignment(horizontal='center')
ws['B2'].number_format = '#,##0'
ws.freeze_panes = 'B2'
ws.column_dimensions['A'].width = 20
ws.merge_cells('F1:H2')
ws['F1'] = 'Merged title'
ws['F1'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
ws['A6'].border = Border(top=Side(style='thin'))
ws2 = wb.create_sheet('Notes')
ws2['A1'] = '=Sales!B6*2'
ws2['A2'] = '="Total: "&TEXT(Sales!B6,"#,##0")'
wb.save('_testin/opx.xlsx')
print('ok')
