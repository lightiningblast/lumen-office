import sys, zipfile
import openpyxl
f = sys.argv[1] if len(sys.argv) > 1 else '_testout/cells_tables.xlsx'
wb = openpyxl.load_workbook(f)
ws = wb.worksheets[0]
print('sheets', wb.sheetnames)
for t in ws.tables.values():
    print('table', t.name, t.ref, t.tableStyleInfo.name, 'totals', t.totalsRowCount, [c.name for c in t.tableColumns])
print('header row', [c.value for c in ws[1]])
print('row 7', [c.value for c in ws[7]])
z = zipfile.ZipFile(f)
sx = z.read('xl/worksheets/sheet1.xml').decode('utf8')
print('sparkline groups', sx.count('<x14:sparklineGroup '), 'tableParts', sx.count('<tablePart '))
wb.save('_testout/cells_tables_resaved.xlsx')
print('resaved ok')
