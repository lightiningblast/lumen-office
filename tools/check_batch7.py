"""Opens the batch-7 test files with python-pptx / python-docx / openpyxl and prints what each library sees."""
import zipfile
from pptx import Presentation
import docx
import openpyxl

p = Presentation('_testout/slides_master.pptx')
m = p.slide_masters[0]
print('pptx master shapes:', [(s.shape_type, s.name) for s in m.shapes])
for i, sl in enumerate(p.slides):
    print('  slide', i + 1, 'showMasterSp=', sl._element.get('showMasterSp'), 'shapes', len(sl.shapes))
p.save('_testout/resaved/slides_master.pptx')

d = docx.Document('_testout/word_cite.docx')
print('docx paragraphs:', [para.text for para in d.paragraphs])
z = zipfile.ZipFile('_testout/word_cite.docx')
print('docx parts:', [n for n in z.namelist() if 'customXml' in n])
body = z.read('word/document.xml').decode('utf8')
print('citation fields:', body.count('CITATION'), 'sdt citation tags:', body.count('<w:citation/>'))
d.save('_testout/resaved/word_cite.docx')

wb = openpyxl.load_workbook('_testout/cells_solver.xlsx')
print('xlsx sheets:', wb.sheetnames, 'custom props:', [(c.name, c.value[:60]) for c in wb.custom_doc_props.props] if hasattr(wb, 'custom_doc_props') else 'n/a')
wb.save('_testout/resaved/cells_solver.xlsx')
print('all three re-saved ok')
