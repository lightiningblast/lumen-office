import sys, glob
import docx
for f in (sys.argv[1:] or glob.glob('_testout/*.docx')):
    d = docx.Document(f)
    print(f, 'paragraphs', len(d.paragraphs), 'tables', len(d.tables), 'sections', len(d.sections))
    d.save(f.replace('.docx', '_resaved.docx').replace('_testout', '_testout/resaved'))
print('all opened ok')
