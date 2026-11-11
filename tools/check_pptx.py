import sys
from pptx import Presentation
p = Presentation(sys.argv[1])
print('size', p.slide_width, p.slide_height, 'slides', len(p.slides))
for i, s in enumerate(p.slides):
    print('--- slide', i + 1, s.slide_layout.name)
    for sh in s.shapes:
        t = sh.text_frame.text[:40].replace('\n', ' | ') if sh.has_text_frame else ''
        print(' ', sh.shape_type, sh.name, sh.is_placeholder, round(sh.left / 9525), round(sh.top / 9525), repr(t))
        if sh.has_table if hasattr(sh, 'has_table') else False:
            print('   table', [[c.text for c in r.cells] for r in sh.table.rows])
