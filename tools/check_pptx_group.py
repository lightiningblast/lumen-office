import sys
from pptx import Presentation
f = sys.argv[1] if len(sys.argv) > 1 else '_testout/slides_group.pptx'
p = Presentation(f)
for i, sl in enumerate(p.slides):
    print('slide', i + 1, 'hidden' if sl._element.get('show') == '0' else 'shown')
    for sh in sl.shapes:
        kids = len(sh.shapes) if sh.shape_type == 6 else ''
        print('  ', sh.shape_type, sh.name, kids)
p.save('_testout/slides_group_resaved.pptx')
print('resaved ok')
