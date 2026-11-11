from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
s = prs.slides.add_slide(prs.slide_layouts[0])
s.shapes.title.text = 'Quarterly Review'
s.placeholders[1].text = 'Prepared with python-pptx'
s2 = prs.slides.add_slide(prs.slide_layouts[1])
s2.shapes.title.text = 'Highlights'
tf = s2.placeholders[1].text_frame
tf.text = 'Revenue up 12%'
p = tf.add_paragraph(); p.text = 'New markets opened'; p.level = 1
p = tf.add_paragraph(); p.text = 'Team grew to 40'
box = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9), Inches(5.5), Inches(3), Inches(1.2))
box.fill.solid(); box.fill.fore_color.rgb = RGBColor(0x21, 0x73, 0x46)
box.text_frame.text = 'On track'
tb = s2.shapes.add_textbox(Inches(1), Inches(6.3), Inches(5), Inches(0.6))
r = tb.text_frame.paragraphs[0].add_run(); r.text = 'Bold red note'; r.font.bold = True; r.font.size = Pt(20); r.font.color.rgb = RGBColor(200, 0, 0)
tbl = s2.shapes.add_table(2, 2, Inches(1), Inches(4.6), Inches(5), Inches(1.2)).table
tbl.cell(0, 0).text = 'A'; tbl.cell(0, 1).text = 'B'; tbl.cell(1, 0).text = '1'; tbl.cell(1, 1).text = '2'
prs.save('_testin/pyppt.pptx')
print('ok')
