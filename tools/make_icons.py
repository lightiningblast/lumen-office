"""Draws the app icons (Write, Cells, Slides, Office Apps) in a detailed, macOS-style look:
rounded-square body with a colour gradient, light from above, soft shadows and a realistic object on top.
Drawn at 2048 px and scaled down, so every size (16 px to 1024 px) is sharp.

  python tools/make_icons.py      -> icons/<app>-1024.png, -512, -256, -192, and <app>.ico (16..256)
"""
import os
from PIL import Image, ImageDraw, ImageFilter, ImageChops

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, 'icons')
S = 2048                       # drawing size (downsampled at the end)
BODY = (164, 164, 1884, 1884)  # icon body inside the canvas (room for its shadow)
RAD = 400


def hexc(h, a=255):
    h = h.lstrip('#')
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16), a)


def vgrad(size, top, bottom):
    """vertical gradient image (RGBA)"""
    w, h = size
    mask = Image.linear_gradient('L').resize((w, h))
    a = Image.new('RGBA', (w, h), top)
    b = Image.new('RGBA', (w, h), bottom)
    return Image.composite(b, a, mask)


def rr_mask(size, box, r):
    m = Image.new('L', size, 0)
    ImageDraw.Draw(m).rounded_rectangle(box, r, fill=255)
    return m


def shadow(canvas, mask, offset=(0, 30), blur=40, alpha=110):
    sh = Image.new('RGBA', canvas.size, (0, 0, 0, 0))
    m = mask.point(lambda v: v * alpha // 255)
    sh.putalpha(m)
    sh = sh.filter(ImageFilter.GaussianBlur(blur))
    moved = Image.new('RGBA', canvas.size, (0, 0, 0, 0))
    moved.paste(sh, offset)
    canvas.alpha_composite(moved)


def paste_masked(canvas, img, mask):
    layer = Image.new('RGBA', canvas.size, (0, 0, 0, 0))
    layer.paste(img, (0, 0), mask)
    canvas.alpha_composite(layer)


def body(top, bottom):
    """the rounded-square app body with gradient, top sheen and edge light"""
    c = Image.new('RGBA', (S, S), (0, 0, 0, 0))
    m = rr_mask((S, S), BODY, RAD)
    shadow(c, m, (0, 34), 44, 120)
    paste_masked(c, vgrad((S, S), hexc(top), hexc(bottom)), m)
    # soft light from above
    sheen = vgrad((S, S), (255, 255, 255, 70), (255, 255, 255, 0))
    sm = rr_mask((S, S), (BODY[0], BODY[1], BODY[2], BODY[1] + 900), RAD)
    paste_masked(c, sheen, ImageChops.multiply(sm, m))
    # thin inner edge highlight
    edge = Image.new('L', (S, S), 0)
    ImageDraw.Draw(edge).rounded_rectangle(BODY, RAD, outline=255, width=6)
    paste_masked(c, Image.new('RGBA', (S, S), (255, 255, 255, 60)), edge)
    return c


def paper(c, box, fold=170, tint='#ffffff'):
    """a white sheet with a folded top-right corner and a soft shadow"""
    x0, y0, x1, y1 = box
    m = Image.new('L', (S, S), 0)
    ImageDraw.Draw(m).polygon([(x0, y0 + 30), (x0 + 30, y0), (x1 - fold, y0), (x1, y0 + fold), (x1, y1 - 30), (x1 - 30, y1), (x0 + 30, y1), (x0, y1 - 30)], fill=255)
    shadow(c, m, (0, 26), 34, 120)
    paste_masked(c, vgrad((S, S), hexc(tint), hexc('#e9edf2')), m)
    f = Image.new('L', (S, S), 0)
    ImageDraw.Draw(f).polygon([(x1 - fold, y0), (x1 - fold, y0 + fold - 20), (x1 - fold + 20, y0 + fold), (x1, y0 + fold)], fill=255)
    shadow(c, f, (-10, 14), 16, 90)
    paste_masked(c, vgrad((S, S), hexc('#f7f9fb'), hexc('#cfd6de')), f)


def finish(c, name):
    img = c.resize((1024, 1024), Image.LANCZOS)
    os.makedirs(OUT, exist_ok=True)
    img.save(os.path.join(OUT, f'{name}-1024.png'))
    for n in (512, 256, 192):
        img.resize((n, n), Image.LANCZOS).save(os.path.join(OUT, f'{name}-{n}.png'))
    img.save(os.path.join(OUT, f'{name}.ico'), sizes=[(16, 16), (20, 20), (24, 24), (32, 32), (40, 40), (48, 48), (64, 64), (96, 96), (128, 128), (256, 256)])


def write_icon():
    c = body('#5aa9ff', '#1a4fd0')
    paper(c, (520, 380, 1440, 1660))
    d = ImageDraw.Draw(c)
    d.rounded_rectangle((640, 560, 1110, 610), 25, fill=hexc('#2f6fe4'))          # heading
    for i, w in enumerate((640, 600, 640, 520, 640, 580, 420)):                 # text lines
        y = 720 + i * 108
        d.rounded_rectangle((640, y, 640 + w, y + 34), 17, fill=hexc('#b9c3cf'))
    # fountain pen lying across the page
    pen = Image.new('RGBA', (S, S), (0, 0, 0, 0))
    pd = ImageDraw.Draw(pen)
    pd.rounded_rectangle((900, 980, 1840, 1120), 70, fill=hexc('#1f2937'))       # barrel
    pd.rounded_rectangle((900, 990, 1840, 1035), 22, fill=hexc('#4b5563'))       # shine on barrel
    pd.rectangle((1180, 980, 1240, 1120), fill=hexc('#d4a63a'))                 # gold band
    pd.polygon([(900, 990), (900, 1110), (690, 1050)], fill=hexc('#d4a63a'))    # nib
    pd.line([(890, 1050), (720, 1050)], fill=hexc('#7a5a12'), width=8)
    pen = pen.rotate(32, resample=Image.BICUBIC, center=(1260, 1050))
    pm = pen.split()[3]
    shadow(c, pm, (16, 40), 30, 120)
    c.alpha_composite(pen)
    finish(c, 'write')


def cells_icon():
    c = body('#5fdc84', '#0f8a3c')
    paper(c, (420, 420, 1630, 1600), fold=150)
    d = ImageDraw.Draw(c)
    gx0, gy0, gx1, gy1 = 520, 560, 1530, 1480
    d.rectangle((gx0, gy0, gx1, gy0 + 130), fill=hexc('#d7f2df'))                # header row
    d.rectangle((gx0, gy0, gx0 + 180, gy1), fill=hexc('#eef7f1'))                # row-label column
    for x in range(gx0, gx1 + 1, 202):
        d.line([(x, gy0), (x, gy1)], fill=hexc('#c5d3cb'), width=6)
    for y in range(gy0, gy1 + 1, 130):
        d.line([(gx0, y), (gx1, y)], fill=hexc('#c5d3cb'), width=6)
    d.rectangle((gx0 + 202 * 2 + 3, gy0 + 130 * 3 + 3, gx0 + 202 * 3 - 3, gy0 + 130 * 4 - 3), fill=hexc('#ffffff'))
    d.rectangle((gx0 + 202 * 2, gy0 + 130 * 3, gx0 + 202 * 3, gy0 + 130 * 4), outline=hexc('#169c47'), width=12)  # selected cell
    # chart card on top
    card = rr_mask((S, S), (980, 1080, 1760, 1720), 70)
    shadow(c, card, (0, 30), 40, 130)
    paste_masked(c, Image.new('RGBA', (S, S), hexc('#ffffff')), card)
    d = ImageDraw.Draw(c)
    for i, (h, col) in enumerate(((250, '#a5e3b8'), (380, '#4cc574'), (300, '#22a653'), (470, '#0e7f3a'))):
        x = 1070 + i * 160
        d.rounded_rectangle((x, 1640 - h, x + 110, 1640), 22, fill=hexc(col))
    d.line([(1050, 1650), (1690, 1650)], fill=hexc('#c9d4cd'), width=8)
    finish(c, 'cells')


def slides_icon():
    c = body('#ffb554', '#e2520c')
    d = ImageDraw.Draw(c)
    # stand
    leg = hexc('#3a3f47')
    d.line([(1024, 1450), (760, 1800)], fill=leg, width=48)
    d.line([(1024, 1450), (1288, 1800)], fill=leg, width=48)
    d.line([(1024, 1450), (1024, 1760)], fill=leg, width=40)
    # screen with frame
    frame = rr_mask((S, S), (380, 440, 1668, 1460), 60)
    shadow(c, frame, (0, 34), 40, 130)
    paste_masked(c, vgrad((S, S), hexc('#ffffff'), hexc('#e9edf2')), frame)
    d = ImageDraw.Draw(c)
    d.rounded_rectangle((330, 400, 1718, 480), 40, fill=hexc('#4b5160'))          # top bar
    d.rounded_rectangle((520, 600, 940, 660), 30, fill=hexc('#c2410c'))           # title
    for i, w in enumerate((420, 360, 400)):
        d.rounded_rectangle((520, 760 + i * 110, 520 + w, 800 + i * 110), 20, fill=hexc('#bcc5d0'))
    # pie chart
    box = (1060, 640, 1500, 1080)
    d.pieslice(box, -90, 120, fill=hexc('#f97316'))
    d.pieslice(box, 120, 230, fill=hexc('#fbbf24'))
    d.pieslice(box, 230, 270, fill=hexc('#3b82f6'))
    d.pieslice((box[0] + 30, box[1] - 30, box[2] + 30, box[3] - 30), -90, -60 + 0, fill=hexc('#ea580c'))
    finish(c, 'slides')


def office_icon():
    c = body('#6b7280', '#2b2f36')
    tiles = [('#5aa9ff', '#1a4fd0', (520, 520)), ('#5fdc84', '#0f8a3c', (1068, 520)), ('#ffb554', '#e2520c', (520, 1068)), ('#f4f4f6', '#c9ccd2', (1068, 1068))]
    for top, bot, (x, y) in tiles:
        m = rr_mask((S, S), (x, y, x + 460, y + 460), 110)
        shadow(c, m, (0, 18), 22, 120)
        paste_masked(c, vgrad((S, S), hexc(top), hexc(bot)), m)
    d = ImageDraw.Draw(c)
    for i in range(3):
        d.rounded_rectangle((1150, 1180 + i * 90, 1450, 1215 + i * 90), 18, fill=hexc('#8a8f99'))
    finish(c, 'office')


if __name__ == '__main__':
    write_icon(); cells_icon(); slides_icon(); office_icon()
    print('icons written to', OUT)
