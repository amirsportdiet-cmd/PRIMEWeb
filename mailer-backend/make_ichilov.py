# -*- coding: utf-8 -*-
"""ממיר קובצי הנחיות אסותא -> איכילוב (מיקום, חניה, לוגו)."""
import fitz, os, glob, io, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

HDR1 = 'הינך מוזמנ/ת להגיע ליום המחקר, שיתקיים באיכילוב -'
HDR2 = 'המרכז הרפואי ת"א ע"ש סוראסקי, דרך וייצמן 6'
HDR3 = '(מגדל אריסון - קומה 13, המכון האנדוקריני).'
PARK_A = 'הסדרת חניה ליום המחקר: בהגיעך תקבל מאיתנו כרטיס חניה לחניון ביה"ח.'
PARK_B = 'יש לשים לב כי חנית בחניון ביה"ח (דרך וייצמן 6, תל אביב).'

def box(txt, size):
    return (f'<div style="text-align:center;direction:rtl;font-family:Arial;'
            f'font-size:{size}pt;font-weight:bold;line-height:1.2;">{txt}</div>')

def rows_of(page):
    """מקבץ מקטעי RTL לשורות ויזואליות: [(rect, text)]"""
    segs = []
    for b in page.get_text('dict')['blocks']:
        if b.get('type') != 0: continue
        for ln in b.get('lines', []):
            t = ''.join(s['text'] for s in ln['spans'])
            if t.strip(): segs.append((fitz.Rect(ln['bbox']), t))
    segs.sort(key=lambda s: ((s[0].y0 + s[0].y1) / 2, -s[0].x1))
    rows, cur, cy = [], [], None
    for r, t in segs:
        c = (r.y0 + r.y1) / 2
        if cy is None or abs(c - cy) <= 4:
            cur.append((r, t)); cy = c if cy is None else cy
        else:
            rows.append(cur); cur = [(r, t)]; cy = c
    if cur: rows.append(cur)
    out = []
    for grp in rows:
        rect = grp[0][0]
        for r, _ in grp[1:]: rect |= r
        out.append((rect, ''.join(t for _, t in grp)))
    return out

def convert(src, dst):
    doc = fitz.open(src); log = []
    for pno in range(len(doc)):
        page = doc[pno]
        rows = rows_of(page); jobs = []
        for i, (r, t) in enumerate(rows):
            if 'אסותא' in t and 'מוזמנ' in t:                      # שורת הכותרת
                jobs.append((r, box(HDR1, 10)))
                if i + 1 < len(rows) and ('הברזל' in rows[i+1][1] or 'קומה' in rows[i+1][1]):
                    jobs.append((rows[i+1][0], box(HDR2 + ' ' + HDR3, 8.5)))
                log.append(f'p{pno}:מיקום')
            elif 'חניון' in t and 'הברזל' in t:                    # שורת החניה
                rect = r
                # השורה שמעל ("הסדרת חניה...") חופפת אנכית — נרנדר את שתיהן יחד
                if i > 0 and 'חניה' in rows[i-1][1] and rows[i-1][0].y1 >= r.y0 - 1:
                    rect = rect | rows[i-1][0]
                    jobs.append((rect, box(PARK_A + '<br>' + PARK_B, 10)))
                else:
                    jobs.append((rect, box(PARK_B, 10)))
                log.append(f'p{pno}:חניה')
        for r, _ in jobs:
            page.add_redact_annot(fitz.Rect(r.x0-8, r.y0+0.5, r.x1+8, r.y1-0.5), fill=(1,1,1))
        heads = [rr for im in page.get_images(full=True)
                    for rr in page.get_image_rects(im[0]) if rr.y0 < 160]
        if heads:
            L = min(heads, key=lambda r: r.x0)
            page.add_redact_annot(fitz.Rect(L.x0-2, L.y0-2, L.x1+2, L.y1+2), fill=(1,1,1))
            log.append(f'p{pno}:לוגו')
        if jobs or heads: page.apply_redactions()
        for r, html in jobs:
            page.insert_htmlbox(fitz.Rect(r.x0-16, r.y0-3, r.x1+16, r.y1+8), html)
    doc.save(dst, garbage=3, deflate=True); doc.close()
    return log

os.makedirs('attachments_ichilov', exist_ok=True)
for src in sorted(glob.glob('attachments/*.pdf')):
    base = os.path.basename(src)
    if 'שתן' in base or 'הסכמה' in base:
        print('דילוג:', base); continue
    dst = os.path.join('attachments_ichilov', base.replace('אסותא', 'איכילוב'))
    print('OK', os.path.basename(dst), convert(src, dst))
