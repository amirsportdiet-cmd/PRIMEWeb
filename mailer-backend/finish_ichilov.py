# -*- coding: utf-8 -*-
"""מוסיף לקובצי איכילוב: לוגו איכילוב + פרטי קשר של אמיר."""
import fitz, glob, os, io, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

CONTACT = ('<div style="text-align:center;direction:rtl;font-family:Arial;font-size:10pt;'
           'line-height:1.5;">אמיר רובין - 052-4844497 או למייל primerct2026@gmail.com</div>')

def rows_of(page):
    segs = []
    for b in page.get_text('dict')['blocks']:
        if b.get('type') != 0: continue
        for ln in b.get('lines', []):
            t = ''.join(s['text'] for s in ln['spans'])
            if t.strip(): segs.append((fitz.Rect(ln['bbox']), t))
    segs.sort(key=lambda s: ((s[0].y0+s[0].y1)/2, -s[0].x1))
    rows, cur, cy = [], [], None
    for r, t in segs:
        c = (r.y0+r.y1)/2
        if cy is None or abs(c-cy) <= 4: cur.append((r,t)); cy = c if cy is None else cy
        else: rows.append(cur); cur=[(r,t)]; cy=c
    if cur: rows.append(cur)
    out = []
    for g in grp_iter(rows):
        pass
    res = []
    for g in rows:
        rect = g[0][0]
        for r,_ in g[1:]: rect |= r
        res.append((rect, ''.join(t for _,t in g)))
    return res
def grp_iter(x): return []

for path in sorted(glob.glob('attachments/*איכילוב*.pdf')):
    if 'הסכמה' in path: continue
    doc = fitz.open(path); log = []
    for pno in range(len(doc)):
        page = doc[pno]
        rows = rows_of(page)
        # --- פרטי קשר: מאחדים את שתי שורות אנשי הקשר לשורה אחת ---
        idx = [i for i,(r,t) in enumerate(rows) if 'assuta.co.il' in t or 'שירן' in t or 'הילה' in t]
        if idx:
            rect = rows[idx[0]][0]
            for i in idx[1:]: rect |= rows[i][0]
            page.add_redact_annot(fitz.Rect(rect.x0-10, rect.y0-1, rect.x1+10, rect.y1+2), fill=(1,1,1))
            page.apply_redactions()
            page.insert_htmlbox(fitz.Rect(rect.x0-40, rect.y0-2, rect.x1+40, rect.y1+8), CONTACT)
            log.append(f'p{pno}:קשר')
        # --- לוגו איכילוב במקום החלל שנשאר ---
        heads = [rr for im in page.get_images(full=True)
                    for rr in page.get_image_rects(im[0]) if rr.y0 < 160]
        if heads:
            # החלל השמאלי (שם היה לוגו אסותא): משמאל לתמונה השמאלית ביותר
            L = min(heads, key=lambda r: r.x0)
            slot = fitz.Rect(70, 20, 209, 130)
            page.insert_image(slot, filename='logo_ich.png', keep_proportion=True)
            log.append(f'p{pno}:לוגו')
    doc.saveIncr() if False else doc.save(path + '.tmp', garbage=3, deflate=True)
    doc.close(); os.replace(path + '.tmp', path)
    print('OK', os.path.basename(path), log)
