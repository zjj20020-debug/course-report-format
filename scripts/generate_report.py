"""
课程报告生成器 — 哈尔滨工程大学学报格式
用法: python generate_report.py [report_content.json]

输出: RIS_6G_Report.docx （同目录）
需要: python3, pandoc(>=3.x), node, npm模块 docx, lxml
"""
import sys, json, os, subprocess, tempfile, shutil

# ── 0. check deps ──
if subprocess.run(['pandoc','--version'], capture_output=True).returncode != 0:
    sys.exit('❌ pandoc not found. Install: winget install pandoc')
try: from lxml import etree
except ImportError: sys.exit('❌ lxml not found. Install: pip install lxml')

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))  # skill scripts dir
JSON_PATH  = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.getcwd(), 'report_content.json')
OUT_DIR    = os.path.dirname(os.path.abspath(JSON_PATH))
OUT_PATH   = os.path.join(OUT_DIR, 'RIS_6G_Report.docx')

# ── 1. load ──
with open(JSON_PATH, encoding='utf-8') as f:
    data = json.load(f)

# ── 2. ensure docx installed ──
if not os.path.isfile(os.path.join(SCRIPT_DIR, 'node_modules', 'docx', 'package.json')):
    r = subprocess.run(['npm','install','docx'], cwd=SCRIPT_DIR,
                       capture_output=True, text=True)
    if r.returncode != 0:
        sys.exit(f'❌ npm install docx failed in {SCRIPT_DIR}: {r.stderr}')

# ── 3. generate structure docx via JS ──
work_dir = tempfile.mkdtemp(prefix='crf_')
struct_args  = os.path.join(work_dir, 'args.json')
struct_out   = os.path.join(work_dir, 'structure.docx')
with open(struct_args, 'w', encoding='utf-8') as f:
    json.dump({'title': data['title'], 'abst': data['abstract'],
               'keywords': data['keywords'], 'output': struct_out}, f)

gen_js = os.path.join(SCRIPT_DIR, 'generate_course_report.js')
r = subprocess.run(['node', gen_js, struct_args], capture_output=True, text=True)
if r.returncode != 0:
    sys.exit(f'❌ Node structure failed: {r.stderr}')

# ── 4. markdown for pandoc ──
lines = []
for sec in data['sections']:
    if 'refs' in sec: continue  # refs handled separately
    lines += [f'# {sec["heading"]}', '']
    if 'subsections' in sec:
        for sub in sec['subsections']:
            lines += [f'## {sub["subheading"]}', '']
            for p in sub['paragraphs']:
                lines += [p, '']
    if 'paragraphs' in sec:
        for p in sec['paragraphs']:
            lines += [p, '']
md_path = os.path.join(work_dir, 'body.md')
with open(md_path, 'w', encoding='utf-8') as f:
    f.write('\n'.join(lines))

# ── 5. pandoc ──
pandoc_out = os.path.join(work_dir, 'pandoc.docx')
ref_doc = os.path.join(SCRIPT_DIR, 'heu_template.docx')
r = subprocess.run([
    'pandoc', md_path, '-o', pandoc_out,
    '--mathml', '-f', 'markdown+tex_math_dollars',
    '--reference-doc=' + ref_doc,
], capture_output=True, text=True)
if r.returncode != 0:
    sys.exit(f'❌ pandoc failed: {r.stderr}')

# ── 6. XML merge ──
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'

import zipfile
s_zip = zipfile.ZipFile(struct_out, 'r')
b_zip = zipfile.ZipFile(pandoc_out, 'r')
s_xml = etree.fromstring(s_zip.read('word/document.xml'))
b_xml = etree.fromstring(b_zip.read('word/document.xml'))
s_zip.close(); b_zip.close()

s_body = s_xml.find('.//{%s}body' % W)
b_body = b_xml.find('.//{%s}body' % W)

# find section boundaries in structure
all_kids = list(s_body)
sec1_end = None
sec2_sectpr = None
part1 = []
part2 = []
in_part2 = False
for c in all_kids:
    if c.tag == '{%s}p' % W:
        ppr = c.find('{%s}pPr' % W)
        if ppr is not None and ppr.find('{%s}sectPr' % W) is not None:
            part1.append(c)
            in_part2 = True
            continue
    if not in_part2:
        part1.append(c)
    elif c.tag == '{%s}sectPr' % W:
        sec2_sectpr = c
        break
    else:
        part2.append(c)  # placeholder in section 2

# Replace section 2 content with pandoc body
body_parts = [c for c in b_body if c.tag != '{%s}sectPr' % W]
s_body.clear()
for c in part1: s_body.append(c)
for c in body_parts: s_body.append(c)

# ── 7. references ──
ref_section = None
for sec in data['sections']:
    if 'refs' in sec:
        ref_section = sec; break

def add_ref_para(text, eng_font):
    p = etree.SubElement(s_body, '{%s}p' % W)
    ppr = etree.SubElement(p, '{%s}pPr' % W)
    etree.SubElement(ppr, '{%s}jc' % W).set('{%s}val' % W, 'both')
    ind = etree.SubElement(ppr, '{%s}ind' % W)
    ind.set('{%s}firstLine' % W, '0')
    sp = etree.SubElement(ppr, '{%s}spacing' % W)
    sp.set('{%s}after' % W, '0'); sp.set('{%s}line' % W, '240'); sp.set('{%s}lineRule' % W, 'exact')
    r = etree.SubElement(p, '{%s}r' % W)
    rpr = etree.SubElement(r, '{%s}rPr' % W)
    rf = etree.SubElement(rpr, '{%s}rFonts' % W)
    rf.set('{%s}ascii' % W, eng_font); rf.set('{%s}hAnsi' % W, eng_font)
    rf.set('{%s}eastAsia' % W, '宋体'); rf.set('{%s}cs' % W, eng_font)
    etree.SubElement(rpr, '{%s}sz' % W).set('{%s}val' % W, '18')
    etree.SubElement(rpr, '{%s}szCs' % W).set('{%s}val' % W, '18')
    t = etree.SubElement(r, '{%s}t' % W)
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = text

if ref_section:
    # title: 4号黑体 bold 14pt, 单倍行距
    p = etree.SubElement(s_body, '{%s}p' % W)
    ppr = etree.SubElement(p, '{%s}pPr' % W)
    etree.SubElement(ppr, '{%s}jc' % W).set('{%s}val' % W, 'left')
    sp = etree.SubElement(ppr, '{%s}spacing' % W)
    sp.set('{%s}before' % W, '200'); sp.set('{%s}after' % W, '200')
    sp.set('{%s}line' % W, '240'); sp.set('{%s}lineRule' % W, 'auto')
    r = etree.SubElement(p, '{%s}r' % W)
    rpr = etree.SubElement(r, '{%s}rPr' % W)
    rf = etree.SubElement(rpr, '{%s}rFonts' % W)
    rf.set('{%s}ascii' % W, '黑体'); rf.set('{%s}hAnsi' % W, '黑体')
    rf.set('{%s}eastAsia' % W, '黑体'); rf.set('{%s}cs' % W, '黑体')
    etree.SubElement(rpr, '{%s}b' % W)
    etree.SubElement(rpr, '{%s}sz' % W).set('{%s}val' % W, '28')
    etree.SubElement(rpr, '{%s}szCs' % W).set('{%s}val' % W, '28')
    etree.SubElement(r, '{%s}t' % W, {'{http://www.w3.org/XML/1998/namespace}space': 'preserve'}).text = ref_section['heading']

    for ref in ref_section['refs']:
        if ref['type'] == 'en':
            add_ref_para(ref['text'], 'Times New Roman')
        elif ref['type'] == 'cn':
            add_ref_para(ref['cn'], '宋体')
            add_ref_para(ref['en'], 'Times New Roman')

# append final sectPr
if sec2_sectpr is not None:
    s_body.append(sec2_sectpr)

# ── 8. write ──
out_zip = zipfile.ZipFile(OUT_PATH, 'w', zipfile.ZIP_DEFLATED)
with zipfile.ZipFile(struct_out, 'r') as z:
    for item in z.namelist():
        if item == 'word/document.xml': continue
        out_zip.writestr(item, z.read(item))
out_zip.writestr('word/document.xml',
    etree.tostring(s_xml, xml_declaration=True, encoding='UTF-8', standalone=True))
with zipfile.ZipFile(pandoc_out, 'r') as z:
    existing = set(zipfile.ZipFile(struct_out, 'r').namelist())
    for item in z.namelist():
        if item not in existing and item.startswith('word/'):
            out_zip.writestr(item, z.read(item))
out_zip.close()

shutil.rmtree(work_dir)
print(f'[OK] {OUT_PATH}')
