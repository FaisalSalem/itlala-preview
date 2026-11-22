"""Make a Claude Docs Word export read right-to-left for Arabic.

Usage: python3 tools/rtl_docx.py input.docx output.docx
- Arabic paragraphs, runs and tables become RTL; English-only blocks stay LTR.
- Page becomes A4 with RTL section direction.
- List indents use logical start/end so numbers sit on the right in Arabic lists.
- Arabic font (Arial for complex script) and ar-SA language are set as defaults.
"""
import re, sys, zipfile

AR = re.compile('[؀-ۿ]')
LAT = re.compile('[A-Za-z]')

def fix_paragraph(m, in_table=False):
    p = m.group(0)
    text = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', p))
    arabic = bool(AR.search(text)) or text.strip() == '' or not LAT.search(text)
    if not arabic:
        return p.replace('<w:bidi/>', '')
    if '<w:pPr>' not in p and '<w:pPr/>' not in p:
        p = p.replace('<w:p>', '<w:p><w:pPr></w:pPr>', 1)
    p = p.replace('<w:pPr/>', '<w:pPr></w:pPr>')
    if '<w:bidi/>' not in p:
        # bidi must come after pStyle/numPr in pPr order; before jc
        p = re.sub(r'(<w:pPr>(?:<w:pStyle[^>]*/>)?(?:<w:keepNext/>)?(?:<w:numPr>.*?</w:numPr>)?)', r'\1<w:bidi/>', p, count=1)
    p = re.sub(r'<w:jc w:val="[a-z]+"/>', '', p)
    # Justify: in Word/LibreOffice the last line follows RTL (right side); viewers that
    # ignore RTL still show the text aligned on the right edge. Headings keep the default.
    if not in_table and 'w:val="Heading' not in p:
        p = p.replace('<w:bidi/>', '<w:bidi/><w:jc w:val="both"/>', 1)
    # runs with Arabic get rtl
    def run(rm):
        r = rm.group(0)
        if AR.search(r) and '<w:rtl/>' not in r:
            if '<w:rPr>' in r:
                r = r.replace('</w:rPr>', '<w:rtl/></w:rPr>', 1)
            else:
                r = r.replace('<w:r>', '<w:r><w:rPr><w:rtl/></w:rPr>', 1)
        return r
    return re.sub(r'<w:r>.*?</w:r>', run, p, flags=re.S)

def fix_table(m):
    t = m.group(0)
    text = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', t))
    if AR.search(text):
        if '<w:bidiVisual/>' not in t:
            t = t.replace('<w:tblPr>', '<w:tblPr><w:bidiVisual/>', 1)
    else:
        t = t.replace('<w:bidiVisual/>', '')
    return t

def fix_document(x):
    x = re.sub(r'<w:tbl>.*?</w:tbl>', fix_table, x, flags=re.S)
    parts = re.split(r'(<w:tbl>.*?</w:tbl>)', x, flags=re.S)
    x = ''.join(re.sub(r'<w:p>.*?</w:p>', (lambda m, t=part.startswith('<w:tbl>'): fix_paragraph(m, t)), part, flags=re.S)
                for part in parts)
    def sect(m):
        s = m.group(0)
        s = re.sub(r'<w:pgSz[^>]*/>', '<w:pgSz w:w="11906" w:h="16838"/>', s)
        s = s.replace('<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"',
                      '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"')
        if '<w:bidi/>' not in s:
            s = s.replace('</w:sectPr>', '<w:bidi/></w:sectPr>')
        return s
    return re.sub(r'<w:sectPr>.*?</w:sectPr>', sect, x, flags=re.S)

def fix_numbering(x):
    x = re.sub(r'<w:ind w:left="(\d+)" w:hanging="(\d+)"/>', r'<w:ind w:start="\1" w:hanging="\2"/>', x)
    return x.replace('<w:lvlJc w:val="left"/>', '<w:lvlJc w:val="start"/>')

def fix_styles(x):
    x = re.sub(r'<w:rFonts [^>]*/>', '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Arial"/>', x, count=1)
    x = x.replace('<w:lang w:val="en-US"/>', '<w:lang w:val="en-US" w:bidi="ar-SA"/>', 1)
    x = re.sub(r'<w:ind w:left="(\d+)"/>', r'<w:ind w:start="\1"/>', x)
    return x

def main(src, dst):
    zin = zipfile.ZipFile(src)
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == 'word/document.xml':
                data = fix_document(data.decode('utf-8')).encode('utf-8')
            elif item.filename == 'word/numbering.xml':
                data = fix_numbering(data.decode('utf-8')).encode('utf-8')
            elif item.filename == 'word/styles.xml':
                data = fix_styles(data.decode('utf-8')).encode('utf-8')
            zout.writestr(item, data)

if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
