
def get_lines_from_docx(docx_url):
    return get_lines_from_xml(docx_to_xml(docx_url))

def docx_to_xml(srcurl):
    import zipfile as Z
    import xml.etree.ElementTree as ET
    z=Z.ZipFile(srcurl,'r')
    return ET.fromstring(z.read('word/document.xml').decode('UTF-8'))

def get_lines_from_xml(doc):
    xmlprepender='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    body=doc[0]
    paragraphs=[paragraph for paragraph in body if paragraph.tag==xmlprepender+'p']
    if len(paragraphs)==0:
        print("No paragraphs!")
        return None
    
    lines=[]
    for paragraph in paragraphs:
        rs=[r for r in paragraph if r.tag==xmlprepender+'r']
        part=""
        for r in rs:
            ts=[t for t in r if t.tag==xmlprepender+'t']
            for t in ts:
                part+=t.text
            
        lines.append(part)
    return lines

def count_words(stringdata):
    import re
    return len([i for i in re.split('[ ]+',stringdata) if i])

