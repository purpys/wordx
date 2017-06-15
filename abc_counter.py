xmlprepender='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

import xml.etree.ElementTree as ET
import zipfile as Z

people=["A","B","C"]

def choose_folder():
    import tkinter
    root=tkinter.Tk()
    result =tkinter.filedialog.askdirectory()
    root.withdraw()
    return result

def list_docs(path):
    import os
    files=[path+os.path.sep+file for file in os.listdir(path) if file.endswith(".docx") and not file.startswith("~") and not file.startswith(".")]
    return files

def docx_to_xml(srcurl):
    z=Z.ZipFile(srcurl,'r')
    return ET.fromstring(z.read('word/document.xml').decode('UTF-8'))
    
def dump_report(srcurl, result, lines):
    for a in people:
        lines.append(a + ": " + str(result[a]))

def get_lines(doc):
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

def count_words(data):
    import re
    return len([i for i in re.split('[ ]+',data) if i])

def parse_header(line, logger):
    import re
    split=re.split('[ ]*\\([0-9\\.\\-: ]+\\)[:]*', line)
    if len(split) != 2:
        logger.append("Not counted (irrelevant?):" + line)
        return None
    if len(split[0].strip())!=1:
        logger.append("Not counted (error):" + line)
        return None
    return (split[0], split[1],)

def process_lines(lines, loglines):
    result={}
    for peop in people:
        result[peop]=0
    
    for line in lines:
        if not line:
            continue
        p=parse_header(line, loglines)
        if not p:
            continue
        if p[0] in people:
            result[p[0]]+=count_words(p[1])

    return result

def process(srcurl):
    dump_report(srcurl,process_lines(get_lines(docx_to_xml(srcurl))))

if __name__=="__main__":
    print("Script started")
    docs=list_docs(choose_folder())
    for doc in docs:
        lines=[]
        reportfile=doc.rsplit(".")[0]+"_report.txt"
        reportlines=[]
        dump_report(doc, process_lines(get_lines(docx_to_xml(doc)),lines), reportlines)

        with open(reportfile, 'w') as f:
            for line in reportlines:
                f.write(line)
                f.write("\n")
            f.write("\n\nIssues with document:\n\n")
            for line in lines:
                f.write(line)
                f.write("\n")
    print("Script finished")
