import json,re
from pathlib import Path
from docx import Document
from docx.shared import Cm,Pt,RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT,WD_CELL_VERTICAL_ALIGNMENT

ROOT=Path(__file__).resolve().parent
OUT=ROOT.parent/'output'
OUT.mkdir(exist_ok=True)

def load(name): return json.loads((ROOT/name).read_text(encoding='utf-8'))
def clean_heading(text):
    text=re.sub(r'(?<=\d)[—–-](?=\d)','至',str(text))
    return re.sub(r'\s+',' ',re.sub(r'[^\w\s\u4e00-\u9fff]',' ',text).replace('_',' ')).strip()
def font(style,size,bold=False):
    style.font.name='Calibri';style.font.size=Pt(size);style.font.bold=bold;style.font.italic=False;style.font.color.rgb=RGBColor(0,0,0)
    style.element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),'宋体')

def make(title,intro):
    d=Document();s=d.sections[0];s.page_height=Cm(29.7);s.page_width=Cm(21)
    s.top_margin=Cm(1.9);s.bottom_margin=Cm(1.8);s.left_margin=Cm(2.1);s.right_margin=Cm(2.1)
    s.header_distance=Cm(.7);s.footer_distance=Cm(.8)
    font(d.styles['Normal'],10.5)
    pf=d.styles['Normal'].paragraph_format;pf.line_spacing=1.22;pf.space_after=Pt(4);pf.widow_control=True
    for name,size in [('Title',20),('Subtitle',11),('Heading 1',15),('Heading 2',12.5),('Heading 3',11)]:
        font(d.styles[name],size,name!='Subtitle');d.styles[name].paragraph_format.space_before=Pt(12 if 'Heading' in name else 0);d.styles[name].paragraph_format.space_after=Pt(6)
    for sty in d.styles:
        for border in list(sty.element.findall('.//'+qn('w:pBdr'))):border.getparent().remove(border)
    d.core_properties.author='';d.core_properties.title=title;d.core_properties.subject='试题解析与学习资料'
    d.add_paragraph(clean_heading(title),style='Title')
    d.add_paragraph('2025年国家金融监督管理总局考试 计算机岗 考生回忆版',style='Subtitle')
    for p in (intro if isinstance(intro,list) else [intro]):
        d.add_paragraph(p)
    p=s.footer.paragraphs[0];p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run();r.font.size=Pt(9)
    f=OxmlElement('w:fldSimple');f.set(qn('w:instr'),'PAGE');r._r.addnext(f)
    return d

def heading(d,text,level=1):
    return d.add_heading(clean_heading(text),level)

def para(d,text,label=None,keep=False):
    if text is None:return
    if isinstance(text,(list,tuple)):
        for x in text:para(d,x)
        return
    p=d.add_paragraph();p.paragraph_format.keep_with_next=keep
    if label:p.add_run(label+'  ').bold=True
    p.add_run(str(text));return p

def hyperlink(p,text,url):
    rel=p.part.relate_to(url,'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',is_external=True)
    h=OxmlElement('w:hyperlink');h.set(qn('r:id'),rel)
    r=OxmlElement('w:r');pr=OxmlElement('w:rPr');c=OxmlElement('w:color');c.set(qn('w:val'),'1F4E79');pr.append(c);r.append(pr)
    t=OxmlElement('w:t');t.text=text;r.append(t);h.append(r);p._p.append(h)

def refs(d,sources):
    if not sources:return
    heading(d,'核验依据',1)
    para(d,'以下资料用于核对相关题目的概念、条款或技术边界。法规题以题目所对应的考试时点和条款为学习语境。链接可在Word中打开。')
    for i,src in enumerate(sources,1):
        p=d.add_paragraph();p.add_run(f'[{i}] ');hyperlink(p,src['title'],src['url'])

def keylines(d,qs):
    disputed=any(q.get('answer',q['key'])!=q['key'] for q in qs)
    heading(d,'原卷答案速查' if disputed else '答案速查',1)
    for start in range(0,len(qs),10):
        para(d,'　'.join(f"{q['id']} {q['key']}"+('*' if q.get('answer',q['key'])!=q['key'] else '') for q in qs[start:start+10]))
    para(d,'本表保留原卷给定答案。带星号的题存在适用条件或复核差异，复核建议和依据见正文。' if disputed else '速查采用本册复核结论；条件不充分或原答案有争议的题，必须阅读正文说明。')

def build_questions(data,filename,computer=False):
    d=make(data['title'],data['intro']);qs=data['questions']
    if computer:
        heading(d,'学习与复盘建议')
        para(d,'本册覆盖第21至100题，含单选60题和多选20题。每题提供题干选项、原卷答案、你的作答、复核结论、逐项解析和知识点。按原卷给定答案，你答对45题、答错35题；该统计用于定位原作答记录，争议题请以本册说明单独复核。')
        para(d,'建议按软件工程与信息系统、数据结构与数据库、操作系统与网络、安全与金融科技风险五条线复习。先完成错题，再对照知识点解释所有选项。多选题特别检查“不符合”“主要”“可能”等限制。')
        conflicts=[q for q in qs if q.get('answer')!=q['key'] or any(w in q.get('explanation','') for w in ['争议','条件不足','题干不足'])]
        if conflicts:para(d,'重点复核题号：'+ '、'.join(str(q['id']) for q in conflicts)+'。这些题的术语、条件或原给答案需要结合正文理解。')
    for sec in data.get('roadmap',[]):
        heading(d,sec['heading']);para(d,sec['paragraphs'])
    keylines(d,qs)
    sources=[];index={}
    heading(d,'逐题解析')
    for q in qs:
        typ='多选题' if 81<=q['id']<=100 else '单选题'
        heading(d,f"第{q['id']}题 {typ}",2)
        para(d,q['stem'],keep=True)
        for letter,option in q['options'].items():para(d,f'{letter}  {option}',keep=letter!='D')
        answer=q.get('answer',q['key']);flag='与原卷一致' if q['user']==q['key'] else '原作答错误'
        para(d,f"原卷给定答案 {q['key']}　你的答案 {q['user']}　{flag}",label='作答记录')
        para(d,str(answer),label='复核结论')
        para(d,q['explanation'],label='解题分析')
        for letter,t in q['option_explanations'].items():para(d,t,label=f'选项{letter}')
        heading(d,'知识点讲解',3)
        para(d,q['knowledge']);para(d,q['pitfall'],label='易错提醒',keep=bool(q.get('sources')))
        nums=[]
        for src in q.get('sources',[]):
            if src['url'] not in index:index[src['url']]=len(sources)+1;sources.append(src)
            nums.append(str(index[src['url']]))
        if nums:para(d,'、'.join('['+n+']' for n in nums),label='核验依据')
    refs(d,sources);path=OUT/filename;d.save(path);print(path)

def build_english():
    data=load('english.json');d=make(data['title'],data['intro'])
    if data.get('stats'):
        st=data['stats']
        if isinstance(st,dict):
            para(d,f"共{st['total']}题，按原卷给定答案答对{st['original_key_correct']}题，正确率{st['original_key_rate']}。按本文复核建议答对{st['recommended_key_correct']}题，正确率{st['recommended_key_rate']}。{st.get('note','')}",label='作答复盘')
        else:para(d,st,label='作答复盘')
    for ps in data['passages']:
        heading(d,f"第{ps['id']}篇 {ps['title']}")
        para(d,f"对应原题 {ps['range']}",label='范围')
        if ps.get('source_note'):para(d,ps['source_note'],label='阅读说明')
        heading(d,'原文与分段翻译',2)
        for p in ps['paragraphs']:
            heading(d,f"第{p['id']}段",3);para(d,p['en']);para(d,p['zh'],label='译文')
            if p.get('role'):para(d,p['role'],label='段落作用')
        heading(d,'篇章结构',2);para(d,ps['structure'])
        for q in ps['questions']:
            heading(d,f"第{q['id']}题 {q['type']}",2)
            para(d,q['stem_en'],keep=True);para(d,q['stem_zh'],label='题意')
            for op in q['options']:
                para(d,f"{op['letter']}  {op['en']}");para(d,op['zh'],label='选项译文')
            para(d,f"原卷给定答案 {q['original_key']}　你的答案 {q['user_answer']}　复核建议 {q['recommended_key']}",label='答案')
            para(d,q['evidence'],label='原文定位');para(d,q['reasoning'],label='解题分析')
            for op in q['options']:para(d,f"{op['verdict']}。{op['analysis']}",label=f"选项{op['letter']}")
            para(d,q['method'],label='方法');para(d,q['user_review'],label='你的作答复盘')
            if q.get('caution'):para(d,q['caution'],label='需注意')
    path=OUT/'03_英语逐题解析.docx';d.save(path);print(path)

def build_vocab():
    data=load('vocabulary.json');d=make(data['title'],data['intro'])
    for sec in data['sections']:
        heading(d,sec['title'])
        heading(d,'单词',2)
        for x in sec['words']:para(d,f"{x['en']}  {x.get('pos','')}  {x['zh']}")
        heading(d,'常用词汇组合与搭配',2)
        for x in sec['phrases']:para(d,f"{x['en']}  {x['zh']}")
    path=OUT/'04_英语单词与常用搭配.docx';d.save(path);print(path)

def comparison_table(d,data):
    headers=data['headers'];rows=data['rows'];n=len(headers)
    t=d.add_table(rows=1,cols=n);t.autofit=False;t.alignment=WD_TABLE_ALIGNMENT.CENTER
    widths={2:[5,11.8],3:[3.3,6.75,6.75],4:[2.6,4.2,5,5]}.get(n,[16.8/n]*n)
    for c,w in zip(t.columns,widths):c.width=Cm(w)
    for c,w in zip(t.rows[0].cells,widths):c.width=Cm(w)
    tr=t.rows[0]._tr.get_or_add_trPr();repeat=OxmlElement('w:tblHeader');repeat.set(qn('w:val'),'true');tr.append(repeat)
    for c,txt in zip(t.rows[0].cells,headers):c.text=str(txt)
    for row in rows:
        vals=list(row.values()) if isinstance(row,dict) else row
        cs=t.add_row().cells
        for c,txt,w in zip(cs,vals,widths):c.text=str(txt);c.width=Cm(w)
    for ri,row in enumerate(t.rows):
        split=OxmlElement('w:cantSplit');row._tr.get_or_add_trPr().append(split)
        for c in row.cells:
            c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
            pr=c._tc.get_or_add_tcPr();b=OxmlElement('w:tcBorders')
            for side in ['top','left','bottom','right','insideH','insideV']:
                e=OxmlElement('w:'+side);e.set(qn('w:val'),'single');e.set(qn('w:sz'),'4');e.set(qn('w:color'),'D9D9D9');b.append(e)
            pr.append(b);m=OxmlElement('w:tcMar')
            for side in ['top','left','bottom','right']:
                e=OxmlElement('w:'+side);e.set(qn('w:w'),'80');e.set(qn('w:type'),'dxa');m.append(e)
            pr.append(m)
            if ri==0:
                shade=OxmlElement('w:shd');shade.set(qn('w:fill'),'E8EFF7');pr.append(shade)
            for p in c.paragraphs:
                p.paragraph_format.line_spacing=1.1;p.paragraph_format.space_before=Pt(2);p.paragraph_format.space_after=Pt(2)
                for r in p.runs:r.font.size=Pt(9.5);r.bold=ri==0
    d.add_paragraph().paragraph_format.space_after=Pt(2)

def anycontent(d,x):
    if x is None:return
    if isinstance(x,str):para(d,x)
    elif isinstance(x,list):
        for z in x:anycontent(d,z)
    elif isinstance(x,dict):
        if 'headers' in x and 'rows' in x:
            comparison_table(d,x);return
        for k,v in x.items():
            if k in ['title','heading']:heading(d,v,3)
            elif k in ['en','zh','question','answer','explanation','trap','example','method','type','stem','prompt','analysis','reason','source','steps','signal','principle','check','text','key','sentence','translation']:
                labels={'en':'英文','zh':'中文','question':'题目','answer':'答案','explanation':'解析','trap':'陷阱','example':'例子','method':'方法','type':'题型','stem':'题干','prompt':'题目','analysis':'解析','reason':'原因','source':'来源','steps':'步骤','signal':'识别信号','principle':'原理','check':'检查','text':'内容','key':'答案','sentence':'原句','translation':'译文'}
                if isinstance(v,str):para(d,v,label=labels[k])
                else:anycontent(d,v)
            else:anycontent(d,v)

def build_patterns():
    data=load('reading_patterns.json');d=make(data['title'],data['intro'])
    for sec in data['sections']:
        heading(d,sec['title'])
        for k,v in sec.items():
            if k!='title':anycontent(d,v)
    path=OUT/'05_英语篇章阅读题型与设题陷阱.docx';d.save(path);print(path)

if __name__=='__main__':
    import sys
    mode=sys.argv[1] if len(sys.argv)>1 else 'all'
    if mode in ['all','other']:build_questions(load('other.json'),'01_其他模块经济金融基础解析.docx')
    if mode in ['all','computer']:
        a=load('computer_a.json');b=load('computer_b.json');a['questions']+=b['questions'];a['title']='2025年金融监管总局计算机岗 计算机模块逐题解析';a['intro']='本册围绕80道计算机专业题，将选项判断与知识体系结合，帮助你在订正错题后掌握同类题的推理方法。';build_questions(a,'02_计算机模块逐题解析.docx',True)
    if mode in ['all','english']:build_english();build_vocab();build_patterns()
