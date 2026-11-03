import copy
import hashlib
import json
import shutil
from pathlib import Path

from docx import Document

ROOT = Path(__file__).resolve().parent
SOURCE = ROOT.parent / 'output' / '04_英语单词与常用搭配.docx'
STAGED = ROOT / 'vocab_forms_revised.docx'
BACKUP = ROOT / 'vocab_before_forms_edit.docx'

# Only normalize actual inflections and transparent, useful adjective families.
# Independent dictionary headwords are not reduced to uncommon etymological roots.
REPLACEMENTS = {
    'creased': 'crease  v. 起皱；使有折痕；creased（-ed / adj.）有折痕的',
    'clinking': 'clink  v. 叮当作响；clinking（-ing，作定语）叮当作响的',
    'endangered': 'endanger  v. 危及；endangered（-ed / adj.）濒危的',
    'plumbing': 'plumb  v. 装设水管；plumbing（n.）管道系统；本篇指支付基础设施',
    'dehumanizing': 'dehumanize  v. 使失去人性；dehumanizing（-ing / adj.）非人性化的',
    'intrusive': 'intrude  v. 侵扰；intrusive（adj.）侵扰的；侵犯私人空间的',
    'bullying': 'bully  v. 欺凌；bullying（-ing / n.）欺凌',
    'undetected': 'detect  v. 发现；undetected（adj.）未被发现的',
    'anonymised': 'anonymise  v. 使匿名化；anonymised（-ed / adj.）经匿名化处理的',
    'entitled': 'entitle  v. 使享有权利；entitled（adj.）有权享有的（选项）',
    'emerging': 'emerge  v. 出现；emerging（-ing / adj.）新出现的',
    'philosophical': 'philosophy  n. 哲学；philosophical（adj.）哲学的',
    'systemic': 'system  n. 系统；systemic（adj.）系统性的；结构性的',
    'preceding': 'precede  v. 在……之前；preceding（-ing / adj.）先前的',
    'universally': 'universal  adj. 普遍的；universally（adv.）普遍地；一致地（选项）',
    'uniformly': 'uniform  adj. 统一的；uniformly（adv.）统一地；一致地（选项）',
    'complementary': 'complement  v. 补充；complementary（adj.）互补的（选项）',
}

INTRO_REPLACEMENTS = {
    2: '单词先列原形或常用基础词，再在同一行列出原文中的 -ing、-ed 或相关形容词、副词形式，并分别给出汉语解释。已是原形的词直接保留；词汇搭配按完整表达呈现。',
    3: 'n.=名词；v.=动词；adj.=形容词；adv.=副词。-ing、-ed 表示词形，词性按本篇用法标注；-ing 形式也可作名词或定语。原文 dechumanizing 按规范 dehumanizing 收录。',
}


def replace_preserving_format(paragraph, text):
    # Preserve the paragraph properties and the first run's formatting.
    template_rpr = None
    if paragraph.runs and paragraph.runs[0]._r.rPr is not None:
        template_rpr = copy.deepcopy(paragraph.runs[0]._r.rPr)
    paragraph.clear()
    run = paragraph.add_run(text)
    if template_rpr is not None:
        run._r.insert(0, template_rpr)


def main():
    original_hash = hashlib.sha256(SOURCE.read_bytes()).hexdigest()
    shutil.copy2(SOURCE, BACKUP)
    doc = Document(SOURCE)
    original = [p.text for p in doc.paragraphs]
    changes = []
    found = set()
    in_words = False
    for i, paragraph in enumerate(doc.paragraphs):
        if i in INTRO_REPLACEMENTS:
            new_text = INTRO_REPLACEMENTS[i]
        else:
            if paragraph.style.name == 'Heading 2':
                in_words = paragraph.text.strip() == '单词'
            if not in_words or paragraph.style.name != 'Normal':
                continue
            word = paragraph.text.split()[0] if paragraph.text.strip() else ''
            if word not in REPLACEMENTS:
                continue
            new_text = REPLACEMENTS[word]
            found.add(word)
        changes.append({'index': i, 'before': paragraph.text, 'after': new_text})
        replace_preserving_format(paragraph, new_text)
    assert found == set(REPLACEMENTS), set(REPLACEMENTS) - found
    assert len(doc.paragraphs) == len(original)
    changed_indices = {c['index'] for c in changes}
    assert all(p.text == original[i] for i, p in enumerate(doc.paragraphs) if i not in changed_indices)
    doc.save(STAGED)
    verify = Document(STAGED)
    assert len(verify.paragraphs) == len(original)
    manifest = {'source': str(SOURCE), 'source_sha256': original_hash, 'staged': str(STAGED),
                'paragraphs': len(original), 'word_lines_updated': len(found), 'changes': changes}
    (ROOT / 'vocab_forms_edit_manifest.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f'Updated {len(found)} word lines; retained {len(original)} paragraphs; staged at {STAGED}')


if __name__ == '__main__':
    main()
