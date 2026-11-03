import copy
import hashlib
import json
import re
import shutil
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent
SOURCE = ROOT.parent / 'output' / '04_英语单词与常用搭配.docx'
STAGED = ROOT / 'vocab_ipa_revised.docx'
BACKUP = ROOT / 'vocab_before_ipa_edit.docx'


def main():
    entries = {}
    for name in ['ipa_a.json', 'ipa_b.json', 'ipa_c.json']:
        part = json.loads((ROOT / name).read_text(encoding='utf-8'))['entries']
        assert not (entries.keys() & part.keys()), 'Overlapping pronunciation records'
        entries.update(part)
    current = json.loads((ROOT / 'ipa_current_words.json').read_text(encoding='utf-8'))
    before_hash = hashlib.sha256(SOURCE.read_bytes()).hexdigest()
    doc = Document(SOURCE)
    before_text = [p.text for p in doc.paragraphs]
    shutil.copy2(SOURCE, BACKUP)
    changes = []
    used = set()
    for item in current:
        paragraph = doc.paragraphs[item['i']]
        assert paragraph.text == item['text'], 'Source changed during pronunciation lookup'
        text = paragraph.text
        head = re.match(r'[A-Za-z][A-Za-z-]*', text).group()
        forms = re.findall(r'([A-Za-z][A-Za-z-]*)(?=（)', text)
        for word in [head] + forms:
            assert word in entries, f'Missing pronunciation for {word}'
            used.add(word)
            record = entries[word]
            if word == 'implant':
                variants = record['variants']
                text = f"implant  n. /{variants['n.']}/ 植入物；v. /{variants['v.']}/ 植入"
                break
            ipa = record['ipa'].strip('/').replace('(r)', '')
            text = re.sub(r'(?<![A-Za-z])' + re.escape(word) + r'(?![A-Za-z])', lambda m: m.group() + ' /' + ipa + '/', text, count=1)
        replace_text(paragraph, text)
        changes.append({'index': item['i'], 'before': item['text'], 'after': text})
    intro = {
        2: '单词先列原形、英式音标、词性与汉语义，再在同一词条中保留原文中的 -ing、-ed 或相关派生形式及其音标、释义。词汇搭配按完整表达呈现。',
        3: '音标采用英式读音；ˈ 表示主重音，ˌ 表示次重音。n.=名词；v.=动词；adj.=形容词；adv.=副词。-ing、-ed 表示词形，词性按本篇用法标注。implant 的名词与动词重音不同，已分别列出。',
    }
    for i, text in intro.items():
        changes.append({'index': i, 'before': doc.paragraphs[i].text, 'after': text})
        replace_text(doc.paragraphs[i], text)
    assert len(doc.paragraphs) == len(before_text)
    changed = {x['index'] for x in changes}
    assert all(p.text == before_text[i] for i, p in enumerate(doc.paragraphs) if i not in changed)
    doc.save(STAGED)
    manifest = {'source': str(SOURCE), 'source_sha256': before_hash, 'staged': str(STAGED),
                'word_lines': len(current), 'pronunciation_entries': len(used), 'changes': changes,
                'sources': {word: entries[word] for word in sorted(used)}}
    (ROOT / 'vocab_ipa_edit_manifest.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f'Staged {len(current)} word lines with {len(used)} pronunciation entries at {STAGED}')


def replace_text(paragraph, text):
    original_rpr = None
    if paragraph.runs and paragraph.runs[0]._r.rPr is not None:
        original_rpr = copy.deepcopy(paragraph.runs[0]._r.rPr)
    paragraph.clear()
    # Arial contains the IPA glyphs needed here and is present on this Windows host.
    for chunk in re.split(r'(/[^/\n]*[ɑ-ʯˈˌ][^/\n]*/)', text):
        if not chunk:
            continue
        run = paragraph.add_run(chunk)
        if original_rpr is not None:
            run._r.insert(0, copy.deepcopy(original_rpr))
        if chunk.startswith('/') and chunk.endswith('/'):
            run.font.name = 'Arial'
            run._element.get_or_add_rPr().rFonts.set(qn('w:hAnsi'), 'Arial')


if __name__ == '__main__':
    main()
