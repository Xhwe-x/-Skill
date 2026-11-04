"""Text-cue checks only; not semantic, legal, scoring or evidence verification."""
import argparse
import collections
import json
import re
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path


def read_text(path):
    p = Path(path)
    if p.suffix.lower() != '.docx': return p.read_text(encoding='utf-8')
    with zipfile.ZipFile(p) as z:
        root = ET.fromstring(z.read('word/document.xml'))
    ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    return '\n'.join(''.join(n.text or '' for n in para.iter(ns + 't')) for para in root.iter(ns + 'p'))


def check(text, manifest, mode='draft'):
    issues = []
    def add(level, code, message): issues.append({'level': level, 'code': code, 'message': message})
    if manifest.get('project_name') and manifest['project_name'] not in text:
        add('REVIEW', 'PROJECT_NAME', '未找到当前项目全称')
    for value in manifest.get('required_literals', []):
        if value not in text: add('REVIEW', 'REQUIRED_CUE', '未找到约定文字线索：' + value)
    for value in manifest.get('forbidden_literals', []):
        if value and value in text: add('BLOCKER', 'FORBIDDEN_CUE', '出现本项目禁止残留：' + value)
    for req in manifest.get('requirements', []):
        cues = req.get('keywords', [])
        if cues and not any(cue in text for cue in cues):
            add('REVIEW', 'REQUIREMENT_CUE', '需人工检查 ' + req.get('id', '') + ' ' + req.get('title', ''))
    placeholders = re.findall(r'【[^】]*(?:待填|待补|待核|待确认)[^】]*】|\bTODO\b|\{\{[^}]+\}\}', text)
    if placeholders: add('BLOCKER' if mode == 'final' else 'REVIEW', 'PLACEHOLDERS', '存在待补标记：' + '、'.join(sorted(set(placeholders))[:8]))
    if manifest.get('project_type') == 'design-consulting' and re.search(r'(?:我方|本公司|投标人)(?:将|应|须)?(?:负责|承担)(?:本项目)?(?:全部|所有)?(?:设备采购|工程施工|软件开发实施)', text):
        add('REVIEW', 'SCOPE', '检查设计咨询是否承担了额外采购、施工或开发义务')
    counts = collections.Counter(line.strip() for line in text.splitlines() if len(line.strip()) >= 80)
    dup = [s[:60] for s, n in counts.items() if n > 1]
    if dup: add('REVIEW', 'DUPLICATES', '重复长段落：' + '；'.join(dup[:5]))
    if len(text.strip()) < 100: add('REVIEW', 'SHORT_TEXT', '正文过短，检查是否提取或生成不完整')
    return {'check_type': 'text-cues-only', 'requires_semantic_review': True, 'issues': issues,
            'limitation': '不证明语义覆盖、技术可行性、证据真实性、法规适用、评审得分或实际页数'}


def main():
    p = argparse.ArgumentParser()
    p.add_argument('--manifest', required=True); p.add_argument('--draft', required=True)
    p.add_argument('--mode', choices=['draft', 'final'], default='draft'); p.add_argument('--out')
    a = p.parse_args()
    result = check(read_text(a.draft), json.loads(Path(a.manifest).read_text(encoding='utf-8')), a.mode)
    output = json.dumps(result, ensure_ascii=False, indent=2)
    if a.out: Path(a.out).write_text(output, encoding='utf-8')
    else: print(output)
    raise SystemExit(2 if any(i['level'] == 'BLOCKER' for i in result['issues']) else 0)


if __name__ == '__main__': main()
