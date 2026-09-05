#!/usr/bin/env python3
"""Check portable packaging; optional private keywords are never printed."""
import argparse
from pathlib import Path
import re
import subprocess
import sys

ROOT = Path(__file__).resolve().parents[1]


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--private-keywords', type=Path)
    args = parser.parse_args()
    errors = []
    keyword_lines = []
    if args.private_keywords:
        if not args.private_keywords.expanduser().is_file():
            raise SystemExit('Private keyword file missing; privacy scan not completed')
        keyword_lines = [s.strip() for s in args.private_keywords.expanduser().read_text().splitlines()
                         if s.strip() and not s.lstrip().startswith('#')]
    if (ROOT / '.git').exists():
        paths = subprocess.check_output(['git', '-C', str(ROOT), 'ls-files', '--cached', '--others', '--exclude-standard'], text=True).splitlines()
    else:
        paths = [str(p.relative_to(ROOT)) for p in ROOT.rglob('*') if p.is_file()
                 and not any(part in {'.git', 'node_modules', '__pycache__', '.venv'} for part in p.relative_to(ROOT).parts)]
    for relative in sorted(set(paths)):
        path = ROOT / relative
        if not path.is_file():
            continue
        if path.suffix.lower() in {'.pdf', '.docx', '.pptx', '.png', '.jpg'}:
            errors.append(relative + ': private/output artifact')
        text = path.read_text(encoding='utf-8', errors='replace')
        if any(word in text for word in keyword_lines):
            errors.append(relative + ': private keyword match (value withheld)')
        if path.suffix == '.md' and relative != 'prompt/SKILL.md':
            for target in re.findall(r'\[[^\]]+\]\(([^)]+)\)', text):
                if '://' in target or target.startswith('#'):
                    continue
                if not (path.parent / target.split('#')[0]).exists():
                    errors.append(relative + ': broken relative link')
    skill = (ROOT / 'SKILL.md').read_text()
    if not skill.startswith('---\nname: flight-theory-notes\n'):
        errors.append('Invalid root Skill metadata')
    result = subprocess.run([sys.executable, str(ROOT / 'tools/export_prompt.py'), '--check'])
    if result.returncode:
        errors.append('Stale prompt export')
    if errors:
        raise SystemExit('\n'.join(errors))
    print('Package checks passed; ' + ('private-keyword scan completed' if args.private_keywords else 'private-keyword scan NOT run'))


if __name__ == '__main__':
    main()
