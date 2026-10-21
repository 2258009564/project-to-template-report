"""Check the actual Git index, or source files before Git initialization."""
from pathlib import Path
import ast
import re
import subprocess


def main():
    root = Path(__file__).resolve().parents[1]
    if (root / '.git').exists():
        raw = subprocess.check_output(['git', 'ls-files', '-z'], cwd=root)
        files = [root / p for p in raw.decode('utf-8').split('\0') if p]
    else:
        files = [p for p in root.rglob('*') if p.is_file() and '__pycache__' not in p.parts]
    assert files, 'Empty package'
    for path in files:
        rel = path.relative_to(root)
        assert not {'memories', 'private', '__pycache__'} & set(rel.parts), rel
        assert path.suffix in {'.md', '.py', '.ps1', '.yaml', '.yml'} or path.name in {'.gitignore', 'LICENSE'}, rel
        text = path.read_text(encoding='utf-8-sig')
        assert '\x1b' not in text, rel
        for pattern in (r'[A-Za-z]:[/\\]Users[/\\][^/\\\s]+', r'gh[pousr]_[A-Za-z0-9]{30,}', r'github_pat_[A-Za-z0-9_]{30,}', r'-----BEGIN (?:RSA |OPENSSH )?PRIVATE KEY-----'):
            assert not re.search(pattern, text), f'Potential private data: {rel}'
        if path.suffix == '.py':
            ast.parse(text, filename=str(rel))
        if path.suffix == '.md':
            for link in re.findall(r'\]\(([^)]+)\)', text):
                if '://' not in link and not link.startswith('#'):
                    assert (path.parent / link.split('#')[0]).exists(), (rel, link)
    print(f'PASS: {len(files)} package files, references, Python syntax and basic privacy patterns')


if __name__ == '__main__':
    main()
