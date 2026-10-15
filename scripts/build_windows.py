"""Build and verify a data-free Windows distribution. Run from any directory."""
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import zipfile

ROOT = Path(__file__).resolve().parents[1]
BUILD = ROOT / '.build'
RELEASE = ROOT / 'release'


def run(args, **kwargs):
    subprocess.run([str(x) for x in args], cwd=ROOT, check=True, **kwargs)


def verify_report(folder):
    report = json.loads((folder / 'regression.json').read_text(encoding='utf-8'))
    if not report['success'] or report['run'] < 77:
        raise RuntimeError(report['details'])
    print(f"Passed {report['run']} regression tests", flush=True)


def archive(destination, files):
    with zipfile.ZipFile(destination, 'w', zipfile.ZIP_DEFLATED) as z:
        for source, name in files:
            z.write(source, 'Матучет/' + name)


def main():
    if sys.platform != 'win32':
        raise SystemExit('Build on Windows with Python x64.')
    version = (ROOT / 'VERSION').read_text(encoding='utf-8').strip()
    import re
    if not re.fullmatch(r'\d+\.\d+\.\d+', version):
        raise ValueError('VERSION must contain a numeric three-part version')
    seed = json.loads((ROOT / 'seed.json').read_text(encoding='utf-8'))
    if any(seed[k] for k in ['equipment', 'employees', 'objects', 'issues', 'legacy']):
        raise ValueError('Public builds require empty initial data')
    if f"'version':'{version}'" not in (ROOT / 'store.py').read_text(encoding='utf-8'):
        raise ValueError('Update the application version in store.py')
    BUILD.mkdir(exist_ok=True)
    # This directory is exclusively owned by this script; never use a user data path.
    qa = BUILD / 'verification'
    if qa.exists():
        shutil.rmtree(qa)
    qa.mkdir()
    env = dict(os.environ, MATUCHET_PREFERENCES=str(qa / 'appearance.json'))
    source_tests = qa / 'source'
    source_tests.mkdir()
    env['MATUCHET_DATA'] = str(source_tests)
    run([sys.executable, ROOT / 'main.py', '--self-test'], env=env)
    verify_report(source_tests)
    # Check empty startup and settings against the shipped seed.
    sys.path.insert(0, str(ROOT))
    from store import Store
    store = Store(qa / 'empty', ROOT)
    snapshot = store.snapshot()
    if any(snapshot[k] for k in ['assets', 'employees', 'acts', 'history']):
        raise ValueError('New installations must have no working data')
    run([sys.executable, '-m', 'PyInstaller', '--noconfirm', '--clean', '--onedir',
         '--windowed', '--name', 'Матучет', '--distpath', BUILD / 'dist',
         '--workpath', BUILD / 'pyinstaller', '--specpath', BUILD,
         '--add-data', f'{ROOT / "ui"};ui',
         '--add-data', f'{ROOT / "seed.json"};.',
         '--add-data', f'{ROOT / "Шаблон_акта.docx"};.',
         '--add-data', f'{ROOT / "Шаблон_приема_передачи.docx"};.',
         '--hidden-import', '_cffi_backend', '--collect-all', 'webview', ROOT / 'main.py'], env=env)
    dist = BUILD / 'dist' / 'Матучет'
    frozen_tests = qa / 'frozen'
    frozen_tests.mkdir()
    env['MATUCHET_DATA'] = str(frozen_tests)
    run([dist / 'Матучет.exe', '--self-test'], env=env, timeout=180)
    verify_report(frozen_tests)
    resources = dist / '_internal'
    if json.loads((resources / 'seed.json').read_text(encoding='utf-8')) != seed:
        raise ValueError('Packaged seed mismatch')
    files = [(p, p.relative_to(dist).as_posix()) for p in sorted(dist.rglob('*')) if p.is_file()]
    if any(p.suffix.lower() in ['.sqlite', '.db', '.log'] or 'Данные' in p.parts for p, _ in files):
        raise ValueError('Unexpected working data in distribution')
    files += [(ROOT / name, name) for name in ['README.md', 'CHANGELOG.md', 'VERSION']]
    RELEASE.mkdir(exist_ok=True)
    outputs = [RELEASE / f'Matuchet-{version}-{kind}.zip' for kind in ['Windows', 'Update']]
    # Neither archive includes a user data directory. Different names clarify installation intent.
    for path in outputs:
        archive(path, files)
    (RELEASE / 'SHA256SUMS.txt').write_text(''.join(
        hashlib.sha256(p.read_bytes()).hexdigest() + '  ' + p.name + '\n' for p in outputs), encoding='utf-8')
    print(f'Release {version} ready in {RELEASE}', flush=True)


if __name__ == '__main__':
    main()
