"""Prepare only verified bundled .NET libraries extracted from a downloaded ZIP."""
import hashlib
import json
import os
from pathlib import Path
import sys


def prepare(resources=None):
    if os.name != 'nt' or (resources is None and not getattr(sys, 'frozen', False)):
        return []
    root = Path(resources or sys._MEIPASS).resolve()
    manifest = json.loads((root / 'runtime_manifest.json').read_text(encoding='utf-8'))
    if manifest.get('format') != 1 or not manifest.get('files'):
        raise ValueError('Поврежден список компонентов программы. Распакуйте полный архив заново.')
    files = []
    for name, expected in manifest['files'].items():
        relative = Path(name)
        allowed = name == 'pythonnet/runtime/Python.Runtime.dll' or (
            name.startswith('webview/lib/') and relative.suffix.lower() == '.dll')
        path = (root / relative).resolve()
        if not allowed or not path.is_relative_to(root) or '..' in relative.parts:
            raise ValueError('Недопустимый путь компонента программы.')
        if not path.is_file() or hashlib.sha256(path.read_bytes()).hexdigest() != expected:
            raise ValueError('Компонент программы отсутствует или отличается от сборки: ' + name +
                             '. Распакуйте полный архив в новую папку.')
        files.append((name, path))
    unblocked = []
    # Do not change system policies or unblock arbitrary files/user documents.
    # Explorer may propagate the ZIP's Internet-zone marker onto these exact DLLs.
    for name, path in files:
        marker = Path(str(path) + ':Zone.Identifier')
        try:
            marker.unlink()
            unblocked.append(name)
        except FileNotFoundError:
            pass
        except OSError as exc:
            raise ValueError('Windows блокирует компонент ' + name +
                '. Откройте свойства скачанного ZIP, нажмите «Разблокировать» и распакуйте архив заново в доступную для записи папку.') from exc
    return unblocked


def check(resources, output):
    """Load the real GUI dependencies, without opening a window or a user database."""
    import traceback
    report = {'success': False}
    try:
        report['unblocked'] = prepare(resources)
        import clr
        clr.AddReference('System.Windows.Forms')
        from webview.util import interop_dll_path
        for name in ['Microsoft.Web.WebView2.Core.dll', 'Microsoft.Web.WebView2.WinForms.dll']:
            clr.AddReference(interop_dll_path(name))
        from Microsoft.Web.WebView2.WinForms import WebView2
        report['success'] = WebView2 is not None
    except Exception:
        report['error'] = traceback.format_exc()
    target = Path(output)
    target.mkdir(parents=True, exist_ok=True)
    (target / 'runtime-check.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    return 0 if report['success'] else 1
