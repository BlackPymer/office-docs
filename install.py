#!/usr/bin/env python3
"""Установка office-docs: venv со зависимостями. Одинаково на Windows и POSIX.

Использование:
    python install.py                     # Linux/macOS
    py -3 install.py                      # Windows
    python install.py --with-markitdown   # + markitdown (docx/pptx/pdf -> md)
    python install.py --venv ГДЕ          # свой каталог venv

Что делает:
  1. проверяет версию Python (нужен 3.10+),
  2. создаёт venv (по умолчанию <папка скилла>/.venv),
  3. ставит requirements.txt (и markitdown, если попросили),
  4. проверяет импорты и компиляцию скриптов,
  5. ищет LibreOffice и подсказывает, что делать, если его нет.

LibreOffice и pandoc этим скриптом НЕ ставятся: они нужны только для
конвертации форматов и проверки рендера (см. README.md).

Файл сознательно написан так, чтобы его можно было запустить и Python 2 —
тогда вместо стека исключений будет понятное «нужен Python 3.10+».
"""

from __future__ import print_function

import argparse
import os
import shutil
import subprocess
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
MIN_VERSION = (3, 10)
CORE = ("docx", "pptx", "openpyxl", "pymupdf", "pypdf", "PIL", "lxml")
MARKITDOWN_SPEC = "markitdown==0.1.8"

CHECK_IMPORTS = (
    "import importlib\n"
    "bad = []\n"
    "for m in %r:\n"
    "    try:\n"
    "        importlib.import_module(m)\n"
    "    except Exception as e:\n"
    "        bad.append('%%s: %%s' %% (m, e))\n"
    "print('OK' if not bad else 'BAD: ' + ' | '.join(bad))\n"
) % (list(CORE),)


def is_windows():
    return os.name == "nt"


def venv_python(venv):
    """Путь к интерпретатору внутри venv (Windows раскладывает его в Scripts/)."""
    if is_windows():
        return os.path.join(venv, "Scripts", "python.exe")
    return os.path.join(venv, "bin", "python")


def say(msg):
    print(msg)
    sys.stdout.flush()


def call(cmd, quiet=False):
    """Запустить команду, стримя вывод. Возвращает код возврата."""
    if not quiet:
        say("  $ " + " ".join(cmd))
    p = subprocess.Popen(cmd)
    return p.wait()


def capture(cmd):
    """Запустить команду и вернуть (код, объединённый вывод)."""
    try:
        p = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
        out, _ = p.communicate()
    except OSError as e:
        return 127, str(e)
    if isinstance(out, bytes):
        out = out.decode("utf-8", "replace")
    return p.returncode, out


def libreoffice_hints():
    if is_windows():
        return [
            "  winget install TheDocumentFoundation.LibreOffice",
            "  choco install libreoffice-fresh        # если есть Chocolatey",
            "  scoop install libreoffice              # если есть Scoop",
            "  или распакуйте портативную сборку в " + os.path.join(ROOT, "prefix") + "\\",
            "  или задайте  set SOFFICE=C:\\путь\\к\\soffice.exe",
        ]
    if sys.platform == "darwin":
        return [
            "  brew install --cask libreoffice",
            "  или задайте  export SOFFICE=/Applications/LibreOffice.app/Contents/MacOS/soffice",
        ]
    return [
        "  sudo apt install libreoffice        # Debian/Ubuntu",
        "  sudo dnf install libreoffice-core   # Fedora",
        "  или задайте  export SOFFICE=/путь/к/soffice",
    ]


def usage_examples(venv_dir):
    scripts = os.path.join(ROOT, "scripts")
    if is_windows():
        runner = "py -3"
        sep = "\\"
        return [
            '  %s %s%sscripts%svalidate.py файл.docx --soffice' % (runner, ROOT, sep, sep),
            '  %s %s%sscripts%srender_check.py файл.docx' % (runner, ROOT, sep, sep),
            '  %s %s%stests%ssmoke.py' % (runner, ROOT, sep, sep),
        ]
    return [
        "  %s/scripts/validate.py файл.docx --soffice" % ROOT,
        "  %s/scripts/render_check.py файл.docx" % ROOT,
        "  %s/tests/smoke.py" % ROOT,
    ]


def main():
    parser = argparse.ArgumentParser(description="Установка office-docs (venv + зависимости).")
    parser.add_argument("--with-markitdown", action="store_true",
                        help="дополнительно поставить markitdown (~300 МБ)")
    parser.add_argument("--venv", default=os.path.join(ROOT, ".venv"),
                        help="куда ставить venv (по умолчанию <папка скилла>/.venv)")
    args = parser.parse_args()

    if sys.version_info < MIN_VERSION:
        say("Нужен Python 3.10 или новее, а запущен %d.%d." % sys.version_info[:2])
        say("Windows: поставьте Python с python.org или `winget install Python.Python.3.12`,")
        say("после чего запускайте `py -3 install.py`.")
        say("Linux: установите python3 (и python3-venv), затем `./install.sh`.")
        return 1

    venv_dir = os.path.abspath(os.path.expanduser(args.venv))
    vpy = venv_python(venv_dir)

    say("==> Python: %s (%d.%d)" % (sys.executable, sys.version_info[0], sys.version_info[1]))

    if os.path.exists(vpy):
        say("==> venv уже есть: %s" % venv_dir)
    else:
        say("==> Создаю venv: %s" % venv_dir)
        rc = call([sys.executable, "-m", "venv", venv_dir], quiet=True)
        if rc != 0 or not os.path.exists(vpy):
            say("Не удалось создать venv.")
            if is_windows():
                say("Проверьте, что Python поставлен целиком (с модулем venv):")
                say("  py -3 -m venv .venv")
            else:
                say("На Debian/Ubuntu модуль venv ставится отдельно: sudo apt install python3-venv")
            return 1

    say("==> Ставлю зависимости")
    call([vpy, "-m", "pip", "install", "--upgrade", "pip", "-q", "--disable-pip-version-check"],
         quiet=True)
    rc = call([vpy, "-m", "pip", "install", "-q", "--disable-pip-version-check",
               "-r", os.path.join(ROOT, "requirements.txt")])
    if rc != 0:
        say("pip install завершился с ошибкой — зависимости не поставлены.")
        return 1
    if args.with_markitdown:
        say("==> Ставлю markitdown (тяжёлый пакет, это долго)")
        rc = call([vpy, "-m", "pip", "install", "-q", "--disable-pip-version-check",
                   MARKITDOWN_SPEC])
        if rc != 0:
            say("markitdown не поставился — остальное установлено, это не критично.")
        else:
            rc_md, out_md = capture([vpy, "-c", "import markitdown; print('OK')"])
            if not out_md.strip().startswith("OK"):
                say("    markitdown поставился, но не импортируется: %s" % out_md.strip()[:200])

    say("==> Проверяю импорты")
    rc, out = capture([vpy, "-c", CHECK_IMPORTS])
    out = out.strip()
    if rc != 0 or not out.startswith("OK"):
        say("    НЕ ИМПОРТИРУЮТСЯ: %s" % out)
        return 1
    say("    все библиотеки на месте: %s" % ", ".join(CORE))

    say("==> Проверяю компиляцию скриптов")
    scripts_dir = os.path.join(ROOT, "scripts")
    rc, out = capture([vpy, "-m", "compileall", "-q", scripts_dir])
    if rc != 0:
        say("    ошибка компиляции:\n%s" % out)
        return 1
    shutil.rmtree(os.path.join(scripts_dir, "__pycache__"), ignore_errors=True)

    say("==> Ищу LibreOffice (soffice)")
    sys.path.insert(0, scripts_dir)
    try:
        from _common import find_soffice, find_pandoc  # noqa: E402

        soffice = find_soffice()
        pandoc = find_pandoc()
    except Exception as e:  # noqa: BLE001 — _common должен импортироваться всегда
        say("    _common не импортируется: %s" % e)
        soffice, pandoc = None, None

    if soffice:
        say("    найден: %s" % soffice)
    else:
        say("    НЕ НАЙДЕН — конвертация форматов и проверка рендера работать не будут.")
        for line in libreoffice_hints():
            say(line)
    if pandoc:
        say("    pandoc: %s" % pandoc)
    else:
        say("    pandoc не найден — без него не будет конвертации в .md (всё остальное работает).")
    say("    poppler/pdftoppm не нужен: страницы рисует PyMuPDF из зависимостей.")

    say("")
    say("Готово. Скрипты сами находят venv, поэтому запускайте их как есть:")
    for line in usage_examples(venv_dir):
        say(line)
    return 0


if __name__ == "__main__":
    sys.exit(main())
