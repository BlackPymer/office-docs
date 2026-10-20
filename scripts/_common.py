"""Общие помощники для office-скриптов скилла office-docs.

Работает одинаково на Linux, macOS и Windows:

* находит LibreOffice (``soffice``/``soffice.exe``) — в локальной установке
  рядом со скиллом, в пользовательских и системных каталогах, в PATH;
* запускает его headless с отдельным временным профилем (иначе параллельные
  запуски зависают на блокировке), с явным таймаутом и без всплывающих
  консолей на Windows;
* находит Python из venv скилла и, если текущему интерпретатору не хватает
  библиотек, перезапускает скрипт им. На Windows это единственный надёжный
  способ: shebang там не работает, а ``py -3 script.py`` берёт случайный Python.

Здесь только стандартная библиотека — модуль обязан импортироваться любым
Python 3.10+, даже до установки зависимостей.
"""

from __future__ import annotations

import glob
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Iterable, Sequence
from urllib.parse import quote

# Корень скилла (родитель scripts/) — ищем зависимости рядом с кодом.
SKILL_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Библиотеки, которые ставит install.py — ими пользуется большинство скриптов.
CORE_MODULES = ("docx", "pptx", "openpyxl", "pymupdf", "pypdf", "PIL", "lxml")

# Переменные окружения: защита от зацикливания ре-экзека и ручное указание Python.
_BOOTSTRAP_ENV = "OFFICE_DOCS_BOOTSTRAPPED"
PYTHON_ENV = "OFFICE_DOCS_PYTHON"
# Только для тестов: форсировать ветку другой ОС (см. tests/smoke.py).
_FORCE_PLATFORM_ENV = "OFFICE_DOCS_FORCE_PLATFORM"


# --------------------------------------------------------------------------
# Платформа
# --------------------------------------------------------------------------
def is_windows() -> bool:
    """True на Windows.

    Значение можно переопределить переменной ``OFFICE_DOCS_FORCE_PLATFORM``
    (``windows``/``posix``) — этим пользуется smoke-тест, чтобы проверять
    обе ветки путей на одной машине.
    """
    forced = os.environ.get(_FORCE_PLATFORM_ENV, "").strip().lower()
    if forced in ("windows", "nt", "win32"):
        return True
    if forced in ("posix", "linux", "darwin", "macos"):
        return False
    return os.name == "nt"


def no_window_kwargs() -> dict:
    """Не давать дочерним процессам открывать своё окно консоли на Windows."""
    if is_windows():
        return {"creationflags": getattr(subprocess, "CREATE_NO_WINDOW", 0)}
    return {}


def user_installation_url(profile: str | os.PathLike[str]) -> str:
    """``file://``-URL для ``-env:UserInstallation``.

    На Windows путь выглядит как ``C:\\Users\\...``, поэтому обычный
    ``file://C:\\...`` LibreOffice не понимает — нужен ``file:///C:/...``.
    """
    raw = str(profile)
    if not is_windows():
        p = Path(raw)
        try:
            return p.resolve().as_uri()
        except ValueError:  # путь другой платформы (тесты)
            pass
    posix = raw.replace("\\", "/")
    if len(posix) > 1 and posix[1] == ":":  # C:/... -> /C:/...
        posix = "/" + posix
    if not posix.startswith("/"):
        posix = "/" + posix
    return "file://" + quote(posix, safe="/:")


# --------------------------------------------------------------------------
# LibreOffice
# --------------------------------------------------------------------------
def _legacy_dirs(name: str) -> list[str]:
    """Каталоги установок прошлых версий скилла (venv/prefix), Windows и POSIX."""
    dirs: list[str] = []
    if is_windows():
        local = os.environ.get("LOCALAPPDATA")
        if local:
            dirs.append(os.path.join(local, "office-skills", name))
    dirs.append(os.path.expanduser(os.path.join("~", ".local", "share", "office-skills", name)))
    return dirs


def _system_soffice_dirs() -> list[str]:
    """Каталоги с бинарём soffice у системных установок LibreOffice."""
    dirs: list[str] = []
    if is_windows():
        for var in ("ProgramFiles", "ProgramW6432", "ProgramFiles(x86)", "LOCALAPPDATA"):
            base = os.environ.get(var)
            if not base:
                continue
            dirs.append(os.path.join(base, "LibreOffice", "program"))  # winget/choco/обычный MSI
            dirs.append(os.path.join(base, "Programs", "LibreOffice", "program"))  # per-user установка
        home = os.path.expanduser("~")
        dirs.append(os.path.join(home, "scoop", "apps", "libreoffice", "current", "program"))
    else:
        dirs.append("/Applications/LibreOffice.app/Contents/MacOS")  # macOS
    return dirs


def soffice_candidates() -> list[str]:
    """Все места, где может лежать soffice, в порядке приоритета.

    1. ``$SOFFICE`` — явное указание пользователя.
    2. ``prefix/`` внутри этой папки скилла (так ставит портативную сборку
       install.py/README);
    3. пользовательские места прошлых установок (``~/.local/share``,
       ``%LOCALAPPDATA%``);
    4. системные установки: Program Files, scoop, ``/Applications/LibreOffice.app``.

    PATH проверяется отдельно — в :func:`find_soffice`.
    """
    exe = "soffice.exe" if is_windows() else "soffice"
    cands: list[str] = []

    env = os.environ.get("SOFFICE")
    if env:
        cands.append(env)

    prefixes = [os.path.join(SKILL_ROOT, "prefix")] + _legacy_dirs("prefix")
    patterns = (
        os.path.join("opt", "libreoffice*", "program", exe),  # распакованный архив (Linux/macOS)
        os.path.join("LibreOffice*", "program", exe),         # распакованный архив (Windows)
        os.path.join("libreoffice*", "program", exe),
        os.path.join("program", exe),                         # скопировали содержимое установки
    )
    for prefix in prefixes:
        for pattern in patterns:
            cands += sorted(glob.glob(os.path.join(prefix, pattern)))

    for prog_dir in _system_soffice_dirs():
        cands.append(os.path.join(prog_dir, exe))
    return cands


def find_soffice() -> str | None:
    """Порядок: ``$SOFFICE`` -> локальная установка -> системные каталоги -> PATH."""
    cands = soffice_candidates()
    for name in ("soffice", "soffice.exe", "soffice.com", "libreoffice"):
        p = shutil.which(name)
        if p:
            cands.append(p)
    for c in cands:
        if c and os.path.exists(c):
            return c
    return None


def find_pandoc() -> str | None:
    """pandoc нужен только для конвертации в Markdown (docx/odt/rtf/html -> md)."""
    p = shutil.which("pandoc") or shutil.which("pandoc.exe")
    if p:
        return p
    cands = [
        os.path.expanduser(os.path.join("~", ".local", "bin", "pandoc")),
        os.path.join(SKILL_ROOT, "bin", "pandoc"),  # portable-раскладка рядом со скиллом
    ]
    if is_windows():
        for var in ("LOCALAPPDATA", "ProgramFiles", "ProgramFiles(x86)", "ProgramW6432"):
            base = os.environ.get(var)
            if base:
                cands.append(os.path.join(base, "Pandoc", "pandoc.exe"))
    for c in cands:
        if os.path.exists(c):
            return c
    return None


def find_pdftoppm() -> str | None:
    """pdftoppm из poppler — теперь необязателен: страницы рисует PyMuPDF."""
    return shutil.which("pdftoppm") or shutil.which("pdftoppm.exe")


def soffice_env() -> dict:
    env = os.environ.copy()
    if not is_windows():
        env["SAL_USE_VCLPLUGIN"] = "svp"  # headless без X11/Wayland
    return env


def run_soffice(
    args: Sequence[str], timeout: int = 180, profile: str | None = None
) -> subprocess.CompletedProcess:
    """Запуск soffice с временным профилем. Бросает RuntimeError, если soffice нет."""
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError(
            "LibreOffice (soffice) не найден. Поставьте его системным пакетом "
            "(apt install libreoffice / winget install TheDocumentFoundation.LibreOffice), "
            "положите портативную сборку в <папку скилла>/prefix/ "
            "или задайте SOFFICE=/путь/к/soffice"
        )
    with tempfile.TemporaryDirectory(prefix="lo_profile_") as prof:
        url = user_installation_url(profile or prof)
        full = [
            f"-env:UserInstallation={url}",
            "--headless",
            "--norestore",
            "--nolockcheck",
            "--nofirststartwizard",
        ] + list(args)
        return subprocess.run(
            [soffice] + full,
            env=soffice_env(),
            timeout=timeout,
            capture_output=True,
            text=True,
            **no_window_kwargs(),
        )


def convert_with_soffice(
    src: str, outdir: str, fmt: str, timeout: int = 180
) -> subprocess.CompletedProcess:
    """``--convert-to <fmt>`` из src в каталог outdir."""
    os.makedirs(outdir, exist_ok=True)
    return run_soffice(
        ["--convert-to", fmt, "--outdir", outdir, os.path.abspath(src)],
        timeout=timeout,
    )


# --------------------------------------------------------------------------
# Python: venv скилла и ре-экзек
# --------------------------------------------------------------------------
def venv_python_candidates() -> list[str]:
    """Пути к интерпретатору внутри venv скилла, в порядке приоритета.

    На Windows venv раскладывается в ``Scripts/python.exe``, на POSIX — в
    ``bin/python``; поддерживаются оба варианта, потому что venv может быть
    создан и под Windows, и под WSL/Git Bash.
    """
    names = (os.path.join("Scripts", "python.exe"), os.path.join("bin", "python"))
    cands: list[str] = []
    explicit = os.environ.get(PYTHON_ENV)
    if explicit:
        cands.append(explicit)
    # <скилл>/.venv — так ставит install.py; <скилл>/scripts/.venv — для ручных экспериментов
    for root in (SKILL_ROOT, os.path.join(SKILL_ROOT, "scripts")):
        cands += [os.path.join(root, ".venv", n) for n in names]
    # прошлые установки: там venv лежит целиком (…/office-skills/venv/bin/python)
    for legacy in _legacy_dirs("venv"):
        cands += [os.path.join(legacy, n) for n in names]
    seen: set[str] = set()
    unique: list[str] = []
    for c in cands:
        if c not in seen:
            seen.add(c)
            unique.append(c)
    return unique


def missing_modules(needs: Iterable[str]) -> list[str]:
    """Какие из модулей не импортируются текущим интерпретатором."""
    import importlib

    bad: list[str] = []
    for name in needs:
        try:
            importlib.import_module(name)
        except Exception:  # noqa: BLE001 — любая ошибка импорта значит «нет библиотеки»
            bad.append(name)
    return bad


def _can_import(python: str, needs: Sequence[str]) -> bool:
    code = "".join(f"import {m};" for m in needs)
    try:
        r = subprocess.run(
            [python, "-c", code],
            capture_output=True,
            timeout=180,
            **no_window_kwargs(),
        )
    except (OSError, subprocess.SubprocessError):
        return False
    return r.returncode == 0


def _usable_python(needs: Sequence[str]) -> str | None:
    """Первый НЕ текущий интерпретатор, который умеет импортировать needs."""
    current = os.path.abspath(sys.executable)
    for cand in venv_python_candidates():
        if not cand or os.path.abspath(cand) == current or not os.path.exists(cand):
            continue
        if _can_import(cand, needs):
            return cand
    return None


def find_runtime_python(needs: Sequence[str] = CORE_MODULES) -> str | None:
    """Интерпретатор с нужными библиотеками: текущий, иначе из venv скилла.

    Используется ``scripts/py.py``: запустить свой код так, чтобы docx/pptx/
    openpyxl/pymupdf точно были доступны, на любой платформе.
    """
    if not missing_modules(needs):
        return sys.executable
    return _usable_python(needs)


def _reexec(python: str) -> None:
    """Перезапустить текущий скрипт найденным интерпретатором. Не возвращается."""
    script = os.path.abspath(sys.argv[0])
    env = os.environ.copy()
    env[_BOOTSTRAP_ENV] = "1"
    argv = [python, script] + sys.argv[1:]
    if not is_windows():
        os.execve(python, argv, env)
    # На Windows os.execv оставляет лишний процесс и кривую консоль — запускаем
    # дочерний процесс и возвращаем его код.
    sys.exit(subprocess.run(argv, env=env, **no_window_kwargs()).returncode)


def ensure_runtime(needs: Sequence[str] = CORE_MODULES) -> None:
    """Гарантировать, что скрипт выполняется интерпретатором с нужными библиотеками.

    На Linux/macOS это обычно уже сделал shebang-трampoline, и вызов ничего не
    делает. На Windows shebang не работает, поэтому скрипт мог быть запущен
    любым Python — тогда находим venv скилла и перезапускаем скрипт им.
    """
    if not needs or os.environ.get(_BOOTSTRAP_ENV):
        return
    missing = missing_modules(needs)
    if not missing:
        return
    python = _usable_python(needs)
    if python:
        _reexec(python)
    hint = (
        "  py -3 install.py      # Windows"
        if is_windows()
        else "  ./install.sh          # Linux/macOS"
    )
    raise SystemExit(
        "Не хватает библиотек: "
        + ", ".join(missing)
        + "\nПоставьте их установщиком скилла:\n"
        + hint
        + f"\nЛибо укажите готовый Python: set {PYTHON_ENV}=C:\\path\\python.exe (Windows)"
        f" / export {PYTHON_ENV}=/path/python (Linux/macOS)"
    )
