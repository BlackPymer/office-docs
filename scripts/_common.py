"""Общие помощники для office-скриптов скилла office-docs.

Находит LibreOffice и запускает его безопасно: отдельный временный профиль
(иначе параллельные запуски зависают на блокировке), headless-бэкенд svp,
явный таймаут.
"""

from __future__ import annotations

import glob
import os
import shutil
import subprocess
import tempfile

# Корень скилла (родитель scripts/) — ищем зависимости рядом с кодом.
SKILL_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def soffice_candidates() -> list[str]:
    """Все места, где может лежать soffice, в порядке приоритета.

    1. $SOFFICE — явное указание пользователя.
    2. prefix/ внутри этого репозитория (так ставит install.sh).
    3. Старое пользовательское место ~/.local/share/office-skills/prefix
       (обратная совместимость).
    4. PATH.
    """
    cands: list[str] = []
    env = os.environ.get("SOFFICE")
    if env:
        cands.append(env)
    for prefix in (
        os.path.join(SKILL_ROOT, "prefix"),
        os.path.expanduser("~/.local/share/office-skills/prefix"),
    ):
        cands += sorted(glob.glob(os.path.join(prefix, "opt", "libreoffice*", "program", "soffice")))
    return cands


def find_soffice() -> str | None:
    """Порядок: $SOFFICE -> локальная установка -> PATH."""
    cands = soffice_candidates()
    for name in ("soffice", "libreoffice"):
        p = shutil.which(name)
        if p:
            cands.append(p)
    for c in cands:
        if c and os.path.exists(c):
            return c
    return None


def soffice_env() -> dict:
    env = os.environ.copy()
    env["SAL_USE_VCLPLUGIN"] = "svp"  # headless: не требует X11
    return env


def run_soffice(args: list[str], timeout: int = 180) -> subprocess.CompletedProcess:
    """Запуск soffice с временным профилем. Бросает RuntimeError, если soffice нет."""
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError(
            "LibreOffice (soffice) не найден. Установите его install.sh-ом "
            "(кладёт в <корень скилла>/prefix/opt/) или системным пакетом "
            "(apt install libreoffice), либо задайте SOFFICE=/путь/к/soffice"
        )
    with tempfile.TemporaryDirectory(prefix="lo_profile_") as prof:
        full = [
            f"-env:UserInstallation=file://{prof}",
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
        )


def convert_with_soffice(
    src: str, outdir: str, fmt: str, timeout: int = 180
) -> subprocess.CompletedProcess:
    """--convert-to <fmt> из src в каталог outdir."""
    os.makedirs(outdir, exist_ok=True)
    return run_soffice(
        ["--convert-to", fmt, "--outdir", outdir, os.path.abspath(src)],
        timeout=timeout,
    )
