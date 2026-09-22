#!/usr/bin/env python3
"""off2pdf · 一次性把 Office 文档转成 PDF（高保真版预览用；用完即走，不驻留）

用途：
  插件自带的纯 JS 引擎是"好看、秒开、零依赖"，但在复杂排版/特效上必不如原版。
  需要"跟原版一模一样"时，用本脚本把它转成 PDF —— 矢量、字体、分页 100% 保真，
  再用插件内置的 pdf.js 渲染（和 Codex 给 PDF 用的同一套渲染器）。

依赖：LibreOffice 的 `soffice`（唯一能忠实解析 97-2003/复杂 OOXML 的引擎之一）。
  它不在了脚本会明确报错；这不影响日常预览。

用法：
  python3 ~/.hermes/scripts/off2pdf.py ~/path/report.docx        # 生成 report.pdf（同目录）
  python3 ~/.hermes/scripts/off2pdf.py --out ~/tmp a.docx b.xlsx
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

SOFFICE_CANDIDATES = (
    "/opt/homebrew/bin/soffice",
    "/usr/local/bin/soffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
)

OFFICE_EXTS = {
    ".doc", ".docx", ".docm", ".dot", ".dotx", ".odt", ".rtf", ".txt",
    ".xls", ".xlsx", ".xlsm", ".ods", ".csv",
    ".ppt", ".pptx", ".pps", ".ppsx", ".odp", ".dps", ".et", ".wps",
    ".ofd", ".pages", ".numbers", ".key",
}


def find_soffice() -> str | None:
    for candidate in SOFFICE_CANDIDATES:
        if os.path.isfile(candidate) and os.access(candidate, os.X_OK):
            return candidate
    return shutil.which("soffice")


def convert_one(soffice: str, source: Path, out_dir: Path) -> tuple[bool, str]:
    out_dir.mkdir(parents=True, exist_ok=True)
    target = out_dir / (source.stem + ".pdf")

    if target.is_file() and target.stat().st_mtime >= source.stat().st_mtime:
        return True, f"跳过（已是最新）：{target}"

    workdir = Path(tempfile.mkdtemp(prefix="off2pdf-"))
    profile = workdir / "loprofile"
    try:
        started = time.time()
        proc = subprocess.run(
            [
                soffice,
                f"-env:UserInstallation=file://{profile}",
                "--headless",
                "--norestore",
                "--convert-to",
                "pdf",
                "--outdir",
                str(workdir),
                str(source),
            ],
            capture_output=True,
            text=True,
            timeout=300,
        )
        produced = sorted(workdir.glob("*.pdf"))
        if not produced:
            detail = (proc.stderr or proc.stdout or "").strip()[-300:]
            return False, f"转换失败：{source.name}（{detail or 'soffice 没有产出 pdf'}）"

        shutil.move(str(produced[0]), str(target))
        elapsed = time.time() - started
        return True, f"✓ {source.name} → {target.name}（{elapsed:.1f}s，{target.stat().st_size / 1024 / 1024:.1f} MB）"
    except subprocess.TimeoutExpired:
        return False, f"转换超时：{source.name}"
    finally:
        shutil.rmtree(workdir, ignore_errors=True)


def main() -> int:
    parser = argparse.ArgumentParser(description="把 Office 文档转成 PDF（一次性脚本，高保真预览用）")
    parser.add_argument("files", nargs="+", help="要转换的文档")
    parser.add_argument("--out", help="输出目录（默认写到源文件所在目录）")
    args = parser.parse_args()

    soffice = find_soffice()
    if not soffice:
        print("✗ 没找到 LibreOffice（soffice）。", file=sys.stderr)
        print("  高保真转换需要它；日常预览不受影响（插件内置引擎照常渲染）。", file=sys.stderr)
        return 2

    print(f"[off2pdf] soffice = {soffice}")
    failures = 0
    for raw in args.files:
        source = Path(raw).expanduser().resolve()
        if not source.is_file():
            print(f"✗ 文件不存在：{source}", file=sys.stderr)
            failures += 1
            continue
        if source.suffix.lower() not in OFFICE_EXTS:
            print(f"✗ 不支持的扩展名（{source.suffix or '无'}）：{source.name}", file=sys.stderr)
            failures += 1
            continue

        out_dir = Path(args.out).expanduser().resolve() if args.out else source.parent
        ok, message = convert_one(soffice, source, out_dir)
        print(("" if ok else "✗ ") + message, file=None if ok else sys.stderr)
        failures += 0 if ok else 1

    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
