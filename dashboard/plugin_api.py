"""office-viewer 插件后端（最小版）。

历史：这里曾经冻结过一套 office_preview 转换器（LibreOffice / 原生解析）把文档转成 HTML。
2026-09-22 起预览改由**纯 JS 内核**在渲染进程内完成（desktop/assets/office-viewer/），
不再需要任何转换器或外部程序，因此转换相关代码全部移除。

保留的三个端点只作为桌面桥的兜底（正常情况下前端走 window.hermesDesktop）：
  GET /list?path=     列目录
  GET /dataurl?path=  文件内容 data URL
  GET /open?path=     用系统默认程序打开
"""
from __future__ import annotations

import base64
import mimetypes
import os
import subprocess
import sys
from urllib.parse import unquote

from fastapi import APIRouter, HTTPException, Query

router = APIRouter()


def _expand(path: str) -> str:
    path = unquote(path or '')
    if path in ('~', '') or path.startswith('~/'):
        path = os.path.join(os.path.expanduser('~'), path[2:] if path.startswith('~/') else '')
    return os.path.abspath(path)


@router.get('/list')
def list_dir(path: str = Query('~/.hermes/workspace')):
    target = _expand(path)
    if not os.path.isdir(target):
        raise HTTPException(status_code=404, detail='not a directory')

    entries = []
    with os.scandir(target) as scan:
        for item in scan:
            try:
                entries.append({
                    'name': item.name,
                    'path': item.path,
                    'dir': item.is_dir(),
                    'size': None if item.is_dir() else item.stat().st_size,
                    'mtime': item.stat().st_mtime,
                })
            except OSError:
                continue

    parent = os.path.dirname(target.rstrip('/')) or None

    return {
        'path': target,
        'parent': parent,
        'home': os.path.expanduser('~'),
        'roots': [{'name': '主目录', 'path': os.path.expanduser('~')}, {'name': '根目录', 'path': '/'}],
        'entries': entries,
    }


@router.get('/dataurl')
def data_url(path: str = Query(...)):
    target = _expand(path)
    if not os.path.isfile(target):
        raise HTTPException(status_code=404, detail='not a file')

    with open(target, 'rb') as handle:
        payload = handle.read()
    mime = mimetypes.guess_type(target)[0] or 'application/octet-stream'

    return {'dataurl': f'data:{mime};base64,{base64.b64encode(payload).decode()}', 'bytes': len(payload)}


@router.get('/open')
def open_external(path: str = Query(...)):
    target = _expand(path)
    if not os.path.exists(target):
        raise HTTPException(status_code=404, detail='not found')
    if sys.platform == 'darwin':
        subprocess.Popen(['open', target])
    elif sys.platform.startswith('win'):
        os.startfile(target)  # noqa: S606
    else:
        subprocess.Popen(['xdg-open', target])

    return {'ok': True}
