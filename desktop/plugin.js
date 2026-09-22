/**
 * office-viewer · 桌面半（Hermes Desktop 插件）
 *
 * 目标：把「doc/docx/xls/xlsx/ppt/pptx/wps/et/dps/ofd/pages/numbers/key/epub/
 * eml/msg/zip/psd/ai/dxf/sqlite/ipynb …」在本机渲染出来，直接在这台桌面端里看。
 *
 * 三块界面
 *   1. `::office{file="…"}` —— 消息内产物卡片。**点一下就呼出侧栏查看器**（用户 2026-09-21 口径）。
 *   2. 侧栏面板（placement: right）—— 常驻的「文档预览」：默认是**文件浏览器**（类访达，可访问全盘），
 *      点文件就地铺满预览（右上角：用系统程序打开 / 关闭）。
 *   3. 整页路由 `/office-viewer` + 侧栏导航 + ⌘K 命令 + 快捷键。
 *
 * 渲染从哪来（2026-09-22 起）：**纯 JS 渲染内核**，随插件携带在
 *   desktop/assets/office-viewer/office-viewer.js（约 2.0 MB，含 docx/doc/xlsx/xls/pptx/csv/rtf/ofd 引擎）
 *   desktop/assets/office-viewer/pptx.worker.js（pptx 的 Worker）
 *   desktop/assets/office-viewer/marked.umd.js（Markdown 渲染）
 * 内核在渲染进程内的 iframe 里解析渲染；文件字节由桌面桥（window.hermesDesktop）读入后直接递给引擎。
 * 全程**不起服务、不写临时文件、不经过任何外部转换器、不依赖 LibreOffice**。
 * 其余格式：md → marked；html / 文本 / 脚本 → 内嵌只读；图片 / PDF → 内嵌显示。
 *
 * 约定（SDK 硬约束）：只允许 import `@hermes/plugin-sdk`、`react`、`react/jsx-runtime`。
 */

import {
  atom,
  host,
  useValue,
  cn,
  Button,
  Codicon,
  Tip,
  ScrollArea,
  PANES_AREA,
  ROUTES_AREA,
  SIDEBAR_NAV_AREA,
  PALETTE_AREA,
  KEYBINDS_AREA,
  TRANSCRIPT_DIRECTIVE_AREA
} from '@hermes/plugin-sdk'
import { jsx, jsxs } from 'react/jsx-runtime'
import { useCallback, useEffect, useMemo, useRef, useState } from 'react'

const PLUGIN_ID = 'office-viewer'
const PANE_KEY = 'viewer'                       // 贡献 id
const PANE_ID = `${PLUGIN_ID}:${PANE_KEY}`      // 宿主里的全名：<pluginId>:<paneId>
const ROUTE = '/office-viewer'

/** 当前要看的文件（三个界面共用一份状态）。 */
const $file = atom(null)
/** 变更计数：换文件/点重新渲染都 +1，Viewer 据此重取。 */
const $nonce = atom(0)

/** ctx 的能力，供组件使用（贡献只拿到 ctx，组件拿不到）。 */
let doors = null

/** 探针：把现场情况写进插件后端日志（只用于排障，失败静默）。 */
function beacon(event, detail) {
  try {
    void doors?.rest(`/beacon?event=${encodeURIComponent(event)}&detail=${encodeURIComponent(String(detail || ''))}`)
      ?.catch?.(() => {})
  } catch {
    /* 探针不影响功能 */
  }
}

/**
 * 打开右侧栏的查看器并切到指定文件。
 *
 * 为什么不是"点一下就弹窗"：用户口径（2026-09-21）——产物卡片点击后要在**右侧边栏**看。
 * 实现顺序：
 *   1. revealPane 前台化本插件面板（含取消收起/取消最小化）；
 *   2. 一个 tick 后检查面板是否真的可见（host.paneVisibility）；
 *   3. 不可见（右列被用户收起等）→ 退到主区标签（host.openWorkspace），并只提示一次。
 */
function openInPane(path) {
  if (path) {
    $file.set(path)
  }
  beacon('openInPane', path)

  const canReveal = typeof host.revealPane === 'function'
  beacon('revealSupported', String(canReveal))

  if (canReveal) {
    try {
      host.revealPane(PANE_ID)
    } catch (error) {
      host.notifyError?.(error, '无法呼出 Office 查看器面板')
    }
  } else if (typeof host.openWorkspace === 'function') {
    // 老版本桌面端没有 revealPane：直接开到主区，别让点击看起来"没反应"。
    host.openWorkspace('office-viewer:viewer', {
      title: 'Office 查看器',
      render: () => jsx(Viewer, { path: $file.get(), compact: false })
    })
  }

  setTimeout(() => {
    let visible = null
    try {
      if (typeof host.paneVisibility === 'function') {
        visible = !!host.paneVisibility(PANE_ID).get()
      }
    } catch {
      visible = null
    }
    beacon('paneVisibility', visible === null ? 'unsupported' : String(visible))

    if (visible === false && typeof host.openWorkspace === 'function') {
      host.openWorkspace('office-viewer:viewer', {
        title: 'Office 查看器',
        render: () => jsx(Viewer, { path: $file.get(), compact: false })
      })
      if (!openInPane.warned) {
        openInPane.warned = true
        host.notify?.({
          kind: 'info',
          message: '右侧栏当前是收起的，已在主区打开 Office 查看器（⌘J 可开右列）'
        })
      }
    }
  }, 120)
}

/** 文件名（路径里取尾段）。 */
function baseName(path) {
  return String(path || '').split('/').pop() || String(path || '')
}

/**
 * 「请本插件预览某个文件」的约定事件。
 *   谁在用：桌面美化插件（desktop-beautify）—— 用户在右侧栏「文件」面板里点 doc/xls/ppt/md/txt 等，
 *   它就把这次点击从应用手里截下来，改发这个事件，于是文件在**这个**标签里预览，
 *   而不是让应用另开一个预览标签（用户口径 2026-09-22）。
 *   为什么用 window 事件：磁盘插件之间不能互相 import，只能靠这样一个约定。
 *   约定：`window.dispatchEvent(new CustomEvent('hermes-office-open', { detail: { path } }))`
 *   —— `path` 必须是绝对路径（面板就按绝对路径读文件）。
 */
const EXTERNAL_OPEN_EVENT = 'hermes-office-open'

function onExternalOpenRequest(event) {
  const path = String(event?.detail?.path ?? '')

  if (path.startsWith('/')) {
    openInPane(path)
  }
}

/** 人类可读体积。 */
function fmtSize(bytes) {
  if (typeof bytes !== 'number' || bytes < 0) return ''
  const units = ['B', 'KB', 'MB', 'GB']
  let value = bytes
  let i = 0
  while (value >= 1024 && i < units.length - 1) {
    value /= 1024
    i += 1
  }
  return `${value >= 10 || i === 0 ? Math.round(value) : value.toFixed(1)} ${units[i]}`
}

/** 小工具条按钮（主题化，无硬编码颜色）。 */
function BarButton({ icon, label, onClick, disabled, tip }) {
  const hasLabel = Boolean(label)
  const button = jsxs(Button, {
    size: hasLabel ? 'xs' : 'icon-xs',
    variant: 'ghost',
    disabled: !!disabled,
    onClick,
    className: 'gap-1.5',
    children: [
      icon ? jsx(Codicon, { name: icon, spinning: icon === 'loading' }) : null,
      hasLabel ? jsx('span', { children: label }) : null
    ]
  })
  return tip ? jsx(Tip, { label: tip, children: button }) : button
}


/* ══════════════════ 文件浏览器（访达式）＋ 铺满面板的预览 ══════════════════
 * 用户口径（2026-09-22）：右栏「文档预览」要**取代「文件」标签**的全部能力 ——
 *   · 打开就是文件浏览器（默认落工作区），能一路点进任何目录，**全局磁盘**都能去；
 *   · 上一级 / 后退 / 前进 / 可点的路径栏（像 macOS 访达）；
 *   · 点文件 ⇒ 就地铺满整个面板预览，右上角两个按钮：用系统默认应用打开、关闭；
 *   · 图标一律素色矢量（codicon），不同格式在文件名右侧挂个**带色小标签**。
 *  数据来自本插件后端半的 `/list`（列目录，任意绝对路径）、`/html`（渲染）、`/dataurl`、`/open`。
 */

const WORKSPACE_HOME = '~/.hermes/workspace'

/** 桌面端预加载桥（`window.hermesDesktop`）：列目录 / 读成 data URL。
 *  **必须优先用它** —— 插件的 `/list` 走 `ctx.rest` ⇒ 打到 dashboard 的 plugin API，
 *  而本机后端是 `hermes serve`（headless，web UI 关闭）⇒ 那条路 404（实测：
 *  "Headless backend (hermes serve): web UI disabled"）。桥是桌面端自己的 IPC，永远在。 */
function desktopBridge() {
  try {
    return typeof window !== 'undefined' ? window.hermesDesktop ?? null : null
  } catch {
    return null
  }
}

function bridgeListing(raw, requested) {
  const entries = (raw?.entries ?? []).map(entry => {
    const name = entry.name
    const path = entry.path

    return {
      name,
      path,
      dir: !!entry.isDirectory,
      ext: entry.isDirectory ? '' : extOf(path),
      size: null,
      mtime: 0,
      hidden: name.startsWith('.'),
      supported: false
    }
  })

  const dir = String(requested || '')

  return {
    path: dir,
    parent: dir.replace(/\/+$/, '').split('/').slice(0, -1).join('/') || null,
    home: '~',
    roots: [{ name: '主目录', path: '~' }, { name: '根目录', path: '/' }],
    entries
  }
}

async function listDirectory(path) {
  const bridge = desktopBridge()

  if (typeof bridge?.readDir === 'function') {
    const raw = await bridge.readDir(path)

    if (raw?.error) throw new Error(raw.error)

    return bridgeListing(raw, path)
  }

  if (doors?.rest) {
    return await doors.rest(`/list?path=${encodeURIComponent(path)}`)
  }

  throw new Error('桌面桥与插件后端都不可用')
}

/** 图片/PDF 这类不需要转换的格式：直接读成 data URL 在面板里显示。 */
const DATAURL_EXTS = ['png', 'jpg', 'jpeg', 'gif', 'webp', 'svg', 'bmp', 'tiff', 'avif', 'heic', 'ico']

async function readDataUrl(path) {
  const bridge = desktopBridge()

  if (typeof bridge?.readFileDataUrl === 'function') {
    return await bridge.readFileDataUrl(path)
  }

  if (doors?.rest) {
    const payload = await doors.rest(`/dataurl?path=${encodeURIComponent(path)}`)

    return payload?.dataurl ?? null
  }

  return null
}

/** 格式家族：颜色 + 素色图标（用户点名：Word 蓝 / Excel 绿 / PPT 橙 / PDF 红 / md 灰）。 */
const FORMAT_FAMILIES = [
  { color: '#2563eb', exts: ['doc', 'docx', 'rtf', 'pages', 'wps', 'odt'], icon: 'file-text' },
  { color: '#16a34a', exts: ['xls', 'xlsx', 'csv', 'numbers', 'et', 'ods'], icon: 'table' },
  { color: '#ea580c', exts: ['ppt', 'pptx', 'key', 'dps', 'odp'], icon: 'file-media' },
  { color: '#dc2626', exts: ['pdf'], icon: 'file-pdf' },
  { color: '#6b7280', exts: ['md', 'markdown', 'txt', 'log', 'text'], icon: 'markdown' },
  { color: '#0891b2', exts: ['png', 'jpg', 'jpeg', 'gif', 'webp', 'svg', 'bmp', 'tiff', 'heic', 'avif'], icon: 'file-media' },
  { color: '#7c3aed', exts: ['zip', 'rar', '7z', 'tar', 'gz', 'bz2', 'xz'], icon: 'file-zip' },
  { color: '#0f766e', exts: ['json', 'yml', 'yaml', 'toml', 'ini', 'conf', 'py', 'js', 'mjs', 'cjs', 'ts', 'tsx', 'jsx', 'sh', 'bash', 'zsh', 'html', 'htm', 'css', 'scss', 'xml', 'sql', 'go', 'rs', 'java', 'c', 'h', 'cpp'], icon: 'file-code' },
  { color: '#64748b', exts: ['epub', 'mobi'], icon: 'book' }
]

const IMAGE_EXTS = ['png', 'jpg', 'jpeg', 'gif', 'webp', 'svg', 'bmp', 'tiff', 'heic', 'avif']

/** 交付产物格式：一个格式独一份图标，页面形状 + 格式字母（跟内置 `file-pdf` 同一套视觉语言）。
 *  脚本类 / 图片类 / 压缩包不在这里 —— 它们按用户口径共用图标（见 FORMAT_FAMILIES）。 */
const ICON_FAMILIES = [
  { icon: 'file-text', exts: ['doc', 'docx', 'rtf', 'odt', 'wps', 'pages'] },
  { icon: 'table', exts: ['xls', 'xlsx', 'csv', 'ods', 'et', 'numbers'] },
  { icon: 'notebook', exts: ['ppt', 'pptx', 'odp', 'key', 'dps'] },
  { icon: 'symbol-text', exts: ['txt', 'log', 'text'] },
  { icon: 'code', exts: ['html', 'htm'] },
  { icon: 'markdown', exts: ['md', 'markdown'] },
  { icon: 'json', exts: ['json', 'yml', 'yaml', 'toml', 'ini', 'conf', 'xml'] },
  { icon: 'file-pdf', exts: ['pdf'] },
  { icon: 'file-zip', exts: ['zip', 'rar', '7z', 'tar', 'gz', 'bz2', 'xz'] },
  { icon: 'book', exts: ['epub', 'mobi'] },
  { icon: 'notebook-template', exts: ['ipynb'] },
  { icon: 'file-code', exts: ['py', 'js', 'mjs', 'cjs', 'ts', 'tsx', 'jsx', 'sh', 'bash', 'zsh', 'css', 'scss', 'sql', 'go', 'rs', 'java', 'c', 'h', 'cpp'] },
  { icon: 'file-media', exts: ['png', 'jpg', 'jpeg', 'gif', 'webp', 'svg', 'bmp', 'tiff', 'heic', 'avif', 'mp3', 'wav', 'm4a', 'flac', 'mp4', 'mov', 'mkv', 'avi', 'webm'] }
]

/** 格式 → codicon。用户口径（2026-09-22）：常用交付格式要一眼可辨、各用各的；
 *  脚本类共用 `file-code`、图片/音视频共用 `file-media`。
 *  局限（codicon 本身只有这几颗带格式指向的图标）：docx 与 txt 只能靠 `file-text` /
 *  `symbol-text` 区分，pptx 无专属图标，暂以 `notebook` 代。 */
function iconForPath(path) {
  const ext = extOf(path)

  return ICON_FAMILIES.find(item => item.exts.includes(ext))?.icon ?? 'file'
}

/** 文件图标（统一走 codicon，素色）。 */
function FileGlyph({ path, className }) {
  return jsx(Codicon, { name: iconForPath(path), className })
}

function extOf(path) {
  const name = String(path || '').split('/').pop() || ''
  const dot = name.lastIndexOf('.')

  return dot > 0 ? name.slice(dot + 1).toLowerCase() : ''
}

function familyOf(path) {
  const ext = extOf(path)

  return { color: '#94a3b8', ext, icon: 'file', family: FORMAT_FAMILIES.find(item => item.exts.includes(ext)) }
}

function humanSize(bytes) {
  if (bytes === null || bytes === undefined) return ''
  let value = Number(bytes)
  const units = ['B', 'KB', 'MB', 'GB', 'TB']
  let i = 0

  while (value >= 1024 && i < units.length - 1) { value /= 1024; i += 1 }

  return `${value >= 10 || i === 0 ? Math.round(value) : value.toFixed(1)} ${units[i]}`
}

function shortDate(stamp) {
  if (!stamp) return ''
  const d = new Date(stamp * 1000)
  const pad = n => String(n).padStart(2, '0')

  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`
}

/** 格式小标签：有色的框 + 里面的扩展名。 */

/** 用系统默认应用打开（后端 `/open` 走 macOS 的 `open`，最稳；再兜一层宿主 os 口）。 */
async function openWithSystemApp(path) {
  try {
    if (doors?.rest) {
      await doors.rest(`/open?path=${encodeURIComponent(path)}`)

      return true
    }
  } catch {
    /* 落到宿主那口 */
  }
  try {
    return await doors?.os?.openExternal?.(`file://${encodeURI(path)}`)
  } catch {
    return false
  }
}


/* ─────────────── 纯 JS 渲染内核：免服务 / 不依赖 LibreOffice / 不转 HTML ───────────────
 * 内核来自 cloudcli-office-viewer（@file-viewer/* 系列，浏览器内解析 OOXML / OLE2）。
 * 它在 iframe 里跑：内核与 pptx Worker 都以 blob URL 注入，文件字节由桌面桥读入后直接递过去，
 * 全程不起服务、不写临时文件、不经过任何外部转换器。
 * 与 DESIGN.md 的硬性原则一致：不传任何懒加载开关（lazySlides/lazyMedia 恒为 false）。 */

const ENGINE_DIR = '~/.hermes/plugins/office-viewer/desktop/assets/office-viewer'
/** 新内核接管的格式；md / html / pdf / 图片 / 纯文本仍走原生路径。 */
const ENGINE_EXTS = ['docx', 'doc', 'xlsx', 'xls', 'csv', 'pptx', 'ppt', 'wps', 'et', 'dps', 'ofd', 'rtf']

const assetTextCache = new Map()

/** 读内核资产（走桌面桥；插件后端仅作兜底）。 */
async function readEngineAsset(name) {
  const bridge = desktopBridge()

  if (typeof bridge?.readFileDataUrl === 'function') {
    return await bridge.readFileDataUrl(`${ENGINE_DIR}/${name}`)
  }

  if (doors?.rest) {
    const payload = await doors.rest(`/dataurl?path=${encodeURIComponent(`${ENGINE_DIR}/${name}`)}`)

    return payload?.dataurl ?? null
  }

  return null
}

/** 内核 / Worker 的文本只读一次，之后走内存缓存。 */
async function engineAssetText(name) {
  if (assetTextCache.has(name)) return assetTextCache.get(name)

  const promise = (async () => {
    const url = await readEngineAsset(name)

    if (!url) throw new Error(`读不到渲染内核资产：${name}`)

    const response = await fetch(url)

    return await response.text()
  })()

  assetTextCache.set(name, promise)

  return promise
}

/** 读内核资产为字节（wasm / 字体这类二进制，不能当文本读）。 */
const assetBytesCache = new Map()

async function engineAssetBytes(name) {
  if (assetBytesCache.has(name)) return assetBytesCache.get(name)

  const promise = (async () => {
    const url = await readEngineAsset(name)

    if (!url) throw new Error(`读不到渲染内核资产：${name}`)

    return new Uint8Array(await (await fetch(url)).arrayBuffer())
  })()

  assetBytesCache.set(name, promise)

  return promise
}

/** 把内核源码/Worker 源码嵌进 iframe 时，避免 "</script" 提前闭合脚本块。 */
function escapeScriptText(text) {
  return String(text).replace(/<\/script/gi, '<\\/script')
}

/** iframe 引导页：内核以 blob URL 载入，Worker 以 text/plain 内联后转 blob。 */
function engineBootstrap(engineUrl, workerText) {
  return `<!doctype html><html><head><meta charset="utf-8">
<style>html,body{margin:0;height:100%;background:#fff}
#ov-host{height:100%;min-height:100%;display:flex;flex-direction:column}
#ov-host > *{flex:1 1 auto;min-height:0}
.ov-status{padding:14px;font:12px/1.7 -apple-system,"PingFang SC","Helvetica Neue",sans-serif;color:#6b7280}
.ov-status.ov-error{color:#b91c1c;white-space:pre-wrap}</style>
</head><body>
<div id="ov-status" class="ov-status">正在载入渲染内核…</div>
<div id="ov-host"></div>
<script type="text/plain" id="ov-worker">${escapeScriptText(workerText)}</script>
<script src="${engineUrl}"></script>
<script>
try {
  var node = document.getElementById('ov-worker');
  var text = (node && node.textContent) || '';
  if (text.trim()) {
    var url = URL.createObjectURL(new Blob([text], { type: 'text/javascript' }));
    globalThis.__OV_RESOLVE_ASSET__ = function (rel) { return rel === 'pptx/pptx.worker.js' ? url : ''; };
  }
  globalThis.__OV_ERR__ = globalThis.CloudCLIOfficeViewer ? '' : '内核未挂载 CloudCLIOfficeViewer';
} catch (e) { globalThis.__OV_ERR__ = String((e && e.message) || e); }
</script>
</body></html>`
}

/** 起一个装了内核的 iframe，挂在 host 上，等它可用。 */
async function mountEngineFrame(host, name) {
  const [engineText, workerText] = await Promise.all([engineAssetText('office-viewer.js'), engineAssetText('pptx.worker.js')])
  const engineUrl = URL.createObjectURL(new Blob([engineText], { type: 'text/javascript' }))
  const iframe = document.createElement('iframe')

  iframe.setAttribute('title', name)
  iframe.style.cssText = 'width:100%;height:100%;border:0;display:block;background:#fff'
  iframe.srcdoc = engineBootstrap(engineUrl, workerText)
  host.appendChild(iframe)

  await new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error('渲染内核加载超时（30s）')), 30_000)

    iframe.addEventListener('load', () => {
      clearTimeout(timer)

      const win = iframe.contentWindow

      if (!win || !win.CloudCLIOfficeViewer || typeof win.CloudCLIOfficeViewer.render !== 'function') {
        reject(new Error(win?.__OV_ERR__ || '渲染内核未就绪'))

        return
      }

      resolve()
    })
  })

  return { iframe, engineUrl }
}

/* ─────────────── 「精确版（PDF）」：高保真查看入口 ───────────────
 * 插件不能执行命令（插件 SDK 无 exec 能力），所以这里做两件能做对的事：
 *   1) 若同目录已有同名 PDF（由 off2pdf.py 生成）⇒ 直接在面板里用 pdf.js 打开（矢量、原版保真）；
 *   2) 若还没有 ⇒ 弹出提示，把要跑的命令给你一键复制。
 * 不开任何常驻进程。 */

const OFFICE_EXTS_FOR_PDF = ['doc', 'docx', 'docm', 'rtf', 'odt', 'xls', 'xlsx', 'xlsm', 'ods', 'csv', 'ppt', 'pptx', 'pps', 'ppsx', 'dps', 'et', 'wps', 'ofd']

function ExactPdfButton({ path }) {
  const [open, setOpen] = useState(false)
  const [note, setNote] = useState('')
  const [checking, setChecking] = useState(false)

  const command = `python3 ~/.hermes/scripts/off2pdf.py "${path}"`
  const dir = String(path).slice(0, String(path).lastIndexOf('/'))
  const stem = baseName(path).replace(/\.[^.]+$/, '')

  const lookForPdf = useCallback(async () => {
    setChecking(true)
    setNote('')

    try {
      const bridge = desktopBridge()
      const listing = await bridge?.readDir?.(dir)
      const entries = listing?.entries || listing || []
      const hit = entries.find(entry => !entry.isDirectory && entry.name === `${stem}.pdf`)

      if (hit?.path) {
        setOpen(false)
        $file.set(hit.path)

        return
      }

      setOpen(true)
      setNote(`还没有 ${stem}.pdf —— 先在终端跑一次下面这条命令（约 2–10 秒）。`)
    } catch (failure) {
      setOpen(true)
      setNote(`列目录失败：${String(failure?.message || failure)}`)
    } finally {
      setChecking(false)
    }
  }, [dir, stem])

  const copyCommand = useCallback(() => {
    void doors?.os?.writeClipboard?.(command)
    setNote('命令已复制，粘到终端跑一次即可。')
  }, [command])

  return jsxs('span', {
    className: 'relative inline-flex',
    children: [
      jsx(BarButton, {
        icon: 'file-pdf',
        label: '',
        tip: '精确版（转 PDF 后用矢量渲染）',
        onClick: () => { void lookForPdf() }
      }),
      open && !checking
        ? jsxs('div', {
          className: cn(
            'absolute right-0 top-full z-50 mt-1 w-[26rem] rounded border p-3 text-left shadow-lg',
            'border-(--ui-stroke-secondary) bg-(--ui-bg-primary) text-(--ui-text-secondary)'
          ),
          children: [
            jsx('div', { className: 'mb-2 text-xs', children: '精确版＝把这个文档转成 PDF 再看（排版与字体与原版一致）' }),
            note ? jsx('div', { className: 'mb-2 text-[0.6875rem] text-(--ui-text-tertiary)', children: note }) : null,
            jsx('pre', {
              className: 'mb-2 select-all overflow-x-auto whitespace-pre rounded border border-(--ui-stroke-secondary) bg-(--ui-bg-secondary) p-2 text-[0.6875rem]',
              children: command
            }),
            jsxs('div', {
              className: 'flex items-center gap-2',
              children: [
                jsx(BarButton, { icon: 'copy', label: '复制命令', onClick: copyCommand }),
                jsx(BarButton, { icon: 'refresh', label: '我已生成，重新检查', onClick: () => { void lookForPdf() } }),
                jsx(BarButton, { icon: 'close', label: '收起', onClick: () => setOpen(false) })
              ]
            }),
            jsx('div', { className: 'mt-2 text-[0.625rem] text-(--ui-text-quaternary)', children: '转换需要 LibreOffice 的 soffice；纯脚本无法忠实解析复杂 Office 排版。' })
          ]
        })
        : null
    ]
  })
}

/* ─────────────── 旧版二进制演示（.ppt / .dps）专用引擎 ───────────────
 * `@file-viewer/ppt`（Flyfish 公开带水印运行时，非开源，manifest 声明 watermarkRequired）。
 * 水印由 WASM 画在最终帧上，无法用 CSS 去除 —— 用户已裁决接受。
 * 该引擎自成一体（无相对 import），因此 index/worker/wasm/字体全部以 blob URL 注入。 */

const LEGACY_PPT_DIR = 'ppt'
const LEGACY_PPT_EXTS = ['ppt', 'dps']

/** 挂载老 ppt 引擎：父窗口建 blob，iframe 内 mount，字节用 iframe 的 Uint8Array 构造。 */
async function mountLegacyPptFrame(host, name) {
  const [indexBytes, workerBytes, wasmBytes, fontBytes] = await Promise.all([
    engineAssetBytes(`${LEGACY_PPT_DIR}/index.mjs`),
    engineAssetBytes(`${LEGACY_PPT_DIR}/worker.mjs`),
    engineAssetBytes(`${LEGACY_PPT_DIR}/ppt-native.wasm`),
    engineAssetBytes(`${LEGACY_PPT_DIR}/ppt-font-cjk.otf`)
  ])
  const blobUrls = [
    URL.createObjectURL(new Blob([indexBytes], { type: 'text/javascript' })),
    URL.createObjectURL(new Blob([workerBytes], { type: 'text/javascript' })),
    URL.createObjectURL(new Blob([wasmBytes], { type: 'application/wasm' })),
    URL.createObjectURL(new Blob([fontBytes], { type: 'font/otf' }))
  ]
  let [indexUrl, workerUrl, wasmUrl, fontUrl] = blobUrls

  // ⚠️ 该引擎在源码里以 `new URL('./ppt-native.wasm', import.meta.url)` 之类的**相对解析**兜底，
  // 而 `import.meta.url` 在 blob 模块里是 `blob:…` ⇒ 相对解析抛 "Invalid URL"（且是急切求值，
  // 传进来的正确 URL 也救不了）。修法：把相对说明符直接换成我们的 blob URL，并给
  // `import.meta.url` 一个无害的绝对基址。
  {
    const decoder = new TextDecoder()
    const swap = (text, pairs) => pairs.reduce((acc, [from, to]) => acc.split(from).join(to), text)
    const indexUrls = [
      ['./ppt-native.wasm', wasmUrl],
      ['./ppt-font-cjk.otf', fontUrl],
      ['./worker.mjs', workerUrl],
      ['./index.mjs', indexUrl]
    ]
    const pairs = indexUrls.flatMap(([spec, url]) => [
      [`'${spec}'`, JSON.stringify(url)],
      [`"${spec}"`, JSON.stringify(url)]
    ])
    const patchedIndex = swap(decoder.decode(indexBytes), pairs).split('import.meta.url').join("'http://localhost/ov-ppt/'")

    URL.revokeObjectURL(indexUrl)
    indexUrl = URL.createObjectURL(new Blob([patchedIndex], { type: 'text/javascript' }))
    blobUrls[0] = indexUrl
  }

  const iframe = document.createElement('iframe')

  iframe.setAttribute('title', name)
  iframe.style.cssText = 'width:100%;height:100%;border:0;display:block;background:#fff'
  iframe.srcdoc = `<!doctype html><html><head><meta charset="utf-8">
<style>html,body{margin:0;height:100%;background:#f4f5f7}
#ov-host{height:100%;min-height:100%;overflow:auto;padding:8px}
.ov-status{padding:14px;font:12px/1.7 -apple-system,"PingFang SC","Helvetica Neue",sans-serif;color:#6b7280}
.ov-status.ov-error{color:#b91c1c;white-space:pre-wrap}</style>
</head><body>
<div id="ov-status" class="ov-status">正在载入旧版演示引擎…（首次较慢，约 10 秒）</div>
<div id="ov-host"></div>
<script type="module">
const status = document.getElementById('ov-status');
const host = document.getElementById('ov-host');
/** 自适应面板**宽度**：对引擎根节点整体 zoom（布局与容器高度一起缩，不是只缩画面）。
 *  引擎默认按原始像素铺开、每页还外套固定尺寸容器 —— 只压 canvas 宽度无效。
 *  按宽度适配（而不是"整页装进面板"）：宽度铺满、页面偏高时纵向滚动，避免两侧留大片空白。 */
function fitToPanel(mounted) {
  const root = (mounted && (mounted.root || mounted.document)) || host.firstElementChild;
  const page = host.querySelector('canvas:not([width="1"])') || host.querySelector('canvas');
  if (!root || !page) return;
  root.style.zoom = '1';
  const natW = page.width || 1210;
  const availW = Math.max(160, host.clientWidth - 8);
  root.style.zoom = String(Math.max(0.1, Math.min(availW / natW, 3)));
  globalThis.__OV_PPT_SCALE__ = root.style.zoom;
}

globalThis.OvLegacyPpt = {
  ready: (async () => {
    const mod = await import(${JSON.stringify(indexUrl)});
    const runtime = await mod.createPptViewer({
      wasmUrl: ${JSON.stringify(wasmUrl)},
      fontUrl: ${JSON.stringify(fontUrl)},
      workerUrl: ${JSON.stringify(workerUrl)},
      worker: 'auto'
    });
    globalThis.__OV_PPT_RUNTIME__ = runtime;
    return true;
  })(),
  mount: async (bytes) => {
    try {
      await globalThis.OvLegacyPpt.ready;
      status.textContent = '正在解析演示文稿…';
      const mounted = await globalThis.__OV_PPT_RUNTIME__.mount(host, bytes);
      globalThis.__OV_PPT_MOUNTED__ = mounted;
      status.remove();
      fitToPanel(mounted);
      // 面板尺寸变化不一定触发 window.resize（右栏拖宽、面板切换都不会）⇒ 盯住宿主本身
      if (globalThis.ResizeObserver) {
        const observer = new ResizeObserver(() => fitToPanel(globalThis.__OV_PPT_MOUNTED__));
        observer.observe(host);
        globalThis.__OV_PPT_RESIZE_OBSERVER__ = observer;
      }
      [60, 400, 1500, 4000].forEach(ms => setTimeout(() => fitToPanel(globalThis.__OV_PPT_MOUNTED__), ms));
      return (host.querySelectorAll('canvas') || []).length;
    } catch (e) {
      status.textContent = '旧版演示引擎渲染失败：' + String((e && e.message) || e);
      status.className = 'ov-status ov-error';
      throw e;
    }
  }
};
</script></body></html>`
  host.appendChild(iframe)

  await new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error('旧版演示引擎加载超时（60s）')), 60_000)

    iframe.addEventListener('load', () => {
      clearTimeout(timer)
      resolve()
    })
  })

  return { iframe, blobUrls }
}

/** 把文件字节交给老 ppt 引擎。 */
async function renderLegacyPpt(iframe, bytes) {
  const win = iframe.contentWindow
  const engine = win?.OvLegacyPpt

  if (!engine || typeof engine.mount !== 'function') throw new Error('旧版演示引擎未就绪')

  return await engine.mount(new win.Uint8Array(bytes))
}

/** 把文件字节交给内核，渲染进 iframe 内的 host。 */
async function renderWithEngine(iframe, bytes, name) {
  const win = iframe.contentWindow
  const status = win.document.getElementById('ov-status')
  const target = win.document.getElementById('ov-host')

  if (status) {
    status.textContent = '正在解析文档…'
    status.classList.remove('ov-error')
  }

  // ⚠️ 字节必须用 iframe 世界的 Uint8Array 构造：跨 realm 递进去，引擎内部的
  // `instanceof Uint8Array` 会判假（旧 .doc 引擎就报 "Unsupported input type"）。
  const bytesInFrame = new win.Uint8Array(bytes)
  const engine = await win.CloudCLIOfficeViewer.render(bytesInFrame, target, { name })

  if (status) status.remove()

  return engine
}

/** 解码文本：先看 BOM，再严格试 UTF-8，失败则回退中文/日文常见编码。
 *  背景：Windows 出的 .txt/.md/.html 多是 GBK，按 UTF-8 硬解就是一堆乱码。 */
function decodeTextBytes(bytes) {
  const head = bytes.subarray(0, 3)

  if (bytes.length > 2 && head[0] === 0xff && head[1] === 0xfe) return new TextDecoder('utf-16le').decode(bytes.subarray(2))
  if (bytes.length > 2 && head[0] === 0xfe && head[1] === 0xff) return new TextDecoder('utf-16be').decode(bytes.subarray(2))
  if (bytes.length > 3 && head[0] === 0xef && head[1] === 0xbb && head[2] === 0xbf) return new TextDecoder('utf-8').decode(bytes.subarray(3))

  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(bytes)
  } catch {
    /* 不是合法 UTF-8：继续试中文/日文编码 */
  }

  for (const encoding of ['gb18030', 'big5', 'shift_jis', 'euc-kr', 'windows-1252']) {
    try {
      const text = new TextDecoder(encoding, { fatal: true }).decode(bytes)

      if (text) return text
    } catch {
      /* 换下一个 */
    }
  }

  return new TextDecoder('utf-8').decode(bytes)
}

/** 读文本：先取字节再智能解码（不要用 fetch().text()，那会按 UTF-8 硬解）。 */
async function readTextSmart(path) {
  try {
    const url = await readDataUrl(path)

    if (url) {
      const bytes = new Uint8Array(await (await fetch(url)).arrayBuffer())

      return decodeTextBytes(bytes)
    }
  } catch (failure) {
    // 超过 data URL 上限时走桌面端给插件的全量文本读取（readFileText 会截断，不能用）。
    const bridge = desktopBridge()

    if (typeof bridge?.readPluginSource === 'function') {
      const result = await bridge.readPluginSource(path)

      if (result && typeof result.text === 'string' && !result.truncated) return result.text
    }

    throw failure
  }

  throw new Error('读不到文件内容')
}

/** 桌面桥单文件读取上限：默认 256 MB（设置 → 聊天 → 本地文件读取上限，最高 4 GB）。 */
const DATA_URL_LIMIT_BYTES = 256 * 1024 * 1024

function isTooLargeFailure(error) {
  return /too large|File preview failed/i.test(String(error?.message || error))
}

function openExternalFile(path) {
  return doors?.os?.openExternal?.(`file://${encodeURI(String(path))}`)
}

function openInAppBrowser(path) {
  const bridge = desktopBridge()
  const target = `file://${encodeURI(String(path))}`

  if (typeof bridge?.openPreviewInBrowser === 'function') {
    void bridge.openPreviewInBrowser(target)

    return true
  }

  return false
}

/** 统一的"预览失败"卡片：把两条退路都摆出来。 */
function PreviewFailure({ message, path, onRetry, onOpenExternal }) {
  const tooLarge = isTooLargeFailure({ message })
  const sizeText = (() => {
    const match = /\((\d+) bytes;/.exec(String(message) || '')

    return match ? `${(Number(match[1]) / 1024 / 1024).toFixed(1)} MB` : ''
  })()

  return jsxs('div', {
    className: 'flex h-full flex-col items-center justify-center gap-3 bg-(--ui-bg-primary) p-6 text-center',
    children: [
      jsx(Codicon, { name: 'warning', className: 'text-2xl text-(--ui-text-tertiary)' }),
      jsx('div', {
        className: 'max-w-[36rem] text-xs text-(--ui-text-secondary)',
        children: tooLarge
          ? `这个文件${sizeText ? ` ${sizeText}` : ''}超过内嵌预览上限（16 MB），不能在面板里渲染。`
          : message
      }),
      !tooLarge ? jsx('pre', { className: 'max-h-28 max-w-[36rem] overflow-auto whitespace-pre-wrap rounded border border-(--ui-stroke-secondary) bg-(--ui-bg-secondary) p-2 text-[0.6875rem] text-(--ui-text-tertiary)', children: message }) : null,
      tooLarge
        ? jsx('div', { className: 'max-w-[36rem] text-[0.6875rem] text-(--ui-text-tertiary)', children: '这是「设置 → 聊天 → 预览 / 图片加载大小上限」的额度，最高可调到 4 GB。下面第一个按钮会把它调到 256 MB，然后重试。' })
        : null,
      jsxs('div', {
        className: 'flex items-center gap-2',
        children: [
          tooLarge && onRetry
            ? jsx(BarButton, {
              icon: 'settings-gear',
              label: '调高上限到 256 MB 并重试',
              onClick: () => {
                void (async () => {
                  try {
                    await desktopBridge()?.dataUrlReadMax?.set?.(256)
                  } catch {
                    /* 设不了就照旧重试，让上层报错 */
                  }
                  onRetry()
                })()
              }
            })
            : null,
          !tooLarge && onRetry ? jsx(BarButton, { icon: 'refresh', label: '重试', onClick: onRetry }) : null,
          onOpenExternal ? jsx(BarButton, { icon: 'link-external', label: '用系统程序打开', onClick: onOpenExternal }) : null,
          path ? jsx(BarButton, { icon: 'browser', label: '在应用浏览器打开', onClick: () => { host.notify?.({ kind: 'info', message: openInAppBrowser(path) ? '已在浏览器面板打开' : '当前版本没有这个入口' }) } }) : null
        ]
      })
    ]
  })
}

/** 走桌面桥把文件读成字节。 */
async function readFileBytes(path) {
  const url = await readDataUrl(path)

  if (!url) throw new Error('读不到文件内容')

  const response = await fetch(url)

  return new Uint8Array(await response.arrayBuffer())
}

/** 查看器里的"新内核"视图：自己管 iframe 的生死。 */
function EngineView({ path, name, onOpenExternal }) {
  const hostRef = useRef(null)
  const [error, setError] = useState('')
  const [engine, setEngine] = useState('')
  const [nonce, setNonce] = useState(0)

  useEffect(() => {
    let cancelled = false
    // ⚠️ 必须在 effect 作用域声明：放进下面的 IIFE 里，cleanup 就引用不到了
    //（关闭预览会触发 unmount ⇒ ReferenceError ⇒ 宿主报 "failed to render"）。
    let engineUrl = ''
    let extraBlobUrls = []
    const host = hostRef.current

    if (!host) return undefined

    host.innerHTML = ''
    setError('')
    setEngine('')

    ;(async () => {
      try {
        // 旧版二进制演示（.ppt/.dps）：内核解不了 OLE2 演示容器。
        const legacy = LEGACY_PPT_EXTS.includes(extOf(path))

        if (legacy) {
          // 插件自带的旧 ppt 引擎：零外部依赖、无服务、无临时文件（每页带第三方水印）。
          const mounted = await mountLegacyPptFrame(host, name)

          extraBlobUrls = mounted.blobUrls
          if (cancelled) return

          const bytes = await readFileBytes(path)

          if (cancelled) return

          const pages = await renderLegacyPpt(mounted.iframe, bytes)

          if (!cancelled) setEngine(`legacy-ppt ${pages}`)

          return
        }

        const mounted = await mountEngineFrame(host, name)

        engineUrl = mounted.engineUrl
        if (cancelled) return

        const bytes = await readFileBytes(path)

        if (cancelled) return

        const tag = await renderWithEngine(mounted.iframe, bytes, name)

        if (!cancelled) setEngine(String(tag || ''))
      } catch (failure) {
        if (!cancelled) setError(String(failure?.message || failure))
      }
    })()

    return () => {
      cancelled = true
      host.innerHTML = ''
      if (engineUrl) URL.revokeObjectURL(engineUrl)
      extraBlobUrls.forEach(url => URL.revokeObjectURL(url))
    }
  }, [path, name, nonce])

  return jsxs('div', {
    className: 'relative h-full min-h-0',
    children: [
      jsx('div', { ref: hostRef, className: 'h-full w-full' }),
      error
        ? jsx('div', {
          className: 'absolute inset-0',
          children: jsx(PreviewFailure, {
            message: `新内核渲染失败：${error}`,
            path,
            onRetry: () => setNonce(n => n + 1),
            onOpenExternal: () => (onOpenExternal ? onOpenExternal() : openExternalFile(path))
          })
        })
        : null
    ]
  })
}


/* ─────────────── 文本 / Markdown / HTML / PDF / 图片 预览（不依赖任何外部服务） ─────────────── */

const MARKED_ASSET = 'marked.umd.js'

const MARKDOWN_EXTS = ['md', 'markdown', 'mdx']
const HTML_PREVIEW_EXTS = ['html', 'htm']
const TEXT_PREVIEW_EXTS = [
  'txt', 'log', 'text', 'json', 'yml', 'yaml', 'toml', 'ini', 'conf', 'env',
  'py', 'js', 'mjs', 'cjs', 'ts', 'tsx', 'jsx', 'sh', 'bash', 'zsh', 'fish',
  'css', 'scss', 'less', 'sql', 'go', 'rs', 'java', 'kt', 'c', 'h', 'cpp', 'hpp',
  'rb', 'php', 'lua', 'pl', 'swift', 'vue', 'svelte', 'xml', 'csv', 'tsv', 'diff', 'patch'
]

function escapeHtmlText(text) {
  return String(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

/** 文本 / md / html 的 iframe 文档：内容以 text/plain 内联，文件里的脚本一律不执行。 */
function textBootstrap({ mode, payload, markedUrl, title }) {
  const css = `html,body{margin:0;height:100%;background:#fff;color:#1f2328}
body{font:13px/1.75 -apple-system,"PingFang SC","Helvetica Neue",sans-serif}
#ov-host{max-width:52rem;margin:0 auto;padding:2rem 2.25rem 4rem}
pre.ov-text{margin:0;padding:1rem 1.25rem;font:12px/1.7 ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;white-space:pre-wrap;word-break:break-word;tab-size:4;color:#1f2328}
h1,h2,h3,h4{line-height:1.35;margin:1.6em 0 .6em;font-weight:600}
h1{font-size:1.7em;border-bottom:1px solid #e5e7eb;padding-bottom:.3em}
h2{font-size:1.35em;border-bottom:1px solid #eef0f2;padding-bottom:.25em}
h3{font-size:1.15em}
p,ul,ol,blockquote,table{margin:.85em 0}
ul,ol{padding-left:1.6em}li{margin:.25em 0}
a{color:#0053fd;text-decoration:none}a:hover{text-decoration:underline}
code{font-family:ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;font-size:.92em;background:#f4f5f7;padding:.12em .35em;border-radius:4px}
pre{background:#f6f7f9;border:1px solid #e5e7eb;border-radius:6px;padding:.9em 1em;overflow:auto}
pre code{background:none;padding:0}
blockquote{border-left:3px solid #d8dbe0;color:#57606a;padding:.1em 0 .1em 1em;margin-left:0}
table{border-collapse:collapse;width:100%}th,td{border:1px solid #e5e7eb;padding:.5em .7em;text-align:left}
th{background:#fafbfc;font-weight:600}img{max-width:100%}
hr{border:none;border-top:1px solid #e5e7eb;margin:1.8em 0}
.ov-empty{color:#9aa0a6;font-style:italic}`

  return `<!doctype html><html><head><meta charset="utf-8"><title>${escapeHtmlText(title)}</title>
<style>${css}</style></head><body><div id="ov-host"></div>
<script type="text/plain" id="ov-payload">${String(payload).replace(/<\/script/gi, '<\\/script')}</script>
${mode === 'markdown' && markedUrl ? `<script src="${markedUrl}"></script>` : ''}
<script>
(function () {
  var host = document.getElementById('ov-host');
  var node = document.getElementById('ov-payload');
  var raw = (node && node.textContent) || '';
  var esc = function (text) { return text.replace(/[&<>]/g, function (c) { return ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' })[c] }) };
  try {
    if ('${mode}' === 'markdown' && globalThis.marked && typeof globalThis.marked.parse === 'function') {
      host.innerHTML = globalThis.marked.parse(raw, { gfm: true });
    } else if ('${mode}' === 'html') {
      var doc = new DOMParser().parseFromString(raw, 'text/html');
      doc.querySelectorAll('script').forEach(function (n) { n.remove() });
      (doc.head ? doc.head.querySelectorAll('style,link[rel=stylesheet]') : []).forEach(function (n) { document.head.appendChild(n.cloneNode(true)) });
      host.innerHTML = doc.body ? doc.body.innerHTML : esc(raw);
    } else {
      host.innerHTML = '<pre class="ov-text">' + esc(raw) + '</pre>';
    }
    if (!host.textContent.trim()) host.innerHTML = '<div class="ov-empty">（空文件）</div>';
  } catch (e) {
    host.innerHTML = '<pre class="ov-text">渲染失败：' + esc(String((e && e.message) || e)) + '</pre>';
  }
})();
</script></body></html>`
}

/** 文本 / Markdown / HTML 视图。 */
function TextView({ path, name, kind }) {
  const hostRef = useRef(null)
  const [error, setError] = useState('')

  useEffect(() => {
    let cancelled = false
    let markedUrl = ''
    const host = hostRef.current

    if (!host) return undefined

    host.innerHTML = ''
    setError('')

    ;(async () => {
      try {
        const payload = await readTextSmart(path)

        if (cancelled) return

        if (kind === 'markdown') {
          const marked = await engineAssetText(MARKED_ASSET).catch(() => '')

          if (marked) markedUrl = URL.createObjectURL(new Blob([marked], { type: 'text/javascript' }))
        }

        const iframe = document.createElement('iframe')

        iframe.setAttribute('title', name)
        iframe.style.cssText = 'width:100%;height:100%;border:0;display:block;background:#fff'
        iframe.srcdoc = textBootstrap({ mode: kind, payload, markedUrl, title: name })
        host.appendChild(iframe)
      } catch (failure) {
        if (!cancelled) setError(String(failure?.message || failure))
      }
    })()

    return () => {
      cancelled = true
      host.innerHTML = ''
      if (markedUrl) URL.revokeObjectURL(markedUrl)
    }
  }, [path, name, kind])

  return jsxs('div', {
    className: 'relative h-full min-h-0',
    children: [
      jsx('div', { ref: hostRef, className: 'h-full w-full' }),
      error
        ? jsxs('div', {
          className: 'absolute inset-0 flex items-center justify-center gap-2 bg-(--ui-bg-primary) p-6 text-center text-xs text-(--ui-text-secondary)',
          children: [jsx(Codicon, { name: 'warning' }), `读不了这个文件：${error}`]
        })
        : null
    ]
  })
}

/** 图片视图。 */
function ImageView({ path, name }) {
  const [url, setUrl] = useState('')
  const [error, setError] = useState('')

  useEffect(() => {
    let cancelled = false

    setUrl('')
    setError('')
    readDataUrl(path)
      .then(next => { if (!cancelled) (next ? setUrl(next) : setError('读不到文件内容')) })
      .catch(failure => { if (!cancelled) setError(String(failure?.message || failure)) })

    return () => { cancelled = true }
  }, [path])

  if (error) return jsx('div', { className: 'flex h-full items-center justify-center text-xs text-(--ui-text-secondary)', children: `读不了这张图：${error}` })

  return jsx('div', {
    className: 'flex h-full items-center justify-center overflow-auto bg-(--ui-bg-secondary) p-2',
    children: url ? jsx('img', { src: url, alt: name, className: 'max-h-full max-w-full object-contain' }) : jsx('span', { className: 'text-xs text-(--ui-text-tertiary)', children: '正在读取…' })
  })
}

/** PDF 视图：pdf.js（4.10.38，随插件携带）逐页渲染成 canvas，可滚动连看。 */
function PdfView({ path, name }) {
  const hostRef = useRef(null)
  const [error, setError] = useState('')
  const [nonce, setNonce] = useState(0)

  useEffect(() => {
    let cancelled = false
    let blobUrls = []
    const host = hostRef.current

    if (!host) return undefined

    host.innerHTML = ''
    setError('')

    ;(async () => {
      try {
        const [pdfSrc, workerSrc, dataUrl] = await Promise.all([
          engineAssetText('pdf.min.mjs'),
          engineAssetText('pdf.worker.min.mjs'),
          readDataUrl(path)
        ])

        if (cancelled) return
        if (!dataUrl) throw new Error('读不到文件内容')

        const pdfUrl = URL.createObjectURL(new Blob([pdfSrc], { type: 'text/javascript' }))
        const workerUrl = URL.createObjectURL(new Blob([workerSrc], { type: 'text/javascript' }))

        blobUrls = [pdfUrl, workerUrl]

        const iframe = document.createElement('iframe')

        iframe.setAttribute('title', name)
        iframe.style.cssText = 'width:100%;height:100%;border:0;display:block;background:#f4f5f7'
        iframe.srcdoc = `<!doctype html><html><head><meta charset="utf-8">
<style>html,body{margin:0;background:#f4f5f7}
#ov-host{display:flex;flex-direction:column;align-items:center;gap:14px;padding:14px 0 28px}
canvas{background:#fff;box-shadow:0 1px 4px rgba(15,23,42,.18);border-radius:2px;max-width:100%;height:auto}
.ov-status{font:12px/1.7 -apple-system,"PingFang SC",sans-serif;color:#6b7280;padding:14px}
.ov-status.ov-error{color:#b91c1c;white-space:pre-wrap}</style></head>
<body><div class="ov-status" id="ov-status">正在载入 PDF 引擎…</div><div id="ov-host"></div>
<script type="text/plain" id="ov-pdf">${String(dataUrl).replace(/<\/script/gi, '<\\/script')}</script>
<script type="module">
const status = document.getElementById('ov-status');
const host = document.getElementById('ov-host');
try {
  const pdfjs = await import(${JSON.stringify(pdfUrl)});
  pdfjs.GlobalWorkerOptions.workerSrc = ${JSON.stringify(workerUrl)};
  const raw = (document.getElementById('ov-pdf').textContent || '');
  const bytes = new Uint8Array(await (await fetch(raw)).arrayBuffer());
  const doc = await pdfjs.getDocument({ data: bytes }).promise;
  status.textContent = '共 ' + doc.numPages + ' 页，正在渲染…';
  const dpr = Math.min(window.devicePixelRatio || 1, 2);
  for (let n = 1; n <= doc.numPages; n += 1) {
    const page = await doc.getPage(n);
    const base = page.getViewport({ scale: 1 });
    const width = Math.min(base.width, host.clientWidth ? host.clientWidth - 28 : base.width);
    const viewport = page.getViewport({ scale: width / base.width });
    const canvas = document.createElement('canvas');
    canvas.width = Math.floor(viewport.width * dpr);
    canvas.height = Math.floor(viewport.height * dpr);
    canvas.style.width = Math.floor(viewport.width) + 'px';
    host.appendChild(canvas);
    await page.render({ canvasContext: canvas.getContext('2d'), viewport, transform: dpr === 1 ? undefined : [dpr, 0, 0, dpr, 0, 0] }).promise;
  }
  status.remove();
} catch (e) {
  status.textContent = 'PDF 渲染失败：' + String((e && e.message) || e);
  status.className = 'ov-status ov-error';
}
</script></body></html>`
        host.appendChild(iframe)
      } catch (failure) {
        if (!cancelled) setError(String(failure?.message || failure))
      }
    })()

    return () => {
      cancelled = true
      host.innerHTML = ''
      blobUrls.forEach(url => URL.revokeObjectURL(url))
    }
  }, [path, name, nonce])

  return jsxs('div', {
    className: 'relative h-full min-h-0',
    children: [
      jsx('div', { ref: hostRef, className: 'h-full w-full' }),
      error
        ? jsx('div', {
          className: 'absolute inset-0',
          children: jsx(PreviewFailure, { message: `读不了这个 PDF：${error}`, path, onRetry: () => setNonce(n => n + 1), onOpenExternal: () => openExternalFile(path) })
        })
        : null
    ]
  })
}

/** HTML 预览：内容进 sandbox iframe（保留脚本但拿不到宿主来源），超 5 MB 降级为源码。 */
const HTML_PREVIEW_MAX_BYTES = 5 * 1024 * 1024

function HtmlPreview({ path, name }) {
  const hostRef = useRef(null)
  const [error, setError] = useState('')

  useEffect(() => {
    let cancelled = false
    const host = hostRef.current

    if (!host) return undefined

    host.innerHTML = ''
    setError('')

    ;(async () => {
      try {
        const html = await readTextSmart(path)

        if (cancelled) return

        if (html.length > HTML_PREVIEW_MAX_BYTES) {
          setError('文件超过 5 MB，已按源码显示')

          const iframe = document.createElement('iframe')

          iframe.setAttribute('title', name)
          iframe.style.cssText = 'width:100%;height:100%;border:0;background:#fff'
          iframe.srcdoc = textBootstrap({ mode: 'text', payload: html, markedUrl: '', title: name })
          host.appendChild(iframe)

          return
        }

        const iframe = document.createElement('iframe')

        iframe.setAttribute('title', name)
        // 隔离：允许脚本运行（保真的前提），但**不给同源** ⇒ 页面里的脚本碰不到 Hermes 本体。
        iframe.setAttribute('sandbox', 'allow-scripts allow-forms allow-modals allow-popups')
        iframe.setAttribute('referrerpolicy', 'no-referrer')
        iframe.style.cssText = 'width:100%;height:100%;border:0;display:block;background:#fff'
        iframe.srcdoc = html
        host.appendChild(iframe)
      } catch (failure) {
        if (!cancelled) setError(String(failure?.message || failure))
      }
    })()

    return () => { cancelled = true; host.innerHTML = '' }
  }, [path, name])

  return jsxs('div', {
    className: 'relative h-full min-h-0',
    children: [
      jsx('div', { ref: hostRef, className: 'h-full w-full' }),
      error
        ? jsx('div', { className: 'absolute inset-x-0 top-0 bg-(--ui-bg-secondary) px-2 py-1 text-[0.6875rem] text-(--ui-text-tertiary)', children: error })
        : null
    ]
  })
}

/* ─────────────── 浏览器本体 ─────────────── */

function FileBrowser() {
  const [dir, setDir] = useState(WORKSPACE_HOME)
  const [trail, setTrail] = useState([WORKSPACE_HOME])
  const [cursor, setCursor] = useState(0)
  const [data, setData] = useState(null)
  const [error, setError] = useState(null)
  const [busy, setBusy] = useState(false)
  const [showHidden, setShowHidden] = useState(false)
  const [nonce, setNonce] = useState(0)

  useEffect(() => {
    let cancelled = false

    setBusy(true)
    setError(null)

    const run = async () => {
      try {
        const payload = await listDirectory(dir)

        if (cancelled) return

        setData(payload)
      } catch (err) {
        if (!cancelled) setError(String(err?.message || err))
      } finally {
        if (!cancelled) setBusy(false)
      }
    }

    void run()

    return () => { cancelled = true }
  }, [dir, nonce])

  const go = useCallback((target, push = true) => {
    if (!target || target === dir) return

    setDir(target)

    if (push) {
      setTrail(prev => {
        const next = prev.slice(0, cursor + 1)

        next.push(target)
        setCursor(next.length - 1)

        return next
      })
    }
  }, [cursor, dir])

  const back = () => { if (cursor > 0) { setCursor(cursor - 1); setDir(trail[cursor - 1]) } }
  const forward = () => { if (cursor < trail.length - 1) { setCursor(cursor + 1); setDir(trail[cursor + 1]) } }

  const entries = (data?.entries ?? []).filter(item => showHidden || !item.hidden)

  // 路径栏：`/Users/you/项目` → [💻][Users][you][项目]，每一段都能点
  const crumbs = (() => {
    const parts = String(data?.path || dir).split('/').filter(Boolean)
    const out = [{ label: '', path: '/', root: true }]
    let acc = ''

    for (const part of parts) {
      acc += `/${part}`
      out.push({ label: part, path: acc })
    }

    return out
  })()

  return jsxs('div', {
    className: 'flex h-full min-h-0 w-full flex-col text-xs',
    children: [
      // 工具条 + 路径栏
      jsxs('div', {
        className: 'shrink-0 border-b border-(--ui-border) px-1.5 py-1',
        children: [
          jsxs('div', {
            className: 'flex items-center gap-0.5',
            children: [
              jsx(BarButton, { icon: 'arrow-up', tip: '上一级', disabled: !data?.parent, onClick: () => go(data?.parent) }),
              jsx(BarButton, { icon: 'chevron-left', tip: '后退', disabled: cursor <= 0, onClick: back }),
              jsx(BarButton, { icon: 'chevron-right', tip: '前进', disabled: cursor >= trail.length - 1, onClick: forward }),
              jsx('span', { className: 'mx-1 h-4 w-px bg-(--ui-border)' }),
              jsx(BarButton, { icon: 'home', tip: '工作区', onClick: () => go(WORKSPACE_HOME) }),
              jsx('span', { className: 'mx-1 flex-1' }),
              jsx(BarButton, { icon: showHidden ? 'eye' : 'eye-closed', tip: showHidden ? '不显示隐藏项' : '显示隐藏项', onClick: () => setShowHidden(!showHidden) }),
              jsx(BarButton, { icon: 'refresh', tip: '刷新', onClick: () => setNonce(count => count + 1) }),
              jsx(BarButton, { icon: 'folder-opened', tip: '用系统对话框选一个文件', onClick: async () => {
                try {
                  const picked = await doors?.os?.pickOpenPath?.({ title: '选择要点开的文件' })

                  if (picked) $file.set(picked)
                } catch { /* 取消即无事 */ }
              } })
            ]
          }),
          jsxs('div', {
            className: 'mt-1 flex flex-wrap items-center gap-0.5 text-(--ui-text-secondary)',
            children: crumbs.flatMap((crumb, index) => {
              const node = jsx('button', {
                type: 'button',
                className: 'max-w-[12rem] truncate rounded-[4px] px-1.5 py-0.5 hover:bg-(--ui-bg-editor)',
                onClick: () => go(crumb.path),
                title: crumb.path,
                children: crumb.root ? jsx(Codicon, { name: 'device-desktop' }) : crumb.label
              }, `crumb-${index}`)

              return index === 0 ? [node] : [jsx('span', { className: 'text-(--ui-text-quaternary)', children: '›' }, `sep-${index}`), node]
            })
          })
        ]
      }),
      // 目录内容
      jsx(ScrollArea, {
        className: 'min-h-0 flex-1',
        children: jsxs('div', {
          className: 'py-1',
          children: [
            error
              ? jsx('div', { className: 'px-3 py-2 text-(--ui-text-tertiary)', children: `读不了这个目录：${error}` })
              : null,
            busy && !data ? jsx('div', { className: 'px-3 py-2 text-(--ui-text-tertiary)', children: '正在读取…' }) : null,
            !error && data && entries.length === 0
              ? jsx('div', { className: 'px-3 py-2 text-(--ui-text-tertiary)', children: '这个目录里没有可显示的内容' })
              : null,
            ...entries.map(item => {

              return jsx('button', {
                type: 'button',
                className: 'flex w-full items-center gap-2 rounded-[4px] px-2 py-1 text-left hover:bg-(--ui-bg-editor)',
                onClick: () => (item.dir ? go(item.path) : $file.set(item.path)),
                title: item.path,
                children: [
                  item.dir
                    ? jsx(Codicon, { name: 'folder', className: 'shrink-0 text-(--ui-text-tertiary)' })
                    : jsx(FileGlyph, { path: item.path, className: 'shrink-0 text-(--ui-text-tertiary)' }),
                  jsx('span', { className: 'min-w-0 flex-1 truncate', children: item.name }),
                  jsx('span', { className: 'w-16 shrink-0 text-right text-(--ui-text-quaternary)', children: item.dir ? '' : humanSize(item.size) }),
                  jsx('span', { className: 'w-28 shrink-0 text-right text-(--ui-text-quaternary)', children: item.dir ? '' : shortDate(item.mtime) })
                ]
              }, item.path)
            })
          ]
        })
      })
    ]
  })
}

/* ─────────────────────────── 查看器本体 ─────────────────────────── */

function Viewer({ path, compact }) {
  const nonce = useValue($nonce)
  const [state, setState] = useState({ status: 'idle' })
  const frameRef = useRef(null)

  useEffect(() => {
    if (!path) {
      setState({ status: 'idle' })
      return
    }
    let cancelled = false
    setState({ status: 'loading' })

    // Office 系（docx/xlsx/pptx/doc/xls/wps/et/dps/rtf/ofd/csv）走纯 JS 内核：
    // 不起服务、不依赖 LibreOffice、不转 HTML，字节直接喂给引擎渲染。
    if (ENGINE_EXTS.includes(extOf(path))) {
      setState({ status: 'engine' })

      return () => { cancelled = true }
    }

    const ext = extOf(path)

    // PDF → 内嵌阅读器；图片 → 直接内嵌（都走桌面桥，不经过任何转换器）。
    if (ext === 'pdf') {
      setState({ status: 'pdf' })

      return () => { cancelled = true }
    }

    if (DATAURL_EXTS.includes(ext)) {
      setState({ status: 'image' })

      return () => { cancelled = true }
    }

    // Markdown（marked 渲染）/ HTML（内嵌，禁脚本）/ 纯文本与脚本（等宽只读）。
    if (MARKDOWN_EXTS.includes(ext)) {
      setState({ status: 'text', kind: 'markdown' })

      return () => { cancelled = true }
    }

    if (HTML_PREVIEW_EXTS.includes(ext)) {
      setState({ status: 'html' })

      return () => { cancelled = true }
    }

    if (TEXT_PREVIEW_EXTS.includes(ext)) {
      setState({ status: 'text', kind: 'text' })

      return () => { cancelled = true }
    }

    // 其余（旧二进制 .ppt / 认不出的容器）：不再有任何后端转换，直接说明并给"用系统程序打开"。
    beacon('viewer-unsupported', `${path} | ${ext}`)
    setState({ status: 'unsupported', meta: { ext } })

    return () => { cancelled = true }
  }, [path, nonce])

  const refresh = useCallback(() => {
    $nonce.set($nonce.get() + 1)
  }, [])

  const revealInFinder = useCallback(() => {
    void doors?.os?.revealPath?.(path)
  }, [path])

  const openExternal = useCallback(() => {
    void doors?.os?.openExternal?.(`file://${encodeURI(path)}`)
  }, [path])

  const copyPath = useCallback(() => {
    void doors?.os?.writeClipboard?.(path)
    host.notify?.({ kind: 'info', message: '路径已复制' })
  }, [path])

  if (!path) {
    // 用户口径（2026-09-22）：面板默认态不再是「还没有选文件」，而是**文件浏览器**本身
    // ——它要取代「文件」标签的全部能力（浏览 + 就地预览）。
    return jsx(FileBrowser, {})
  }

  return jsxs('div', {
    className: 'flex h-full min-h-0 flex-col',
    children: [
      jsxs('div', {
        className: cn(
          'flex shrink-0 items-center gap-1 border-b px-2 py-1.5',
          'border-(--ui-stroke-secondary) text-(--ui-text-secondary)'
        ),
        children: [
          jsx(FileGlyph, { path, className: 'shrink-0 text-[0.95em]' }),
          jsx(Tip, {
            label: path,
            children: jsx('span', {
              className: 'min-w-0 flex-1 truncate text-xs',
              children: baseName(path)
            })
          }),
          state.meta?.size ? jsx('span', { className: 'shrink-0 text-[0.6875rem] text-(--ui-text-quaternary)', children: fmtSize(state.meta.size) }) : null,
          jsx(BarButton, { icon: 'refresh', label: '', tip: '重新渲染', onClick: refresh }),
          jsx(BarButton, { icon: 'folder-opened', label: '', tip: '在访达中显示', onClick: revealInFinder }),
          jsx(BarButton, { icon: 'link-external', label: '', tip: '用系统程序打开', onClick: openExternal }),
          jsx(BarButton, { icon: 'copy', label: '', tip: '复制路径', onClick: copyPath }),
          OFFICE_EXTS_FOR_PDF.includes(extOf(path)) ? jsx(ExactPdfButton, { path }) : null,
          // 关闭预览：回到文件浏览（用户点名要的那颗按钮）。
          jsx(BarButton, { icon: 'close', label: '', tip: '关闭预览（回到文件浏览）', onClick: () => $file.set(null) })
        ]
      }),
      jsxs('div', {
        className: 'relative min-h-0 flex-1',
        children: [
          state.status === 'loading'
            ? jsxs('div', {
                className: 'flex h-full items-center justify-center gap-2 text-xs text-(--ui-text-tertiary)',
                children: [jsx(Codicon, { name: 'loading', spinning: true, className: 'text-base' }), '正在渲染…（首次打开某格式可能几秒）']
              })
            : null,
          state.status === 'error'
            ? jsxs('div', {
                className: 'flex h-full flex-col items-center justify-center gap-2 p-6 text-center',
                children: [
                  jsx(Codicon, { name: 'warning', className: 'text-2xl text-(--ui-text-tertiary)' }),
                  jsx('div', { className: 'text-xs text-(--ui-text-secondary)', children: '渲染失败' }),
                  jsx('pre', {
                    className: 'max-h-40 overflow-auto whitespace-pre-wrap rounded border border-(--ui-stroke-secondary) bg-(--ui-bg-secondary) p-2 text-[0.6875rem] text-(--ui-text-tertiary)',
                    children: state.message
                  }),
                  jsx(Button, { size: 'xs', onClick: refresh, children: '重试' })
                ]
              })
            : null,
          state.status === 'unsupported'
            ? jsxs('div', {
                className: 'flex h-full flex-col items-center justify-center gap-2 p-6 text-center',
                children: [
                  jsx(Codicon, { name: 'question', className: 'text-2xl text-(--ui-text-tertiary)' }),
                  jsx('div', { className: 'text-xs text-(--ui-text-secondary)', children: `${state.meta?.ext || '该'} 还不支持（本插件内置转换器覆盖 ${state.meta?.known_formats ?? '—'} 种扩展名）` }),
                  jsx(BarButton, { icon: 'link-external', label: '用系统程序打开', onClick: openExternal })
                ]
              })
            : null,
          state.status === 'engine'
            ? jsx(EngineView, { path, name: baseName(path), onOpenExternal: openExternal })
            : null,
          state.status === 'text'
            ? jsx(TextView, { path, name: baseName(path), kind: state.kind })
            : null,
          state.status === 'html'
            ? jsx(HtmlPreview, { path, name: baseName(path) })
            : null,
          state.status === 'pdf'
            ? jsx(PdfView, { path, name: baseName(path) })
            : null,
          state.status === 'image'
            ? jsx(ImageView, { path, name: baseName(path) })
            : null,
          state.status === 'ready'
            ? jsx('iframe', {
                ref: frameRef,
                title: baseName(path),
                srcDoc: state.html,
                sandbox: 'allow-scripts allow-popups allow-popups-to-escape-sandbox allow-forms',
                className: 'h-full w-full border-0 bg-white',
                style: { colorScheme: 'light' }
              })
            : null
        ]
      }),
    ]
  })
}

/* ─────────────────────────── 三种宿主界面 ─────────────────────────── */

/** 侧栏面板。 */
function ViewerPane() {
  const file = useValue($file)
  useEffect(() => {
    beacon('pane-mount', '')
  }, [])
  return jsx(Viewer, { path: file, compact: false })
}

/** 整页视图。 */
function ViewerPage() {
  const file = useValue($file)
  return jsx('div', {
    className: 'flex h-full min-h-0 flex-col',
    children: jsx(Viewer, { path: file, compact: false })
  })
}

/**
 * 消息内产物卡片：`::office{file="…"}`
 *
 * 用户口径（2026-09-21）：点卡片 → 侧栏自动呼出查看器。
 * 所以卡片本身就是按钮，点击做两件事：记住文件 + revealPane。
 */
function OfficeCard({ path, label }) {
  const [open, setOpen] = useState(false)   // 仅当显式点「展开」时才用得着
  const [expandable, setExpandable] = useState(false)
  const name = label || baseName(path)
  const valid = typeof path === 'string' && path.startsWith('/')

  if (!valid) {
    return jsx('div', {
      className: 'rounded border border-(--ui-stroke-secondary) px-3 py-2 text-xs text-(--ui-text-tertiary)',
      children: 'Office 查看器：卡片参数无效（需要绝对路径的 file="…"）'
    })
  }

  return jsxs('div', {
    className: 'my-2 overflow-hidden rounded-lg border border-(--ui-stroke-secondary)',
    children: [
      jsxs('div', {
        className: 'flex items-center gap-2 px-3 py-2',
        children: [
          jsx(Codicon, { name: 'file-media', className: 'text-base text-(--ui-text-tertiary)' }),
          jsx('div', {
            className: 'min-w-0 flex-1 cursor-pointer text-sm text-(--ui-text-secondary)',
            title: `${path}\n（点击在右侧栏打开）`,
            onClick: () => openInPane(path),
            children: name
          }),
          jsx(Tip, {
            label: '在右侧栏打开（⌘J 可切换右列）',
            children: jsx(Button, {
              size: 'xs',
              onClick: () => openInPane(path),
              children: '预览'
            })
          }),
          jsx(BarButton, {
            icon: 'chevron-down',
            label: '',
            tip: '在本条消息里内联展开',
            onClick: () => {
              setExpandable(true)
              setOpen(true)
            }
          }),
          jsx(BarButton, { icon: 'link-external', label: '', tip: '用系统程序打开', onClick: () => void doors?.os?.openExternal?.(`file://${encodeURI(path)}`) })
        ]
      }),
      open && expandable
        ? jsx('div', {
            className: 'h-[60vh] min-h-64 border-t border-(--ui-stroke-secondary)',
            children: jsx(Viewer, { path, compact: true })
          })
        : null
    ]
  })
}

/* ─────────────────────────── 注册 ─────────────────────────── */

export default {
  id: PLUGIN_ID,
  name: 'Office 查看器',
  description: '在本机渲染 Office / WPS / iWork / OFD / epub / 邮件 / 压缩包 / 设计稿等长尾格式，点产物卡片即在侧栏直接看。零常驻服务。',
  defaultEnabled: true,

  register(ctx) {
    doors = { rest: ctx.rest, os: ctx.os, storage: ctx.storage }
    beacon('register', '桌面半已加载')

    // 别的插件请本插件预览某个文件（约定事件，见 EXTERNAL_OPEN_EVENT 注释）。
    // removeEventListener 先跑一遍：重复 register 时不叠加监听。
    window.removeEventListener(EXTERNAL_OPEN_EVENT, onExternalOpenRequest)
    window.addEventListener(EXTERNAL_OPEN_EVENT, onExternalOpenRequest)
    ctx.onDispose?.(() => window.removeEventListener(EXTERNAL_OPEN_EVENT, onExternalOpenRequest))

    ctx.registerMany([
      {
        id: PANE_KEY,
        area: PANES_AREA,
        // 标签名按用户口径改（2026-09-22）：只动这一处标签名；插件自己的 ⌘K 命令与
        // 「在主区打开」的页标题仍是「Office 查看器」，那是另一处文案，用户点名的是标签。
        title: '文档预览',
        // 落位（2026-09-21 实测修正）：必须**进右列**，不能锚到 workspace。
        //   · placement:'right'  → 语义角色＝右列（桌面端源码：非 left/main 即 'right'）；
        //   · dock.pos:'center'  → 以「文件浏览器」为锚**叠成标签**（而不是分割出新列）；
        //   · 之前误锚 workspace+pos:'right'，实测落进主区那一组（和聊天/review 并排），
        //     所以点卡片看着像"在聊天区打开"，而不是右侧栏。enforceDockedPanes 会把
        //     已在树里的面板按 dock 归位，故改这条声明即可自动搬回右列，无需手动拖。
        //   · uncloseable  → 桌面端"关掉唯一面板＝禁用整个插件"，锁掉 × 防误关。
        data: {
          placement: 'right',
          dock: { pane: 'file-browser', pos: 'center' },
          width: '460px',
          uncloseable: true
        },
        render: () => jsx(ViewerPane, {})
      },
      {
        id: 'page',
        area: ROUTES_AREA,
        data: { path: ROUTE },
        render: () => jsx(ViewerPage, {})
      },
      {
        id: 'nav',
        area: SIDEBAR_NAV_AREA,
        data: { path: ROUTE, label: '文档预览', codicon: 'preview' } // 用户口径：侧栏这一行也叫「文档预览」
      },
      {
        id: 'cmd-open',
        area: PALETTE_AREA,
        data: {
          id: 'office-viewer.open',
          label: 'Office 查看器：打开面板',
          keywords: ['office', 'office-viewer', '查看器', 'preview'],
          run: () => openInPane(null)
        }
      },
      {
        id: 'cmd-pick',
        area: PALETTE_AREA,
        data: {
          id: 'office-viewer.pick',
          label: 'Office 查看器：选择文件…',
          keywords: ['office', 'office-viewer', '选择', '文件'],
          run: async () => {
            const picked = await doors?.os?.pickOpenPath?.({ title: '选择要预览的文档' })
            if (picked) openInPane(picked)
          }
        }
      },
      {
        id: 'cmd-refresh',
        area: PALETTE_AREA,
        data: {
          id: 'office-viewer.refresh',
          label: 'Office 查看器：重新渲染当前文件',
          keywords: ['office', 'viewer', '刷新', '重渲染'],
          run: () => $nonce.set($nonce.get() + 1)
        }
      },
      {
        id: 'open-pane',
        area: KEYBINDS_AREA,
        data: {
          id: 'office-viewer.open-pane',
          label: 'Office 查看器：呼出面板',
          category: 'Office 查看器',
          defaults: ['mod+shift+o'],
          run: () => openInPane(null)
        }
      },
      {
        id: 'card',
        area: TRANSCRIPT_DIRECTIVE_AREA,
        data: {
          name: 'office',
          render: ({ attrs }) => jsx(OfficeCard, { path: attrs?.file || attrs?.path, label: attrs?.name })
        }
      }
    ])
  }
}
