# hermes-office-viewer

> **Hermes 桌面端的文档预览插件** —— 右侧栏一个「文档预览」标签，等于一个类访达的文件浏览器
> ＋ 在面板里就地预览**几乎所有常见文档**：Office 全家桶、PDF、Markdown、HTML、图片、代码与文本。
>
> **纯 JavaScript 渲染，零常驻服务，不依赖 LibreOffice 或任何外部转换器，不改桌面端一行源码。**

它是「不改宿主源码、纯注入扩展」的写法：装上即在渲染进程里工作，**停用/删除插件后界面 100% 还原**。

---

## 能力一览

| 格式 | 用什么渲染 | 备注 |
|---|---|---|
| `docx` `doc` `docm` `xlsx` `xls` `xlsm` `csv` `tsv` `pptx` `pps` `rtf` `odt` `ods` `wps` `et` `ofd` | 随插件携带的**纯 JS 内核**（`@file-viewer/*` 系列构建产物） | 按**文件头签名**分派，不只看扩展名；表格读真实样式与填充色 |
| `ppt` `dps`（97-2003 二进制） | **Flyfish Public Watermarked Runtime**（WASM） | 内核不支持 OLE2 容器，故单独走这条；每页右下角带 `Flyfish Viewer` 水印（**许可要求，不可去除**） |
| `pdf` | **pdf.js** | 逐页 canvas，DPR 高清，连续滚动 |
| `md` `markdown` `mdx` | **marked** | 纯前端解析 |
| `html` `htm` | **沙箱 iframe** | 保留脚本执行，但不给同源 —— 页面碰不到宿主 |
| `png` `jpg` `jpeg` `gif` `webp` `svg` `bmp` `tiff` `heic` `avif` | 内嵌等比显示 | 走桌面桥取字节 |
| `txt` `log` `json` `yml` `py` `js` `sh` `css` `sql` … | 等宽只读 | **多编码嗅探**：UTF-8 / GBK / Big5 / Shift-JIS / EUC-KR / Windows-1252，中文不乱码 |

**功能面**

- **文件浏览器**：默认落 `~/.hermes/workspace`；`↑上一级 / ←后退 / →前进`（带历史栈）、可点击的分段路径栏、
  按格式映射的图标、文件与目录混排、可打开系统访达 / 用系统程序打开 / 复制路径。
- **就地预览**：点文件即在面板里铺满渲染，顶部窄工具条（重新渲染 / 在访达中显示 / 用系统程序打开 / 复制路径 / 关闭）。
- **高保真兜底**：工具条上的「**精确版（PDF）**」—— 同目录已有同名 `.pdf` 就直接切过去用 pdf.js 看；
  没有则给出命令，跑一次 `off2pdf.py` 后即可一键打开（矢量、字体、分页 100% 保真）。
- **自适应面板**：用 `ResizeObserver` 跟随面板宽度（拖窄拖宽都不变形）；老 `.ppt` 按面积自适应宽度。

---

## 环境要求

| 必须 | 说明 |
|---|---|
| **Hermes 桌面端**（带桌面插件机制／`@hermes/plugin-sdk`） | 本插件是它的桌面半插件 |
| 无其他依赖 | **不需要** Node.js、**不需要** Python 包、**不需要** LibreOffice |

可选：LibreOffice（`soffice`）—— 只有「精确版（PDF）」与「老 `.ppt` 转 `.pptx`」这两个**按需**功能用得到
（`brew install --cask libreoffice`）。不装完全不影响日常预览。

---

## 安装

```bash
git clone https://github.com/xionglaoshi/hermes-office-viewer.git
cd hermes-office-viewer
bash install.sh            # 先看会做什么：bash install.sh --dry-run
```

`install.sh` 做三件事：把插件复制到 `~/.hermes/plugins/office-viewer/`（既有同名目录会先备份成 `.bak-<时间戳>`）、
用 hermes CLI 启用它、把可选的 `scripts/off2pdf.py` 放到 `~/.hermes/scripts/`。

然后在应用里收尾（**必须**，插件的桌面半是强制 opt-in 的）：

1. `⌘K` →「技能与工具」→「桌面插件」→ **重新扫描**
2. 打开「**Office 查看器**」开关
3. `⌘R`（或重启桌面端）；`⌘J` 呼出右侧栏，切到「**文档预览**」标签

### 手动安装（不用脚本）

```bash
mkdir -p ~/.hermes/plugins
cp -R hermes-office-viewer ~/.hermes/plugins/office-viewer
~/.hermes/hermes-agent/venv/bin/hermes plugins enable office-viewer   # 或改 config.yaml 的 plugins.enabled
```

> ⚠️ **不要**手工往 `~/.hermes/desktop-plugins/` 里放副本 —— 那是桌面端自己管理的目录，
> 无 `.hermes-package.json` 的手工副本会被当成「用户有意独立安装」，之后永不复写、列表卡在「复制中…」。

### 更新

```bash
cd hermes-office-viewer && git pull
bash install.sh          # 再 ⌘K →「重新扫描」
```

---

## 目录结构

```
office-viewer/
├── plugin.yaml                  # 插件清单（agent 半）
├── desktop/
│   ├── plugin.js                # 桌面半：唯一入口（贡献面 / 面板 / 命令 / 全部渲染逻辑）
│   └── assets/office-viewer/    # 随插件携带的渲染资产（约 22 MB，缺一不可）
│       ├── office-viewer.js     # 纯 JS 内核（Office 系）
│       ├── pptx.worker.js       # PPTX Worker
│       ├── pdf.min.mjs / pdf.worker.min.mjs   # pdf.js
│       ├── marked.umd.js        # Markdown
│       ├── ppt/                 # 老 .ppt/.dps 引擎（WASM + 中文字体）
│       └── SOURCE.txt           # 资产来源与 sha256
├── dashboard/plugin_api.py      # 后端兜底（桥不可用时的列目录 / 取数据 / 打开）
├── scripts/off2pdf.py           # 可选的「精确版 PDF」助手（一次性脚本，不驻留）
├── install.sh
├── LICENSE                      # MIT（只覆盖作者代码）
└── THIRD-PARTY-NOTICES.md       # 第三方组件与各自许可（**发布/商用前请读**）
```

---

## 实现要点（想移植到你自己的插件可以抄这些）

1. **不改宿主源码，全部走注入**：面板用插件贡献面（`placement:'right'` ＋ `dock` 到既有面板、
   `uncloseable:true`），界面改造走 DOM/CSS 注入 ⇒ 停用即还原。
2. **文件字节走桌面桥**：`window.hermesDesktop.readDir / readFileDataUrl`（**桥优先、插件后端兜底**），
   不新起服务、不落临时文件。桥的单文件读取上限默认 256 MB，超出时把按钮做成「调高上限到 256 MB 并重试」，
   **不静默改用户设置**。
3. **引擎与 Worker 用 blob 运行时注入**：插件文件本身是被当 blob 模块加载的，
   相对路径 import 与 `new URL('./x.wasm', import.meta.url)` 都会炸 ⇒ 从磁盘读字节、造 blob URL，
   再把源码文本里的相对说明符改写成 blob 地址后注入。
4. **跨 realm 传字节要用 iframe 自己的构造器**：`new iframe.contentWindow.Uint8Array(bytes)`，
   否则子文档里 `instanceof Uint8Array` 判假（旧 `.doc` 引擎会直接报 `Unsupported input type`）。
5. **渲染容器必须有确定高度**（`height:100%`，不是 `min-height`）—— 按视口加载的表格引擎在高度为 0 的容器里
   会只出页签、正文空白，**而且不报错**。
6. **自适应用 `ResizeObserver` 观察渲染宿主**，不要只监听 `window.resize`（拖面板宽度不会触发 window resize）；
   缩放用 `zoom` 而不是 `transform: scale`（后者只缩画面、容器仍占原高）。
7. **文本自己解码**，别用 `fetch(dataUrl).text()`（那是硬按 UTF-8 解，中文 GBK 必乱码）。
8. **每个新能力配独立开关**（⌘K 命令 ＋ 自己的 storage 键），出问题能一键回到宿主原生行为。

---

## 已知边界（诚实清单）

- **老 `.ppt` / `.dps` 带水印**（每页右下角 `Flyfish Viewer`）：这是该引擎的许可要求，由 WASM 画进最终帧，
  **不要**试图用 CSS 或叠白块去除。要无水印只有一条路 —— 把 `.ppt` 另存/转成 `.pptx`（本仓库的 `pptx` 路径无水印）。
- **xlsx**：图表不还原、公式只显示缓存值不重算；依赖外部链接的工作簿可能缺数据。
- **pptx**：不播放动画与转场。
- **超大文件**：受桌面桥读取上限限制（默认 256 MB，可在「设置 → 聊天」改，范围 1 MB–4 GB）；插件会明确提示而不是静默失败。
- **老 `.ppt` 的自适应**在个别面板宽度下可能仍不铺满（引擎内部固定尺寸容器所致）—— 换 `.pptx` 是最快的解法。
- 面板里的「隐藏项 / 系统对话框选文件」等入口按窄口径设计；不提供「粘贴绝对路径」这类输入框（刻意为之）。

---

## 许可

- **作者代码**：**MIT**（见 `LICENSE`）。
- **第三方组件**：渲染内核（`@file-viewer/*` 构建产物）、pdf.js（Apache-2.0）、marked（MIT）、
  老 `.ppt` 引擎（Flyfish Public Watermarked Runtime，**非开源**）各有自己的条款。
  **再分发 / 商用前请先读 [`THIRD-PARTY-NOTICES.md`](THIRD-PARTY-NOTICES.md)**，
  其中老 `.ppt` 引擎必须随附其 `LICENSE`/`NOTICE`、保留水印、不得做成独立替代品。

## 致谢

- [Hermes](https://hermes-agent.nousresearch.com/docs) —— 桌面端与插件 SDK
- [pdf.js](https://github.com/mozilla/pdf.js) · [marked](https://github.com/markedjs/marked) · [@file-viewer](https://github.com/file-viewer)
- 设计取舍受了 ChatGPT/Codex 桌面端文档渲染的启发（其 Office 内核为闭源 WASM，故本项目走纯 JS 路线）

## 相关项目

- [hermes-desktop-beautify](https://github.com/xionglaoshi/hermes-desktop-beautify) —— Hermes 桌面端美化插件；
  它把本插件作为「文档预览」面板与「会话卡片点击 → 文档预览」的运行时依赖。
- [meeting-recorder](https://github.com/xionglaoshi/meeting-recorder) —— 会议记录工具（录音→实时转写→纪要）。
