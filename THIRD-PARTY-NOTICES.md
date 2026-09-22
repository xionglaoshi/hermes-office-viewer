# 第三方组件与许可 / Third-Party Notices

本仓库**作者的代码**以 MIT 发布（见 `LICENSE`）。但仓库里随包携带了若干**第三方组件**，
它们**不适用** MIT，各自条款如下。分发本仓库或基于它做产品前，请阅读本节。

## 1. 预构建渲染内核（Office 系）

| 文件 | 说明 |
|---|---|
| `desktop/assets/office-viewer/office-viewer.js`（约 2.0 MB） | `@file-viewer/*` 系列（renderer-word / renderer-spreadsheet / pptx / doc / renderer-ofd）的预构建产物，覆盖 docx / doc / xlsx / xls / wps / et / dps / rtf / ofd / csv |
| `desktop/assets/office-viewer/pptx.worker.js`（约 567 KB） | 同上系列的 PPTX Worker |

- 该内核是**构建产物而非源码**，内含上游的组件清单与许可文本（OFL 字体、MIT、MPL-2.0 等）。
- 其中 docx 渲染路径带有上游的授权校验调用（`assertViewerLicense`）。
- **本仓库作者声明持有随本仓库分发该内核的权利**；若你要**再分发、商用或做成对外服务**，
  请自行与上游确认你的授权范围。上游来源线索见 `desktop/assets/office-viewer/SOURCE.txt`。

## 2. pdf.js

- `desktop/assets/office-viewer/pdf.min.mjs`、`pdf.worker.min.mjs`
- 许可：**Apache License 2.0** —— 全文见 `LICENSES/pdfjs-LICENSE.txt`
- 上游：https://github.com/mozilla/pdf.js

## 3. marked

- `desktop/assets/office-viewer/marked.umd.js`（v16.4.2）
- 许可：**MIT** —— 全文见 `LICENSES/marked-LICENSE.md`
- 上游：https://github.com/markedjs/marked

## 4. 老 .ppt / .dps 引擎（Flyfish Public Watermarked Runtime）

- 目录：`desktop/assets/office-viewer/ppt/`（`index.mjs` `worker.mjs` `frame-cache.mjs`
  `ppt-native.wasm` `ppt-font-cjk.otf` `manifest.json`）
- 许可：`@file-viewer/ppt` 的 **Flyfish Public Watermarked Runtime License v2**
  —— **不是开源许可**（明确声明 "not an open-source license"）。
- 条款要点（原文见 `desktop/assets/office-viewer/ppt/LICENSE` 与 `NOTICE`，**必须随附、不得删改**）：

  1. 允许个人与组织用于个人、内部、商业、生产、SaaS、托管与面向客户的应用；
  2. **允许**打包、复制、托管、镜像与再分发 —— 作为**集成依赖**（含 npm、CDN/IIFE、Docker、
     离线、copy-assets 与 **GitHub Release** 分发）；
  3. **必须保留水印**（每页右下角 `Flyfish Viewer`，由 WebAssembly 绘制在最终帧上）；
     不得去除、遮挡或绕过水印；
  4. **不得**把它作为 `@file-viewer/ppt` 的独立替代品对外提供（即不能做成"另一个 PPT 预览 SDK"）。
- 上游：https://github.com/file-viewer

## 5. 字体与其他内嵌组件

内核内嵌字体（Carlito / Cousine / NotoSansSymbols2 / Tinos 等）为 **SIL OFL 1.1**；
其余内嵌库（base64-js、pako、fzstd、ag-psd 等）为 **MIT**，`introspect-wasm` 为 **MPL-2.0**。
这些组件的许可文本已随构建产物一同内嵌在 `office-viewer.js` 里（`LICENSE.*.txt`）。

---

**一句话**：作者代码 MIT；内核与老 .ppt 引擎是第三方构建产物，各有条款 ——
日常自用、公司内部用、按上面第 4 节的条件再分发都没问题；要做成对外产品请先核对你的授权。
