#!/usr/bin/env bash
# ============================================================================
#  hermes-office-viewer · 安装脚本
#
#  把本仓库安装成一个 Hermes 桌面端插件：
#    ① 复制 <仓库> → ~/.hermes/plugins/office-viewer/（含约 22 MB 渲染资产，缺一不可）
#    ② 用 hermes CLI 启用插件（写 config.yaml 的 plugins.enabled）
#    ③ 把可选的 off2pdf.py 放到 ~/.hermes/scripts/（「精确版（PDF）」按钮用得到）
#
#  用法：
#    bash install.sh              # 真装
#    bash install.sh --dry-run    # 只看会做什么，不写盘
#
#  注意：这会用仓库内容覆盖同名的既有插件目录（默认先备份一份 .bak-<时间戳>）。
# ============================================================================
set -euo pipefail

DRY_RUN=0
[ "${1:-}" = "--dry-run" ] && DRY_RUN=1

SRC_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
HERMES_HOME="${HERMES_HOME:-${HOME}/.hermes}"
PLUGIN_DIR="${HERMES_HOME}/plugins/office-viewer"
SCRIPTS_DIR="${HERMES_HOME}/scripts"
STAMP="$(date +%Y%m%d-%H%M%S)"

say() { printf '%s\n' "$*"; }
run() {
  if [ "${DRY_RUN}" = "1" ]; then
    say "   [dry-run] $*"
  else
    eval "$@"
  fi
}

say "──────── hermes-office-viewer 安装 ────────"
say "源目录：${SRC_DIR}"
say "HERMES_HOME：${HERMES_HOME}"
[ "${DRY_RUN}" = "1" ] && say "模式：dry-run（不写盘）"

# ── 0 前置检查 ─────────────────────────────────────────────────────────────
if [ ! -d "${HERMES_HOME}" ]; then
  say "✗ 找不到 ${HERMES_HOME} —— 请先安装 Hermes 桌面端，或用 HERMES_HOME=... 指定位置"
  exit 1
fi

for f in desktop/plugin.js plugin.yaml \
         desktop/assets/office-viewer/office-viewer.js \
         desktop/assets/office-viewer/pptx.worker.js \
         desktop/assets/office-viewer/pdf.min.mjs \
         desktop/assets/office-viewer/marked.umd.js \
         desktop/assets/office-viewer/ppt/ppt-native.wasm; do
  if [ ! -f "${SRC_DIR}/${f}" ]; then
    say "✗ 仓库不完整，缺 ${f} —— 请确认 clone 完整（或仓库被部分下载）"
    exit 1
  fi
done
say "✓ 仓库文件完整"

# ── 1 既有目录先备份 ───────────────────────────────────────────────────────
if [ -d "${PLUGIN_DIR}" ]; then
  say "→ 已存在插件目录，先备份为 office-viewer.bak-${STAMP}"
  run "mv '${PLUGIN_DIR}' '${PLUGIN_DIR}.bak-${STAMP}'"
fi

# ── 2 复制插件本体 ─────────────────────────────────────────────────────────
say "→ 复制到 ${PLUGIN_DIR}（约 22 MB）"
run "mkdir -p '${PLUGIN_DIR}'"
run "rsync -a --exclude '.git/' --exclude '__pycache__/' --exclude '.DS_Store' \
      '${SRC_DIR}/' '${PLUGIN_DIR}/'"

# ── 3 启用插件 ─────────────────────────────────────────────────────────────
HERMES_BIN=""
for cand in "${HERMES_HOME}/hermes-agent/venv/bin/hermes" "$(command -v hermes 2>/dev/null || true)"; do
  [ -n "${cand}" ] && [ -x "${cand}" ] && HERMES_BIN="${cand}" && break
done

if [ -n "${HERMES_BIN}" ]; then
  say "→ 启用插件：${HERMES_BIN} plugins enable office-viewer"
  run "'${HERMES_BIN}' plugins enable office-viewer" || \
    say "   ⚠ CLI 启用失败（不影响后续步骤）—— 你也可以在应用里开开关"
else
  say "⚠ 没找到 hermes CLI，跳过自动启用 —— 在应用里开开关即可"
fi

# ── 4 可选：off2pdf.py（「精确版（PDF）」用） ──────────────────────────────
if [ -f "${SRC_DIR}/scripts/off2pdf.py" ]; then
  say "→ 安装可选脚本 ${SCRIPTS_DIR}/off2pdf.py"
  run "mkdir -p '${SCRIPTS_DIR}'"
  run "cp -p '${SRC_DIR}/scripts/off2pdf.py' '${SCRIPTS_DIR}/off2pdf.py'"
fi

# ── 5 完成 ─────────────────────────────────────────────────────────────────
say ""
if [ "${DRY_RUN}" = "1" ]; then
  say "dry-run 结束：以上命令都没有真正执行。"
else
  say "✓ 装完了。最后两步在应用里做："
fi
say "   ① ⌘K →「技能与工具」→「桌面插件」→ 重新扫描"
say "   ② 打开「Office 查看器」开关，再 ⌘R（或重启桌面端）"
say "   ③ 之后 ⌘J 打开右侧栏，在「文档预览」标签里浏览与预览"
say ""
say "可选：要「精确版（PDF）」就把 LibreOffice 装上（brew install --cask libreoffice），"
say "      然后在面板里点那颗按钮，按提示跑一次 off2pdf.py。"
