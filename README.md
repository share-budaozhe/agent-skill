# agent-skill

智能体工作流 skill 分享。

本仓库用于存放和分享可在 Claude Code 中直接使用的 skills。

## 目录结构

```
agent-skill/
├── pdf-to-markdown/          # PDF 批量转 Markdown 的 skill
│   ├── SKILL.md              # skill 定义与使用说明
│   └── scripts/
│       └── extract_text.py   # 核心转换脚本
├── image-to-visio-redraw/    # 图片转 Visio 矢量重绘的 skill
│   ├── SKILL.md              # skill 定义与使用说明
│   ├── references/
│   │   └── 01-看图拆解方法论.md
│   └── scripts/
│       ├── build_visio.ps1   # 生成 vsdx
│       ├── export_png.ps1    # 导出 PNG 预览
│       └── template_gen_data.py
└── README.md                 # 本文件
```

## 已有 skills

| 名称 | 说明 |
|------|------|
| [pdf-to-markdown](./pdf-to-markdown/SKILL.md) | 将 PDF 批量转换为 Markdown，自动检测文本层，扫描版自动跳过并提示 |
| [image-to-visio-redraw](./image-to-visio-redraw/SKILL.md) | 把图片（截图/照片/PPT 导出图）里的架构图按内容重绘为原生可编辑的 Visio 矢量图（.vsdx），不贴图、逐元素可编辑 |

---

## 安装说明

### 1. 克隆仓库

```bash
git clone https://github.com/share-budaozhe/agent-skill.git
```

### 2. 安装 skill 到 Claude Code

Skills 需要放到 Claude Code 的 `.claude/skills/` 目录下，有两种安装位置：

**方式 A：安装到用户级（所有项目可用，推荐）**

```bash
# 在 agent-skill 目录下执行
mkdir -p ~/.claude/skills
cp -r pdf-to-markdown ~/.claude/skills/
```

**方式 B：安装到某个项目（仅该项目可用）**

```bash
cp -r pdf-to-markdown /path/to/your/project/.claude/skills/
```

安装后目录结构应为：

```
~/.claude/skills/
└── pdf-to-markdown/
    ├── SKILL.md
    └── scripts/
        └── extract_text.py
```

### 3. 安装依赖

`pdf-to-markdown` 依赖 Python 和 `pypdfium2`：

```bash
pip install pypdfium2
```

`image-to-visio-redraw` 依赖：

- **Microsoft Visio**（Office16，需注册 COM `Visio.Application`）
- **Windows PowerShell 5.1+**
- **Python 3**（生成中间 JSON，标准库即可）

> 该 skill 为 Windows 专属，运行时按 SKILL.md 中的工作流调用 `scripts/build_visio.ps1` 生成 `.vsdx`。

### 4. 使用

重启 Claude Code（或重新加载 skills 列表），然后直接描述需求即可触发，例如：

> 把这个目录下的 PDF 转成 md

skill 会自动完成 PDF 文本层检测、批量转换，并将扫描版 PDF 单独列出提示。

更详细的使用说明见 [pdf-to-markdown/SKILL.md](./pdf-to-markdown/SKILL.md)。
