<div align="center">

<img src="./assets/hero-banner.svg" alt="academic-report-docx Hero Banner" width="100%"/>

<br/>

[![Claude Code](https://img.shields.io/badge/Claude_Code-Skill_Ready-0f172a?style=flat-square&logo=anthropic&logoColor=white)](#五五大智能体安装与调用指南)
[![OpenAI Codex](https://img.shields.io/badge/OpenAI_Codex-openai.yaml-10a37f?style=flat-square&logo=openai&logoColor=white)](#五五大智能体安装与调用指南)
[![dsh DeepSeek Harness](https://img.shields.io/badge/dsh-DeepSeek_Harness-2563eb?style=flat-square)](#五五大智能体安装与调用指南)
[![ZCode](https://img.shields.io/badge/ZCode-Skill_Compatible-4f46e5?style=flat-square)](#五五大智能体安装与调用指南)
[![WorkBuddy](https://img.shields.io/badge/WorkBuddy-Docx_Skill-059669?style=flat-square)](#五五大智能体安装与调用指南)
[![License: MIT](https://img.shields.io/badge/License-MIT-slate?style=flat-square)](./LICENSE)

**把“客观平实的社科工作汇报文风”与“宋体 + Times New Roman 纯黑三段式三线表 Word 规范”凝练为一套跨平台 Agent 技能。**

[设计初衷](#一为什么需要这个技能) · [行文规范](#二行文风格规范客观平实的学术汇报体) · [版式解剖](#三word-版式规范gbt-77131-三段式三线表) · [架构流水线](#四三阶段闭环编译流水线) · [五大终端安装](#五五大智能体安装与调用指南) · [命令行速查](#六独立命令行用法cli-quickstart)

</div>

---

## 一、为什么需要这个技能？

用大模型生成发给合作者、课题组负责人（PI）或审稿团队的**阶段性研究汇报 Word 文档**时，通常会在“文风”和“排版”两端同时踩坑：

1. **文风摇摆失准**：要么写成满篇“我把……整理好了，你看合不合适”的**第一人称口语聊天腔**，缺乏科研文档应有的客观分寸；要么滑向另一个极端，堆砌“赋能、双轨奏疏、藩臬铨选、跨库实体链接、定稿宽表”这类**晦涩古语与 AI 句法压缩饼干**，还在中文句子里每隔几个字塞一个英文缩写括号，让人读三行就发紧。
2. **Word 排版粗糙失范**：Pandoc 或常规脚本直出的 `.docx` 往往带着默认的蓝色标题、等线/Calibri 字体、带竖线和灰底的网格表、无首行缩进的松散段落，不符合中文社科与自然科学期刊通用的**“中文宋体 + 西文数字 Times New Roman + 纯黑配色 + 标准三段式三线表”**学术规范。

`academic-report-docx` 将经过真实大型实证项目打磨定型的**汇报行文准则**与**底层 OpenXML 精排引擎**打包为开箱即用的跨终端技能，原生适配 **Claude Code、Codex、dsh（DeepSeek Harness）、ZCode 与 WorkBuddy**。

---

## 二、行文风格规范：客观平实的学术汇报体

<p align="center">
  <img src="./assets/prose-style-comparison.svg" alt="行文风格三栏对照图" width="100%"/>
</p>

本技能内置自动文风门禁脚本 [`scripts/audit_report_prose.py`](./scripts/audit_report_prose.py)，在编译 Word 前逐行执行 5 项硬核检查：

| 检查维度 | 强制标准 | 自动拦截反例 | 标准替代表述 |
| :--- | :--- | :--- | :--- |
| **1. 汇报人称视角** | **零第一 / 第二人称代词** | `我`、`我们`、`咱们`、`你`、`您` | `本报告汇总……`、`清洗程序按……`、`需确认……` |
| **2. 拒绝晦涩古词** | 用现代社科直白准确术语 | `奏疏 / 藩臬 / 铨选 / 奏销 / 抑或 / 确系` | `奏折与题本 / 省内下属官员 / 人事任免 / 财务报销 / 还是 / 确实` |
| **3. 拒绝 AI 套话俗喻** | 拆解复合名词砖块与口语俗喻 | `跨库实体链接`、`定稿宽表`、`毛坯 / 精装修 / 对账` | 直接写明具体处理动作、输入输出与核对规则 |
| **4. 正文去括号路障** | 叙述句内禁塞纯英文缩写括号 | *“条件随机缺失（MAR）”*、*“假说一（H1）”* | 正文纯中文流畅叙事；斜体变量名 `*var*` 统一归入三线表 |
| **5. 均匀信息密度** | 单分句 **≤ 38 汉字**，自然序数分段 | 50 字不换气的长定语分句、反引号 `` ` ``、`【】` | 短句衔接；枚举用段首“第一，…… 第二，……”保护首行缩进 |

---

## 三、Word 版式规范：GB/T 7713.1 三段式三线表

<p align="center">
  <img src="./assets/three-line-table-anatomy.svg" alt="学术 Word 版式与三段式三线表解剖图" width="100%"/>
</p>

编译脚本 [`scripts/md_to_academic_docx.py`](./scripts/md_to_academic_docx.py) 直接重写 `.docx` 压缩包内的 `word/styles.xml`、`word/document.xml` 与 `word/footer1.xml`，在 OpenXML 层面锁定全部版式细节：

- **中英数双字体分离**：所有中文字符（`w:eastAsia`）强制绑定 **`宋体`（SimSun）**；所有英文字母、阿拉伯数字与西文标点（`w:ascii / w:hAnsi`）强制绑定 **`Times New Roman`**；公式（`<m:oMath>`）绑定 **`Cambria Math`**。
- **全域 100% 纯黑色（`#000000`）**：自动剥离全部蓝色标题、灰色副标题、超链接染色、代码块灰底与单元格底纹。
- **标准 A4 段落层级**：A4 纸张（上下边距 `2.54 cm`、左右边距 `2.80 cm`），页脚居中 `10pt` 自动页码；标题三号加粗居中、一级标题四号加粗居左、二级标题小四加粗居左（均开启 `keepNext` 防孤行）；正文**小四（`12pt`）、1.5 倍行距、首行缩进 2 字符、两端对齐**。
- **标准三段式三线表**：
  1. **第一段（表上表题）**：`表 X ……` 位于表正上方，五号（`10.5pt`）加粗居中，锁定 `keepNext` 防止表题与表体跨页断开；
  2. **第二段（三线表体）**：顶线 **`1.5pt`**（`w:sz="12"`）、表头底线 **`0.75pt`**（`w:sz="6"`）、底线 **`1.5pt`**（`w:sz="12"`），**彻底清除全部竖线与表体内部横线**；首行自动设为跨页重复表头（`w:tblHeader`），5 列及以上宽表自动切换为小五号（`9pt`）防折行，原生支持 `A. ...` / `B. ...` 子面板加粗分组行；
  3. **第三段（表下表注）**：`注：……` 紧贴表底下方，小五号（`9pt`）两端对齐、无首行缩进，交代聚类标准误、控制变量与显著性水平。

---

## 四、三阶段闭环编译流水线

<p align="center">
  <img src="./assets/multi-agent-workflow.svg" alt="五大智能体适配与三阶段流水线" width="100%"/>
</p>

---

## 五、五大智能体安装与调用指南

本仓库同时采用**根目录直装结构**（`SKILL.md` + `agents/openai.yaml`）与**目录索引镜像结构**（`skills/academic-report-docx/`），兼容一键全平台脚本安装、单终端克隆安装以及插件市场安装。

### 方式一：一键全平台自动安装（推荐）

克隆仓库后运行内置安装脚本，即可一次性部署到本机 **Claude Code、Codex、dsh、ZCode、WorkBuddy** 全部技能目录：

**Windows（PowerShell）：**
```powershell
git clone https://github.com/icgma/academic-report-docx.git
cd academic-report-docx
.\install.ps1
# 如只需安装到指定终端：
# .\install.ps1 -Target claude,codex,dsh,zcode,workbuddy
```

**macOS / Linux（Bash）：**
```bash
git clone https://github.com/icgma/academic-report-docx.git
cd academic-report-docx
chmod +x install.sh && ./install.sh all
```

---

### 方式二：按各智能体终端单独安装与调用

| 智能体终端 | 全局技能目录（用户级） | 项目级目录（团队共享） | 专属适配特性与触发方式 |
| :--- | :--- | :--- | :--- |
| **Claude Code** | `~/.claude/skills/academic-report-docx` | `<repo>/.claude/skills/academic-report-docx` | 支持直接克隆，或通过 `/skill-install https://github.com/icgma/academic-report-docx` 扫描 `skills/` 目录自动安装 |
| **OpenAI Codex** | `~/.codex/skills/academic-report-docx`<br/>`~/.agents/skills/academic-report-docx` | `<repo>/.agents/skills/academic-report-docx` | 内置 [`agents/openai.yaml`](./agents/openai.yaml) 与矢量/位图双尺寸图标，支持 `$academic-report-docx` 显式唤起 |
| **dsh (DeepSeek Harness)** | `~/.dsh/skills/academic-report-docx` | `<repo>/.dsh/skills/academic-report-docx`<br/>`<repo>/.agents/skills/academic-report-docx` | 严格遵循 `@deepseek-ai/dsh-skill-filesystem` 规范（`kebab-case` 名称 + `whenToUse` + Chokidar 免重启热重载） |
| **ZCode** | `~/.zcode/skills/academic-report-docx` | `<repo>/.zcode/skills/academic-report-docx` | 原生识别 `SKILL.md` 触发词，在实证分析与协作者审阅汇报阶段自动挂载 |
| **WorkBuddy** | `~/.workbuddy/skills/academic-report-docx` | `<repo>/.workbuddy/skills/academic-report-docx` | 原生提供 `display_name_zh`、`display_name_en` 与 `description_en` 双语技能面板元数据 |

<details>
<summary><b>展开查看各终端独立单行克隆命令</b></summary>

```bash
# 1. Claude Code
git clone https://github.com/icgma/academic-report-docx.git ~/.claude/skills/academic-report-docx

# 2. OpenAI Codex
git clone https://github.com/icgma/academic-report-docx.git ~/.codex/skills/academic-report-docx

# 3. dsh (DeepSeek Harness — 写入后 Chokidar 自动热加载，无需重启会话)
git clone https://github.com/icgma/academic-report-docx.git ~/.dsh/skills/academic-report-docx

# 4. ZCode
git clone https://github.com/icgma/academic-report-docx.git ~/.zcode/skills/academic-report-docx

# 5. WorkBuddy
git clone https://github.com/icgma/academic-report-docx.git ~/.workbuddy/skills/academic-report-docx
```

</details>

### 对话触发提示词示例

在任一智能体中输入类似指令即可自动激活本技能：

> “请调用 `academic-report-docx` 技能，把当前项目的数据清洗、缺失值检验和回归结果整理成一份发给合作者审阅的工作汇报 Word 文档（使用客观汇报风格、不要用第一人称，宋体 + Times New Roman、纯黑、三段式三线表）。”

---

## 六、独立命令行用法（CLI Quickstart）

除在智能体对话中自动调用外，你也可以直接在终端使用本技能自带的两个 Python 工具脚本。

### 环境依赖
- **Python 3.9+**，安装 `lxml`：`pip install lxml`
- **Pandoc（推荐）**：用于将 `$...$` 与 `$$...$$` 编译为原生 Word OMML 数学公式（若当前机器未安装 Pandoc，脚本会自动回退至内置纯 Python OpenXML 渲染器，同样保证 100% 通过字体、颜色与三线表校验）。

### 1. 运行汇报文风与表格三段式检查
```bash
python scripts/audit_report_prose.py examples/sample_empirical_report.md
```
输出示例：
```text
=== Academic Report Prose Audit: sample_empirical_report.md ===
Result: PASS (0 prose, pronoun, clause-length, or 3-part table issues detected!)
```

### 2. 编译标准学术 Word 文档并执行 OpenXML 断言验证
```bash
python scripts/md_to_academic_docx.py examples/sample_empirical_report.md \
  -o examples/sample_empirical_report.docx \
  --verify
```
输出示例：
```text
[OK] Compiled Academic Word DOCX via Pandoc (pandoc.exe) + OpenXML Post-Processor: examples/sample_empirical_report.docx
[PASS] OpenXML Verification Gate passed: 2 Three-Line Table(s), 2,697 chars, 100% 宋体 + Times New Roman, pure black #000000.
```

> 仓库内已附带一份编译并核验通过的实证汇报样稿，可直接下载在 Word / WPS 中打开查看排版效果：
> - Markdown 源稿：[`examples/sample_empirical_report.md`](./examples/sample_empirical_report.md)
> - 标准 Word 成稿：[`examples/sample_empirical_report.docx`](./examples/sample_empirical_report.docx)

---

## 七、仓库目录结构

```text
academic-report-docx/
├── SKILL.md                             # 核心技能指令文件（五端通用 Frontmatter + 三步标准作业流）
├── agents/
│   └── openai.yaml                      # OpenAI Codex 原生界面配置与默认提示词
├── assets/
│   ├── hero-banner.svg                  # README 顶部概览横幅与 A4 样张预览
│   ├── prose-style-comparison.svg       # 三栏文风正反对比图解
│   ├── three-line-table-anatomy.svg     # 学术版式与 GB/T 7713.1 三段式三线表解剖图
│   ├── multi-agent-workflow.svg         # 五大智能体适配与三阶段闭环架构图
│   └── icon.svg / icon.png              # 技能矢量与位图图标
├── scripts/
│   ├── audit_report_prose.py            # 客观汇报文风与三段式表格自动门禁检查器
│   ├── md_to_academic_docx.py           # Markdown -> 学术规范 Word 编译器与 OpenXML 断言校验器
│   └── sync_skill_mirror.py             # 同步生成 skills/ 子目录镜像以兼容多技能市场扫描器
├── references/
│   ├── prose_style_guide.md             # 详细文风指南与三组高频场景正反改写范例
│   └── openxml_typesetting_spec.md      # OpenXML 底层字号、行距、页边距与三线表参数表
├── examples/
│   ├── sample_empirical_report.md       # 完整实证研究汇报标准模板（含公式与多面板三线表）
│   └── sample_empirical_report.docx     # 预编译通过 XML 核验的 Word 成品样稿
├── skills/
│   └── academic-report-docx/            # 面向 Claude /skill-install 与目录索引安装器的镜像包
├── install.ps1                          # Windows PowerShell 一键多终端安装脚本
├── install.sh                           # macOS / Linux Bash 一键多终端安装脚本
└── LICENSE                              # MIT 开源协议
```

---

## License

Released under the [MIT License](./LICENSE) © 2026 [icgma](https://github.com/icgma).
