---
name: academic-report-docx
display_name_zh: "学术汇报与三线表 Word 生成"
display_name_en: "Academic Report & Three-Line Table DOCX"
description: >-
  客观平实的中文学术工作汇报写作与规范 Word（.docx）自动排版技能。强制执行“零第一/第二人称客观汇报体、
  去生僻古词与AI工程套话、去句中英文括号路障、单分句不超38字”的行文规范，并通过 Pandoc OMML 与
  OpenXML 后处理脚本一键生成纯黑配色（#000000）、中文宋体 + 西文数字 Times New Roman、小四 1.5 倍行距
  首行缩进 2 字符及标准 GB/T 7713.1 三段式三线表的学术级 Word 文档。
  Triggers on: 学术汇报Word, 汇报风格word, 三段式表格, 三线表word, 宋体Times New Roman, 协作者审阅文档,
  研究进展汇报docx, 学术规范排版, academic-report-docx.
description_en: >-
  Draft plain-spoken, objective Chinese academic/research progress reports (zero 1st/2nd-person pronouns,
  zero obscure archaisms or AI buzzwords, zero inline English bracket roadblocks) and compile Markdown
  into strictly formatted academic Word (.docx) documents with SimSun + Times New Roman, pure black (#000000),
  A4 layout, 12pt 1.5x 2-char-indent body text, native OMML math, and GB/T 7713.1 Three-Line Tables.
whenToUse: >-
  当用户需要撰写或改写发给合作者、导师（PI）或课题组的阶段性研究工作汇报、数据清洗与实证模型审阅说明，
  或需要将包含统计表格与公式的中文文稿导出为符合学术规范（宋体 + Times New Roman、纯黑、标准三段式三线表）
  的 Word（.docx）文档时调用本技能。
user-invocable: true
disable-model-invocation: false
version: "1.0.0"
---

# 中文学术工作汇报行文与规范三线表 Word 生成技能

本技能将**“平实客观的中文社科汇报文风”**与**“严格的学术 Word（`.docx`）版式规范”**封装为一套端到端闭环工作流。无论是从零起草研究进展说明、把口语化草稿或 AI 腔初稿改写为专业汇报，还是将 Markdown 编译为可直接发送审阅的 Word 文档，均须严格执行以下规范并通过脚本自动核验。

---

## 一、行文风格铁律（客观平实的学术汇报体）

### 1. 零第一/第二人称，保持客观干练的汇报视角
- **严禁人称代词**：全文不得出现 `我`、`我们`、`咱们`、`你`、`您` 等对话体或书信体代词。
- **标准主语与句式**：统一采用客观研究主体（`本报告`、`项目`、`清洗阶段`、`跨库匹配程序`、`主回归模型`、`稳健性检验`）或干净的无主语动作句式：
  - ❌ *“我把最近做完的数据清洗整理在这份文档里了，想请你帮忙把把关……”*
  - ✅ *“本报告汇总近期完成的四部分工作：……文末列出七个涉及制度口径与模型细节的具体问题，供审阅把关并商定定稿方案。”*

### 2. 平实直白，拒绝生僻古语、口语俗喻与 AI 工程黑话
- **不用生僻古语与半文言虚词**：不用晦涩古官制词或公文套词（如 `奏疏 / 藩臬 / 铨选 / 奏销 / 审拟 / 抑或 / 确系 / 予以 / 系……所致`），改用现代学术直白表述（如 `奏折与题本 / 省内下属官员 / 日常人事任免 / 地方财务报销 / 命盗刑案判决 / 还是 / 确实 / 在脚注中说明 / 源于……`）。
- **拆解 AI 复合名词砖块与工程黑话**：不用 `跨库实体链接与空间校准`、`定稿宽表`、`合并流水线`、`机械内生偏误` 等生造复合词组，直接写成“谁对什么数据做了哪步处理、解决了什么问题”。
- **禁用不合学术语境的口语俗喻**：严禁出现 `毛坯 / 精装修 / 对账` 等非正式比喻词。

### 3. 正文去括号路障，精确变量与统计量归入三线表
- **严禁句中英文缩写括号**：中文叙述句内不得条件反射式夹带英文括号（如 ❌ *“条件随机缺失（MAR）”*、*“人名权威库（IHP）”*、*“假说一（H1）”*），正文只用纯粹流畅的中文把逻辑讲透。
- **表文分工明确**：所有英文变量名（斜体 `*var_name*`）、模型编号、回归系数、标准误与样本量，100% 完整保留在**三段式三线表**与独立数学公式（`$$...$$`）中，既消除正文阅读卡顿，又保证实证细节零丢失。
- **控制单句认知负荷**：单个逗号分句不超过 **38 个汉字**；正文不用反引号 `` `code` `` 或方头括号 `【】`；枚举条目优先写成自然段首序数词（`第一，……`、`第二，……`），避免自动编号列表破坏 Word 首行缩进版式。

> 需要对照完整正反改写范例时，请先阅读 [references/prose_style_guide.md](./references/prose_style_guide.md)。

---

## 二、学术 Word（`.docx`）排版硬规范

所有导出的 `.docx` 文档必须通过底层 OpenXML 校验，严禁使用花哨配色、网格线表格或默认等线/Calibri 字体：

| 版式维度 | 强制规范参数 | OpenXML 底层映射 |
| :--- | :--- | :--- |
| **中文字体** | 全文统一 **宋体（SimSun）** | `w:rFonts w:eastAsia="宋体"` |
| **西文与数字** | 全文统一 **Times New Roman** | `w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"` |
| **数学公式** | 原生 Word OMML 公式（**Cambria Math**） | `<m:oMath>` / `<m:oMathPara>`（严禁纯文本乱码或代码底纹） |
| **全篇配色** | **100% 纯黑色**（零蓝标、零灰字、零底纹） | 所有 `w:color` 强制锁定 `000000`，清除 `w:shd` 与 `w:highlight` |
| **页面与页码** | 标准 **A4**，上下 `2.54 cm`、左右 `2.80 cm` | 页脚居中 `10pt` Times New Roman 自动页码 |
| **标题层级** | 标题三号加粗居中；一级四号加粗居左；二级小四加粗居左 | 标题开启 `w:keepNext`，防止标题孤立在页底 |
| **正文段落** | **小四（12pt）**、**1.5 倍行距**、**首行缩进 2 字符** | `w:sz="24"`、`w:line="360"`、`w:firstLineChars="200"`、两端对齐 |
| **三段式三线表** | **表上表题 + 标准三线表体 + 表下表注** | 顶底线 `1.5pt (sz=12)`、栏目线 `0.75pt (sz=6)`、零竖线零内横线 |

### 三段式三线表 Markdown 源文件标准写法

在编写待转换的 `.md` 源稿时，每一张表格必须严格具备**表题、表体、表注**三段连续结构（中间各空一行）：

```markdown
表 1 全样本清洗前后核心指标对比（N = 570,653）

| 数据维度与核心字段 | 清洗前原始状态 | 清洗后定稿状态 | 主要修补来源与规则说明 |
| :--- | :---: | :---: | :--- |
| A. 人物传记与履历字段 |  |  |  |
| 权威传记匹配率（*bio_match*） | 41.89% | 99.52% | 接入权威人物传记库按唯一编号关联 |
| 出生年份有效率（*birth_year*） | 51.87% | 77.87% | 清空公式错误值并回填四位公元生年 |

注：面板 A 统计全量去重档案；括号内斜体为底层数据表对应字段名。*** $p < 0.001$，** $p < 0.01$，* $p < 0.05$。
```

> 完整的 OpenXML 字段规范详见 [references/openxml_typesetting_spec.md](./references/openxml_typesetting_spec.md)。

---

## 三、三步标准执行工作流（必须跑通脚本校验）

定位本技能目录 `<SKILL_DIR>`（即当前 `SKILL.md` 所在文件夹），严格按以下三步执行：

### Step 1：起草或改写符合汇报规范的 Markdown 源稿
1. 盘点全部需要保留的实证数字、回归系数、样本量与文件路径（严禁手写编造或遗漏任何统计量）。
2. 按 YAML 头部（`title`、`subtitle`）+ 中文分级标题（`# 一、……`、`## （一）……`）+ 三段式表格语法组织源稿。

### Step 2：运行行文风格自动诊断脚本（Prose Gate）
```bash
python "<SKILL_DIR>/scripts/audit_report_prose.py" "<input.md>"
```
- 脚本会自动扫描：第一/第二人称代词、禁用生僻古语与 AI 套话、超过 38 字的无标点长分句、正文句中英文括号、反引号与方头括号、以及每张表格是否配齐上方 `表 X` 表题与下方 `注：` 表注。
- 若返回非零退出码，必须根据报错行号逐处修改源稿并重跑，直至输出 `PASS`。

### Step 3：编译并校验标准学术 Word 文档（OpenXML Gate）
```bash
python "<SKILL_DIR>/scripts/md_to_academic_docx.py" "<input.md>" -o "<output.docx>" --verify
```
- 脚本调用 Pandoc（将 `$...$` / `$$...$$` 转为原生 Word OMML 公式；若环境未装 Pandoc 则自动启用内置纯 Python OpenXML 渲染器），随后重写 `styles.xml`、`document.xml` 与 `footer1.xml` 完成宋体 + Times New Roman、纯黑与三段式三线表精排。
- `--verify` 会自动解包生成的 `.docx` 并执行 XML 断言核验（检查字体集合、颜色集合、三线表磅值与表题表注配对），全部断言通过后方可向用户交付文件。
