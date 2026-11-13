---
name: academic-report-docx
display_name_zh: "学术汇报与三线表 Word 生成"
display_name_en: "Academic Report & Three-Line Table DOCX"
description: >-
  客观平实的中文学术工作汇报写作与规范 Word（.docx）自动排版技能（v1.1.0）。强制执行“零第一/第二人称客观汇报体（支持引语豁免与领域词库切换）、
  去生僻古词与AI工程套话、去句中英文括号路障、单分句不超38字”的行文规范，内置实证数字表文交叉对账门禁，并通过 Pandoc / 纯 Python 原生 OMML
  与 OpenXML 引擎一键生成纯黑配色（#000000）、中文宋体 + 西文数字 Times New Roman、小四 1.5 倍行距首行缩进 2 字符、宽表自动横向分节、
  双层跨列表头三段式三线表及规范插图（表上、图下）的学术级 Word 文档，支持对既有杂乱 .docx 直接重排（--restyle）。
  Triggers on: 学术汇报Word, 汇报风格word, 三段式表格, 三线表word, 宋体Times New Roman, 协作者审阅文档,
  研究进展汇报docx, 学术规范排版, word格式重排, academic-report-docx.
description_en: >-
  Draft plain-spoken, objective Chinese academic/research progress reports (context-aware zero 1st/2nd-person
  pronoun gate with quote exemption, empirical number cross-verification) and compile Markdown or restyle
  existing .docx files into strictly formatted academic Word (.docx) documents with SimSun + Times New Roman,
  pure black (#000000), A4 portrait/auto-landscape sections, native OMML math (zero-Pandoc ready), normalized
  figures, and GB/T 7713.1 Three-Line Tables with 2-row spanned headers.
whenToUse: >-
  当用户需要撰写或改写发给合作者、导师（PI）或课题组的阶段性研究工作汇报、数据清洗与实证模型审阅说明，
  或需要将包含统计表格、回归宽表、公式与插图的 Markdown 文稿（或格式杂乱的既有 .docx 文档）导出为符合学术规范
  （宋体 + Times New Roman、纯黑、标准三段式三线表、表题在表上图题在图下）的 Word（.docx）文档时调用本技能。
user-invocable: true
disable-model-invocation: false
version: "1.1.0"
---

# 中文学术工作汇报行文与规范三线表 Word 生成技能（v1.1.0）

本技能将**“平实客观的中文社科汇报文风”**、**“实证数字表文交叉核验”**与**“严格的学术 Word（`.docx`）OpenXML 版式引擎”**封装为端到端闭环工作流。无论是从零起草研究进展说明、改写口语化/AI 腔初稿、核验表文数字一致性，还是将既有杂乱 `.docx` 一键重排为规范三线表文档，均须按以下标准执行并通过脚本校验。

---

## 一、行文风格铁律（语境感知的客观学术汇报体）

### 1. 零第一/第二人称（直接引语与档案摘录自动豁免）
- **叙述语境严禁代词**：汇报正文不得出现 `我`、`我们`、`咱们`、`你`、`您` 等对话体或书信体代词，统一采用客观研究主体（`本报告`、`清洗程序`、`主回归模型`、`稳健性检验`）或无主语动作句式。
- **直接引语智能豁免（v1.1.0）**：置于中文双引号 `“……”` 或引用块 `> ……` 内的受访者原话、问卷题项与原始档案引文自动豁免代词与分句长度检查，避免误伤质性访谈与史料摘录。

### 2. 平实直白，支持可插拔学科词表（`configs/prose_rules.json`）
- **通用核心规则（`--profile core`，默认开启）**：拦截半文言公文套词（`抑或 / 确系 / 予以`）、AI 复合工程黑话（`跨库实体链接 / 定稿宽表 / 赋能 / 抓手`）与非正式口语俗喻（`毛坯 / 精装修 / 对账`）。
- **历史制度扩展词表（`--profile core,historical`）**：在涉及清代/近代官制史料汇报时按需叠加，提示将生僻古官制词（`奏疏 / 藩臬 / 铨选 / 奏销`）替换为现代学术通用表述；也可通过 `--rules custom.json` 加载团队自定义词表。

### 3. 正文去括号路障，严格区分“表题在表上、图题在图下”
- **严禁句中英文缩写括号**：中文叙述句内不夹带纯英文缩写括号，所有斜体变量名（`*var_name*`）、系数与样本量归入三线表与公式。
- **控制单句负荷**：单个无标点汉字分句不超过 **38 字**；正文不用反引号 `` `code` `` 或方头括号 `【】`。
- **图表题序铁律**：表格必须为 **`表 X ……`（上）+ 表体 + `注：……`（下）**；插图必须为 **`![说明](路径)`（上）+ `图 X ……`（下）**，严禁把图题写在图片上方。

> 正反改写对照详见 [references/prose_style_guide.md](./references/prose_style_guide.md)。

---

## 二、学术 Word（`.docx`）排版硬规范（v1.1.0 全特性）

| 版式维度 | 强制规范参数 | OpenXML 底层映射 |
| :--- | :--- | :--- |
| **双字体分离** | 中文 **宋体（SimSun）**；西文与数字 **Times New Roman** | `w:eastAsia="宋体"` + `w:ascii/hAnsi="Times New Roman"` |
| **原生数学公式** | **Cambria Math** 原生 OMML 公式（无 Pandoc 环境亦原生生成） | `<m:oMath>` / `<m:oMathPara>`（支持分式、根式、巨算子、上下标与重音符） |
| **全域纯黑配色** | **100% 纯黑色（`#000000`）** | 清除全部蓝字标题、超链接下划线、底纹 `w:shd` 与高亮 `w:highlight` |
| **自适应横纵分节** | 正文纵向 A4；**`≥ 7` 列宽表或 `<!-- landscape -->` 自动切换横向 A4 分节页** | 表前闭合纵向 `<w:sectPr>`，表注后闭合横向 `<w:sectPr w:orient="landscape"/>` 并切回纵向 |
| **三段式三线表** | 顶底线 `1.5pt`、栏目线 `0.75pt`、**双层跨列表头子线 `0.5pt`**、零竖线 | 支持 `>` 跨列合并 `<w:gridSpan>`、跨页双行表头重复与回归双行标准误紧凑间距 |
| **规范学术插图** | 自动等比限宽 `≤ 14.5 cm`、居中清零首行缩进、锁定 `keepNext` 防图文分离 | 自动处理 `.png/.jpg/.svg`，紧贴下方五号加粗居中 `图 X` 与小五号图注 |
| **文件占用保护** | 自动检测 Windows Word/WPS `~$` 锁文件与写入权限锁 | 目标文件被打开占用时自动安全回退输出至 `*_updated.docx`，绝不崩溃丢稿 |

### 高级语法示例：双层跨列表头宽表与规范插图

```markdown
<!-- landscape -->
表 2 分区域与分行政层级的多模型稳健性回归结果

| 解释变量 | 东部城市子样本 | > | > | 中西部城市子样本 | > | > | 全样本汇总 |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
|  | (1) 基线 | (2) 加控制 | (3) 工具变量 | (4) 基线 | (5) 加控制 | (6) 工具变量 | (7) 双固定效应 |
| 政务数字化指数（*digital_gov*） | -1.842*** | -1.621*** | -1.705*** | -1.214** | -1.108** | -1.190** | -1.516*** |
|  | (0.291) | (0.275) | (0.312) | (0.388) | (0.364) | (0.402) | (0.264) |
| 观测数 N | 1,200 | 1,200 | 1,200 | 1,800 | 1,800 | 1,800 | 3,000 |

注：括号内为地级市聚类稳健标准误；首行 `>` 自动向右合并为跨列组表头并绘制 `0.5pt` 栏目子线（cmidrule）。

![动态效应轨迹](figures/event_study.png)

图 1 改革前后跨部门响应时长事件研究法动态效应估计

注：竖线代表 95% 置信区间，基期为改革前第 1 期。
```

> 底层 OpenXML 完整参数详见 [references/openxml_typesetting_spec.md](./references/openxml_typesetting_spec.md)。

---

## 三、标准执行工作流（三大自动门禁）

定位本技能目录 `<SKILL_DIR>`（当前 `SKILL.md` 所在文件夹），按以下流程执行：

### Step 1：运行语境感知行文与图表题序检查（Prose Gate）
```bash
python "<SKILL_DIR>/scripts/audit_report_prose.py" "<input.md>"
# 如涉及历史官制题材可叠加历史词库：--profile core,historical
```

### Step 2：运行实证数字表文交叉对账门禁（Number Gate）
```bash
# 内置表文一致性核验（自动拦截正负号反转与高精度系数抄错）：
python "<SKILL_DIR>/scripts/verify_numbers.py" "<input.md>"

# 若有统计脚本导出的 JSON/CSV 结果文件，绑定外部真值源逐数对账：
python "<SKILL_DIR>/scripts/verify_numbers.py" "<input.md>" --data-source "<results.json>"
```

### Step 3：编译或重排标准学术 Word 文档并执行 OpenXML 断言（OpenXML Gate）
```bash
# 场景 A：从 Markdown 编译（自动串联 Prose Gate + Number Gate + OpenXML 验证）
python "<SKILL_DIR>/scripts/md_to_academic_docx.py" "<input.md>" -o "<output.docx>" --verify

# 场景 B：对合作者发来的既有杂乱 .docx 直接一键清洗重排为宋体 + Times New Roman + 纯黑三线表
python "<SKILL_DIR>/scripts/md_to_academic_docx.py" "<messy.docx>" --restyle -o "<clean.docx>" --verify
```
