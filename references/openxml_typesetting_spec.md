# OpenXML 学术 Word 版式与三段式三线表底层技术规范（v1.1.0）

本规范定义 `scripts/md_to_academic_docx.py` 与 `scripts/latex_to_omml.py` 对 `.docx` 压缩包内 `word/document.xml`、`word/styles.xml`、`word/_rels/document.xml.rels` 与 `word/footer1.xml` 执行的精确 OpenXML 重写参数。

---

## 1. 纵横混排页面分节与页脚页码（`w:sectPr` & `footer1.xml`）

- **纵向标准页面（Portrait A4）**：`w:pgSz w:w="11906" w:h="16838"`（`21.0 cm × 29.7 cm`），上下边距 `1440 dxa`（`2.54 cm`），左右边距 `1588 dxa`（`2.80 cm`）。
- **宽表自动横向分节页（Landscape A4）**：
  - 触发条件：表格逻辑列数 `≥ 7`（可通过 `--landscape-cols N` 调整，`0` 关闭），或表题上方声明 `<!-- landscape -->`。
  - 分节机制：在宽表表题前一段落 `<w:pPr>` 注入纵向闭合 `<w:sectPr>`，在宽表表注 `<w:pPr>` 注入横向闭合 `<w:sectPr>`（`w:pgSz w:w="16838" w:h="11906" w:orient="landscape"`，四周边距 `1440 dxa`），表后正文自动切回纵向 A4 且零多余空白页。
- **页码域代码**：所有分节统一关联 `footer1.xml`，底部居中 `PAGE` 动态域，字号 `10pt`（`w:sz="20"`），字体 `Times New Roman`，纯黑 `000000`。

---

## 2. 字体分离、全域纯黑与原生 OMML 公式（`w:rPr` & `m:oMath`）

| 文本类型 | `w:eastAsia` | `w:ascii` / `w:hAnsi` / `w:cs` | `w:color` | 底纹/高亮/下划线清理 |
| :--- | :--- | :--- | :--- | :--- |
| 普通文本、标题与超链接 | `宋体` | `Times New Roman` | `000000` | 移除 `w:shd`, `w:highlight`, `w:bdr`, `w:u` |
| OMML 数学公式 (`m:r`) | `宋体` | `Cambria Math` | `000000` | 原生支持 `<m:f>`, `<m:rad>`, `<m:nary>`, `<m:acc>`, `<m:sSubSup>` |

---

## 3. 段落、标题与图表题注字号表（`w:pPr`）

| 段落角色 | 字号（磅 / half-pt） | 对齐方式 (`w:jc`) | 行距 (`w:line`) | 首行缩进 | 分页控制 |
| :--- | :--- | :--- | :--- | :--- | :--- |
| 文档主标题 (`Title`) | 三号 `16pt` (`32`) 加粗 | 居中 `center` | `1.25x` (`300`) | 无 | `w:keepNext` |
| 文档副标题 (`Subtitle`) | `11pt` (`22`) 常规 | 居中 `center` | `1.25x` (`300`) | 无 | `w:keepNext` |
| 一级标题 (`Heading 1`) | 四号 `14pt` (`28`) 加粗 | 左对齐 `left` | `1.5x` (`360`) | 无 | `w:keepNext` |
| 二/三级标题 (`Heading 2/3`) | 小四 `12pt` (`24`) 加粗 | 左对齐 `left` | `1.5x` (`360`) | 无 | `w:keepNext` |
| 正文叙述段 (`Body`) | 小四 `12pt` (`24`) 常规 | 两端对齐 `both` | `1.5x` (`360`) | `2字符` (`firstLineChars="200"`, `480 dxa`) | 自动 |
| 独立数学公式 (`oMathPara`) | 小四 `12pt` (`24`) | 居中 `center` | `1.25x` (`300`) | 无 | 自动 |
| 表格上方表题 (`表 X ...`) | 五号 `10.5pt` (`21`) 加粗 | 居中 `center` | `1.25x` (`300`) | 无 | `w:keepNext` |
| 表格下方表注 (`注：...`) | 小五 `9pt` (`18`) 常规 | 两端对齐 `both` | `1.17x` (`280`) | 无 | 自动 |
| 插图载体段 (`w:drawing`) | 等比限宽 `≤ 14.5 cm` (`5,220,000 EMU`) | 居中 `center` | 单倍 (`240`) | **强制归零** (`firstLine="0"`) | `w:keepNext` |
| 插图下方图题 (`图 X ...`) | 五号 `10.5pt` (`21`) 加粗 | 居中 `center` | `1.25x` (`300`) | 无 | 含图注时 `w:keepNext` |
| 插图下方图注 (`注：...`) | 小五 `9pt` (`18`) 常规 | 两端对齐 `both` | `1.17x` (`280`) | 无 | 自动 |

---

## 4. GB/T 7713.1 标准三段式三线表与双层跨列表头（`w:tbl`）

1. **整体宽度与对齐**：`w:tblW w:w="5000" w:type="pct"`（100% 版心宽度），水平居中。
2. **三线及双层栏目子线磅值（`1/8 pt` 单位）**：
   - **顶线（Top Rule）**：`w:val="single" w:sz="12" w:color="000000"`（`1.5 pt`）
   - **跨列分组子线（Group `cmidrule`）**：双层表头首行跨列分组单元格（`<w:gridSpan>`）底边 `w:val="single" w:sz="4" w:color="000000"`（`0.5 pt`），左侧空白桩列底边 `none`
   - **主栏目线（Header Bottom Rule）**：末行表头底边 `w:val="single" w:sz="6" w:color="000000"`（`0.75 pt`）
   - **底线（Bottom Rule）**：末行底边与表底边 `w:val="single" w:sz="12" w:color="000000"`（`1.5 pt`）
   - **零竖线与零内横线**：`left`、`right`、`insideH`、`insideV` 及数据行上下边框全部强制设为 `w:val="none"`。
3. **跨页保护、双行回归标准误与面板子标题**：
   - 单层或双层表头全部写入 `w:tblHeader`（跨页自动重复完整表头），所有行写入 `w:cantSplit`。
   - 当数据行首列为空且其余列为括号标准误 `(0.291)` 时，自动收紧与上方系数行的行距（`line=220, before=0`）并在系数行开启 `keepNext` 防止系数与标准误跨页拆散。
   - 当某行首列匹配 `^(Panel\s+[A-Z]|[A-Z]\.)\s` 且其余单元格为空时，自动识别为面板分组子标题行并加粗居左显示。
