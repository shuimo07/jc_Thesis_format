# jc_Thesis_format

**成都锦城学院 本科毕业论文（设计）格式规范 + 可复用排版工具库**

这里存放的是**论文格式本身**——页面设置、页眉页脚、字体字号、行距缩进、
标题层级、图表题注、三线表、页码分节、封面与前置部分的结构。
不放任何具体论文内容。

---

## 目录结构

```
docs/format-checklist.md          ★ 格式修改清单：要改的地方全集（逐处可勾选）
docs/notes-and-pitfalls.md        ★ 注意事项与踩坑记录 + 症状对照表
docs/format-spec.md               完整格式规范（含从截图反推参数的换算口径）
reference/pages/p01..p57.png       57 页参考框架图（按论文页序归档）
reference/README.md                参考图逐页索引：每页是什么 + 该对照哪些格式点
reference/source-map.tsv           顺序号 → 原文件名 → 字节数 → SHA256
scripts/jc_thesis_format.py       排版工具库：把规范固化成可调用的 python-docx 原语
scripts/build_template.py         生成空白模板 论文模板.docx（只有格式，无内容）
scripts/measure_reference.py      从参考页面截图反推版面参数
scripts/refresh_and_export.py     用本机 Word 刷新目录域 + 导出 PDF + 栅格化逐页校验
template/论文模板.docx             生成好的空白模板（10 页）
requirements.txt
```

## 从哪开始看

| 你想做什么 | 看这个 |
|---|---|
| 把手上这篇论文的格式改对 | [`docs/format-checklist.md`](docs/format-checklist.md) |
| 想先知道哪里容易踩坑 | [`docs/notes-and-pitfalls.md`](docs/notes-and-pitfalls.md) |
| 想核对某页的标准长什么样 | [`reference/README.md`](reference/README.md) |
| 想直接拿一份格式正确的模板 | [`template/论文模板.docx`](template/论文模板.docx) |
| 想用脚本批量排 | [`docs/format-spec.md`](docs/format-spec.md) + `scripts/` |

## 参考框架图

`reference/pages/` 下是 57 页参考截图，按论文实际页序排列：

```
p01             封面（不编号）
p02 – p07       学术诚信声明 I / 版权使用授权书 II / 摘要 III / Abstract IV / 目录 V,VI
p08 – p55       正文 1 – 48（含图、三线表、插图、结论）
p56             参考文献
p57             致谢
```

**正文页码 = 顺序号 − 7**。逐页说明见 [`reference/README.md`](reference/README.md)。

## 快速开始

```bash
pip install -r requirements.txt

# 1) 生成一份空白模板，照着填内容即可
python scripts/build_template.py -o 论文模板.docx

# 2) 自己写脚本按规范排版
python - <<'PY'
from docx import Document
from scripts.jc_thesis_format import *
doc = Document()
apply_page(doc.sections[0])
add_chapter_body_styles(doc)
add_toc_styles(doc)
build_cover(doc, title='题目', college='计算机与软件学院', major='软件工程',
            name='姓名', sid='学号')
new_section(doc, 'upperRoman', 1)          # 前置部分：罗马数字页码
build_integrity_statement(doc)
build_authorization(doc)
build_cn_abstract(doc, '摘要正文', ('关键词1','关键词2'))
build_en_abstract(doc, 'Abstract text', ('keyword 1','keyword 2'))
toc_heading(doc); toc_field(doc, [(1,'第一章　绪论')])
new_section(doc, 'decimal', 1)             # 正文：阿拉伯数字从 1 起
enable_update_fields(doc)
doc.save('论文.docx')
PY

# 3) 刷新目录页码并导出 PDF 校验
python scripts/refresh_and_export.py 论文.docx --raster
```

## 规范速查

| 项目 | 取值 |
|---|---|
| 页面 | A4，上下 2.5 / 左 3.0 / 右 2.5 cm，版心宽 15.48 cm |
| 页眉 | 「成都锦城学院毕业论文（设计）」宋体五号居中 + 双线 |
| 页码 | 封面无 → 前置部分大写罗马数字 → 正文阿拉伯数字从 1 起 |
| 一级标题 | 黑体小二加粗左对齐，段前 30 磅，固定行距 28 磅，另起页 |
| 二级标题 | 黑体四号加粗，段前 5 磅 |
| 三级标题 | 黑体小四加粗，段前 7 磅 |
| 正文 | 宋体小四，首行缩进 2 字符，两端对齐，**固定行距 22 磅** |
| 题注 | 黑体五号加粗居中，**表题在表上、图题在图下** |
| 表格 | 三线表：顶/底 1.5pt，表头下 0.75pt，无竖线，表内宋体五号 |
| 插图 | 宽度统一 15.0 cm，居中 |
| 目录 | TOC 域，小四，固定行距 20 磅，右对齐点线前导 |

细节与依据见 [`docs/format-spec.md`](docs/format-spec.md)。

## 踩过的坑（改之前先看）

1. **中文字体要显式写 `w:eastAsia`**。只设 `font.name` 中文会回落默认字体。
2. **页脚 PAGE 域必须写全 begin / instrText / separate / 缓存值 / end**，
   只写 instrText 会把 ` PAGE ` 当普通文字打印出来。
3. **行尾空白 Word 不画下划线**，封面填充线右侧要用「右制表位 + underscore 前导符」。
4. **`w:pPr` 和 `w:tblBorders` 的子元素顺序是法定的**，插错位置 Word 报文档损坏。
5. **表头下那条线画在单元格 `w:tcBorders/w:bottom` 上**，不是表格级 `insideH`。
6. **改完一定要渲染校验**，只读 XML 发现不了下划线被吞、表格线错行这类问题。

> 展开版（每个坑的原因、正确做法、以及「出现什么症状去哪一节」的对照表）见
> [`docs/notes-and-pitfalls.md`](docs/notes-and-pitfalls.md)。

## 依赖

- `python-docx`（读写 docx）
- `lxml`、`pillow`
- `pywin32`（仅刷新域/导出 PDF 需要，且需本机装有 Word）
- `pymupdf`（仅 PDF 栅格化校验需要）
