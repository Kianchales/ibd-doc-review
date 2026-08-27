# ibd-doc-review

A 股 IBD 投行文档质检 Skill。

对 Word 文档按招股书或审核问询反馈回复的样式规范进行排版：招股书版样式涵盖报告、备忘录、尽调报告等正式文档；反馈回复版样式适用于审核问询回复与落实函回复。

样式定义全部来自模板文件实证，修改模板即可调整输出。仅处理格式，不修改文档内容。

## 功能特性

- **两套样式体系开箱即用**：招股书版（含报告类正式文档）+ 反馈回复版，样式定义全部模板实证
- **模板即样式源**：样式由 `assets/templates/` 下的模板 docx 定义——**修改/替换模板 = 定制输出样式**，无需改代码
- **三场景自动路由**：已有 docx 整套套样式（apply-template）/ 新建 docx（模板基底填充）/ 局部微调（实时编辑）
- **内置校验门禁**：`scripts/check_styles.py` 自动检查必备样式、裸段落、空段落、标题跳级、内容完整性（--verify-content）、序号段落核对（--check-numbering）
- **三件套交付**（2026-08-27）：格式修改输出 **Word 修订稿**（--revise 基于原文件生成：w:pPrChange 格式更改修订 + trackRevisions，审阅可接受/拒绝）+ 格式问题清单（--diff 生成：位置/原文/改成什么）+ 修改统计
- **投行基本格式核对**（2026-08-27，只读不改文件）：`scripts/check_content.py` 按「文字 / 数据 / 表格」三大组别×严重程度（HIGH/MEDIUM/LOW）组织 14 个核对项——标题序号连续性、用词规范、日期统一、多余空格与重复标点、释义简称、国家城市合规（外部清单）、金额千分位两位小数、指标数值一致、合计与占比计算、跨表勾稽、表格字号体系、数字右对齐、空单元格提示、不适用符号统一
- **零依赖校验脚本**：两个脚本均纯 Python 标准库（zipfile/re），无需安装任何包

## 触发词

| 场景 | 触发词示例 |
|------|-----------|
| 招股书样式（含报告/备忘录/尽调等正式文档） | 「按招股书样式」「招股书格式」「招股书排版」「招股书章节样式」「按招股书章节格式写」「把这篇改成招股书格式」「这篇按招股书排一下」「套用 000-009 样式」「套 000-009」「报告模板样式」「按报告模板」「套模板样式」「应用报告样式」 |
| 反馈回复样式 | 「按反馈回复样式」「按问询回复格式」「反馈回复排版」「反馈回复排版规范」「问询函回复格式」「落实函格式」「审核问询回复格式」「按问询函格式排版」 |
| 通用排版要求 | 「把这个 Word 改成 XXX 格式」「排版规范」「按投行规范排版」「字体字号统一」「排版规范一点」「文档格式统一」 |
| 局部样式调整 | 「标题改成黑体」「正文首行缩进」「表格改成三线表」「单位行右对齐」 |

## 安装

1. 将本仓库 `ibd-doc-review/` 目录复制到技能目录：
   - WorkBuddy：`~/.workbuddy/skills/ibd-doc-review/`
   - Claude Code / Codex / Cursor 等支持 skills 的平台：对应平台 skills 目录
2. 或在终端：`git clone <仓库地址> ~/.workbuddy/skills/ibd-doc-review`

> 安装后技能即自动可用，无需额外配置（校验脚本零依赖）。

## 版本

当前版本 **v0.7.0**（frontmatter `version` 为准；版本演进见 `CHANGELOG.md`）。想固定某个版本使用时，checkout 对应 git tag 或下载对应 Release。

## 使用流程（S1-S7）

```
识别场景（招股书版/反馈回复版）→ 加载样式资产（assets/templates/ 下模板）
→ 工具路由（新建=tencent-docx / 已有套样式=minimax-docx apply-template / 局部=实时编辑）
→ 样式映射（内容段落→pStyle：0011/001 监管问题黑体、000 正文、002-007 标题、008/009 表格辅助）
→ 序号九级链 + 表格三线表
→ 校验门禁（check_styles.py + apply-template XSD 校验）
→ 交付声明
```

校验示例：

```bash
python scripts/check_styles.py --input 你的文档.docx --scenario 反馈回复        # 成品文档校验
python scripts/check_styles.py --input 你的模板.docx --scenario 反馈回复 --mode template  # 样式库校验
python scripts/check_styles.py --input 套样式后.docx --verify-content 原文.docx  # 内容完整性（严禁改原文）
python scripts/check_styles.py --input 你的文档.docx --check-numbering          # 序号段落核对
python scripts/check_styles.py --input 修正稿.docx --diff 原文.docx            # 格式修改清单+统计
python scripts/check_styles.py --input 样式化结果.docx --revise 原文件.docx     # 生成 Word 修订稿（原文件+格式更改修订）
python scripts/check_content.py --input 你的文档.docx                          # 投行格式核对（全部，只读不改文件）
python scripts/check_content.py --input 你的文档.docx --checks text,data       # 只查文字类+数据类
python scripts/check_content.py --input 你的文档.docx --checks calc,geo        # 只查指定子项
```

## 模板自定义（改模板 = 改输出样式）

`assets/templates/` 下的模板 docx 是**样式源**：

| 模板 | 控制什么 |
|------|---------|
| `报告模板.docx` | 招股书版 000-009 样式（正文/各级标题/单位/备注的字体、字号、对齐、缩进、间距） |
| `反馈回复样式.docx` | 反馈回复版 0011/001/000-009 样式（含监管问题黑体、回复宋体对照） |
| `表格模板.docx` | 全文档三线表规范（框线、字号、表头、对齐） |

**自定义方式**：用 Word 打开对应模板 docx，修改样式（如把 000 正文字号改成小五、把标题颜色改成深蓝），保存后重新运行校验——输出将自动跟随新模板。样式 ID 与文档段落映射见 `references/style-map.md`。

> **skillhub 分发版说明**：skillhub 平台不接受 .docx/.zip 资产，模板 docx 由 GitHub 仓库提供——`git clone https://github.com/Kianchales/ibd-doc-review` 获取完整版，或从仓库下载 `assets/templates/` 下 3 个模板 docx 放入本 skill 同目录。本仓库（GitHub 版）自带完整模板，开箱即用。

## 目录结构

```
ibd-doc-review/
├── SKILL.md                    # 主文件：触发词 + 场景识别 + 工具路由 + S1-S7 流程 + 铁律 + 降级协议
├── README.md                   # 本文件
├── assets/
│   └── templates/              # 样式源模板（可自定义替换）
│       ├── 报告模板.docx       # 招股书版 000-009
│       ├── 反馈回复样式.docx   # 反馈回复版 0011/001/000-009
│       └── 表格模板.docx       # 三线表规范
├── references/
│   ├── style-map.md            # 两套样式体系完整映射表（模板实证）
│   ├── rules.md                # 标题序号九级链 + 表格三线表 + 排版铁律
│   └── examples.md             # 典型用例（触发→执行→产出 全链路）
└── scripts/
    ├── check_styles.py         # 样式校验脚本（纯标准库零依赖）
    ├── check_content.py        # 投行格式核对脚本：三大组别 14 项内容级核对（只读不改文件）
    └── tests/
        └── test_check_styles.py # 校验脚本自测（14 用例）
```

## 样式体系速览

- **正文 000**：宋体 12pt，首行缩进 2 字符，1.5 倍行距，段前/段后 0.5 行
- **一级标题**：招股书版 001（黑体 16pt 居中分页前）/ 反馈回复版 0011（黑体 16pt 两端对齐）
- **监管问题 001**（反馈回复版）：黑体缩进——监管问题黑体、回复宋体对照本身是审核合规展示
- **表格**：三线表 v2（100% 页宽自动调整、纯黑边框上下 1.5pt/内部 0.5pt、两级表头支持跨列/跨行合并、表头居中加粗、数字右对齐、合计行加粗跨列、行高 397 禁跨页拆行、单元格 10.5pt 垂直居中）

## 排版铁律（内置规则，与 SKILL.md 一致）

0. 只改格式、严禁修改原文内容：套样式后须 `--verify-content` 校验，文本逐字不一致即 FAIL
1. 模板先行：优先以模板为基底，保留 styles.xml 样式表
2. pStyle 优先于手写格式：能套命名样式就不硬编码字体字号
3. 序号段落判定（标题 vs 正文）："1、""（1）"开头的段落按**段落长短 + 上下文是否分段**判定——短标题+后段独立展开→标题样式；长句+连续序号→000 正文（双维度算法，主判据=上下文分段）
4. 段落间禁止空行/空段落：间距由样式 spacing 控制（校验脚本已内置检测）
5. 反馈回复合规展示：监管问题必须 001 黑体、回复正文必须 000 宋体
6. 不覆盖用户手动格式：apply-template 前先与用户确认

## 依赖

- **必需**：无（校验脚本纯标准库）
- **可选**：`tencent-docx` / `minimax-docx` / 本地 Office 编辑工具（用于文档生成/套用/微调，SKILL.md 有完整路由说明）

## 许可

MIT License（模板 docx 为样式定义示例，可自由修改使用）。
