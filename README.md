# ibd-doc-review

投行文档的格式活儿，都归这个 Skill 管。

两件事：

1. **套样式**。招股书、反馈回复、报告、备忘录、尽调报告，给你排成投行的标准格式。样式由模板决定——想改格式，改模板就行。
2. **查问题**。把文档过一遍，看序号连不连续、日期写法统不统一、金额有没有加千分位、标点全半角对不对、简称有没有乱用、表格字号合不合规……查完给一份问题清单，标好位置和改法。

它只动格式，不动你的内容。

## 怎么用

装了就能用。两个脚本，都是 Python 自带库，不用装任何东西。

**套样式**：

```bash
python scripts/check_styles.py --input 你的文档.docx --scenario 招股书     # 套招股书样式
python scripts/check_styles.py --input 你的文档.docx --scenario 反馈回复   # 套反馈回复样式
python scripts/check_styles.py --input 套完的.docx --verify-content 原文.docx  # 确认没动内容
python scripts/check_styles.py --input 修正稿.docx --diff 原文.docx       # 改了哪些格式，出一份清单
python scripts/check_styles.py --input 套完的.docx --revise 原文件.docx    # 用 Word 修订模式标出格式改动，可接受/拒绝
```

**查问题**（只读，不改文件）：

```bash
python scripts/check_content.py --input 你的文档.docx                    # 全部查一遍
python scripts/check_content.py --input 你的文档.docx --checks text      # 只查文字类
python scripts/check_content.py --input 你的文档.docx --checks calc,geo  # 只查计算和敏感词
```

查完在文档同目录生成一份 `_格式核对报告.md`：总览一张表，然后按「必须改 / 建议改 / 自己看着办」三档列出每一条问题，带位置、原文、问题、改法。表格里的问题会标出表格首行内容，方便你找到是哪张表。

查什么，简单说：

- **文字**：标题序号（跳号、重号、倒退都算）、错别字（如「帐」应写「账」）、日期写法（除了「X年X月X日」最正式，其余写法要统一成一种）、多余空格和重复标点、简称（同名不同义、没定义就用）、国家地区表述（内置一份敏感词清单，可用 `--geo-file` 追加）
- **数据**：金额有没有千分位和两位小数（百分比、文号、标准号、地址、案号这些不算金额，会跳过）、同一个指标前后数值是否一致、表格合计行加总对不对、占比列加起来是不是 100%、跨表同名科目数值是否对得上
- **表格**：字号是不是五号（放不下可用小五，封面提示框和签字页不查）、数字有没有右对齐、空单元格（只给一条汇总）、「不适用」符号是否统一

## 模板在哪改

`assets/templates/` 下三个模板就是样式源：

| 模板 | 管什么 |
|------|--------|
| `报告模板.docx` | 招股书版样式（正文、各级标题、字体字号、对齐缩进） |
| `反馈回复样式.docx` | 反馈回复版样式（监管问题黑体、回复宋体） |
| `表格模板.docx` | 三线表规范（框线、表头、对齐） |

用 Word 打开模板改样式，保存，输出自动跟着变。样式 ID 和段落的对应关系见 `references/style-map.md`。

> skillhub 上装的版本不带模板（平台不收 docx 文件），模板从 GitHub 仓库拿：`git clone https://github.com/Kianchales/ibd-doc-review`，或用仓库里的 `assets/templates/` 三个文件。

## 目录

```
ibd-doc-review/
├── SKILL.md                    # 触发词、场景识别、流程、铁律
├── README.md                   # 本文件
├── assets/templates/           # 样式源模板（可自己改）
├── references/
│   ├── style-map.md            # 两套样式体系的样式对照表
│   ├── rules.md                # 序号规则、表格规范、排版铁律
│   ├── sensitive_terms.json    # 国家/地区敏感词清单
│   └── examples.md             # 使用例子
└── scripts/
    ├── check_styles.py         # 套样式 + 样式校验
    ├── check_content.py        # 内容级格式核对（只读）
    └── tests/                  # 自测用例
```

## 几条规矩

- 只改格式，不改内容。套完样式必须用 `--verify-content` 验一遍，逐字对不上就 Fail，不算完。
- 能套命名样式（000-009 这套）就别手写字体字号，不然改起来麻烦。
- 段落之间不要空行，间距由样式控制。
- 反馈回复里，监管问题必须黑体、回复正文必须宋体，这是审核的要求。

## 依赖

- 必需：无。
- 可选：文档生成/编辑工具（新建、套模板、局部微调时用，SKILL.md 里有说明）。

## 许可

MIT。
