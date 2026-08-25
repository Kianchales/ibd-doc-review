# Changelog

本文件记录 ipo-doc-formatting 的版本演进。版本规则：**功能/结构变更 bump minor（0.2.0→0.3.0），规则裁定/缺陷修复 bump patch（0.3.0→0.3.1）**；模板 docx 更新在对应版本下记录，不单独 bump。发布时打 `vX.Y.Z` tag。

## [0.3.0] - 2026-08-25

### 新增
- **内容完整性校验**：`check_styles.py --verify-content <原文.docx>`，套样式后文本与原文逐字对比（过滤合并单元格空文本），严禁修改原文内容
- **铁律 0「只改格式、严禁修改原文内容」**（用户裁定，最高优先级）：不得增删改移任何文字内容（含表格内文字/标点/空格/数字/单位）
- 自测扩至 8 用例（内容一致性 PASS/FAIL、表格内段落豁免回归）

### 修正
- **表格规范 v2**（表格模板.docx 内置示例表格完整实证 + 用户裁定）：
  - 宽度 100% 页宽 + Word 自动调整（`tblW 5000 pct` + `tblLayout autofit`）
  - 边框纯黑 000000（弃模板原 010000，用户裁定）
  - **两级表头**：组头跨列合并（gridSpan）+ 类别跨行合并（vMerge）+ 子头（金额/占比）
  - 数字右对齐、首列左对齐、合计行加粗跨列
  - 行属性：cantSplit + trHeight=397 + 行居中 + 表头行 tblHeader
  - 单元格边距 57 dxa、全表 10.5pt 垂直居中

### 模板留痕
- `assets/templates/` 模板源未更换；表格规范理解从「styles.xml 样式表」深化到「document.xml 内置示例」逐单元格实证

## [0.2.0] - 2026-08-25

### 新增
- 文档类型收敛为**两类**：招股书版（含报告/备忘录/尽调等正式文档）+ 反馈回复版（用户裁定）
- 模板迁至 `assets/templates/`（skill-creator 标准：输出资源归 assets；**模板即样式源，替换模板即定制输出**）
- **失败降级协议**：tencent-docx → minimax-docx → 脚本兜底（标注局限）→ Markdown，逐级降级不硬撑（用户裁定）
- `references/examples.md`：3 个典型用例全链路 + 降级示例 + 模板自定义示例
- `scripts/tests/test_check_styles.py`：自测 6 用例
- 触发词口语化补充（共 20 个）；frontmatter 加 version 字段

### 修正
- **段落间禁止空行/空段落**（用户裁定）：间距由样式 spacing 控制，脚本内置空段落检测（自闭合+配对两种形式）

### 模板留痕
- 模板 docx 从专家团 references/templates/ 复制入 skill（首次打包，自包含）

## [0.1.0] - 2026-08-25

### 新增
- 初版：招股书版 / 反馈回复版 / 报告版三体系样式应用（000-009、反馈回复 0011/001 监管问题黑体）
- `scripts/check_styles.py`：成品文档校验（必备样式/裸段落/标题跳级）+ 模板样式库校验
- 样式定义全部模板 styles.xml 实证提取
