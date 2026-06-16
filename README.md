# Word MathType Rebuild

一个用于将 Microsoft Word 原生公式逐个重建为 MathType 公式的 Codex skill。

A Codex skill for rebuilding Microsoft Word native equations as MathType equations one by one.

该 skill 适用于论文、手稿或技术文档编辑场景：需要把 Word OMML 原生公式替换为新插入的 MathType 公式，尤其是需要“右编号”格式的行间公式。默认不使用 MathType 的批量 `Convert Equations` 转换功能。

This skill is designed for manuscript and technical-document editing workflows where Word OMML equations must be replaced by newly inserted MathType equations, especially display equations that need MathType right-numbered formatting. It does not use MathType's bulk `Convert Equations` command by default.

## 功能 / What It Does

- 从 `.docx` 中提取行间 Word 原生公式，供逐个核对。
- Extracts display equations from a `.docx` for review.
- 辅助将行间公式映射到 Word COM 的 `OMaths` 索引。
- Helps map display equations to Word COM `OMaths` indexes.
- 将 Word 光标定位到目标原公式之后，准备插入新的 MathType 公式。
- Prepares the Word selection immediately after a target original equation.
- 指导 Codex 操作 MathType 插件，新建右编号 MathType 公式。
- Guides Codex through inserting a new MathType right-numbered equation.
- 只有在新的 MathType 对象已经插入后，才删除原 Word 公式。
- Deletes the original Word equation only after the new MathType object exists.
- 提供基于 Word COM 和 DOCX XML 的校验流程。
- Provides validation guidance using Word COM and DOCX XML checks.

## 环境要求 / Requirements

- Windows
- Microsoft Word
- 已安装 Word 插件的 MathType
- MathType with the Word add-in installed
- Python with `pywin32`
- Codex Computer Use，用于直接操作 Word 和 MathType 窗口
- Codex Computer Use for direct Word/MathType UI operation

可选但推荐：

Optional but recommended:

- Codex `documents` skill：当 LibreOffice 可用时，用于最终 DOCX 渲染和版面检查。
- Codex `documents` skill for final DOCX render/QA when LibreOffice is available.

## 安装 / Installation

将本仓库克隆或复制到 Codex skills 目录：

Clone or copy this repository into your Codex skills directory:

```powershell
git clone https://github.com/syczk301/word-mathtype-rebuild.git "$env:USERPROFILE\.codex\skills\word-mathtype-rebuild"
```

然后开启新的 Codex 会话，或刷新 skills，使 `$word-mathtype-rebuild` 可用。

Then start a new Codex session or refresh skills so `$word-mathtype-rebuild` becomes available.

## 使用方式 / Usage

在 Codex 中可以这样调用：

Invoke the skill in Codex with a request like:

```text
Use $word-mathtype-rebuild to replace the display Word equations in this DOCX with new MathType right-numbered equations one by one.
```

也可以用中文描述：

You can also ask in Chinese:

```text
使用 $word-mathtype-rebuild，把这个 Word 文档里的行间原生公式逐个新建为 MathType 右编号公式，然后删除原 Word 公式。
```

该流程需要在工作目录中准备公式映射文件：

The workflow expects a formula mapping file in the working directory:

```python
FORMULAS = [
    r"first formula TeX",
    r"second formula TeX",
]

DISPLAY_OMATH_INDEXES = [2, 8]
```

其中 `FORMULAS[n-1]` 与 `DISPLAY_OMATH_INDEXES[n-1]` 对应第 `n` 个行间公式。

Each `FORMULAS[n-1]` entry corresponds to `DISPLAY_OMATH_INDEXES[n-1]` for display formula `n`.

## 内置脚本 / Included Scripts

### `scripts/extract_display_formulas.py`

从 DOCX 中提取行间 OMML 公式：

Extract display OMML formulas from a DOCX:

```powershell
python scripts\extract_display_formulas.py manuscript.docx --out display_formulas_text.txt
```

### `scripts/prepare_formula_in_word.py`

打开或复用 Word 文档，将插入点定位到目标原公式之后，并把目标 TeX 写入文本文件：

Open or reuse a Word document, select the insertion point after a target original equation, and write the target TeX to a text file:

```powershell
python scripts\prepare_formula_in_word.py 3 --doc manuscript.docx --mapping formula_mapping.py
```

之后 Codex 应通过 Word 的 MathType 插件新建右编号公式，将 TeX 粘贴到 MathType，关闭 MathType 写回 Word，然后删除原 Word 公式。

After this, Codex should use Word's MathType add-in to insert a new right-numbered MathType equation, paste the TeX into MathType, close MathType to write back, and then delete the original Word equation.

## 注意事项 / Notes

- 行间公式建议按原始索引从后往前处理，避免删除后索引前移。
- Process display equations from highest original index to lowest to avoid index shifts.
- 编辑前必须备份 DOCX。
- Back up the DOCX before editing.
- MathType 的 TeX 解析并不完美；如果某个命令渲染异常，应改用更简单的等价写法。
- MathType TeX parsing is imperfect. If a command renders incorrectly, use a simpler equivalent expression.
- 如果 Word 提示该命令不能在 Equation Builder 公式内执行，说明光标仍在 Word 原生公式内部，需要把光标移到公式外再重试。
- If Word reports that the command does not work inside Equation Builder equations, move the cursor outside the Word native formula and retry.
- 该 skill 是给 Codex 使用的操作流程，不是独立的命令行批量转换器。
- This skill is an operational workflow for Codex, not a standalone batch conversion CLI.

## 许可证 / License

尚未指定许可证。

No license has been specified yet.
