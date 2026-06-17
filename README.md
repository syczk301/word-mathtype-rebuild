# Word MathType Rebuild

<p align="right">
  <a href="./README.en.md"><strong>English</strong></a>
</p>

一个用于将 Microsoft Word 原生公式逐个重建为 MathType 公式的 Codex skill。

该 skill 适用于论文、手稿或技术文档编辑场景：需要把 Word OMML 原生公式替换为新插入的 MathType 公式，尤其是需要“右编号”格式的行间公式。默认不使用 MathType 的批量 `Convert Equations` 转换功能。

## 功能

- 从 `.docx` 中提取行间 Word 原生公式，供逐个核对。
- 自动生成行间公式映射文件，将行间公式映射到 Word COM 的 `OMaths` 索引。
- 将 Word 光标定位到目标原公式之后，准备插入新的 MathType 公式。
- 指导 Codex 操作 MathType 插件，新建右编号 MathType 公式。
- 只有在新的 MathType 对象已经插入后，才删除原 Word 公式。
- 提供基于 Word COM 和 DOCX XML 的校验流程。

## 环境要求

- Windows
- Microsoft Word
- 已安装 Word 插件的 MathType
- Python with `pywin32`
- Codex Computer Use，用于直接操作 Word 和 MathType 窗口

可选但推荐：

- Codex `documents` skill：当 LibreOffice 可用时，用于最终 DOCX 渲染和版面检查。

## 安装

将本仓库克隆或复制到 Codex skills 目录：

```powershell
git clone https://github.com/syczk301/word-mathtype-rebuild.git "$env:USERPROFILE\.codex\skills\word-mathtype-rebuild"
```

然后开启新的 Codex 会话，或刷新 skills，使 `$word-mathtype-rebuild` 可用。

## 使用方式

在 Codex 中可以这样调用：

```text
使用 $word-mathtype-rebuild，把这个 Word 文档里的行间原生公式逐个新建为 MathType 右编号公式，然后删除原 Word 公式。
```

提取脚本会自动在工作目录中生成公式映射文件：

```python
FORMULAS = [
    r"first formula TeX",
    r"second formula TeX",
]

DISPLAY_OMATH_INDEXES = [2, 8]
```

其中 `FORMULAS[n-1]` 与 `DISPLAY_OMATH_INDEXES[n-1]` 对应第 `n` 个行间公式。生成的 `FORMULAS` 来自 OMML 线性文本，插入 MathType 前仍需核对并按需要改写成 MathType 可解析的 TeX。

## 内置脚本

### `scripts/extract_display_formulas.py`

从 DOCX 中提取行间 OMML 公式，并自动生成 `formula_mapping.py`。默认 `--scope auto` 会优先提取 `m:oMathPara` 行间公式；如果没有行间公式但仍有 Word 原生公式，则退回映射全部 `m:oMath` 并打印警告。

```powershell
python scripts\extract_display_formulas.py manuscript.docx --out display_formulas_text.txt --mapping-out formula_mapping.py
```

### `scripts/prepare_formula_in_word.py`

打开或复用 Word 文档，将插入点定位到目标原公式之后，并把目标 TeX 写入文本文件：

```powershell
python scripts\prepare_formula_in_word.py 3 --doc manuscript.docx --mapping formula_mapping.py
```

之后 Codex 应通过 Word 的 MathType 插件新建右编号公式，将 TeX 粘贴到 MathType，关闭 MathType 写回 Word，然后删除原 Word 公式。

## 注意事项

- 行间公式建议按原始索引从后往前处理，避免删除后索引前移。
- 编辑前必须备份 DOCX。
- MathType 的 TeX 解析并不完美；如果某个命令渲染异常，应改用更简单的等价写法。
- 如果 Word 提示该命令不能在 Equation Builder 公式内执行，说明光标仍在 Word 原生公式内部，需要把光标移到公式外再重试。
- 该 skill 是给 Codex 使用的操作流程，不是独立的命令行批量转换器。

## 许可证

本项目使用 [MIT License](./LICENSE)。
