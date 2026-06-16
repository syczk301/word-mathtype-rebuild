# Word MathType Rebuild

A Codex skill for rebuilding Microsoft Word native equations as MathType equations one by one.

This skill is designed for manuscript editing workflows where Word OMML equations must be replaced by newly inserted MathType equations, especially display equations that need MathType right-numbered formatting. It does not use MathType's bulk "Convert Equations" command by default.

## What It Does

- Extracts display equations from a `.docx` for review.
- Helps map display equations to Word COM `OMaths` indexes.
- Prepares Word selection immediately after a target original equation.
- Guides Codex through inserting a new MathType right-numbered equation.
- Deletes the original Word equation only after the new MathType object exists.
- Provides validation guidance using Word COM and DOCX XML checks.

## Requirements

- Windows
- Microsoft Word
- MathType with the Word add-in installed
- Python with `pywin32`
- Codex Computer Use for direct Word/MathType UI operation

Optional but recommended:

- Codex `documents` skill for final DOCX render/QA when LibreOffice is available.

## Installation

Clone or copy this folder into your Codex skills directory:

```powershell
git clone https://github.com/syczk301/word-mathtype-rebuild.git "$env:USERPROFILE\.codex\skills\word-mathtype-rebuild"
```

Then start a new Codex session or refresh skills so `$word-mathtype-rebuild` becomes available.

## Usage

Invoke the skill in Codex with a request like:

```text
Use $word-mathtype-rebuild to replace the display Word equations in this DOCX with new MathType right-numbered equations one by one.
```

The workflow expects a formula mapping file in the working directory:

```python
FORMULAS = [
    r"first formula TeX",
    r"second formula TeX",
]

DISPLAY_OMATH_INDEXES = [2, 8]
```

Each `FORMULAS[n-1]` entry corresponds to `DISPLAY_OMATH_INDEXES[n-1]`.

## Included Scripts

### `scripts/extract_display_formulas.py`

Extracts display OMML formulas from a DOCX:

```powershell
python scripts\extract_display_formulas.py manuscript.docx --out display_formulas_text.txt
```

### `scripts/prepare_formula_in_word.py`

Opens or reuses a Word document, selects the insertion point after a target original equation, and writes the target TeX to a text file:

```powershell
python scripts\prepare_formula_in_word.py 3 --doc manuscript.docx --mapping formula_mapping.py
```

After this, Codex should use Word's MathType add-in to insert a new right-numbered MathType equation, paste the TeX into MathType, close MathType to write back, and then delete the original Word equation.

## Notes

- Process display equations from highest original index to lowest to avoid index shifts.
- Back up the DOCX before editing.
- MathType TeX parsing is imperfect. If a command renders incorrectly, use a simpler equivalent expression.
- If Word reports that the command does not work inside Equation Builder equations, move the cursor outside the Word native formula and retry.

## License

No license has been specified yet.
