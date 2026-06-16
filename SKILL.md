---
name: word-mathtype-rebuild
description: Rebuild Microsoft Word native equations as MathType equations in .docx manuscripts. Use when Codex must operate Word with the MathType add-in to replace Word OMML formulas one by one, especially display equations that must be newly inserted as MathType right-numbered equations and original Word equations deleted afterward, without using MathType bulk conversion.
---

# Word MathType Rebuild

Use this skill for fragile Word + MathType conversion work where the user wants formulas recreated manually instead of bulk converted.

## Operating Rules

- Do not use MathType "Convert Equations" for this workflow unless the user explicitly asks for bulk conversion.
- Process formulas one by one. For display equations, create a new MathType right-numbered equation first, then delete the original Word equation.
- Preserve a reversible backup before editing the DOCX.
- Work in descending order of original display-equation indexes when deleting originals. This avoids lower original indexes shifting.
- Save after each successful formula, or at least after each small batch.
- If Word or MathType becomes busy, press `Esc`, wait, and retry COM access. Do not force-close Word unless a valid backup exists.

## Required Tools

- Windows Microsoft Word with the MathType Word add-in installed.
- Computer Use for interacting with Word and MathType windows.
- Python with `pywin32` for Word COM automation.
- The `documents` skill is recommended for final DOCX QA. If LibreOffice/`soffice` is missing, state that render QA could not run and use Word COM + OOXML structural checks instead.

## Workflow

1. Back up the input DOCX.
   - Use a clear name such as `<stem>_before_mathtype_rebuild.docx`.
   - Never overwrite this backup.

2. Extract display formulas.
   - Use `scripts/extract_display_formulas.py <docx> --out display_formulas_text.txt`.
   - Review each extracted formula and prepare a TeX list manually. MathType TeX parsing is imperfect; avoid commands it mishandles.
   - Known safer replacements:
     - Use `floor(...)` instead of `\lfloor ... \rfloor` if MathType renders floor delimiters as broken glyphs.
     - Use plain function names like `Top-u` when `\operatorname{Top}\text{-}u` renders poorly.

3. Record original display equation indexes.
   - Display equations in DOCX XML are `m:oMathPara`; Word COM indexes all `OMaths`.
   - Use the extraction script output to map display equations to Word COM `OMaths` indexes.
   - Process display formulas from highest original index to lowest.

4. For each display formula:
   - Put the intended TeX into the clipboard.
   - Use `scripts/prepare_formula_in_word.py <n> --doc <path> --mapping <formula_mapping.py>` to place the Word selection immediately after the original formula.
   - Trigger MathType "Insert Right Numbered Display Equation".
     - Preferred UI path: Word MathType tab -> the right-numbered equation button, labeled "Right Number" or localized as right-numbered.
     - If available, run the add-in macro in a background Python process:
       `Word.Application.Run("MathTypeCommands.UILib.MTCommand_InsertRightNumberedDispEqn")`.
     - Do not let the macro call block the main process indefinitely; launch it separately if needed.
   - In the MathType editor, click inside the blank editor, paste the TeX, confirm it rendered correctly, and close the MathType window with `Alt+F4` to write it back to Word.
   - If MathType inserted an empty right-numbered row, delete the extra blank row and keep the real equation row.
   - Delete the original Word equation range with COM only after the new MathType equation is present.
   - Remove isolated punctuation-only paragraphs left by the deleted original, such as `,\r` or `.\r`, when they are clearly leftovers.

5. Validate after each batch.
   - Word COM count expectations are only a hint; COM can report stale `OMaths`.
   - Prefer saved DOCX XML checks:
     - `word/document.xml` should have fewer `<m:oMath>` elements by the number of display formulas replaced.
     - `word/document.xml` should have no `<m:oMathPara>` for replaced display formulas.
     - MathType display equations appear as Word objects, typically `<w:object>`.
   - Reopen the DOCX in Word to ensure it is not corrupt.

6. Finalize.
   - Save the active working document back to the requested output path.
   - Keep the backup.
   - Run render QA through the `documents` skill when possible.
   - Report exactly what was processed and any skipped items.

## Computer Use Notes

- Initialize Computer Use with the bundled runtime when available.
- Use `sky.list_windows()` if `sky.list_apps()` misses a MathType window.
- Typical MathType editor paste sequence:
  - activate MathType window
  - call `get_window_state`
  - click near the blank editor area
  - press `Ctrl+V`
  - inspect screenshot visually
  - press `Alt+F4`
- If Word shows "This command does not work inside Equation Builder equations", the cursor is inside a Word native formula. Use COM to reselect a collapsed range after the formula and trigger MathType again.

## Minimal Mapping Module Format

Create a `formula_mapping.py` in the working directory:

```python
FORMULAS = [
    r"first formula TeX",
    r"second formula TeX",
]

DISPLAY_OMATH_INDEXES = [2, 8]
```

The helper scripts expect `FORMULAS[n-1]` and `DISPLAY_OMATH_INDEXES[n-1]` to correspond to display formula number `n`.
