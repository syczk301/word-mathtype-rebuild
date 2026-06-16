import argparse
import importlib.util
from pathlib import Path

import win32com.client as win32


def load_mapping(path):
    spec = importlib.util.spec_from_file_location("formula_mapping", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def main():
    parser = argparse.ArgumentParser(description="Prepare Word selection after an original display formula.")
    parser.add_argument("n", type=int, help="1-based display formula number in mapping")
    parser.add_argument("--doc", required=True, help="DOCX path to open or use if already active")
    parser.add_argument("--mapping", default="formula_mapping.py")
    parser.add_argument("--tex-out", default="current_tex.txt")
    args = parser.parse_args()

    mapping = load_mapping(args.mapping)
    idx = mapping.DISPLAY_OMATH_INDEXES[args.n - 1]
    tex = mapping.FORMULAS[args.n - 1]

    word = win32.Dispatch("Word.Application")
    word.Visible = True
    doc_path = str(Path(args.doc).resolve())
    doc = None
    for i in range(1, word.Documents.Count + 1):
        candidate = word.Documents(i)
        if str(candidate.FullName).lower() == doc_path.lower():
            doc = candidate
            break
    if doc is None:
        doc = word.Documents.Open(doc_path)

    rng = doc.OMaths(idx).Range.Duplicate
    print("target", args.n, idx, rng.Start, rng.End, repr(rng.Text[:120]))
    rng.Collapse(0)
    rng.InsertAfter("\r")
    rng.Collapse(0)
    rng.Select()
    word.Activate()

    Path(args.tex_out).write_text(tex, encoding="utf-8")
    print(f"prepared formula {args.n}; wrote TeX to {args.tex_out}")


if __name__ == "__main__":
    main()
