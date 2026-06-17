import argparse
from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET


NS = {
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}


def py_raw_string(value):
    value = value.replace("\\", "\\\\").replace('"""', '\\"\\"\\"')
    return f'r"""{value}"""'


def collect_display_formulas(root):
    formulas = []
    omath_index = 0
    active_display = None

    def walk(node):
        nonlocal omath_index, active_display
        previous_display = active_display

        if node.tag == f"{{{NS['m']}}}oMathPara":
            active_display = {
                "display_no": len(formulas) + 1,
                "omath_indexes": [],
                "text_parts": [],
            }

        if node.tag == f"{{{NS['m']}}}oMath":
            omath_index += 1
            if active_display is not None:
                active_display["omath_indexes"].append(omath_index)

        if node.tag == f"{{{NS['m']}}}t" and node.text and active_display is not None:
            active_display["text_parts"].append(node.text)

        for child in list(node):
            walk(child)

        if node.tag == f"{{{NS['m']}}}oMathPara":
            text = "".join(active_display["text_parts"]).strip()
            indexes = active_display["omath_indexes"]
            formulas.append(
                {
                    "number": active_display["display_no"],
                    "kind": "display",
                    "omath_indexes": indexes,
                    "text": text,
                }
            )
            active_display = previous_display

    walk(root)
    return formulas


def collect_all_omaths(root):
    formulas = []
    omath_index = 0

    for node in root.iter():
        if node.tag == f"{{{NS['m']}}}oMath":
            omath_index += 1
            formulas.append(
                {
                    "number": omath_index,
                    "kind": "omath",
                    "omath_indexes": [omath_index],
                    "text": "".join(
                        t.text
                        for t in node.iter()
                        if t.tag == f"{{{NS['m']}}}t" and t.text
                    ).strip(),
                }
            )
    return formulas


def write_text_listing(path, formulas):
    with open(path, "w", encoding="utf-8") as f:
        for item in formulas:
            number = item["number"]
            kind = item["kind"]
            indexes = item["omath_indexes"]
            text = item["text"]
            first_idx = indexes[0] if indexes else None
            all_indexes = ", ".join(str(i) for i in indexes) if indexes else "none"
            f.write(
                f"{number}. {kind}; OMath index {first_idx}; "
                f"OMath indexes [{all_indexes}]: {text}\n"
            )


def write_mapping_module(path, formulas, source_docx, scope):
    lines = [
        '"""Auto-generated display formula mapping.',
        "",
        "Review FORMULAS before using MathType. The extracted strings are OMML",
        "linear text, not guaranteed MathType TeX.",
        '"""',
        "",
        f"SOURCE_DOCX = {str(Path(source_docx).resolve())!r}",
        f"EXTRACTION_SCOPE = {scope!r}",
        "",
        "FORMULAS = [",
    ]
    for item in formulas:
        number = item["number"]
        kind = item["kind"]
        indexes = item["omath_indexes"]
        text = item["text"]
        first_idx = indexes[0] if indexes else None
        lines.append(f"    # {kind} {number}, Word OMath index {first_idx}")
        lines.append(f"    {py_raw_string(text)},")
    lines.extend(
        [
            "]",
            "",
            "DISPLAY_OMATH_INDEXES = [",
        ]
    )
    for item in formulas:
        number = item["number"]
        kind = item["kind"]
        indexes = item["omath_indexes"]
        first_idx = indexes[0] if indexes else None
        lines.append(f"    {first_idx},  # {kind} {number}")
    lines.extend(
        [
            "]",
            "",
            "DISPLAY_OMATH_INDEX_GROUPS = [",
        ]
    )
    for item in formulas:
        number = item["number"]
        kind = item["kind"]
        indexes = item["omath_indexes"]
        lines.append(f"    {indexes!r},  # {kind} {number}")
    lines.append("]")
    Path(path).write_text("\n".join(lines) + "\n", encoding="utf-8")


def main():
    parser = argparse.ArgumentParser(description="Extract display OMML formulas from a DOCX.")
    parser.add_argument("docx")
    parser.add_argument("--out", default="display_formulas_text.txt")
    parser.add_argument(
        "--mapping-out",
        default="formula_mapping.py",
        help="Write a formula_mapping.py module for prepare_formula_in_word.py. Use an empty value to skip.",
    )
    parser.add_argument(
        "--scope",
        choices=("auto", "display", "all"),
        default="auto",
        help="auto: display formulas if present, otherwise all Word native formulas.",
    )
    args = parser.parse_args()

    with zipfile.ZipFile(args.docx) as z:
        root = ET.fromstring(z.read("word/document.xml"))

    display_formulas = collect_display_formulas(root)
    if args.scope == "display":
        formulas = display_formulas
        scope = "display"
    elif args.scope == "all":
        formulas = collect_all_omaths(root)
        scope = "all"
    elif display_formulas:
        formulas = display_formulas
        scope = "display"
    else:
        formulas = collect_all_omaths(root)
        scope = "all-fallback"

    write_text_listing(args.out, formulas)
    print(f"wrote {len(formulas)} formulas with scope {scope} to {args.out}")
    if scope == "all-fallback" and formulas:
        print("warning: no m:oMathPara display formulas found; mapped all m:oMath formulas instead")
    if args.mapping_out:
        write_mapping_module(args.mapping_out, formulas, args.docx, scope)
        print(f"wrote mapping module to {args.mapping_out}")


if __name__ == "__main__":
    main()
