import argparse
import zipfile
import xml.etree.ElementTree as ET


NS = {
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}


def iter_text(node):
    for t in node.iter():
        if t.tag == f"{{{NS['m']}}}t" and t.text:
            yield t.text


def main():
    parser = argparse.ArgumentParser(description="Extract display OMML formulas from a DOCX.")
    parser.add_argument("docx")
    parser.add_argument("--out", default="display_formulas_text.txt")
    args = parser.parse_args()

    with zipfile.ZipFile(args.docx) as z:
        root = ET.fromstring(z.read("word/document.xml"))

    lines = []
    omath_index = 0
    for elem in root.iter():
        if elem.tag == f"{{{NS['m']}}}oMath":
            omath_index += 1
        if elem.tag == f"{{{NS['m']}}}oMathPara":
            text = "".join(iter_text(elem)).strip()
            lines.append((len(lines) + 1, omath_index, text))

    with open(args.out, "w", encoding="utf-8") as f:
        for display_no, omath_idx, text in lines:
            f.write(f"{display_no}. OMath index around {omath_idx}: {text}\n")

    print(f"wrote {len(lines)} display formulas to {args.out}")


if __name__ == "__main__":
    main()
