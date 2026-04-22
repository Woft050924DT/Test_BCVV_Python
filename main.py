from __future__ import annotations

import argparse
import csv
import re
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable
import xml.etree.ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}
ET.register_namespace("w", W_NS)

ENDING_LABELS = {
    ".": "period",
    ";": "semicolon",
    ":": "colon",
    "!": "exclamation",
    "?": "question",
    "none": "none",
}
IGNORED_STYLE_KEYWORDS = {
    "toc",
    "tableofcontents",
    "header",
    "footer",
    "page number",
    "index",
}
WARNING_KEYWORDS = {
    "warning",
    "caution",
    "note",
    "avertissement",
    "mise en garde",
    "remarque",
}
ABBREVIATION_PATTERN = re.compile(
    r"(?i)\b(?:max|min|approx|appt|env|etc|ref|no|nr|n°|fig|al|art|sec|dr|mr|mrs|ms|sr|jr)\.$"
)


@dataclass
class Item:
    item_id: str
    kind: str
    text: str
    normalized_text: str
    ending: str
    style: str = ""
    section_path: tuple[str, ...] = ()
    group_key: str = ""
    group_label: str = ""
    locator: str = ""
    reason_context: str = ""
    element: ET.Element | None = None
    metadata: dict[str, str | int] = field(default_factory=dict)


@dataclass
class Finding:
    item: Item
    expected: str
    reason: str


def qn(tag: str) -> str:
    prefix, name = tag.split(":")
    if prefix != "w":
        raise ValueError(f"Unsupported namespace prefix: {prefix}")
    return f"{{{W_NS}}}{name}"


def iter_text_nodes(element: ET.Element) -> Iterable[ET.Element]:
    return element.findall(".//w:t", NS)


def extract_text(element: ET.Element) -> str:
    return "".join(node.text or "" for node in iter_text_nodes(element))


def normalize_whitespace(text: str) -> str:
    text = text.replace("\xa0", " ")
    return re.sub(r"\s+", " ", text).strip()


def text_snippet(text: str, limit: int = 140) -> str:
    collapsed = normalize_whitespace(text)
    if len(collapsed) <= limit:
        return collapsed
    return collapsed[: limit - 3].rstrip() + "..."


def canonical_style_name(style: str) -> str:
    return normalize_whitespace(style).lower()


def is_ignored_style(style: str) -> bool:
    lowered = canonical_style_name(style)
    return any(keyword in lowered for keyword in IGNORED_STYLE_KEYWORDS)


def looks_like_warning(text: str) -> bool:
    lowered = normalize_whitespace(text).lower()
    return any(keyword in lowered for keyword in WARNING_KEYWORDS)


def is_abbreviation_period(text: str) -> bool:
    trimmed = normalize_whitespace(text)
    if not trimmed.endswith("."):
        return False
    if ABBREVIATION_PATTERN.search(trimmed):
        return True
    return bool(re.search(r"(?:\b[A-Za-z]\.){2,}$", trimmed))


def classify_ending(text: str) -> str:
    trimmed = normalize_whitespace(text)
    if not trimmed:
        return "none"
    if trimmed.endswith("."):
        return "none" if is_abbreviation_period(trimmed) else "."
    for marker in (";", ":", "!", "?"):
        if trimmed.endswith(marker):
            return marker
    return "none"


def paragraph_style(paragraph: ET.Element) -> str:
    style = paragraph.find("./w:pPr/w:pStyle", NS)
    if style is None:
        return ""
    return style.attrib.get(qn("w:val"), "")


def heading_level(style: str) -> int | None:
    match = re.search(r"heading\s*(\d+)$", canonical_style_name(style))
    if match:
        return int(match.group(1))
    match = re.search(r"titre\s*(\d+)$", canonical_style_name(style))
    if match:
        return int(match.group(1))
    return None


def iter_body_blocks(body: ET.Element) -> Iterable[ET.Element]:
    for child in body:
        local = child.tag.rsplit("}", 1)[-1]
        if local in {"p", "tbl"}:
            yield child


def add_highlight(element: ET.Element, color: str = "yellow") -> None:
    runs = element.findall(".//w:r", NS)
    if not runs:
        return
    for run in runs:
        rpr = run.find("w:rPr", NS)
        if rpr is None:
            rpr = ET.Element(qn("w:rPr"))
            run.insert(0, rpr)
        highlight = rpr.find("w:highlight", NS)
        if highlight is None:
            highlight = ET.SubElement(rpr, qn("w:highlight"))
        highlight.set(qn("w:val"), color)


def parse_document(root: ET.Element) -> list[Item]:
    items: list[Item] = []
    body = root.find("w:body", NS)
    if body is None:
        return items

    section_stack: list[str] = []
    para_index = 0
    table_index = 0

    for block in iter_body_blocks(body):
        local = block.tag.rsplit("}", 1)[-1]
        if local == "p":
            text = normalize_whitespace(extract_text(block))
            style = paragraph_style(block)
            if not text or is_ignored_style(style):
                para_index += 1
                continue

            level = heading_level(style)
            if level is not None:
                while len(section_stack) >= level:
                    section_stack.pop()
                section_stack.append(text)

            item = Item(
                item_id=f"p-{para_index}",
                kind="paragraph",
                text=text,
                normalized_text=text.casefold(),
                ending=classify_ending(text),
                style=style or "Normal",
                section_path=tuple(section_stack),
                locator=f"paragraph {para_index + 1}",
                element=block,
                metadata={"paragraph_index": para_index},
            )
            items.append(item)
            para_index += 1
            continue

        table_items = parse_table(block, table_index, tuple(section_stack))
        items.extend(table_items)
        table_index += 1

    return items


def parse_table(table: ET.Element, table_index: int, section_path: tuple[str, ...]) -> list[Item]:
    rows = table.findall("./w:tr", NS)
    matrix: list[list[ET.Element]] = [row.findall("./w:tc", NS) for row in rows]
    if not matrix:
        return []

    col_count = max((len(row) for row in matrix), default=0)
    header_texts = extract_header_texts(matrix[0], col_count)
    data_start = detect_table_data_start(matrix)

    items: list[Item] = []
    for row_idx, row_cells in enumerate(matrix[data_start:], start=data_start):
        if is_warning_row(row_cells):
            continue
        for col_idx in range(col_count):
            if col_idx >= len(row_cells):
                continue
            cell = row_cells[col_idx]
            text = normalize_whitespace(extract_text(cell))
            if not text:
                continue
            if looks_like_warning(text) and col_idx == 0:
                continue

            header = header_texts[col_idx] if col_idx < len(header_texts) else ""
            row_signature = classify_row_signature(row_cells)
            items.append(
                Item(
                    item_id=f"t{table_index}-r{row_idx}-c{col_idx}",
                    kind="table_cell",
                    text=text,
                    normalized_text=text.casefold(),
                    ending=classify_ending(text),
                    section_path=section_path,
                    locator=f"table {table_index + 1}, row {row_idx + 1}, col {col_idx + 1}",
                    element=cell,
                    metadata={
                        "table_index": table_index,
                        "row_index": row_idx,
                        "col_index": col_idx,
                        "header": header,
                        "row_signature": row_signature,
                    },
                )
            )
    return items


def extract_header_texts(first_row: list[ET.Element], col_count: int) -> list[str]:
    headers = []
    for col_idx in range(col_count):
        if col_idx < len(first_row):
            headers.append(text_snippet(extract_text(first_row[col_idx]), 50))
        else:
            headers.append("")
    return headers


def detect_table_data_start(matrix: list[list[ET.Element]]) -> int:
    if len(matrix) <= 1:
        return 0
    first_row_texts = [normalize_whitespace(extract_text(cell)) for cell in matrix[0]]
    non_empty = [text for text in first_row_texts if text]
    if non_empty and all(len(text) <= 60 for text in non_empty):
        return 1
    return 0


def classify_row_signature(row_cells: list[ET.Element]) -> str:
    lengths = []
    for cell in row_cells:
        text = normalize_whitespace(extract_text(cell))
        if not text:
            lengths.append("empty")
        elif len(text) <= 24:
            lengths.append("short")
        elif len(text) <= 80:
            lengths.append("medium")
        else:
            lengths.append("long")
    return "|".join(lengths)


def is_warning_row(row_cells: list[ET.Element]) -> bool:
    texts = [normalize_whitespace(extract_text(cell)) for cell in row_cells]
    non_empty = [text for text in texts if text]
    if not non_empty:
        return True
    combined = " ".join(non_empty).lower()
    return len(non_empty) <= 2 and any(combined.startswith(keyword) or f" {keyword}" in combined[:80] for keyword in WARNING_KEYWORDS)


def build_groups(items: list[Item]) -> dict[str, list[Item]]:
    groups: dict[str, list[Item]] = {}
    paragraph_runs = build_paragraph_runs([item for item in items if item.kind == "paragraph"])
    for group_key, grouped_items, label in paragraph_runs:
        for item in grouped_items:
            item.group_key = group_key
            item.group_label = label
        groups[group_key] = grouped_items

    for group_key, grouped_items, label in build_table_groups([item for item in items if item.kind == "table_cell"]):
        for item in grouped_items:
            item.group_key = group_key
            item.group_label = label
        groups[group_key] = grouped_items
    return groups


def build_paragraph_runs(paragraphs: list[Item]) -> list[tuple[str, list[Item], str]]:
    groups: list[tuple[str, list[Item], str]] = []
    current: list[Item] = []
    current_key = ""
    current_label = ""

    def flush() -> None:
        nonlocal current, current_key, current_label
        if current:
            groups.append((current_key, current[:], current_label))
        current = []
        current_key = ""
        current_label = ""

    for item in paragraphs:
        style_key = canonical_style_name(item.style)
        section_key = " > ".join(item.section_path[-2:]) if item.section_path else "root"
        key = f"paragraph::{style_key}::{section_key}"
        label = f"Paragraphs style={item.style} section={section_key}"
        boundary = heading_level(item.style) is not None or looks_like_warning(item.text)

        if not current:
            current = [item]
            current_key = key
            current_label = label
            if boundary:
                flush()
            continue

        if key != current_key or boundary:
            flush()
            current = [item]
            current_key = key
            current_label = label
            if boundary:
                flush()
            continue

        current.append(item)

    flush()
    return groups


def build_table_groups(table_cells: list[Item]) -> list[tuple[str, list[Item], str]]:
    grouped: dict[str, list[Item]] = {}
    labels: dict[str, str] = {}

    for item in table_cells:
        table_index = int(item.metadata["table_index"])
        col_index = int(item.metadata["col_index"])
        header = str(item.metadata.get("header") or f"col {col_index + 1}")
        row_signature = str(item.metadata.get("row_signature") or "")
        key = f"table::{table_index}::col::{col_index}::shape::{row_signature}"
        label = f"Table {table_index + 1} column {col_index + 1} header={header}"
        grouped.setdefault(key, []).append(item)
        labels[key] = label

    return [(key, values, labels[key]) for key, values in grouped.items()]


def is_candidate_group(group: list[Item]) -> bool:
    if len(group) < 3:
        return False
    if group[0].kind == "paragraph":
        average_length = sum(len(item.text) for item in group) / len(group)
        style_name = canonical_style_name(group[0].style)
        if style_name in {"normal", ""} and not group[0].section_path:
            return False
        if style_name in {"normal", ""} and average_length > 110:
            return False
    endings = Counter(item.ending for item in group)
    if len(endings) <= 1:
        return False
    dominant_ending, dominant_count = endings.most_common(1)[0]
    if dominant_ending == "none" and dominant_count == len(group):
        return False
    return dominant_count >= 2 and dominant_count / len(group) >= 0.6


def detect_findings(groups: dict[str, list[Item]]) -> list[Finding]:
    findings: list[Finding] = []
    for group_key, group_items in groups.items():
        if not is_candidate_group(group_items):
            continue
        endings = Counter(item.ending for item in group_items)
        expected, expected_count = endings.most_common(1)[0]

        for item in group_items:
            if item.ending == expected:
                continue
            reason = (
                f"Ending '{format_ending(item.ending)}' differs from dominant ending "
                f"'{format_ending(expected)}' in group ({expected_count}/{len(group_items)} items)."
            )
            findings.append(Finding(item=item, expected=expected, reason=reason))
    return findings


def format_ending(ending: str) -> str:
    return ENDING_LABELS.get(ending, ending)


def write_report(findings: list[Finding], output_csv: Path) -> None:
    output_csv.parent.mkdir(parents=True, exist_ok=True)
    with output_csv.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.writer(handle)
        writer.writerow(
            ["Text snippet", "Detected ending", "Expected ending", "Group identifier", "Reason for flagging", "Location"]
        )
        for finding in findings:
            writer.writerow(
                [
                    text_snippet(finding.item.text),
                    format_ending(finding.item.ending),
                    format_ending(finding.expected),
                    finding.item.group_label or finding.item.group_key,
                    finding.reason,
                    finding.item.locator,
                ]
            )


def write_highlighted_docx(input_docx: Path, output_docx: Path, findings: list[Finding]) -> None:
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    for finding in findings:
        if finding.item.element is not None:
            add_highlight(finding.item.element)

    document_xml = ET.tostring(DOCUMENT_ROOT, encoding="utf-8", xml_declaration=True)
    with zipfile.ZipFile(input_docx, "r") as source, zipfile.ZipFile(output_docx, "w") as target:
        for entry in source.infolist():
            if entry.filename == "word/document.xml":
                target.writestr(entry, document_xml)
            else:
                target.writestr(entry, source.read(entry.filename))


def load_document(input_docx: Path) -> ET.Element:
    with zipfile.ZipFile(input_docx) as archive:
        return ET.fromstring(archive.read("word/document.xml"))


def explain_groups(groups: dict[str, list[Item]]) -> str:
    paragraph_groups = sum(1 for key in groups if key.startswith("paragraph::"))
    table_groups = sum(1 for key in groups if key.startswith("table::"))
    return (
        "Grouping logic:\n"
        "- Paragraphs are grouped by consecutive runs that share style and local section context.\n"
        "- Table cells are grouped within the same table and column, further split by row-shape signature.\n"
        "- A group is checked only when it has at least 3 items and one ending style clearly dominates.\n"
        f"- Built {paragraph_groups} paragraph groups and {table_groups} table groups.\n"
    )


def write_summary(summary_path: Path, groups: dict[str, list[Item]], findings: list[Finding]) -> None:
    summary_path.write_text(
        explain_groups(groups)
        + f"\nFindings: {len(findings)}\n",
        encoding="utf-8",
    )


def process_document(input_docx: Path, output_docx: Path, output_csv: Path, summary_path: Path) -> list[Finding]:
    global DOCUMENT_ROOT
    DOCUMENT_ROOT = load_document(input_docx)
    items = parse_document(DOCUMENT_ROOT)
    groups = build_groups(items)
    findings = detect_findings(groups)
    write_report(findings, output_csv)
    write_summary(summary_path, groups, findings)
    write_highlighted_docx(input_docx, output_docx, findings)
    return findings


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Detect punctuation inconsistencies in DOCX files.")
    parser.add_argument("input_docx", type=Path, nargs="?", help="Path to the source .docx file")
    parser.add_argument(
        "--output-docx",
        type=Path,
        default=Path("output/highlighted_punctuation_review.docx"),
        help="Path to write the highlighted .docx output",
    )
    parser.add_argument(
        "--output-report",
        type=Path,
        default=Path("output/punctuation_report.csv"),
        help="Path to write the CSV report",
    )
    parser.add_argument(
        "--output-summary",
        type=Path,
        default=Path("output/grouping_summary.txt"),
        help="Path to write the grouping summary",
    )
    return parser


def main() -> int:
    parser = build_arg_parser()
    args = parser.parse_args()
    input_docx = resolve_input_docx(args.input_docx)

    findings = process_document(
        input_docx=input_docx,
        output_docx=args.output_docx,
        output_csv=args.output_report,
        summary_path=args.output_summary,
    )
    print(f"Processed {input_docx}")
    print(f"Findings: {len(findings)}")
    print(f"Highlighted DOCX: {args.output_docx}")
    print(f"CSV report: {args.output_report}")
    print(f"Summary: {args.output_summary}")
    return 0


def resolve_input_docx(cli_value: Path | None) -> Path:
    if cli_value is not None:
        return cli_value

    local_sample = Path("BCVV_Assignment_Sample_file.docx")
    if local_sample.exists():
        return local_sample

    downloads_sample = Path.home() / "Downloads" / "BCVV_Assignment_Sample_file.docx"
    if downloads_sample.exists():
        return downloads_sample

    raise SystemExit(
        "No input_docx was provided and no default sample file was found. "
        "Pass a .docx path or place BCVV_Assignment_Sample_file.docx in the project folder or Downloads."
    )


DOCUMENT_ROOT: ET.Element


if __name__ == "__main__":
    raise SystemExit(main())
