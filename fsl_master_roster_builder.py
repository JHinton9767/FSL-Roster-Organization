from __future__ import annotations

import argparse
import re
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm", ".xltx", ".xltm"}

STANDARD_COLUMNS = [
    "Year",
    "Source File",
    "Source Sheet",
    "Chapter",
    "Last Name",
    "First Name",
    "Banner ID",
    "Email",
    "Status",
    "Semester Joined",
    "Position",
]

HEADER_ALIASES = {
    "last_name": [
        "last name",
        "lastname",
        "surname",
        "member last name",
    ],
    "first_name": [
        "first name",
        "firstname",
        "given name",
        "member first name",
    ],
    "banner_id": [
        "banner id",
        "student id",
        "id",
        "banner",
        "student number",
        "banner number",
    ],
    "email": [
        "email",
        "e-mail",
        "email address",
        "student email",
    ],
    "status": [
        "status",
        "member status",
        "membership status",
        "roster status",
    ],
    "semester_joined": [
        "semester joined",
        "joined",
        "join term",
        "semester initiated",
        "term joined",
        "semester admitted",
        "initiation term",
    ],
    "position": [
        "position",
        "office",
        "role",
        "member/council",
        "member council",
        "title",
    ],
    "chapter": [
        "chapter",
        "organization",
        "org",
        "group",
        "fraternity/sorority",
        "fsl organization",
    ],
}

STATUS_MAP = {
    "A": "Active",
    "AL": "Alumni",
    "G": "Graduated",
    "I": "Inactive",
    "S": "Suspended",
    "N": "New Member",
    "RS": "Resigned",
    "RV": "Revoked",
    "T": "Transfer",
}


@dataclass
class ExtractedRow:
    year: str
    source_file: str
    source_sheet: str
    chapter: str
    last_name: str
    first_name: str
    banner_id: str
    email: str
    status: str
    semester_joined: str
    position: str

    def as_list(self) -> List[str]:
        return [
            self.year,
            self.source_file,
            self.source_sheet,
            self.chapter,
            self.last_name,
            self.first_name,
            self.banner_id,
            self.email,
            self.status,
            self.semester_joined,
            self.position,
        ]


def clean_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    text = re.sub(r"\s+", " ", text)
    return text


def canonical_header(value: object) -> str:
    text = clean_text(value).lower()
    text = text.replace("_", " ")
    text = re.sub(r"[^a-z0-9 ]+", "", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalize_status(value: str) -> str:
    raw = clean_text(value)
    upper = raw.upper()
    if upper in STATUS_MAP:
        return STATUS_MAP[upper]
    return raw


def infer_year(path: Path) -> str:
    for part in path.parts:
        if re.fullmatch(r"(19|20)\d{2}", part):
            return part
    match = re.search(r"(19|20)\d{2}", path.stem)
    if match:
        return match.group(0)
    return "Unknown"


def infer_chapter(path: Path, sheet_name: str) -> str:
    parent = path.parent.name
    stem = path.stem
    if re.fullmatch(r"(19|20)\d{2}", parent):
        return stem
    if parent.lower() not in {"", "raw_rosters", "rosters"}:
        return parent
    return stem if stem.lower() != sheet_name.lower() else ""


def score_header_row(values: List[object]) -> Tuple[int, Dict[str, int]]:
    matched: Dict[str, int] = {}
    canon = [canonical_header(v) for v in values]
    for idx, header in enumerate(canon):
        for standard_name, aliases in HEADER_ALIASES.items():
            if header in aliases and standard_name not in matched:
                matched[standard_name] = idx
    return len(matched), matched


def find_header_row(ws) -> Tuple[Optional[int], Dict[str, int]]:
    best_score = 0
    best_row_idx = None
    best_map: Dict[str, int] = {}
    for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=min(ws.max_row, 20), values_only=True), start=1):
        score, header_map = score_header_row(list(row))
        if score > best_score:
            best_score = score
            best_row_idx = row_idx
            best_map = header_map
    required = {"last_name", "first_name"}
    if best_row_idx is None or best_score < 3 or not required.issubset(best_map):
        return None, {}
    return best_row_idx, best_map


def get_cell(row: Tuple[object, ...], index: Optional[int]) -> str:
    if index is None:
        return ""
    if index >= len(row):
        return ""
    return clean_text(row[index])


def row_is_empty(values: Iterable[str]) -> bool:
    return all(not clean_text(v) for v in values)


def extract_rows_from_workbook(path: Path, verbose: bool = False) -> Tuple[List[ExtractedRow], List[str]]:
    rows: List[ExtractedRow] = []
    issues: List[str] = []

    try:
        wb = load_workbook(path, data_only=True, read_only=True)
    except Exception as exc:
        issues.append(f"FAILED to open {path}: {exc}")
        return rows, issues

    year = infer_year(path)

    for ws in wb.worksheets:
        header_row_idx, header_map = find_header_row(ws)
        if header_row_idx is None:
            issues.append(f"Skipped {path.name} | sheet '{ws.title}': no usable header row found.")
            continue

        for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
            last_name = get_cell(row, header_map.get("last_name"))
            first_name = get_cell(row, header_map.get("first_name"))
            banner_id = get_cell(row, header_map.get("banner_id"))
            email = get_cell(row, header_map.get("email"))
            status = normalize_status(get_cell(row, header_map.get("status")))
            semester_joined = get_cell(row, header_map.get("semester_joined"))
            position = get_cell(row, header_map.get("position"))
            chapter = get_cell(row, header_map.get("chapter")) or infer_chapter(path, ws.title)

            core_values = [last_name, first_name, banner_id, email, status, semester_joined, position, chapter]
            if row_is_empty(core_values):
                continue

            # Skip likely note/footer rows.
            if not last_name and not first_name:
                continue

            rows.append(
                ExtractedRow(
                    year=year,
                    source_file=str(path.name),
                    source_sheet=ws.title,
                    chapter=chapter,
                    last_name=last_name,
                    first_name=first_name,
                    banner_id=banner_id,
                    email=email,
                    status=status,
                    semester_joined=semester_joined,
                    position=position,
                )
            )

    if verbose:
        print(f"Processed {path}")
    return rows, issues


def dedupe_rows(rows: List[ExtractedRow]) -> List[ExtractedRow]:
    seen = set()
    deduped: List[ExtractedRow] = []
    for row in rows:
        key = tuple(v.strip().lower() for v in row.as_list())
        if key in seen:
            continue
        seen.add(key)
        deduped.append(row)
    return deduped


def autosize_columns(ws) -> None:
    max_widths = defaultdict(int)
    for row in ws.iter_rows(values_only=True):
        for idx, value in enumerate(row, start=1):
            length = len(clean_text(value))
            if length > max_widths[idx]:
                max_widths[idx] = length
    for idx, width in max_widths.items():
        ws.column_dimensions[get_column_letter(idx)].width = min(max(width + 2, 12), 28)


def style_header(ws) -> None:
    fill = PatternFill("solid", fgColor="1F4E78")
    font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = fill
        cell.font = font


def write_summary_sheet(wb: Workbook, rows: List[ExtractedRow], issues: List[str]) -> None:
    ws = wb.active
    ws.title = "Summary"
    ws.append(["Metric", "Value"])
    style_header(ws)

    by_year = defaultdict(int)
    with_banner = 0
    missing_banner = 0
    for row in rows:
        by_year[row.year] += 1
        if row.banner_id:
            with_banner += 1
        else:
            missing_banner += 1

    metrics = [
        ["Total extracted rows", len(rows)],
        ["Rows with Banner ID", with_banner],
        ["Rows missing Banner ID", missing_banner],
        ["Years found", len(by_year)],
    ]
    for item in metrics:
        ws.append(item)

    ws.append([])
    ws.append(["Year", "Row Count"])
    for cell in ws[ws.max_row]:
        cell.fill = PatternFill("solid", fgColor="D9EAF7")
        cell.font = Font(bold=True)
    for year in sorted(by_year.keys()):
        ws.append([year, by_year[year]])

    ws.append([])
    ws.append(["Import Issues"])
    ws[ws.max_row][0].font = Font(bold=True)
    if issues:
        for issue in issues:
            ws.append([issue])
    else:
        ws.append(["None"])
    autosize_columns(ws)


def write_year_sheets(wb: Workbook, rows: List[ExtractedRow], chunk_size: int = 1000) -> None:
    grouped = defaultdict(list)
    for row in rows:
        grouped[row.year].append(row)

    for year in sorted(grouped.keys()):
        year_rows = grouped[year]
        year_rows.sort(key=lambda r: (
            r.banner_id or "ZZZZZZZZ",
            r.last_name.lower(),
            r.first_name.lower(),
            r.chapter.lower(),
            r.source_file.lower(),
            r.source_sheet.lower(),
        ))

        for start in range(0, len(year_rows), chunk_size):
            end = min(start + chunk_size, len(year_rows))
            label_start = start + 1
            label_end = end
            sheet_name = f"{year}_{label_start:04d}_{label_end:04d}"
            ws = wb.create_sheet(title=sheet_name[:31])
            ws.append(STANDARD_COLUMNS)
            style_header(ws)
            for row in year_rows[start:end]:
                ws.append(row.as_list())
            ws.freeze_panes = "A2"
            autosize_columns(ws)


def build_master_roster(input_root: Path, output_file: Path, chunk_size: int, keep_duplicates: bool, verbose: bool) -> None:
    all_rows: List[ExtractedRow] = []
    issues: List[str] = []

    files = sorted([p for p in input_root.rglob("*") if p.suffix.lower() in SUPPORTED_EXTENSIONS])
    if not files:
        raise FileNotFoundError(
            f"No Excel files found under {input_root}. Supported types: {', '.join(sorted(SUPPORTED_EXTENSIONS))}"
        )

    for path in files:
        extracted, file_issues = extract_rows_from_workbook(path, verbose=verbose)
        all_rows.extend(extracted)
        issues.extend(file_issues)

    if not keep_duplicates:
        all_rows = dedupe_rows(all_rows)

    wb = Workbook()
    write_summary_sheet(wb, all_rows, issues)
    write_year_sheets(wb, all_rows, chunk_size=chunk_size)
    wb.save(output_file)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Build a single FSL master roster workbook from year-by-year folders of chapter rosters. "
            "The output workbook creates sheets separated by year and chunked into groups of 1000 rows."
        )
    )
    parser.add_argument("input_root", help="Root folder containing year folders of roster Excel files.")
    parser.add_argument(
        "-o",
        "--output",
        default="Master_FSL_Roster.xlsx",
        help="Output workbook path. Default: Master_FSL_Roster.xlsx",
    )
    parser.add_argument(
        "--chunk-size",
        type=int,
        default=1000,
        help="Number of rows per year sheet. Default: 1000",
    )
    parser.add_argument(
        "--keep-duplicates",
        action="store_true",
        help="Keep exact duplicate rows instead of removing them.",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Print each workbook as it is processed.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    input_root = Path(args.input_root).expanduser().resolve()
    output_file = Path(args.output).expanduser().resolve()

    build_master_roster(
        input_root=input_root,
        output_file=output_file,
        chunk_size=args.chunk_size,
        keep_duplicates=args.keep_duplicates,
        verbose=args.verbose,
    )
    print(f"Master roster created: {output_file}")


if __name__ == "__main__":
    main()
