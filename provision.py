#!/usr/bin/env python3
import argparse
import csv
import re
import zipfile
from collections import Counter
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, Iterable, List, Tuple
from xml.etree import ElementTree
from xml.sax.saxutils import escape


def normalize_vsn(vsn: str) -> str:
    cleaned = vsn.strip()
    if not cleaned:
        return ""
    for dash in ("−", "–", "—", "‑", "‒", "﹣", "－"):
        cleaned = cleaned.replace(dash, "-")
    parts = [part for part in re.split(r"[-\s]+", cleaned) if part]
    if not parts:
        return ""
    if len(parts) == 1:
        return parts[0]
    ag = parts[0]
    vsnr = "".join(parts[1:])
    if vsnr.isdigit() and len(vsnr) < 9:
        vsnr = vsnr.zfill(9)
    return f"{ag}-{vsnr}"


def build_vsn(ag: str, vsnr: str) -> str:
    ag_clean = ag.strip()
    vsnr_clean = vsnr.strip()
    if vsnr_clean.isdigit() and len(vsnr_clean) < 9:
        vsnr_clean = vsnr_clean.zfill(9)
    return normalize_vsn(f"{ag_clean}-{vsnr_clean}")


def detect_dialect(path: Path) -> csv.Dialect:
    sample = path.read_text(encoding="utf-8-sig")[:4096]
    sniffer = csv.Sniffer()
    try:
        return sniffer.sniff(sample)
    except csv.Error:
        return csv.excel


def read_csv(path: Path) -> Tuple[List[Dict[str, str]], List[str]]:
    dialect = detect_dialect(path)
    with path.open(encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle, dialect=dialect)
        rows = [row for row in reader]
    return rows, reader.fieldnames or []


def _xlsx_column_index(cell_ref: str) -> int:
    letters = ""
    for char in cell_ref:
        if char.isalpha():
            letters += char.upper()
        else:
            break
    index = 0
    for char in letters:
        index = index * 26 + (ord(char) - ord("A") + 1)
    return index - 1


def _xlsx_text(element: ElementTree.Element) -> str:
    if element is None:
        return ""
    if element.text:
        return element.text
    return "".join(child.text or "" for child in element.findall(".//{*}t"))


def read_xlsx(path: Path) -> Tuple[List[Dict[str, str]], List[str]]:
    shared_strings: List[str] = []
    with zipfile.ZipFile(path) as archive:
        if "xl/sharedStrings.xml" in archive.namelist():
            shared_tree = ElementTree.fromstring(archive.read("xl/sharedStrings.xml"))
            for shared in shared_tree.findall(".//{*}si"):
                shared_strings.append(_xlsx_text(shared))

        sheet_path = "xl/worksheets/sheet1.xml"
        if "xl/workbook.xml" in archive.namelist() and "xl/_rels/workbook.xml.rels" in archive.namelist():
            workbook_tree = ElementTree.fromstring(archive.read("xl/workbook.xml"))
            rels_tree = ElementTree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
            sheet = workbook_tree.find(".//{*}sheet")
            if sheet is not None:
                rel_id = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                if rel_id:
                    rel = rels_tree.find(f".//{{*}}Relationship[@Id='{rel_id}']")
                    if rel is not None:
                        target = rel.attrib.get("Target", "")
                        if target:
                            sheet_path = target.lstrip("/")
                            if not sheet_path.startswith("xl/"):
                                sheet_path = f"xl/{sheet_path}"

        sheet_tree = ElementTree.fromstring(archive.read(sheet_path))

    rows: List[Dict[str, str]] = []
    headers: List[str] = []
    for row in sheet_tree.findall(".//{*}sheetData/{*}row"):
        values: Dict[int, str] = {}
        for cell in row.findall("{*}c"):
            cell_ref = cell.attrib.get("r", "")
            index = _xlsx_column_index(cell_ref)
            cell_type = cell.attrib.get("t")
            raw_value = _xlsx_text(cell.find("{*}v")) if cell_type != "inlineStr" else _xlsx_text(cell)
            if cell_type == "s":
                try:
                    value = shared_strings[int(raw_value)]
                except (ValueError, IndexError):
                    value = raw_value
            else:
                value = raw_value
            values[index] = value or ""
        if not headers:
            max_index = max(values.keys(), default=-1)
            headers = [values.get(idx, "").strip() for idx in range(max_index + 1)]
            continue
        if not headers:
            continue
        row_dict = {headers[idx]: values.get(idx, "").strip() for idx in range(len(headers))}
        rows.append(row_dict)
    return rows, headers


def write_csv(path: Path, fieldnames: Iterable[str], rows: Iterable[Dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def load_ameise(path: Path) -> Dict[str, Dict[str, str]]:
    rows, _ = read_csv(path)
    mapping: Dict[str, Dict[str, str]] = {}
    for row in rows:
        vsn = normalize_vsn(row.get("VSN", ""))
        if not vsn:
            continue
        mapping[vsn] = {
            "VMT": row.get("VMT", "").strip(),
            "Vorname / Ansprechpartner": row.get("Vorname / Ansprechpartner", "").strip(),
            "Nachname / Firma": row.get("Nachname / Firma", "").strip(),
            "Gesellschaft": row.get("Gesellschaft", "").strip(),
            "Sparte": row.get("Sparte", "").strip(),
        }
    return mapping


def parse_rate(value: str) -> float | None:
    cleaned = value.strip()
    if not cleaned:
        return None
    cleaned = cleaned.replace("%", "").replace("\u00a0", "").replace(" ", "")
    if cleaned.count(",") == 1 and "." in cleaned:
        cleaned = cleaned.replace(".", "").replace(",", ".")
    elif cleaned.count(",") == 1 and "." not in cleaned:
        cleaned = cleaned.replace(",", ".")
    try:
        rate = float(cleaned)
    except ValueError:
        return None
    if rate > 1:
        return rate / 100
    return rate


def load_vermittlerliste(
    path: Path,
) -> Dict[str, Dict[str, str | float | None]]:
    rows, _ = read_csv(path)
    mapping: Dict[str, Dict[str, str | float | None]] = {}
    for row in rows:
        vmt = row.get("VMT", "").strip()
        if not vmt:
            continue
        mapping[vmt] = {
            "Name": row.get("Name", "").strip(),
            "Provisionssatz": parse_rate(row.get("Provisionssatz", "")),
        }
    return mapping


def sanitize_filename(value: str) -> str:
    cleaned = re.sub(r"[^\w\.-]+", "_", value.strip(), flags=re.UNICODE)
    return cleaned or "UNMATCHED"


def build_ods_content_xml(
    sheet_name: str,
    fieldnames: List[str],
    rows: List[Dict[str, str]],
    date_fields: Iterable[str],
    currency_fields: Iterable[str],
) -> str:
    date_field_set = {name for name in date_fields}
    currency_field_set = {name for name in currency_fields}

    def cell(value: str, is_date: bool, is_currency: bool) -> str:
        if is_date:
            display, iso_value = normalize_date(value)
            if iso_value:
                escaped = escape(display)
                return (
                    "<table:table-cell table:style-name=\"DateCell\" office:value-type=\"date\" "
                    f"office:date-value=\"{iso_value}\">"
                    f"<text:p>{escaped}</text:p>"
                    "</table:table-cell>"
                )
            value = display
        if is_currency:
            amount = parse_amount_eur(value)
            if amount is not None:
                escaped = escape(format_amount_display(amount))
                return (
                    "<table:table-cell table:style-name=\"CurrencyCell\" "
                    "office:value-type=\"currency\" office:currency=\"EUR\" "
                    f"office:value=\"{amount}\">"
                    f"<text:p>{escaped}</text:p>"
                    "</table:table-cell>"
                )
        escaped = escape(value)
        return (
            "<table:table-cell office:value-type=\"string\">"
            f"<text:p>{escaped}</text:p>"
            "</table:table-cell>"
        )

    header_cells = "".join(cell(name, False, False) for name in fieldnames)
    row_xml = f"<table:table-row>{header_cells}</table:table-row>"
    rows_xml = []
    for row in rows:
        cells = "".join(
            cell(
                str(row.get(name, "")),
                name in date_field_set,
                name in currency_field_set,
            )
            for name in fieldnames
        )
        rows_xml.append(f"<table:table-row>{cells}</table:table-row>")
    table_rows = row_xml + "".join(rows_xml)
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
    xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
    office:version="1.2">
  <office:automatic-styles>
    <style:style style:name="DateCell" style:family="table-cell" style:data-style-name="date1"/>
    <number:date-style style:name="date1" number:automatic-order="true">
      <number:day number:style="long"/>
      <number:text>.</number:text>
      <number:month number:style="long"/>
      <number:text>.</number:text>
      <number:year number:style="long"/>
    </number:date-style>
    <style:style style:name="CurrencyCell" style:family="table-cell" style:data-style-name="currency1"/>
    <number:currency-style style:name="currency1">
      <number:number number:decimal-places="2" number:grouping="true" number:min-integer-digits="1"/>
      <number:text> </number:text>
      <number:currency-symbol>€</number:currency-symbol>
    </number:currency-style>
  </office:automatic-styles>
  <office:body>
    <office:spreadsheet>
      <table:table table:name="{escape(sheet_name)}">
        {table_rows}
      </table:table>
    </office:spreadsheet>
  </office:body>
</office:document-content>
"""


def write_ods(
    path: Path,
    sheet_name: str,
    fieldnames: List[str],
    rows: List[Dict[str, str]],
    date_fields: Iterable[str],
    currency_fields: Iterable[str],
) -> None:
    content_xml = build_ods_content_xml(
        sheet_name, fieldnames, rows, date_fields, currency_fields
    )
    manifest_xml = """<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest
    xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"
    manifest:version="1.2">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
</manifest:manifest>
"""
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr(
            "mimetype",
            "application/vnd.oasis.opendocument.spreadsheet",
            compress_type=zipfile.ZIP_STORED,
        )
        archive.writestr("content.xml", content_xml, compress_type=zipfile.ZIP_DEFLATED)
        archive.writestr("META-INF/manifest.xml", manifest_xml, compress_type=zipfile.ZIP_DEFLATED)


def _pdf_escape(text: str) -> str:
    escaped = text.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")
    return escaped.encode("cp1252", errors="replace").decode("cp1252")


def _format_pdf_cell(value: str, width: int, align_right: bool = False) -> str:
    trimmed = value.strip()
    if len(trimmed) > width:
        return trimmed[: max(0, width - 1)] + "…"
    if align_right:
        return trimmed.rjust(width)
    return trimmed.ljust(width)


def write_pdf(
    path: Path,
    title: str,
    fieldnames: List[str],
    rows: List[Dict[str, str]],
) -> None:
    page_width = 842
    page_height = 595
    margin_x = 40
    margin_y = 40
    font_size = 10
    line_height = 12
    lines_per_page = int((page_height - 2 * margin_y) / line_height)

    column_widths = [12, 16, 16, 12, 10, 10, 10]
    if len(fieldnames) != len(column_widths):
        column_widths = [max(8, 80 // max(1, len(fieldnames)))] * len(fieldnames)

    header_line = " | ".join(
        _format_pdf_cell(name, width) for name, width in zip(fieldnames, column_widths)
    )
    separator_line = "-+-".join("-" * width for width in column_widths)

    def row_line(row: Dict[str, str]) -> str:
        return " | ".join(
            _format_pdf_cell(
                str(row.get(name, "")),
                width,
                align_right=name == "Betrag",
            )
            for name, width in zip(fieldnames, column_widths)
        )

    data_lines = [row_line(row) for row in rows]

    pages: List[List[str]] = []
    current: List[str] = []
    for line in data_lines:
        if len(current) >= lines_per_page - 4:
            pages.append(current)
            current = []
        current.append(line)
    if current:
        pages.append(current)

    contents: List[str] = []
    for page_index, page_lines in enumerate(pages, start=1):
        lines: List[str] = []
        lines.append(title if len(pages) == 1 else f"{title} (Seite {page_index}/{len(pages)})")
        lines.append("")
        lines.append(header_line)
        lines.append(separator_line)
        lines.extend(page_lines)
        text_lines = [f"({_pdf_escape(line)}) Tj T*" for line in lines]
        text_stream = "\n".join(
            [
                "BT",
                f"/F1 {font_size} Tf",
                f"{margin_x} {page_height - margin_y} Td",
                f"{line_height} TL",
                *text_lines,
                "ET",
            ]
        )
        contents.append(text_stream)

    objects: List[bytes] = []

    def add_object(data: str) -> int:
        objects.append(data.encode("cp1252"))
        return len(objects)

    font_obj = add_object(
        "<< /Type /Font /Subtype /Type1 /Name /F1 /BaseFont /Courier "
        "/Encoding /WinAnsiEncoding >>"
    )

    content_obj_ids: List[int] = []
    for stream in contents:
        stream_bytes = stream.encode("cp1252")
        content_obj_ids.append(
            add_object(f"<< /Length {len(stream_bytes)} >>\nstream\n{stream}\nendstream")
        )

    pages_obj_id = add_object("<< /Type /Pages /Kids [] /Count 0 >>")

    page_obj_ids: List[int] = []
    for content_obj_id in content_obj_ids:
        page_obj_ids.append(
            add_object(
                f"<< /Type /Page /Parent {pages_obj_id} 0 R /MediaBox [0 0 {page_width} {page_height}] "
                f"/Resources << /Font << /F1 {font_obj} 0 R >> >> "
                f"/Contents {content_obj_id} 0 R >>"
            )
        )

    pages_kids = " ".join(f"{page_id} 0 R" for page_id in page_obj_ids)
    objects[pages_obj_id - 1] = (
        f"<< /Type /Pages /Kids [{pages_kids}] /Count {len(page_obj_ids)} >>"
    ).encode("cp1252")
    catalog_obj = add_object(f"<< /Type /Catalog /Pages {pages_obj_id} 0 R >>")

    xref_positions = []
    pdf_parts = [b"%PDF-1.4\n"]
    for idx, obj in enumerate(objects, start=1):
        xref_positions.append(sum(len(part) for part in pdf_parts))
        pdf_parts.append(f"{idx} 0 obj\n".encode("cp1252"))
        pdf_parts.append(obj)
        pdf_parts.append(b"\nendobj\n")
    xref_start = sum(len(part) for part in pdf_parts)
    xref_entries = ["0000000000 65535 f "]
    for pos in xref_positions:
        xref_entries.append(f"{pos:010d} 00000 n ")
    pdf_parts.append(f"xref\n0 {len(xref_entries)}\n".encode("cp1252"))
    pdf_parts.append("\n".join(xref_entries).encode("cp1252"))
    pdf_parts.append(
        f"\ntrailer\n<< /Size {len(xref_entries)} /Root {catalog_obj} 0 R >>\nstartxref\n{xref_start}\n%%EOF\n".encode(
            "cp1252"
        )
    )

    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(b"".join(pdf_parts))


def format_amount_eur(value: str) -> str:
    cleaned = value.strip()
    if not cleaned:
        return ""
    if cleaned.endswith("€"):
        return cleaned
    return f"{cleaned} €"


def parse_amount_eur(value: str) -> float | None:
    cleaned = value.strip()
    if not cleaned:
        return None
    cleaned = cleaned.replace("€", "").replace("\u00a0", "").replace(" ", "")
    if cleaned.count(",") == 1 and "." in cleaned:
        cleaned = cleaned.replace(".", "").replace(",", ".")
    elif cleaned.count(",") == 1 and "." not in cleaned:
        cleaned = cleaned.replace(",", ".")
    try:
        return float(cleaned)
    except ValueError:
        return None


def format_amount_display(amount: float) -> str:
    formatted = f"{amount:,.2f}"
    formatted = formatted.replace(",", "_").replace(".", ",").replace("_", ".")
    return f"{formatted} €"


def normalize_date(value: str) -> Tuple[str, str]:
    cleaned = value.strip()
    if not cleaned:
        return "", ""

    candidates = [cleaned]
    if "T" in cleaned:
        candidates.append(cleaned.split("T", 1)[0])
    if " " in cleaned:
        candidates.append(cleaned.split(" ", 1)[0])

    date_formats = [
        "%d.%m.%Y",
        "%d.%m.%y",
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%d/%m/%Y",
        "%d/%m/%y",
        "%m/%d/%Y",
        "%m/%d/%y",
    ]

    for candidate in candidates:
        for fmt in date_formats:
            try:
                parsed = datetime.strptime(candidate, fmt)
            except ValueError:
                continue
            return parsed.strftime("%d.%m.%Y"), parsed.strftime("%Y-%m-%d")

    if re.fullmatch(r"\d+(\.\d+)?", cleaned):
        try:
            serial = float(cleaned)
        except ValueError:
            serial = 0.0
        if serial:
            base_date = datetime(1899, 12, 30)
            parsed = base_date + timedelta(days=serial)
            return parsed.strftime("%d.%m.%Y"), parsed.strftime("%Y-%m-%d")

    return cleaned, ""


def match_insurer(
    insurer: str, path: Path, ameise_map: Dict[str, Dict[str, str]]
) -> Tuple[List[Dict[str, str]], Counter, List[str]]:
    rows, fieldnames = read_csv(path)
    matched_rows: List[Dict[str, str]] = []
    counts = Counter()
    missing_fields: List[str] = []

    if "ag" not in fieldnames or "vsnr" not in fieldnames:
        if "ag" not in fieldnames:
            missing_fields.append("ag")
        if "vsnr" not in fieldnames:
            missing_fields.append("vsnr")
        return matched_rows, counts, missing_fields

    for row in rows:
        vsn = build_vsn(row.get("ag", ""), row.get("vsnr", ""))
        ameise_details = ameise_map.get(vsn, {})
        vmt = ameise_details.get("VMT", "")
        status = "matched" if vmt else "unmatched"
        counts[vmt or "UNMATCHED"] += 1
        abrechnungsbetrag = format_amount_eur(row.get("abrechnungsbetrag", ""))
        enriched = dict(row)
        enriched.update(
            {
                "insurer": insurer,
                "VSN": vsn,
                "VMT": vmt,
                "Vorname / Ansprechpartner": ameise_details.get("Vorname / Ansprechpartner", ""),
                "Nachname / Firma": ameise_details.get("Nachname / Firma", ""),
                "Gesellschaft": ameise_details.get("Gesellschaft", ""),
                "Sparte": ameise_details.get("Sparte", ""),
                "abrechnungsbetrag": abrechnungsbetrag,
                "match_status": status,
            }
        )
        matched_rows.append(enriched)

    return matched_rows, counts, missing_fields


def match_vema(
    path: Path, ameise_map: Dict[str, Dict[str, str]]
) -> Tuple[List[Dict[str, str]], Counter, List[str]]:
    rows, fieldnames = read_csv(path)
    matched_rows: List[Dict[str, str]] = []
    counts = Counter()
    missing_fields: List[str] = []

    required_fields = ["Vertragsnummer", "Fälligkeit", "Betrag"]
    for field in required_fields:
        if field not in fieldnames:
            missing_fields.append(field)
    if missing_fields:
        return matched_rows, counts, missing_fields

    for row in rows:
        vsn = normalize_vsn(row.get("Vertragsnummer", ""))
        ameise_details = ameise_map.get(vsn, {})
        vmt = ameise_details.get("VMT", "")
        status = "matched" if vmt else "unmatched"
        counts[vmt or "UNMATCHED"] += 1
        amount = format_amount_eur(row.get("Betrag", ""))
        enriched = dict(row)
        date_display, _ = normalize_date(row.get("Fälligkeit", ""))
        enriched.update(
            {
                "insurer": "VEMA",
                "VSN": vsn,
                "VMT": vmt,
                "Vorname / Ansprechpartner": ameise_details.get("Vorname / Ansprechpartner", ""),
                "Nachname / Firma": ameise_details.get("Nachname / Firma", ""),
                "Gesellschaft": ameise_details.get("Gesellschaft", ""),
                "Sparte": ameise_details.get("Sparte", ""),
                "abrechnungsbetrag": amount,
                "beg_wirk_dat": date_display or row.get("Fälligkeit", ""),
                "Betrag": amount,
                "match_status": status,
            }
        )
        matched_rows.append(enriched)

    return matched_rows, counts, missing_fields


def match_fonds_finanz(
    path: Path, ameise_map: Dict[str, Dict[str, str]]
) -> Tuple[List[Dict[str, str]], Counter, List[str]]:
    rows, fieldnames = read_xlsx(path)
    matched_rows: List[Dict[str, str]] = []
    counts = Counter()
    missing_fields: List[str] = []

    required_fields = ["Abrechnungsdatum", "Vertragsnummer extern", "Summe in EUR"]
    for field in required_fields:
        if field not in fieldnames:
            missing_fields.append(field)
    if missing_fields:
        return matched_rows, counts, missing_fields

    for row in rows:
        vsn = normalize_vsn(row.get("Vertragsnummer extern", ""))
        ameise_details = ameise_map.get(vsn, {})
        vmt = ameise_details.get("VMT", "")
        status = "matched" if vmt else "unmatched"
        counts[vmt or "UNMATCHED"] += 1
        amount = format_amount_eur(row.get("Summe in EUR", ""))
        enriched = dict(row)
        date_display, _ = normalize_date(row.get("Abrechnungsdatum", ""))
        enriched.update(
            {
                "insurer": "Fonds Finanz",
                "VSN": vsn,
                "VMT": vmt,
                "Vorname / Ansprechpartner": ameise_details.get("Vorname / Ansprechpartner", ""),
                "Nachname / Firma": ameise_details.get("Nachname / Firma", ""),
                "Gesellschaft": ameise_details.get("Gesellschaft", ""),
                "Sparte": ameise_details.get("Sparte", ""),
                "abrechnungsbetrag": amount,
                "beg_wirk_dat": date_display or row.get("Abrechnungsdatum", ""),
                "Summe in EUR": amount,
                "match_status": status,
            }
        )
        matched_rows.append(enriched)

    return matched_rows, counts, missing_fields


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Match insurer settlement CSVs to Ameise contract list and assign sub-brokers.",
        epilog="Example: provision.py --ameise ameise.csv --kravag kravag.csv --vema vema.csv",
    )
    parser.add_argument("--ameise", required=True, type=Path, help="CSV export from Ameise CRM")
    parser.add_argument("--kravag", type=Path, nargs="*", default=[], help="KRAVAG settlement CSVs")
    parser.add_argument("--rv", type=Path, nargs="*", default=[], help="R+V settlement CSVs")
    parser.add_argument(
        "--vema",
        type=Path,
        nargs="*",
        default=[],
        help="VEMA settlement CSVs (optional)",
    )
    parser.add_argument(
        "--ff",
        type=Path,
        nargs="*",
        default=[],
        help="Fonds Finanz settlement XLSX files (optional)",
    )
    parser.add_argument(
        "--vermittlerliste",
        type=Path,
        help="Optional Vermittlerliste with columns VMT, Name, Provisionssatz",
    )
    parser.add_argument(
        "--output-dir", type=Path, default=Path("output"), help="Directory for outputs"
    )
    return parser.parse_args()


def run_provision(
    ameise: Path,
    kravag: Iterable[Path],
    rv: Iterable[Path],
    vema: Iterable[Path],
    ff: Iterable[Path],
    vermittlerliste: Path | None,
    output_dir: Path,
) -> None:
    ameise_map = load_ameise(ameise)
    vermittler_map = load_vermittlerliste(vermittlerliste) if vermittlerliste else {}
    date_prefix = datetime.now().strftime("%Y%m%d")

    ods_fieldnames = [
        "VSN",
        "Vorname / Ansprechpartner",
        "Nachname / Firma",
        "Gesellschaft",
        "Sparte",
        "Datum",
        "Betrag",
    ]

    all_rows: List[Dict[str, str]] = []
    summary = Counter()
    missing_columns: Dict[str, List[str]] = {}

    for insurer, paths in (("KRAVAG", kravag), ("R+V", rv)):
        for path in paths:
            matched, counts, missing_fields = match_insurer(insurer, path, ameise_map)
            all_rows.extend(matched)
            summary.update(counts)
            if missing_fields:
                missing_columns.setdefault(insurer, [])
                for field in missing_fields:
                    if field not in missing_columns[insurer]:
                        missing_columns[insurer].append(field)

    for path in vema:
        matched, counts, missing_fields = match_vema(path, ameise_map)
        all_rows.extend(matched)
        summary.update(counts)
        if missing_fields:
            missing_columns.setdefault("VEMA", [])
            for field in missing_fields:
                if field not in missing_columns["VEMA"]:
                    missing_columns["VEMA"].append(field)

    for path in ff:
        matched, counts, missing_fields = match_fonds_finanz(path, ameise_map)
        all_rows.extend(matched)
        summary.update(counts)
        if missing_fields:
            missing_columns.setdefault("FF", [])
            for field in missing_fields:
                if field not in missing_columns["FF"]:
                    missing_columns["FF"].append(field)

    if missing_columns:
        for insurer, fields in missing_columns.items():
            missing = ", ".join(fields)
            raise SystemExit(f"Missing required columns in {insurer} file: {missing}")

    if not all_rows:
        raise SystemExit("No insurer rows processed. Provide at least one settlement CSV.")

    fieldnames: List[str] = []
    for row in all_rows:
        for key in row.keys():
            if key not in fieldnames:
                fieldnames.append(key)
    write_csv(output_dir / "matched_rows.csv", fieldnames, all_rows)
    unmatched_rows = [row for row in all_rows if row.get("match_status") == "unmatched"]
    write_csv(output_dir / "unmatched_rows.csv", fieldnames, unmatched_rows)

    summary_rows = [
        {"VMT": vmt, "count": str(count)}
        for vmt, count in sorted(summary.items(), key=lambda item: item[0])
    ]
    write_csv(output_dir / "summary_by_vmt.csv", ["VMT", "count"], summary_rows)
    unmatched_ods = output_dir / "vmt" / f"{sanitize_filename('UNMATCHED')}.ods"
    if unmatched_ods.exists():
        unmatched_ods.unlink()

    rows_by_vmt: Dict[str, List[Dict[str, str]]] = {}
    for row in all_rows:
        vmt = row.get("VMT") or "UNMATCHED"
        rows_by_vmt.setdefault(vmt, []).append(row)

    for vmt, rows in rows_by_vmt.items():
        if vmt == "UNMATCHED":
            continue
        vermittler_info = vermittler_map.get(vmt, {})
        vermittler_name = vermittler_info.get("Name") or ""
        name_suffix = f"_{sanitize_filename(vermittler_name)}" if vermittler_name else ""
        filename_base = f"{date_prefix}_{sanitize_filename(vmt)}{name_suffix}"
        ods_rows = []
        total_amount = 0.0
        has_amount = False
        for row in rows:
            ods_row = dict(row)
            date_display, _ = normalize_date(row.get("beg_wirk_dat", ""))
            ods_row["Datum"] = date_display or row.get("beg_wirk_dat", "")
            amount = parse_amount_eur(row.get("abrechnungsbetrag", ""))
            provision_rate = vermittler_info.get("Provisionssatz")
            if amount is not None and isinstance(provision_rate, float):
                amount *= provision_rate
                ods_row["Betrag"] = format_amount_display(amount)
            else:
                ods_row["Betrag"] = row.get("abrechnungsbetrag", "")
                amount = parse_amount_eur(ods_row["Betrag"])
            if amount is not None:
                total_amount += amount
                has_amount = True
            ods_rows.append(ods_row)
        if has_amount:
            total_row = {field: "" for field in ods_fieldnames}
            total_row["Nachname / Firma"] = "Summe"
            total_row["Betrag"] = format_amount_display(total_amount)
            ods_rows.append(total_row)
        write_ods(
            output_dir / "vmt" / f"{filename_base}.ods",
            vmt,
            ods_fieldnames,
            ods_rows,
            date_fields=["Datum"],
            currency_fields=["Betrag"],
        )
        write_pdf(
            output_dir / "vmt" / f"{filename_base}.pdf",
            vmt,
            ods_fieldnames,
            ods_rows,
        )


def main() -> None:
    args = parse_args()
    run_provision(
        ameise=args.ameise,
        kravag=args.kravag,
        rv=args.rv,
        vema=args.vema,
        ff=args.ff,
        vermittlerliste=args.vermittlerliste,
        output_dir=args.output_dir,
    )


if __name__ == "__main__":
    main()
