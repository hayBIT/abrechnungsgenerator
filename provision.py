#!/usr/bin/env python3
import argparse
import csv
import re
import zipfile
from collections import Counter
from pathlib import Path
from typing import Dict, Iterable, List, Tuple
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


def sanitize_filename(value: str) -> str:
    cleaned = re.sub(r"[^\w\.-]+", "_", value.strip(), flags=re.UNICODE)
    return cleaned or "UNMATCHED"


def build_ods_content_xml(sheet_name: str, fieldnames: List[str], rows: List[Dict[str, str]]) -> str:
    def cell(value: str) -> str:
        escaped = escape(value)
        return (
            "<table:table-cell office:value-type=\"string\">"
            f"<text:p>{escaped}</text:p>"
            "</table:table-cell>"
        )

    header_cells = "".join(cell(name) for name in fieldnames)
    row_xml = f"<table:table-row>{header_cells}</table:table-row>"
    rows_xml = []
    for row in rows:
        cells = "".join(cell(str(row.get(name, ""))) for name in fieldnames)
        rows_xml.append(f"<table:table-row>{cells}</table:table-row>")
    table_rows = row_xml + "".join(rows_xml)
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    office:version="1.2">
  <office:body>
    <office:spreadsheet>
      <table:table table:name="{escape(sheet_name)}">
        {table_rows}
      </table:table>
    </office:spreadsheet>
  </office:body>
</office:document-content>
"""


def write_ods(path: Path, sheet_name: str, fieldnames: List[str], rows: List[Dict[str, str]]) -> None:
    content_xml = build_ods_content_xml(sheet_name, fieldnames, rows)
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


def format_amount_eur(value: str) -> str:
    cleaned = value.strip()
    if not cleaned:
        return ""
    if cleaned.endswith("€"):
        return cleaned
    return f"{cleaned} €"


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
                "beg_wirk_dat": row.get("Fälligkeit", ""),
                "Betrag": amount,
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
        "--output-dir", type=Path, default=Path("output"), help="Directory for outputs"
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    ameise_map = load_ameise(args.ameise)

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

    for insurer, paths in (("KRAVAG", args.kravag), ("R+V", args.rv)):
        for path in paths:
            matched, counts, missing_fields = match_insurer(insurer, path, ameise_map)
            all_rows.extend(matched)
            summary.update(counts)
            if missing_fields:
                missing_columns.setdefault(insurer, [])
                for field in missing_fields:
                    if field not in missing_columns[insurer]:
                        missing_columns[insurer].append(field)

    for path in args.vema:
        matched, counts, missing_fields = match_vema(path, ameise_map)
        all_rows.extend(matched)
        summary.update(counts)
        if missing_fields:
            missing_columns.setdefault("VEMA", [])
            for field in missing_fields:
                if field not in missing_columns["VEMA"]:
                    missing_columns["VEMA"].append(field)

    if missing_columns:
        for insurer, fields in missing_columns.items():
            missing = ", ".join(fields)
            raise SystemExit(f"Missing required columns in {insurer} CSV: {missing}")

    if not all_rows:
        raise SystemExit("No insurer rows processed. Provide at least one settlement CSV.")

    fieldnames: List[str] = []
    for row in all_rows:
        for key in row.keys():
            if key not in fieldnames:
                fieldnames.append(key)
    write_csv(args.output_dir / "matched_rows.csv", fieldnames, all_rows)
    unmatched_rows = [row for row in all_rows if row.get("match_status") == "unmatched"]
    write_csv(args.output_dir / "unmatched_rows.csv", fieldnames, unmatched_rows)

    summary_rows = [
        {"VMT": vmt, "count": str(count)}
        for vmt, count in sorted(summary.items(), key=lambda item: item[0])
    ]
    write_csv(args.output_dir / "summary_by_vmt.csv", ["VMT", "count"], summary_rows)
    unmatched_ods = args.output_dir / "vmt" / f"{sanitize_filename('UNMATCHED')}.ods"
    if unmatched_ods.exists():
        unmatched_ods.unlink()

    rows_by_vmt: Dict[str, List[Dict[str, str]]] = {}
    for row in all_rows:
        vmt = row.get("VMT") or "UNMATCHED"
        rows_by_vmt.setdefault(vmt, []).append(row)

    for vmt, rows in rows_by_vmt.items():
        if vmt == "UNMATCHED":
            continue
        filename = f"{sanitize_filename(vmt)}.ods"
        ods_rows = []
        for row in rows:
            ods_row = dict(row)
            ods_row["Datum"] = row.get("beg_wirk_dat", "")
            ods_row["Betrag"] = row.get("abrechnungsbetrag", "")
            ods_rows.append(ods_row)
        write_ods(args.output_dir / "vmt" / filename, vmt, ods_fieldnames, ods_rows)


if __name__ == "__main__":
    main()
