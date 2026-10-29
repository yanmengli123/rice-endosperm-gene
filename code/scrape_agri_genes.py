#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Download and export the complete RAP-DB "Agri. genes" table.

The web page https://rapdb.dna.naro.go.jp/agri_genes/agri_gene_list.html
renders its table client-side from a single JSON endpoint:
    https://rapdb.dna.naro.go.jp/agri_genes/agri_genes.json

The page's "Per page" selector only paginates the browser view (10/20/50/100);
the JSON endpoint returns every record at once, so no 11-page copy/paste loop is
needed and nothing is lost to manual selection.

Outputs (written next to this script, i.e. ...\\RAP-DB\\Agri. genes):
    agri_genes.json                       raw source payload
    RAPDB_Agri_Genes.xlsx                 formatted workbook (all records)
    RAPDB_Agri_Genes.csv                  same main table, UTF-8 with BOM
    RAPDB_Agri_Genes_references.csv       one row per transcript-reference
    RAPDB_Agri_Genes_endosperm_hits.xlsx  endosperm keyword preview subset

Usage:
    python scrape_agri_genes.py              # download fresh data
    python scrape_agri_genes.py --use-local  # reuse agri_genes.json
"""

from __future__ import annotations

import argparse
import datetime as dt
import html
import json
import re
import sys
import time
from pathlib import Path
from typing import Any, Iterable

import pandas as pd
import requests
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

DATA_ROOT = Path(__file__).resolve().parent.parent
BASE_DIR = DATA_ROOT / "RAP-DB" / "Agri. genes"
PAGE_URL = "https://rapdb.dna.naro.go.jp/agri_genes/agri_gene_list.html"
DATA_URL = "https://rapdb.dna.naro.go.jp/agri_genes/agri_genes.json"
SITE_ROOT = "https://rapdb.dna.naro.go.jp"

RAW_JSON = BASE_DIR / "agri_genes.json"
XLSX_PATH = BASE_DIR / "RAPDB_Agri_Genes.xlsx"
CSV_PATH = BASE_DIR / "RAPDB_Agri_Genes.csv"
REF_CSV_PATH = BASE_DIR / "RAPDB_Agri_Genes_references.csv"
HITS_XLSX_PATH = BASE_DIR / "RAPDB_Agri_Genes_endosperm_hits.xlsx"

# Number of transcript records the site itself reports for this release.
EXPECTED_ROWS = 1093

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    ),
    "Accept": "application/json, text/plain, */*",
    "Referer": PAGE_URL,
}

# Conservative endosperm / grain-filling keyword set, matched case-insensitively
# against gene names, gene symbols, trait ontology terms, PO terms and paper titles.
ENDOSPERM_PATTERN = re.compile(
    r"\b("
    r"endosperm\w*|amylose|amylopectin|starch\w*|waxy|wx|"
    r"glutelin\w*|prolamin\w*|globulin\w*|aleurone|"
    r"storage protein\w*|grain filling|chalkiness|chalky|"
    r"amylase\w*|sucrose synthase|adp-glucose|granule-bound"
    r")\b",
    re.IGNORECASE,
)

CONTROL_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
WHITESPACE = re.compile(r"[ \t\r\f\v]+")
# "TO:0000074 - blast disease" -> "blast disease (TO:0000074)", matching the site.
ONTOLOGY_TERM = re.compile(r"^(TO|PO):(\d+)\s*-\s*(.+)$")


def clean(value: Any) -> str:
    """HTML-unescape a cell value and make it safe/compact for Excel."""
    if value is None:
        return ""
    text = html.unescape(str(value))
    text = CONTROL_CHARS.sub("", text)
    text = WHITESPACE.sub(" ", text)
    return text.strip()


def format_ontology(values: Iterable[str] | None) -> str:
    """Return 'term (TO:0000074); term2 (...)' for a list of ontology strings."""
    if not values:
        return ""
    formatted: list[str] = []
    for raw in values:
        term = clean(raw)
        match = ONTOLOGY_TERM.match(term)
        if match:
            prefix, number, label = match.groups()
            term = f"{label.strip()} ({prefix}:{number})"
        if term:
            formatted.append(term)
    return "; ".join(formatted)


def format_reference(ref: dict[str, Any] | None) -> str:
    """One-line citation, mirroring the 'References' column of the web table."""
    if not isinstance(ref, dict):
        return ""
    authors = clean(ref.get("authorText") or ref.get("firstAuthor"))
    source = clean(ref.get("source"))
    volume = clean(ref.get("volume"))
    issue = clean(ref.get("issue"))
    pages = clean(ref.get("pages"))
    year = clean(ref.get("pubDateYear"))
    lit_type = clean(ref.get("literatureType"))

    citation = ""
    if authors:
        citation += f"{authors} "
    if source:
        citation += f"{source}, "
    if volume:
        citation += volume
    if issue:
        citation += f"({issue})"
    if volume or issue:
        citation += ":"
    if pages:
        citation += pages
    if year:
        citation += f" ({year})"
    if lit_type and lit_type.lower() != "none":
        citation += f" [{lit_type}]"
    return citation.strip(" ,")



def download_json(force: bool = False) -> list[dict[str, Any]]:
    """Fetch the complete agri-gene payload, with retries, cached on disk."""
    if RAW_JSON.exists() and not force:
        print(f"[info] reusing cached payload: {RAW_JSON}")
        with RAW_JSON.open(encoding="utf-8") as handle:
            data = json.load(handle)
    else:
        last_error: Exception | None = None
        for attempt in range(1, 6):
            try:
                print(f"[info] GET {DATA_URL} (attempt {attempt}/5)")
                response = requests.get(DATA_URL, headers=HEADERS, timeout=120)
                response.raise_for_status()
                data = response.json()
                RAW_JSON.write_text(
                    json.dumps(data, ensure_ascii=False), encoding="utf-8"
                )
                print(f"[ok] saved raw payload -> {RAW_JSON.name}")
                break
            except Exception as exc:  # noqa: BLE001 - retry on any transport error
                last_error = exc
                print(f"[warn] {exc}")
                time.sleep(2 * attempt)
        else:
            raise SystemExit(f"[error] download failed: {last_error}")

    if not isinstance(data, list) or not data:
        raise SystemExit("[error] unexpected payload: expected a non-empty JSON list")

    print(f"[ok] records in payload: {len(data):,}")
    if len(data) != EXPECTED_ROWS:
        print(
            f"[warn] expected {EXPECTED_ROWS:,} transcript records, got {len(data):,} "
            "- RAP-DB may have released a new curation."
        )
    return data


def build_tables(records: list[dict[str, Any]]) -> dict[str, pd.DataFrame]:
    """Convert the raw JSON into the main table plus flattened side tables."""
    main_rows: list[dict[str, Any]] = []
    ref_rows: list[dict[str, Any]] = []
    to_rows: list[dict[str, Any]] = []
    po_rows: list[dict[str, Any]] = []

    for index, record in enumerate(records, start=1):
        locus = clean(record.get("locus"))
        transcript = clean(record.get("transcript"))
        seqid = clean(record.get("seqid"))
        start = record.get("start")
        end = record.get("end")
        position = f"{seqid}:{start}..{end}" if seqid and start and end else ""

        to_terms = record.get("to") or []
        po_terms = record.get("po") or []
        references_raw = record.get("references")
        references = references_raw.values() if isinstance(references_raw, dict) else []
        references = [ref for ref in references if isinstance(ref, dict)]

        citations = [format_reference(ref) for ref in references]
        citations = [citation for citation in citations if citation]
        pubmed_ids = [
            clean(ref.get("linkTo")).rstrip("/").split("/")[-1]
            for ref in references
            if ref.get("linkTo")
        ]

        main_rows.append(
            {
                "No.": index,
                "Locus ID": locus,
                "Transcript ID": transcript,
                "Position": position,
                "Chromosome": seqid,
                "Start": start,
                "End": end,
                "Strand": clean(record.get("strand")),
                "Gene symbols": clean(record.get("gene_symbols")),
                "Gene names": clean(record.get("gene_names")),
                "Trait Ontology (TO)": format_ontology(to_terms),
                "Plant Ontology (PO)": format_ontology(po_terms),
                "Related traits (in Japanese)": clean(record.get("traits")),
                "References": "\n".join(citations),
                "No. of references": len(citations),
                "PubMed IDs": "; ".join(pubmed_ids),
                "Locus URL": f"{SITE_ROOT}/locus/?name={locus}" if locus else "",
                "Transcript URL": (
                    f"{SITE_ROOT}/transcript/?name={transcript}" if transcript else ""
                ),
                "JBrowse URL": (
                    f"{SITE_ROOT}/jbrowse/?data=data%2Firgsp1&loc={seqid}%3A{start}..{end}"
                    if position
                    else ""
                ),
            }
        )

        for ref in references:
            pubmed_id = (
                clean(ref.get("linkTo")).rstrip("/").split("/")[-1]
                if ref.get("linkTo")
                else ""
            )
            ref_rows.append(
                {
                    "Locus ID": locus,
                    "Transcript ID": transcript,
                    "PubMed ID": pubmed_id,
                    "Literature type": clean(ref.get("literatureType")),
                    "Title": clean(ref.get("title") or ref.get("planeText")),
                    "First author": clean(ref.get("firstAuthor")),
                    "Author text": clean(ref.get("authorText")),
                    "Journal": clean(ref.get("source")),
                    "Volume": clean(ref.get("volume")),
                    "Issue": clean(ref.get("issue")),
                    "Pages": clean(ref.get("pages")),
                    "Year": clean(ref.get("pubDateYear")),
                    "Link": clean(ref.get("linkTo")),
                    "Citation": format_reference(ref),
                }
            )

        for term in to_terms:
            text = clean(term)
            match = ONTOLOGY_TERM.match(text)
            to_rows.append(
                {
                    "Locus ID": locus,
                    "Transcript ID": transcript,
                    "TO ID": f"TO:{match.group(2)}" if match else "",
                    "TO term": match.group(3).strip() if match else text,
                }
            )

        for term in po_terms:
            text = clean(term)
            match = ONTOLOGY_TERM.match(text)
            po_rows.append(
                {
                    "Locus ID": locus,
                    "Transcript ID": transcript,
                    "PO ID": f"PO:{match.group(2)}" if match else "",
                    "PO term": match.group(3).strip() if match else text,
                }
            )

    return {
        "main": pd.DataFrame(main_rows),
        "references": pd.DataFrame(ref_rows),
        "to": pd.DataFrame(to_rows),
        "po": pd.DataFrame(po_rows),
    }


def find_endosperm_hits(
    main_df: pd.DataFrame, ref_df: pd.DataFrame
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Keyword preview of endosperm / storage-compound related agri. genes."""
    gene_hits: list[dict[str, Any]] = []

    for row in main_df.to_dict("records"):
        gene_text = " ".join(
            str(row.get(column) or "")
            for column in (
                "Gene symbols",
                "Gene names",
                "Trait Ontology (TO)",
                "Plant Ontology (PO)",
            )
        )
        tokens = {match.lower() for match in ENDOSPERM_PATTERN.findall(gene_text)}
        if not tokens:
            continue
        transcript = row.get("Transcript ID", "")
        gene_hits.append(
            {
                "Locus ID": row.get("Locus ID", ""),
                "Transcript ID": transcript,
                "Gene symbols": row.get("Gene symbols", ""),
                "Gene names": row.get("Gene names", ""),
                "Trait Ontology (TO)": row.get("Trait Ontology (TO)", ""),
                "Plant Ontology (PO)": row.get("Plant Ontology (PO)", ""),
                "Related traits (in Japanese)": row.get(
                    "Related traits (in Japanese)", ""
                ),
                "Matched keywords": "; ".join(sorted(tokens)),
            }
        )

    ref_hits: list[dict[str, Any]] = []
    if not ref_df.empty:
        for row in ref_df.to_dict("records"):
            title = str(row.get("Title") or "")
            tokens = {match.lower() for match in ENDOSPERM_PATTERN.findall(title)}
            if tokens:
                ref_hits.append(
                    {
                        "Locus ID": row.get("Locus ID", ""),
                        "Transcript ID": row.get("Transcript ID", ""),
                        "PubMed ID": row.get("PubMed ID", ""),
                        "Title": title,
                        "Journal": row.get("Journal", ""),
                        "Year": row.get("Year", ""),
                        "Matched keywords": "; ".join(sorted(tokens)),
                        "Link": row.get("Link", ""),
                    }
                )

    return pd.DataFrame(gene_hits), pd.DataFrame(ref_hits)


def build_readme(meta: dict[str, Any]) -> pd.DataFrame:
    """Two-column provenance / dictionary sheet."""
    rows = [
        ("RAP-DB Agri. genes - complete export", ""),
        ("Source page", PAGE_URL),
        ("Data endpoint used", DATA_URL),
        ("Retrieved (local time)", meta["retrieved"]),
        ("Transcript records", meta["records"]),
        ("Unique loci", meta["loci"]),
        ("Reference links (transcript x paper)", meta["refs"]),
        ("Unique PubMed / paper records", meta["unique_refs"]),
        ("Records with gene symbols", meta["with_symbols"]),
        ("Records with gene names", meta["with_names"]),
        ("Records with trait ontology (TO)", meta["with_to"]),
        ("Records with plant ontology (PO)", meta["with_po"]),
        ("Records with related traits (Japanese)", meta["with_traits"]),
        ("", ""),
        (
            "Pagination note",
            "The page's Per page selector (10/20/50/100 rows) only paginates the browser "
            "view; it is not a server-side limit. The table is rendered from "
            "/agri_genes/agri_genes.json, which returns all records in one response, "
            "so every transcript was captured without an 11-page copy/paste loop.",
        ),
        (
            "Text handling",
            "Gene symbols / names were HTML-unescaped (e.g. &alpha; -> Greek alpha, "
            "&beta; -> Greek beta) and stripped of control characters so they display "
            "correctly in Excel.",
        ),
        (
            "Sheet 'Agri. genes'",
            "One row per transcript, mirroring the website columns plus coordinate, "
            "URL and reference-count helpers. Position = seqid:start..end (IRGSP-1.0, 1-based).",
        ),
        (
            "Sheet 'References'",
            "One row per transcript-literature link: PMID, title, authors, journal, "
            "volume/issue/pages, year, link.",
        ),
        ("Sheet 'TO_terms' / 'PO_terms'", "One row per transcript-ontology term pair (long format)."),
        (
            "Sheet 'Endosperm keyword hits'",
            "Preview subset: records whose symbols/names/TO/PO match the endosperm keyword "
            "regex below. It is a convenience filter, not a curated endosperm gene list - "
            "verify each hit before use.",
        ),
        ("Endosperm keyword regex", ENDOSPERM_PATTERN.pattern),
        ("Script", "scrape_agri_genes.py (run: python scrape_agri_genes.py)"),
        ("Raw payload", RAW_JSON.name),
    ]
    return pd.DataFrame(rows, columns=["Item", "Value"])


HEADER_FILL = PatternFill("solid", fgColor="1F6B3B")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=10)
LINK_FONT = Font(color="0563C1", underline="single", size=10)


def style_sheet(
    worksheet: Any,
    widths: list[float],
    wrap_columns: tuple[int, ...] = (),
    link_columns: tuple[int, ...] = (),
    freeze: str = "A2",
) -> None:
    """Apply RAP-DB-ish header styling, filters, widths, wrapping and links."""
    for cell in worksheet[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    worksheet.row_dimensions[1].height = 30

    for index, width in enumerate(widths, start=1):
        worksheet.column_dimensions[get_column_letter(index)].width = width

    last_column = get_column_letter(worksheet.max_column)
    worksheet.auto_filter.ref = f"A1:{last_column}{worksheet.max_row}"
    worksheet.freeze_panes = freeze

    for column in wrap_columns:
        for row in range(2, worksheet.max_row + 1):
            cell = worksheet.cell(row=row, column=column)
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    for column in link_columns:
        for row in range(2, worksheet.max_row + 1):
            cell = worksheet.cell(row=row, column=column)
            if cell.value:
                cell.hyperlink = str(cell.value)
                cell.font = LINK_FONT


MAIN_WIDTHS = [6, 15, 17, 22, 11, 10, 10, 7, 30, 42, 45, 35, 40, 70, 10, 16, 40, 45, 42]
MAIN_WRAP = (9, 10, 11, 12, 13, 14)
MAIN_LINKS = (17, 18, 19)


def cap_cell_length(frame: pd.DataFrame, limit: int = 32000) -> pd.DataFrame:
    """Guard against Excel's 32,767-character per-cell limit."""
    for column in frame.columns:
        if frame[column].dtype == object:
            frame[column] = frame[column].map(
                lambda value: value
                if not isinstance(value, str) or len(value) <= limit
                else value[: limit - 3] + "..."
            )
    return frame


def write_workbooks(
    tables: dict[str, pd.DataFrame],
    hits: pd.DataFrame,
    hits_refs: pd.DataFrame,
    readme: pd.DataFrame,
) -> None:
    with pd.ExcelWriter(XLSX_PATH, engine="openpyxl") as writer:
        readme.to_excel(writer, sheet_name="README", index=False)
        tables["main"].to_excel(writer, sheet_name="Agri. genes", index=False)
        tables["references"].to_excel(writer, sheet_name="References", index=False)
        tables["to"].to_excel(writer, sheet_name="TO_terms", index=False)
        tables["po"].to_excel(writer, sheet_name="PO_terms", index=False)
        hits.to_excel(writer, sheet_name="Endosperm keyword hits", index=False)

        workbook = writer.book
        style_sheet(workbook["README"], [34, 110], wrap_columns=(2,))
        style_sheet(
            workbook["Agri. genes"],
            MAIN_WIDTHS,
            wrap_columns=MAIN_WRAP,
            link_columns=MAIN_LINKS,
        )
        style_sheet(
            workbook["References"],
            [15, 17, 12, 14, 70, 14, 18, 28, 8, 7, 10, 7, 42, 80],
            wrap_columns=(5, 14),
            link_columns=(13,),
        )
        style_sheet(workbook["TO_terms"], [15, 17, 12, 45], wrap_columns=(4,))
        style_sheet(workbook["PO_terms"], [15, 17, 12, 45], wrap_columns=(4,))
        style_sheet(
            workbook["Endosperm keyword hits"],
            [15, 17, 30, 45, 45, 35, 30, 24],
            wrap_columns=(3, 4, 5, 6, 7),
        )

    with pd.ExcelWriter(HITS_XLSX_PATH, engine="openpyxl") as writer:
        hits.to_excel(writer, sheet_name="Endosperm gene hits", index=False)
        hits_refs.to_excel(writer, sheet_name="Endosperm paper hits", index=False)
        workbook = writer.book
        style_sheet(
            workbook["Endosperm gene hits"],
            [15, 17, 30, 45, 45, 35, 30, 24],
            wrap_columns=(3, 4, 5, 6, 7),
        )
        style_sheet(
            workbook["Endosperm paper hits"],
            [15, 17, 12, 75, 28, 7, 22, 42],
            wrap_columns=(4,),
            link_columns=(8,),
        )


def validate_outputs(expected_rows: int) -> None:
    """Re-open the workbook and confirm sheet sizes before reporting success."""
    from openpyxl import load_workbook

    workbook = load_workbook(XLSX_PATH, read_only=True)
    try:
        for name in workbook.sheetnames:
            sheet = workbook[name]
            print(f"[check] sheet '{name}': {sheet.max_row} rows x {sheet.max_column} cols")
        main_sheet = workbook["Agri. genes"]
        data_rows = main_sheet.max_row - 1
        if data_rows != expected_rows:
            raise SystemExit(
                f"[error] 'Agri. genes' has {data_rows} data rows, expected {expected_rows}"
            )
    finally:
        workbook.close()
    print(f"[ok] workbook verified: {XLSX_PATH}")





def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Download the complete RAP-DB Agri. genes table and export it to Excel/CSV."
    )
    parser.add_argument(
        "--refresh",
        action="store_true",
        help="force a fresh download even if agri_genes.json already exists",
    )
    parser.add_argument(
        "--use-local",
        action="store_true",
        help="never touch the network; reuse the existing agri_genes.json",
    )
    args = parser.parse_args(argv)

    if args.use_local and not RAW_JSON.exists():
        raise SystemExit(f"[error] --use-local given but {RAW_JSON.name} is missing")

    records = download_json(force=args.refresh and not args.use_local)
    tables = build_tables(records)

    main_df = cap_cell_length(tables["main"])
    refs_df = cap_cell_length(tables["references"])
    to_df = cap_cell_length(tables["to"])
    po_df = cap_cell_length(tables["po"])
    tables = {"main": main_df, "references": refs_df, "to": to_df, "po": po_df}

    hits, hits_refs = find_endosperm_hits(main_df, refs_df)
    hits = cap_cell_length(hits)
    hits_refs = cap_cell_length(hits_refs)

    if refs_df.empty:
        unique_refs = 0
    else:
        unique_refs = int(refs_df["PubMed ID"].replace("", pd.NA).dropna().nunique())

    meta = {
        "retrieved": dt.datetime.now().astimezone().strftime("%Y-%m-%d %H:%M:%S UTC%z"),
        "records": len(main_df),
        "loci": main_df["Locus ID"].nunique(),
        "refs": len(refs_df),
        "unique_refs": unique_refs,
        "with_symbols": int((main_df["Gene symbols"] != "").sum()),
        "with_names": int((main_df["Gene names"] != "").sum()),
        "with_to": int((main_df["Trait Ontology (TO)"] != "").sum()),
        "with_po": int((main_df["Plant Ontology (PO)"] != "").sum()),
        "with_traits": int((main_df["Related traits (in Japanese)"] != "").sum()),
    }
    readme = build_readme(meta)

    main_df.to_csv(CSV_PATH, index=False, encoding="utf-8-sig")
    refs_df.to_csv(REF_CSV_PATH, index=False, encoding="utf-8-sig")

    write_workbooks(tables, hits, hits_refs, readme)
    validate_outputs(len(records))

    print("\n=== Summary ===")
    print(f"Transcripts            : {meta['records']:,}")
    print(f"Unique loci            : {meta['loci']:,}")
    print(f"Reference links        : {meta['refs']:,}")
    print(f"Unique paper records   : {meta['unique_refs']:,}")
    print(f"With gene symbols      : {meta['with_symbols']:,}")
    print(f"With gene names        : {meta['with_names']:,}")
    print(f"With TO terms          : {meta['with_to']:,}")
    print(f"With PO terms          : {meta['with_po']:,}")
    print(f"With related traits    : {meta['with_traits']:,}")
    print(f"Endosperm keyword hits : {len(hits):,} transcripts / {len(hits_refs):,} papers")
    print(f"\nWorkbook : {XLSX_PATH}")
    print(f"CSV      : {CSV_PATH}")
    print(f"Refs CSV : {REF_CSV_PATH}")
    print(f"Hits     : {HITS_XLSX_PATH}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
