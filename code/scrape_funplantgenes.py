#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Scrape the complete Oryza sativa section of funPlantGenes (HENAU), including
every gene detail page.

List pages:
    https://funplantgenes.henau.edu.cn/categories/oryza-sativa/
    https://funplantgenes.henau.edu.cn/categories/oryza-sativa/page/N/   (N = 2..103)
Gene pages (three groups appear in this category):
    https://funplantgenes.henau.edu.cn/genes/oryza_sativa/<slug>/
    https://funplantgenes.henau.edu.cn/genes/prgdb/<slug>/
    https://funplantgenes.henau.edu.cn/genes/pcfg/<slug>/

What is captured per gene (every collapsible section on the page):
    Gene Details, Functional Descriptions / Functional Annotation,
    Functional Keywords, Literature and News (title, DOI, PMID),
    Gene Resources (UniProt, EMBL, KEGG, InterPro, Pfam, ...),
    Homologs / Homologous (species + gene ID + link),
    Sequences (cDNA / CDS / protein).

Outputs (in the workspace's funPlantGenes directory):
    FunPlantGenes_Oryza_sativa.xlsx      README + 7 data sheets
    FunPlantGenes_Oryza_sativa_genes.csv main table
    FunPlantGenes_Oryza_sativa_literature.csv
    FunPlantGenes_Oryza_sativa_keywords.csv
    FunPlantGenes_Oryza_sativa_homologs.csv
    FunPlantGenes_Oryza_sativa_resources.csv
    FunPlantGenes_Oryza_sativa_sequences.csv    (manifest: lengths + FASTA header)
    sequences_cdna.fa / sequences_cds.fa / sequences_dna.fa / sequences_protein.fa
    cache/index.json, cache/html/*.html.gz      resumable download cache
    download_failures.csv                       only if something never succeeded

Typical run:
    python scrape_funplantgenes.py                 # full scrape (resumable)
    python scrape_funplantgenes.py --limit 20      # smoke test on 20 genes
    python scrape_funplantgenes.py --skip-download # rebuild outputs from cache
"""

from __future__ import annotations

import argparse
import datetime as dt
import gzip
import html
import json
import re
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Any, Iterable

import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

DATA_ROOT = Path(__file__).resolve().parent.parent
BASE_DIR = DATA_ROOT / "funPlantGenes"
SITE = "https://funplantgenes.henau.edu.cn"
LIST_URL = f"{SITE}/categories/oryza-sativa/"
CACHE_DIR = BASE_DIR / "cache"
HTML_DIR = CACHE_DIR / "html"
INDEX_JSON = CACHE_DIR / "index.json"

XLSX_PATH = BASE_DIR / "FunPlantGenes_Oryza_sativa.xlsx"
GENES_CSV = BASE_DIR / "FunPlantGenes_Oryza_sativa_genes.csv"
LIT_CSV = BASE_DIR / "FunPlantGenes_Oryza_sativa_literature.csv"
KW_CSV = BASE_DIR / "FunPlantGenes_Oryza_sativa_keywords.csv"
HOM_CSV = BASE_DIR / "FunPlantGenes_Oryza_sativa_homologs.csv"
RES_CSV = BASE_DIR / "FunPlantGenes_Oryza_sativa_resources.csv"
SEQ_CSV = BASE_DIR / "FunPlantGenes_Oryza_sativa_sequences.csv"
FAIL_CSV = BASE_DIR / "download_failures.csv"
FASTA_DIR = BASE_DIR

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}

# Endosperm / grain-filling preview filter for this rice-focused project.
ENDOSPERM_PATTERN = re.compile(
    r"\b("
    r"endosperm\w*|amylose|amylopectin|starch\w*|waxy|wx|"
    r"glutelin\w*|prolamin\w*|globulin\w*|aleurone|"
    r"storage protein\w*|grain filling|chalkiness|chalky|"
    r"amylase\w*|sucrose synthase|adp-glucose|granule-bound"
    r")\b",
    re.IGNORECASE,
)


def clean(value: Any) -> str:
    """Collapse NBSPs / whitespace and drop control characters."""
    if value is None:
        return ""
    text = str(value).replace("\xa0", " ").replace("\u00ad", "")
    text = CONTROL_CHARS.sub("", text)
    text = re.sub(r"[ \t\r\f\v]+", " ", text)
    text = re.sub(r"\n{2,}", "\n", text)
    return text.strip()


def make_session() -> requests.Session:
    session = requests.Session()
    session.headers.update(HEADERS)
    adapter = requests.adapters.HTTPAdapter(pool_connections=4, pool_maxsize=8)
    session.mount("https://", adapter)
    return session


def fetch(session: requests.Session, url: str, attempts: int = 5) -> requests.Response:
    """GET with retries; the site sporadically drops TLS handshakes."""
    last_error: Exception | None = None
    for attempt in range(1, attempts + 1):
        try:
            response = session.get(url, timeout=60)
            if response.status_code in (200, 404, 410):
                return response
            last_error = RuntimeError(f"HTTP {response.status_code}")
        except Exception as exc:  # noqa: BLE001 - network flakiness
            last_error = exc
        time.sleep(min(8, 1.5 * attempt) + (0.3 * attempt))
    raise RuntimeError(f"failed after {attempts} attempts: {last_error}")


def decode_html(response: requests.Response) -> str:
    """The server sends no charset; requests would fall back to latin-1."""
    return response.content.decode("utf-8", errors="replace")


def gene_key(group: str, slug: str) -> str:
    return f"{group}__{slug}"


def gzip_path(group: str, slug: str) -> Path:
    return HTML_DIR / f"{gene_key(group, slug)}.html.gz"


def missing_path(group: str, slug: str) -> Path:
    return HTML_DIR / f"{gene_key(group, slug)}.missing"


def save_html(group: str, slug: str, text: str) -> None:
    HTML_DIR.mkdir(parents=True, exist_ok=True)
    # Binary writes avoid changing the site's CRLF bytes on Windows.
    gzip_path(group, slug).write_bytes(gzip.compress(text.encode("utf-8"), mtime=0))


def load_html(group: str, slug: str) -> str | None:
    path = gzip_path(group, slug)
    if not path.exists():
        return None
    with gzip.open(path, "rt", encoding="utf-8") as handle:
        return handle.read()

CONTROL_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
SEQUENCE_CHARS = re.compile(r"[^A-Za-z*]")
PMID_RE = re.compile(r"PMID[:\s]*(\d+)", re.IGNORECASE)
DOI_RE = re.compile(r"10\.\d{4,9}/[^\s;]+")

_print_lock = threading.Lock()


def log(message: str) -> None:
    with _print_lock:
        print(f"[{dt.datetime.now():%H:%M:%S}] {message}", flush=True)



def parse_list_page(html: str, page: int) -> tuple[list[dict[str, Any]], int | None]:
    soup = BeautifulSoup(html, "lxml")
    entries: list[dict[str, Any]] = []
    for card in soup.select("a.card[href]"):
        href = card.get("href", "").strip()
        match = re.match(r"^/genes/([^/]+)/([^/]+)/?$", href)
        if not match:
            continue
        title_tag = card.select_one(".card-title")
        desc_tag = card.select_one(".card-description")
        entries.append(
            {
                "group": match.group(1),
                "slug": match.group(2),
                "url": f"{SITE}{href}",
                "card_symbol": clean(title_tag.get_text(" ", strip=True)) if title_tag else "",
                "card_description": clean(desc_tag.get_text(" ", strip=True)) if desc_tag else "",
                "list_page": page,
            }
        )
    total = None
    match = re.search(r"Total pages:\s*(\d+)", soup.get_text(" ", strip=True))
    if match:
        total = int(match.group(1))
    return entries, total


def enumerate_genes(refresh: bool = False) -> dict[str, Any]:
    """Walk all 103 list pages and cache the gene index as JSON."""
    if INDEX_JSON.exists() and not refresh:
        data = json.loads(INDEX_JSON.read_text(encoding="utf-8"))
        log(f"[index] cached: {len(data['genes']):,} genes / {data['total_pages']} pages")
        return data

    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    session = make_session()
    log(f"[index] GET {LIST_URL}")
    response = fetch(session, LIST_URL)
    response.raise_for_status()
    first_entries, total_pages = parse_list_page(decode_html(response), 1)
    if not total_pages:
        raise SystemExit("[error] could not find 'Total pages' on the first list page")
    log(f"[index] total pages reported: {total_pages}; page 1 -> {len(first_entries)} genes")

    genes: list[dict[str, Any]] = []
    seen: set[str] = set()

    def add(entries: Iterable[dict[str, Any]]) -> int:
        added = 0
        for entry in entries:
            if entry["url"] in seen:
                continue
            seen.add(entry["url"])
            genes.append(entry)
            added += 1
        return added

    add(first_entries)
    for page in range(2, total_pages + 1):
        url = f"{LIST_URL}page/{page}/"
        for attempt in range(1, 4):
            try:
                response = fetch(session, url)
                response.raise_for_status()
                entries, _ = parse_list_page(decode_html(response), page)
                added = add(entries)
                log(f"[index] page {page}/{total_pages}: {len(entries)} cards (+{added})")
                break
            except Exception as exc:  # noqa: BLE001
                log(f"[warn] list page {page} attempt {attempt}: {exc}")
                time.sleep(3 * attempt)
        else:
            raise SystemExit(f"[error] could not fetch list page {page}")
        time.sleep(0.1)

    data = {
        "source": LIST_URL,
        "fetched_at": dt.datetime.now().astimezone().strftime("%Y-%m-%d %H:%M:%S UTC%z"),
        "total_pages": total_pages,
        "genes": genes,
    }
    INDEX_JSON.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    log(f"[index] saved {len(genes):,} unique genes -> {INDEX_JSON.name}")
    return data


_thread_local = threading.local()


def thread_session() -> requests.Session:
    session = getattr(_thread_local, "session", None)
    if session is None:
        session = make_session()
        _thread_local.session = session
    return session


def download_all(
    genes: list[dict[str, Any]], workers: int = 6, limit: int | None = None
) -> None:
    """Download every gene page into cache/html/*.html.gz (skips cached pages)."""
    HTML_DIR.mkdir(parents=True, exist_ok=True)
    selected = genes[:limit] if limit else genes
    todo = [
        gene
        for gene in selected
        if not gzip_path(gene["group"], gene["slug"]).exists()
        and not missing_path(gene["group"], gene["slug"]).exists()
    ]
    log(
        f"[download] genes={len(selected):,} cached={len(selected) - len(todo):,} "
        f"to_fetch={len(todo):,} workers={workers}"
    )
    if not todo:
        return

    failures: list[dict[str, str]] = []
    done = 0
    started = time.time()

    def work(gene: dict[str, Any]) -> tuple[str, dict[str, Any], str]:
        session = thread_session()
        try:
            response = fetch(session, gene["url"])
            if response.status_code == 200:
                text = decode_html(response)
                if len(text) < 2000:
                    raise RuntimeError(f"suspiciously short page ({len(text)} bytes)")
                save_html(gene["group"], gene["slug"], text)
                return ("ok", gene, "")
            missing_path(gene["group"], gene["slug"]).write_text(
                str(response.status_code), encoding="utf-8"
            )
            return ("missing", gene, f"HTTP {response.status_code}")
        except Exception as exc:  # noqa: BLE001
            return ("fail", gene, str(exc))
        finally:
            time.sleep(0.05)

    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = {pool.submit(work, gene): gene for gene in todo}
        for future in as_completed(futures):
            status, gene, error = future.result()
            done += 1
            if status == "fail":
                failures.append(
                    {"url": gene["url"], "group": gene["group"], "slug": gene["slug"], "error": error}
                )
            if done % 100 == 0 or done == len(todo):
                elapsed = max(time.time() - started, 1e-6)
                rate = done / elapsed
                remaining = (len(todo) - done) / rate if rate else 0
                log(
                    f"[download] {done:,}/{len(todo):,} "
                    f"({rate:.1f}/s, ~{remaining / 60:.1f} min left), failures={len(failures)}"
                )

    # second chance for transient failures
    if failures:
        log(f"[download] retrying {len(failures)} failures once more")
        retry_failures: list[dict[str, str]] = []
        for item in failures:
            gene = next(g for g in todo if g["url"] == item["url"])
            status, _, error = work(gene)
            if status == "fail":
                retry_failures.append(item)
        failures = retry_failures

    if failures:
        pd.DataFrame(failures).to_csv(FAIL_CSV, index=False, encoding="utf-8-sig")
        log(f"[warn] {len(failures)} pages still failing -> {FAIL_CSV.name}")
    elif FAIL_CSV.exists():
        FAIL_CSV.unlink()
    log("[download] finished")




def parse_label_items(body: Any) -> list[tuple[str, str]]:
    """Parse <li><strong>Label:</strong> value</li> rows into (label, value)."""
    items: list[tuple[str, str]] = []
    for li in body.select("li"):
        strong = li.find("strong")
        if not strong:
            continue
        label = clean(strong.get_text(" ", strip=True)).rstrip(":").strip()
        if not label:
            continue
        text = clean(li.get_text(" ", strip=True))
        text = re.sub(rf"^{re.escape(label)}\s*:\s*", "", text, flags=re.IGNORECASE).strip()
        items.append((label, text))
    return items


def extract_alias_symbol(sections: dict[str, Any]) -> str:
    """Read 'Gene Symbol' while preserving NBSP-separated aliases as '; '."""
    body = sections.get("Gene Details")
    if body is None:
        return ""
    for li in body.select("li"):
        strong = li.find("strong")
        if strong is None:
            continue
        label = clean(strong.get_text(" ", strip=True)).rstrip(":").strip()
        if label.lower() not in ("gene symbol", "suggest gene name"):
            continue
        raw = li.decode_contents().replace("\xa0", "; ").replace("&nbsp;", "; ")
        raw = re.sub(r"<[^>]+>", " ", raw)
        raw = html.unescape(raw)
        raw = re.sub(rf"^\s*{re.escape(label)}\s*:\s*", "", raw, flags=re.IGNORECASE)
        raw = re.sub(r"\s*;\s*", "; ", raw)
        raw = re.sub(r"(?:;\s*)+", "; ", raw).strip("; ")
        return clean(raw)
    return ""


def parse_sequences(body: Any) -> list[dict[str, str]]:
    sequences: list[dict[str, str]] = []
    for accordion in body.select(".accordion"):
        header_tag = accordion.select_one(".accordion-header")
        type_name = clean(header_tag.get_text(" ", strip=True)) if header_tag else ""
        holder = accordion.select_one(".text-to-copy")
        if holder is None:
            continue
        if "cdna" in type_name.lower():
            seq_type = "cDNA"
        elif "cds" in type_name.lower():
            seq_type = "CDS"
        elif "dna" in type_name.lower():
            seq_type = "DNA"
        elif "protein" in type_name.lower():
            seq_type = "Protein"
        else:
            seq_type = type_name or "Sequence"

        # Multi-transcript genes list one <li> per sequence; never concatenate them.
        records = holder.select("li") or [holder]
        for record in records:
            raw = record.get_text("\n")
            lines = [line.strip() for line in raw.splitlines() if line.strip()]
            if not lines:
                continue
            fasta_header = lines[0] if lines[0].startswith(">") else ""
            payload = "".join(lines[1:] if fasta_header else lines)
            # Preserve source characters. Alphabet problems are reported separately,
            # never silently removed from an official-source export.
            sequence = re.sub(r"\s+", "", payload)
            if not sequence:
                continue
            if seq_type == "Protein":
                length = len(sequence.rstrip("*"))
            else:
                length = len(sequence.replace("*", ""))
            sequences.append(
                {
                    "type": seq_type,
                    "fasta_header": fasta_header,
                    "length": str(length),
                    "sequence": sequence,
                }
            )
    return sequences


def parse_literature(body: Any) -> list[dict[str, str]]:
    papers: list[dict[str, str]] = []
    for li in body.select("li"):
        title = ""
        doi = ""
        pmid = ""
        link = ""
        for anchor in li.select("a"):
            href = anchor.get("href", "")
            text = clean(anchor.get_text(" ", strip=True))
            if "pubmed" in href:
                pmid = text or href.rstrip("/").split("/")[-1]
            elif "doi.org" in href:
                if DOI_RE.fullmatch(text):
                    doi = doi or text
                elif not title:
                    title = text
                    link = href
            elif not title and len(text) > 3:
                title = text
                link = href
        raw = clean(li.get_text(" ", strip=True))
        if not pmid:
            match = PMID_RE.search(raw)
            pmid = match.group(1) if match else ""
        if not doi:
            match = DOI_RE.search(raw)
            doi = match.group(0).rstrip(".;,") if match else ""
        if not title:
            title = clean(re.split(r"DOI:|PMID:", raw)[0]).strip(" .;")
        if title or doi or pmid:
            papers.append(
                {
                    "title": title,
                    "doi": doi,
                    "pmid": pmid,
                    "link": link or (f"https://doi.org/{doi}" if doi else ""),
                }
            )
    return papers


def parse_gene(
    gene: dict[str, Any], html: str
) -> tuple[dict[str, Any], dict[str, list[dict[str, str]]]]:
    """Extract every section of one gene page into a flat row + child tables."""
    soup = BeautifulSoup(html, "lxml")
    sections: dict[str, Any] = {}
    for notice in soup.select("div.notice"):
        head = notice.select_one(".notice-head p")
        body = notice.select_one(".notice-body")
        if head is None or body is None:
            continue
        sections[clean(head.get_text(" ", strip=True))] = body

    details = (
        dict(parse_label_items(sections["Gene Details"])) if "Gene Details" in sections else {}
    )
    detail_labels = {key.casefold(): value for key, value in details.items()}

    def detail(name: str, default: str = "") -> str:
        return detail_labels.get(name.casefold(), default)

    descriptions: list[str] = []
    for name in ("Functional Descriptions", "Functional Annotation", "Domain Class Descriptions"):
        if name in sections:
            descriptions.extend(
                clean(li.get_text(" ", strip=True))
                for li in sections[name].select("li")
                if clean(li.get_text(" ", strip=True))
            )

    keywords: list[str] = []
    if "Functional Keywords" in sections:
        for anchor in sections["Functional Keywords"].select("a"):
            text = clean(anchor.get_text(" ", strip=True))
            if text and text not in keywords:
                keywords.append(text)

    domain_class = ""
    if "Domain Features" in sections:
        for li in sections["Domain Features"].select("li"):
            text = clean(li.get_text(" ", strip=True))
            match = re.match(r"Domain Class\s*:\s*(.+)", text, flags=re.IGNORECASE)
            if match:
                domain_class = match.group(1).strip()
                break

    literature = (
        parse_literature(sections["Literature and News"])
        if "Literature and News" in sections
        else []
    )

    resources: dict[str, str] = {}
    if "Gene Resources" in sections:
        for label, value in parse_label_items(sections["Gene Resources"]):
            resources[label] = value

    homologs: list[dict[str, str]] = []
    for name in ("Homologs", "Homologous"):
        if name not in sections:
            continue
        for li in sections[name].select("li"):
            strong = li.find("strong")
            if strong is None:
                continue
            species = clean(strong.get_text(" ", strip=True)).rstrip(":").strip()
            anchors = li.select("a")
            homolog_id = ", ".join(clean(a.get_text(" ", strip=True)) for a in anchors)
            if not homolog_id:
                homolog_id = clean(li.get_text(" ", strip=True))
                homolog_id = re.sub(
                    rf"^{re.escape(species)}\s*:\s*", "", homolog_id
                ).strip()
            link = anchors[0].get("href", "") if anchors else ""
            if link.startswith("/"):
                link = f"{SITE}{link}"
            if species or homolog_id:
                homologs.append(
                    {"species": species, "homolog_id": homolog_id, "link": link}
                )

    sequences = parse_sequences(sections["Sequences"]) if "Sequences" in sections else []
    seq_lengths: dict[str, int] = {}
    for item in sequences:
        seq_lengths.setdefault(item["type"], int(item["length"]))
    resource_lookup = {key.lower(): value for key, value in resources.items()}

    def pick_resource(*names: str) -> str:
        for name in names:
            value = resource_lookup.get(name.lower())
            if value:
                return value
        return ""

    known = {
        "Gene Details",
        "Functional Descriptions",
        "Functional Annotation",
        "Domain Class Descriptions",
        "Functional Keywords",
        "Domain Features",
        "Literature and News",
        "Gene Resources",
        "Homologs",
        "Homologous",
        "Sequences",
    }
    extra_sections = []
    for name, body in sections.items():
        if name in known:
            continue
        text = clean(body.get_text(" ", strip=True))
        if text:
            extra_sections.append(f"{name}: {text}")

    row = {
        "Page Gene ID": gene["slug"],
        "Gene Key": f"{gene['group']}__{gene['slug']}",
        "Gene Group": gene["group"],
        "Gene Symbol": extract_alias_symbol(sections)
        or detail("Gene Symbol")
        or detail("Gene Name")
        or detail("Suggest Gene Name")
        or gene.get("card_symbol", ""),
        "MSU Gene ID": detail("MSU gene ID"),
        "RAPdb Gene ID": detail("RAPdb gene ID"),
        "Gene ID (page)": detail("Gene ID"),
        "Genome": detail("Genome"),
        "Species": detail("Species"),
        "Description": detail("Description"),
        "Source": detail("Source"),
        "Domain Class": domain_class,
        "Functional Descriptions": " || ".join(descriptions),
        "Functional Keywords": "; ".join(keywords),
        "No. of Keywords": len(keywords),
        "No. of References": len(literature),
        "PMIDs": "; ".join(sorted({paper["pmid"] for paper in literature if paper["pmid"]})),
        "DOIs": "; ".join(sorted({paper["doi"] for paper in literature if paper["doi"]})),
        "No. of Homologs": len(homologs),
        "Homologs": "; ".join(f"{h['species']}: {h['homolog_id']}" for h in homologs),
        "UniProt": pick_resource("UniProt"),
        "EnsemblPlants": pick_resource("EnsemblPlants"),
        "KEGG": pick_resource("KEGG"),
        "InterPro": pick_resource("InterPro"),
        "Pfam": pick_resource("Pfam"),
        "cDNA length": seq_lengths.get("cDNA", ""),
        "CDS length": seq_lengths.get("CDS", ""),
        "Protein length": seq_lengths.get("Protein", ""),
        "DNA length": seq_lengths.get("DNA", ""),
        "No. of sequence records": len(sequences),
        "Page URL": gene["url"],
        "Additional Sections": " || ".join(extra_sections),
        "Status": "ok" if sections else "no-data",
    }

    children = {
        "literature": [
            {"gene": row["Gene Key"], "symbol": row["Gene Symbol"], **paper}
            for paper in literature
        ],
        "keywords": [
            {"gene": row["Gene Key"], "symbol": row["Gene Symbol"], "keyword": keyword}
            for keyword in keywords
        ],
        "homologs": [
            {"gene": row["Gene Key"], "symbol": row["Gene Symbol"], **homolog}
            for homolog in homologs
        ],
        "resources": [
            {"gene": row["Gene Key"], "symbol": row["Gene Symbol"], "database": db, "value": value}
            for db, value in resources.items()
        ],
        "sequences": [
            {
                "gene": row["Gene Key"],
                "symbol": row["Gene Symbol"],
                **item,
                "record_id": f"{row['Gene Key']}|{item['type']}|{n:04d}",
                "source_url": gene["url"],
            }
            for n, item in enumerate(sequences, 1)
        ],
    }
    return row, children


def parse_all(genes: list[dict[str, Any]], limit: int | None = None) -> dict[str, Any]:
    """Parse every cached page; pages without HTML are reported as missing."""
    selected = genes[:limit] if limit else genes
    rows: list[dict[str, Any]] = []
    literature: list[dict[str, str]] = []
    keywords: list[dict[str, str]] = []
    homologs: list[dict[str, str]] = []
    resources: list[dict[str, str]] = []
    sequences: list[dict[str, str]] = []
    missing: list[dict[str, str]] = []

    for index, gene in enumerate(selected, start=1):
        html = load_html(gene["group"], gene["slug"])
        if html is None:
            marker = missing_path(gene["group"], gene["slug"])
            status = marker.read_text(encoding="utf-8").strip() if marker.exists() else "no-html"
            rows.append(
                {
                    "Page Gene ID": gene["slug"],
                    "Gene Key": f"{gene['group']}__{gene['slug']}",
                    "Gene Group": gene["group"],
                    "Gene Symbol": gene.get("card_symbol", ""),
                    "MSU Gene ID": "",
                    "RAPdb Gene ID": "",
                    "Gene ID (page)": "",
                    "Genome": "",
                    "Species": "",
                    "Description": gene.get("card_description", ""),
                    "Source": "",
                    "Domain Class": "",
                    "Functional Descriptions": "",
                    "Functional Keywords": "",
                    "No. of Keywords": 0,
                    "No. of References": 0,
                    "PMIDs": "",
                    "DOIs": "",
                    "No. of Homologs": 0,
                    "Homologs": "",
                    "UniProt": "",
                    "EnsemblPlants": "",
                    "KEGG": "",
                    "InterPro": "",
                    "Pfam": "",
                    "cDNA length": "",
                    "CDS length": "",
                    "Protein length": "",
                    "DNA length": "",
                    "No. of sequence records": 0,
                    "Page URL": gene["url"],
                    "Additional Sections": "",
                    "Status": status,
                }
            )
            missing.append({"url": gene["url"], "status": status})
            continue

        row, children = parse_gene(gene, html)
        rows.append(row)
        literature.extend(children["literature"])
        keywords.extend(children["keywords"])
        homologs.extend(children["homologs"])
        resources.extend(children["resources"])
        sequences.extend(children["sequences"])

        if index % 500 == 0 or index == len(selected):
            log(f"[parse] {index:,}/{len(selected):,} pages")

    main_df = pd.DataFrame(rows)
    main_df = main_df.reindex(columns=MAIN_COLUMNS)
    main_df.insert(0, "No.", range(1, len(main_df) + 1))

    def frame(data: list[dict[str, Any]], columns: list[str]) -> pd.DataFrame:
        return pd.DataFrame(data, columns=columns) if data else pd.DataFrame(columns=columns)

    return {
        "main": main_df,
        "literature": frame(
            literature, ["gene", "symbol", "title", "doi", "pmid", "link"]
        ),
        "keywords": frame(keywords, ["gene", "symbol", "keyword"]),
        "homologs": frame(
            homologs, ["gene", "symbol", "species", "homolog_id", "link"]
        ),
        "resources": frame(resources, ["gene", "symbol", "database", "value"]),
        "sequences": frame(
            sequences,
            ["gene", "symbol", "type", "fasta_header", "length", "sequence", "record_id", "source_url"],
        ),
        "missing": frame(missing, ["url", "status"]),
    }


def find_endosperm_hits(
    main_df: pd.DataFrame, literature: pd.DataFrame
) -> pd.DataFrame:
    """Keyword preview over gene symbols/descriptions/keywords and paper titles."""
    titles: dict[str, list[str]] = {}
    if not literature.empty:
        for row in literature.to_dict("records"):
            titles.setdefault(row.get("gene", ""), []).append(str(row.get("title") or ""))

    hits: list[dict[str, Any]] = []
    for row in main_df.to_dict("records"):
        text = " ".join(
            str(row.get(column) or "")
            for column in (
                "Gene Symbol",
                "Description",
                "Domain Class",
                "Functional Descriptions",
                "Functional Keywords",
            )
        )
        text += " " + " ".join(titles.get(row.get("Gene Key", ""), []))
        tokens = {match.lower() for match in ENDOSPERM_PATTERN.findall(text)}
        if not tokens:
            continue
        hits.append(
            {
                "Gene Key": row.get("Gene Key", ""),
                "Page Gene ID": row.get("Page Gene ID", ""),
                "MSU Gene ID": row.get("MSU Gene ID", ""),
                "Gene Symbol": row.get("Gene Symbol", ""),
                "Domain Class": row.get("Domain Class", ""),
                "Functional Keywords": row.get("Functional Keywords", ""),
                "Matched keywords": "; ".join(sorted(tokens)),
                "No. of References": row.get("No. of References", ""),
                "Page URL": row.get("Page URL", ""),
            }
        )
    return pd.DataFrame(hits)


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


def build_readme(meta: dict[str, Any]) -> pd.DataFrame:
    rows = [
        ("funPlantGenes (HENAU) - Oryza sativa complete export", ""),
        ("Source category page", LIST_URL),
        ("Gene page pattern", f"{SITE}/genes/<group>/<slug>/"),
        ("List pages crawled", meta["pages"]),
        ("Retrieved (local time)", meta["retrieved"]),
        ("Genes in index", meta["index_genes"]),
        ("Genes exported to this workbook", meta["genes"]),
        ("Gene page groups", meta["groups"]),
        ("Pages parsed successfully", meta["ok"]),
        ("Pages unavailable (404/no data)", meta["missing"]),
        ("Reference links (paper x gene)", meta["refs"]),
        ("Unique PMIDs", meta["unique_pmids"]),
        ("Functional keyword rows", meta["keywords"]),
        ("Homolog rows", meta["homologs"]),
        ("Gene-resource rows", meta["resources"]),
        ("Sequence records (cDNA/CDS/DNA/protein)", meta["sequences"]),
        ("Genes with symbol", meta["with_symbol"]),
        ("Genes with functional description", meta["with_desc"]),
        ("Genes with >=1 reference", meta["with_refs"]),
        ("Genes with sequences", meta["with_seq"]),
        ("", ""),
        (
            "Method note",
            "Every list page (1..103) was parsed for /genes/... cards, then each gene page "
            "was downloaded once (cache/html/*.html.gz) and every collapsible section was "
            "extracted: Gene Details, Functional Descriptions/Annotation, Functional Keywords, "
            "Literature and News, Gene Resources, Homologs/Homologous, Sequences. "
            "Requests are rate-limited and retried; rerunning resumes from the cache.",
        ),
        (
            "PCFG records",
            "Part of the category links /genes/pcfg/... entries (putative flowering genes "
            "from PlantCFG). Their pages use 'Gene ID', 'Suggest Gene Name', 'Description', "
            "'Functional Annotation' and 'Homologous'; these are mapped into the same columns.",
        ),
        (
            "PRGDB records",
            "/genes/prgdb/... entries come from PRGdb 4.0 and use 'Gene Name', "
            "'Domain Class Descriptions' (mapped to Functional Descriptions) and "
            "'Domain Features' (Domain Class column). Sequence types follow the website's "
            "labels, including any source mislabelling; see SOURCE_ISSUES.csv.",
        ),
        (
            "Sequences",
            "Full sequences are written to sequences_cdna.fa / sequences_cds.fa / "
            "sequences_dna.fa / sequences_protein.fa (wrapped at 60 aa/nt). Multi-transcript "
            "genes contribute one FASTA record and one 'Sequences' sheet row per transcript; "
            "'cDNA/CDS/DNA/Protein length' in the Genes sheet refer to the first record, and "
            "'No. of sequence records' counts all records of that gene. Root FASTA files "
            "use unique record_id values. Original headers and unchanged sequences are "
            "also exported to official_fasta/. The manifest preserves fasta_header and source_url. "
            "Source identity/type anomalies are retained and listed in SOURCE_ISSUES.csv.",
        ),
        (
            "Text handling",
            "The server omits a charset and requests would decode as latin-1; pages are decoded "
            "as UTF-8. NBSPs are normalised to spaces and control characters removed.",
        ),
        ("Script", "scrape_funplantgenes.py (python scrape_funplantgenes.py)"),
        ("Cache", "cache/index.json + cache/html/*.html.gz; retain for source provenance"),
    ]
    return pd.DataFrame(rows, columns=["Item", "Value"])


def write_fasta(sequences: pd.DataFrame) -> dict[str, int]:
    targets = {
        "cDNA": FASTA_DIR / "sequences_cdna.fa",
        "CDS": FASTA_DIR / "sequences_cds.fa",
        "DNA": FASTA_DIR / "sequences_dna.fa",
        "Protein": FASTA_DIR / "sequences_protein.fa",
    }
    counts: dict[str, int] = {}
    original_dir = FASTA_DIR / "official_fasta"
    original_dir.mkdir(parents=True, exist_ok=True)
    for seq_type, path in targets.items():
        subset = sequences[sequences["type"] == seq_type] if not sequences.empty else sequences
        with path.open("w", encoding="utf-8", newline="\n") as handle, (original_dir / path.name).open("w", encoding="utf-8", newline="\n") as original:
            for row in subset.to_dict("records"):
                source_header = str(row.get("fasta_header") or f">{row.get('gene', '')}")
                record_id = str(row.get("record_id") or row.get("gene", ""))
                header = f">{record_id} source_header={source_header.lstrip('>')}"
                symbol = str(row.get("symbol") or "")
                if symbol:
                    header = f"{header} | {symbol}"
                handle.write(header + "\n")
                original.write(source_header + "\n")
                sequence = str(row.get("sequence") or "")
                for start in range(0, len(sequence), 60):
                    handle.write(sequence[start : start + 60] + "\n")
                    original.write(sequence[start : start + 60] + "\n")
        counts[seq_type] = len(subset)
    return counts


GENE_WIDTHS = [
    5, 16, 28, 13, 22, 16, 15, 14, 18, 12, 45, 20, 14, 70, 35, 8, 8, 18, 18, 8,
    45, 14, 16, 14, 24, 12, 10, 10, 12, 10, 10, 45, 50, 8,
]

MAIN_COLUMNS = [
    "Page Gene ID",
    "Gene Key",
    "Gene Group",
    "Gene Symbol",
    "MSU Gene ID",
    "RAPdb Gene ID",
    "Gene ID (page)",
    "Genome",
    "Species",
    "Description",
    "Source",
    "Domain Class",
    "Functional Descriptions",
    "Functional Keywords",
    "No. of Keywords",
    "No. of References",
    "PMIDs",
    "DOIs",
    "No. of Homologs",
    "Homologs",
    "UniProt",
    "EnsemblPlants",
    "KEGG",
    "InterPro",
    "Pfam",
    "cDNA length",
    "CDS length",
    "Protein length",
    "DNA length",
    "No. of sequence records",
    "Page URL",
    "Additional Sections",
    "Status",
]


def write_outputs(
    tables: dict[str, pd.DataFrame], hits: pd.DataFrame, readme: pd.DataFrame
) -> dict[str, int]:
    main_df = cap_cell_length(tables["main"])
    literature = cap_cell_length(tables["literature"])
    keywords = cap_cell_length(tables["keywords"])
    homologs = cap_cell_length(tables["homologs"])
    resources = cap_cell_length(tables["resources"])
    sequences = tables["sequences"]
    manifest = sequences.drop(columns=["sequence"], errors="ignore").copy()
    hits_df = cap_cell_length(hits)

    main_df.to_csv(GENES_CSV, index=False, encoding="utf-8-sig")
    literature.to_csv(LIT_CSV, index=False, encoding="utf-8-sig")
    keywords.to_csv(KW_CSV, index=False, encoding="utf-8-sig")
    homologs.to_csv(HOM_CSV, index=False, encoding="utf-8-sig")
    resources.to_csv(RES_CSV, index=False, encoding="utf-8-sig")
    manifest.to_csv(SEQ_CSV, index=False, encoding="utf-8-sig")

    fasta_counts = write_fasta(sequences)

    with pd.ExcelWriter(XLSX_PATH, engine="openpyxl") as writer:
        readme.to_excel(writer, sheet_name="README", index=False)
        main_df.to_excel(writer, sheet_name="Genes", index=False)
        literature.to_excel(writer, sheet_name="Literature", index=False)
        keywords.to_excel(writer, sheet_name="Keywords", index=False)
        homologs.to_excel(writer, sheet_name="Homologs", index=False)
        resources.to_excel(writer, sheet_name="Gene Resources", index=False)
        manifest.to_excel(writer, sheet_name="Sequences", index=False)
        hits_df.to_excel(writer, sheet_name="Endosperm keyword hits", index=False)

        workbook = writer.book
        style_sheet(workbook["README"], [40, 110], wrap_columns=(2,))
        style_sheet(
            workbook["Genes"],
            GENE_WIDTHS,
            wrap_columns=(11, 14, 15, 21, 33),
            link_columns=(32,),
        )
        style_sheet(
            workbook["Literature"],
            [16, 22, 70, 26, 12, 45],
            wrap_columns=(3,),
            link_columns=(6,),
        )
        style_sheet(workbook["Keywords"], [16, 22, 30])
        style_sheet(
            workbook["Homologs"], [16, 22, 28, 22, 45], link_columns=(5,)
        )
        style_sheet(workbook["Gene Resources"], [16, 22, 22, 60], wrap_columns=(4,))
        style_sheet(workbook["Sequences"], [16, 22, 10, 32, 10, 46, 65], link_columns=(7,))
        style_sheet(
            workbook["Endosperm keyword hits"],
            [28, 16, 16, 22, 14, 35, 24, 8, 45],
            wrap_columns=(6,),
            link_columns=(9,),
        )
    return fasta_counts


def validate_outputs(expected_rows: int) -> None:
    from openpyxl import load_workbook

    workbook = load_workbook(XLSX_PATH, read_only=True)
    try:
        for name in workbook.sheetnames:
            sheet = workbook[name]
            log(f"[check] sheet '{name}': {sheet.max_row:,} rows x {sheet.max_column} cols")
        data_rows = workbook["Genes"].max_row - 1
        if data_rows != expected_rows:
            raise SystemExit(
                f"[error] 'Genes' has {data_rows:,} rows, expected {expected_rows:,}"
            )
    finally:
        workbook.close()
    log(f"[ok] workbook verified: {XLSX_PATH}")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Scrape all Oryza sativa genes (list + detail pages) from funPlantGenes."
    )
    parser.add_argument("--refresh-index", action="store_true", help="re-crawl the 103 list pages")
    parser.add_argument("--workers", type=int, default=6, help="parallel gene page downloads")
    parser.add_argument("--limit", type=int, help="only process the first N genes (smoke test)")
    parser.add_argument(
        "--skip-download",
        action="store_true",
        help="do not hit the network; rebuild outputs from cache/html",
    )
    parser.add_argument(
        "--enumerate-only", action="store_true", help="only crawl the list pages and exit"
    )
    args = parser.parse_args(argv)

    index = enumerate_genes(refresh=args.refresh_index)
    genes = index["genes"][: args.limit] if args.limit else index["genes"]
    log(f"[main] genes selected: {len(genes):,} (index total {len(index['genes']):,})")

    if args.enumerate_only:
        for gene in genes[:5]:
            log(f"[sample] {gene['card_symbol']} -> {gene['url']}")
        return 0

    if not args.skip_download:
        download_all(genes, workers=args.workers)
    else:
        log("[main] --skip-download: using cached HTML only")

    tables = parse_all(genes)
    hits = find_endosperm_hits(tables["main"], tables["literature"])

    main_df = tables["main"]
    ok_mask = main_df["Status"] == "ok"
    with_seq = (
        main_df["cDNA length"].astype(str).ne("")
        | main_df["CDS length"].astype(str).ne("")
        | main_df["DNA length"].astype(str).ne("")
        | main_df["Protein length"].astype(str).ne("")
    )
    literature = tables["literature"]
    unique_pmids = 0
    if not literature.empty and "pmid" in literature:
        unique_pmids = int(literature["pmid"].replace("", pd.NA).dropna().nunique())

    meta = {
        "pages": index["total_pages"],
        "retrieved": dt.datetime.now().astimezone().strftime("%Y-%m-%d %H:%M:%S UTC%z"),
        "index_genes": len(index["genes"]),
        "genes": len(main_df),
        "groups": "; ".join(
            f"{group}: {count:,}"
            for group, count in main_df["Gene Group"].value_counts().items()
        ),
        "ok": int(ok_mask.sum()),
        "missing": int((~ok_mask).sum()),
        "refs": len(literature),
        "unique_pmids": unique_pmids,
        "keywords": len(tables["keywords"]),
        "homologs": len(tables["homologs"]),
        "resources": len(tables["resources"]),
        "sequences": len(tables["sequences"]),
        "with_symbol": int(main_df["Gene Symbol"].astype(str).str.strip().ne("").sum()),
        "with_desc": int(main_df["Functional Descriptions"].astype(str).str.strip().ne("").sum()),
        "with_refs": int((main_df["No. of References"].fillna(0) > 0).sum()),
        "with_seq": int(with_seq.sum()),
    }
    readme = build_readme(meta)
    fasta_counts = write_outputs(tables, hits, readme)
    validate_outputs(len(main_df))

    print("\n=== Summary ===")
    print(f"List pages             : {meta['pages']}")
    print(f"Genes in index         : {meta['index_genes']:,}")
    print(f"Gene pages parsed      : {meta['genes']:,} (ok {meta['ok']:,}, unavailable {meta['missing']:,})")
    print(f"Reference links        : {meta['refs']:,} (unique PMIDs {meta['unique_pmids']:,})")
    print(f"Keyword rows           : {meta['keywords']:,}")
    print(f"Homolog rows           : {meta['homologs']:,}")
    print(f"Gene-resource rows     : {meta['resources']:,}")
    print(f"Sequences              : {meta['sequences']:,} "
          f"(cDNA {fasta_counts.get('cDNA', 0):,}, CDS {fasta_counts.get('CDS', 0):,}, "
          f"DNA {fasta_counts.get('DNA', 0):,}, "
          f"protein {fasta_counts.get('Protein', 0):,})")
    print(f"Genes with symbol/desc : {meta['with_symbol']:,} / {meta['with_desc']:,}")
    print(f"Endosperm keyword hits : {len(hits):,}")
    print(f"\nWorkbook : {XLSX_PATH}")
    print(f"CSV      : {GENES_CSV}")
    print(f"FASTA    : {FASTA_DIR / 'sequences_cdna.fa'} (+ cds / protein)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
