#!/usr/bin/env python3
import argparse
import os
import re
import json
import time
import hashlib
import sqlite3
from dataclasses import dataclass
from pathlib import Path
from typing import List, Dict, Optional, Tuple

# Optional: pdf text extraction if available
try:
    import pdfplumber  # type: ignore
except Exception:
    pdfplumber = None

# --------- CONFIG (safe defaults; no deletions, no writes to Dropbox) ----------
DEFAULT_EXT_GROUPS = {
    "drawings": {".pdf", ".dwg", ".dxf"},
    "specs": {".pdf", ".doc", ".docx"},
    "spreadsheets": {".xls", ".xlsx", ".csv"},
    "text": {".txt", ".rtf", ".md"},
    "emails": {".eml", ".msg", ".mht"},
    "images": {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".heic"},
    "archives": {".zip", ".7z", ".rar"},
    "other": set()
}

KEYWORDS_DRAWINGS = ["drawings", "plans", "plan set", "ifc", "permit set", "cd set", "issuance", "architectural", "mep"]
KEYWORDS_SPECS    = ["spec", "specifications", "project manual", "div 01", "section", "addenda"]
KEYWORDS_ADDENDA  = ["addendum", "bulletin", "sketch", "asi", "rfi response", "revision bulletin"]

# Program/operational context keywords (non-governing, never scope by itself)
PROGRAM_CONTEXT_KEYWORDS = ["weight loss", "bariatric", "nutrition", "program", "clinic operations"]

SAME_NAME_RISK_TERMS = ["archive", "old", "closed out", "2021", "2022", "2023"]

# --------- DATA STRUCTURES ----------
@dataclass
class FileRec:
    path: str
    name: str
    ext: str
    size: int
    mtime: float
    sha256: Optional[str] = None
    group: str = "other"
    score: float = 0.0
    class_hint: str = "unknown"
    project_hits: int = 0
    option_hits: int = 0
    vendor_hits: int = 0
    context_hits: int = 0
    sample_text: str = ""

# --------- DB (continuous crawl memory) ----------
def db_init(db_path: str):
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("""
    CREATE TABLE IF NOT EXISTS files (
        path TEXT PRIMARY KEY,
        name TEXT,
        ext TEXT,
        size INTEGER,
        mtime REAL,
        sha256 TEXT,
        group_name TEXT
    )
    """)
    cur.execute("CREATE INDEX IF NOT EXISTS idx_files_name ON files(name)")
    conn.commit()
    return conn

def sha256_file(path: str, chunk_size=1024*1024) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while True:
            chunk = f.read(chunk_size)
            if not chunk:
                break
            h.update(chunk)
    return h.hexdigest()

def classify_group(ext: str) -> str:
    ext = ext.lower()
    for g, exts in DEFAULT_EXT_GROUPS.items():
        if ext in exts:
            return g
    return "other"

def safe_read_pdf_text(path: str, max_pages: int = 2, max_chars: int = 6000) -> str:
    if pdfplumber is None:
        return ""
    try:
        txt_parts = []
        with pdfplumber.open(path) as pdf:
            for i in range(min(max_pages, len(pdf.pages))):
                t = pdf.pages[i].extract_text() or ""
                if t.strip():
                    txt_parts.append(t)
        txt = "\n".join(txt_parts)
        return txt[:max_chars]
    except Exception:
        return ""

def safe_read_textlike(path: str, max_chars: int = 6000) -> str:
    # For txt/csv/md. Avoid binary.
    try:
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            return f.read(max_chars)
    except Exception:
        return ""

def normalize(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip().lower()

def count_hits(hay: str, needles: List[str]) -> int:
    h = normalize(hay)
    return sum(1 for n in needles if normalize(n) in h)

def project_tokenize(project: str) -> List[str]:
    # Conservative tokens: split on dash, spaces, punctuation.
    base = re.sub(r"[^\w\s]", " ", project)
    toks = [t for t in base.split() if len(t) >= 3]
    # Add variants
    variants = set(toks)
    variants |= {t.replace("saint", "st") for t in toks}
    variants |= {t.replace("st", "saint") for t in toks}
    return sorted(variants)

def discover_files(root: str) -> List[FileRec]:
    recs: List[FileRec] = []
    for dirpath, _, filenames in os.walk(root):
        for fn in filenames:
            p = os.path.join(dirpath, fn)
            try:
                st = os.stat(p)
            except Exception:
                continue
            ext = Path(fn).suffix.lower()
            recs.append(FileRec(
                path=p,
                name=fn,
                ext=ext,
                size=st.st_size,
                mtime=st.st_mtime,
                group=classify_group(ext)
            ))
    return recs

def score_file(rec: FileRec, project_tokens: List[str], vendor_tokens: List[str]) -> FileRec:
    name_l = normalize(rec.name)
    path_l = normalize(rec.path)

    # Base score: project token hits in filename/path
    rec.project_hits = sum(1 for t in project_tokens if t.lower() in name_l or t.lower() in path_l)

    # Option hits: only from name/path at this phase
    rec.option_hits = count_hits(rec.name + " " + rec.path, ["option 1", "option 2", "opt 1", "opt 2", "alternate", "alt"])

    # Vendor hits: helps ranking but never makes governing by itself
    rec.vendor_hits = sum(1 for v in vendor_tokens if v and v.lower() in (name_l + " " + path_l))

    # Program context hits: allows “weight loss” detection as context only
    rec.context_hits = count_hits(rec.name + " " + rec.path, PROGRAM_CONTEXT_KEYWORDS)

    # Class hint from name/path keywords
    blob = f"{rec.name} {rec.path}"
    if count_hits(blob, KEYWORDS_ADDENDA) > 0:
        rec.class_hint = "addenda"
    elif count_hits(blob, KEYWORDS_SPECS) > 0:
        rec.class_hint = "specs"
    elif count_hits(blob, KEYWORDS_DRAWINGS) > 0:
        rec.class_hint = "drawings"
    else:
        rec.class_hint = "unknown"

    # Same-name risk heuristic (archive-ish paths)
    risk = 1 if any(rt in path_l for rt in SAME_NAME_RISK_TERMS) else 0

    # Scoring: conservative, explainable
    score = 0.0
    score += rec.project_hits * 5.0
    score += rec.option_hits * 1.5
    score += min(rec.vendor_hits, 5) * 0.5
    score += rec.context_hits * 0.2
    score += 2.0 if rec.class_hint in ("drawings", "specs", "addenda") else 0.0
    score -= 2.0 * risk

    # Small bump for likely project folders
    if any(k in path_l for k in ["project", "jobs", "estimating", "construction", "ti", "tenant"]):
        score += 0.5

    rec.score = score
    return rec

def update_index(conn, recs: List[FileRec], hash_changed_only: bool = True):
    cur = conn.cursor()
    for r in recs:
        cur.execute("SELECT size, mtime, sha256 FROM files WHERE path=?", (r.path,))
        row = cur.fetchone()
        needs_hash = True
        if row and hash_changed_only:
            old_size, old_mtime, old_hash = row
            if old_size == r.size and float(old_mtime) == float(r.mtime) and old_hash:
                r.sha256 = old_hash
                needs_hash = False
        if needs_hash:
            try:
                r.sha256 = sha256_file(r.path)
            except Exception:
                r.sha256 = None

        cur.execute("""
        INSERT INTO files(path, name, ext, size, mtime, sha256, group_name)
        VALUES(?,?,?,?,?,?,?)
        ON CONFLICT(path) DO UPDATE SET
          name=excluded.name,
          ext=excluded.ext,
          size=excluded.size,
          mtime=excluded.mtime,
          sha256=excluded.sha256,
          group_name=excluded.group_name
        """, (r.path, r.name, r.ext, r.size, r.mtime, r.sha256, r.group))
    conn.commit()

def pick_top(recs: List[FileRec], n: int = 60) -> List[FileRec]:
    return sorted(recs, key=lambda r: r.score, reverse=True)[:n]

def extract_identifier_text(rec: FileRec) -> str:
    # Only used for GOVERNING SET PROOF attempts; limited reading
    if rec.ext == ".pdf":
        return safe_read_pdf_text(rec.path)
    if rec.ext in (".txt", ".md", ".csv"):
        return safe_read_textlike(rec.path)
    return ""

def make_report(project: str, dropbox_root: str, out_dir: str,
                recs_scored: List[FileRec], top: List[FileRec],
                project_tokens: List[str], vendor_tokens: List[str]) -> Tuple[str, str]:
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    out_dir = os.path.expanduser(out_dir)
    os.makedirs(out_dir, exist_ok=True)

    # Historical candidates: same-name risk = project hits but archive-ish path
    historical = [r for r in top if r.project_hits > 0 and any(rt in normalize(r.path) for rt in SAME_NAME_RISK_TERMS)]
    governing_candidates = [r for r in top if r.project_hits > 0 and r.class_hint in ("drawings", "specs", "addenda")]

    # Program context detections (weight loss, etc.)
    program_context = [r for r in top if r.context_hits > 0]

    # Governing proof attempts (lightweight)
    proof = []
    for r in governing_candidates[:20]:
        sample = extract_identifier_text(r)
        # Proof check: do any project tokens appear in first pages text?
        hits = sum(1 for t in project_tokens if t.lower() in normalize(sample))
        proof.append({
            "path": r.path,
            "type": r.class_hint,
            "project_identifier_hits_in_content": hits,
            "note": "content scan limited to first pages/first chars; no scope extraction",
        })

    report_lines = []
    report_lines.append("PROJECT ARTIFACT DISCOVERY REPORT (LIVE)\n")
    report_lines.append(f"Project: {project}")
    report_lines.append(f"Timestamp: {ts}")
    report_lines.append(f"Dropbox Root: {dropbox_root}")
    report_lines.append("Mode: index everything; read-only\n")

    report_lines.append("1) PROJECT FINGERPRINT (RANKING ONLY)")
    report_lines.append(f"- Project tokens: {', '.join(project_tokens[:25])}{'...' if len(project_tokens)>25 else ''}")
    report_lines.append(f"- Vendor tokens loaded: {len([v for v in vendor_tokens if v])}")
    report_lines.append("")

    report_lines.append("2) TOP CANDIDATE ARTIFACTS (by score)")
    for r in top[:40]:
        report_lines.append(f"- score={r.score:.1f} type_hint={r.class_hint:8} group={r.group:12} hits(project={r.project_hits}, option={r.option_hits}, vendor={r.vendor_hits}, context={r.context_hits}) :: {r.path}")
    report_lines.append("")

    report_lines.append("3) GOVERNING SET PROOF (PRELIMINARY)")
    report_lines.append("Rule: no artifact can be treated as governing unless it passes proof checks.")
    for p in proof[:20]:
        report_lines.append(f"- {p['type']:8} content_project_hits={p['project_identifier_hits_in_content']} :: {p['path']}")
    report_lines.append("")

    report_lines.append("4) HISTORICAL CANDIDATES (PATTERN LIBRARY ONLY)")
    if historical:
        for r in historical[:30]:
            report_lines.append(f"- SAME-NAME RISK :: {r.path}")
    else:
        report_lines.append("- none detected in top candidates")
    report_lines.append("")

    report_lines.append("5) PROGRAM / OPERATIONAL CONTEXT (NON-GOVERNING)")
    if program_context:
        for r in program_context[:30]:
            report_lines.append(f"- context_keyword_hit :: {r.path}")
    else:
        report_lines.append("- none detected in top candidates")
    report_lines.append("")

    report_lines.append("6) NEXT AUTHORIZED STEP")
    report_lines.append("After you accept this report: REQUIRED TRADE determination + scope element extraction with citations.")
    report_txt = "\n".join(report_lines)

    safe_name = re.sub(r"[^A-Za-z0-9_-]+", "_", project)[:80]
    report_path = os.path.join(out_dir, f"{safe_name}_artifact_discovery_report.txt")
    json_path   = os.path.join(out_dir, f"{safe_name}_artifact_discovery_report.json")

    with open(report_path, "w", encoding="utf-8") as f:
        f.write(report_txt)

    payload = {
        "project": project,
        "timestamp": ts,
        "dropbox_root": dropbox_root,
        "top_candidates": [r.__dict__ for r in top],
        "governing_proof": proof,
        "historical_candidates": [r.__dict__ for r in historical],
        "program_context_candidates": [r.__dict__ for r in program_context],
        "rules": {
            "index_everything": True,
            "read_only": True,
            "historical_quarantined": True,
            "program_context_non_governing": True,
        }
    }
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2)

    return report_path, json_path

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dropbox-root", required=True)
    ap.add_argument("--project", required=True)
    ap.add_argument("--out-dir", required=True)
    ap.add_argument("--db-path", default=os.path.expanduser("~/TCL_BRAIN/dropbox_index.sqlite"))
    ap.add_argument("--vendor-list", default="")  # optional path to project vendor selection or roster
    args = ap.parse_args()

    dropbox_root = args.dropbox_root
    project = args.project

    # Optional vendor tokens (helps ranking; not authority)
    vendor_tokens: List[str] = []
    if args.vendor_list and os.path.exists(os.path.expanduser(args.vendor_list)):
        # expects a CSV/XLSX with a Vendor column somewhere; keep simple for v1
        try:
            import pandas as pd  # type: ignore
            vpath = os.path.expanduser(args.vendor_list)
            if vpath.lower().endswith(".csv"):
                df = pd.read_csv(vpath)
            else:
                df = pd.read_excel(vpath)
            for c in df.columns:
                if "vendor" in c.lower() or "sub" in c.lower():
                    vendor_tokens = [str(x).strip() for x in df[c].dropna().tolist()]
                    break
        except Exception:
            vendor_tokens = []
    vendor_tokens = [normalize(v) for v in vendor_tokens if isinstance(v, str) and v.strip()]

    project_tokens = project_tokenize(project)

    # 1) Crawl (index everything)
    recs = discover_files(dropbox_root)

    # 2) Index + hash changed/new
    os.makedirs(os.path.dirname(os.path.expanduser(args.db_path)), exist_ok=True)
    conn = db_init(os.path.expanduser(args.db_path))
    update_index(conn, recs, hash_changed_only=True)

    # 3) Score
    recs_scored = [score_file(r, project_tokens, vendor_tokens) for r in recs]

    # 4) Pick top + report
    top = pick_top(recs_scored, n=80)
    report_path, json_path = make_report(project, dropbox_root, args.out_dir, recs_scored, top, project_tokens, vendor_tokens)

    print("OK")
    print(f"Report: {report_path}")
    print(f"JSON:   {json_path}")
    print(f"DB:     {os.path.expanduser(args.db_path)}")

if __name__ == "__main__":
    main()
