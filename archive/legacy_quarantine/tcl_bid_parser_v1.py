#!/usr/bin/env python3
"""TCL Bid Parser v1 (GO B.1)

- Parses bid files (PDF/DOCX; DOC best-effort) into structured JSON objects.
- No external lookups. Warn-only.
- Output schema is stable for Command Center integration.

Usage:
  python3 tcl_bid_parser_v1.py --bids ./bids --out ./outputs/bids_parsed.json

Notes:
  - PDF parsing uses pdfplumber. DOCX uses python-docx.
  - For legacy .doc files, parsing is not reliable without additional tooling; those will be flagged.
"""

import argparse, json, re, datetime
from pathlib import Path

def normalize_whitespace(s: str) -> str:
    return re.sub(r'[ \t]+',' ', re.sub(r'\r','',s)).strip()

def extract_text(path: Path, max_pages: int = 10) -> str:
    ext = path.suffix.lower()
    if ext == '.pdf':
        try:
            import pdfplumber
        except Exception as e:
            return ''
        txt=''
        try:
            with pdfplumber.open(str(path)) as pdf:
                for i in range(min(max_pages, len(pdf.pages))):
                    txt += (pdf.pages[i].extract_text() or '') + '\n'
        except Exception:
            return ''
        return txt
    if ext == '.docx':
        try:
            from docx import Document
            doc = Document(str(path))
            return '\n'.join([p.text for p in doc.paragraphs])
        except Exception:
            return ''
    # .doc and others: unsupported in v1
    return ''

def infer_trade(filename: str, text: str) -> str:
    f=filename.lower()
    t=text.lower()
    if 'electric' in f: return 'ELECTRICAL'
    if 'mechanical' in f or 'hvac' in f: return 'MECHANICAL'
    if 'convergint' in f or 'fire alarm' in t or (' fa ' in f): return 'FIRE ALARM'
    if 'patriot' in f or 'sprinkler' in t: return 'FIRE SPRINKLER'
    if 'interiors' in f or 'gwb' in f or 'drywall' in t: return 'INTERIORS'
    if 'wells' in f or 'plumb' in t: return 'PLUMBING'
    if 'archer' in f: return 'GC/GENERAL'
    return 'UNKNOWN'

def parse_total(filename: str, text: str):
    patterns=[
        r'Proposal Total Sum:\s*\$([\d,]+\.\d{2})',
        r'TOTAL BID AMOUNT:\s*\$([\d,]+(?:\.\d{2})?)',
        r'BASE PRICE\s*\$([\d,]+(?:\.\d{2})?)',
        r'Base price.*?\$([\d,]+\.\d{2})',
        r'\bTOTAL\s*\$([\d,]+\.\d{2})\b',
        r'\bTotal\s*\$([\d,]+\.\d{2})\b',
    ]
    if 'convergint' in filename.lower():
        base=re.search(r'CONVERGINT BASE FIRE ALARM:\s*\$\s*([\d,]+)', text, re.I)
        pp=re.search(r'Plans and Permit Breakout:\s*\$\s*([\d,]+)', text, re.I)
        amt=0.0
        notes=[]
        if base:
            amt+=float(base.group(1).replace(',','')); notes.append('Base FA')
        if pp:
            amt+=float(pp.group(1).replace(',','')); notes.append('Plans/Permit')
        return (amt if amt>0 else None), notes or ['unparsed']
    for pat in patterns:
        m=re.search(pat, text, re.I|re.S)
        if m:
            return float(m.group(1).replace(',','')), [pat]
    return None, ['unparsed']

def parse_project_fields(text: str):
    lines=[normalize_whitespace(l) for l in text.splitlines() if normalize_whitespace(l)]
    project=None; address=None
    for l in lines[:250]:
        m=re.search(r'Project\s*[:\-]\s*(.+)$', l, re.I)
        if m and len(m.group(1))>2 and not project:
            project=m.group(1).strip()
        m=re.search(r'Address\s*[:\-]\s*(.+)$', l, re.I)
        if m and len(m.group(1))>5 and not address:
            address=m.group(1).strip()
    if not project:
        for l in lines[:250]:
            if re.search(r'lachini', l, re.I):
                project=l.strip()
                break
    return project, address

def parse_sections(text: str, keyword_regex: str, limit: int = 50):
    lines=[normalize_whitespace(l) for l in text.splitlines()]
    out=[]
    for l in lines:
        if re.search(keyword_regex, l, re.I):
            cleaned=re.sub(r'^[\-\*\u2022]+\s*','',l).strip()
            if cleaned:
                out.append(cleaned)
    seen=set(); res=[]
    for x in out:
        k=x.lower()
        if k not in seen:
            seen.add(k); res.append(x)
    return res[:limit]

def confidence_score(text: str) -> float:
    conf=0.0
    if re.search(r'lachini', text, re.I): conf+=0.5
    if re.search(r'woodinville', text, re.I): conf+=0.3
    if re.search(r'14490|147th', text, re.I): conf+=0.2
    return min(conf,1.0)

def build_bid_object(path: Path):
    txt=extract_text(path)
    trade=infer_trade(path.name, txt)
    total, total_notes=parse_total(path.name, txt)
    project, addr=parse_project_fields(txt)
    return {
        'file': path.name,
        'trade': trade,
        'bidder_name': path.stem,
        'project_name_raw': project,
        'project_address_raw': addr,
        'base_total': total,
        'total_parse_notes': total_notes,
        'allowances': parse_sections(txt, r'\bAllow'),
        'alternates': parse_sections(txt, r'\bAlternat'),
        'exclusions': parse_sections(txt, r'\bExclud'),
        'clarifications': parse_sections(txt, r'\bClarif|Qualification|Assumption'),
        'schedule_notes': parse_sections(txt, r'\bLead time|Schedule|Duration|Weeks?\b'),
        'confidence_score': confidence_score(txt),
        'warnings': (['UNSUPPORTED_FILETYPE'] if path.suffix.lower() not in ('.pdf','.docx') else []),
        'parsed_at': datetime.datetime.utcnow().isoformat()+'Z',
    }

def main():
    ap=argparse.ArgumentParser()
    ap.add_argument('--bids', required=True, help='Folder containing bid files')
    ap.add_argument('--out', required=True, help='Output JSON path')
    args=ap.parse_args()

    bid_dir=Path(args.bids)
    files=[p for p in bid_dir.rglob('*') if p.is_file()]
    bids=[build_bid_object(p) for p in files]
    out_path=Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps({'bids':bids}, indent=2), encoding='utf-8')
    print(f'WROTE {out_path} ({len(bids)} bids)')

if __name__=='__main__':
    main()
