import io, re, hashlib
from datetime import datetime
from collections import defaultdict
from pypdf import PdfReader

MONTHS = {
    'january':1,'february':2,'march':3,'april':4,'may':5,'june':6,
    'july':7,'august':8,'september':9,'october':10,'november':11,'december':12
}
STATUS_PATTERNS = [
    'Not Verified','Late Visit','Schedule Visit','Enter Visit','Accept Visit'
]

STREET_RE = re.compile(r"\b(\d{2,5})\s+(RUE|AV(?:E|ENUE)?|BOULEVARD|BLVD|CHEMIN|CH\.?|DE\b|D\b)", re.I)
DATE_RE = re.compile(
    r"(?P<sd>\d{1,2})\s+(?P<sm>[A-Za-z]+)\s+(?P<sy>\d{4})\s+to\s+"
    r"(?P<ed>\d{1,2})\s+(?P<em>[A-Za-z]+)\s+(?P<ey>\d{4})\s+"
    r"(?P<hours>\d+(?:\.\d+)?)\s+Mondelez Merch Support",
    re.I | re.S,
)


def _date(d, m, y):
    return datetime(int(y), MONTHS[m.lower()], int(d)).date().isoformat()


def normalize_space(s: str) -> str:
    return re.sub(r"\s+", " ", s or "").strip()


def extract_text(pdf_bytes: bytes) -> str:
    reader = PdfReader(io.BytesIO(pdf_bytes))
    return "\n".join((p.extract_text() or "") for p in reader.pages)


def parse_pdf_bytes(pdf_bytes: bytes):
    text = extract_text(pdf_bytes)
    return parse_text(text)


def parse_text(text: str):
    # Preserve original line breaks where useful, but normalize odd PDF hyphenation.
    text = text.replace('\u00ad', '').replace('\ufffe', '-')
    chunks = text.split('MONDELEZ RETAILERS')[1:]
    visits = []
    for chunk in chunks:
        chunk = chunk.strip()
        m_id = re.match(r"(?P<id>\d{9})\s+(?P<rest>.*)", chunk, re.S)
        if not m_id:
            continue
        source_id = m_id.group('id')
        rest = m_id.group('rest')
        m_date = DATE_RE.search(rest)
        if not m_date:
            continue
        before = rest[:m_date.start()].strip()
        after = rest[m_date.end():]
        start = _date(m_date.group('sd'), m_date.group('sm'), m_date.group('sy'))
        end = _date(m_date.group('ed'), m_date.group('em'), m_date.group('ey'))
        hours = float(m_date.group('hours'))

        # Split store name vs address using the first actual street number/keyword.
        one_line = normalize_space(before)
        street = STREET_RE.search(one_line)
        if street:
            store = one_line[:street.start()].strip(' -')
            address = one_line[street.start():].strip()
        else:
            lines = [normalize_space(x) for x in before.splitlines() if normalize_space(x)]
            idx = next((i for i,l in enumerate(lines) if re.match(r"\d{2,5}\s+", l)), None)
            if idx is None:
                store, address = one_line, ''
            else:
                store = normalize_space(' '.join(lines[:idx]))
                address = normalize_space(' '.join(lines[idx:]))

        status = ''
        for pat in STATUS_PATTERNS:
            if re.search(re.escape(pat), after, re.I):
                status = pat
                break
        if not status:
            # fallback: text following closing project parenthesis
            tail = normalize_space(after)
            status = tail[-40:].strip()

        visits.append({
            'source_id': source_id,
            'store': store,
            'address': address,
            'hours': hours,
            'status': status,
            'week_start': start,
            'week_end': end,
        })

    # occurrence index is necessary because the same store can appear twice in one week.
    counts = defaultdict(int)
    for v in visits:
        key = (v['week_start'], v['week_end'], v['source_id'])
        counts[key] += 1
        v['occurrence'] = counts[key]
        v['visit_key'] = f"{v['week_start']}|{v['week_end']}|{v['source_id']}|{v['occurrence']}"
    return visits


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def compare_visits(old_visits, new_visits):
    # MyWork puede dejar de incluir semanas ya vencidas en un PDF nuevo.
    # No tratamos esas semanas completas como "visitas eliminadas".
    new_weeks = {(v['week_start'], v['week_end']) for v in (new_visits or [])}
    old_scoped = [v for v in (old_visits or []) if (v['week_start'], v['week_end']) in new_weeks]
    old = {v['visit_key']: v for v in old_scoped}
    new = {v['visit_key']: v for v in (new_visits or [])}
    added = [new[k] for k in new.keys() - old.keys()]
    removed = [old[k] for k in old.keys() - new.keys()]
    changed = []
    for k in old.keys() & new.keys():
        diffs = {}
        for field in ('hours','status','store','address'):
            if old[k].get(field) != new[k].get(field):
                diffs[field] = {'old': old[k].get(field), 'new': new[k].get(field)}
        if diffs:
            changed.append({'visit': new[k], 'diffs': diffs})
    return {
        'added': sorted(added, key=lambda v:(v['week_start'],v['source_id'],v['occurrence'])),
        'removed': sorted(removed, key=lambda v:(v['week_start'],v['source_id'],v['occurrence'])),
        'changed': sorted(changed, key=lambda x:(x['visit']['week_start'],x['visit']['source_id'],x['visit']['occurrence'])),
    }
