import io
import re
import hashlib
from datetime import datetime
from collections import defaultdict
from pypdf import PdfReader


MONTHS = {
    "january": 1,
    "february": 2,
    "march": 3,
    "april": 4,
    "may": 5,
    "june": 6,
    "july": 7,
    "august": 8,
    "september": 9,
    "october": 10,
    "november": 11,
    "december": 12,
}


STATUS_PATTERNS = [
    "Not Verified",
    "Late Visit",
    "Schedule Visit",
    "Enter Visit",
    "Accept Visit",
]


DATE_RE = re.compile(
    r"(?P<sd>\d{1,2})\s+"
    r"(?P<sm>[A-Za-z]+)\s+"
    r"(?P<sy>\d{4})\s+"
    r"to\s+"
    r"(?P<ed>\d{1,2})\s+"
    r"(?P<em>[A-Za-z]+)\s+"
    r"(?P<ey>\d{4})\s+"
    r"(?P<hours>\d+(?:\.\d+)?)\s+"
    r"Mondelez\s+Merch\s+Support",
    re.I | re.S,
)


def normalize_space(value):
    return re.sub(
        r"\s+",
        " ",
        value or ""
    ).strip()


def parse_date(day, month, year):

    month_number = MONTHS.get(
        month.lower()
    )

    if not month_number:
        return None

    return datetime(
        int(year),
        month_number,
        int(day),
    ).date().isoformat()


def extract_text(pdf_bytes):

    reader = PdfReader(
        io.BytesIO(pdf_bytes)
    )

    pages = []

    for page in reader.pages:

        text = page.extract_text() or ""

        pages.append(text)

    return "\n".join(pages)


def split_store_address(before):

    lines = [
        normalize_space(line)
        for line in before.splitlines()
        if normalize_space(line)
    ]

    if not lines:
        return "", ""

    # Buscar la primera línea que empieza con número de calle
    address_index = None

    for i, line in enumerate(lines):

        if re.match(
            r"^\d{2,5}\s+",
            line
        ):

            address_index = i
            break

    if address_index is None:

        one_line = normalize_space(
            before
        )

        street_match = re.search(
            r"\b\d{2,5}\s+"
            r"(?:RUE|AV(?:E|ENUE)?|"
            r"BOULEVARD|BLVD|CHEMIN|CH\.?|DE|D)\b",
            one_line,
            re.I,
        )

        if street_match:

            store = one_line[
                :street_match.start()
            ].strip(" -")

            address = one_line[
                street_match.start():
            ].strip()

            return store, address

        return one_line, ""

    store = normalize_space(
        " ".join(
            lines[:address_index]
        )
    )

    address = normalize_space(
        " ".join(
            lines[address_index:]
        )
    )

    return store, address


def parse_pdf_bytes(pdf_bytes):

    text = extract_text(
        pdf_bytes
    )

    return parse_text(
        text
    )


def parse_text(text):

    if not text:
        return []

    # Limpieza de caracteres raros que a veces introduce PDF
    text = (
        text
        .replace("\u00ad", "")
        .replace("\ufffe", "-")
        .replace("\u2010", "-")
        .replace("\u2011", "-")
    )

    # MyWork usa este encabezado para cada visita
    chunks = re.split(
        r"MONDELEZ\s+RETAILERS",
        text,
        flags=re.I,
    )

    visits = []

    for raw_chunk in chunks[1:]:

        chunk = raw_chunk.strip()

        # ID de tienda
        id_match = re.match(
            r"(?P<id>\d{9})\s+(?P<rest>.*)",
            chunk,
            re.S,
        )

        if not id_match:
            continue

        source_id = id_match.group(
            "id"
        )

        rest = id_match.group(
            "rest"
        )

        # Fecha + horas
        date_match = DATE_RE.search(
            rest
        )

        if not date_match:
            continue

        week_start = parse_date(
            date_match.group("sd"),
            date_match.group("sm"),
            date_match.group("sy"),
        )

        week_end = parse_date(
            date_match.group("ed"),
            date_match.group("em"),
            date_match.group("ey"),
        )

        if not week_start or not week_end:
            continue

        hours = float(
            date_match.group(
                "hours"
            )
        )

        before = rest[
            :date_match.start()
        ].strip()

        after = rest[
            date_match.end():
        ]

        store, address = (
            split_store_address(
                before
            )
        )

        if not store:
            continue

        status = ""

        for pattern in STATUS_PATTERNS:

            if re.search(
                re.escape(pattern),
                after,
                re.I,
            ):

                status = pattern
                break

        if not status:

            status = normalize_space(
                after
            )[:80]

        visits.append(
            {
                "source_id": source_id,
                "store": store,
                "address": address,
                "hours": hours,
                "status": status,
                "week_start": week_start,
                "week_end": week_end,
            }
        )

    # --------------------------------------------------------
    # OCCURRENCE
    #
    # Walmart y Super C pueden aparecer dos veces
    # en la misma semana.
    # --------------------------------------------------------

    counts = defaultdict(int)

    for visit in visits:

        key = (
            visit["week_start"],
            visit["week_end"],
            visit["source_id"],
        )

        counts[key] += 1

        visit["occurrence"] = (
            counts[key]
        )

        visit["visit_key"] = (
            f"{visit['week_start']}|"
            f"{visit['week_end']}|"
            f"{visit['source_id']}|"
            f"{visit['occurrence']}"
        )

    return visits


def sha256_bytes(data):

    return hashlib.sha256(
        data
    ).hexdigest()


def compare_visits(
    old_visits,
    new_visits,
):

    new_weeks = {
        (
            v["week_start"],
            v["week_end"]
        )
        for v in (
            new_visits or []
        )
    }

    old_scoped = [
        v
        for v in (
            old_visits or []
        )
        if (
            v["week_start"],
            v["week_end"]
        )
        in new_weeks
    ]

    old = {
        v["visit_key"]: v
        for v in old_scoped
    }

    new = {
        v["visit_key"]: v
        for v in (
            new_visits or []
        )
    }

    added = [
        new[k]
        for k in (
            new.keys()
            - old.keys()
        )
    ]

    removed = [
        old[k]
        for k in (
            old.keys()
            - new.keys()
        )
    ]

    changed = []

    for key in (
        old.keys()
        & new.keys()
    ):

        diffs = {}

        for field in (
            "hours",
            "status",
            "store",
            "address",
        ):

            if (
                old[key].get(field)
                !=
                new[key].get(field)
            ):

                diffs[field] = {
                    "old": old[key].get(
                        field
                    ),
                    "new": new[key].get(
                        field
                    ),
                }

        if diffs:

            changed.append(
                {
                    "visit": new[key],
                    "diffs": diffs,
                }
            )

    return {
        "added": sorted(
            added,
            key=lambda v: (
                v["week_start"],
                v["source_id"],
                v["occurrence"],
            ),
        ),

        "removed": sorted(
            removed,
            key=lambda v: (
                v["week_start"],
                v["source_id"],
                v["occurrence"],
            ),
        ),

        "changed": sorted(
            changed,
            key=lambda x: (
                x["visit"]["week_start"],
                x["visit"]["source_id"],
                x["visit"]["occurrence"],
            ),
        ),
    }
