"""
generate_news_format.py

Parses the current Transfer Management PDF (photo-grid visual layout) and the
previous Transfer News PDF (table format), cross-references them, and produces
a 'News Format' DataFrame for the next Transfer News Word document.

Pipeline logic (matches how the mission prints the news):
  1. Every missionary in Transfer Management is looked up in the previous
     Transfer News by last name (area/zone-aware, so duplicate last names like
     MENSAH vs MENSAH.J resolve to the right person).
  2. Found          -> row with current zone/area/companion, plus a Status
                       classification (STAYING / AREA CHANGE / ZONE TRANSFER,
                       with NEW LEADERSHIP / RELEASED notes).
  3. Not found, but the companion's previous row said "NEW MISSIONARY"
                    -> they arrived last transfer masked and are now revealed:
                       they get their own row (JC), title inferred from their
                       companion.
  4. Not found, and unexplained (or named in the INCOMING section)
                    -> genuinely new this transfer: no row of their own, and
                       the trainer's companion cell shows "NEW MISSIONARY".

Assignment codes are derived from the Transfer Management badges plus the
training context:
  - DL whose junior is new/still-in-training -> DT (district leader training)
  - non-DL senior with a T marker on an unmatched junior, or with a masked
    new junior -> TR
  - ZL/STL/AP pairs -> ZL1/ZL2, STL1/STL2, AP1/AP2
  - juniors -> JC (unless part of a ZL2/STL2/AP2/SA pair)
  - SA (special assignment) is inherited from the previous news when the
    pair carries no badge.

Consumed by data/generate_transfer_sheet.py when a previous_news_pdf is
supplied.
"""
from __future__ import annotations

import logging
import re
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import pandas as pd
import pdfplumber

logger = logging.getLogger(__name__)

# ── Constants ──────────────────────────────────────────────────────────────────

# Badge codes that appear above missionary photos in Transfer Management
TM_BADGE_RE = re.compile(r"^(ZL1|ZL2|ZL|STL1|STL2|STL|AP1|AP2|AP|DL|DT|TR|SC|JC)$")

# Canonical zone names, keyed by alphanumeric-only uppercase form
ZONE_CANON: dict[str, str] = {
    "UYO": "UYO",
    "UYOCENTRAL": "UYO CENTRAL",
    "UDOUMANA": "UDO UMANA",
    "ETINAN": "ETINAN",
    "ETINANNORTH": "ETINAN NORTH",
    "IBIONO": "IBIONO",
    "EKET": "EKET",
    "IBESIKPO": "IBESIKPO",
    "ORON": "ORON",
    "UKATARAN": "UKAT ARAN",
    "UKATNSIT": "UKAT-NSIT",
    "IKOTUSEKONG": "IKOT USEKONG",
    "ITAMNKEMBA": "ITAM-NKEMBA",
}

# Zone heading pattern in the Transfer News PDF. Generic on purpose: any
# "01 SOMETHING ZONE" heading counts, so brand-new zones are picked up
# automatically without code changes.
TN_ZONE_RE = re.compile(r"^\d{0,2}\s*([A-Z][A-Z0-9&.\- ]*?)\s+ZONE\b", re.IGNORECASE)

# Transfer News column header fragments (old format uses NAME/ASSIGNMENT,
# the new printed format uses MISSIONARY/ROLE — both are accepted).
TN_COL_NAME = ("MISSIONARY", "NAME")
TN_COL_ASSIGN = ("ASSIGN", "ROLE")
TN_COL_ZONE = ("ZONE",)
TN_COL_AREA = ("AREA",)
TN_COL_COMPANION = ("COMPANION",)

MONTHS = [
    "JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE",
    "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER",
]

# Vertical anchors of the two district blocks on each TM zone page.
# Each block: district header, area band (may contain T training markers),
# badge-code row, name row.
TM_BLOCKS = [
    {"district": (58, 80), "area": (80, 118), "codes": (118, 130), "names": (130, 145)},
    {"district": (178, 200), "area": (200, 237), "codes": (237, 249), "names": (249, 264)},
]

# Horizontal offset between a badge code and the photo/name it belongs to
BADGE_X_OFFSET = 14

LEADER_EXEMPT = {"JC", "SC", ""}


def _is_leader(assignment: str) -> bool:
    return str(assignment or "").strip().upper() not in LEADER_EXEMPT


# ── Data classes ───────────────────────────────────────────────────────────────

@dataclass
class MissionaryRecord:
    display_name: str    # As shown in the previous TN, e.g. "ELDER MENSAH.J"
    title: str           # "ELDER" or "SISTER"
    last_name: str       # Normalized: "MENSAH.J" or "SMITH"
    assignment: str
    section_zone: str    # Zone section they appear under (their zone before last transfer)
    assigned_zone: str   # Value of the previous news ZONE column (their zone last transfer)
    area: str
    companion: str
    zone_order: int = 0
    row_order: int = 0


@dataclass
class TMMissionary:
    name: str            # Last name as printed in Transfer Management
    x: float             # x-position of the name (for badge/T association)
    badge: str = ""      # Raw badge code above the photo, if any
    has_t: bool = False  # T (training) marker on this missionary
    # Filled in by the matcher:
    prev: MissionaryRecord | None = None
    state: str = "unmatched"   # matched | revealed | masked | unmatched


@dataclass
class AreaRecord:
    zone: str
    district: str
    area: str
    members: list[TMMissionary] = field(default_factory=list)

    @property
    def senior(self) -> TMMissionary | None:
        return self.members[0] if self.members else None

    @property
    def junior(self) -> TMMissionary | None:
        """First companion after the senior — kept for call sites that only
        care about a single partner; N-member groups should use .members."""
        return self.members[1] if len(self.members) > 1 else None


# ── Utility helpers ────────────────────────────────────────────────────────────

def _norm(name: str) -> str:
    return str(name or "").strip().upper()


def _norm_key(text: str) -> str:
    """Alphanumeric-only uppercase form for fuzzy comparison."""
    return re.sub(r"[^A-Z0-9]", "", _norm(text))


# Words that can precede "ZONE"/appear in zone cells but are never zone names
_ZONE_BLACKLIST = {
    "NEW", "EXISTING", "MISSIONARY", "NAME", "ASSIGNMENT", "ROLE", "AREA",
    "COMPANION", "INDEX", "ALL", "ZONES", "NAMEOFMISSIONARYASSIGNMENT",
}


def _canon_zone(text: str, registry: dict[str, str] | None = None) -> str:
    """Canonical display name for a zone. Unknown zones are accepted as-is
    (cleaned), so brand-new zones work without any code changes; near-misses
    of known zones (typos like UKAT URAN for UKAT ARAN) snap to the known
    spelling instead of creating a phantom zone."""
    key = _norm_key(text)
    if not key or key in _ZONE_BLACKLIST or any(w in key for w in ("MISSIONARY", "ASSIGNMENT", "COMPANION")):
        return ""
    if registry and key in registry:
        return registry[key]
    if key in ZONE_CANON:
        return ZONE_CANON[key]

    known: dict[str, str] = dict(ZONE_CANON)
    if registry:
        known.update(registry)
    if len(key) >= 6:
        from difflib import SequenceMatcher

        best = max(known, key=lambda k: SequenceMatcher(None, key, k).ratio(), default=None)
        if best and SequenceMatcher(None, key, best).ratio() >= 0.85:
            return known[best]

    cleaned = re.sub(r"\s+", " ", _norm(text)).strip()
    if len(key) < 3 or TM_BADGE_RE.match(cleaned):
        return ""
    return cleaned


def _base_name(key: str) -> str:
    """MENSAH.J -> MENSAH (suffix used by the mission to disambiguate)."""
    return key.split(".")[0]


def _area_match(a: str, b: str) -> bool:
    na, nb = _norm_key(a), _norm_key(b)
    if not na or not nb:
        return False
    return na == nb or (len(na) >= 5 and na in nb) or (len(nb) >= 5 and nb in na)


def _slot_area_match(a: str, b: str) -> bool:
    """Looser area comparison used to link revealed trainees to last
    transfer's masked slots — also tolerates letter transpositions like
    OKIOTA vs OKOITA (same letters, different order)."""
    if _area_match(a, b):
        return True
    la = "".join(sorted(re.sub(r"[^A-Z]", "", _norm(a))))
    lb = "".join(sorted(re.sub(r"[^A-Z]", "", _norm(b))))
    return len(la) >= 5 and la == lb


def _is_new_missionary_text(text: str) -> bool:
    return "NEWMISSIONARY" in _norm_key(text)


_COMPANION_SPLIT_RE = re.compile(r"\s*(?:&|,|\bAND\b)\s*")


def _split_companions(text: str) -> list[str]:
    """Extract each companion's last name from a companion cell that may
    name one person ("ELDER FOO"), or a trio/foursome ("ELDER FOO & ELDER
    BAR", or the older "ELDERS FOO AND BAR" form)."""
    names: list[str] = []
    for piece in _COMPANION_SPLIT_RE.split(str(text or "").strip()):
        piece = piece.strip()
        if not piece:
            continue
        m = re.match(r"^(?:ELDERS?|SISTERS?)\s*(.+)$", piece)
        name = m.group(1).strip() if m else piece
        if name and not _is_new_missionary_text(name):
            names.append(name)
    return names


# ── Character-stream helpers (Transfer Management) ────────────────────────────
#
# Two names/areas can overlap horizontally in the TM export, so x-sorted word
# extraction interleaves their letters ("Andrew"+"Ampofo" -> "AndreAwmpofo").
# page.chars preserves content-stream order, where each text object stays
# contiguous, so segmenting on backward x-jumps recovers the real strings.

def _band_chars(chars: list[dict], lo: float, hi: float) -> list[dict]:
    return [c for c in chars if lo <= c["top"] < hi]


def _stream_segments(chars: list[dict], gap: float = 4.0) -> list[list[dict]]:
    segments: list[list[dict]] = []
    current: list[dict] = []
    for c in chars:
        if current:
            p = current[-1]
            new_line = abs(c["top"] - p["top"]) > 3
            back_jump = c["x0"] < p["x1"] - 1.0
            too_far = c["x0"] - p["x1"] > gap
            if new_line or back_jump or too_far:
                segments.append(current)
                current = []
        current.append(c)
    if current:
        segments.append(current)
    return segments


def _segment_text(segment: list[dict]) -> str:
    return "".join(c["text"] for c in segment)


_CAMEL_BOUNDARY = re.compile(r"(?<=[a-z])(?=[A-Z])|(?<=[A-Z])(?=[A-Z][a-z])")


def _name_tokens(chars: list[dict]) -> list[tuple[str, float]]:
    """Split a name row into (last_name, x0) tokens, in x order."""
    tokens: list[tuple[str, float]] = []
    for segment in _stream_segments(chars, gap=4.0):
        text = _segment_text(segment)
        boundaries = [0] + [m.start() for m in _CAMEL_BOUNDARY.finditer(text)] + [len(text)]
        for start, end in zip(boundaries, boundaries[1:]):
            token = text[start:end].strip()
            if len(token) > 1:
                tokens.append((token, float(segment[start]["x0"])))
    tokens.sort(key=lambda t: t[1])
    return tokens


def _footer_free_chars(page) -> list[dict]:
    return [c for c in page.chars if 28 < c["top"] < 760]


# ── INCOMING section (page 1 of the TM export) ────────────────────────────────

def parse_incoming(pdf_path: Path) -> set[str]:
    """Return normalized last names from the INCOMING section on page 1."""
    incoming: set[str] = set()
    stop_words = ("UNASSIGNED", "CUSTOM", "ASSIGNMENT")

    with pdfplumber.open(str(pdf_path)) as pdf:
        chars = _footer_free_chars(pdf.pages[0])
        rows: dict[int, list[dict]] = defaultdict(list)
        for c in chars:
            rows[round(c["top"] / 6) * 6].append(c)

        in_section = False
        for y in sorted(rows):
            text = " ".join(
                _segment_text(seg) for seg in _stream_segments(rows[y], gap=4.0)
            ).strip()
            upper = text.upper()
            if "INCOMING" in upper:
                in_section = True
                continue
            if in_section:
                if any(upper.startswith(w) for w in stop_words):
                    break
                for seg in _stream_segments(rows[y], gap=4.0):
                    seg_text = _segment_text(seg)
                    boundaries = [0] + [m.start() for m in _CAMEL_BOUNDARY.finditer(seg_text)] + [len(seg_text)]
                    for start, end in zip(boundaries, boundaries[1:]):
                        token = seg_text[start:end].strip()
                        if len(token) > 1:
                            incoming.add(_norm(token))

    logger.info("INCOMING missionaries: %s", sorted(incoming))
    return incoming


# ── TM zone pages ──────────────────────────────────────────────────────────────
#
# Each area is printed as a bordered photo card that can hold 2, 3 or more
# missionaries (companion trios/foursomes are real), or zero (a closed area
# still prints its name in the area band but has no photo at all). A card's
# outer border can visually overlap its neighbor's when its own area name is
# long (the export doesn't reflow), so card *borders* aren't a reliable
# boundary. Each area's own label always starts almost exactly where its
# first missionary's name starts, though (confirmed empirically: within ~1.5pt
# across ordinary pairs, trios and long compound names alike) — so the
# label x-positions themselves are the boundaries: sort them left to right and
# assign every name to the closest-preceding label, exactly like the mission's
# own layout groups them. A label with no name in its range is a closed area.

def _parse_block(chars: list[dict], block: dict, zone: str) -> list[AreaRecord]:
    name_chars = _band_chars(chars, *block["names"])
    tokens = _name_tokens(name_chars)
    if not tokens:
        return []

    district_chars = _band_chars(chars, *block["district"])
    district = " ".join(
        _segment_text(seg).strip() for seg in _stream_segments(district_chars, gap=6.0)
    ).strip()

    missionaries = [TMMissionary(name=t[0], x=t[1]) for t in tokens]

    def nearest(target_x: float) -> TMMissionary:
        return min(missionaries, key=lambda m: abs(m.x - target_x))

    # Badge codes sit slightly right of the name start; T markers sit above the
    # trainee's photo.
    for segment in _stream_segments(_band_chars(chars, *block["codes"]), gap=2.5):
        text = _segment_text(segment).strip()
        x0 = float(segment[0]["x0"])
        if TM_BADGE_RE.match(text):
            owner = nearest(x0 - BADGE_X_OFFSET)
            if not owner.badge:
                owner.badge = text
        elif text == "T":
            nearest(x0).has_t = True

    # Area band: single 'T' segments are training markers; every other
    # segment is one area's full label (compound names with internal spaces
    # or slashes still come through as one segment — the gaps inside them are
    # smaller than the gap to the next area's label).
    labels: list[tuple[float, str]] = []
    for segment in _stream_segments(_band_chars(chars, *block["area"]), gap=5.0):
        text = _segment_text(segment).strip()
        if not text:
            continue
        if text == "T":
            nearest(float(segment[0]["x0"])).has_t = True
            continue
        labels.append((float(segment[0]["x0"]), text))

    if not labels:
        # No area labels at all on this row — one catch-all area rather than
        # silently dropping every missionary in the block.
        members = sorted(missionaries, key=lambda m: m.x)
        return [AreaRecord(zone=zone, district=district, area="UNKNOWN", members=members)]

    labels.sort(key=lambda p: p[0])
    boundaries = [x0 - 5 for x0, _ in labels]

    def label_index(x0: float) -> int:
        index = 0
        for i, boundary in enumerate(boundaries):
            if x0 >= boundary:
                index = i
        return index

    area_members: dict[int, list[TMMissionary]] = defaultdict(list)
    for m in missionaries:
        area_members[label_index(m.x)].append(m)

    records: list[AreaRecord] = []
    for i, (_x0, text) in enumerate(labels):
        area = re.sub(r"\s+", " ", text)
        area = re.sub(r"-\s+", "-", area).strip().strip("/").strip()
        members = sorted(area_members.get(i, []), key=lambda m: m.x)
        records.append(AreaRecord(zone=zone, district=district, area=area or "UNKNOWN", members=members))

    return records


def parse_tm_zones(pdf_path: Path) -> list[AreaRecord]:
    """Parse the zone pages of the Transfer Management PDF."""
    all_areas: list[AreaRecord] = []

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages[1:]:
            chars = _footer_free_chars(page)
            if not chars:
                continue

            header_chars = [c for c in chars if c["top"] < 50]
            header_text = "".join(
                _segment_text(seg) + " " for seg in _stream_segments(sorted(header_chars, key=lambda c: c["x0"]), gap=4.0)
            ).strip()
            if "ZONE" not in header_text.upper():
                continue
            zone_raw = re.sub(r"\s*Zone\s*$", "", header_text, flags=re.IGNORECASE).strip()
            zone = _canon_zone(zone_raw) or _norm(zone_raw)

            for block in TM_BLOCKS:
                all_areas.extend(_parse_block(chars, block, zone))

    logger.info("Parsed %d areas from Transfer Management", len(all_areas))
    return all_areas


# ── Parse the previous Transfer News PDF ──────────────────────────────────────

def parse_zone_sections(pdf_path: Path) -> list[str]:
    """Pre-pass: every zone name mentioned as a heading or zone-index entry, in
    document order. This is the canonical zone ORDER carried from one transfer
    news to the next (requirement: never re-sort zones)."""
    finder = re.compile(r"([A-Z][A-Z0-9&.\- ]*?)\s+ZONE\b")
    sections: list[str] = []
    seen: set[str] = set()

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            text = (page.extract_text() or "").upper()
            for line in text.splitlines():
                for m in finder.finditer(line):
                    name = re.sub(r"^\d+\s*", "", re.sub(r"\s+", " ", m.group(1))).strip()
                    canon = _canon_zone(name)
                    key = _norm_key(canon)
                    if not key or key in seen:
                        continue
                    seen.add(key)
                    sections.append(canon)

    logger.info("Zone order from previous news: %s", sections)
    return sections


def parse_transfer_title(tm_pdf: Path) -> str:
    """Derive the transfer title (e.g. 'JUNE / JULY 2026 TRANSFER NEWS') from
    the month/year printed on page 1 of the Transfer Management export."""
    try:
        with pdfplumber.open(str(tm_pdf)) as pdf:
            text = (pdf.pages[0].extract_text() or "").upper()
    except Exception:
        return ""
    m = re.search(r"\b(" + "|".join(MONTHS) + r")\s+(20\d{2})\b", text)
    if m:
        month, year = m.group(1), int(m.group(2))
        prev_month = MONTHS[(MONTHS.index(month) - 1) % 12]
        return f"{prev_month} / {month} {year} TRANSFER NEWS"

    # Some iMOS exports print only "August Transfer" and omit the year.
    # In that format the printed month is the complete transfer label, not
    # the second half of a two-month range. Use the source file's timestamp
    # for the year so hosted runs do not accidentally inherit the server's
    # current month or timezone.
    month_only = re.search(r"\b(" + "|".join(MONTHS) + r")\s+TRANSFER\b", text)
    if month_only:
        try:
            year = datetime.fromtimestamp(tm_pdf.stat().st_mtime).year
        except OSError:
            year = datetime.now().year
        return f"{month_only.group(1)} {year} TRANSFER NEWS"
    return ""


def parse_previous_news(
    pdf_path: Path,
) -> tuple[dict[str, list[MissionaryRecord]], dict[str, list[MissionaryRecord]], list[str]]:
    """
    Parse the previous Transfer News PDF (table format).

    Returns (records, companion_mentions, section_zones):
      records: normalized_last_name -> their own row(s).
      companion_mentions: normalized_last_name -> row(s) whose companion column
        named this person (used to recover people whose own row was missed).
      section_zones: zone names in document order (canonical zone order).
    """
    records: dict[str, list[MissionaryRecord]] = defaultdict(list)
    companion_mentions: dict[str, list[MissionaryRecord]] = defaultdict(list)
    section_zones = parse_zone_sections(pdf_path)
    zone_registry = {_norm_key(z): z for z in section_zones}
    current_zone = "UNKNOWN"
    zone_order = 0
    row_order = 0
    column_map: dict[str, int | None] | None = None

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            words = page.extract_words(keep_blank_chars=False)
            rows_by_y: dict[float, list[dict]] = defaultdict(list)
            for w in words:
                rows_by_y[round(w["top"] / 6) * 6].append(w)

            # Map y -> zone name for section headings on this page
            zone_y: dict[float, str] = {}
            for y, row_words in sorted(rows_by_y.items()):
                text = " ".join(w["text"] for w in sorted(row_words, key=lambda w: w["x0"])).strip()
                m = TN_ZONE_RE.match(text)
                if m:
                    zone_y[y] = _canon_zone(m.group(1), zone_registry) or m.group(1).upper()

            for tbl_obj in page.find_tables():
                tbl_top = tbl_obj.bbox[1]

                headings_above = [(y, z) for y, z in zone_y.items() if y <= tbl_top]
                if headings_above:
                    _, current_zone = max(headings_above, key=lambda t: t[0])
                    zone_order += 1

                table = tbl_obj.extract()
                if not table:
                    continue

                # Find header row; some PDFs print it once and continue the
                # roster with headerless tables — reuse the last layout then.
                header_idx = None
                for i, row in enumerate(table[:4]):
                    upper = [str(c or "").upper().strip() for c in row]
                    has_name = any(f in c for c in upper for f in TN_COL_NAME)
                    has_assign = any(f in c for c in upper for f in TN_COL_ASSIGN)
                    if has_name and has_assign:
                        header_idx = i
                        break

                if header_idx is not None:
                    headers = [str(c or "").upper().strip() for c in table[header_idx]]

                    def col(fragments: tuple[str, ...]) -> int | None:
                        return next(
                            (i for i, h in enumerate(headers) if any(f in h for f in fragments)),
                            None,
                        )

                    column_map = {
                        "name": col(TN_COL_NAME),
                        "assign": col(TN_COL_ASSIGN),
                        "zone": col(TN_COL_ZONE),
                        "area": col(TN_COL_AREA),
                        "comp": col(TN_COL_COMPANION),
                    }
                    data_rows = table[header_idx + 1 :]
                else:
                    if column_map is None:
                        # Assume the document's fixed column order.
                        column_map = {"name": 0, "assign": 1, "zone": 2, "area": 3, "comp": 4}
                    data_rows = table

                ci_name = column_map["name"]
                if ci_name is None:
                    continue
                ci_assign = column_map["assign"]
                ci_zone = column_map["zone"]
                ci_area = column_map["area"]
                ci_comp = column_map["comp"]

                for row in data_rows:
                    if not row:
                        continue
                    name_raw = str(row[ci_name] if ci_name is not None and ci_name < len(row) else "").strip().upper()
                    if not name_raw or len(name_raw) < 3:
                        continue

                    # Cells can render with no space between title and name
                    # ("ELDERISAAC") and wrap hyphenated names across lines.
                    name_clean = re.sub(r"\s+", " ", name_raw.replace("-\n", "-").replace("\n", " ")).strip()
                    match = re.match(r"^(ELDER|SISTER)\s*(.+)$", name_clean)
                    if not match:
                        # Not a missionary row (cover tables, stat tiles,
                        # assignment key, ...) — every roster row starts with
                        # ELDER or SISTER.
                        continue
                    title, last_name = match.group(1), match.group(2).strip()
                    if last_name in ("", "#", "MISSIONARY"):
                        continue
                    display_name = f"{title} {last_name}"

                    def cell(index: int | None) -> str:
                        if index is None or index >= len(row):
                            return ""
                        return re.sub(r"\s+", " ", str(row[index] or "")).strip().upper()

                    assignment = cell(ci_assign)
                    assigned_zone = _canon_zone(cell(ci_zone), zone_registry)
                    area = cell(ci_area)
                    companion = cell(ci_comp)

                    rec = MissionaryRecord(
                        display_name=display_name,
                        title=title,
                        last_name=last_name,
                        assignment=assignment,
                        section_zone=current_zone,
                        assigned_zone=assigned_zone,
                        area=area,
                        companion=companion,
                        zone_order=zone_order,
                        row_order=row_order,
                    )
                    records[_norm(last_name)].append(rec)
                    row_order += 1

                    # Index every companion mentioned in the cell — trios and
                    # foursomes print as "ELDER X & ELDER Y" (or the older
                    # "ELDERS X AND Y" form), not just a single name.
                    for companion_last in _split_companions(companion):
                        companion_mentions[_norm(companion_last)].append(rec)

    logger.info("Parsed %d unique last names from previous Transfer News", len(records))
    return dict(records), dict(companion_mentions), section_zones


# ── Cross-reference: global matching ──────────────────────────────────────────

def _candidate_keys(name: str, prev: dict[str, list[MissionaryRecord]], base_index: dict[str, list[str]]) -> list[str]:
    key = _norm(name)
    keys = []
    if key in prev:
        keys.append(key)
    for other in base_index.get(_base_name(key), []):
        if other != key:
            keys.append(other)
    return keys


def _match_all(
    areas: list[AreaRecord],
    prev: dict[str, list[MissionaryRecord]],
    companion_mentions: dict[str, list[MissionaryRecord]],
    incoming: set[str],
) -> set[int]:
    """Resolve every TM missionary against the previous news (each previous
    record is consumed at most once, so duplicate last names distribute
    correctly). Sets .prev and .state on each TMMissionary in place and
    returns the ids of the consumed previous records (unconsumed records
    belong to missionaries no longer in the mission)."""
    base_index: dict[str, list[str]] = defaultdict(list)
    for key in prev:
        base_index[_base_name(key)].append(key)

    consumed: set[int] = set()
    people: list[tuple[TMMissionary, AreaRecord, bool]] = []
    for area in areas:
        for i, member in enumerate(area.members):
            people.append((member, area, i > 0))

    # Juniors named in the INCOMING section are new this transfer, full stop.
    for person, area, is_junior in people:
        if is_junior and _norm(person.name) in incoming:
            person.state = "masked"

    def candidates(person: TMMissionary) -> list[MissionaryRecord]:
        found: list[MissionaryRecord] = []
        for key in _candidate_keys(person.name, prev, base_index):
            found.extend(r for r in prev[key] if id(r) not in consumed)
        return found

    def take(person: TMMissionary, rec: MissionaryRecord) -> None:
        person.prev = rec
        person.state = "matched"
        consumed.add(id(rec))

    # Pass A: same-area matches first (pins down duplicate last names).
    for person, area, _ in people:
        if person.state != "unmatched":
            continue
        area_hits = [r for r in candidates(person) if _area_match(r.area, area.area)]
        if area_hits:
            exact = [r for r in area_hits if _norm(r.last_name) == _norm(person.name)]
            take(person, (exact or area_hits)[0])

    # Pass B/C: remaining candidates, preferring exact key and zone agreement.
    for person, area, _ in people:
        if person.state != "unmatched":
            continue
        cands = candidates(person)
        if not cands:
            continue

        def score(rec: MissionaryRecord) -> tuple[int, int]:
            exact = int(_norm(rec.last_name) == _norm(person.name))
            prev_zone = rec.assigned_zone or _canon_zone(rec.section_zone)
            zone_hit = int(bool(prev_zone) and prev_zone == area.zone)
            return (exact, zone_hit)

        take(person, max(cands, key=score))

    # Recover seniors whose own previous row is missing from the previous news
    # but who were named in someone's companion cell. Juniors are excluded:
    # an unknown junior is handled by the masked/revealed logic below, and a
    # companion mention there is usually a different missionary with the same
    # last name.
    for person, area, is_junior in people:
        if person.state != "unmatched" or is_junior:
            continue
        mentions = companion_mentions.get(_norm(person.name), [])
        if mentions:
            mention = mentions[0]
            person.prev = MissionaryRecord(
                display_name=f"{mention.title} {_norm(person.name)}",
                title=mention.title,
                last_name=_norm(person.name),
                assignment="",
                section_zone=mention.section_zone,
                assigned_zone=mention.assigned_zone,
                area=mention.area,
                companion=mention.last_name,
                zone_order=mention.zone_order,
                row_order=mention.row_order,
            )
            person.state = "matched"

    # Classify what's left. A junior with no previous record is either a
    # missionary who arrived last transfer masked as "NEW MISSIONARY" (now
    # revealed, so they get their own row) or someone genuinely new this
    # transfer (still masked). They count as revealed when either their
    # senior's previous companion cell said "NEW MISSIONARY", or one of last
    # transfer's masked slots was in the area they now serve in.
    masked_slots = [
        rec.area
        for recs in prev.values()
        for rec in recs
        if _is_new_missionary_text(rec.companion)
    ]
    for area in areas:
        if not area.members:
            continue  # closed area — no one to classify
        senior, juniors = area.members[0], area.members[1:]
        for junior in juniors:
            if junior.state != "unmatched":
                continue
            senior_had_new = senior.prev is not None and _is_new_missionary_text(senior.prev.companion)
            slot_here = any(_slot_area_match(slot, area.area) for slot in masked_slots)
            if senior_had_new or slot_here:
                junior.state = "revealed"
            else:
                junior.state = "masked"
                if junior.has_t and senior.prev is not None and _area_match(senior.prev.area, area.area):
                    logger.warning(
                        "Junior of %s in %s / %s treated as NEW MISSIONARY, but the "
                        "senior stayed in the same area — verify whether this is a "
                        "continuing trainee who should be revealed by name (%s)",
                        senior.name, area.zone, area.area, junior.name,
                    )

        if senior.state == "unmatched":
            # A senior always gets a row: they were revealed some transfer ago
            # even if the previous news never printed them.
            senior.state = "revealed"
            logger.warning(
                "Senior %s (%s / %s) not found in previous news — "
                "verify the printed name/title",
                senior.name, area.zone, area.area,
            )
        for junior in juniors:
            if junior.state == "masked":
                logger.info(
                    "New missionary (masked): junior of %s in %s / %s",
                    senior.name, area.zone, area.area,
                )

    return consumed


# ── Assignment derivation ──────────────────────────────────────────────────────

def _family(badge: str) -> str:
    badge = _norm(badge)
    for fam in ("ZL", "STL", "AP"):
        if badge.startswith(fam):
            return fam
    return badge


def _group_assignments(area: AreaRecord) -> list[str]:
    """One assignment code per member of area.members (senior first). Same
    rules as a 2-person pair, generalized to N companions: every non-senior
    member is judged independently against the training/masked/revealed
    logic, and any one of them can push the shared senior to DT/TR."""
    members = area.members
    if not members:
        return []
    senior = members[0]
    juniors = members[1:]
    s_badge = _norm(senior.badge)
    fam = _family(s_badge)

    if fam in ("ZL", "STL", "AP"):
        assignments = [f"{fam}{i + 1}" for i in range(len(members))]
    elif s_badge in ("DL", "DT", "TR", "SA"):
        assignments = [s_badge] + ["JC"] * len(juniors)
    else:
        assignments = ["SC"] + ["JC"] * len(juniors)

    # SA (special assignment) carries over from the previous news for
    # whichever badge-less members were SA last time.
    for i, member in enumerate(members):
        if not _norm(member.badge) and member.prev is not None and _norm(member.prev.assignment) == "SA":
            assignments[i] = "SA"

    for i, junior in enumerate(juniors, start=1):
        junior_trainee_marker = junior.has_t and junior.state != "matched"
        if junior.state == "masked" or junior_trainee_marker:
            # Actively training a brand-new missionary.
            assignments[0] = "DT" if assignments[0] == "DL" else "TR"
            assignments[i] = "JC"
        elif junior.state == "revealed":
            # Junior arrived last transfer; a DL finishing their training is a
            # DT (district leader training), others keep their badge role.
            if assignments[0] == "DL":
                assignments[0] = "DT"
            assignments[i] = "JC"

    return assignments


# ── Status classification ─────────────────────────────────────────────────────

def _classify(prev_rec: MissionaryRecord, current_zone: str, current_area: str, new_assignment: str) -> tuple[str, str]:
    """Return (status_text, previous_zone_display)."""
    prev_zone = prev_rec.assigned_zone or _canon_zone(prev_rec.section_zone) or _norm(prev_rec.section_zone)

    if _norm_key(prev_zone) and _norm_key(prev_zone) != _norm_key(current_zone):
        movement = "ZONE TRANSFER"
    elif not _area_match(prev_rec.area, current_area):
        movement = "AREA CHANGE"
    else:
        movement = "STAYING"

    prev_assign = _norm(prev_rec.assignment)
    new_assign = _norm(new_assignment)
    role = ""
    if _is_leader(new_assign) and not _is_leader(prev_assign):
        role = " - NEW LEADERSHIP"
    elif _is_leader(prev_assign) and not _is_leader(new_assign):
        role = " - RELEASED FROM LEADERSHIP"
    elif _is_leader(prev_assign) and _is_leader(new_assign) and prev_assign != new_assign:
        role = f" - ROLE CHANGE ({prev_assign} -> {new_assign})"

    return movement + role, prev_zone


# ── Build the News Format table ────────────────────────────────────────────────

@dataclass
class NewsResult:
    news: pd.DataFrame       # one row per missionary (News Format)
    changes: pd.DataFrame    # detected differences vs the previous transfer
    transfer_title: str      # e.g. "JUNE / JULY 2026 TRANSFER NEWS"
    zone_order: list[str]    # canonical zone order (previous order + new zones)


def build_news_format(
    areas: list[AreaRecord],
    prev: dict[str, list[MissionaryRecord]],
    companion_mentions: dict[str, list[MissionaryRecord]] | None = None,
    incoming: set[str] | None = None,
    section_order: list[str] | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, list[str]]:
    """
    Cross-reference the current Transfer Management areas against the previous
    Transfer News. Returns (news_df, changes_df, zone_order):

      news_df: one row per missionary (new arrivals stay masked as their
        trainer's "NEW MISSIONARY" companion), grouped by the zone each
        missionary served in last transfer, in the previous news' zone order,
        preserving each missionary's previous position within the zone
        (leaders who were AP/ZL/STL keep their historical placement).
      changes_df: departures, new zones, closed zones/areas etc.
      zone_order: the previous zone order with any new zones appended.
    """
    companion_mentions = companion_mentions or {}
    incoming = incoming or set()

    consumed = _match_all(areas, prev, companion_mentions, incoming)

    # Zone order is intentional and carried over from the previous news;
    # zones not present there (brand-new zones) are appended at the end in
    # the order they first appear.
    zone_first_seen: dict[str, int] = {}
    for zone in section_order or []:
        zone_first_seen.setdefault(zone, len(zone_first_seen))
    all_prev = sorted((r for recs in prev.values() for r in recs), key=lambda r: r.row_order)
    for rec in all_prev:
        zone = rec.assigned_zone or _canon_zone(rec.section_zone)
        if zone:
            zone_first_seen.setdefault(zone, len(zone_first_seen))

    def group_order(zone: str) -> int:
        if zone not in zone_first_seen:
            zone_first_seen[zone] = len(zone_first_seen)
        return zone_first_seen[zone]

    rows: list[dict] = []

    # The previous news disambiguates duplicate last names with suffixes
    # (MENSAH.J). Keep the suffix only while the base name is still shared by
    # several missionaries; otherwise print the plain name like the mission does.
    base_counts: dict[str, int] = defaultdict(int)
    for area in areas:
        for member in area.members:
            base_counts[_base_name(_norm(member.name))] += 1

    def display_for(person: TMMissionary, fallback: TMMissionary | None) -> str:
        if person.state == "masked":
            return "NEW MISSIONARY"
        if person.prev is not None and person.state == "matched":
            last = person.prev.last_name
            if "." in last and base_counts.get(_base_name(_norm(last)), 0) <= 1:
                last = _base_name(last)
            return f"{person.prev.title} {last}"
        title = "ELDER"
        if fallback is not None and fallback.prev is not None:
            title = fallback.prev.title
        return f"{title} {_norm(person.name)}"

    for area in areas:
        assignments = _group_assignments(area)
        current_zone = _norm(area.zone)
        current_area = _norm(area.area)

        for i, person in enumerate(area.members):
            if person.state == "masked":
                continue
            assignment = assignments[i]
            others = area.members[:i] + area.members[i + 1 :]
            # A fallback companion for title/zone inference when this person
            # has no previous record of their own — any other group member
            # who does have one.
            fallback = next((o for o in others if o.prev is not None), others[0] if others else None)

            if person.state == "matched" and person.prev is not None:
                status, prev_zone = _classify(person.prev, current_zone, current_area, assignment)
                prev_area = person.prev.area
                prev_assignment = person.prev.assignment
                group_zone = prev_zone or current_zone
                row_order = person.prev.row_order
            else:
                # Revealed: arrived last transfer masked as NEW MISSIONARY.
                if fallback is not None and fallback.prev is not None:
                    group_zone = (
                        fallback.prev.assigned_zone
                        or _canon_zone(fallback.prev.section_zone)
                        or current_zone
                    )
                    status = "NEWLY REVEALED - ARRIVED LAST TRANSFER"
                else:
                    group_zone = current_zone
                    status = "NEWLY REVEALED - VERIFY NAME/TITLE"
                prev_zone = group_zone
                prev_area = ""
                prev_assignment = ""
                row_order = 10 ** 9  # newly revealed missionaries go last

            companions = " & ".join(display_for(o, person) for o in others)

            rows.append({
                "Name of Missionary": display_for(person, fallback),
                "Assignment": assignment,
                "New/Existing Zone": current_zone,
                "New/Existing Area": current_area,
                "New/Existing Companion(s)": companions,
                "Status": status,
                "Previous Zone": prev_zone,
                "Previous Area": prev_area,
                "Previous Assignment": prev_assignment,
                "_zone_order": group_order(group_zone or current_zone),
                "_prev_tier": _leadership_tier(prev_assignment),
                "_row_order": row_order,
                "_district": area.district,
            })

    df = pd.DataFrame(rows)
    if df.empty:
        logger.warning("News Format produced 0 rows — check PDF paths and parsing")
        return df, _build_changes(areas, prev, consumed), list(zone_first_seen)

    # Within each zone: AP, ZL and STL companionships first — judged by the
    # role held LAST transfer, then by each missionary's position in the
    # previous news. This keeps former leaders in their historical spot
    # (a JC who was ZL1 still prints above the ZL2 row) instead of re-sorting
    # everyone by their current calling.
    df = df.sort_values(
        ["_zone_order", "_prev_tier", "_row_order", "Name of Missionary"]
    ).reset_index(drop=True)

    changes = _build_changes(areas, prev, consumed)
    return df, changes, list(zone_first_seen)


def _leadership_tier(assignment: str) -> int:
    a = _norm(assignment)
    if a.startswith("AP") or a == "SA":
        return 0
    if a.startswith("ZL"):
        return 1
    if a.startswith("STL"):
        return 2
    return 3


def _build_changes(
    areas: list[AreaRecord],
    prev: dict[str, list[MissionaryRecord]],
    consumed: set[int],
) -> pd.DataFrame:
    """Summarize everything that changed since the previous transfer:
    departures, new/closed zones and areas, arrivals, revealed trainees."""
    rows: list[dict] = []

    prev_records = [rec for recs in prev.values() for rec in recs]
    prev_zones = {rec.assigned_zone or _canon_zone(rec.section_zone) for rec in prev_records}
    prev_zones.discard("")
    prev_areas = {_norm_key(rec.area): rec.area for rec in prev_records if _norm_key(rec.area)}

    current_zones: list[str] = []
    current_areas: dict[str, str] = {}
    for area in areas:
        if area.zone not in current_zones:
            current_zones.append(area.zone)
        current_areas.setdefault(_norm_key(area.area), area.area)

    # Departed: previous rows no one in the current Transfer Management claimed.
    for rec in sorted(prev_records, key=lambda r: r.row_order):
        if id(rec) not in consumed and not _is_new_missionary_text(rec.display_name):
            rows.append({
                "Change": "DEPARTED / NOT FOUND",
                "Missionary": rec.display_name,
                "Detail": f"was {rec.assignment or 'SC/JC'} in {rec.assigned_zone or rec.section_zone} / {rec.area}",
            })

    for area in areas:
        if not area.members:
            # A card with a printed area name but nobody assigned to it —
            # the area is closing when this transfer takes effect.
            rows.append({
                "Change": "AREA CLOSING (NO ASSIGNMENT)",
                "Missionary": "",
                "Detail": f"{area.zone} / {area.area}",
            })
            continue
        senior = area.members[0]
        for junior in area.members[1:]:
            if junior.state == "masked":
                rows.append({
                    "Change": "NEW MISSIONARY (MASKED)",
                    "Missionary": f"junior of {senior.name.upper()}",
                    "Detail": f"{area.zone} / {area.area}",
                })
        for person in area.members:
            if person.state == "revealed":
                rows.append({
                    "Change": "NEWLY REVEALED",
                    "Missionary": person.name.upper(),
                    "Detail": f"now serving in {area.zone} / {area.area}",
                })

    for zone in current_zones:
        if not any(_norm_key(zone) == _norm_key(z) for z in prev_zones):
            rows.append({"Change": "NEW ZONE", "Missionary": "", "Detail": zone})
    for zone in sorted(prev_zones):
        if not any(_norm_key(zone) == _norm_key(z) for z in current_zones):
            rows.append({"Change": "ZONE NOT IN CURRENT TM", "Missionary": "", "Detail": zone})

    for key, name in current_areas.items():
        if key not in prev_areas:
            rows.append({"Change": "NEW/RENAMED AREA", "Missionary": "", "Detail": name.upper()})
    for key, name in sorted(prev_areas.items()):
        if key not in current_areas:
            rows.append({"Change": "CLOSED/RENAMED AREA", "Missionary": "", "Detail": name.upper()})

    return pd.DataFrame(rows, columns=["Change", "Missionary", "Detail"])


# ── Public entry point ─────────────────────────────────────────────────────────

def generate_news_format(
    tm_pdf: Path,
    prev_news_pdf: Path,
) -> NewsResult:
    """
    Full pipeline: parse current Transfer Management + previous Transfer News
    → NewsResult (news table, change report, transfer title, zone order).
    """
    prev, companion_mentions, section_order = parse_previous_news(prev_news_pdf)
    incoming = parse_incoming(tm_pdf)
    areas = parse_tm_zones(tm_pdf)
    news, changes, zone_order = build_news_format(
        areas, prev, companion_mentions, incoming, section_order
    )

    # Zone order for the NEXT document: previous order first, then any zone
    # that only exists in the current Transfer Management.
    known = {_norm_key(z) for z in zone_order}
    for area in areas:
        if _norm_key(area.zone) not in known:
            known.add(_norm_key(area.zone))
            zone_order.append(area.zone)

    return NewsResult(
        news=news,
        changes=changes,
        transfer_title=parse_transfer_title(tm_pdf),
        zone_order=zone_order,
    )


# ── CLI (debug / standalone) ──────────────────────────────────────────────────

if __name__ == "__main__":
    import sys
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")

    tm = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("data/current_transfer_management.pdf")
    prev_path = Path(sys.argv[2]) if len(sys.argv) > 2 else Path("data/old_transfer_news.pdf")

    result = generate_news_format(tm, prev_path)
    df = result.news
    print(f"\nTitle: {result.transfer_title}")
    print(f"Zone order: {result.zone_order}")
    print(f"News Format: {len(df)} rows | Changes: {len(result.changes)} rows")
    pd.set_option("display.max_columns", None)
    pd.set_option("display.width", 250)
    pd.set_option("display.max_rows", None)
    print(df[[c for c in df.columns if not c.startswith("_")]].to_string(index=False))
    print("\nCHANGES:")
    print(result.changes.to_string(index=False))
