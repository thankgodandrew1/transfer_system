"""
verify_news.py

Cross-checks a generated News Format DataFrame against the mission's own
Transfer Management reports:

  data/current_transfer_report.xlsx  — this cycle's roster (authoritative for
    current position/area/zone/companions; also carries "... Prior To Change"
    columns for whichever people changed).
  data/old_transfer_report.xlsx      — last cycle's roster (authoritative for
    previous position/area/zone/companions, including the TR/DT training
    distinction that current_transfer_report.xlsx doesn't capture).

Both reports are more reliable than the PDF-derived data News Format is built
from (structured export vs. OCR-style text/photo-grid parsing), so wherever
they say something unambiguous, this module corrects the News Format
DataFrame in place and records what it changed. Anything it can't resolve
unambiguously (duplicate names even the reports can't separate, a person
missing from the current report entirely) is left as a blocking error instead
of guessed at.

Run standalone: python verify_news.py <workbook.xlsx> [current_report.xlsx] [old_report.xlsx]
"""
from __future__ import annotations

import html
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl
import pandas as pd

from generate_news_format import MissionaryRecord, _classify, _is_new_missionary_text


def _norm_key(text: str) -> str:
    from text_utils import comparison_key
    return comparison_key(text)


def _area_match(a: str, b: str) -> bool:
    na, nb = _norm_key(a), _norm_key(b)
    if not na or not nb:
        return False
    return na == nb or (len(na) >= 5 and na in nb) or (len(nb) >= 5 and nb in na)


def _family(code: str) -> str:
    code = _norm_key(code)
    for fam in ("ZL", "STL", "AP"):
        if code.startswith(fam):
            return fam
    return code


# ── Report loading ──────────────────────────────────────────────────────────


@dataclass
class ReportRecord:
    zone: str
    district: str
    area: str
    full_name: str          # "Last, First ..." as printed in the report
    last: str
    title: str               # "ELDER" or "SISTER", from the report's Type column
    position: str            # normalized current-position code, e.g. "DL", "ZL1", "SA"
    companions: list[str]    # other companions' full names, this report's own cycle
    prior_area: str | None = None
    prior_position: str | None = None  # normalized, e.g. "DT", "SC"


_COMPOUND_TO_CODE = {
    # old_transfer_report.xlsx's "(FAMILY, ROLE)" scheme -> News Format vocabulary.
    # ROLE is SC/JC/TR: SC = senior/primary of the pair, JC = junior/secondary,
    # TR = actively training a brand-new missionary.
}


def _normalize_position(raw: object) -> str:
    """'(ZL1)' -> 'ZL1'; '(DL, TR)' -> 'DT'; '(ZL, JC)' -> 'ZL'; '(SC)' -> 'SC'."""
    text = str(raw or "").strip()
    if text.startswith("(") and text.endswith(")"):
        text = text[1:-1].strip()
    if not text:
        return ""
    parts = [p.strip().upper() for p in text.split(",")]
    if len(parts) == 1:
        return parts[0]
    family, role = parts[0], parts[1]
    if family == "DL" and role == "TR" and len(parts) == 2:
        return "DT"
    if role == "TR":
        return "TR" if family == "SC" and len(parts) == 2 else "/".join(parts)
    if family in ("ZL", "STL", "AP"):
        if role in ("SC", "JC") and len(parts) == 2:
            return family  # legacy numbered-pair notation
    if family in ("DL", "DT", "SC", "JC") and role in ("SC", "JC") and len(parts) == 2:
        return family
    return "/".join(parts)


def _split_report_companions(raw: object) -> list[str]:
    text = str(raw or "").strip()
    if not text:
        return []
    # current_transfer_report.xlsx joins multiple companions with "/", each as
    # "Last, First (POS)".
    names = []
    for piece in text.split("/"):
        piece = re.sub(r"\s*\([^)]*\)\s*$", "", piece.strip()).strip()
        if piece:
            names.append(piece)
    return names


def _clean_place(raw: object) -> str:
    """Report place names carry HTML entities ('Ikot Ebo 1&amp;2') and stray
    edge slashes ('Afaha Eket/', '/Nkana'); clean them the same way the
    Transfer Management parser does so a corrected value prints cleanly."""
    text = re.sub(r"\s+", " ", html.unescape(str(raw or "")))
    return re.sub(r"-\s+", "-", text).strip().strip("/").strip()


def load_report(path: Path) -> list[ReportRecord]:
    wb = openpyxl.load_workbook(str(path), data_only=True)
    try:
        return _read_report(wb, path)
    finally:
        wb.close()


def _read_report(wb, path: Path) -> list[ReportRecord]:
    required = {"missionary name", "zone", "area", "position", "companion"}
    selected = None
    for sheet in wb.worksheets:
        for number, row in enumerate(sheet.iter_rows(max_row=20, values_only=True), start=1):
            header = {re.sub(r"\s+", " ", str(value or "")).strip().casefold(): i for i, value in enumerate(row)}
            if required <= header.keys():
                selected = sheet, number, header
                break
        if selected:
            break
    if selected is None:
        raise ValueError(f'{path.name}: no roster table was found. Export a Transfer Management Excel report with Missionary Name, Zone, Area, Position and Companion columns.')
    ws, header_row, lower_idx = selected
    idx = {name: lower_idx[name.casefold()] for name in (
        "Missionary Name", "Zone", "Area", "Position", "Companion",
        "District", "Type", "Area Prior To Change", "Position Prior To Change",
    ) if name.casefold() in lower_idx}

    records: list[ReportRecord] = []
    seen = set()
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        name = str(row[idx["Missionary Name"]] or "").strip()
        if not name:
            continue
        key = _norm_key(name)
        if key in seen:
            raise ValueError(f'{path.name}: "{name}" appears more than once. Export a roster containing one row per missionary, or distinguish identical names before uploading.')
        seen.add(key)
        if not _clean_place(row[idx["Zone"]]) or not _clean_place(row[idx["Area"]]):
            raise ValueError(f'{path.name}: "{name}" is missing a zone or area. Complete that assignment in the roster and upload it again.')
        last = name.split(",")[0].strip()
        prior_area = row[idx["Area Prior To Change"]] if "Area Prior To Change" in idx else None
        prior_pos_raw = row[idx["Position Prior To Change"]] if "Position Prior To Change" in idx else None
        type_raw = str(row[idx["Type"]] or "").strip().upper() if "Type" in idx else ""
        records.append(ReportRecord(
            zone=_clean_place(row[idx["Zone"]]),
            district=_clean_place(row[idx["District"]]) if "District" in idx else "",
            area=_clean_place(row[idx["Area"]]),
            full_name=name,
            last=last,
            title="SISTER" if type_raw.startswith("SISTER") else ("ELDER" if type_raw.startswith("ELDER") else ""),
            position=_normalize_position(row[idx["Position"]]),
            companions=_split_report_companions(row[idx["Companion"]]),
            prior_area=_clean_place(prior_area) if prior_area else None,
            prior_position=_normalize_position(prior_pos_raw) if prior_pos_raw else None,
        ))
    if not records:
        raise ValueError(f'{path.name}: the roster has no missionaries. Upload the populated Transfer Management report.')
    return records


# ── Verification result ─────────────────────────────────────────────────────


@dataclass
class Correction:
    row_name: str
    field: str
    old_value: str
    new_value: str
    reason: str


@dataclass
class Issue:
    row_name: str
    detail: str
    severity: str  # "blocking" or "note"


@dataclass
class VerificationResult:
    news: pd.DataFrame
    corrections: list[Correction] = field(default_factory=list)
    blocking_errors: list[Issue] = field(default_factory=list)
    notes: list[Issue] = field(default_factory=list)
    new_missionary_count: int | None = None


def _last_name(display_name: str) -> str:
    parts = str(display_name).split()
    if parts and parts[0].upper() in ("ELDER", "ELDERS", "SISTER", "SISTERS"):
        parts = parts[1:]
    # Strip a trailing first-initial disambiguator ("OKON P." -> "OKON").
    if parts and parts[-1].upper() not in {"JR.", "SR.", "II.", "III.", "IV."} and re.fullmatch(r"[^\W_]+\.", parts[-1]):
        parts = parts[:-1]
    # Keep every remaining word: a last name can be "COLLINS JR".
    return " ".join(parts).split(".")[0]


def _name_disambiguators(records: list[ReportRecord]) -> dict[str, str]:
    result = {}
    for rec in records:
        first = _norm_key(rec.full_name.split(",", 1)[1]) if "," in rec.full_name else ""
        peers = [_norm_key(other.full_name.split(",", 1)[1]) for other in records
            if other is not rec and _norm_key(other.last) == _norm_key(rec.last) and "," in other.full_name]
        if not first or not peers:
            continue
        length = next((n for n in range(1, len(first) + 1) if all(first[:n] != peer[:n] for peer in peers)), len(first))
        result[_norm_key(rec.full_name)] = first[:length]
    return result


def _display_name(rec: "ReportRecord", dup_last_names: set[str], confirmed_new_fullnames: set[str], disambiguators=None) -> str:
    """The name a person should print as: masked if they're a confirmed new
    missionary, disambiguated with a first initial if their last name
    collides with another named current missionary, otherwise plain."""
    if _norm_key(rec.full_name) in confirmed_new_fullnames:
        return "NEW MISSIONARY"
    title = rec.title or "ELDER"
    last = rec.last.upper()
    if _norm_key(rec.last) in dup_last_names:
        first = rec.full_name.split(",", 1)[1].strip() if "," in rec.full_name else ""
        if first:
            prefix = (disambiguators or {}).get(_norm_key(rec.full_name), first[0].upper())
            return f"{title} {last} {prefix}."
    return f"{title} {last}"


def verify(
    news: pd.DataFrame,
    current_report_path: Path | None,
    old_report_path: Path | None = None,
    manual_corrections_path: Path | None = None,
) -> VerificationResult:
    result = VerificationResult(news=news.copy())
    if current_report_path is None or not Path(current_report_path).exists():
        result.notes.append(Issue("", "No current_transfer_report.xlsx supplied — skipping verification.", "note"))
        return result

    current_records = load_report(Path(current_report_path))
    old_records = load_report(Path(old_report_path)) if old_report_path and Path(old_report_path).exists() else []

    current_by_key: dict[tuple[str, str], list[ReportRecord]] = {}
    current_by_last: dict[str, list[ReportRecord]] = {}
    current_by_fullname: dict[str, ReportRecord] = {}
    for rec in current_records:
        current_by_key.setdefault((_norm_key(rec.last), _norm_key(rec.area)), []).append(rec)
        current_by_last.setdefault(_norm_key(rec.last), []).append(rec)
        current_by_fullname[_norm_key(rec.full_name)] = rec

    old_by_fullname: dict[str, ReportRecord] = {_norm_key(r.full_name): r for r in old_records}

    # Ground-truth new-missionary count: everyone in this cycle's full roster
    # who isn't in last cycle's at all — independent of News Format's own
    # row set, since brand-new missionaries are deliberately masked out of it.
    if old_records:
        result.new_missionary_count = sum(
            1 for rec in current_records if _norm_key(rec.full_name) not in old_by_fullname
        )

    # Confirmed new = absent from old_transfer_report.xlsx entirely; these
    # never get a name in print, so they don't count as a "collision" for
    # anyone else who happens to share their last name.
    confirmed_new_fullnames: set[str] = (
        {_norm_key(r.full_name) for r in current_records if _norm_key(r.full_name) not in old_by_fullname}
        if old_records else set()
    )
    named_current = [r for r in current_records if _norm_key(r.full_name) not in confirmed_new_fullnames]
    last_counts: dict[str, int] = {}
    for r in named_current:
        last_counts[_norm_key(r.last)] = last_counts.get(_norm_key(r.last), 0) + 1
    dup_last_names = {k for k, v in last_counts.items() if v > 1}
    disambiguators = _name_disambiguators(named_current)
    def display(rec):
        return _display_name(rec, dup_last_names, confirmed_new_fullnames, disambiguators)
    printed_records = {}
    for rec in named_current:
        printed_records.setdefault(_norm_key(display(rec)), []).append(rec)

    df = result.news
    if df.empty:
        return result

    def find_current(row) -> ReportRecord | None:
        printed = printed_records.get(_norm_key(row["Name of Missionary"]), [])
        last_key = _norm_key(_last_name(row["Name of Missionary"]))
        if len(printed) == 1:
            explicit_initial = _norm_key(row["Name of Missionary"]) != _norm_key(f'{printed[0].title or "ELDER"} {printed[0].last}')
            if explicit_initial or len(current_by_last.get(last_key, [])) == 1:
                return printed[0]
        area_key = _norm_key(row["New/Existing Area"])
        cands = current_by_key.get((last_key, area_key))
        if not cands:
            cands = current_by_key.get((_norm_key(_last_name(row["Name of Missionary"]).rstrip("0123456789")), area_key))
        if not cands:
            by_last = current_by_last.get(last_key)
            if by_last and len(by_last) == 1:
                cands = by_last
        if not cands:
            return None
        if len(cands) > 1:
            # Disambiguate by companion overlap between the News Format row
            # and the report's own companion list.
            news_companions = {
                _norm_key(_last_name(p)) for p in re.split(r"\s*&\s*", str(row["New/Existing Companion(s)"]))
            }
            scored = [
                (rec, len(news_companions & {_norm_key(c.split(",")[0]) for c in rec.companions}))
                for rec in cands
            ]
            scored.sort(key=lambda t: -t[1])
            if scored[0][1] > 0 and (len(scored) == 1 or scored[0][1] > scored[1][1]):
                cands = [scored[0][0]]
        return cands[0] if len(cands) == 1 else None

    def position_matches(news_code: str, report_code: str) -> bool:
        n, r = _norm_key(news_code), _norm_key(report_code)
        if n == r:
            return True
        # Report only numbers the "#1" of a leadership pair; the "#2+" slot
        # prints as the bare family name.
        if r in ("AP", "ZL", "STL") and re.fullmatch(re.escape(r) + r"\d+", str(news_code).strip().upper()):
            return True
        # current_transfer_report.xlsx has no TR/DT vocabulary at all — it
        # can only ever say "SC" or "DL" for someone actively training a
        # brand-new missionary, so that's not a real disagreement. Only
        # old_transfer_report-chained evidence (see below) should move a
        # position into or out of TR/DT.
        if n == "TR" and r in ("SC", ""):
            return True
        if n == "DT" and r == "DL":
            return True
        # Neither report has an SA code either: a special-assignment
        # missionary exports as a plain SC/JC.
        if n == "SA" and r in ("SC", "JC", ""):
            return True
        return False

    def strip_zone_suffix(zone: str) -> str:
        return re.sub(r"\s*ZONE\s*$", "", str(zone or ""), flags=re.IGNORECASE).strip()

    corrections = result.corrections
    idx_to_rec: dict[int, ReportRecord] = {}
    confirmed_new_idx: list[int] = []

    for idx, row in df.iterrows():
        rec = find_current(row)
        if rec is None:
            result.blocking_errors.append(Issue(
                row["Name of Missionary"],
                "Not found in current_transfer_report.xlsx — generation cannot safely verify this row.",
                "blocking",
            ))
            continue
        idx_to_rec[idx] = rec

        row_name = row["Name of Missionary"]

        # Current position.
        if rec.position and not position_matches(row["Assignment"], rec.position):
            corrections.append(Correction(row_name, "Assignment", str(row["Assignment"]), rec.position,
                                           "current_transfer_report.xlsx Position"))
            df.at[idx, "Assignment"] = rec.position

        # Current zone (report zone names always carry a trailing "Zone" that
        # News Format's canonical zone names never do).
        report_zone = strip_zone_suffix(rec.zone)
        if report_zone and _norm_key(report_zone) != _norm_key(row["New/Existing Zone"]) and not _area_match(report_zone, row["New/Existing Zone"]):
            corrections.append(Correction(row_name, "New/Existing Zone", str(row["New/Existing Zone"]), report_zone.upper(),
                                           "current_transfer_report.xlsx Zone"))
            df.at[idx, "New/Existing Zone"] = report_zone.upper()

        # Current area.
        if rec.area and not _area_match(rec.area, row["New/Existing Area"]):
            corrections.append(Correction(row_name, "New/Existing Area", str(row["New/Existing Area"]), rec.area.upper(),
                                           "current_transfer_report.xlsx Area"))
            df.at[idx, "New/Existing Area"] = rec.area.upper()

        # Current companion(s): rebuild from the report's own companion list
        # (which already handles trios/foursomes), masking anyone confirmed
        # new and disambiguating anyone whose last name collides with
        # another named current missionary — the same logic that applies to
        # this row's own name below, so the two sides of a companionship
        # always agree on how each other prints.
        report_comp_lasts = {_norm_key(c.split(",")[0]) for c in rec.companions}
        if rec.companions:
            new_companions = " & ".join(dict.fromkeys(
                display(current_by_fullname[_norm_key(c)])
                if _norm_key(c) in current_by_fullname
                else c.split(",")[0].strip().upper()
                for c in rec.companions
            ))
            if new_companions != str(row["New/Existing Companion(s)"]):
                corrections.append(Correction(row_name, "New/Existing Companion(s)", str(row["New/Existing Companion(s)"]),
                                               new_companions, "current_transfer_report.xlsx Companion"))
                df.at[idx, "New/Existing Companion(s)"] = new_companions

        # This row's own name: disambiguate if it collides with another
        # named current missionary sharing the last name. (Confirmed-new
        # rows are handled by the masking pass below, not here.)
        if _norm_key(rec.full_name) not in confirmed_new_fullnames and _norm_key(rec.last) in dup_last_names:
            expected_name = display(rec)
            if expected_name != row_name:
                corrections.append(Correction(row_name, "Name of Missionary", row_name, expected_name,
                                               f"disambiguated — another current missionary also has last name {rec.last!r}"))
                df.at[idx, "Name of Missionary"] = expected_name
                row_name = expected_name

        # Title: the previous news can carry a wrong ELDER/SISTER forward, and
        # the companion cells above already print the report's title.
        first_word, _, rest = str(row_name).partition(" ")
        if rest and rec.title and first_word in ("ELDER", "SISTER") and first_word != rec.title:
            corrected_name = f"{rec.title} {rest}"
            corrections.append(Correction(row_name, "Name of Missionary", row_name, corrected_name,
                                           "current_transfer_report.xlsx Type"))
            df.at[idx, "Name of Missionary"] = corrected_name
            row_name = corrected_name

        # Previous state, via old_transfer_report.xlsx chained by exact full name.
        old_rec = old_by_fullname.get(_norm_key(rec.full_name))
        if old_rec is not None:
            if old_rec.position and not position_matches(row["Previous Assignment"], old_rec.position):
                corrections.append(Correction(row_name, "Previous Assignment", str(row["Previous Assignment"]), old_rec.position,
                                               "old_transfer_report.xlsx Position"))
                df.at[idx, "Previous Assignment"] = old_rec.position
            if old_rec.area and not _area_match(old_rec.area, row["Previous Area"]):
                corrections.append(Correction(row_name, "Previous Area", str(row["Previous Area"]), old_rec.area.upper(),
                                               "old_transfer_report.xlsx Area"))
                df.at[idx, "Previous Area"] = old_rec.area.upper()
            old_zone = strip_zone_suffix(old_rec.zone)
            if old_zone and _norm_key(old_zone) != _norm_key(row["Previous Zone"]) and not _area_match(old_zone, row["Previous Zone"]):
                corrections.append(Correction(row_name, "Previous Zone", str(row["Previous Zone"]), old_zone.upper(),
                                               "old_transfer_report.xlsx Zone"))
                df.at[idx, "Previous Zone"] = old_zone.upper()

            # Status is derived from the previous state — recompute it now
            # that we know that state is right (this is what actually fixes
            # a "NEWLY REVEALED" mislabel: the person wasn't new, a
            # duplicate-name mismatch just made them look unmatched).
            # Use the dataframe's own (already tolerance-corrected) Previous
            # Assignment rather than old_rec.position directly: the report
            # only numbers the "#1" of a leadership pair, so old_rec.position
            # alone can't distinguish ZL1 from ZL2 and would manufacture a
            # bogus "ROLE CHANGE (ZL -> ZL2)" every time.
            expected_status, _ = _classify(
                MissionaryRecord(
                    display_name=rec.full_name, title="", last_name=rec.last,
                    assignment=str(df.at[idx, "Previous Assignment"]) or "", section_zone="",
                    assigned_zone=strip_zone_suffix(old_rec.zone).upper(),
                    area=old_rec.area or "", companion="",
                ),
                str(df.at[idx, "New/Existing Zone"]), str(df.at[idx, "New/Existing Area"]), str(df.at[idx, "Assignment"]),
            )
            if expected_status != row["Status"]:
                corrections.append(Correction(row_name, "Status", str(row["Status"]), expected_status,
                                               "recomputed from old_transfer_report.xlsx previous state"))
                df.at[idx, "Status"] = expected_status
        elif old_records:
            # Confirmed new (absent from old_transfer_report.xlsx entirely).
            # The mission's convention is that a genuinely new missionary
            # gets no row of their own and no name in print — only their
            # trainer's companion cell shows "NEW MISSIONARY" — regardless of
            # whatever this row currently says (a duplicate-name mismatch
            # from the PDF-fuzzy matcher is exactly why they ended up with a
            # named row in the first place). Masked in a pass below, once
            # every row has been matched.
            confirmed_new_idx.append(idx)

        # A DL/SC training a confirmed-new missionary (absent from
        # old_transfer_report.xlsx entirely) should read DT/TR, even though
        # neither report captures that badge directly pre-transfer.
        if old_records and rec.position in ("DL", "SC", "") and _norm_key(df.at[idx, "Assignment"]) != "SA":
            companion_is_new = any(
                _norm_key(c.split(",")[0]) in report_comp_lasts
                and old_by_fullname.get(_norm_key(c)) is None
                for c in rec.companions
            )
            if companion_is_new:
                expected = "DT" if rec.position == "DL" else "TR"
                if _norm_key(row["Assignment"]) != expected:
                    corrections.append(Correction(row_name, "Assignment", str(row["Assignment"]), expected,
                                                   "training a companion absent from old_transfer_report.xlsx"))
                    df.at[idx, "Assignment"] = expected

    # Coverage: the loop above only checks rows the PDF parser produced, so a
    # district the parser missed would never be noticed. Check the other
    # direction too — every current missionary who isn't masked as new must
    # own exactly one row. Without the old report, masked new missionaries
    # can't be told apart from missing ones, so the check needs it.
    wrongly_masked: list[ReportRecord] = []

    def masked_slot_of(rec: ReportRecord) -> int | None:
        """The row whose companion cell shows rec as NEW MISSIONARY. Everyone
        checked here is named in old_transfer_report.xlsx, so that mask is the
        parser's mistake: a former AP or other leader going to the field whose
        previous-news name it couldn't match, or a trainee who changed areas."""
        key = _norm_key(rec.full_name)
        companions = {_norm_key(c) for c in rec.companions}
        for idx, mate in idx_to_rec.items():
            linked = _norm_key(mate.full_name) in companions or key in {_norm_key(c) for c in mate.companions}
            if linked and _is_new_missionary_text(news.at[idx, "New/Existing Companion(s)"]):
                return idx
        return None

    def undo_trainer_promotion(rec: ReportRecord, mate_idx: int) -> None:
        """The parser made the senior a trainer because of that mask. A former
        leader isn't a trainee, so unless the senior also has a confirmed-new
        companion they go back to the position the report gives them."""
        mate = idx_to_rec[mate_idx]
        trainee = old_by_fullname[_norm_key(rec.full_name)].position in ("JC", "")
        training = any(_norm_key(c) in confirmed_new_fullnames for c in mate.companions)
        current = str(df.at[mate_idx, "Assignment"])
        if not trainee and not training and current in ("TR", "DT") and mate.position in ("SC", "DL", ""):
            restored = mate.position or "SC"
            corrections.append(Correction(str(df.at[mate_idx, "Name of Missionary"]), "Assignment", current, restored,
                                           f"not training {rec.full_name}, who is not a new missionary"))
            df.at[mate_idx, "Assignment"] = restored

    if old_records:
        rows_per_rec: dict[int, int] = {}
        for rec in idx_to_rec.values():
            rows_per_rec[id(rec)] = rows_per_rec.get(id(rec), 0) + 1
        for rec in named_current:
            count = rows_per_rec.get(id(rec), 0)
            where = f"{strip_zone_suffix(rec.zone)} / {rec.area}"
            masked_slot = masked_slot_of(rec) if count == 0 else None
            if count == 0:
                wrongly_masked.append(rec)
                if masked_slot is not None:
                    undo_trainer_promotion(rec, masked_slot)
            elif count > 1:
                result.blocking_errors.append(Issue(
                    display(rec),
                    f"{rec.full_name} ({where}) matches {count} Transfer News rows — only one can be right.",
                    "blocking",
                ))
    else:
        result.notes.append(Issue(
            "", "No old_transfer_report.xlsx supplied — could not check that every current missionary has a row.",
            "note",
        ))

    # Drop every confirmed-new missionary's own row. Their companion cells
    # elsewhere already read "NEW MISSIONARY" (built via _display_name
    # above) — this is what actually enforces the mission's masking
    # convention, rather than just cleaning up the data on a row that
    # should never have existed in the first place.
    if confirmed_new_idx:
        for i in confirmed_new_idx:
            corrections.append(Correction(
                str(df.at[i, "Name of Missionary"]), "(row)", "named row", "removed — masked as NEW MISSIONARY",
                "confirmed new missionary (absent from old_transfer_report.xlsx) — the mission's convention is "
                "no name or row of their own until a later transfer",
            ))
        df = df.drop(index=confirmed_new_idx).reset_index(drop=True)
        result.news = df

    # Give each wrongly masked missionary the row they should have had, built
    # from the reports. This runs before the manual overrides so that an office
    # entry such as "Elder Example | JC" has a row to act on.
    for rec in wrongly_masked:
        old_rec = old_by_fullname[_norm_key(rec.full_name)]
        name = display(rec)
        zone = strip_zone_suffix(rec.zone).upper()
        area = rec.area.upper()
        old_zone = strip_zone_suffix(old_rec.zone).upper()
        assignment = rec.position
        if any(_norm_key(c) in confirmed_new_fullnames for c in rec.companions):
            if assignment in ("SC", "DL", ""):
                assignment = "DT" if assignment == "DL" else "TR"
        status, _ = _classify(
            MissionaryRecord(
                display_name=rec.full_name, title="", last_name=rec.last, assignment=old_rec.position,
                section_zone="", assigned_zone=old_zone, area=old_rec.area or "", companion="",
            ),
            zone, area, rec.position,
        )
        row: dict[str, object] = dict.fromkeys(df.columns, "")
        row.update({
            "Name of Missionary": name,
            "Assignment": assignment,
            "New/Existing Zone": zone,
            "New/Existing Area": area,
            "New/Existing Companion(s)": " & ".join(dict.fromkeys(
                display(current_by_fullname[_norm_key(c)])
                if _norm_key(c) in current_by_fullname
                else c.split(",")[0].strip().upper()
                for c in rec.companions
            )),
            "Status": status,
            "Previous Zone": old_zone,
            "Previous Area": (old_rec.area or "").upper(),
            "Previous Assignment": old_rec.position,
        })
        if "_row_order" in df.columns:
            row["_row_order"] = 10 ** 9
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
        corrections.append(Correction(
            name, "(row)", "masked as NEW MISSIONARY", "row added from the reports",
            "named in both transfer reports; missing PDF row restored from authoritative assignments",
        ))
    result.news = df

    _apply_manual_corrections(df, manual_corrections_path, corrections)

    return result


# ── Manual overrides ────────────────────────────────────────────────────────
#
# For facts neither report captures correctly (e.g. a Special Assignment the
# mission hasn't reflected in Transfer Management yet). Applied last, so it
# can override even the ground-truth reports. CSV columns: Name,Field,Value —
# Name may be a surname or the full printed missionary name; Field is any News
# Format column name; Value replaces it verbatim.

def _load_manual_corrections(path: Path) -> list[tuple[str, str, str]]:
    import csv

    if not path.exists():
        return []
    rows: list[tuple[str, str, str]] = []
    with open(path, newline="", encoding="utf-8-sig") as f:
        for entry in csv.DictReader(f):
            name, field, value = entry.get("Name", "").strip(), entry.get("Field", "").strip(), entry.get("Value", "").strip()
            if name and field:
                rows.append((name, field, value))
    return rows


def _apply_manual_corrections(df: pd.DataFrame, path: Path | None, corrections: list[Correction]) -> None:
    # Never inherit another transfer's local office corrections implicitly.
    if path is None:
        return
    overrides = _load_manual_corrections(path)
    if not overrides:
        return
    pending = []
    for name, field, value in overrides:
        if field not in df.columns:
            raise ValueError(f'Correction for "{name}" uses an unknown field "{field}". Use a News Format column heading.')
        target_key = _norm_key(name)
        exact = [idx for idx in df.index if target_key == _norm_key(str(df.at[idx, "Name of Missionary"]))]
        matches = exact or [idx for idx in df.index if target_key == _norm_key(_last_name(df.at[idx, "Name of Missionary"]))]
        if len(matches) != 1:
            reason = "matches more than one missionary" if matches else "does not match a named missionary"
            raise ValueError(f'Correction for "{name}" {reason}. Copy the complete name from the current Transfer News, including its title and initials. New missionaries remain masked and have no named row.')
        if field == "Assignment":
            from roles import normalize_role
            value = normalize_role(value)
        pending.append((matches[0], field, value))
    # Validate the entire file before changing any rows.
    for idx, field, value in pending:
        old_value = str(df.at[idx, field])
        if old_value == value:
            continue
        corrections.append(Correction(str(df.at[idx, "Name of Missionary"]), field, old_value, value,
                                       f"manual override ({path.name})"))
        df.at[idx, field] = value


# ── CLI / pipeline entry point ──────────────────────────────────────────────


def verify_workbook(
    workbook_path: Path,
    current_report_path: Path | None = None,
    old_report_path: Path | None = None,
) -> VerificationResult:
    df = pd.read_excel(workbook_path, sheet_name="News Format")
    return verify(df, current_report_path, old_report_path)


if __name__ == "__main__":
    wb_path = Path(sys.argv[1]) if len(sys.argv) > 1 else None
    if wb_path is None:
        print("Usage: python verify_news.py <workbook.xlsx> [current_report.xlsx] [old_report.xlsx]")
        sys.exit(1)
    cur = Path(sys.argv[2]) if len(sys.argv) > 2 else Path("data/current_transfer_report.xlsx")
    old = Path(sys.argv[3]) if len(sys.argv) > 3 else Path("data/old_transfer_report.xlsx")
    res = verify_workbook(wb_path, cur, old)
    print(f"{len(res.corrections)} corrections, {len(res.blocking_errors)} blocking errors, {len(res.notes)} notes")
    for c in res.corrections:
        print(f"  FIX  {c.row_name}: {c.field} {c.old_value!r} -> {c.new_value!r}  ({c.reason})")
    for e in res.blocking_errors:
        print(f"  BLOCK {e.row_name}: {e.detail}")
    for n in res.notes[:20]:
        print(f"  note {n.row_name}: {n.detail}")
    if res.new_missionary_count is not None:
        print(f"New missionaries (per ground truth): {res.new_missionary_count}")
