"""Editable per-run roster snapshots. Uploaded originals are never overwritten."""
from collections import Counter
from dataclasses import asdict
import json
from pathlib import Path

import openpyxl
import pandas as pd

from roles import normalize_role
from text_utils import comparison_key
from verify_news import ReportRecord, _display_name, _name_disambiguators, load_report


def save_review(job_dir: Path, state: dict) -> None:
    path = job_dir / "review.json"
    temporary = path.with_suffix(".tmp")
    temporary.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
    temporary.replace(path)


def load_review(job_dir: Path) -> dict:
    return json.loads((job_dir / "review.json").read_text(encoding="utf-8"))


def printable_names(records: list[dict]) -> list[str]:
    counts = Counter(comparison_key(row["last"]) for row in records)
    duplicates = {key for key, count in counts.items() if count > 1}
    parsed = [ReportRecord(**row) for row in records]
    disambiguators = _name_disambiguators(parsed)
    return [_display_name(row, duplicates, set(), disambiguators) for row in parsed]


def initialize_review(job_dir: Path, values: dict, role_overrides=()) -> dict:
    inputs = job_dir / "inputs"
    state = {"revision": 0, "edits": [], "values": dict(values)}
    for source, filename in (("current", "current_transfer_report.xlsx"), ("previous", "old_transfer_report.xlsx")):
        state[source] = [asdict(row) for row in load_report(inputs / filename)]
    # Absorb the familiar quick-role entries before showing the editable roster.
    names = printable_names(state["current"])
    for name, role in role_overrides:
        matches = [i for i, row in enumerate(state["current"]) if comparison_key(name) in {
            comparison_key(names[i]), comparison_key(row["last"]), comparison_key(row["full_name"]),
        }]
        if len(matches) != 1:
            raise ValueError(f'Role override "{name}" must match one missionary. Leave the quick overrides empty and select that person on the review screen instead.')
        row = state["current"][matches[0]]
        state["edits"].append({"source": "current", "name": row["full_name"], "before": dict(row), "reason": "Uploaded quick-role override"})
        row["position"] = role
    state["values"]["current_role_overrides"] = ""
    save_review(job_dir, state)
    return state


def review_summary(state: dict) -> dict:
    current = {comparison_key(row["full_name"]): row for row in state["current"]}
    previous = {comparison_key(row["full_name"]): row for row in state["previous"]}
    arrivals = [dict(row, index=i) for i, row in enumerate(state["current"]) if comparison_key(row["full_name"]) not in previous]
    departures = [dict(row, index=i) for i, row in enumerate(state["previous"]) if comparison_key(row["full_name"]) not in current]
    warnings = []
    if not current.keys() & previous.keys():
        warnings.append("The rosters have no names in common. Check that these are consecutive full mission rosters, or link confirmed name variations below.")
    if len(arrivals) > max(10, len(current) // 4):
        warnings.append("More than a quarter of the current roster appears new. Check the previous roster and name variations before confirming.")
    if current.keys() == previous.keys() and all(
        (row["zone"], row["area"], row["position"], row["companions"]) ==
        (previous[key]["zone"], previous[key]["area"], previous[key]["position"], previous[key]["companions"])
        for key, row in current.items()
    ):
        warnings.append("Both rosters contain identical assignments. Confirm that the current and previous files are correct.")
    return {"current_count": len(current), "previous_count": len(previous), "arrivals": arrivals,
            "matched_count": len(current.keys() & previous.keys()),
            "departures": departures, "warnings": warnings,
            "assignment_changes": sum((row["zone"], row["area"], row["position"]) !=
                (previous[key]["zone"], previous[key]["area"], previous[key]["position"])
                for key, row in current.items() if key in previous)}


def _index(value, records):
    try:
        index = int(value)
    except (ValueError, TypeError):
        raise ValueError("Select a missionary from the roster.") from None
    if index < 0 or index >= len(records):
        raise ValueError("Select a missionary from this roster.")
    return index


def edit_assignment(state: dict, form) -> None:
    source = form.get("source")
    if source not in ("current", "previous"):
        raise ValueError("Choose the current roster or previous baseline.")
    records = state[source]
    index = _index(form.get("missionary"), records)
    zone, area, reason = (str(form.get(key, "")).strip() for key in ("zone", "area", "reason"))
    if not zone or not area or len(zone) > 120 or len(area) > 160:
        raise ValueError("Enter a zone and area (up to 120 and 160 characters).")
    if not reason or len(reason) > 500:
        raise ValueError("Explain the correction or emergency transfer in up to 500 characters.")
    if any(ord(c) < 32 for c in zone + area + reason):
        raise ValueError("Use plain text for the zone, area and reason.")
    role = normalize_role(form.get("role", ""))
    companions = {_index(value, records) for value in form.getlist("companions")}
    companions.discard(index)
    group = companions | {index}
    group_names = {records[i]["full_name"] for i in group}
    before = [dict(row, companions=list(row["companions"])) for row in records]
    # Updating a group removes stale links to its former companions, then
    # creates reciprocal links. Never infer companionship from a shared area.
    for i, row in enumerate(records):
        if i in group:
            row.update(zone=zone, area=area, companions=[records[j]["full_name"] for j in sorted(group) if j != i])
        else:
            row["companions"] = [name for name in row["companions"] if name not in group_names]
    records[index]["position"] = role
    state["edits"].append({"source": source, "name": records[index]["full_name"], "reason": reason,
        "changes": [{"before": old, "after": new} for old, new in zip(before, records) if old != new]})
    state["revision"] += 1


def link_identity(state: dict, form) -> None:
    current = state["current"][_index(form.get("current_identity"), state["current"])]
    previous = state["previous"][_index(form.get("previous_identity"), state["previous"])]
    summary = review_summary(state)
    if current["full_name"] not in {row["full_name"] for row in summary["arrivals"]} or previous["full_name"] not in {row["full_name"] for row in summary["departures"]}:
        raise ValueError("Only link an apparent arrival to an apparent departure. Existing matched people cannot be merged.")
    reason = str(form.get("reason", "")).strip()
    if not reason or len(reason) > 500:
        raise ValueError("Explain why these two names identify the same person.")
    old_name = previous["full_name"]
    state["edits"].append({"source": "previous", "name": old_name, "new_name": current["full_name"], "reason": reason})
    previous.update(full_name=current["full_name"], last=current["last"])
    for row in state["previous"]:
        row["companions"] = [current["full_name"] if name == old_name else name for name in row["companions"]]
    state["revision"] += 1


def undo_last(state: dict) -> None:
    if not state["edits"]:
        raise ValueError("There are no saved corrections to undo.")
    edit = state["edits"].pop()
    records = state[edit["source"]]
    if "new_name" in edit:
        for row in records:
            if row["full_name"] == edit["new_name"]:
                row.update(full_name=edit["name"], last=edit["name"].split(",")[0].strip())
            row["companions"] = [edit["name"] if name == edit["new_name"] else name for name in row["companions"]]
    else:
        changes = edit.get("changes", [{"before": edit.get("before")}])
        for change in changes:
            old = change["before"]
            index = next(i for i, row in enumerate(records) if row["full_name"] == old["full_name"])
            records[index] = old
    state.setdefault("undone", []).append({"revision": state["revision"], "correction": edit})
    state["revision"] += 1


def export_report(records: list[dict], path: Path) -> None:
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Reviewed roster"
    sheet.append(["Zone", "District", "Area", "Missionary Name", "Type", "Position", "Companion"])
    for row in records:
        sheet.append([row[key] for key in ("zone", "district", "area", "full_name", "title", "position")] + [" / ".join(row["companions"])])
    for row in sheet:
        for cell in row:
            if isinstance(cell.value, str):
                cell.data_type = "s"
    sheet.freeze_panes = "A2"
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)
    workbook.close()


def export_review(job_dir: Path, state: dict) -> tuple[Path, Path]:
    paths = []
    for source in ("current", "previous"):
        path = job_dir / "reviewed" / f"{source}_report.xlsx"
        export_report(state[source], path)
        paths.append(path)
    return tuple(paths)


def baseline_workbook(state: dict, path: Path) -> None:
    """A complete pre-transfer snapshot for movement FROM locations."""
    rows = state["previous"]
    names = printable_names(rows)
    # A departure can change surname disambiguation between cycles. Common
    # people must use the SAME printed name as current verified News, or a
    # real move would look like an arrival plus a departure to Movements.
    previous_keys = {comparison_key(row["full_name"]) for row in rows}
    continuing = [row for row in state["current"] if comparison_key(row["full_name"]) in previous_keys]
    continuing_names = dict(zip((comparison_key(row["full_name"]) for row in continuing), printable_names(continuing)))
    for i, row in enumerate(rows):
        key = comparison_key(row["full_name"])
        if key in continuing_names:
            names[i] = continuing_names[key]
        elif names[i] in continuing_names.values():
            names[i] = f'{row["title"] or "ELDER"} {row["full_name"].replace(",", " ")}'.upper()
    if len({comparison_key(name) for name in names}) != len(names):
        raise ValueError("Two previous missionaries have indistinguishable printed names. Add distinguishing initials in the source report before generating movements.")
    name_map = dict(zip((row["full_name"] for row in rows), names))
    frame = pd.DataFrame([{"Name of Missionary": names[i], "Assignment": row["position"],
        "New/Existing Zone": row["zone"], "New/Existing Area": row["area"],
        "New/Existing Companion(s)": " & ".join(name_map.get(name, name) for name in row["companions"]),
        "Previous Zone": row["zone"], "Previous Area": row["area"], "Previous Assignment": row["position"], "Status": "BASELINE"}
        for i, row in enumerate(rows)])
    with pd.ExcelWriter(path) as writer:
        frame.to_excel(writer, sheet_name="News Format", index=False)
        pd.DataFrame([{"Metric": "New Missionaries (Incoming)", "Value": 0}]).to_excel(writer, sheet_name="Meta", index=False)
