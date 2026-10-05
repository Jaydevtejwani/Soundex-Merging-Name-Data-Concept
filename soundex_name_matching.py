"""
Collecting Data As Per Soundex Name Percentage
------------------------------------------------
Desktop Excel name-matching and consolidation application.

Requirements:
    Python 3.10+
    pip install pandas openpyxl rapidfuzz

Run:
    python soundex_name_matching.py

The GUI lets the user choose:
    1. Input Excel file
    2. Output Excel file
    3. Matching Merge Percentage (default 80)

Output workbook:
    Original Data
    Processed Data
    Match Details
    Final Merged Data

Notes:
- All source columns are detected dynamically.
- No fixed list of data fields is required.
- The matching engine combines normalized name similarity,
  token/order similarity, component similarity, and phonetic similarity.
- Matching is deliberately conservative: phonetic similarity alone
  cannot create a merge.
- Conflicting nonblank values are preserved in Final Merged Data
  using a pipe-separated representation, rather than silently
  overwriting one value with another.
"""

from __future__ import annotations

import math
import os
import re
import subprocess
import sys
import threading
import time
import traceback
import unicodedata
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from rapidfuzz import fuzz, process as rf_process


APP_TITLE = "Collecting Data As Per Soundex Name Percentage"
DEFAULT_THRESHOLD = 80.0

# These weights are intentionally explicit and easy to tune later.
WEIGHTS = {
    "full": 0.35,
    "token_set": 0.20,
    "first": 0.15,
    "last": 0.15,
    "middle": 0.05,
    "phonetic": 0.10,
}

# Candidate generation is broader than the final merge threshold.
# This prevents a high final threshold from being defeated by small
# differences during candidate lookup.
CANDIDATE_SCORE_FLOOR = 55.0
MAX_CANDIDATES_PER_RECORD = 30


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).strip()


def normalize_basic(value: Any) -> str:
    s = clean_text(value)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.upper()
    s = re.sub(r"[^A-Z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def compact_alnum(value: Any) -> str:
    return re.sub(r"[^A-Z0-9]", "", normalize_basic(value))


def soundex(value: Any) -> str:
    """Standard English Soundex, applied to the normalized name."""
    s = re.sub(r"[^A-Z]", "", normalize_basic(value))
    if not s:
        return ""

    first = s[0]
    mapping = {
        **dict.fromkeys("BFPV", "1"),
        **dict.fromkeys("CGJKQSXZ", "2"),
        **dict.fromkeys("DT", "3"),
        "L": "4",
        **dict.fromkeys("MN", "5"),
        "R": "6",
    }

    result = first
    previous = mapping.get(first, "")
    for ch in s[1:]:
        code = mapping.get(ch, "")
        if code and code != previous:
            result += code
            if len(result) == 4:
                break
        previous = code

    return (result + "000")[:4]


def name_tokens(name: str) -> list[str]:
    return normalize_basic(name).split()


def split_name(name: str) -> tuple[str, str, str]:
    """
    Returns first, middle, last using a practical heuristic.

    For:
      Rahul Sharma       -> Rahul, "", Sharma
      Rahul S Sharma     -> Rahul, S, Sharma
      Sharma Rahul       -> Rahul, "", Sharma
      Sharma, Rahul S    -> Rahul, S, Sharma
    """
    raw = clean_text(name)
    raw = raw.replace(",", " , ")
    parts = [normalize_basic(p) for p in raw.split() if normalize_basic(p)]

    if not parts:
        return "", "", ""
    if len(parts) == 1:
        return parts[0], "", parts[0]

    # Comma means the portion before it is treated as surname.
    if "," in raw:
        left, right = raw.split(",", 1)
        last_parts = name_tokens(left)
        right_parts = name_tokens(right)
        if right_parts:
            first = right_parts[0]
            middle = " ".join(right_parts[1:-0]) if False else " ".join(right_parts[1:])
            last = " ".join(last_parts)
            return first, middle, last

    # Heuristic for surname-first without comma:
    # if there are 2+ tokens, compare in both orientations later.
    # Here we retain natural first/last order.
    first = parts[0]
    last = parts[-1]
    middle = " ".join(parts[1:-1])
    return first, middle, last


def canonical_name_tokens(name: str) -> list[str]:
    """
    Returns a deterministic token representation independent of token order.
    """
    return sorted(name_tokens(name))


def token_set_similarity(a: str, b: str) -> float:
    na = normalize_basic(a)
    nb = normalize_basic(b)
    if not na or not nb:
        return 0.0
    return float(fuzz.token_set_ratio(na, nb))


def order_aware_similarity(a: str, b: str) -> float:
    na = normalize_basic(a)
    nb = normalize_basic(b)
    if not na or not nb:
        return 0.0

    # Compare natural order and reversed order; this handles
    # "Sharma Rahul" vs "Rahul Sharma".
    direct = fuzz.ratio(na, nb)
    ra = " ".join(reversed(name_tokens(na)))
    rb = " ".join(reversed(name_tokens(nb)))
    reversed_score = fuzz.ratio(ra, rb)
    cross_a = fuzz.ratio(na, rb)
    cross_b = fuzz.ratio(ra, nb)
    return float(max(direct, reversed_score, cross_a, cross_b))


def component_similarity(a: str, b: str) -> tuple[float, float, float]:
    af, am, al = split_name(a)
    bf, bm, bl = split_name(b)

    # Also evaluate the opposite orientation for surname-first names.
    direct_first = fuzz.ratio(af, bf) if af and bf else 0.0
    direct_last = fuzz.ratio(al, bl) if al and bl else 0.0

    swapped_first = fuzz.ratio(af, bl) if af and bl else 0.0
    swapped_last = fuzz.ratio(al, bf) if al and bf else 0.0

    if (swapped_first + swapped_last) > (direct_first + direct_last):
        first_score, last_score = swapped_first, swapped_last
    else:
        first_score, last_score = direct_first, direct_last

    if am and bm:
        middle_score = float(fuzz.ratio(am, bm))
    elif not am and not bm:
        middle_score = 100.0
    else:
        # Missing middle name should not be treated as a strong mismatch.
        middle_score = 65.0

    return float(first_score), float(middle_score), float(last_score)


def phonetic_similarity(a: str, b: str) -> float:
    """
    Phonetic similarity is based on Soundex of the whole name and
    individual components. It is supporting evidence, not proof.
    """
    af, am, al = split_name(a)
    bf, bm, bl = split_name(b)

    whole_a = soundex(a)
    whole_b = soundex(b)

    score = 100.0 if whole_a and whole_a == whole_b else 0.0

    component_matches = 0
    component_total = 0
    for x, y in ((af, bf), (al, bl)):
        if x and y:
            component_total += 1
            if soundex(x) == soundex(y):
                component_matches += 1

    if component_total:
        component_score = 100.0 * component_matches / component_total
        score = max(score, component_score)

    return score


def calculate_match(a: str, b: str) -> dict[str, Any]:
    na = normalize_basic(a)
    nb = normalize_basic(b)

    if not na or not nb:
        return {
            "match_percentage": 0.0,
            "full_similarity": 0.0,
            "token_similarity": 0.0,
            "first_similarity": 0.0,
            "middle_similarity": 0.0,
            "last_similarity": 0.0,
            "phonetic_similarity": 0.0,
            "soundex_a": soundex(a),
            "soundex_b": soundex(b),
        }

    if na == nb:
        return {
            "match_percentage": 100.0,
            "full_similarity": 100.0,
            "token_similarity": 100.0,
            "first_similarity": 100.0,
            "middle_similarity": 100.0,
            "last_similarity": 100.0,
            "phonetic_similarity": 100.0,
            "soundex_a": soundex(a),
            "soundex_b": soundex(b),
        }

    full = order_aware_similarity(a, b)
    token = token_set_similarity(a, b)
    first, middle, last = component_similarity(a, b)
    phon = phonetic_similarity(a, b)

    score = (
        WEIGHTS["full"] * full
        + WEIGHTS["token_set"] * token
        + WEIGHTS["first"] * first
        + WEIGHTS["last"] * last
        + WEIGHTS["middle"] * middle
        + WEIGHTS["phonetic"] * phon
    )

    return {
        "match_percentage": round(max(0.0, min(100.0, score)), 2),
        "full_similarity": round(full, 2),
        "token_similarity": round(token, 2),
        "first_similarity": round(first, 2),
        "middle_similarity": round(middle, 2),
        "last_similarity": round(last, 2),
        "phonetic_similarity": round(phon, 2),
        "soundex_a": soundex(a),
        "soundex_b": soundex(b),
    }


@dataclass
class MatchRecord:
    row_index: int
    document_id: str
    full_name: str
    unique_id: str
    normalized_name: str
    soundex_code: str


class UnionFind:
    def __init__(self, n: int):
        self.parent = list(range(n))
        self.rank = [0] * n

    def find(self, x: int) -> int:
        while self.parent[x] != x:
            self.parent[x] = self.parent[self.parent[x]]
            x = self.parent[x]
        return x

    def union(self, a: int, b: int) -> None:
        ra, rb = self.find(a), self.find(b)
        if ra == rb:
            return
        if self.rank[ra] < self.rank[rb]:
            ra, rb = rb, ra
        self.parent[rb] = ra
        if self.rank[ra] == self.rank[rb]:
            self.rank[ra] += 1


def find_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    normalized = {normalize_basic(c): c for c in df.columns}
    for candidate in candidates:
        key = normalize_basic(candidate)
        if key in normalized:
            return normalized[key]
    return None


def make_unique_id(document_id: Any, full_name: Any) -> str:
    return f"{clean_text(document_id)}_ {clean_text(full_name)}"


def nonblank(value: Any) -> bool:
    if value is None:
        return False
    try:
        if pd.isna(value):
            return False
    except Exception:
        pass
    return clean_text(value) != ""


def values_for_cell(series: pd.Series) -> list[str]:
    values = []
    seen = set()
    for value in series:
        if nonblank(value):
            s = clean_text(value)
            key = normalize_basic(s)
            if key not in seen:
                seen.add(key)
                values.append(s)
    return values


def merged_cell(values: list[str]) -> str:
    if not values:
        return ""
    if len(values) == 1:
        return values[0]
    return " | ".join(values)


def safe_sheet_name(name: str) -> str:
    return re.sub(r"[:\\/?*\[\]]", "_", name)[:31]


def style_workbook(path: str) -> None:
    wb = load_workbook(path)
    header_fill = PatternFill(fill_type="solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)

    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # Keep width reasonable. Do not scan huge sheets excessively.
        for col_idx in range(1, ws.max_column + 1):
            letter = get_column_letter(col_idx)
            max_len = 0
            # Sample up to first 200 non-header rows for width.
            for row_idx in range(2, min(ws.max_row, 201) + 1):
                value = ws.cell(row_idx, col_idx).value
                if value is not None:
                    max_len = max(max_len, min(len(str(value)), 45))
            ws.column_dimensions[letter].width = max(10, min(max_len + 2, 45))

    wb.save(path)


class Processor:
    def __init__(self, progress_callback=None, status_callback=None):
        self.progress_callback = progress_callback or (lambda *args: None)
        self.status_callback = status_callback or (lambda *args: None)
        self.start_time = time.time()

    def update(self, fraction: float, task: str, processed: int, total: int) -> None:
        elapsed = max(time.time() - self.start_time, 0.001)
        if processed > 0:
            eta = elapsed * (total - processed) / processed
        else:
            eta = 0.0
        self.progress_callback(
            max(0.0, min(1.0, fraction)),
            task,
            processed,
            total,
            eta,
        )

    def process(self, input_path: str, output_path: str, threshold: float) -> dict[str, Any]:
        self.start_time = time.time()
        self.status_callback("Reading Excel file...")

        # dtype=object prevents unwanted numeric coercion for IDs/account numbers.
        df = pd.read_excel(input_path, dtype=object)
        if df.empty:
            raise ValueError("The input Excel file contains no data rows.")

        original_columns = list(df.columns)
        if len(original_columns) == 0:
            raise ValueError("No columns were found in the input Excel file.")

        document_col = find_column(df, ["Document ID", "DocumentID", "Document_Id"])
        name_col = find_column(df, ["Full Name", "FullName", "Name"])

        if document_col is None:
            raise ValueError("Could not find a 'Document ID' column.")
        if name_col is None:
            raise ValueError("Could not find a 'Full Name' column.")

        total_rows = len(df)
        self.update(0.05, "Excel loaded", 0, total_rows)

        # Preserve original data exactly as read.
        original_df = df.copy()

        self.status_callback("Creating Unique ID and normalized name data...")
        processed_df = df.copy()

        # Add processing columns without overwriting source columns.
        unique_ids = []
        normalized_names = []
        soundex_codes = []

        for i, row in df.iterrows():
            doc = row[document_col]
            name = row[name_col]
            unique_ids.append(make_unique_id(doc, name))
            normalized_names.append(normalize_basic(name))
            soundex_codes.append(soundex(name))
            if (i + 1) % max(1, total_rows // 100) == 0 or i + 1 == total_rows:
                self.update(
                    0.05 + 0.10 * (i + 1) / total_rows,
                    "Preparing records",
                    i + 1,
                    total_rows,
                )

        # Avoid collision if source already contains these names.
        processing_cols = ["Unique ID", "Normalized Name", "Soundex Code", "Match Group"]
        for col in processing_cols:
            if col in processed_df.columns:
                raise ValueError(
                    f"Input already contains a reserved processing column '{col}'. "
                    "Please rename/remove that source column and try again."
                )

        processed_df["Unique ID"] = unique_ids
        processed_df["Normalized Name"] = normalized_names
        processed_df["Soundex Code"] = soundex_codes

        records: list[MatchRecord] = []
        for i in range(total_rows):
            records.append(
                MatchRecord(
                    row_index=i,
                    document_id=clean_text(df.iloc[i][document_col]),
                    full_name=clean_text(df.iloc[i][name_col]),
                    unique_id=unique_ids[i],
                    normalized_name=normalized_names[i],
                    soundex_code=soundex_codes[i],
                )
            )

        self.status_callback("Building candidate name pairs...")
        # Candidate generation using RapidFuzz and soundex buckets.
        normalized_names_list = normalized_names
        soundex_buckets: dict[str, list[int]] = defaultdict(list)
        token_first_buckets: dict[str, list[int]] = defaultdict(list)

        for i, rec in enumerate(records):
            if rec.soundex_code:
                soundex_buckets[rec.soundex_code].append(i)
            tokens = name_tokens(rec.normalized_name)
            if tokens:
                token_first_buckets[tokens[0]].append(i)

        uf = UnionFind(total_rows)
        match_rows: list[dict[str, Any]] = []
        pair_seen: set[tuple[int, int]] = set()

        # Compare exact normalized names first.
        exact_groups: dict[str, list[int]] = defaultdict(list)
        for i, n in enumerate(normalized_names):
            if n:
                exact_groups[n].append(i)

        for indices in exact_groups.values():
            if len(indices) > 1:
                base = indices[0]
                for other in indices[1:]:
                    uf.union(base, other)
                    pair_seen.add((min(base, other), max(base, other)))
                    m = calculate_match(records[base].full_name, records[other].full_name)
                    match_rows.append(self._match_detail(records[base], records[other], m, threshold, "Exact normalized name"))

        # Candidate lookup for every row.
        for i, rec in enumerate(records):
            candidate_ids = set()

            # Soundex bucket.
            candidate_ids.update(soundex_buckets.get(rec.soundex_code, []))

            # Same first token bucket.
            toks = name_tokens(rec.normalized_name)
            if toks:
                candidate_ids.update(token_first_buckets.get(toks[0], []))

            # Fuzzy retrieval from all normalized names.
            if rec.normalized_name:
                fuzzy = rf_process.extract(
                    rec.normalized_name,
                    normalized_names_list,
                    scorer=fuzz.token_set_ratio,
                    limit=MAX_CANDIDATES_PER_RECORD + 1,
                    score_cutoff=CANDIDATE_SCORE_FLOOR,
                )
                for _, score, idx in fuzzy:
                    candidate_ids.add(idx)

            for j in candidate_ids:
                if j <= i:
                    continue
                pair = (i, j)
                if pair in pair_seen:
                    continue
                pair_seen.add(pair)

                m = calculate_match(rec.full_name, records[j].full_name)
                if m["match_percentage"] > CANDIDATE_SCORE_FLOOR:
                    # Only threshold-qualified matches create a merge edge.
                    if m["match_percentage"] > threshold:
                        uf.union(i, j)
                        status = "Merge Candidate"
                    else:
                        status = "Below Threshold"
                    match_rows.append(
                        self._match_detail(
                            rec,
                            records[j],
                            m,
                            threshold,
                            status,
                        )
                    )

            if (i + 1) % max(1, total_rows // 100) == 0 or i + 1 == total_rows:
                self.update(
                    0.15 + 0.45 * (i + 1) / total_rows,
                    "Comparing names",
                    i + 1,
                    total_rows,
                )

        self.status_callback("Resolving match groups...")
        groups: dict[int, list[int]] = defaultdict(list)
        for i in range(total_rows):
            groups[uf.find(i)].append(i)

        # Only give group IDs to actual multi-row groups.
        group_id_by_row: dict[int, str] = {}
        multi_groups = []
        counter = 1
        for root, members in groups.items():
            if len(members) > 1:
                gid = f"G{counter:04d}"
                counter += 1
                multi_groups.append((gid, members))
                for idx in members:
                    group_id_by_row[idx] = gid
            else:
                group_id_by_row[members[0]] = ""

        processed_df["Match Group"] = [
            group_id_by_row.get(i, "") for i in range(total_rows)
        ]

        # Enrich match detail rows with final group IDs.
        for row in match_rows:
            ia = row["_row_a"]
            ib = row["_row_b"]
            row["Match Group"] = group_id_by_row.get(ia, "") or group_id_by_row.get(ib, "")
            row.pop("_row_a", None)
            row.pop("_row_b", None)

        match_details_df = pd.DataFrame(match_rows)

        self.status_callback("Consolidating matched records...")
        final_columns = original_columns + [
            "Match Group",
            "Source Row Count",
            "Source Document IDs",
            "Source Unique IDs",
            "Group Match Range",
            "Conflict Columns",
        ]

        final_rows = []
        total_groups = len(multi_groups)

        # Multi-row groups.
        for group_pos, (gid, members) in enumerate(multi_groups, start=1):
            sub = df.iloc[members]
            output_row: dict[str, Any] = {}

            conflicts = []
            for col in original_columns:
                vals = values_for_cell(sub[col])
                output_row[col] = merged_cell(vals)
                if len(vals) > 1:
                    conflicts.append(str(col))

            # Match range from details belonging to group.
            group_scores = []
            for mr in match_rows:
                if mr.get("Match Group") == gid:
                    try:
                        group_scores.append(float(mr["Match Percentage"]))
                    except Exception:
                        pass

            output_row["Match Group"] = gid
            output_row["Source Row Count"] = len(members)
            output_row["Source Document IDs"] = " | ".join(
                values_for_cell(sub[document_col])
            )
            output_row["Source Unique IDs"] = " | ".join(
                unique_ids[i] for i in members
            )
            output_row["Group Match Range"] = (
                f"{min(group_scores):.2f}% - {max(group_scores):.2f}%"
                if group_scores
                else ""
            )
            output_row["Conflict Columns"] = " | ".join(conflicts)
            final_rows.append(output_row)

            if group_pos % max(1, total_groups // 100 or 1) == 0 or group_pos == total_groups:
                self.update(
                    0.60 + 0.25 * group_pos / max(1, total_groups),
                    "Consolidating matched groups",
                    group_pos,
                    max(1, total_groups),
                )

        # Keep unmatched rows as individual final records so no source record is lost.
        grouped_members = {idx for _, members in multi_groups for idx in members}
        unmatched = [i for i in range(total_rows) if i not in grouped_members]

        for pos, idx in enumerate(unmatched, start=1):
            row = df.iloc[idx]
            output_row = {col: row[col] for col in original_columns}
            output_row["Match Group"] = ""
            output_row["Source Row Count"] = 1
            output_row["Source Document IDs"] = clean_text(row[document_col])
            output_row["Source Unique IDs"] = unique_ids[idx]
            output_row["Group Match Range"] = ""
            output_row["Conflict Columns"] = ""
            final_rows.append(output_row)

            if pos % max(1, len(unmatched) // 100 or 1) == 0 or pos == len(unmatched):
                self.update(
                    0.85 + 0.05 * pos / max(1, len(unmatched)),
                    "Adding unmatched records",
                    pos,
                    max(1, len(unmatched)),
                )

        final_df = pd.DataFrame(final_rows, columns=final_columns)

        # Clean internal columns in match details and ensure useful ordering.
        if not match_details_df.empty:
            preferred = [
                "Match Group",
                "Document ID 1",
                "Unique ID 1",
                "Name 1",
                "Document ID 2",
                "Unique ID 2",
                "Name 2",
                "Match Percentage",
                "Full Similarity",
                "Token Similarity",
                "First Similarity",
                "Middle Similarity",
                "Last Similarity",
                "Phonetic Similarity",
                "Soundex 1",
                "Soundex 2",
                "Status",
                "Threshold",
                "Reason",
            ]
            ordered = [c for c in preferred if c in match_details_df.columns]
            ordered += [c for c in match_details_df.columns if c not in ordered]
            match_details_df = match_details_df[ordered]
        else:
            match_details_df = pd.DataFrame(
                columns=[
                    "Match Group",
                    "Document ID 1",
                    "Unique ID 1",
                    "Name 1",
                    "Document ID 2",
                    "Unique ID 2",
                    "Name 2",
                    "Match Percentage",
                    "Full Similarity",
                    "Token Similarity",
                    "First Similarity",
                    "Middle Similarity",
                    "Last Similarity",
                    "Phonetic Similarity",
                    "Soundex 1",
                    "Soundex 2",
                    "Status",
                    "Threshold",
                    "Reason",
                ]
            )

        # Excel has a 31,000-ish column limit in practical use; warn early.
        if final_df.shape[1] > 16000:
            raise ValueError(
                f"The output would contain {final_df.shape[1]} columns, which is beyond "
                "Excel's worksheet column limit."
            )

        self.status_callback("Writing output Excel workbook...")
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            original_df.to_excel(writer, sheet_name="Original Data", index=False)
            processed_df.to_excel(writer, sheet_name="Processed Data", index=False)
            match_details_df.to_excel(writer, sheet_name="Match Details", index=False)
            final_df.to_excel(writer, sheet_name="Final Merged Data", index=False)

        self.status_callback("Formatting Excel workbook...")
        style_workbook(output_path)

        elapsed = time.time() - self.start_time
        potential_matches = sum(
            1 for r in match_rows if r.get("Status") == "Merge Candidate"
        )

        self.update(1.0, "Completed", total_rows, total_rows)
        self.status_callback("Processing completed.")

        return {
            "total_records": total_rows,
            "total_columns": len(original_columns),
            "potential_matches": potential_matches,
            "match_groups": len(multi_groups),
            "output_path": output_path,
            "elapsed_seconds": elapsed,
        }

    @staticmethod
    def _match_detail(
        a: MatchRecord,
        b: MatchRecord,
        m: dict[str, Any],
        threshold: float,
        status: str,
    ) -> dict[str, Any]:
        if status == "Exact normalized name":
            reason = "Exact normalized name match."
            final_status = "Merge Candidate"
        elif status == "Merge Candidate":
            reason = (
                "Calculated matching percentage is above the configured threshold; "
                "grouped as a potential match."
            )
            final_status = "Merge Candidate"
        else:
            reason = "Similarity found, but it does not exceed the configured threshold."
            final_status = "Below Threshold"

        return {
            "_row_a": a.row_index,
            "_row_b": b.row_index,
            "Match Group": "",
            "Document ID 1": a.document_id,
            "Unique ID 1": a.unique_id,
            "Name 1": a.full_name,
            "Document ID 2": b.document_id,
            "Unique ID 2": b.unique_id,
            "Name 2": b.full_name,
            "Match Percentage": m["match_percentage"],
            "Full Similarity": m["full_similarity"],
            "Token Similarity": m["token_similarity"],
            "First Similarity": m["first_similarity"],
            "Middle Similarity": m["middle_similarity"],
            "Last Similarity": m["last_similarity"],
            "Phonetic Similarity": m["phonetic_similarity"],
            "Soundex 1": m["soundex_a"],
            "Soundex 2": m["soundex_b"],
            "Status": final_status,
            "Threshold": threshold,
            "Reason": reason,
        }


def open_path(path: str) -> None:
    path_obj = Path(path).resolve()
    if not path_obj.exists():
        return
    if sys.platform.startswith("win"):
        os.startfile(str(path_obj))
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path_obj)])
    else:
        subprocess.Popen(["xdg-open", str(path_obj)])


class App:
    def __init__(self):
        import tkinter as tk
        from tkinter import filedialog, messagebox, ttk

        self.tk = tk
        self.ttk = ttk
        self.filedialog = filedialog
        self.messagebox = messagebox

        self.root = tk.Tk()
        self.root.title(APP_TITLE)
        self.root.geometry("900x650")
        self.root.minsize(820, 600)

        self.input_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.threshold_var = tk.StringVar(value=str(int(DEFAULT_THRESHOLD)))
        self.progress_var = tk.DoubleVar(value=0)
        self.percent_var = tk.StringVar(value="0%")
        self.task_var = tk.StringVar(value="Waiting...")
        self.rows_var = tk.StringVar(value="Rows Processed: 0 / 0")
        self.eta_var = tk.StringVar(value="Estimated Time Remaining: --")
        self.status_var = tk.StringVar(value="Status: Ready")

        self.process_button = None
        self.log_text = None
        self.progress_bar = None

        self._build_ui()

    def _build_ui(self):
        tk = self.tk
        ttk = self.ttk

        outer = ttk.Frame(self.root, padding=18)
        outer.pack(fill="both", expand=True)

        title = ttk.Label(
            outer,
            text=APP_TITLE,
            font=("Segoe UI", 18, "bold"),
        )
        title.pack(anchor="center", pady=(0, 18))

        card = ttk.LabelFrame(outer, text="Input / Processing Settings", padding=15)
        card.pack(fill="x")

        self._path_row(card, 0, "1. Select File Path", self.input_var, self._browse_input)
        self._path_row(card, 1, "2. Save Excel File Path", self.output_var, self._browse_output)

        ttk.Label(card, text="3. Matching Merge Percentage").grid(
            row=2, column=0, sticky="w", pady=8
        )

        threshold_frame = ttk.Frame(card)
        threshold_frame.grid(row=2, column=1, sticky="w", padx=(10, 0), pady=8)
        threshold_entry = ttk.Entry(
            threshold_frame,
            textvariable=self.threshold_var,
            width=10,
        )
        threshold_entry.pack(side="left")
        ttk.Label(threshold_frame, text=" %").pack(side="left")

        ttk.Label(
            card,
            text=(
                "Records are considered merge candidates when their calculated "
                "matching percentage is greater than the selected threshold."
            ),
            wraplength=700,
        ).grid(row=3, column=1, sticky="w", padx=(10, 0), pady=(0, 10))

        card.columnconfigure(1, weight=1)

        action = ttk.Frame(outer)
        action.pack(fill="x", pady=15)

        self.process_button = ttk.Button(
            action,
            text="PROCESS",
            command=self._start_processing,
        )
        self.process_button.pack()

        progress_frame = ttk.LabelFrame(outer, text="Processing Progress", padding=15)
        progress_frame.pack(fill="both", expand=True)

        self.progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=100,
            mode="determinate",
        )
        self.progress_bar.pack(fill="x", pady=(0, 6))

        top_info = ttk.Frame(progress_frame)
        top_info.pack(fill="x")

        ttk.Label(top_info, textvariable=self.percent_var).pack(side="left")
        ttk.Label(top_info, textvariable=self.rows_var).pack(side="right")

        ttk.Label(progress_frame, textvariable=self.task_var).pack(
            anchor="w", pady=(12, 3)
        )
        ttk.Label(progress_frame, textvariable=self.eta_var).pack(anchor="w")
        ttk.Label(progress_frame, textvariable=self.status_var).pack(
            anchor="w", pady=(3, 10)
        )

        self.log_text = tk.Text(
            progress_frame,
            height=12,
            wrap="word",
            state="disabled",
            font=("Consolas", 9),
        )
        self.log_text.pack(fill="both", expand=True)

    def _path_row(self, parent, row, label, variable, command):
        ttk = self.ttk
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=8)

        frame = ttk.Frame(parent)
        frame.grid(row=row, column=1, sticky="ew", padx=(10, 0), pady=8)
        parent.columnconfigure(1, weight=1)

        ttk.Entry(frame, textvariable=variable).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(frame, text="Browse", command=command).pack(
            side="left", padx=(8, 0)
        )

    def _browse_input(self):
        path = self.filedialog.askopenfilename(
            title="Select Input Excel File",
            filetypes=[
                ("Excel files", "*.xlsx *.xlsm"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self.input_var.set(path)

            # Suggest output path automatically.
            p = Path(path)
            self.output_var.set(str(p.with_name(p.stem + "_Soundex_Matched.xlsx")))

    def _browse_output(self):
        path = self.filedialog.asksaveasfilename(
            title="Save Output Excel File",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if path:
            self.output_var.set(path)

    def _log(self, message: str):
        def do():
            self.log_text.configure(state="normal")
            self.log_text.insert("end", message + "\n")
            self.log_text.see("end")
            self.log_text.configure(state="disabled")
        self.root.after(0, do)

    def _start_processing(self):
        if getattr(self, "_running", False):
            return

        input_path = self.input_var.get().strip()
        output_path = self.output_var.get().strip()

        if not input_path:
            self.messagebox.showerror(APP_TITLE, "Please select the input Excel file.")
            return
        if not Path(input_path).exists():
            self.messagebox.showerror(APP_TITLE, "The selected input file does not exist.")
            return
        if not output_path:
            self.messagebox.showerror(APP_TITLE, "Please select the output Excel path.")
            return

        try:
            threshold = float(self.threshold_var.get().strip())
        except ValueError:
            self.messagebox.showerror(APP_TITLE, "Matching Merge Percentage must be a number.")
            return

        if not 0 <= threshold <= 100:
            self.messagebox.showerror(APP_TITLE, "Matching Merge Percentage must be between 0 and 100.")
            return

        if Path(input_path).resolve() == Path(output_path).resolve():
            self.messagebox.showerror(
                APP_TITLE,
                "Input and output files must be different."
            )
            return

        self._running = True
        self.process_button.configure(state="disabled")
        self._log("=" * 70)
        self._log(f"Started: {input_path}")
        self._log(f"Output:  {output_path}")
        self._log(f"Threshold: greater than {threshold:.2f}%")

        thread = threading.Thread(
            target=self._worker,
            args=(input_path, output_path, threshold),
            daemon=True,
        )
        thread.start()

    def _worker(self, input_path: str, output_path: str, threshold: float):
        try:
            processor = Processor(
                progress_callback=self._progress,
                status_callback=self._status,
            )
            result = processor.process(input_path, output_path, threshold)

            self._log(
                f"Completed. Records={result['total_records']}, "
                f"Columns={result['total_columns']}, "
                f"Potential Matches={result['potential_matches']}, "
                f"Groups={result['match_groups']}, "
                f"Time={result['elapsed_seconds']:.1f}s"
            )

            self.root.after(0, lambda: self._completed_dialog(result))
        except Exception as exc:
            error_text = f"{type(exc).__name__}: {exc}"
            self._log(error_text)
            self._log(traceback.format_exc())
            self.root.after(0, lambda: self._error_dialog(error_text))
        finally:
            self.root.after(0, self._finish_ui)

    def _progress(self, fraction, task, processed, total, eta):
        def do():
            pct = fraction * 100
            self.progress_var.set(pct)
            self.percent_var.set(f"{pct:.0f}%")
            self.task_var.set(f"Current Task: {task}")
            self.rows_var.set(f"Rows Processed: {processed:,} / {total:,}")
            if eta > 0:
                mins, secs = divmod(int(eta), 60)
                hours, mins = divmod(mins, 60)
                if hours:
                    eta_text = f"{hours:02d}:{mins:02d}:{secs:02d}"
                else:
                    eta_text = f"{mins:02d}:{secs:02d}"
                self.eta_var.set(f"Estimated Time Remaining: {eta_text}")
            else:
                self.eta_var.set("Estimated Time Remaining: --")
        self.root.after(0, do)

    def _status(self, text):
        self.root.after(0, lambda: self.status_var.set(f"Status: {text}"))

    def _completed_dialog(self, result):
        self.status_var.set("Status: Processing completed.")
        self.task_var.set("Current Task: Completed")
        self.progress_var.set(100)
        self.percent_var.set("100%")
        self.rows_var.set(
            f"Rows Processed: {result['total_records']:,} / {result['total_records']:,}"
        )
        self.eta_var.set("Estimated Time Remaining: 00:00")

        msg = (
            "PROCESSING COMPLETED\n\n"
            f"Total Records: {result['total_records']:,}\n"
            f"Source Columns: {result['total_columns']:,}\n"
            f"Potential Matches: {result['potential_matches']:,}\n"
            f"Match Groups: {result['match_groups']:,}\n\n"
            f"Output:\n{result['output_path']}"
        )

        answer = self.messagebox.askyesno(
            APP_TITLE,
            msg + "\n\nOpen the output Excel file now?",
        )
        if answer:
            open_path(result["output_path"])

    def _error_dialog(self, text):
        self.status_var.set("Status: Error.")
        self.messagebox.showerror(
            APP_TITLE,
            "Processing failed.\n\n" + text,
        )

    def _finish_ui(self):
        self._running = False
        self.process_button.configure(state="normal")

    def run(self):
        self.root.mainloop()


def main():
    app = App()
    app.run()


if __name__ == "__main__":
    main()
