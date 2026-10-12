"""
Collecting Data As Per Soundex Name Percentage
------------------------------------------------
Desktop Excel name-matching and consolidation application.

Requirements:
    Python 3.10+
    pip install pandas openpyxl rapidfuzz customtkinter jellyfish

Run:
    python soundex_name_matching.py

Output workbook:
    Original Data
    Processed Data
    Match Details
    Final Merged Data
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
import jellyfish
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from rapidfuzz import fuzz, process as rf_process
import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk

# Set up CustomTkinter appearance
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

APP_TITLE = "Collecting Data As Per Soundex Name Percentage"
DEFAULT_THRESHOLD = 80.0

WEIGHTS = {
    "full": 0.35,
    "token_set": 0.20,
    "first": 0.15,
    "last": 0.15,
    "middle": 0.05,
    "phonetic": 0.10,
}

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
    raw = clean_text(name)
    raw = raw.replace(",", " , ")
    parts = [normalize_basic(p) for p in raw.split() if normalize_basic(p)]

    if not parts:
        return "", "", ""
    if len(parts) == 1:
        return parts[0], "", parts[0]

    if "," in raw:
        left, right = raw.split(",", 1)
        last_parts = name_tokens(left)
        right_parts = name_tokens(right)
        if right_parts:
            first = right_parts[0]
            middle = " ".join(right_parts[1:])
            last = " ".join(last_parts)
            return first, middle, last

    first = parts[0]
    last = parts[-1]
    middle = " ".join(parts[1:-1])
    return first, middle, last


def canonical_name_tokens(name: str) -> list[str]:
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
        middle_score = 65.0

    return float(first_score), float(middle_score), float(last_score)


def phonetic_similarity(a: str, b: str) -> float:
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
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )

        for col_idx in range(1, ws.max_column + 1):
            letter = get_column_letter(col_idx)
            max_len = 0

            for row_idx in range(
                2,
                min(ws.max_row, 201) + 1
            ):
                value = ws.cell(row_idx, col_idx).value

                if value is not None:
                    max_len = max(
                        max_len,
                        min(len(str(value)), 45)
                    )

            ws.column_dimensions[letter].width = max(
                10,
                min(max_len + 2, 45)
            )

    wb.save(path)


class Processor:
    def __init__(
        self,
        progress_callback=None,
        status_callback=None,
        selected_columns=None
    ):
        self.progress_callback = progress_callback or (
            lambda *args: None
        )
        self.status_callback = status_callback or (
            lambda *args: None
        )
        self.selected_columns = selected_columns or []
        self.start_time = time.time()

    def update(
        self,
        fraction: float,
        task: str,
        processed: int,
        total: int
    ) -> None:

        elapsed = max(
            time.time() - self.start_time,
            0.001
        )

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

    def _match_detail(self, r1: MatchRecord, r2: MatchRecord, m: dict[str, Any], threshold: float, status: str) -> dict[str, Any]:
        return {
            "Match Group": "",
            "Document ID 1": r1.document_id,
            "Unique ID 1": r1.unique_id,
            "Name 1": r1.full_name,
            "Document ID 2": r2.document_id,
            "Unique ID 2": r2.unique_id,
            "Name 2": r2.full_name,
            "Match Percentage": m["match_percentage"],
            "Full Similarity": m["full_similarity"],
            "Token Similarity": m["token_similarity"],
            "First Similarity": m["first_similarity"],
            "Middle Similarity": m["middle_similarity"],
            "Last Similarity": m["last_similarity"],
            "Phonetic Similarity": m["phonetic_similarity"],
            "Soundex 1": m["soundex_a"],
            "Soundex 2": m["soundex_b"],
            "Status": status,
            "Threshold": threshold,
            "Reason": "Soundex/String similarity match",
            "_row_a": r1.row_index,
            "_row_b": r2.row_index,
        }

    def process(
        self,
        input_path: str,
        output_path: str,
        threshold: float
    ) -> dict[str, Any]:

        self.start_time = time.time()

        self.status_callback("Reading Excel file...")

        df = pd.read_excel(
            input_path,
            dtype=object
        )

        if df.empty:
            raise ValueError(
                "The input Excel file contains no data rows."
            )

        original_columns = list(df.columns)

        if len(original_columns) == 0:
            raise ValueError(
                "No columns were found in the input Excel file."
            )

        document_col = find_column(
            df,
            [
                "Document ID",
                "DocumentID",
                "Document_Id"
            ]
        )

        name_col = find_column(
            df,
            [
                "Full Name",
                "FullName",
                "Name"
            ]
        )

        if document_col is None:
            raise ValueError(
                "Could not find a 'Document ID' column."
            )

        if name_col is None:
            raise ValueError(
                "Could not find a 'Full Name' column."
            )

        total_rows = len(df)

        self.update(
            0.05,
            "Excel loaded",
            0,
            total_rows
        )

        original_df = df.copy()

        self.status_callback(
            "Creating Unique ID and normalized name data..."
        )

        processed_df = df.copy()

        unique_ids = []
        normalized_names = []
        soundex_codes = []

        for i, row in df.iterrows():

            doc = row[document_col]
            name = row[name_col]

            unique_ids.append(
                make_unique_id(doc, name)
            )

            normalized_names.append(
                normalize_basic(name)
            )

            soundex_codes.append(
                soundex(name)
            )

            if (
                (i + 1) % max(1, total_rows // 100) == 0
                or i + 1 == total_rows
            ):
                self.update(
                    0.05
                    + 0.10 * (i + 1) / total_rows,
                    "Preparing records",
                    i + 1,
                    total_rows,
                )

        processing_cols = [
            "Unique ID",
            "Normalized Name",
            "Soundex Code",
            "Match Group"
        ]

        for col in processing_cols:
            if col in processed_df.columns:
                raise ValueError(
                    f"Input already contains a reserved processing "
                    f"column '{col}'. Please rename/remove that "
                    f"source column and try again."
                )

        processed_df["Unique ID"] = unique_ids
        processed_df["Normalized Name"] = normalized_names
        processed_df["Soundex Code"] = soundex_codes

        records: list[MatchRecord] = []

        for i in range(total_rows):

            records.append(
                MatchRecord(
                    row_index=i,
                    document_id=clean_text(
                        df.iloc[i][document_col]
                    ),
                    full_name=clean_text(
                        df.iloc[i][name_col]
                    ),
                    unique_id=unique_ids[i],
                    normalized_name=normalized_names[i],
                    soundex_code=soundex_codes[i],
                )
            )

        self.status_callback(
            "Building candidate name pairs..."
        )

        normalized_names_list = normalized_names

        soundex_buckets: dict[str, list[int]] = defaultdict(list)
        token_first_buckets: dict[str, list[int]] = defaultdict(list)
        document_buckets: dict[str, list[int]] = defaultdict(list)

        for i, rec in enumerate(records):

            if rec.soundex_code:
                soundex_buckets[
                    rec.soundex_code
                ].append(i)

            tokens = name_tokens(
                rec.normalized_name
            )

            if tokens:
                token_first_buckets[
                    tokens[0]
                ].append(i)

            document_buckets[
                rec.document_id
            ].append(i)

        uf = UnionFind(total_rows)

        match_rows: list[dict[str, Any]] = []

        pair_seen: set[tuple[int, int]] = set()

        exact_groups: dict[
            tuple[str, str],
            list[int]
        ] = defaultdict(list)

        for i, n in enumerate(normalized_names):

            if n:
                exact_groups[
                    (
                        records[i].document_id,
                        n
                    )
                ].append(i)

        for indices in exact_groups.values():

            if len(indices) > 1:

                base = indices[0]

                for other in indices[1:]:

                    if (
                        records[base].document_id
                        != records[other].document_id
                    ):
                        continue

                    uf.union(
                        base,
                        other
                    )

                    pair_seen.add(
                        (
                            min(base, other),
                            max(base, other)
                        )
                    )

                    m = calculate_match(
                        records[base].full_name,
                        records[other].full_name
                    )

                    match_rows.append(
                        self._match_detail(
                            records[base],
                            records[other],
                            m,
                            threshold,
                            "Exact normalized name"
                        )
                    )

        for i, rec in enumerate(records):

            candidate_ids = set()

            same_document_ids = set(
                document_buckets.get(
                    rec.document_id,
                    []
                )
            )

            for candidate_id in soundex_buckets.get(
                rec.soundex_code,
                []
            ):
                if candidate_id in same_document_ids:
                    candidate_ids.add(candidate_id)

            toks = name_tokens(
                rec.normalized_name
            )

            if toks:

                for candidate_id in token_first_buckets.get(
                    toks[0],
                    []
                ):
                    if candidate_id in same_document_ids:
                        candidate_ids.add(candidate_id)

            if rec.normalized_name:

                fuzzy = rf_process.extract(
                    rec.normalized_name,
                    normalized_names_list,
                    scorer=fuzz.token_set_ratio,
                    limit=MAX_CANDIDATES_PER_RECORD + 1,
                    score_cutoff=CANDIDATE_SCORE_FLOOR,
                )

                for _, score, idx in fuzzy:

                    if idx in same_document_ids:
                        candidate_ids.add(idx)

            for j in candidate_ids:

                if j <= i:
                    continue

                if (
                    records[i].document_id
                    != records[j].document_id
                ):
                    continue

                pair = (i, j)

                if pair in pair_seen:
                    continue

                pair_seen.add(pair)

                m = calculate_match(
                    rec.full_name,
                    records[j].full_name
                )

                if (
                    m["match_percentage"]
                    > CANDIDATE_SCORE_FLOOR
                ):

                    if (
                        m["match_percentage"]
                        > threshold
                    ):
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

            if (
                (i + 1) % max(1, total_rows // 100) == 0
                or i + 1 == total_rows
            ):
                self.update(
                    0.15
                    + 0.45 * (i + 1) / total_rows,
                    "Comparing names",
                    i + 1,
                    total_rows,
                )

        self.status_callback(
            "Resolving match groups..."
        )

        groups: dict[int, list[int]] = defaultdict(list)

        for i in range(total_rows):
            groups[
                uf.find(i)
            ].append(i)

        group_id_by_row: dict[int, str] = {}

        multi_groups = []

        counter = 1

        for root, members in groups.items():

            if len(members) > 1:

                gid = f"G{counter:04d}"

                counter += 1

                multi_groups.append(
                    (
                        gid,
                        members
                    )
                )

                for idx in members:
                    group_id_by_row[idx] = gid

            else:
                group_id_by_row[
                    members[0]
                ] = ""

        processed_df["Match Group"] = [
            group_id_by_row.get(i, "")
            for i in range(total_rows)
        ]

        for row in match_rows:

            ia = row["_row_a"]
            ib = row["_row_b"]

            row["Match Group"] = (
                group_id_by_row.get(ia, "")
                or group_id_by_row.get(ib, "")
            )

            row.pop("_row_a", None)
            row.pop("_row_b", None)

        match_details_df = pd.DataFrame(
            match_rows
        )

        self.status_callback(
            "Consolidating matched records..."
        )

        final_columns = original_columns + [
            "Match Group",
            "Source Row Count",
            "Source Document IDs",
            "Source Unique IDs",
            "Group Match Range",
            "Conflict Columns",
            "Information Same Merge",
            "Information Name Same Don't Merge",
        ]

        final_rows = []

        total_groups = len(multi_groups)

        for group_pos, (
            gid,
            members
        ) in enumerate(
            multi_groups,
            start=1
        ):

            sub = df.iloc[members]
            
            has_conflict = False
            if self.selected_columns:
                for col in self.selected_columns:
                    if col in df.columns:
                        val_set = set(str(val) for val in sub[col].dropna().unique())
                        if len(val_set) > 1:
                            has_conflict = True
                            break

            if has_conflict:
                for idx_m in members:
                    row_m = df.iloc[idx_m]
                    output_row = {col: row_m[col] for col in original_columns}
                    output_row["Match Group"] = gid
                    output_row["Source Row Count"] = 1
                    output_row["Source Document IDs"] = clean_text(row_m[document_col])
                    output_row["Source Unique IDs"] = unique_ids[idx_m]
                    output_row["Group Match Range"] = ""
                    output_row["Conflict Columns"] = ""
                    output_row["Information Same Merge"] = ""
                    
                    conflict_details = []
                    for col in self.selected_columns:
                        if col in row_m:
                            conflict_details.append(f"{col} ({row_m[col]})")
                    details_str = " | ".join(conflict_details)
                    output_row["Information Name Same Don't Merge"] = f"{row_m.get(name_col)} | {details_str}" if details_str else str(row_m.get(name_col))
                    
                    final_rows.append(output_row)
            else:
                output_row: dict[str, Any] = {}
                conflicts = []

                for col in original_columns:
                    vals = values_for_cell(
                        sub[col]
                    )
                    output_row[col] = merged_cell(
                        vals
                    )
                    if len(vals) > 1:
                        conflicts.append(
                            str(col)
                        )

                group_scores = []
                for mr in match_rows:
                    if mr.get("Match Group") == gid:
                        try:
                            group_scores.append(
                                float(
                                    mr["Match Percentage"]
                                )
                            )
                        except Exception:
                            pass

                output_row["Match Group"] = gid
                output_row["Source Row Count"] = len(members)
                output_row["Source Document IDs"] = (
                    " | ".join(
                        values_for_cell(
                            sub[document_col]
                        )
                    )
                )
                output_row["Source Unique IDs"] = (
                    " | ".join(
                        unique_ids[i]
                        for i in members
                    )
                )
                output_row["Group Match Range"] = (
                    f"{min(group_scores):.2f}% - "
                    f"{max(group_scores):.2f}%"
                    if group_scores
                    else ""
                )
                output_row["Conflict Columns"] = (
                    " | ".join(conflicts)
                )
                output_row["Information Same Merge"] = ", ".join(values_for_cell(sub[name_col]))
                output_row["Information Name Same Don't Merge"] = ""

                final_rows.append(
                    output_row
                )

            if (
                group_pos
                % max(
                    1,
                    total_groups // 100 or 1
                ) == 0
                or group_pos == total_groups
            ):

                self.update(
                    0.60
                    + 0.25
                    * group_pos
                    / max(1, total_groups),
                    "Consolidating matched groups",
                    group_pos,
                    max(1, total_groups),
                )

        grouped_members = {
            idx
            for _, members in multi_groups
            for idx in members
        }

        unmatched = [
            i
            for i in range(total_rows)
            if i not in grouped_members
        ]

        for pos, idx in enumerate(
            unmatched,
            start=1
        ):

            row = df.iloc[idx]

            output_row = {
                col: row[col]
                for col in original_columns
            }

            output_row["Match Group"] = ""
            output_row["Source Row Count"] = 1
            output_row["Source Document IDs"] = (
                clean_text(
                    row[document_col]
                )
            )
            output_row["Source Unique IDs"] = (
                unique_ids[idx]
            )
            output_row["Group Match Range"] = ""
            output_row["Conflict Columns"] = ""
            output_row["Information Same Merge"] = clean_text(row[name_col])
            output_row["Information Name Same Don't Merge"] = ""

            final_rows.append(
                output_row
            )

            if (
                pos
                % max(
                    1,
                    len(unmatched) // 100 or 1
                ) == 0
                or pos == len(unmatched)
            ):

                self.update(
                    0.85
                    + 0.05
                    * pos
                    / max(1, len(unmatched)),
                    "Adding unmatched records",
                    pos,
                    max(1, len(unmatched)),
                )

        final_df = pd.DataFrame(
            final_rows,
            columns=final_columns
        )

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

            ordered = [
                c
                for c in preferred
                if c in match_details_df.columns
            ]

            ordered += [
                c
                for c in match_details_df.columns
                if c not in ordered
            ]

            match_details_df = match_details_df[
                ordered
            ]

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

        if final_df.shape[1] > 16000:

            raise ValueError(
                f"The output would contain "
                f"{final_df.shape[1]} columns, which is beyond "
                "Excel's worksheet column limit."
            )

        self.status_callback(
            "Writing output Excel workbook..."
        )

        with pd.ExcelWriter(
            output_path,
            engine="openpyxl"
        ) as writer:

            original_df.to_excel(
                writer,
                sheet_name="Original Data",
                index=False
            )

            processed_df.to_excel(
                writer,
                sheet_name="Processed Data",
                index=False
            )

            match_details_df.to_excel(
                writer,
                sheet_name="Match Details",
                index=False
            )

            final_df.to_excel(
                writer,
                sheet_name="Final Merged Data",
                index=False
            )

        self.status_callback(
            "Formatting Excel workbook..."
        )

        style_workbook(
            output_path
        )

        elapsed = (
            time.time()
            - self.start_time
        )

        potential_matches = sum(
            1
            for r in match_rows
            if r.get("Status")
            == "Merge Candidate"
        )

        self.update(
            1.0,
            "Completed",
            total_rows,
            total_rows
        )

        self.status_callback(
            "Processing completed."
        )

        return {
            "total_records": total_rows,
            "total_columns": len(
                original_columns
            ),
        }


# --- STEP 1: Interactive Data Analysis Preview Dialog ---
class AnalysisDialog(ctk.CTkToplevel):
    def __init__(self, parent, df, input_path):
        super().__init__(parent)
        self.title("Data Analysis Preview")
        self.geometry("950x600")
        self.df = df
        self.input_path = input_path
        self.proceed = False

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # --- Top Search Bar ---
        self.filter_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.filter_frame.grid(row=0, column=0, padx=20, pady=(15, 5), sticky="ew")
        self.filter_frame.grid_columnconfigure(1, weight=1)

        self.filter_label = ctk.CTkLabel(self.filter_frame, text="🔍 Filter Data:", font=ctk.CTkFont(weight="bold"))
        self.filter_label.grid(row=0, column=0, padx=(0, 10))

        self.filter_entry = ctk.CTkEntry(self.filter_frame, placeholder_text="Type to search rows instantly...", height=36)
        self.filter_entry.grid(row=0, column=1, sticky="ew")
        self.filter_entry.bind("<KeyRelease>", self.filter_table)

        # --- Data Table Preview Grid ---
        self.table_card = ctk.CTkFrame(self)
        self.table_card.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")
        self.table_card.grid_rowconfigure(0, weight=1)
        self.table_card.grid_columnconfigure(0, weight=1)

        self.table_frame = ctk.CTkFrame(self.table_card, fg_color="transparent")
        self.table_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        self.table_frame.grid_rowconfigure(0, weight=1)
        self.table_frame.grid_columnconfigure(0, weight=1)

        self.tree_scroll_y = ttk.Scrollbar(self.table_frame, orient="vertical")
        self.tree_scroll_x = ttk.Scrollbar(self.table_frame, orient="horizontal")
        
        cols = list(df.columns)
        self.tree = ttk.Treeview(
            self.table_frame, 
            columns=cols, 
            show="headings", 
            yscrollcommand=self.tree_scroll_y.set,
            xscrollcommand=self.tree_scroll_x.set
        )
        self.tree_scroll_y.config(command=self.tree.yview)
        self.tree_scroll_x.config(command=self.tree.xview)

        for col in cols:
            self.tree.heading(str(col), text=str(col))
            self.tree.column(str(col), width=130, anchor="w")

        self.tree.grid(row=0, column=0, sticky="nsew")
        self.tree_scroll_y.grid(row=0, column=1, sticky="ns")
        self.tree_scroll_x.grid(row=1, column=0, sticky="ew")

        self.populate_tree(df)

        # --- Bottom Proceed Button ---
        self.next_btn = ctk.CTkButton(
            self, 
            text="Proceed to Header Selection ➔", 
            font=ctk.CTkFont(size=14, weight="bold"),
            height=44,
            command=self.on_next
        )
        self.next_btn.grid(row=2, column=0, padx=20, pady=15, sticky="ew")

        self.transient(parent)
        self.grab_set()

    def populate_tree(self, data_df):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for _, row in data_df.head(200).iterrows():
            self.tree.insert("", "end", values=list(row.values))

    def filter_table(self, event):
        query = self.filter_entry.get().lower()
        if not query:
            self.populate_tree(self.df)
            return
        filtered = self.df[self.df.astype(str).apply(lambda x: x.str.lower().str.contains(query).any(), axis=1)]
        self.populate_tree(filtered)

    def on_next(self):
        self.proceed = True
        self.destroy()


# --- STEP 2: Dedicated Header Selection Pop-Up Dialog ---
class HeaderSelectionDialog(ctk.CTkToplevel):
    def __init__(self, parent, columns):
        super().__init__(parent)
        self.title("Select Non-Merge Conflict Headers")
        self.geometry("500x550")
        self.columns = columns
        self.selected_headers = []
        self.confirmed = False

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # Instruction Header
        self.title_label = ctk.CTkLabel(
            self, 
            text="Select Headers for Duplicate Non-Merge Rule:\n(If values differ between matched names, don't merge)", 
            font=ctk.CTkFont(size=13, weight="bold"),
            justify="center"
        )
        self.title_label.grid(row=0, column=0, padx=20, pady=(20, 10))

        # Checkbox Scrollable Frame
        self.scroll_frame = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.scroll_frame.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")

        self.header_vars = {}
        
        # "None" Checkbox Option
        self.none_var = ctk.BooleanVar(value=True)
        self.none_chk = ctk.CTkCheckBox(
            self.scroll_frame, 
            text="None (Bypass Rule)", 
            variable=self.none_var, 
            command=self.on_none_selected,
            font=ctk.CTkFont(weight="bold")
        )
        self.none_chk.pack(anchor="w", pady=(5, 12))

        # Populate all dataset columns as checkboxes
        for col in columns:
            var = ctk.BooleanVar(value=False)
            chk = ctk.CTkCheckBox(self.scroll_frame, text=str(col), variable=var, command=self.on_column_selected)
            chk.pack(anchor="w", pady=6)
            self.header_vars[str(col)] = var

        # Final Export Button
        self.export_btn = ctk.CTkButton(
            self, 
            text="🚀 Confirm & Final Export", 
            font=ctk.CTkFont(size=14, weight="bold"),
            height=46,
            fg_color="#2b8a3e", hover_color="#236b31",
            command=self.on_confirm
        )
        self.export_btn.grid(row=2, column=0, padx=20, pady=20, sticky="ew")

        self.transient(parent)
        self.grab_set()

    def on_none_selected(self):
        if self.none_var.get():
            for col, var in self.header_vars.items():
                var.set(False)

    def on_column_selected(self):
        if any(var.get() for var in self.header_vars.values()):
            self.none_var.set(False)

    def on_confirm(self):
        if not self.none_var.get():
            self.selected_headers = [col for col, var in self.header_vars.items() if var.get()]
        else:
            self.selected_headers = []
        
        self.confirmed = True
        self.destroy()


# --- Main Application GUI Class ---
class SoundexApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title(APP_TITLE)
        self.geometry("680x520")
        self.resizable(False, False)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(5, weight=1)

        self.title_label = ctk.CTkLabel(
            self, 
            text=APP_TITLE, 
            font=ctk.CTkFont(size=18, weight="bold")
        )
        self.title_label.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="w")

        self.settings_frame = ctk.CTkFrame(self)
        self.settings_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        self.settings_frame.grid_columnconfigure(1, weight=1)

        self.input_label = ctk.CTkLabel(self.settings_frame, text="1. Select File Path", anchor="w")
        self.input_label.grid(row=0, column=0, padx=15, pady=12, sticky="w")

        self.input_entry = ctk.CTkEntry(self.settings_frame, placeholder_text="Select input Excel file...")
        self.input_entry.grid(row=0, column=1, padx=10, pady=12, sticky="ew")

        self.input_browse = ctk.CTkButton(self.settings_frame, text="Browse", width=90, command=self.browse_input)
        self.input_browse.grid(row=0, column=2, padx=15, pady=12)

        self.output_label = ctk.CTkLabel(self.settings_frame, text="2. Save Excel File Path", anchor="w")
        self.output_label.grid(row=1, column=0, padx=15, pady=12, sticky="w")

        self.output_entry = ctk.CTkEntry(self.settings_frame, placeholder_text="Select output Excel destination...")
        self.output_entry.grid(row=1, column=1, padx=10, pady=12, sticky="ew")

        self.output_browse = ctk.CTkButton(self.settings_frame, text="Browse", width=90, command=self.browse_output)
        self.output_browse.grid(row=1, column=2, padx=15, pady=12)

        self.thresh_label = ctk.CTkLabel(self.settings_frame, text="3. Matching Merge Percentage", anchor="w")
        self.thresh_label.grid(row=2, column=0, padx=15, pady=12, sticky="w")

        self.thresh_entry = ctk.CTkEntry(self.settings_frame, width=70)
        self.thresh_entry.insert(0, str(int(DEFAULT_THRESHOLD)))
        self.thresh_entry.grid(row=2, column=1, padx=10, pady=12, sticky="w")

        self.thresh_percent = ctk.CTkLabel(self.settings_frame, text="%")
        self.thresh_percent.grid(row=2, column=1, padx=(85, 0), pady=12, sticky="w")

        self.process_btn = ctk.CTkButton(
            self, 
            text="Analyze Data", 
            font=ctk.CTkFont(size=15, weight="bold"),
            height=42,
            command=self.open_analysis_dialog
        )
        self.process_btn.grid(row=2, column=0, padx=20, pady=10, sticky="ew")

        self.progress_frame = ctk.CTkFrame(self, corner_radius=8)
        self.progress_frame.grid(row=3, column=0, padx=20, pady=10, sticky="ew")
        self.progress_frame.grid_columnconfigure(0, weight=1)

        self.progress_bar = ctk.CTkProgressBar(self.progress_frame)
        self.progress_bar.set(0.0)
        self.progress_bar.grid(row=0, column=0, padx=15, pady=(15, 5), sticky="ew")

        self.rows_status_label = ctk.CTkLabel(self.progress_frame, text="Rows Processed: 0 / 0", anchor="e")
        self.rows_status_label.grid(row=1, column=0, padx=15, pady=(0, 5), sticky="e")

        self.status_textbox = ctk.CTkTextbox(self.progress_frame, height=90)
        self.status_textbox.grid(row=2, column=0, padx=15, pady=(0, 15), sticky="ew")
        self.status_textbox.insert("0.0", "Waiting...\nStatus: Ready")
        self.status_textbox.configure(state="disabled")

    def browse_input(self):
        filename = filedialog.askopenfilename(
            title="Select Input Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if filename:
            self.input_entry.delete(0, "end")
            self.input_entry.insert(0, filename)
            
            base_dir = os.path.dirname(filename)
            base_name = os.path.splitext(os.path.basename(filename))[0]
            default_out = os.path.join(base_dir, f"{base_name}_Consolidated.xlsx")
            self.output_entry.delete(0, "end")
            self.output_entry.insert(0, default_out)

    def browse_output(self):
        filename = filedialog.asksaveasfilename(
            title="Save Consolidated Excel File As",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if filename:
            self.output_entry.delete(0, "end")
            self.output_entry.insert(0, filename)

    def log_status(self, text):
        self.status_textbox.configure(state="normal")
        self.status_textbox.delete("0.0", "end")
        self.status_textbox.insert("0.0", text)
        self.status_textbox.configure(state="disabled")

    def set_progress(self, fraction, task, processed, total, eta):
        self.after(0, lambda: self.progress_bar.set(fraction))
        self.after(0, lambda: self.rows_status_label.configure(text=f"Rows Processed: {processed} / {total}"))
        self.after(0, lambda: self.log_status(f"{task}\nStatus: Processing..."))

    def open_analysis_dialog(self):
        input_path = self.input_entry.get().strip()
        output_path = self.output_entry.get().strip()

        if not input_path or not os.path.exists(input_path):
            messagebox.showerror("Error", "Please select a valid input Excel file path.")
            return

        if not output_path:
            messagebox.showerror("Error", "Please specify a valid save Excel file path.")
            return

        try:
            threshold = float(self.thresh_entry.get().strip())
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid numeric percentage threshold.")
            return

        try:
            df = pd.read_excel(input_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read input file:\n{str(e)}")
            return

        # Step 1: Open Data Analysis Preview Dialog
        analysis_dlg = AnalysisDialog(self, df, input_path)
        self.wait_window(analysis_dlg)

        if not analysis_dlg.proceed:
            return

        # Step 2: Open Dedicated Header Selection Dialog
        header_dlg = HeaderSelectionDialog(self, list(df.columns))
        self.wait_window(header_dlg)

        if header_dlg.confirmed:
            selected_headers = header_dlg.selected_headers
            self.process_btn.configure(state="disabled", text="Processing...")
            threading.Thread(
                target=self.run_backend_processing,
                args=(input_path, output_path, threshold, selected_headers),
                daemon=True
            ).start()

    def run_backend_processing(self, input_path, output_path, threshold, selected_headers):
        try:
            processor = Processor(
                progress_callback=self.set_progress,
                status_callback=lambda text: self.after(0, lambda: self.log_status(text)),
                selected_columns=selected_headers
            )
            result = processor.process(input_path, output_path, threshold)
            
            self.after(0, lambda: self.log_status("Completed successfully!\nStatus: Ready"))
            self.after(0, lambda: messagebox.showinfo(
                "Success", 
                f"Processing complete!\nTotal Records: {result['total_records']}\nSaved to:\n{output_path}"
            ))
        except Exception as e:
            tb = traceback.format_exc()
            self.after(0, lambda: self.log_status(f"Error: {str(e)}"))
            self.after(0, lambda: messagebox.showerror("Processing Error", f"An error occurred:\n{str(e)}"))
            print(tb, file=sys.stderr)
        finally:
            self.after(0, lambda: self.process_btn.configure(state="normal", text="Analyze Data"))


if __name__ == "__main__":
    app = SoundexApp()
    app.mainloop()
