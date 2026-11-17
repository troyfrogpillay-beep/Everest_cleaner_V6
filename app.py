import io
import re
import difflib
from copy import copy
from dataclasses import dataclass, field
from datetime import datetime
from typing import Tuple, List, Dict, Any, Optional, NamedTuple

import openpyxl
from openpyxl.utils import get_column_letter
import pandas as pd
from email_validator import validate_email, EmailNotValidError
import streamlit as st

st.set_page_config(page_title="Everest Data Cleaner", layout="wide")

# ==============================================================================
# PRE-CLEANING STEPS: CLIENT SECTION EXTRACTION (ADDITIVE - NEVER TOUCHES THE
# SOURCE WORKSHEET OR THE DATAFRAME USED BY THE CLEANER)
# ==============================================================================

SUMMARY_SHEET_NAME = "Client Premium Summary"
DETAIL_SHEET_NAME = "Client Detailed Schedule"

SUMMARY_HEADERS = ["Payment Date", "Month End", "Policy Name", "Client Name", "Risk Payment"]
DETAIL_HEADERS = ["Client Name", "Member ID", "Surname", "First Name", "Policy Description", "Risk Premium"]


class ExtractionError(Exception):
    """Raised with a user-readable message when an extraction cannot proceed."""


class CellSnapshot(NamedTuple):
    """Independent copy of one source cell (value + type + formatting)."""
    value: Any
    data_type: str
    number_format: str
    font: Any
    fill: Any
    border: Any
    alignment: Any
    protection: Any


@dataclass
class ExtractedSection:
    section_label: str
    default_sheet_name: str
    source_sheet: str
    header_row: int
    first_row: int
    last_row: int
    first_col: int
    last_col: int
    header: List[CellSnapshot] = field(default_factory=list)
    rows: List[List[CellSnapshot]] = field(default_factory=list)
    source_rows: List[int] = field(default_factory=list)
    col_widths: Dict[int, float] = field(default_factory=dict)

    @property
    def section_range(self) -> str:
        return (f"'{self.source_sheet}'!{get_column_letter(self.first_col)}{self.header_row}:"
                f"{get_column_letter(self.last_col)}{self.last_row}")


def _norm_header(val: Any) -> str:
    if not isinstance(val, str):
        return ""
    return re.sub(r"\s+", " ", val.strip()).casefold()


def _norm_client(val: Any) -> str:
    if val is None:
        return ""
    return str(val).strip().casefold()


def _is_blank(val: Any) -> bool:
    return val is None or (isinstance(val, str) and not val.strip())


def _sheet_grid(ws) -> List[tuple]:
    return list(ws.iter_rows(values_only=True))


def _header_positions(row: tuple) -> Dict[str, List[int]]:
    pos: Dict[str, List[int]] = {}
    for idx, v in enumerate(row, start=1):
        key = _norm_header(v)
        if key:
            pos.setdefault(key, []).append(idx)
    return pos


def _matching_header_rows(grid: List[tuple], required: List[str]) -> List[Tuple[int, Dict[str, List[int]]]]:
    req = [_norm_header(h) for h in required]
    hits = []
    for r_idx, row in enumerate(grid, start=1):
        pos = _header_positions(row)
        if all(h in pos for h in req):
            hits.append((r_idx, pos))
    return hits


def _best_partial(grid: List[tuple], required: List[str]) -> Tuple[int, int, List[str]]:
    req = [(h, _norm_header(h)) for h in required]
    best = (0, 0, [h for h, _ in req])
    for r_idx, row in enumerate(grid, start=1):
        pos = _header_positions(row)
        found = [h for h, n in req if n in pos]
        if len(found) > best[1]:
            best = (r_idx, len(found), [h for h, n in req if n not in pos])
    return best


def _extract_section(wb, client_name: str, required: List[str], other_required: List[str],
                     label: str, default_sheet_name: str, stop_at_first_blank_row: bool) -> ExtractedSection:
    if client_name is None or not str(client_name).strip():
        raise ExtractionError(f"Client name is blank. Enter a client name to extract the {label}.")
    target = _norm_client(client_name)

    found: List[Tuple[Any, List[tuple], int, Dict[str, List[int]]]] = []
    best_overall = (