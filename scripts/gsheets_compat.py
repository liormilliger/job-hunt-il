# -*- coding: utf-8 -*-
"""openpyxl-compatible facade backed by Google Sheets.

Every script in this pipeline used to talk to a local .xlsx file through
openpyxl. To move the tracker to Google Sheets (so it works without Microsoft
Excel, on Mac/Linux/Docker, and is viewable/editable from a browser or phone)
without rewriting every script's cell-by-cell logic, this module re-implements
the small slice of openpyxl's API those scripts actually use, on top of the
Sheets API. Callers do:

    import gsheets_compat as openpyxl
    from gsheets_compat import Font, PatternFill, Alignment, Border, Side

...and everything else (ws.cell(), ws.append(), ws.iter_rows(), cell.font =
Font(...), ws.column_dimensions[...].width = ..., wb.save(), etc.) behaves the
same as it did against openpyxl, just against a live Google Sheet instead of a
local file.

Design (deliberately simple over clever, since this can't be exercised against
live Google credentials in every environment that edits it):
  - Each Worksheet holds its ENTIRE grid of values in memory (loaded in one
    shot in load_workbook(), or started empty for a brand-new sheet).
    ws.cell(), ws.append(), ws[...], ws.delete_rows() all mutate that
    in-memory grid directly — no network calls.
  - Cell-level formatting (font/fill/alignment/border) and column/row
    dimensions are recorded as pending requests, not applied immediately.
  - wb.save() is the one point that talks to the network: for every sheet
    that changed, it clears the sheet's data range and rewrites the full grid
    in one values.update call (valueInputOption=USER_ENTERED, so formula
    strings like =HYPERLINK(...) are interpreted, exactly as in Excel), then
    flushes every pending formatting/dimension request in one batchUpdate.
  - Creating/deleting whole SHEETS (create_sheet / del wb["name"]) happens
    immediately against the API, since later formatting calls need a real
    sheetId to address.

Known gaps vs. openpyxl (none of these are exercised by this pipeline's
scripts today — flagged here so a future change doesn't assume otherwise):
  - No cell merging, conditional formatting, charts, or number formats beyond
    what Sheets infers from the value.
  - Column width / row height are converted from Excel's units to Sheets
    pixels with a standard approximation, not an exact match.
  - No support for opening an .xlsx file directly — "tracker" in config.json
    is a Google Sheet ID, not a filename, once setup.py has run.
"""
import os
import re

import jh_config
import google_auth

_CFG = jh_config.load()

_COL_RE = re.compile(r"^([A-Za-z]+)(\d+)$")


# ── column letter <-> index (same algorithm as openpyxl.utils) ─────────────
def _col_to_idx(letters: str) -> int:
    idx = 0
    for ch in letters.upper():
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx


def get_column_letter(idx: int) -> str:
    letters = ""
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(ord("A") + rem) + letters
    return letters


class _Utils:
    get_column_letter = staticmethod(get_column_letter)


utils = _Utils()


def _hex_to_rgb(hexval):
    if not hexval:
        return None
    h = hexval.lstrip("#")
    if len(h) == 8:  # openpyxl sometimes includes an alpha prefix (AARRGGBB)
        h = h[2:]
    if len(h) != 6:
        return None
    r, g, b = (int(h[i:i + 2], 16) / 255.0 for i in (0, 2, 4))
    return {"red": r, "green": g, "blue": b}


# ── style value objects (mirroring the openpyxl.styles names/shape) ────────
class Font:
    def __init__(self, name=None, bold=False, size=None, color=None):
        self.name, self.bold, self.size, self.color = name, bold, size, color


class PatternFill:
    def __init__(self, fill_type="solid", fgColor=None, start_color=None):
        self.fgColor = fgColor or start_color


class Alignment:
    def __init__(self, horizontal=None, vertical=None, wrap_text=False):
        self.horizontal, self.vertical, self.wrap_text = horizontal, vertical, wrap_text


class Side:
    def __init__(self, style=None, color=None):
        self.style, self.color = style, color


class Border:
    def __init__(self, left=None, right=None, top=None, bottom=None):
        self.left, self.right, self.top, self.bottom = left, right, top, bottom


def _format_to_cell_format(font=None, fill=None, alignment=None, border=None):
    fmt = {}
    if font:
        tf = {}
        if font.bold:
            tf["bold"] = True
        if font.name:
            tf["fontFamily"] = font.name
        if font.size:
            tf["fontSize"] = int(font.size)
        rgb = _hex_to_rgb(font.color)
        if rgb:
            tf["foregroundColor"] = rgb
        if tf:
            fmt["textFormat"] = tf
    if fill and fill.fgColor:
        rgb = _hex_to_rgb(fill.fgColor)
        if rgb:
            fmt["backgroundColor"] = rgb
    if alignment:
        if alignment.horizontal:
            fmt["horizontalAlignment"] = alignment.horizontal.upper()
        if alignment.vertical:
            v = alignment.vertical.upper()
            fmt["verticalAlignment"] = "MIDDLE" if v == "CENTER" else v
        fmt["wrapStrategy"] = "WRAP" if alignment.wrap_text else "OVERFLOW_CELL"
    if border:
        b = {}
        for side_name in ("left", "right", "top", "bottom"):
            side = getattr(border, side_name)
            if side and side.style:
                b[side_name] = {"style": "SOLID", "color": _hex_to_rgb(side.color) or
                                 {"red": 0, "green": 0, "blue": 0}}
        if b:
            fmt["borders"] = b
    return fmt


def _fields_for(fmt: dict) -> str:
    fields = []
    if "textFormat" in fmt:
        fields.append("userEnteredFormat.textFormat")
    if "backgroundColor" in fmt:
        fields.append("userEnteredFormat.backgroundColor")
    if "horizontalAlignment" in fmt:
        fields.append("userEnteredFormat.horizontalAlignment")
    if "verticalAlignment" in fmt:
        fields.append("userEnteredFormat.verticalAlignment")
    if "wrapStrategy" in fmt:
        fields.append("userEnteredFormat.wrapStrategy")
    if "borders" in fmt:
        fields.append("userEnteredFormat.borders")
    return ",".join(fields) or "userEnteredFormat"


# ── Cell: thin handle onto (worksheet, row, col) ────────────────────────────
class Cell:
    __slots__ = ("_ws", "row", "column")

    def __init__(self, ws, row, column):
        self._ws, self.row, self.column = ws, row, column

    @property
    def value(self):
        return self._ws._get(self.row, self.column)

    @value.setter
    def value(self, v):
        self._ws._set(self.row, self.column, v)

    @property
    def font(self):
        return None

    @font.setter
    def font(self, f):
        self._ws._add_format(self.row, self.column, font=f)

    @property
    def fill(self):
        return None

    @fill.setter
    def fill(self, f):
        self._ws._add_format(self.row, self.column, fill=f)

    @property
    def alignment(self):
        return None

    @alignment.setter
    def alignment(self, a):
        self._ws._add_format(self.row, self.column, alignment=a)

    @property
    def border(self):
        return None

    @border.setter
    def border(self, b):
        self._ws._add_format(self.row, self.column, border=b)


class _ColumnDims(dict):
    """ws.column_dimensions['A'].width = 20"""
    def __getitem__(self, key):
        if key not in self:
            dict.__setitem__(self, key, _Dim())
        return dict.__getitem__(self, key)


class _RowDims(dict):
    """ws.row_dimensions[1].height = 28"""
    def __getitem__(self, key):
        if key not in self:
            dict.__setitem__(self, key, _Dim())
        return dict.__getitem__(self, key)


class _Dim:
    def __init__(self):
        self.width = None
        self.height = None


class Worksheet:
    def __init__(self, workbook, title, sheet_id, grid=None):
        self.parent = workbook
        self.title = title
        self.sheet_id = sheet_id
        self._grid = grid if grid is not None else []   # list[list[value]]
        self._pending_formats = {}    # (row, col) -> merged format dict
        self.column_dimensions = _ColumnDims()
        self.row_dimensions = _RowDims()
        self._dirty = grid is None    # brand-new sheets always need a save

    # -- grid access -----------------------------------------------------
    def _ensure(self, row, col):
        while len(self._grid) < row:
            self._grid.append([])
        r = self._grid[row - 1]
        while len(r) < col:
            r.append(None)

    def _get(self, row, col):
        if row - 1 < len(self._grid) and col - 1 < len(self._grid[row - 1]):
            return self._grid[row - 1][col - 1]
        return None

    def _set(self, row, col, value):
        self._ensure(row, col)
        self._grid[row - 1][col - 1] = value
        self._dirty = True

    def _add_format(self, row, col, font=None, fill=None, alignment=None, border=None):
        fmt = _format_to_cell_format(font=font, fill=fill, alignment=alignment, border=border)
        if not fmt:
            return
        existing = self._pending_formats.get((row, col), {})
        existing.update(fmt)
        self._pending_formats[(row, col)] = existing
        self._dirty = True

    # -- openpyxl-shaped API ----------------------------------------------
    @property
    def max_row(self):
        return len(self._grid)

    @property
    def max_column(self):
        return max((len(r) for r in self._grid), default=0)

    def cell(self, row, column, value=None):
        if value is not None:
            self._set(row, column, value)
        else:
            self._ensure(row, column)
        return Cell(self, row, column)

    def append(self, values):
        row_idx = len(self._grid) + 1
        self._grid.append(list(values))
        self._dirty = True
        return [Cell(self, row_idx, c) for c in range(1, len(values) + 1)]

    def __getitem__(self, key):
        if isinstance(key, int):
            self._ensure(key, self.max_column or 1)
            width = len(self._grid[key - 1])
            return tuple(Cell(self, key, c) for c in range(1, width + 1))
        m = _COL_RE.match(str(key))
        if not m:
            raise KeyError(key)
        col = _col_to_idx(m.group(1))
        row = int(m.group(2))
        self._ensure(row, col)
        return Cell(self, row, col)

    def iter_rows(self, min_row=1, max_row=None, values_only=False):
        max_row = max_row or self.max_row
        for r in range(min_row, max_row + 1):
            width = self.max_column
            if values_only:
                yield tuple(self._get(r, c) for c in range(1, width + 1))
            else:
                yield tuple(Cell(self, r, c) for c in range(1, width + 1))

    def delete_rows(self, idx, amount=1):
        del self._grid[idx - 1: idx - 1 + amount]
        self._dirty = True


class Workbook:
    def __init__(self, spreadsheet_id=None, title="job-hunt-il tracker"):
        self._cfg = _CFG
        self._service = google_auth.get_sheets_service(self._cfg)
        self._sheets = {}     # title -> Worksheet
        self._order = []      # sheet titles, in order
        if spreadsheet_id:
            self.spreadsheet_id = spreadsheet_id
            self._load_all()
        else:
            body = {
                "properties": {"title": title},
                "sheets": [{"properties": {"title": "Sheet", "index": 0}}],
            }
            created = self._service.spreadsheets().create(
                body=body, fields="spreadsheetId,sheets(properties(sheetId,title,index))"
            ).execute()
            self.spreadsheet_id = created["spreadsheetId"]
            for sp in created["sheets"]:
                p = sp["properties"]
                ws = Worksheet(self, p["title"], p["sheetId"], grid=[])
                self._sheets[p["title"]] = ws
                self._order.append(p["title"])

    # -- loading -----------------------------------------------------------
    def _load_all(self):
        meta = self._service.spreadsheets().get(
            spreadsheetId=self.spreadsheet_id,
            fields="sheets(properties(sheetId,title,index))",
        ).execute()
        props = sorted(meta["sheets"], key=lambda s: s["properties"]["index"])
        titles = [p["properties"]["title"] for p in props]
        ranges = [f"'{t}'!A1:ZZ20000" for t in titles]
        vr = self._service.spreadsheets().values().batchGet(
            spreadsheetId=self.spreadsheet_id, ranges=ranges,
        ).execute()
        for p, value_range in zip(props, vr.get("valueRanges", [])):
            title = p["properties"]["title"]
            grid = value_range.get("values", [])
            ws = Worksheet(self, title, p["properties"]["sheetId"], grid=grid)
            ws._dirty = False
            self._sheets[title] = ws
            self._order.append(title)

    # -- openpyxl-shaped API -------------------------------------------------
    @property
    def sheetnames(self):
        return list(self._order)

    @property
    def active(self):
        return self._sheets[self._order[0]]

    def __getitem__(self, name):
        return self._sheets[name]

    def __delitem__(self, name):
        ws = self._sheets[name]
        self._service.spreadsheets().batchUpdate(
            spreadsheetId=self.spreadsheet_id,
            body={"requests": [{"deleteSheet": {"sheetId": ws.sheet_id}}]},
        ).execute()
        del self._sheets[name]
        self._order.remove(name)

    def create_sheet(self, title):
        resp = self._service.spreadsheets().batchUpdate(
            spreadsheetId=self.spreadsheet_id,
            body={"requests": [{"addSheet": {"properties": {"title": title}}}]},
        ).execute()
        sheet_id = resp["replies"][0]["addSheet"]["properties"]["sheetId"]
        ws = Worksheet(self, title, sheet_id, grid=[])
        self._sheets[title] = ws
        self._order.append(title)
        return ws

    def save(self, path=None):
        """path is accepted (and ignored) for call-site compatibility with
        openpyxl — the workbook already knows its own spreadsheet ID."""
        requests = []
        for title, ws in self._sheets.items():
            if not ws._dirty and not ws._pending_formats and not ws.column_dimensions \
                    and not ws.row_dimensions:
                continue

            # 1. values: clear the whole sheet, then rewrite the current grid
            #    in one shot. Simplest correct thing given writes here are
            #    infrequent (a handful of runs a day) — not a hot path.
            if ws._dirty:
                self._service.spreadsheets().values().clear(
                    spreadsheetId=self.spreadsheet_id, range=f"'{title}'!A1:ZZ20000",
                    body={},
                ).execute()
                if ws._grid:
                    last_col = max((len(r) for r in ws._grid), default=0)
                    end_col = get_column_letter(max(last_col, 1))
                    rng = f"'{title}'!A1:{end_col}{len(ws._grid)}"
                    self._service.spreadsheets().values().update(
                        spreadsheetId=self.spreadsheet_id, range=rng,
                        valueInputOption="USER_ENTERED",
                        body={"values": [[("" if v is None else v) for v in row]
                                          for row in ws._grid]},
                    ).execute()
                ws._dirty = False

            # 2. cell formatting
            for (row, col), fmt in ws._pending_formats.items():
                requests.append({
                    "repeatCell": {
                        "range": {
                            "sheetId": ws.sheet_id,
                            "startRowIndex": row - 1, "endRowIndex": row,
                            "startColumnIndex": col - 1, "endColumnIndex": col,
                        },
                        "cell": {"userEnteredFormat": fmt},
                        "fields": _fields_for(fmt),
                    }
                })
            ws._pending_formats.clear()

            # 3. column widths (approximate: openpyxl "characters" -> pixels)
            for letter, dim in ws.column_dimensions.items():
                if dim.width is None:
                    continue
                idx = _col_to_idx(letter)
                requests.append({
                    "updateDimensionProperties": {
                        "range": {"sheetId": ws.sheet_id, "dimension": "COLUMNS",
                                  "startIndex": idx - 1, "endIndex": idx},
                        "properties": {"pixelSize": int(dim.width * 7 + 5)},
                        "fields": "pixelSize",
                    }
                })
            ws.column_dimensions.clear()

            # 4. row heights (approximate: openpyxl "points" -> pixels)
            for row_num, dim in ws.row_dimensions.items():
                if dim.height is None:
                    continue
                requests.append({
                    "updateDimensionProperties": {
                        "range": {"sheetId": ws.sheet_id, "dimension": "ROWS",
                                  "startIndex": row_num - 1, "endIndex": row_num},
                        "properties": {"pixelSize": int(dim.height * 4 / 3)},
                        "fields": "pixelSize",
                    }
                })
            ws.row_dimensions.clear()

        if requests:
            # Sheets caps batchUpdate size in practice; chunk defensively.
            for i in range(0, len(requests), 400):
                self._service.spreadsheets().batchUpdate(
                    spreadsheetId=self.spreadsheet_id,
                    body={"requests": requests[i:i + 400]},
                ).execute()


def load_workbook(spreadsheet_id, read_only=False):
    """read_only is accepted for call-site compatibility (brief.py) and
    ignored — Sheets reads are always a fresh fetch, there's no local lock
    concept to bypass."""
    return Workbook(spreadsheet_id=spreadsheet_id)
