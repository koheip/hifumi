"""ヒフミのコア処理：Excel の読み込み・列の自動判定・集計・レポート生成。

すべてメモリ上（bytes）で完結し、ファイルをサーバーに保存しません。
Web（FastAPI）に依存しない純粋な関数だけで書いているので、単体テストがしやすい構成です。
"""
from __future__ import annotations

import io
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime, time
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---- 制限値（無料サーバーを守るための設計） --------------------------------
MAX_FILE_BYTES = 2 * 1024 * 1024   # アップロード上限 2MB
MAX_ROWS = 20_000                  # 読み込む最大行数（データ行）
MAX_COLS = 50                      # 読み込む最大列数
MAX_GROUPS_FOR_CHART = 15          # グラフに載せる最大グループ数
PREVIEW_ROWS = 10                  # 画面プレビューの行数
BLANK_LABEL = "(空欄)"

# 列の自動判定に使うキーワード（見つかった列を優先する）
VALUE_KEYWORDS = ("金額", "売上", "合計", "費用", "支出", "収入", "amount", "total", "sales", "price")
GROUP_KEYWORDS = ("カテゴリ", "分類", "店舗", "部門", "部署", "担当", "種別", "区分", "category", "group", "dept")


class AppError(Exception):
    """利用者に見せてよい（日本語の）エラーメッセージを持つ例外。"""

    def __init__(self, message: str, status_code: int = 400):
        super().__init__(message)
        self.message = message
        self.status_code = status_code


# ---- データ構造 ----------------------------------------------------------
@dataclass
class Table:
    headers: list[str]
    rows: list[list[Any]]          # 元の値（数値・文字列・日付など）
    sheet_name: str = ""


@dataclass
class GroupStat:
    name: str
    count: int
    total: float
    average: float
    share: float                   # 全体に対する構成比（0〜1）


@dataclass
class Summary:
    group_col: str
    value_col: str
    row_count: int                 # 集計対象になった行数
    skipped: int                   # 数値として読めず除外した行数
    total: float
    average: float
    minimum: float
    maximum: float
    groups: list[GroupStat] = field(default_factory=list)


# ---- 読み込み ------------------------------------------------------------
def validate_upload(filename: str | None, data: bytes) -> None:
    """ファイル名・サイズ・中身の形式をチェックする。問題があれば AppError。"""
    if not filename:
        raise AppError("ファイルが選択されていません。.xlsx ファイルを選んでください。")
    lower = filename.lower()
    if lower.endswith(".xls"):
        raise AppError("古い形式（.xls）には対応していません。Excel で「名前を付けて保存」から .xlsx 形式に変えてお試しください。")
    if not lower.endswith(".xlsx"):
        raise AppError("対応しているのは .xlsx ファイルのみです。拡張子をご確認ください。")
    if len(data) == 0:
        raise AppError("ファイルが空です。データの入った .xlsx ファイルを選んでください。")
    if len(data) > MAX_FILE_BYTES:
        raise AppError(f"ファイルサイズが大きすぎます（上限 {MAX_FILE_BYTES // 1024 // 1024}MB）。データを減らしてお試しください。", 413)
    # .xlsx は zip 形式。拡張子だけ偽装したファイルをここで弾く。
    if not zipfile.is_zipfile(io.BytesIO(data)):
        raise AppError("Excel ファイルとして読み込めませんでした。ファイルが壊れていないか、本当に .xlsx かご確認ください。")


def _cell_to_plain(value: Any) -> Any:
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d") if not isinstance(value, datetime) or value.time() == time(0) else value.strftime("%Y-%m-%d %H:%M")
    return value


def read_table(data: bytes) -> Table:
    """先頭シートを読み込み、1行目（空行は飛ばす）を見出しとして Table にする。"""
    try:
        wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    except Exception:
        raise AppError("Excel ファイルとして読み込めませんでした。ファイルが壊れていないかご確認ください。")
    try:
        ws = wb.worksheets[0]
        sheet_name = ws.title
        raw_rows = []
        for row in ws.iter_rows(values_only=True):
            row = list(row[:MAX_COLS])
            if any(v not in (None, "") for v in row):
                raw_rows.append(row)
            if len(raw_rows) > MAX_ROWS + 1:
                raise AppError(f"行数が多すぎます（上限 {MAX_ROWS:,} 行）。データを分けてお試しください。", 413)
    finally:
        wb.close()

    if len(raw_rows) < 2:
        raise AppError("データが見つかりません。1行目に見出し、2行目以降にデータが入っているかご確認ください。")

    header_row = raw_rows[0]
    width = max(i for i, v in enumerate(header_row) if v not in (None, "")) + 1
    headers: list[str] = []
    for i in range(width):
        name = str(header_row[i]).strip() if header_row[i] not in (None, "") else f"列{i + 1}"
        base, n = name, 2
        while name in headers:           # 見出しの重複は連番で区別する
            name = f"{base}_{n}"
            n += 1
        headers.append(name)

    rows = []
    for r in raw_rows[1:]:
        r = (r + [None] * width)[:width]
        rows.append([_cell_to_plain(v) for v in r])
    return Table(headers=headers, rows=rows, sheet_name=sheet_name)


# ---- 列の自動判定 --------------------------------------------------------
def to_number(value: Any) -> float | None:
    """数値として読める値なら float を返す。読めなければ None。"""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        s = value.strip().replace(",", "").replace("¥", "").replace("￥", "").replace("円", "").replace(" ", "")
        if not s:
            return None
        try:
            return float(s)
        except ValueError:
            return None
    return None


def numeric_columns(table: Table) -> list[str]:
    """空でない値の 80% 以上が数値として読める列を、数値列とみなす。"""
    result = []
    for i, h in enumerate(table.headers):
        values = [r[i] for r in table.rows if r[i] not in (None, "")]
        if values and sum(to_number(v) is not None for v in values) / len(values) >= 0.8:
            result.append(h)
    return result


def group_candidates(table: Table) -> list[str]:
    """グループ分けに向く列（数値以外で、種類が多すぎず少なすぎない列）。"""
    nums = set(numeric_columns(table))
    result = []
    for i, h in enumerate(table.headers):
        if h in nums:
            continue
        uniq = {str(r[i]) if r[i] not in (None, "") else BLANK_LABEL for r in table.rows}
        if 2 <= len(uniq) <= 50 and len(uniq) < len(table.rows):   # 全行が別の値（ID など）の列は除く
            result.append(h)
    return result


def _pick(columns: list[str], keywords: tuple[str, ...], prefer_last: bool) -> str | None:
    for h in columns:
        if any(k in h.lower() for k in keywords):
            return h
    if not columns:
        return None
    return columns[-1] if prefer_last else columns[0]


def detect_columns(table: Table) -> tuple[str | None, str | None]:
    """(グループ列, 数値列) を自動判定して返す。見つからなければ None。"""
    value_col = _pick(numeric_columns(table), VALUE_KEYWORDS, prefer_last=True)
    group_col = _pick(group_candidates(table), GROUP_KEYWORDS, prefer_last=False)
    return group_col, value_col


# ---- 集計 ----------------------------------------------------------------
def summarize(table: Table, group_col: str, value_col: str) -> Summary:
    if group_col not in table.headers:
        raise AppError(f"グループ列「{group_col}」が見つかりません。")
    if value_col not in table.headers:
        raise AppError(f"数値列「{value_col}」が見つかりません。")
    if group_col == value_col:
        raise AppError("グループ列と数値列には、別々の列を選んでください。")
    gi, vi = table.headers.index(group_col), table.headers.index(value_col)

    sums: dict[str, float] = {}
    counts: dict[str, int] = {}
    skipped = 0
    for r in table.rows:
        num = to_number(r[vi])
        if num is None:
            skipped += 1
            continue
        key = str(r[gi]).strip() if r[gi] not in (None, "") else BLANK_LABEL
        sums[key] = sums.get(key, 0.0) + num
        counts[key] = counts.get(key, 0) + 1

    n = sum(counts.values())
    if n == 0:
        raise AppError(f"「{value_col}」列に数値として読めるデータがありません。別の列を選んでください。")
    total = sum(sums.values())
    all_values = [v for r in table.rows if (v := to_number(r[vi])) is not None]
    groups = [
        GroupStat(name=k, count=counts[k], total=sums[k], average=sums[k] / counts[k],
                  share=(sums[k] / total) if total else 0.0)
        for k in sorted(sums, key=lambda k: sums[k], reverse=True)
    ]
    return Summary(group_col=group_col, value_col=value_col, row_count=n, skipped=skipped,
                   total=total, average=total / n, minimum=min(all_values), maximum=max(all_values),
                   groups=groups)


# ---- レポート（Excel）の生成 ---------------------------------------------
_HEADER_FILL = PatternFill("solid", fgColor="2F5D50")
_HEADER_FONT = Font(bold=True, color="FFFFFF")
_THIN = Side(style="thin", color="BBBBBB")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_NUM_FMT = "#,##0.##"


def _style_header(ws, row: int, first_col: int, last_col: int) -> None:
    for c in range(first_col, last_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill, cell.font, cell.border = _HEADER_FILL, _HEADER_FONT, _BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")


def _autofit(ws) -> None:
    """各列の最大文字数から列幅を決める（全角は2文字分として数える）。"""
    widths: dict[int, int] = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            text = f"{cell.value:,.2f}" if isinstance(cell.value, float) else str(cell.value)
            length = sum(2 if ord(ch) > 255 else 1 for ch in text)
            widths[cell.column] = max(widths.get(cell.column, 0), length)
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = min(max(w + 3, 10), 50)


def build_report(table: Table, summary: Summary) -> bytes:
    """集計サマリー・元データの2シートからなる整形済み Excel を bytes で返す。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "集計サマリー"

    ws["A1"] = "ヒフミ 集計レポート"
    ws["A1"].font = Font(bold=True, size=16, color="2F5D50")
    ws["A2"] = f"グループ列：{summary.group_col}　／　集計する数値列：{summary.value_col}"
    ws["A3"] = f"作成日時：{datetime.now().strftime('%Y-%m-%d %H:%M')}"

    # 全体の統計
    ws["A5"], ws["B5"] = "項目", "値"
    _style_header(ws, 5, 1, 2)
    overall = [("集計対象の行数", summary.row_count), ("合計", summary.total), ("平均", summary.average),
               ("最小", summary.minimum), ("最大", summary.maximum), ("グループ数", len(summary.groups))]
    for i, (label, val) in enumerate(overall, start=6):
        ws.cell(row=i, column=1, value=label).border = _BORDER
        c = ws.cell(row=i, column=2, value=val)
        c.border, c.number_format = _BORDER, _NUM_FMT

    # グループ別
    start = 6 + len(overall) + 2
    for j, h in enumerate([summary.group_col, "件数", "合計", "平均", "構成比"], start=1):
        ws.cell(row=start, column=j, value=h)
    _style_header(ws, start, 1, 5)
    for i, g in enumerate(summary.groups, start=start + 1):
        values = [g.name, g.count, g.total, g.average, g.share]
        for j, v in enumerate(values, start=1):
            cell = ws.cell(row=i, column=j, value=v)
            cell.border = _BORDER
            if isinstance(v, str):
                cell.data_type = "s"        # 数式として解釈させない
            elif j == 5:
                cell.number_format = "0.0%"
            else:
                cell.number_format = _NUM_FMT
    last = start + len(summary.groups)

    # 棒グラフ（合計の大きい上位だけ）
    n_chart = min(len(summary.groups), MAX_GROUPS_FOR_CHART)
    chart = BarChart()
    chart.type, chart.style = "col", 10
    chart.title = f"{summary.group_col}ごとの{summary.value_col} 合計"
    chart.y_axis.title, chart.x_axis.title = summary.value_col, summary.group_col
    chart.add_data(Reference(ws, min_col=3, min_row=start, max_row=start + n_chart), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=start + 1, max_row=start + n_chart))
    chart.legend = None
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.width, chart.height = 18, 9
    ws.add_chart(chart, "G5")
    _autofit(ws)
    ws.column_dimensions["A"].width = max(ws.column_dimensions["A"].width or 0, 24)

    # 元データ
    wd = wb.create_sheet("元データ")
    wd.append(table.headers)
    _style_header(wd, 1, 1, len(table.headers))
    for r in table.rows:
        wd.append(r)
    for row in wd.iter_rows(min_row=2):
        for cell in row:
            if isinstance(cell.value, str):
                cell.data_type = "s"        # 「=」始まりの文字列を数式として実行させない
            elif isinstance(cell.value, (int, float)):
                cell.number_format = _NUM_FMT
    wd.freeze_panes = "A2"
    wd.auto_filter.ref = wd.dimensions
    _autofit(wd)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()
