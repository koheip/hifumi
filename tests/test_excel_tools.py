import io

import pytest
from openpyxl import Workbook, load_workbook

import excel_tools as xt
from sample_data import build_sample_xlsx


def make_xlsx(rows):
    wb = Workbook()
    ws = wb.active
    for r in rows:
        ws.append(r)
    for row in ws.iter_rows():          # 「=」始まりでも、数式ではなく文字列として保存する
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                c.data_type = "s"
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


DATA = [
    ["店舗", "商品", "売上金額"],
    ["渋谷", "A", 100],
    ["渋谷", "B", 300],
    ["新宿", "A", 200],
    ["新宿", "B", 400],
    ["大阪", "A", 1000],
]


def test_read_table_and_detect_columns():
    table = xt.read_table(make_xlsx(DATA))
    assert table.headers == ["店舗", "商品", "売上金額"]
    assert len(table.rows) == 5
    assert xt.numeric_columns(table) == ["売上金額"]
    assert xt.detect_columns(table) == ("店舗", "売上金額")


def test_summarize_totals_and_groups():
    table = xt.read_table(make_xlsx(DATA))
    s = xt.summarize(table, "店舗", "売上金額")
    assert s.total == 2000 and s.row_count == 5 and s.average == 400
    assert (s.minimum, s.maximum) == (100, 1000)
    # 合計の大きい順：大阪1000, 新宿600, 渋谷400
    assert [(g.name, g.total, g.count) for g in s.groups] == [("大阪", 1000, 1), ("新宿", 600, 2), ("渋谷", 400, 2)]
    assert s.groups[0].share == pytest.approx(0.5)
    assert sum(g.share for g in s.groups) == pytest.approx(1.0)


def test_numbers_written_as_text_and_blanks():
    rows = [["部門", "費用"], ["営業", "1,200"], ["営業", "¥800円"], [None, 500], ["総務", "不明"]]
    s = xt.summarize(xt.read_table(make_xlsx(rows)), "部門", "費用")
    assert s.total == 2500 and s.skipped == 1
    assert {g.name for g in s.groups} == {"営業", xt.BLANK_LABEL}


def test_summarize_errors():
    table = xt.read_table(make_xlsx(DATA))
    with pytest.raises(xt.AppError):
        xt.summarize(table, "存在しない列", "売上金額")
    with pytest.raises(xt.AppError):
        xt.summarize(table, "売上金額", "売上金額")
    with pytest.raises(xt.AppError):
        xt.summarize(table, "店舗", "商品")           # 数値として読めない列


@pytest.mark.parametrize("name,data,status", [
    (None, b"x", 400),
    ("data.csv", b"a,b", 400),
    ("old.xls", b"x", 400),
    ("empty.xlsx", b"", 400),
    ("fake.xlsx", b"this is not excel", 400),
    ("big.xlsx", b"0" * (xt.MAX_FILE_BYTES + 1), 413),
])
def test_validate_upload_rejects(name, data, status):
    with pytest.raises(xt.AppError) as e:
        xt.validate_upload(name, data)
    assert e.value.status_code == status


def test_read_table_without_data():
    with pytest.raises(xt.AppError):
        xt.read_table(make_xlsx([["見出しだけ"]]))


def test_build_report_structure():
    data = make_xlsx(DATA)
    table = xt.read_table(data)
    summary = xt.summarize(table, "店舗", "売上金額")
    wb = load_workbook(io.BytesIO(xt.build_report(table, summary)))
    assert wb.sheetnames == ["集計サマリー", "元データ"]
    ws, wd = wb["集計サマリー"], wb["元データ"]
    assert len(ws._charts) == 1                      # 棒グラフが入っている
    assert wd.max_row == 6 and wd["A1"].value == "店舗"
    assert wd["A1"].font.bold and wd.freeze_panes == "A2"
    assert wd.column_dimensions["A"].width >= 10     # 列幅が設定されている
    texts = [c.value for row in ws.iter_rows() for c in row]
    assert 2000 in texts and "大阪" in texts


def test_formula_injection_is_neutralized():
    rows = [["名前", "金額"], ["=1+1", 10], ["@SUM(A1)", 20], ["通常", 30]]
    table = xt.read_table(make_xlsx(rows))
    s = xt.summarize(table, "名前", "金額")
    wb = load_workbook(io.BytesIO(xt.build_report(table, s)))
    cell = wb["元データ"]["A2"]
    assert cell.data_type == "s" and cell.value == "=1+1"


def test_sample_file_is_usable():
    table = xt.read_table(build_sample_xlsx())
    assert len(table.rows) == 60
    group, value = xt.detect_columns(table)
    assert (group, value) == ("店舗", "売上金額")
    s = xt.summarize(table, group, value)
    assert s.total == sum(r[6] for r in table.rows)
