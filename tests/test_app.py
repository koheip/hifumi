import base64
import io

from fastapi.testclient import TestClient
from openpyxl import Workbook, load_workbook

import excel_tools as xt
from main import app
from sample_data import build_sample_xlsx

client = TestClient(app)
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def upload(data, name="test.xlsx", url="/analyze", **form):
    return client.post(url, files={"file": (name, data, XLSX)}, data=form)


def test_index_and_health():
    r = client.get("/")
    assert r.status_code == 200 and "ヒフミ" in r.text and "sample.xlsx" in r.text
    assert client.get("/health").json() == {"status": "ok"}


def test_sample_download_is_valid_xlsx():
    r = client.get("/sample.xlsx")
    assert r.status_code == 200 and r.headers["content-type"] == XLSX
    assert load_workbook(io.BytesIO(r.content)).sheetnames == ["売上データ"]


def test_analyze_shows_summary_and_preview():
    r = upload(build_sample_xlsx())
    assert r.status_code == 200
    assert "集計結果" in r.text and "渋谷店" in r.text and "売上金額" in r.text
    assert 'name="data_b64"' in r.text


def test_analyze_with_chosen_columns():
    r = upload(build_sample_xlsx(), group_col="カテゴリ", value_col="数量")
    assert r.status_code == 200 and "カテゴリごとの数量" in r.text


def test_round_trip_upload_then_download_report():
    r = upload(build_sample_xlsx(), url="/report")
    assert r.status_code == 200 and r.headers["content-type"] == XLSX
    assert "attachment" in r.headers["content-disposition"]
    wb = load_workbook(io.BytesIO(r.content))
    assert wb.sheetnames == ["集計サマリー", "元データ"]
    assert wb["元データ"].max_row == 61


def test_report_from_carried_data_without_reupload():
    """結果画面が持っているデータ（base64）だけでレポートを作れる。"""
    sample = build_sample_xlsx()
    r = client.post("/report", data={"data_b64": base64.b64encode(sample).decode(), "group_col": "店舗", "value_col": "売上金額"})
    assert r.status_code == 200
    assert load_workbook(io.BytesIO(r.content)).sheetnames == ["集計サマリー", "元データ"]


def test_friendly_errors_are_in_japanese():
    cases = [
        (upload(b"a,b\n1,2", name="data.csv"), 400, ".xlsx"),
        (upload(b"not excel", name="fake.xlsx"), 400, "読み込めませんでした"),
        (upload(b"0" * (xt.MAX_FILE_BYTES + 1)), 413, "大きすぎます"),
        (client.post("/analyze", data={}), 400, "選択されていません"),
        (client.post("/analyze", data={"data_b64": "###"}), 400, "引き継げませんでした"),
    ]
    for response, status, word in cases:
        assert response.status_code == status
        assert "うまく処理できませんでした" in response.text and word in response.text


def test_no_numeric_column_message():
    wb = Workbook()
    wb.active.append(["名前", "部署"])
    for n, d in [("a", "x"), ("b", "y"), ("c", "x"), ("d", "y")]:
        wb.active.append([n, d])
    buf = io.BytesIO()
    wb.save(buf)
    r = upload(buf.getvalue())
    assert r.status_code == 400 and "数値の列" in r.text


def test_html_is_escaped():
    wb = Workbook()
    wb.active.append(["店舗", "金額"])
    for s, v in [("<script>alert(1)</script>", 10), ("<script>alert(1)</script>", 5), ("B", 1), ("B", 2)]:
        wb.active.append([s, v])
    buf = io.BytesIO()
    wb.save(buf)
    r = upload(buf.getvalue())
    assert r.status_code == 200 and "<script>alert(1)</script>" not in r.text and "&lt;script&gt;" in r.text
