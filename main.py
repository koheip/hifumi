"""ヒフミ（Excel 集計レポートメーカー）— FastAPI アプリ本体。

起動:  uvicorn main:app --reload
方針:  アップロードされたファイルはメモリ上だけで処理し、サーバーには保存しない。
"""
import base64
import binascii
from pathlib import Path
from urllib.parse import quote

from fastapi import FastAPI, File, Form, Request, UploadFile
from fastapi.responses import HTMLResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

import excel_tools as xt
from sample_data import build_sample_xlsx

BASE_DIR = Path(__file__).resolve().parent
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

app = FastAPI(title="ヒフミ", description="Excel 集計レポートメーカー")
app.mount("/static", StaticFiles(directory=BASE_DIR / "static"), name="static")
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))


# ---- 共通処理 ------------------------------------------------------------
@app.exception_handler(xt.AppError)
async def app_error_handler(request: Request, exc: xt.AppError):
    """想定内のエラーは、日本語メッセージ付きの画面で返す。"""
    return templates.TemplateResponse(request, "error.html", {"message": exc.message}, status_code=exc.status_code)


async def get_file_bytes(upload: UploadFile | None, data_b64: str) -> tuple[str, bytes]:
    """新しくアップロードされたファイル、または結果画面から引き継いだデータ（base64）を取り出す。

    サーバー側に何も保存しない代わりに、結果画面の中にデータを持たせて再利用する（ステートレス設計）。
    """
    if upload is not None and upload.filename:
        data = await upload.read(xt.MAX_FILE_BYTES + 1)   # 上限+1バイトだけ読めば超過を判定できる
        filename = upload.filename
    elif data_b64:
        if len(data_b64) > xt.MAX_FILE_BYTES * 2:
            raise xt.AppError("ファイルサイズが大きすぎます。もう一度アップロードしてください。", 413)
        try:
            data = base64.b64decode(data_b64, validate=True)
        except (binascii.Error, ValueError):
            raise xt.AppError("データを引き継げませんでした。もう一度ファイルをアップロードしてください。")
        filename = "carried.xlsx"
    else:
        raise xt.AppError("ファイルが選択されていません。.xlsx ファイルを選んでください。")
    xt.validate_upload(filename, data)
    return filename, data


def resolve_columns(table: xt.Table, group_col: str, value_col: str) -> tuple[str, str]:
    """空欄なら自動判定の結果を使う。判定できなければ分かりやすいエラーにする。"""
    auto_group, auto_value = xt.detect_columns(table)
    group_col, value_col = group_col or auto_group, value_col or auto_value
    if not value_col:
        raise xt.AppError("集計できる数値の列が見つかりませんでした。数値（金額・数量など）の入った列があるかご確認ください。")
    if not group_col:
        raise xt.AppError("グループ分けに使える列が見つかりませんでした。店舗・カテゴリなど、同じ値が繰り返される列があるかご確認ください。")
    return group_col, value_col


# ---- 画面・API -----------------------------------------------------------
@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse(request, "index.html", {"max_mb": xt.MAX_FILE_BYTES // 1024 // 1024})


@app.post("/analyze", response_class=HTMLResponse)
async def analyze(request: Request, file: UploadFile | None = File(None),
                  data_b64: str = Form(""), group_col: str = Form(""), value_col: str = Form("")):
    """ファイルを読み込み、列を判定して集計し、プレビュー付きの結果画面を返す。"""
    filename, data = await get_file_bytes(file, data_b64)
    table = xt.read_table(data)
    group_col, value_col = resolve_columns(table, group_col, value_col)
    summary = xt.summarize(table, group_col, value_col)
    context = {
        "filename": file.filename if file is not None and file.filename else "（前回アップロードしたファイル）",
        "data_b64": base64.b64encode(data).decode("ascii"),
        "table": table, "summary": summary,
        "headers": table.headers,
        "numeric_cols": xt.numeric_columns(table),
        "preview": table.rows[: xt.PREVIEW_ROWS],
        "total_rows": len(table.rows),
        "preview_rows": xt.PREVIEW_ROWS,
    }
    return templates.TemplateResponse(request, "result.html", context)


@app.post("/report")
async def report(file: UploadFile | None = File(None), data_b64: str = Form(""),
                 group_col: str = Form(""), value_col: str = Form("")):
    """整形済みの Excel レポート（集計サマリー＋元データ）をダウンロードさせる。"""
    _, data = await get_file_bytes(file, data_b64)
    table = xt.read_table(data)
    group_col, value_col = resolve_columns(table, group_col, value_col)
    summary = xt.summarize(table, group_col, value_col)
    content = xt.build_report(table, summary)
    filename = quote("ヒフミ集計レポート.xlsx")
    return Response(content=content, media_type=XLSX_MIME,
                    headers={"Content-Disposition": f"attachment; filename*=UTF-8''{filename}"})


@app.get("/sample.xlsx")
async def sample():
    """動作確認用のサンプル Excel（架空の売上データ）。"""
    filename = quote("ヒフミ_サンプル売上データ.xlsx")
    return Response(content=build_sample_xlsx(), media_type=XLSX_MIME,
                    headers={"Content-Disposition": f"attachment; filename*=UTF-8''{filename}"})


@app.get("/health")
async def health():
    return {"status": "ok"}
