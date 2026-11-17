"""動作確認用のサンプル Excel（架空の店舗売上データ）をメモリ上で作る。"""
import io
import random
from datetime import date, timedelta

from openpyxl import Workbook

STORES = ["渋谷店", "新宿店", "横浜店", "大阪店"]
PRODUCTS = {  # カテゴリ: [(商品名, 単価), ...]
    "飲料": [("緑茶", 150), ("コーヒー", 220), ("炭酸水", 130)],
    "食品": [("おにぎり", 160), ("サンドイッチ", 320), ("カップ麺", 210)],
    "日用品": [("ティッシュ", 280), ("ハンドソープ", 450), ("乾電池", 380)],
    "文具": [("ボールペン", 120), ("ノート", 180), ("付箋", 240)],
}


def build_sample_xlsx(rows: int = 60, seed: int = 2026) -> bytes:
    rnd = random.Random(seed)          # 固定シードなので毎回同じデータになる
    wb = Workbook()
    ws = wb.active
    ws.title = "売上データ"
    ws.append(["日付", "店舗", "カテゴリ", "商品名", "数量", "単価", "売上金額"])
    start = date(2026, 4, 1)
    for _ in range(rows):
        category = rnd.choice(list(PRODUCTS))
        name, price = rnd.choice(PRODUCTS[category])
        qty = rnd.randint(1, 12)
        ws.append([start + timedelta(days=rnd.randint(0, 29)), rnd.choice(STORES),
                   category, name, qty, price, qty * price])
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()
