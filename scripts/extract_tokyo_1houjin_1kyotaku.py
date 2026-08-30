#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
厚労省「介護サービス情報公表システム」オープンデータ（居宅介護支援・全国版）から
東京都内に居宅介護支援事業所を1件だけ持つ法人を抽出し、Excelを生成する。

使い方:
    python3 scripts/extract_tokyo_1houjin_1kyotaku.py [CSVパス]

CSVパス省略時は既定URLからダウンロードを試みる。
"""

import os
import re
import subprocess
import sys
import unicodedata

import pandas as pd
from openpyxl.utils import get_column_letter

CSV_URL = (
    "https://www.mhlw.go.jp/content/12300000/"
    "jigyosho_430_all_20260709180740.csv"
)
DEFAULT_CSV = "jigyosho_430_all.csv"
OUT_XLSX = "tokyo_1houjin_1kyotaku.xlsx"

OUT_COLS = [
    "法人名", "法人所在地", "事業所名", "事業所番号",
    "事業所所在地", "電話番号", "FAX", "ホームページ", "指定年月日",
]


# --------------------------------------------------------------------------
# 1. 取得
# --------------------------------------------------------------------------
def ensure_csv(path):
    if os.path.exists(path):
        return path
    print(f"ダウンロード中: {CSV_URL}")
    subprocess.run(
        ["curl", "-sSL", "--fail", "-o", path, CSV_URL], check=True
    )
    return path


# --------------------------------------------------------------------------
# 2. 読み込み（cp932 → utf-8-sig の順に試す）
# --------------------------------------------------------------------------
def find_header_row(path, encoding, max_scan=10):
    """先頭数行を覗いて、実際の見出し行のインデックスを返す。"""
    with open(path, "r", encoding=encoding, errors="strict") as f:
        for i, line in enumerate(f):
            if i >= max_scan:
                break
            if "事業所" in line and line.count(",") >= 5:
                return i
    return 0


def load_csv(path):
    last_err = None
    for enc in ("cp932", "utf-8-sig"):
        try:
            header_row = find_header_row(path, enc)
            df = pd.read_csv(
                path,
                encoding=enc,
                skiprows=header_row,
                dtype=str,
                low_memory=False,
                on_bad_lines="skip",
            )
            print(f"読み込み成功: encoding={enc}, 見出し行={header_row}, "
                  f"{len(df):,}行 x {len(df.columns)}列")
            return df
        except (UnicodeDecodeError, UnicodeError) as e:
            last_err = e
            continue
    raise RuntimeError(f"cp932 / utf-8-sig いずれでも読めませんでした: {last_err}")


# --------------------------------------------------------------------------
# 列名の解決（オープンデータは年度で列名が揺れるため部分一致で探す）
# --------------------------------------------------------------------------
def pick(cols, patterns, exclude=()):
    """patterns を順に走査し、最初にマッチした列名を返す。無ければ None。"""
    for pat in patterns:
        for c in cols:
            name = str(c)
            if re.search(pat, name) and not any(
                re.search(x, name) for x in exclude
            ):
                return c
    return None


def pick_all(cols, pattern, exclude=()):
    """pattern にマッチする列を出現順に全部返す。"""
    return [
        c for c in cols
        if re.search(pattern, str(c))
        and not any(re.search(x, str(c)) for x in exclude)
    ]


def join_parts(df, cols):
    """住所が都道府県/市区町村/番地/建物名に分割されている場合に連結する。"""
    if not cols:
        return pd.Series([""] * len(df), index=df.index)
    s = df[cols[0]].fillna("")
    for c in cols[1:]:
        s = s.str.cat(df[c].fillna(""), sep="")
    return s.str.strip()


KANA_EXCL = (r"かな", r"カナ", r"ｶﾅ", r"フリガナ", r"ふりがな")


def resolve_columns(df):
    cols = list(df.columns)
    m = {}

    m["法人名"] = pick(
        cols, [r"法人.*名称", r"法人名", r"設置者.*名"], exclude=KANA_EXCL
    )
    m["事業所名"] = pick(
        cols, [r"事業所.*名称", r"事業所名"], exclude=KANA_EXCL
    )
    m["事業所番号"] = pick(cols, [r"事業所番号", r"事業所.*番号"])
    m["電話番号"] = pick(cols, [r"事業所.*電話", r"電話番号", r"電話"])
    m["FAX"] = pick(cols, [r"事業所.*FAX", r"FAX", r"ＦＡＸ", r"ファクシミリ"])
    m["ホームページ"] = pick(
        cols, [r"ホームページ", r"URL", r"ＵＲＬ", r"ウェブ", r"ＨＰ"]
    )
    m["指定年月日"] = pick(
        cols,
        [r"指定年月日", r"指定.*年月日", r"指定.*更新.*年月日"],
        exclude=(r"開始", r"休止", r"廃止", r"失効"),
    )

    # 住所は分割列を連結（無ければ単一列）
    houjin_addr = pick_all(
        cols, r"法人所在地", exclude=KANA_EXCL + (r"コード",)
    )
    if not houjin_addr:
        c = pick(cols, [r"法人.*所在地", r"法人.*住所"], exclude=KANA_EXCL)
        houjin_addr = [c] if c else []
    m["_法人所在地_列"] = houjin_addr

    jigyo_addr = pick_all(
        cols, r"事業所所在地", exclude=KANA_EXCL + (r"コード",)
    )
    if not jigyo_addr:
        c = pick(cols, [r"事業所.*所在地", r"事業所.*住所", r"所在地"],
                 exclude=KANA_EXCL + (r"法人", r"コード"))
        jigyo_addr = [c] if c else []
    m["_事業所所在地_列"] = jigyo_addr

    # 都道府県列（あれば絞り込みに使う）
    m["_都道府県"] = pick(
        cols,
        [r"事業所所在地.*都道府県", r"^都道府県$", r"都道府県名", r"都道府県"],
        exclude=(r"コード", r"法人"),
    )
    return m


# --------------------------------------------------------------------------
# 正規化
# --------------------------------------------------------------------------
_FULL2HALF = str.maketrans(
    {chr(c): chr(c - 0xFEE0) for c in range(0xFF01, 0xFF5F)}
)


def normalize(s):
    """全角英数記号を半角化し、空白（全角含む）をすべて除去する。"""
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    s = str(s)
    s = unicodedata.normalize("NFKC", s)
    s = s.translate(_FULL2HALF)
    s = re.sub(r"\s+", "", s)
    s = s.replace("　", "")
    return s


# --------------------------------------------------------------------------
# Excel 出力
# --------------------------------------------------------------------------
def display_width(v):
    """日本語を2、半角を1として概算幅を出す。"""
    w = 0
    for ch in str(v):
        w += 2 if unicodedata.east_asian_width(ch) in ("F", "W", "A") else 1
    return w


def write_excel(path, target_df, all_df):
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        target_df.to_excel(writer, sheet_name="対象リスト", index=False)
        all_df.to_excel(writer, sheet_name="全東京都", index=False)

        for sheet_name, df in (("対象リスト", target_df), ("全東京都", all_df)):
            ws = writer.sheets[sheet_name]
            ws.freeze_panes = "A2"  # 1行目固定
            for i, col in enumerate(df.columns, start=1):
                widths = [display_width(col)]
                # 全件走査は重いので先頭2000行で概算
                widths += [display_width(v) for v in df[col].head(2000)]
                ws.column_dimensions[get_column_letter(i)].width = min(
                    max(widths) + 2, 60
                )


# --------------------------------------------------------------------------
def main():
    csv_path = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_CSV
    csv_path = ensure_csv(csv_path)

    df = load_csv(csv_path)
    m = resolve_columns(df)

    missing = [k for k in ("法人名", "事業所名") if not m.get(k)]
    if missing:
        raise RuntimeError(
            f"必須列が見つかりません: {missing}\n列一覧: {list(df.columns)}"
        )

    out = pd.DataFrame(index=df.index)
    out["法人名"] = df[m["法人名"]].fillna("").astype(str).str.strip()
    out["法人所在地"] = join_parts(df, m["_法人所在地_列"])
    out["事業所名"] = df[m["事業所名"]].fillna("").astype(str).str.strip()
    for key in ("事業所番号", "電話番号", "FAX", "ホームページ", "指定年月日"):
        col = m.get(key)
        out[key] = (
            df[col].fillna("").astype(str).str.strip()
            if col else ""
        )
    out["事業所所在地"] = join_parts(df, m["_事業所所在地_列"])
    out = out[OUT_COLS]

    # ---- 3. 東京都に絞り込み -------------------------------------------
    if m["_都道府県"]:
        pref = df[m["_都道府県"]].fillna("").astype(str).str.strip()
        mask = pref.str.startswith("東京都") | (pref == "東京")
        how = f'都道府県列「{m["_都道府県"]}」'
    else:
        mask = out["事業所所在地"].str.startswith("東京都")
        how = "事業所所在地の先頭一致"
    tokyo = out[mask].reset_index(drop=True)
    print(f"東京都の絞り込み方法: {how}")

    # ---- 4. 法人名＋法人所在地でグルーピング ---------------------------
    key = (
        tokyo["法人名"].map(normalize)
        + "|"
        + tokyo["法人所在地"].map(normalize)
    )
    counts = key.map(key.value_counts())
    target = tokyo[counts == 1].reset_index(drop=True)

    # ---- 5. Excel 出力 --------------------------------------------------
    write_excel(OUT_XLSX, target, tokyo)

    # ---- 6. 報告 --------------------------------------------------------
    print()
    print(f"東京都の居宅介護支援 総事業所数: {len(tokyo):,} 件")
    print(f"1法人1事業所の件数            : {len(target):,} 件")
    print(f"出力: {OUT_XLSX}")


if __name__ == "__main__":
    main()
