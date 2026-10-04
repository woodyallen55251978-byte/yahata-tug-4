"""元帳R7年度（2025/4/1～2026/3/31）の再計算と既存集計との照合。

ZIP・元資料は読み取りのみ。結果は標準出力へ表示する（Excelは作らない）。
使い方: python3 scripts/01_元帳照合.py   （事前に pip install openpyxl xlrd）
"""
import collections as C
import csv
import datetime as dt
import glob
import hashlib
import os
import unicodedata
import zipfile

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
INPUT = os.path.join(ROOT, "input")
ZIP = os.path.join(ROOT, "八幡4隻検証_資料一式.zip")

# 元帳「タグ」欄の記号。正式船名・会社への対応は未確認のため記号のまま扱う。
CODES_5 = {"T", "S", "Y", "M", "P"}          # 既存集計の「STC」と月別値で照合する記号群
CODES_KANCHO = {"関", "長"}                   # 既存集計の「西海」と月別値で照合する記号群
CODES_EXCLUDED_BY_EXISTING = {"まんじゅ", "せいじゅ"}  # 既存集計で除外されたと見られる記号（理由未確認）


def nfc(s):
    return unicodedata.normalize("NFC", s)


def extract_if_needed():
    if os.path.isdir(INPUT):
        return
    with zipfile.ZipFile(ZIP) as z:
        for info in z.infolist():
            name = info.filename
            if not info.flag_bits & 0x800:
                name = name.encode("cp437").decode("cp932")
            path = os.path.join(INPUT, name)
            if name.endswith("/"):
                os.makedirs(path, exist_ok=True)
                continue
            os.makedirs(os.path.dirname(path), exist_ok=True)
            with open(path, "wb") as f:
                f.write(z.read(info))


def verify_hashes():
    print("== SHA-256照合（00_資料一覧.csv）")
    with open(os.path.join(INPUT, "00_資料一覧.csv"), encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            cands = [p for p in glob.glob(os.path.join(INPUT, row["保存先"], "*"))
                     if nfc(os.path.basename(p)) == nfc(row["ファイル名"])]
            if not cands:
                print("  なし", row["ファイル名"])
                continue
            h = hashlib.sha256(open(cands[0], "rb").read()).hexdigest()
            print("  ", "一致" if h == row["SHA-256"] else "不一致", row["ファイル名"])


def fy_date(month, day):
    return dt.date(2025 if month >= 4 else 2026, month, day)


def minutes(t):
    return t.hour * 60 + t.minute


def load_ledger():
    path = glob.glob(os.path.join(INPUT, "01_元資料", "作業実績*"))[0]
    ws = openpyxl.load_workbook(path, read_only=True)["R7年度"]
    return [(i, r) for i, r in enumerate(ws.iter_rows(values_only=True), 1)
            if i > 1 and any(v is not None for v in r[:21])]


def is_cancel(r):
    return any(isinstance(v, str) and ("ｷｬﾝｾﾙ" in v or "キャンセル" in v) for v in r)


def main():
    extract_if_needed()
    verify_hashes()
    rows = load_ledger()
    tug_rows = [(i, r) for i, r in rows if r[17] is not None]

    print("\n== 件数（シート R7年度）")
    print("  データ行", len(rows), "／うち『作業なし』行", sum(1 for _, r in rows if r[3] == "作業なし"))
    print("  E列=1 の行（作業件数）", sum(1 for _, r in rows if r[4] == 1))
    print("  No欄記入行", sum(1 for _, r in rows if r[2] is not None))
    print("  タグ欄記入行（延べタグ・全記号）", len(tug_rows))
    by_code = C.Counter(r[17] for _, r in tug_rows)
    n5 = sum(by_code[c] for c in CODES_5)
    nk = sum(by_code[c] for c in CODES_KANCHO)
    nx = sum(by_code[c] for c in CODES_EXCLUDED_BY_EXISTING)
    print("  記号T,S,Y,M,P", n5, "／関・長", nk, "／まんじゅ・せいじゅ", nx,
          "／その他", len(tug_rows) - n5 - nk - nx)
    print("  タグ欄×S列", sorted(C.Counter((r[17], r[18]) for _, r in tug_rows).items(), key=lambda kv: -kv[1]))
    canc = [(i, r) for i, r in tug_rows if is_cancel(r)]
    print("  取消し記載のタグ行", len(canc), "／作業", len({r[3] + str(r[0]) + str(r[1]) for _, r in canc}))

    # 日別・記号群別の稼働隻数（元帳の月日＝作業日として集計）
    def grp(code):
        if code in CODES_5:
            return "STC"
        if code in CODES_KANCHO:
            return "西海"
        if code in CODES_EXCLUDED_BY_EXISTING:
            return None
        return "他社"
    days = C.defaultdict(lambda: C.defaultdict(set))
    for _, r in tug_rows:
        g = grp(r[17])
        if g:
            days[fy_date(r[0], r[1])][g].add(r[17])
    months = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]
    alld = [dt.date(2025, 4, 1) + dt.timedelta(n) for n in range(365)]
    print("\n== 月別 日別稼働隻数の合計")
    for g in ["STC", "西海", "他社"]:
        vals = [sum(len(days[d][g]) for d in alld if d.month == m) for m in months]
        print(" ", g, vals, sum(vals))
    print("  他社記号を使用した日数", sum(1 for d in alld if days[d]["他社"]))
    dist = C.Counter()
    for d in alld:
        n = sum(len(v) for v in days[d].values())
        dist["10以上" if n >= 10 else ("2以下" if n <= 2 else n)] += 1
    print("  稼働隻数別の日数", dict(dist))

    # 作業時間（L・M列）を使った同時稼働の候補（時刻の定義は未確認）
    iv, unusable = [], []
    job = None
    for i, r in rows:
        if r[4] == 1 or r[2] is not None:
            job = i
        if r[17] is None:
            continue
        s, e = r[11], r[12]
        if s is None or e is None or (isinstance(e, dt.datetime) and (e.month, e.day) != (1, 1)):
            unusable.append(i)
            continue
        em = minutes(e) + (1440 if isinstance(e, dt.datetime) or minutes(e) < minutes(s) else 0)
        base = dt.datetime.combine(fy_date(r[0], r[1]), dt.time())
        iv.append((base + dt.timedelta(minutes=minutes(s)), base + dt.timedelta(minutes=em), r[17], job, r))
    print("\n== 作業時間で区間化できたタグ行", len(iv), "／不可", len(unusable), unusable)
    for label, sel in [("取消し含む", lambda r: True), ("取消し除く", lambda r: not is_cancel(r)),
                       ("取消し・当直除く", lambda r: not is_cancel(r) and r[8] != "当直")]:
        ev = [(a, b, t, j) for a, b, t, j, r in iv if sel(r)]
        mx, multi = C.defaultdict(int), set()
        for p in sorted({a for a, *_ in ev}):
            act = [(t, j) for a, b, t, j in ev if a <= p < b]
            n = len({t for t, _ in act})
            mx[p.date()] = max(mx[p.date()], n)
            if n >= 3 and len({j for _, j in act}) >= 2:
                multi.add(p.date())
        mx = {d: v for d, v in mx.items() if dt.date(2025, 4, 1) <= d <= dt.date(2026, 3, 31)}
        print(" ", label, "同時3隻以上の日", sum(v >= 3 for v in mx.values()),
              "うち複数本船の重なり", len([d for d in multi if d in mx]),
              "／4隻以上", sum(v >= 4 for v in mx.values()), "／5隻以上", sum(v >= 5 for v in mx.values()))


if __name__ == "__main__":
    main()
