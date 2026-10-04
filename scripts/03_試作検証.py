"""八幡地区_検証試作_v01.xlsx の動作確認（7項目＋基礎確認）。

- 納品用ファイル（実データ 2025/12/6）は変更しない。
- 試験専用コピーは、作成プログラム（02_試作作成.py の build）に架空の実績行を渡して output/試験専用/ に作る。
  架空の入力（前日確定・変更・マスター追加など）もそのコピーにだけ入れる。
- 再計算は LibreOffice（recalc.py）。Excel実機での確認は行っていない（未実施）。
- 結果は output/試験専用/試験結果_v01.json と標準出力へ書き出す（試作確認結果_v01.md の表の元）。

使い方: RECALC=/path/to/recalc.py python3 scripts/03_試作検証.py
"""
import datetime as dt
import glob
import importlib.util
import json
import os
import re
import shutil
import subprocess
import zipfile

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "output", "八幡地区_検証試作_v01.xlsx")
TDIR = os.path.join(ROOT, "output", "試験専用")
RECALC = os.environ.get("RECALC") or glob.glob("/root/.claude/skills/**/xlsx/scripts/recalc.py", recursive=True)[0]
spec = importlib.util.spec_from_file_location("gen", os.path.join(ROOT, "scripts", "02_試作作成.py"))
gen = importlib.util.module_from_spec(spec)
spec.loader.exec_module(gen)
results = []
D = dt.date
T = dt.time


def recalc(path):
    out = subprocess.run(["python3", RECALC, path, "300"], capture_output=True, text=True)
    info = json.loads(out.stdout)
    assert info.get("status") in ("success", "errors_found"), info
    return info


def check(test, no, name, inputs, expected, actual, method="LibreOffice再計算"):
    ok = expected == actual
    results.append({"test": test, "no": no, "name": name, "inputs": inputs, "expected": expected, "actual": actual,
                    "result": "合格" if ok else "不合格", "method": method})
    print(("OK " if ok else "NG ") + no, name, "期待", expected, "実際", actual)


def hhmm(v):
    if isinstance(v, dt.datetime):
        return v.strftime("%m/%d %H:%M") if v.date() != D(1899, 12, 30) else v.strftime("%H:%M")
    return v.strftime("%H:%M") if isinstance(v, dt.time) else v


def trow(d, ship, io, hin, start, end, tug, first=True, sched=None, berth="T1", biko=None):
    r = [None] * 28
    r[0], r[1], r[3], r[8], r[16], r[11], r[12], r[17], r[14], r[24] = d.month, d.day, ship, io, hin, start, end, tug, berth, biko
    r[10] = sched or start
    r[5], r[6] = 1000, "試験"
    if first:
        r[2], r[4] = 1, 1
    return ("試験", tuple(r))


def make_copy(name, rows, target):
    path = os.path.join(TDIR, name)
    jobs = gen.build(rows, target, path)
    return path, jobs


def add_test_ships(wb, ships, company="試験会社"):
    """マスターへ試験会社と試験船を追加（試験専用コピーのみ）。"""
    m = wb["マスター"]
    m["Q15"], m["R15"], m["T15"] = company, "他社", "確認済み"
    for n, (name, base) in enumerate(ships):
        r = 40 + n
        m[f"A{r}"], m[f"C{r}"], m[f"D{r}"], m[f"E{r}"], m[f"J{r}"] = name, company, "タグ", base, "確認済み"


def open_vals(path):
    return openpyxl.load_workbook(path, data_only=True)


# ---------------------------------------------------------------- 1. 同時稼働
def test1():
    tg = D(2025, 11, 1)
    ships = [("試験A船", "門司港"), ("試験B船", "門司港"), ("試験C船", "門司港")]
    rows = [trow(tg, "試験本船1", "入", "鋼材", T(9, 0), T(10, 0), "試験A船"),
            trow(tg, "試験本船1", "入", "鋼材", T(9, 0), T(10, 0), "試験A船", first=False),
            trow(tg, "試験本船2", "入", "鋼材", T(9, 0), T(10, 0), "試験B船"),
            trow(tg, "試験本船3", "出", "鋼材", T(9, 30), T(10, 30), "試験C船")]
    path, _ = make_copy("試験1a_同時稼働と同一記録の重複.xlsx", rows, tg)
    wb = openpyxl.load_workbook(path)
    add_test_ships(wb, ships)
    wb.save(path)
    info = recalc(path)
    v = open_vals(path)
    s, j = v["集計"], v["実績"]
    check(1, "1-1", "A船9:00～10:00・B船9:00～10:00・C船9:30～10:30（別の本船作業）の最大同時は3隻（9:30）",
          "実績（試験）：本船1にA船、本船2にB船、本船3にC船", [3, "09:30", 3], [s["C14"].value, hhmm(s["C15"].value), s["C16"].value])
    check(1, "1-2", "同じA船の同一記録を重複登録しても4隻にしない（警告は「同一記録の重複」）",
          "本船1のA船の行をもう1行追加（同じ時刻）", [4, 3, 2, 0, "同一記録の重複（1隻として数える）"],
          [s["C10"].value, s["C14"].value, s["C25"].value, s["C24"].value, j["AH6"].value])
    check(1, "1-3", "数式エラーなし（試験コピー1a）", "—", 0, info["total_errors"])
    rows_b = [trow(tg, "試験本船1", "入", "鋼材", T(9, 0), T(10, 0), "試験A船"),
              trow(tg, "試験本船2", "入", "鋼材", T(9, 0), T(10, 0), "試験B船"),
              trow(tg, "試験本船3", "出", "鋼材", T(9, 30), T(10, 30), "試験C船"),
              trow(tg, "試験本船4", "転", "鋼材", T(9, 15), T(9, 45), "試験A船")]
    path, _ = make_copy("試験1b_重複配船の警告.xlsx", rows_b, tg)
    wb = openpyxl.load_workbook(path)
    add_test_ships(wb, ships)
    wb.save(path)
    info = recalc(path)
    v = open_vals(path)
    s, j = v["集計"], v["実績"]
    check(1, "1-4", "同じA船を同時刻の別作業（本船4、9:15～9:45）にも割り当てると重複配船の警告。同時隻数は3隻のまま",
          "本船4にA船（9:15～9:45）を追加", [2, "重複配船の警告（同じ船が同じ時間に別作業）", "重複配船の警告（同じ船が同じ時間に別作業）", "", 3],
          [s["C24"].value, j["AH6"].value, j["AH9"].value, j["AH7"].value or "", s["C14"].value])
    check(1, "1-5", "数式エラーなし（試験コピー1b）", "—", 0, info["total_errors"])


# ---------------------------------------------------------------- 2. 複数回の変更
def test2():
    tg = D(2025, 11, 2)
    rows = [trow(tg, "試験本船R", "入", "鋼材", T(9, 0), T(10, 0), "T", sched=T(9, 0))]
    path, jobs = make_copy("試験2_複数回の変更.xlsx", rows, tg)
    lab = jobs[0]["label"]
    wb = openpyxl.load_workbook(path)
    z, h = wb["前日確定"], wb["変更・キャンセル"]
    z["J26"] = T(9, 0)
    z["K26"] = "T"
    h["A26"], h["G26"], h["J26"], h["T26"] = lab, dt.datetime(2025, 11, 2, 7, 0), T(11, 0), "本船都合"
    h["A27"], h["G27"], h["J27"], h["T27"] = lab, dt.datetime(2025, 11, 2, 9, 30), T(13, 0), "荷役都合"
    wb.save(path)
    info = recalc(path)
    v = open_vals(path)
    z, h, s = v["前日確定"], v["変更・キャンセル"], v["集計"]
    check(2, "2-1", "前日確定9:00は上書きされずに残る", "前日確定 J26＝9:00", ["09:00", "11/02 09:00"], [hhmm(z["J26"].value), hhmm(z["Y26"].value)])
    check(2, "2-2", "1回目の変更（9:00→11:00）が残る", "変更 26行：同じ作業、変更後11:00",
          [1, "11/02 09:00", "11/02 09:00", "11:00", "11/02 11:00"],
          [h["D26"].value, hhmm(h["E26"].value), hhmm(h["I26"].value), hhmm(h["J26"].value), hhmm(h["AA26"].value)])
    check(2, "2-3", "2回目の変更（11:00→13:00）が残り、変更前は1回目の変更後になる", "変更 27行：同じ作業、変更後13:00",
          [2, "11/02 09:00", "11/02 11:00", "13:00", "11/02 13:00"],
          [h["D27"].value, hhmm(h["E27"].value), hhmm(h["I27"].value), hhmm(h["J27"].value), hhmm(h["AA27"].value)])
    check(2, "2-4", "前日確定シートに変更回数2・最終13:00、集計に変更2行・作業1件", "—",
          [2, "11/02 13:00", 2, 1], [z["U26"].value, hhmm(z["V26"].value), s["C62"].value, s["C63"].value])
    check(2, "2-5", "数式エラーなし（試験コピー2）", "—", 0, info["total_errors"])


# ---------------------------------------------------------------- 3. 全取消し
def test3():
    tg = D(2025, 11, 3)
    path, _ = make_copy("試験3_全取消し.xlsx", [], tg)
    wb = openpyxl.load_workbook(path)
    z, h = wb["前日確定"], wb["変更・キャンセル"]
    z["C7"], z["C8"], z["C10"] = "【試験】架空の予定表（11/2作成）", dt.datetime(2025, 11, 2, 16, 0), "予定表と照合済み"
    z["D26"], z["E26"], z["F26"], z["J26"], z["S26"] = "試験船P", "入", "N4", T(8, 0), "予定表と照合済み"
    for c, t in zip("KLMN", ["T", "S", "Y", "M"]):
        z[f"{c}26"] = t
    z["D27"], z["E27"], z["F27"], z["J27"], z["S27"] = "試験船Q", "出", "N5", T(13, 0), "予定表と照合済み"
    z["K27"], z["L27"] = "T", "S"
    for r, (t, act) in zip(range(14, 18), [("T", "出勤して待機"), ("S", "出勤して待機"), ("Y", "出勤取消し"), ("M", "出勤取消し")]):
        z[f"B{r}"], z[f"E{r}"], z[f"F{r}"] = t, dt.datetime(2025, 11, 2, 16, 30), act
    h["B7"], h["C7"] = "変更あり（下の表に入力）", "○"
    labs = ["11/03 予定のみ#01 08:00 入 試験船P", "11/03 予定のみ#02 13:00 出 試験船Q"]
    for r, lab in zip((26, 27), labs):
        h[f"A{r}"], h[f"G{r}"], h[f"R{r}"], h[f"S{r}"], h[f"T{r}"] = lab, dt.datetime(2025, 11, 3, 5, 0), "○", "当日・基地発前", "強風・荒天"
    wb.save(path)
    info = recalc(path)
    v = open_vals(path)
    s, z, h = v["集計"], v["前日確定"], v["変更・キャンセル"]
    check(3, "3-1", "実績0件の日でも対象日と前日確定の記録が消えない", "実績なし（11/3）。前日確定に2作業・予定タグ計6",
          ["2025-11-03", "予定表と照合済み", 2, 6, 4, "08:00", "13:00"],
          [str(s["C5"].value)[:10], s["C48"].value, s["C50"].value, s["C51"].value, s["C52"].value, hhmm(z["J26"].value), hhmm(z["J27"].value)])
    check(3, "3-2", "前日の出勤手配4隻と実作業0隻を区別", "出勤手配 T・S・Y・M", [4, 0, 0, 0],
          [s["C53"].value, s["C55"].value, s["C14"].value, s["C9"].value])
    check(3, "3-3", "実際に出勤・待機したかを別に入力・集計できる", "T・S＝出勤して待機、Y・M＝出勤取消し", [2], [s["C54"].value])
    check(3, "3-4", "全作業の取消しと全船休航が記録される", "変更で2作業とも取消し○（当日・基地発前）、日ごとの確認＝変更あり・全船休航○",
          [2, 2, "○", "変更あり（下の表に入力）", "○", "○", "—"],
          [s["C64"].value, s["C62"].value, s["C61"].value, s["C60"].value, z["W26"].value, z["W27"].value, s["C15"].value])
    check(3, "3-5", "数式エラーなし（試験コピー3）", "—", 0, info["total_errors"])


# ---------------------------------------------------------------- 4. 休船からの呼出
def test4():
    tg = D(2025, 11, 4)
    rows = [trow(tg, "試験本船1", "入", "鋼材", T(8, 0), T(9, 0), "T"),
            trow(tg, "試験本船1", "入", "鋼材", T(8, 0), T(9, 0), "S", first=False),
            trow(tg, "試験本船2", "出", "鋼材", T(10, 0), T(11, 0), "T")]
    path, jobs = make_copy("試験4_休船からの呼出.xlsx", rows, tg)
    wb = openpyxl.load_workbook(path)
    h = wb["変更・キャンセル"]
    h["A15"], h["C15"], h["D15"], h["E15"], h["G15"] = "T", dt.datetime(2025, 11, 4, 6, 0), jobs[0]["label"], "常駐船が他作業で不足", "確認済み"
    wb.save(path)
    info = recalc(path)
    v = open_vals(path)
    s, h = v["集計"], v["変更・キャンセル"]
    dvs = [d for d in wb["変更・キャンセル"].data_validations.dataValidation if "A15:A22" in str(d.sqref)]
    check(4, "4-1", "同じ船を1回呼び出し、その後2作業に従事：呼出1回・作業2件", "呼出：T 6:00（最初の作業＝本船1）。実績：Tが本船1と本船2に従事",
          [1, 2, 2, "製鉄曳船", ""], [s["C67"].value, h["J15"].value, s["C68"].value, h["B15"].value, h["H15"].value or ""])
    check(4, "4-2", "呼出した船の欄はマスターの船名一覧から選ぶ設定", "—", ["=船名リスト"], [d.formula1 for d in dvs], "プログラム検査")
    check(4, "4-3", "数式エラーなし（試験コピー4）", "—", 0, info["total_errors"])


# ---------------------------------------------------------------- 5. 日をまたぐ作業
def test5():
    tg = D(2025, 10, 2)
    rows = [trow(D(2025, 10, 1), "試験LNG船A", "入", "LNG", T(23, 50), T(0, 30), "T", berth="ZL"),
            trow(tg, "試験LNG船B", "出", "LNG", T(0, 10), T(0, 40), "S", berth="ZL"),
            trow(tg, "試験LNG船C", "当直", "LNG", T(1, 0), T(3, 0), "Y", berth="ZL"),
            trow(tg, "試験LNG船D", "エスコート", "LNG", T(4, 0), T(4, 30), "M", berth="ZL")]
    path, jobs = make_copy("試験5_日をまたぐ作業.xlsx", rows, tg)
    wb = openpyxl.load_workbook(path)
    z, h = wb["前日確定"], wb["変更・キャンセル"]
    z["J26"] = T(23, 50)
    h["A26"], h["J26"] = jobs[0]["label"], dt.datetime(2025, 10, 2, 0, 20)
    wb.save(path)
    info = recalc(path)
    v = open_vals(path)
    s, j, z, h = v["集計"], v["実績"], v["前日確定"], v["変更・キャンセル"]
    check(5, "5-1", "10/2の0:10～0:30は2隻（A船10/1 23:50～10/2 0:30、B船10/2 0:10～0:40）", "対象日10/2",
          [2, "00:10", 2, 0], [s["C14"].value, hhmm(s["C15"].value), s["G10"].value, s["C17"].value])
    check(5, "5-2", "時刻だけでなく日付を保持（開始日時・終了日時・前日からの継続）", "—",
          ["2025-10-01 23:50", "2025-10-02 00:30", "前日から継続", "○", "2025-10-02 00:10"],
          [str(j["AE6"].value)[:16], str(j["AF6"].value)[:16], j["AG6"].value, j["Q6"].value, str(j["AE7"].value)[:16]])
    check(5, "5-3", "LNG入港・出港・警戒（当直）・エスコートを区別して保持", "入・出・当直・エスコート（品名LNG）",
          ["LNG入港", "LNG出港", "LNG警戒", "LNGエスコート"], [j["AC6"].value, j["AC7"].value, j["AC8"].value, j["AC9"].value])
    check(5, "5-4", "予定・変更でも日付を失わない（前日確定23:50は10/1、変更後は10/2 0:20）", "前日確定 23:50、変更後 2025/10/2 0:20",
          ["10/01 23:50", "10/01 23:50", "10/02 00:20", "10/02 00:20"],
          [hhmm(z["Y26"].value), hhmm(h["I26"].value), hhmm(h["AA26"].value), hhmm(z["V26"].value)])
    check(5, "5-5", "数式エラーなし（試験コピー5）", "—", 0, info["total_errors"])


# ---------------------------------------------------------------- 6. マスター追加
def test6():
    tg = D(2025, 11, 6)
    ships = [("試験船1", "門司港"), ("試験船2", "若松"), ("試験船3", "戸畑"), ("試験船4", "八幡")]
    rows = [trow(tg, "試験本船X", "入", "鋼材", T(9, 0), T(10, 0), "試験船1"),
            trow(tg, "試験本船X", "入", "鋼材", T(9, 0), T(10, 0), "試験船2", first=False),
            trow(tg, "試験本船Y", "出", "鋼材", T(9, 30), T(10, 30), "試験船3"),
            trow(tg, "試験本船Y", "出", "鋼材", T(9, 30), T(10, 30), "試験船4", first=False)]
    path, _ = make_copy("試験6_マスター追加.xlsx", rows, tg)
    wb = openpyxl.load_workbook(path)
    add_test_ships(wb, ships)
    z = wb["前日確定"]
    for r, (n, _) in zip(range(14, 18), ships):
        z[f"B{r}"] = n
    m = wb["マスター"]
    m["BA1"] = "=ROWS(船名リスト)"
    for i in range(4):
        m[f"BA{2 + i}"] = f"=INDEX(船名リスト,{35 + i})"
    m["BB1"] = "=ROWS(会社リスト)"
    wb.save(path)
    info = recalc(path)
    v = open_vals(path)
    s, z, m = v["集計"], v["前日確定"], v["マスター"]
    check(6, "6-1", "試験会社に4隻を追加すると、4隻とも船名の選択肢（船名リスト）に入る", "マスター40～43行に試験船1～4、会社欄に試験会社",
          [38, "試験船1", "試験船2", "試験船3", "試験船4", 10],
          [m["BA1"].value, m["BA2"].value, m["BA3"].value, m["BA4"].value, m["BA5"].value, m["BB1"].value])
    check(6, "6-2", "選んだ4隻の会社・基地港が自動表示", "出勤手配に試験船1～4を選択",
          ["試験会社"] * 4 + ["門司港", "若松", "戸畑", "八幡"],
          [z[f"C{r}"].value for r in range(14, 18)] + [z[f"D{r}"].value for r in range(14, 18)])
    check(6, "6-3", "会社別・会社区分別の集計と出勤手配隻数へ反映", "実績（試験）：本船Xに試験船1・2、本船Yに試験船3・4",
          ["試験会社", 4, 4, 4, 4, 4], [s["F47"].value, s["H47"].value, s["I47"].value, s["C40"].value, s["C53"].value, s["C14"].value])
    check(6, "6-4", "数式エラーなし（試験コピー6）", "—", 0, info["total_errors"])
    # 納品用ファイルに試験の文字が残っていないこと
    hits = []
    for ws in openpyxl.load_workbook(SRC).worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and "試験" in c.value:
                    hits.append(f"{ws.title}!{c.coordinate}")
    check(6, "6-5", "試験船・試験会社は納品用ファイルに残っていない", "納品用ファイルの全セルを検索（「試験」）", [], hits, "プログラム検査")


# ---------------------------------------------------------------- 7. 供給条件・安瀬・高炉原料
def test7():
    tg = D(2025, 11, 15)
    rows = [trow(tg, "試験鉱石船", "入", "鉱石", T(9, 0), T(11, 0), "S"),
            trow(tg, "試験鉱石船", "入", "鉱石", T(9, 0), T(11, 0), "Y", first=False),
            trow(tg, "試験鉱石船", "入", "鉱石", T(9, 0), T(11, 0), "M", first=False),
            trow(tg, "試験鋼材船", "入", "鋼材", T(9, 30), T(10, 30), "P"),
            trow(tg, "試験鋼材船", "入", "鋼材", T(9, 30), T(10, 30), "竜山", first=False),
            trow(tg, "試験安瀬船", "入", "支給炭", T(9, 45), T(10, 45), "さくら"),
            trow(tg, "試験副原料船", "出", "副原料", T(13, 0), T(14, 0), "S")]
    path, _ = make_copy("試験7_供給条件.xlsx", rows, tg)
    wb = openpyxl.load_workbook(path)
    wb["集計"]["C7"] = T(10, 0)
    i = wb["入渠"]
    i["A6"], i["C6"], i["D6"] = "T", D(2025, 11, 10), D(2025, 11, 20)
    o = wb["他港作業"]
    o["A6"], o["C6"], o["D6"], o["E6"], o["F6"] = "関", "試験港", dt.datetime(2025, 11, 15, 8, 0), dt.datetime(2025, 11, 15, 12, 0), "実時刻"
    t = wb["他社配置"]
    for r, (comp, ship, ok) in zip((6, 7, 8), [("シーゲートコーポレーション", "かざし丸", "可"), ("グリーンシッピング", "むさし丸2", None),
                                              ("春風海運", "春風", "不可")]):
        t[f"A{r}"], t[f"B{r}"], t[f"C{r}"], t[f"D{r}"], t[f"E{r}"], t[f"F{r}"] = comp, ship, "門司港", D(2025, 4, 1), D(2026, 3, 31), ok
    wb.save(path)
    info = recalc(path)
    v = open_vals(path)
    s = v["集計"]
    status = {s[f"M{r}"].value: s[f"P{r}"].value for r in range(22, 142) if s[f"M{r}"].value}
    cnt = {s[f"M{r}"].value: s[f"N{r}"].value for r in range(10, 20)}
    check(7, "7-1", "同じ時刻（10:00）に入渠中の船と他港作業中の船は供給可能から除かれる",
          "入渠：T 11/10～11/20、他港作業：関 11/15 8:00～12:00、確認時刻10:00", ["入渠中", "他港作業中", 1, 1, 1, 1],
          [status.get("T"), status.get("関"), cnt["入渠中"], cnt["他港作業中"], s["C73"].value, s["C74"].value])
    check(7, "7-2", "他社の対応可否が空欄の船は「未確認」で、供給可能に自動で加わらない",
          "他社配置：かざし丸＝可、むさし丸2＝空欄、春風＝不可", ["供給可能（他社・対応可）", "未確認（他社の対応可否が空欄）", "対応不可", 1, 1, 7, 1],
          [status.get("かざし丸"), status.get("むさし丸2"), status.get("春風"), cnt["供給可能（他社・対応可）"],
           cnt["供給可能（待機）"], cnt["未確認（他社の対応可否が空欄）"], cnt["対応不可"]])
    check(7, "7-3", "作業中の船・会社未確認の船も供給可能に入らない", "—",
          ["作業中", "作業中", 6, 9, "供給可能（待機）"], [status.get("S"), status.get("さくら"), cnt["作業中"], cnt["未確認（会社未確認）"], status.get("長")])
    check(7, "7-4", "安瀬（按分配船）は応援の集計へ入らない", "支給炭の作業にさくら、鋼材の作業に竜山（どちらも会社未確認）",
          [1, 1], [s["C44"].value, s["C45"].value])
    check(7, "7-5", "高炉原料を除いた集計が、全作業の集計と別に表示される", "鉱石の作業（S・Y・M）、鋼材（P・竜山）、支給炭（さくら）、副原料（S）",
          [7, 4, 2, 6, 3], [s["C30"].value, s["C31"].value, s["C32"].value, s["C33"].value, s["C34"].value])
    check(7, "7-6", "数式エラーなし（試験コピー7）", "—", 0, info["total_errors"])


# ---------------------------------------------------------------- 基礎確認（納品ファイル・プログラム検査）
def independent_day(target):
    path = glob.glob(os.path.join(ROOT, "input", "01_元資料", "作業実績*"))[0]
    ws = openpyxl.load_workbook(path, read_only=True)["R7年度"]
    iv, jobs, job, total = [], 0, 0, {"jobs": 0, "rows": 0}
    for i, r in enumerate(ws.iter_rows(values_only=True), 1):
        if i == 1 or r[0] is None:
            continue
        if r[4] == 1:
            total["jobs"] += 1
        if r[17] is not None:
            total["rows"] += 1
        if r[17] is None or gen.fy_date(r[0], r[1]) != target:
            continue
        if r[4] == 1:
            jobs += 1
            job = jobs
        s, e = r[11], r[12]
        sm, em = s.hour * 60 + s.minute, e.hour * 60 + e.minute
        em += 1440 if em < sm else 0
        iv.append((sm, em, r[17], job))
    best, at, nj, m3 = 0, None, 0, 0
    for t in range(0, 1440, 5):
        act = [(x[2], x[3]) for x in iv if x[0] <= t < x[1]]
        n = len({a for a, _ in act})
        m3 += 5 if n >= 3 else 0
        if n > best:
            best, at, nj = n, t, len({b for _, b in act})
    return {"rows": len(iv), "jobs": jobs, "max": best, "at": f"{at // 60:02d}:{at % 60:02d}", "jobs_at": nj, "min3": m3}, total


def base_checks():
    base = os.path.join(TDIR, "試験0_納品ファイルの再計算.xlsx")
    shutil.copy(SRC, base)
    info = recalc(base)
    check(0, "R1", "納品ファイルを再計算して数式エラーがない", "入力なし（実データ2025/12/6）", 0, info["total_errors"])
    v = open_vals(base)
    s = v["集計"]
    ind, total = independent_day(gen.TARGET)
    check(0, "R2", "元帳の取込件数と同時稼働が、元帳からの独立計算と一致（行数・作業数・最大同時・時刻・本船作業数・3隻以上の分）",
          "元帳R7年度 2025/12/6", [ind["rows"], ind["jobs"], ind["max"], ind["at"], ind["jobs_at"], ind["min3"]],
          [s["C10"].value, s["C9"].value, s["C14"].value, hhmm(s["C15"].value), s["C16"].value, s["C17"].value])
    check(0, "R3", "基礎集計（会社区分別の延べタグ：製鉄曳船・西日本海運・他社・未確認）", "元帳R7年度 2025/12/6", [13, 2, 0, 2],
          [s["C38"].value, s["C39"].value, s["C40"].value, s["C41"].value])
    check(0, "R4", "元帳R7年度全体の基礎集計（作業件数・タグ欄の行数）が前回報告と同じ", "元帳R7年度（読み取りのみ）",
          [1249, 3071], [total["jobs"], total["rows"]], "プログラム検査")

    wb = openpyxl.load_workbook(SRC)
    names = [ws.title for ws in wb.worksheets if ws.sheet_state == "visible"]
    check(0, "P1", "シート名が作成条件どおり", "—",
          ["はじめに", "実績", "前日確定", "変更・キャンセル", "入渠", "他港作業", "他社配置", "標準時間", "マスター", "集計"], names, "プログラム検査")
    with zipfile.ZipFile(SRC) as z:
        parts = z.namelist()
    check(0, "P2", "マクロ・VBA・外部リンクを含まない", "—", [False, False],
          [any("vbaProject" in p for p in parts), any("externalLink" in p for p in parts)], "プログラム検査")
    bad = []
    for ws in wb.worksheets:
        if ws.title == "はじめに":
            continue
        for row in ws.iter_rows(min_row=5, max_row=25):
            for c in row:
                v_ = c.value
                if isinstance(v_, str) and ("【入力" in v_ or "【自動" in v_) and "／" not in v_:
                    if c.column in (1, 2) and not c.fill.fgColor.rgb[-6:] in ("FFE97F", "BFBFBF"):
                        target = ws.cell(c.row, 3)  # 左に見出し、C列が値の欄
                        ok = target.fill.fgColor.rgb[-6:] == ("FFF9C4" if "【入力" in v_ else "EDEDED")
                    else:
                        ok = c.fill.fgColor.rgb[-6:] == ("FFE97F" if "【入力" in v_ else "BFBFBF")
                    if not ok:
                        bad.append(f"{ws.title}!{c.coordinate}")
    check(0, "P3", "見出しの【入力】は黄色・【自動】は灰色", "—", [], bad, "プログラム検査")
    dv = {ws.title: {str(d.sqref): d.formula1 for d in ws.data_validations.dataValidation} for ws in wb.worksheets}
    ship_ranges = {"前日確定": ["B14:B21", "K26:P45"], "変更・キャンセル": ["L26:Q85", "A15:A22"],
                   "入渠": ["A6:A55"], "他港作業": ["A6:A55"], "他社配置": ["B6:B55"]}
    wrong = [f"{s_}!{r}" for s_, rs in ship_ranges.items() for r in rs if dv[s_].get(r) != "=船名リスト"]
    check(0, "P4", "船名の入力欄はすべてマスターの船名一覧（▼）から選ぶ設定", "—", [], wrong, "プログラム検査")
    other = {"前日確定": ["C10", "F14:F21", "S26:S45"], "変更・キャンセル": ["A26:A85", "R26:R85", "C7:C11", "D15:D22"],
             "他社配置": ["A6:A55", "F6:F55"], "マスター": ["C6:C125", "F6:F125", "Y6:Y25"]}
    missing = [f"{s_}!{r}" for s_, rs in other.items() for r in rs if r not in dv[s_]]
    check(0, "P5", "作業・○・会社・選択肢の▼が入力欄に設定されている", "—", [], missing, "プログラム検査")
    sel = (wb.active.title, wb["はじめに"]["B17"].hyperlink.location, wb["前日確定"].sheet_view.selection[0].activeCell)
    check(0, "P6", "初期表示：はじめにのリンクから最初の入力欄（前日確定C7）へ案内", "—", ("はじめに", "'前日確定'!C7", "C7"), sel, "プログラム検査")
    sheets = set(wb.sheetnames)
    broken = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if c.hyperlink is not None:
                    loc = c.hyperlink.location or ""
                    m_ = re.match(r"'?([^'!]+)'?!", loc)
                    if not m_ or m_.group(1) not in sheets:
                        broken.append(f"リンク {ws.title}!{c.coordinate}")
                if isinstance(c.value, str) and c.value.startswith("="):
                    if "#REF!" in c.value:
                        broken.append(f"#REF {ws.title}!{c.coordinate}")
                    for ref in re.findall(r"(?:'([^']+)'|([^\s=(),&:'\"<>+\-*/$]+))!", c.value):
                        name = ref[0] or ref[1]
                        if name not in sheets:
                            broken.append(f"参照 {ws.title}!{c.coordinate}→{name}")
    for nm, dn in wb.defined_names.items():
        for ref in re.findall(r"(?:'([^']+)'|([^\s=(),&:'\"$]+))!", dn.attr_text):
            if (ref[0] or ref[1]) not in sheets:
                broken.append(f"名前 {nm}")
    check(0, "P7", "リンク切れ・参照切れがない（シート内リンク、式のシート参照、名前定義、#REF!）", "—", [], sorted(set(broken))[:20], "プログラム検査")


def main():
    if os.path.isdir(TDIR):
        shutil.rmtree(TDIR)
    os.makedirs(TDIR)
    base_checks()
    for t in (test1, test2, test3, test4, test5, test6, test7):
        t()
    with open(os.path.join(TDIR, "試験結果_v01.json"), "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=1, default=str)
    print("合格", sum(r["result"] == "合格" for r in results), "/", len(results))


if __name__ == "__main__":
    main()
