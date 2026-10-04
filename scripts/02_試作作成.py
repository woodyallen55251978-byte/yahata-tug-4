"""八幡地区_検証試作_v01.xlsx を作成する（実データ1日分の小規模試作）。

- 元帳（input/01_元資料/作業実績…xlsx のシート「R7年度」）から対象日1日分を読み取り、「実績」シートへ取り込む。
- 元資料は読み取りのみ。入力欄は空欄（または「未確認」）で作成し、推測で埋めない。
- 計算条件は output/計算条件_v01.md に記載。

使い方: python3 scripts/02_試作作成.py
"""
import datetime as dt
import glob
import os

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.hyperlink import Hyperlink

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "output", "八幡地区_検証試作_v01.xlsx")

# 対象日：R7年度で同時3隻以上の候補日のうち、複数本船（LNG入港を含む）と他社タグが重なる日。
# 上田資料「最大同時稼働隻数」に記載があり、資料の値（9隻）と今回計算（8隻）が異なるため照合対象として選んだ。
TARGET = dt.date(2025, 12, 6)

# 容量（行数）
JIS_FIRST, JIS_LAST = 6, 505          # 実績
M_FIRST, M_LAST = 6, 125              # マスター（船）
C_FIRST, C_LAST = 6, 45               # マスター（会社）
L_FIRST, L_LAST = 6, 25               # マスター（基地港・選択肢）
SLOT_FIRST = 6                        # _計算 の時間枠の最初の行
SLOTS = 288                           # 5分刻み × 24時間
SLOT_LAST = SLOT_FIRST + SLOTS - 1
GRID_COL0 = 3                         # _計算 の船列の開始（C列）
NSHIP = M_LAST - M_FIRST + 1          # 120列
GRID_LASTCOL = GRID_COL0 + NSHIP - 1
COL_TOTAL = GRID_LASTCOL + 1          # 全体の同時隻数
COL_JOBS = GRID_LASTCOL + 2           # 作業中の本船作業数
COL_STC = GRID_LASTCOL + 3            # 会社区分「製鉄曳船」の同時隻数
COL_NW = GRID_LASTCOL + 4             # 会社区分「西日本海運」の同時隻数
YJ_FIRST, YJ_LAST = 26, 45            # 前日確定：作業行
CH_FIRST, CH_LAST = 26, 85            # 変更・キャンセル：変更行
YB_FIRST, YB_LAST = 15, 22            # 変更・キャンセル：休船呼出
DAY_FIRST, DAY_LAST = 7, 11           # 変更・キャンセル：日ごとの確認

FONT = "游ゴシック"
F_IN = PatternFill("solid", fgColor="FFF9C4")    # 薄黄色＝入力
F_AUTO = PatternFill("solid", fgColor="EDEDED")  # 薄灰色＝自動
F_HEAD_IN = PatternFill("solid", fgColor="FFE97F")
F_HEAD_AUTO = PatternFill("solid", fgColor="BFBFBF")
F_TITLE = PatternFill("solid", fgColor="DDEBF7")
THIN = Side(style="thin", color="A6A6A6")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def font(**kw):
    return Font(name=FONT, size=kw.pop("size", 10), **kw)


def col(c):
    return get_column_letter(c)


def q(sheet):
    return "'" + sheet + "'"


def fy_date(m, d):
    return dt.date(2025 if m >= 4 else 2026, m, d)


# ---------------------------------------------------------------- 元帳の読み取り
def read_target_rows():
    path = glob.glob(os.path.join(ROOT, "input", "01_元資料", "作業実績*"))[0]
    ws = openpyxl.load_workbook(path, read_only=True)["R7年度"]
    rows, prev_cross = [], []
    for i, r in enumerate(ws.iter_rows(values_only=True), 1):
        if i == 1 or r[0] is None or r[17] is None:
            continue
        d = fy_date(r[0], r[1])
        if d == TARGET:
            rows.append((i, r))
        elif d == TARGET - dt.timedelta(1):
            s, e = r[11], r[12]
            if isinstance(e, dt.datetime) and (e.month, e.day) == (1, 1):
                prev_cross.append((i, r))
            elif isinstance(s, dt.time) and isinstance(e, dt.time) and e < s:
                prev_cross.append((i, r))
    return prev_cross + rows


def is_cancel(r):
    return any(isinstance(v, str) and ("ｷｬﾝｾﾙ" in v or "キャンセル" in v) for v in r)


# ---------------------------------------------------------------- 共通の書式
def sheet_head(ws, title, who, dest, width_to=12):
    ws["A1"] = title
    ws["A1"].font = font(size=14, bold=True)
    ws["A2"] = who
    ws["A2"].font = font(size=10)
    ws["A3"] = dest
    ws["A3"].font = font(size=10, color="595959")
    for r in (1, 2, 3):
        for c in range(1, width_to + 1):
            ws.cell(r, c).fill = F_TITLE


def header(ws, row, c, text, kind, width=None):
    cell = ws.cell(row, c, text)
    cell.font = font(bold=True)
    cell.fill = F_HEAD_IN if kind == "in" else F_HEAD_AUTO
    cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    cell.border = BOX
    if width:
        ws.column_dimensions[col(c)].width = width


def paint(ws, r1, r2, c1, c2, kind, fmt=None):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(r, c)
            cell.fill = F_IN if kind == "in" else F_AUTO
            cell.border = BOX
            cell.font = font()
            if fmt:
                cell.number_format = fmt


def section(ws, row, text, c2=10):
    ws.cell(row, 1, text).font = font(size=11, bold=True, color="1F4E79")


def add_list(ws, rng, source, prompt=None):
    dv = DataValidation(type="list", formula1=source, allow_blank=True, showDropDown=False)
    dv.error = "一覧から選んでください（一覧にない場合は「マスター」シートへ追加してください）"
    dv.errorTitle = "選択してください"
    if prompt:
        dv.prompt = prompt
        dv.showInputMessage = True
    ws.add_data_validation(dv)
    dv.add(rng)


def circle(ws, rng):
    add_list(ws, rng, '"○"', "○を選びます（空欄＝該当なし）")


def select(ws, cell_ref):
    ws.sheet_view.selection[0].activeCell = cell_ref
    ws.sheet_view.selection[0].sqref = cell_ref


# ---------------------------------------------------------------- マスター
MASTER_SHIPS = [
    # 選択名, 正式船名, 会社, 種類, 基地港, 情報元, 確認状況
    ("八幡丸", "八幡丸", "製鉄曳船", "タグ", "未確認", "00_作成条件.md", "要確認（基地港）"),
    ("鐵豊丸", "鐵豊丸", "製鉄曳船", "タグ", "未確認", "00_作成条件.md", "要確認（基地港）"),
    ("新豊丸", "新豊丸", "製鉄曳船", "タグ", "未確認", "00_作成条件.md", "要確認（基地港）"),
    ("八豊丸", "八豊丸", "製鉄曳船", "タグ", "未確認", "00_作成条件.md", "要確認（基地港）"),
    ("松豊丸", "松豊丸", "製鉄曳船", "タグ", "未確認", "00_作成条件.md", "要確認（基地港）"),
    ("T", "", "製鉄曳船", "タグ", "未確認", "元帳R7年度の記号。会社は既存集計（西日本G 2,744）との照合上の扱い", "要確認（正式船名・会社）"),
    ("S", "", "製鉄曳船", "タグ", "未確認", "同上", "要確認（正式船名・会社）"),
    ("Y", "", "製鉄曳船", "タグ", "未確認", "同上", "要確認（正式船名・会社）"),
    ("M", "", "製鉄曳船", "タグ", "未確認", "同上", "要確認（正式船名・会社）"),
    ("P", "", "製鉄曳船", "タグ", "未確認", "同上", "要確認（正式船名・会社）"),
    ("関豊丸", "関豊丸", "西日本海運", "タグ", "門司港", "00_作成条件.md／要目表2026.8", "確認済み"),
    ("長豊丸", "長豊丸", "西日本海運", "タグ", "門司港", "00_作成条件.md／要目表2026.8", "確認済み"),
    ("関", "", "西日本海運", "タグ", "未確認", "元帳R7年度の記号。会社は既存集計との照合上の扱い", "要確認（正式船名・会社）"),
    ("長", "", "西日本海運", "タグ", "未確認", "同上", "要確認（正式船名・会社）"),
    ("かざし丸", "かざし丸", "シーゲートコーポレーション", "タグ", "門司港", "要目表2026.8（2025年度の所属は未確認）", "要確認（2025年度の所属）"),
    ("鷲羽丸", "鷲羽丸", "シーゲートコーポレーション", "タグ", "門司港", "要目表2026.8（2025年度の所属は未確認）", "要確認（2025年度の所属）"),
    ("むさし丸2", "むさし丸2", "グリーンシッピング", "タグ", "門司港", "要目表2026.8（2025年度の所属は未確認）", "要確認（2025年度の所属）"),
    ("かなで", "かなで", "グリーンシッピング", "タグ", "門司港", "要目表2026.8（2025年度の所属は未確認）", "要確認（2025年度の所属）"),
    ("かいせい", "かいせい", "日鉄物流", "タグ", "門司港", "要目表2026.8（2025年度の所属は未確認）", "要確認（2025年度の所属）"),
    ("春風", "春風", "春風海運", "タグ", "門司港", "00_作成条件.md（ユーザー指定）", "確認済み"),
    ("栄春", "栄春", "春風海運", "タグ", "門司港", "00_作成条件.md（ユーザー指定）", "確認済み"),
    ("双美", "双美", "春風海運", "タグ", "門司港", "00_作成条件.md（ユーザー指定）", "確認済み"),
    ("春幸", "春幸", "春風海運", "タグ", "門司港", "00_作成条件.md（ユーザー指定）", "確認済み"),
    ("さくら", "", "", "タグ", "未確認", "元帳R7年度（S列「洞」は参考。会社は推定しない）", "要確認（正式船名・会社）"),
    ("みさき", "", "", "タグ", "未確認", "元帳R7年度（S列「洞」は参考。会社は推定しない）", "要確認（正式船名・会社）"),
    ("竜山", "", "", "タグ", "未確認", "元帳R7年度（S列「矢」は参考。会社は推定しない）", "要確認（正式船名・会社）"),
    ("竜峰", "", "", "タグ", "未確認", "元帳R7年度（S列「矢」は参考。会社は推定しない）", "要確認（正式船名・会社）"),
    ("竜昇", "", "", "タグ", "未確認", "元帳R7年度（S列「矢」は参考。会社は推定しない）", "要確認（正式船名・会社）"),
    ("竜海", "", "", "タグ", "未確認", "元帳R7年度（S列「矢」は参考。会社は推定しない）", "要確認（正式船名・会社）"),
    ("慶鳳丸", "", "", "タグ", "未確認", "元帳R7年度", "要確認（正式船名・会社）"),
    ("秀豊", "", "", "タグ", "未確認", "元帳R7年度（秀豊丸と同一船かは未確認）", "要確認（正式船名・会社）"),
    ("秀豊丸", "", "", "タグ", "未確認", "元帳R7年度（秀豊と同一船かは未確認）", "要確認（正式船名・会社）"),
    ("まんじゅ", "", "", "未確認", "未確認", "元帳R7年度。LNGエスコートのみ・時刻空欄。既存集計では除外とみられる", "要確認（種類・会社）"),
    ("せいじゅ", "", "", "未確認", "未確認", "元帳R7年度。LNGエスコートのみ・時刻空欄。既存集計では除外とみられる", "要確認（種類・会社）"),
]
MASTER_COMPANIES = [
    ("製鉄曳船", "製鉄曳船", "00_作成条件.md"),
    ("西日本海運", "西日本海運", "00_作成条件.md"),
    ("シーゲートコーポレーション", "他社", "00_作成条件.md／要目表2026.8"),
    ("グリーンシッピング", "他社", "00_作成条件.md／要目表2026.8"),
    ("日鉄物流", "他社", "00_作成条件.md／要目表2026.8"),
    ("春風海運", "他社", "00_作成条件.md"),
    ("矢野海運", "他社", "00_作成条件.md（船名は未確認）"),
    ("洞海マリン", "他社", "00_作成条件.md（船名は未確認）"),
    ("福島海運", "他社", "00_作成条件.md（船名は未確認）"),
]
BASE_PORTS = ["門司港", "戸畑", "八幡", "若松", "未確認"]
CHOICES = {  # 列, 見出し, 名前, 値
    "W": ("確認状況", "選択_確認状況", ["未確認", "確認済み", "要確認（正式船名・会社）", "要確認（基地港）", "要確認（2025年度の所属）", "要確認（種類・会社）"]),
    "X": ("照合の状況", "選択_照合", ["未確認", "予定表と照合済み", "予定表に記載なし", "予定表なし"]),
    "Y": ("取消しの時点", "選択_取消時点", ["前日", "当日・基地発前", "当日・基地発後", "不明"]),
    "Z": ("変更・取消しの理由", "選択_理由", ["強風・荒天", "視界不良", "本船都合", "荷役都合", "代理店・荷主都合", "バース都合", "タグの都合", "その他", "不明"]),
    "AA": ("時間調整", "選択_時間調整", ["あり", "なし", "不明"]),
    "AB": ("呼出の理由", "選択_呼出理由", ["常駐船が他作業で不足", "入渠中の船の代わり", "他港作業中の船の代わり", "乗組員の出勤手配の不足", "その他", "不明"]),
    "AC": ("作業内容", "選択_作業内容", ["入港", "出港", "転錨・シフト", "LNG入港", "LNG出港", "LNG警戒", "エスコート", "その他"]),
    "AD": ("予定時刻の基準", "選択_時刻基準", ["本船POB", "着岸", "離岸", "作業開始", "基地発", "未確認"]),
    "AE": ("船の種類", "選択_種類", ["タグ", "通船・警戒船", "未確認"]),
    "AF": ("時刻の確認状況", "選択_時刻確認", ["実時刻", "予定・推定", "不明"]),
    "AG": ("会社区分", "選択_会社区分", ["製鉄曳船", "西日本海運", "他社", "未確認"]),
    "AH": ("その日の変更確認", "選択_日確認", ["未確認", "変更なしと確認済み", "変更あり（下の表に入力）"]),
}


def build_master(wb):
    ws = wb.create_sheet("マスター")
    sheet_head(ws, "マスター（会社・船名・基地港・選択肢）",
               "誰が：上田・各担当者　何を見て：要目表・元帳・予定表　どこへ：薄黄色の行の下の空き行へ追加（空行を空けずに続けて入力）",
               "集計先：全シートの選択肢（プルダウン）、実績の会社判定、集計の会社別。正式船名を入れると同じ船として1隻に数えます。", 34)
    heads = [("選択名（元帳・予定表の表記）【入力】", "in", 16), ("正式船名【入力】", "in", 12), ("会社【入力・選択】", "in", 18),
             ("船の種類【入力・選択】", "in", 10), ("基地港【入力・選択】", "in", 10), ("適用開始日【入力】", "in", 11),
             ("適用終了日【入力】", "in", 11), ("情報元【入力】", "in", 30), ("確認状況【入力・選択】", "in", 18),
             ("集計船名【自動】", "auto", 11), ("会社区分【自動】", "auto", 10), ("対象日に有効【自動】", "auto", 8),
             ("照合名【自動】", "auto", 8), ("同一船の先頭【自動】", "auto", 8)]
    for i, (t, k, w) in enumerate(heads, 1):
        header(ws, 5, i, t, k, w)
    ws.row_dimensions[5].height = 45
    paint(ws, M_FIRST, M_LAST, 1, 9, "in")
    paint(ws, M_FIRST, M_LAST, 10, 14, "auto")
    for c in (6, 7):
        for r in range(M_FIRST, M_LAST + 1):
            ws.cell(r, c).number_format = "yyyy/mm/dd"
    for n, row in enumerate(MASTER_SHIPS):
        r = M_FIRST + n
        name, formal, comp, kind, base, src, status = row
        for c, v in zip((1, 2, 3, 4, 5, 8, 9), (name, formal, comp, kind, base, src, status)):
            ws.cell(r, c, v if v != "" else None)
    for r in range(M_FIRST, M_LAST + 1):
        ws.cell(r, 10, f'=IF(A{r}="","",IF(B{r}="",A{r},B{r}))')
        ws.cell(r, 11, f'=IF(A{r}="","",IF(C{r}="","未確認",IFERROR(INDEX($Q${C_FIRST}:$Q${C_LAST},MATCH(C{r},$P${C_FIRST}:$P${C_LAST},0)),"未確認")))')
        ws.cell(r, 12, f'=IF(A{r}="","",IF(AND(OR(F{r}="",F{r}<=集計!$C$5),OR(G{r}="",G{r}>=集計!$C$5)),"○",""))')
        ws.cell(r, 13, f'=IF(L{r}="○",A{r},"")')
        ws.cell(r, 14, f'=IF(OR(J{r}="",L{r}<>"○"),"",IF(COUNTIFS(J${M_FIRST}:J{r},J{r},L${M_FIRST}:L{r},"○")=1,"○",""))')
    # 会社
    for i, (t, k, w) in enumerate([("会社名【入力】", "in", 22), ("会社区分【入力・選択】", "in", 10),
                                     ("情報元【入力】", "in", 24), ("確認状況【入力・選択】", "in", 12)]):
        header(ws, 5, 16 + i, t, k, w)
    paint(ws, C_FIRST, C_LAST, 16, 19, "in")
    for n, (name, kubun, src) in enumerate(MASTER_COMPANIES):
        r = C_FIRST + n
        ws.cell(r, 16, name)
        ws.cell(r, 17, kubun)
        ws.cell(r, 18, src)
        ws.cell(r, 19, "確認済み" if kubun in ("製鉄曳船", "西日本海運") or name == "春風海運" else "未確認")
    # 基地港
    header(ws, 5, 21, "基地港【入力】", "in", 10)
    paint(ws, L_FIRST, L_LAST, 21, 21, "in")
    for n, v in enumerate(BASE_PORTS):
        ws.cell(L_FIRST + n, 21, v)
    # 選択肢
    for letter, (title, _, values) in CHOICES.items():
        c = openpyxl.utils.column_index_from_string(letter) + 1  # V列を空けるため1列右へ
        header(ws, 5, c, title + "【入力】", "in", 14)
        paint(ws, L_FIRST, L_LAST, c, c, "in")
        for n, v in enumerate(values):
            ws.cell(L_FIRST + n, c, v)
    ws.cell(4, 1, "※ グリーンシッピングの応援船が4隻あった場合は、船名を確認のうえ下の空き行へ追加すると、選択肢・集計へ反映されます（元帳R7年度に出る同社の船名は要目表にある「むさし丸2」「かなで」の2隻のみ）。").font = font(color="C00000")
    ws.cell(4, 16, "※ 会社を追加すると、会社の選択肢と「集計」の会社別へ反映されます。").font = font(color="C00000")
    ws.freeze_panes = "B6"
    select(ws, f"A{M_FIRST + len(MASTER_SHIPS)}")

    def dyn(name, c, first, last):
        L = col(c)
        wb.defined_names[name] = DefinedName(
            name, attr_text=f"OFFSET(マスター!${L}${first},0,0,MAX(1,COUNTA(マスター!${L}${first}:${L}${last})),1)")
    dyn("船名リスト", 1, M_FIRST, M_LAST)
    dyn("会社リスト", 16, C_FIRST, C_LAST)
    dyn("基地港リスト", 21, L_FIRST, L_LAST)
    for letter, (_, nm, _) in CHOICES.items():
        dyn(nm, openpyxl.utils.column_index_from_string(letter) + 1, L_FIRST, L_LAST)
    add_list(ws, f"C{M_FIRST}:C{M_LAST}", "=会社リスト")
    add_list(ws, f"D{M_FIRST}:D{M_LAST}", "=選択_種類")
    add_list(ws, f"E{M_FIRST}:E{M_LAST}", "=基地港リスト")
    add_list(ws, f"I{M_FIRST}:I{M_LAST}", "=選択_確認状況")
    add_list(ws, f"Q{C_FIRST}:Q{C_LAST}", "=選択_会社区分")
    add_list(ws, f"S{C_FIRST}:S{C_LAST}", "=選択_確認状況")
    return ws


# ---------------------------------------------------------------- 実績
JIS_COLS = [
    ("作業キー【自動】", 14), ("作業の表示名【自動】", 34), ("作業の先頭行【自動】", 7), ("作業日【自動】", 11),
    ("元帳の行番号【自動】", 8), ("元帳No【自動】", 6), ("本船名【自動】", 18), ("G/T【自動】", 9),
    ("代理店（参考・会社判定に使わない）【自動】", 11), ("入出転【自動】", 7), ("六/部【自動】", 6), ("バース【自動】", 8),
    ("品名【自動】", 8), ("予定時間（元帳・定義未確認）【自動】", 10), ("作業開始（元帳）【自動】", 9), ("作業終了（元帳）【自動】", 9),
    ("終了が翌日【自動】", 7), ("タグ（元帳の記号）【自動】", 9), ("S列の記号（参考）【自動】", 7), ("備考1（元帳）【自動】", 16),
    ("取消しの記載【自動】", 7), ("時刻の扱い【自動】", 16),
    ("集計船名【自動】", 10), ("会社【自動】", 14), ("会社区分【自動】", 10), ("船名・会社の確認状況【自動】", 18),
    ("船の種類【自動】", 9), ("安瀬の扱い【自動】", 12),
    ("開始（分）【自動】", 7), ("終了（分）【自動】", 7), ("同時稼働に含む【自動】", 7), ("作業の開始（分）【自動】", 7),
    ("作業の終了（分）【自動】", 7), ("重複配船の警告【自動】", 18),
]


def build_jisseki(wb, rows):
    ws = wb.create_sheet("実績")
    sheet_head(ws, "実績（元帳から自動表示・再入力しない）",
               "誰が：全員が見るだけ　何を見て：元帳「作業実績R6.4～」シート「R7年度」　どこへ：入力欄はありません（休船からの呼出は「変更・キャンセル」で最初に対応した作業を選びます）",
               f"対象日 {TARGET:%Y/%m/%d}。時刻は元帳の「作業時間」で、基地発・基地着かは未確認のため「候補・時刻要確認」。集計先：「集計」「前日確定」「変更・キャンセル」", 22)
    for i, (t, w) in enumerate(JIS_COLS, 1):
        header(ws, 5, i, t, "auto", w)
    ws.row_dimensions[5].height = 60
    last = JIS_FIRST + len(rows) - 1
    paint(ws, JIS_FIRST, max(last, JIS_FIRST), 1, len(JIS_COLS), "auto")
    jobs, jobno, key = [], 0, None
    for n, (i, r) in enumerate(rows):
        R = JIS_FIRST + n
        d = fy_date(r[0], r[1])
        first = r[4] == 1 or r[2] is not None
        if first or key is None:
            jobno += 1
            key = f"{d:%Y-%m-%d}#{jobno:02d}"
            label = (f"{d:%m/%d} #{jobno:02d} {r[10].strftime('%H:%M') if isinstance(r[10], dt.time) else '--:--'} "
                     f"{r[8] or ''} {r[3]}（{r[14] or ''}）")
            jobs.append({"key": key, "label": label, "ship": r[3], "io": r[8], "berth": r[14], "sched": r[10], "tugs": []})
        jobs[-1]["tugs"].append(r[17])
        s, e = r[11], r[12]
        cross = ""
        bad = False
        if isinstance(e, dt.datetime):
            if (e.year, e.month, e.day) == (1900, 1, 1):
                e = e.time()
                cross = "○"
            else:
                bad = True
        elif isinstance(s, dt.time) and isinstance(e, dt.time) and e < s:
            cross = "○"
        timing = "候補・時刻要確認" if (isinstance(s, dt.time) and isinstance(e, dt.time) and not bad) else "時刻なし（集計対象外）"
        vals = [key, jobs[-1]["label"], "○" if first else "", d, i, r[2], r[3], r[5], r[6], r[8], r[9], r[14],
                (r[16] or "").strip() or None, r[10], s if isinstance(s, dt.time) else None,
                e if isinstance(e, dt.time) else None, cross, r[17], r[18], r[24], "○" if is_cancel(r) else "", timing]
        for c, v in enumerate(vals, 1):
            ws.cell(R, c, v)
        ws.cell(R, 4).number_format = "yyyy/mm/dd"
        for c in (14, 15, 16):
            ws.cell(R, c).number_format = "hh:mm"
        m = f"$M${M_FIRST}:$M${M_LAST}"
        look = lambda c: f"INDEX(マスター!${c}${M_FIRST}:${c}${M_LAST},MATCH(R{R},マスター!{m},0))"
        ws.cell(R, 23, f"=IFERROR({look('J')},R{R})")
        ws.cell(R, 24, f'=IFERROR(IF({look("C")}="","未確認",{look("C")}),"マスター未登録")')
        ws.cell(R, 25, f'=IFERROR({look("K")},"マスター未登録")')
        ws.cell(R, 26, f'=IFERROR(IF({look("I")}="","未確認",{look("I")}),"マスター未登録")')
        ws.cell(R, 27, f'=IFERROR(IF({look("D")}="","未確認",{look("D")}),"マスター未登録")')
        ws.cell(R, 28, f'=IF(M{R}="支給炭","安瀬候補（要確認）","")')
        ws.cell(R, 29, f'=IF(OR(O{R}="",V{R}<>"候補・時刻要確認"),"",ROUND((D{R}-集計!$C$5)*1440+O{R}*1440,0))')
        ws.cell(R, 30, f'=IF(OR(P{R}="",V{R}<>"候補・時刻要確認"),"",ROUND((D{R}-集計!$C$5)*1440+P{R}*1440+IF(Q{R}="○",1440,0),0))')
        ws.cell(R, 31, f'=IF(AND(AC{R}<>"",AD{R}<>"",U{R}<>"○",AA{R}<>"通船・警戒船"),"○","")')
        rng = lambda c: f"${c}${JIS_FIRST}:${c}${JIS_LAST}"
        ws.cell(R, 32, f'=IF(C{R}<>"○","",IF(COUNTIFS({rng("A")},A{R},{rng("AE")},"○")=0,"",_xlfn.MINIFS({rng("AC")},{rng("A")},A{R},{rng("AE")},"○")))')
        ws.cell(R, 33, f'=IF(C{R}<>"○","",IF(COUNTIFS({rng("A")},A{R},{rng("AE")},"○")=0,"",_xlfn.MAXIFS({rng("AD")},{rng("A")},A{R},{rng("AE")},"○")))')
        ws.cell(R, 34, f'=IF(AE{R}<>"○","",IF(COUNTIFS({rng("W")},W{R},{rng("AC")},"<"&AD{R},{rng("AD")},">"&AC{R},{rng("AE")},"○")>1,"同じ船の時間が重なっています",""))')
    for c in range(29, 34):
        ws.column_dimensions[col(c)].hidden = True
    ws.freeze_panes = "H6"
    select(ws, "A6")
    return ws, jobs


# ---------------------------------------------------------------- 前日確定
def build_zenjitsu(wb, jobs):
    ws = wb.create_sheet("前日確定")
    sheet_head(ws, "前日確定（対象日の前日最終の予定表と翌日の出勤手配）",
               "誰が：上田　何を見て：対象日の前日に翌日の出勤を手配した最終の作業予定表　どこへ：①C7～C10　②出勤手配（B14～）　③作業ごとの予定（J列～T列）。予定表にあって実績にない作業は36行目以降の空き行へ追加",
               "集計先：「集計」の予定・出勤手配、「変更・キャンセル」の変更前の時刻・タグ。ここに入れた前日最終の値は変更があっても上書きしません。", 23)
    section(ws, 5, "■ 対象日の予定表")
    labels = [("対象日【自動】", "auto"), ("根拠の予定表（ファイル名・版）【入力】", "in"), ("予定表の確定時点（日時）【入力】", "in"),
              ("照合した人【入力】", "in"), ("照合の状況【入力・選択】", "in")]
    for n, (t, k) in enumerate(labels):
        r = 6 + n
        ws.cell(r, 1, t).font = font(bold=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
        paint(ws, r, r, 3, 5, k)
    ws["C6"] = "=集計!$C$5"
    ws["C6"].number_format = "yyyy/mm/dd"
    ws["C8"].number_format = "yyyy/mm/dd hh:mm"
    ws["C10"] = "未確認"
    ws["F7"] = "◀ 最初の入力はここ（対象日の前日最終の予定表のファイル名・版。2025年度分は本一式に含まれていないため、上田さんの照合で入力）"
    ws["F7"].font = font(bold=True, color="C00000")
    add_list(ws, "C10", "=選択_照合")

    section(ws, 12, "■ 翌日の出勤手配（対象日の前日に手配した船）")
    for i, (t, k, w) in enumerate([("No", "auto", None), ("船名【入力・選択】", "in", None), ("会社【自動】", "auto", None),
                                     ("手配した日時【入力】", "in", None), ("情報元【入力】", "in", None)]):
        header(ws, 13, 1 + i, t, k, w)
    paint(ws, 14, 21, 1, 1, "auto")
    paint(ws, 14, 21, 2, 2, "in")
    paint(ws, 14, 21, 3, 3, "auto")
    paint(ws, 14, 21, 4, 5, "in")
    for r in range(14, 22):
        ws.cell(r, 1, r - 13)
        ws.cell(r, 3, f'=IF(B{r}="","",IFERROR(INDEX(マスター!$C${M_FIRST}:$C${M_LAST},MATCH(B{r},マスター!$M${M_FIRST}:$M${M_LAST},0))&"","マスター未登録"))')
        ws.cell(r, 4).number_format = "yyyy/mm/dd hh:mm"
    add_list(ws, "B14:B21", "=船名リスト", "マスターの船名から選びます")
    ws["A22"] = "出勤手配隻数【自動】"
    ws["A22"].font = font(bold=True)
    ws.merge_cells("A22:B22")
    ws["C22"] = "=IF(COUNTA(B14:B21)=0,0,SUMPRODUCT((B14:B21<>\"\")/COUNTIF(B14:B21,B14:B21&\"\")))"
    paint(ws, 22, 22, 3, 3, "auto")

    section(ws, 24, "■ 作業ごとの前日最終予定（実績の作業は自動表示。予定表にだけある作業は下の空き行へ追加）")
    heads = [("区分【自動】", "auto", 9), ("作業の表示名【自動】", "auto", 34), ("作業キー【自動】", "auto", 14),
             ("本船名【自動／予定のみは入力】", "auto", 18), ("入出転【自動／予定のみは入力】", "auto", 7), ("バース【自動／予定のみは入力】", "auto", 8),
             ("実績のタグ（元帳）【自動】", "auto", 14), ("実績タグ数【自動】", "auto", 7), ("元帳の予定時間（参考・前日確定とは限らない）【自動】", "auto", 11),
             ("前日最終の予定時刻【入力】", "in", 10)] + \
            [(f"予定タグ{c}【入力・選択】", "in", 8) for c in "①②③④⑤⑥"] + \
            [("予定タグ数【自動】", "auto", 7), ("予定タグ（まとめ）【自動】", "auto", 16), ("照合の状況【入力・選択】", "in", 14),
             ("備考【入力】", "in", 18), ("変更回数【自動】", "auto", 7), ("変更後の最終時刻【自動】", "auto", 9), ("取消し【自動】", "auto", 7)]
    for i, (t, k, w) in enumerate(heads, 1):
        header(ws, 25, i, t, k, w)
    ws.row_dimensions[25].height = 60
    paint(ws, YJ_FIRST, YJ_LAST, 1, 9, "auto")
    paint(ws, YJ_FIRST, YJ_LAST, 10, 16, "in")
    paint(ws, YJ_FIRST, YJ_LAST, 17, 18, "auto")
    paint(ws, YJ_FIRST, YJ_LAST, 19, 20, "in")
    paint(ws, YJ_FIRST, YJ_LAST, 21, 23, "auto")
    for n, j in enumerate(jobs):
        r = YJ_FIRST + n
        for c, v in enumerate(["実績あり", j["label"], j["key"], j["ship"], j["io"], j["berth"], "・".join(j["tugs"])], 1):
            ws.cell(r, c, v)
        ws.cell(r, 9, j["sched"])
        ws.cell(r, 19, "未確認")
    for r in range(YJ_FIRST + len(jobs), YJ_LAST + 1):
        n = r - YJ_FIRST - len(jobs) + 1
        paint(ws, r, r, 4, 6, "in")
        ws.cell(r, 1, f'=IF(D{r}="","","予定のみ")')
        ws.cell(r, 3, f'=IF(D{r}="","","{TARGET:%Y-%m-%d}予定のみ#{n:02d}")')
        ws.cell(r, 2, f'=IF(D{r}="","","{TARGET:%m/%d} 予定のみ#{n:02d} "&IF(J{r}="","--:--",TEXT(J{r},"hh:mm"))&" "&E{r}&" "&D{r})')
        ws.cell(r, 7, f'=IF(D{r}="","","（実績なし）")')
    ws.cell(YJ_FIRST + len(jobs) - 1, 24, "▼ ここから下は予定表にだけある作業（実績のない取消し等）を追加する行").font = font(color="C00000")
    chk = f"'変更・キャンセル'"
    for r in range(YJ_FIRST, YJ_LAST + 1):
        ws.cell(r, 8, f'=IF(C{r}="","",COUNTIF(実績!$A${JIS_FIRST}:$A${JIS_LAST},C{r}))')
        ws.cell(r, 17, f'=IF(C{r}="","",COUNTA(K{r}:P{r}))')
        ws.cell(r, 18, f'=_xlfn.TEXTJOIN("・",TRUE,K{r}:P{r})')
        ws.cell(r, 21, f'=IF(C{r}="","",COUNTIF({chk}!$B${CH_FIRST}:$B${CH_LAST},C{r}))')
        ws.cell(r, 22, f'=IF(OR(C{r}="",U{r}=0),"",INDEX({chk}!$AA${CH_FIRST}:$AA${CH_LAST},MATCH(C{r}&"#"&U{r},{chk}!$Z${CH_FIRST}:$Z${CH_LAST},0)))')
        ws.cell(r, 23, f'=IF(C{r}="","",IF(COUNTIFS({chk}!$B${CH_FIRST}:$B${CH_LAST},C{r},{chk}!$R${CH_FIRST}:$R${CH_LAST},"○")>0,"○",""))')
        for c in (9, 10, 22):
            ws.cell(r, c).number_format = "hh:mm"
    add_list(ws, f"K{YJ_FIRST}:P{YJ_LAST}", "=船名リスト", "予定表の曳船名をマスターの船名から選びます")
    add_list(ws, f"S{YJ_FIRST}:S{YJ_LAST}", "=選択_照合")
    add_list(ws, f"E{YJ_FIRST + len(jobs)}:E{YJ_LAST}", '"入,出,転,エスコート,当直,その他"')
    ws.freeze_panes = "C26"
    select(ws, "C7")
    return ws


# ---------------------------------------------------------------- 変更・キャンセル
def build_henkou(wb):
    ws = wb.create_sheet("変更・キャンセル")
    sheet_head(ws, "変更・キャンセル（全作業。1回の変更を1行。前日確定は上書きしない）",
               "誰が：上田　何を見て：当日の変更連絡・変更分の予定表・配船記録　どこへ：①日ごとの確認（A7～）②休船からの呼出（A15～）③変更の記録（A26～で作業を選ぶ）。実績3隻未満の日・全船休航の日も①へ入力",
               "集計先：「前日確定」の変更回数・最終時刻・取消し、「集計」の変更・取消し・呼出。実績のない取消しは「前日確定」の下の空き行へ作業を追加してから選びます。", 25)
    section(ws, 5, "■ 日ごとの変更確認（変更なしと確認できた日と未確認の日を区別）")
    for i, (t, k) in enumerate([("日付【入力】", "in"), ("その日の変更確認【入力・選択】", "in"), ("全船休航【入力・○】", "in"),
                                 ("実績の最大同時稼働（候補）【自動】", "auto"), ("情報元【入力】", "in"), ("確認者【入力】", "in")], 1):
        header(ws, 6, i, t, k)
    paint(ws, DAY_FIRST, DAY_LAST, 1, 3, "in")
    paint(ws, DAY_FIRST, DAY_LAST, 4, 4, "auto")
    paint(ws, DAY_FIRST, DAY_LAST, 5, 6, "in")
    ws.cell(DAY_FIRST, 1, TARGET)
    ws.cell(DAY_FIRST, 2, "未確認")
    for r in range(DAY_FIRST, DAY_LAST + 1):
        ws.cell(r, 1).number_format = "yyyy/mm/dd"
        ws.cell(r, 4, f'=IF(A{r}="","",IF(A{r}=集計!$C$5,集計!$C$14,"この試作では対象日のみ計算"))')
    add_list(ws, f"B{DAY_FIRST}:B{DAY_LAST}", "=選択_日確認")
    circle(ws, f"C{DAY_FIRST}:C{DAY_LAST}")

    section(ws, 13, "■ 休船からの呼出（製鉄曳船・西日本海運。1隻ごとに1行。同じ船の同じ呼出は1回と数える）")
    heads = [("呼出した船【入力・選択】", "in"), ("会社【自動】", "auto"), ("呼出日時【入力】", "in"),
             ("最初に対応した作業【入力・選択】", "in"), ("呼出の理由【入力・選択】", "in"), ("情報元【入力】", "in"),
             ("確認状況【入力・選択】", "in"), ("同一呼出の判定【自動】", "auto"), ("対象会社の判定【自動】", "auto")]
    for i, (t, k) in enumerate(heads, 1):
        header(ws, 14, i, t, k)
    for i, (t, k) in enumerate(heads, 1):
        paint(ws, YB_FIRST, YB_LAST, i, i, k)
    for r in range(YB_FIRST, YB_LAST + 1):
        ws.cell(r, 2, f'=IF(A{r}="","",IFERROR(INDEX(マスター!$C${M_FIRST}:$C${M_LAST},MATCH(A{r},マスター!$M${M_FIRST}:$M${M_LAST},0))&"","マスター未登録"))')
        ws.cell(r, 3).number_format = "yyyy/mm/dd hh:mm"
        ws.cell(r, 8, f'=IF(A{r}="","",IF(COUNTIFS(A${YB_FIRST}:A{r},A{r},C${YB_FIRST}:C{r},C{r})>1,"同一呼出（1回と数える）",""))')
        ws.cell(r, 9, f'=IF(A{r}="","",IF(OR(B{r}="製鉄曳船",B{r}="西日本海運"),"","製鉄曳船・西日本海運以外の船です"))')
    add_list(ws, f"A{YB_FIRST}:A{YB_LAST}", "=船名リスト")
    add_list(ws, f"D{YB_FIRST}:D{YB_LAST}", "=作業リスト")
    add_list(ws, f"E{YB_FIRST}:E{YB_LAST}", "=選択_呼出理由")
    add_list(ws, f"G{YB_FIRST}:G{YB_LAST}", "=選択_確認状況")

    section(ws, 24, "■ 変更・キャンセルの記録（作業を選ぶと本船名・前日確定・変更前が自動表示。2回目以降の変更も同じ作業を選んで行を追加）")
    heads = [("作業の選択【入力・選択】", "in", 34), ("作業キー【自動】", "auto", 14), ("本船名【自動】", "auto", 16),
             ("何回目の変更【自動】", "auto", 7), ("前日確定の時刻【自動】", "auto", 9), ("前日確定のタグ【自動】", "auto", 14),
             ("連絡日時【入力】", "in", 14), ("連絡元【入力】", "in", 10), ("変更前の時刻【自動】", "auto", 9),
             ("変更後の時刻【入力】", "in", 9), ("変更前のタグ【自動】", "auto", 14)] + \
            [(f"変更後のタグ{c}【入力・選択】", "in", 8) for c in "①②③④⑤⑥"] + \
            [("取消し【入力・○】", "in", 7), ("取消しの時点【入力・選択】", "in", 12), ("理由【入力・選択】", "in", 12),
             ("理由の詳細【入力】", "in", 16), ("時間調整【入力・選択】", "in", 8), ("時間調整の内容【入力】", "in", 16),
             ("情報元【入力】", "in", 14), ("確認者【入力】", "in", 8),
             ("照合キー【自動】", "auto", 8), ("変更後の時刻（有効）【自動】", "auto", 8), ("変更後のタグ（有効）【自動】", "auto", 8),
             ("前回の照合キー【自動】", "auto", 8)]
    for i, (t, k, w) in enumerate(heads, 1):
        header(ws, 25, i, t, k, w)
        paint(ws, CH_FIRST, CH_LAST, i, i, k)
    ws.row_dimensions[25].height = 60
    Z = "'前日確定'"
    for r in range(CH_FIRST, CH_LAST + 1):
        m = f'MATCH(A{r},{Z}!$B${YJ_FIRST}:$B${YJ_LAST},0)'
        ws.cell(r, 2, f'=IF(A{r}="","",IFERROR(INDEX({Z}!$C${YJ_FIRST}:$C${YJ_LAST},{m}),"作業が見つかりません"))')
        ws.cell(r, 3, f'=IF(A{r}="","",IFERROR(INDEX({Z}!$D${YJ_FIRST}:$D${YJ_LAST},{m}),""))')
        ws.cell(r, 4, f'=IF(B{r}="","",COUNTIF(B${CH_FIRST}:B{r},B{r}))')
        ws.cell(r, 5, f'=IF(A{r}="","",IFERROR(IF(INDEX({Z}!$J${YJ_FIRST}:$J${YJ_LAST},{m})="","",INDEX({Z}!$J${YJ_FIRST}:$J${YJ_LAST},{m})),""))')
        ws.cell(r, 6, f'=IF(A{r}="","",IFERROR(INDEX({Z}!$R${YJ_FIRST}:$R${YJ_LAST},{m}),""))')
        ws.cell(r, 7).number_format = "yyyy/mm/dd hh:mm"
        if r == CH_FIRST:  # 先頭行は必ず1回目（自分の行を参照しない＝循環参照を避ける）
            ws.cell(r, 9, f'=IF(A{r}="","",E{r})')
            ws.cell(r, 11, f'=IF(A{r}="","",F{r})')
        else:
            up = r - 1
            ws.cell(r, 9, f'=IF(A{r}="","",IF(D{r}=1,E{r},IFERROR(INDEX($AA${CH_FIRST}:$AA${up},MATCH(AC{r},$Z${CH_FIRST}:$Z${up},0)),"")))')
            ws.cell(r, 11, f'=IF(A{r}="","",IF(D{r}=1,F{r},IFERROR(INDEX($AB${CH_FIRST}:$AB${up},MATCH(AC{r},$Z${CH_FIRST}:$Z${up},0)),"")))')
        ws.cell(r, 26, f'=IF(B{r}="","",B{r}&"#"&D{r})')
        ws.cell(r, 27, f'=IF(A{r}="","",IF(J{r}="",I{r},J{r}))')
        ws.cell(r, 28, f'=IF(A{r}="","",IF(COUNTA(L{r}:Q{r})=0,K{r},_xlfn.TEXTJOIN("・",TRUE,L{r}:Q{r})))')
        ws.cell(r, 29, f'=IF(B{r}="","",B{r}&"#"&(D{r}-1))')
        for c in (5, 9, 10, 27):
            ws.cell(r, c).number_format = "hh:mm"
    add_list(ws, f"A{CH_FIRST}:A{CH_LAST}", "=作業リスト", "「前日確定」の作業から選びます")
    add_list(ws, f"L{CH_FIRST}:Q{CH_LAST}", "=船名リスト")
    circle(ws, f"R{CH_FIRST}:R{CH_LAST}")
    add_list(ws, f"S{CH_FIRST}:S{CH_LAST}", "=選択_取消時点")
    add_list(ws, f"T{CH_FIRST}:T{CH_LAST}", "=選択_理由")
    add_list(ws, f"V{CH_FIRST}:V{CH_LAST}", "=選択_時間調整")
    for c in range(26, 30):
        ws.column_dimensions[col(c)].hidden = True
    ws.freeze_panes = "B7"
    select(ws, "B7")
    wb.defined_names["作業リスト"] = DefinedName("作業リスト", attr_text=f"'前日確定'!$B${YJ_FIRST}:$B${YJ_LAST}")
    return ws


# ---------------------------------------------------------------- 入渠・他港作業・他社配置・標準時間
def simple_table(wb, name, title, who, dest, heads, nrows=50, first_input="A6"):
    ws = wb.create_sheet(name)
    sheet_head(ws, title, who, dest, len(heads) + 1)
    for i, h in enumerate(heads, 1):
        t, k, w = h[:3]
        header(ws, 5, i, t, k, w)
        paint(ws, 6, 5 + nrows, i, i, k, h[3] if len(h) > 3 else None)
    ws.row_dimensions[5].height = 45
    ws.freeze_panes = "A6"
    select(ws, first_input)
    return ws


def ship_company(r, c):
    return f'=IF({c}{r}="","",IFERROR(INDEX(マスター!$C${M_FIRST}:$C${M_LAST},MATCH({c}{r},マスター!$M${M_FIRST}:$M${M_LAST},0))&"","マスター未登録"))'


def build_others(wb):
    D = "yyyy/mm/dd"
    DT = "yyyy/mm/dd hh:mm"
    ws = simple_table(wb, "入渠", "入渠（作業に使えない期間）",
                      "誰が：上田（製鉄曳船）・西日本海運担当者（西日本海運）、他社は分かる範囲　何を見て：入渠の記録　どこへ：A6から1隻・1回ごとに1行",
                      "集計先：「集計」の対象日に入渠で使えない船数",
                      [("船名【入力・選択】", "in", 14), ("会社【自動】", "auto", 18), ("入渠開始日【入力】", "in", 12, D),
                       ("作業復帰日【入力】", "in", 12, D), ("情報元【入力】", "in", 24), ("確認状況【入力・選択】", "in", 14),
                       ("対象日に使えない【自動】", "auto", 10)])
    for r in range(6, 56):
        ws.cell(r, 2, ship_company(r, "A"))
        ws.cell(r, 7, f'=IF(OR(A{r}="",C{r}=""),"",IF(AND(集計!$C$5>=C{r},OR(D{r}="",集計!$C$5<D{r})),"○",""))')
    add_list(ws, "A6:A55", "=船名リスト")
    add_list(ws, "F6:F55", "=選択_確認状況")

    ws = simple_table(wb, "他港作業", "他港作業（西日本海運の全他港作業。関豊丸・長豊丸に限定しない）",
                      "誰が：西日本海運担当者　何を見て：西日本海運の作業記録　どこへ：A6から1作業ごとに1行。実時刻が分からない場合は「時刻の確認状況」で区別",
                      "集計先：「集計」の対象日に他港作業で拘束された船数",
                      [("船名【入力・選択】", "in", 14), ("会社【自動】", "auto", 14), ("作業港【入力】", "in", 12),
                       ("基地発日時【入力】", "in", 16, DT), ("基地着日時【入力】", "in", 16, DT),
                       ("時刻の確認状況【入力・選択】", "in", 12), ("情報元【入力】", "in", 20), ("確認者【入力】", "in", 10),
                       ("対象日に拘束【自動】", "auto", 9), ("会社の確認【自動】", "auto", 22)])
    for r in range(6, 56):
        ws.cell(r, 2, ship_company(r, "A"))
        ws.cell(r, 9, f'=IF(OR(A{r}="",D{r}=""),"",IF(AND(D{r}<集計!$C$5+1,OR(E{r}="",E{r}>集計!$C$5)),"○",""))')
        ws.cell(r, 10, f'=IF(A{r}="","",IF(B{r}="西日本海運","","西日本海運以外の船です（確認してください）"))')
    add_list(ws, "A6:A55", "=船名リスト")
    add_list(ws, "F6:F55", "=選択_時刻確認")

    ws = simple_table(wb, "他社配置", "他社配置（若松・関門等の他社の体制）",
                      "誰が：上田（分かる範囲）　何を見て：他社の配置・入渠・他港作業の情報　どこへ：A6から会社・船ごとに1行。配置の記録であり、応援の依頼に応じられたかとは別",
                      "集計先：「集計」の対象日の他社配置の記録数。2025年の配置は将来の応援確約ではありません。",
                      [("会社【入力・選択】", "in", 18), ("船名【入力・選択】", "in", 12), ("配置港【入力】", "in", 10),
                       ("対象期間の開始日【入力】", "in", 12, D), ("対象期間の終了日【入力】", "in", 12, D),
                       ("入渠・他港拘束（分かれば）【入力】", "in", 20), ("情報元【入力】", "in", 20), ("確認日【入力】", "in", 12, D),
                       ("対象日に配置【自動】", "auto", 9), ("船名と会社の一致【自動】", "auto", 20)])
    for r in range(6, 56):
        ws.cell(r, 9, f'=IF(OR(A{r}="",D{r}=""),"",IF(AND(集計!$C$5>=D{r},OR(E{r}="",集計!$C$5<=E{r})),"○",""))')
        ws.cell(r, 10, f'=IF(OR(A{r}="",B{r}=""),"",IF(IFERROR(INDEX(マスター!$C${M_FIRST}:$C${M_LAST},MATCH(B{r},マスター!$M${M_FIRST}:$M${M_LAST},0))&"","")=A{r},"","マスターの会社と異なります"))')
    add_list(ws, "A6:A55", "=会社リスト")
    add_list(ws, "B6:B55", "=船名リスト")

    ws = simple_table(wb, "標準時間", "標準時間（基地発から基地着までの見込み）",
                      "誰が：八幡タグ＝上田、西日本海運タグ＝西日本海運担当者　何を見て：運航の実態　どこへ：A6から会社・基地港・バース・作業内容ごとに1行（LNG入港・出港・警戒は分ける）",
                      "集計先：今後の標準時間による推計（実測とは区別して表示）。右の上田資料の目安は起点が未確認のため参考表示のみ。",
                      [("会社【入力・選択】", "in", 16), ("基地港【入力・選択】", "in", 10), ("バース【入力】", "in", 8),
                       ("作業内容【入力・選択】", "in", 10), ("予定時刻の基準【入力・選択】", "in", 10),
                       ("基地から現場への移動（分）【入力】", "in", 10), ("作業（分）【入力】", "in", 8), ("帰航（分）【入力】", "in", 8),
                       ("合計（分）【自動】", "auto", 8), ("適用開始日【入力】", "in", 11, D), ("適用終了日【入力】", "in", 11, D),
                       ("確認者【入力】", "in", 10), ("情報元【入力】", "in", 16), ("確認状況【入力・選択】", "in", 12)])
    for r in range(6, 56):
        ws.cell(r, 9, f'=IF(COUNT(F{r}:H{r})=0,"",SUM(F{r}:H{r}))')
    add_list(ws, "A6:A55", "=会社リスト")
    add_list(ws, "B6:B55", "=基地港リスト")
    add_list(ws, "D6:D55", "=選択_作業内容")
    add_list(ws, "E6:E55", "=選択_時刻基準")
    add_list(ws, "N6:N55", "=選択_確認状況")
    ref = [("区分", "平均使用隻数", "標準的な拘束時間", "最大の拘束時間"),
           ("LNG", "5.0隻", "入港 六連 回頭5.0H／連れ船4.0H、出港2.0H", "入港6.0H、出港2.5H"),
           ("副原料", "1.8隻", "入港2.5H、出港1.0H", "入港3.5H、出港1.0H"),
           ("安瀬石炭", "3.6隻", "入港 六連 回頭3.0H／連れ船2.5H／部埼3.0H、出港2.0H", "入港4.5H、出港2.0H"),
           ("内浦", "1.7隻", "入港2.0H、出港1.5H", "入港3.0H、出港2.0H"),
           ("八幡", "1.9隻", "入港1.5H、出港1.0H", "入港2.0H、出港2.0H"),
           ("鋼材シフト（日本製鉄作業）", "2.0隻", "シフト2.0H", "シフト2.5H")]
    ws.cell(4, 16, "参考：上田資料「1作業平均使用曳船隻数・曳船拘束時間集計」（基地出発～基地帰着の目安と記載。基地港・起点の詳細は未確認。入力欄ではありません）").font = font(bold=True)
    for n, row in enumerate(ref):
        for i, v in enumerate(row):
            c = ws.cell(5 + n, 16 + i, v)
            c.fill = F_HEAD_AUTO if n == 0 else F_AUTO
            c.border = BOX
            c.font = font(bold=(n == 0))
            c.alignment = Alignment(wrap_text=True, vertical="top")
    for c, w in zip(range(16, 20), (16, 9, 34, 18)):
        ws.column_dimensions[col(c)].width = w


# ---------------------------------------------------------------- _計算（非表示）
def build_calc(wb):
    ws = wb.create_sheet("_計算")
    ws["A1"] = "補助計算（非表示）：5分刻みで、マスターの船ごとに対象日の作業中（実績・取消しを除く・時刻のある行）なら1。同じ集計船名は1列にまとめる。"
    ws["A2"], ws["A3"], ws["A4"] = "会社区分", "集計船名", "当日稼働"
    ws["B1"] = "会社"
    J = lambda c: f"実績!${c}${JIS_FIRST}:${c}${JIS_LAST}"
    for k in range(NSHIP):
        c = GRID_COL0 + k
        L = col(c)
        mr = M_FIRST + k
        ws.cell(1, c, f'=IF(マスター!$N${mr}<>"○","",IF(マスター!$C${mr}="","未確認",マスター!$C${mr}))')
        ws.cell(2, c, f'=IF(マスター!$N${mr}<>"○","",マスター!$K${mr})')
        ws.cell(3, c, f'=IF(マスター!$N${mr}<>"○","",マスター!$J${mr})')
        ws.cell(4, c, f'=IF({L}$3="",0,IF(COUNTIFS({J("W")},{L}$3,{J("AE")},"○")>0,1,0))')
        for s in range(SLOTS):
            r = SLOT_FIRST + s
            ws.cell(r, c, f'=IF({L}$4=0,0,IF(COUNTIFS({J("W")},{L}$3,{J("AC")},"<="&$A{r},{J("AD")},">"&$A{r},{J("AE")},"○")>0,1,0))')
    ws.cell(5, 1, "分")
    ws.cell(5, 2, "時刻")
    for i, t in enumerate(["全体の同時隻数", "作業中の本船作業数", "製鉄曳船（会社要確認）", "西日本海運"]):
        ws.cell(5, COL_TOTAL + i, t)
    g1, g2 = col(GRID_COL0), col(GRID_LASTCOL)
    for s in range(SLOTS):
        r = SLOT_FIRST + s
        ws.cell(r, 1, s * 5)
        ws.cell(r, 2, f"=A{r}/1440").number_format = "hh:mm"
        ws.cell(r, COL_TOTAL, f"=SUM({g1}{r}:{g2}{r})")
        ws.cell(r, COL_JOBS, f'=SUMPRODUCT(({J("C")}="○")*ISNUMBER({J("AF")})*({J("AF")}<=$A{r})*({J("AG")}>$A{r}))')
        ws.cell(r, COL_STC, f'=SUMIF(${g1}$2:${g2}$2,"製鉄曳船",{g1}{r}:{g2}{r})')
        ws.cell(r, COL_NW, f'=SUMIF(${g1}$2:${g2}$2,"西日本海運",{g1}{r}:{g2}{r})')
    ws.sheet_state = "hidden"
    return ws


# ---------------------------------------------------------------- 集計
def build_summary(wb):
    ws = wb.create_sheet("集計")
    sheet_head(ws, "集計（自動表示・入力欄なし）",
               "誰が：全員が見る　何を見て：各シートの入力結果　どこへ：入力欄はありません。不足情報は「未確認」の数で確認します",
               "この試作は入力と集計のつながりを試すものです。3隻と4隻の比較の結論や、4隻の必要性を示すものではありません。", 14)
    calc = "_計算"
    T, JB, ST, NW = (f"{calc}!${col(c)}${SLOT_FIRST}:${col(c)}${SLOT_LAST}" for c in (COL_TOTAL, COL_JOBS, COL_STC, COL_NW))
    B = f"{calc}!$B${SLOT_FIRST}:$B${SLOT_LAST}"
    J = lambda c: f"実績!${c}${JIS_FIRST}:${c}${JIS_LAST}"
    Z = "'前日確定'"
    H = "'変更・キャンセル'"
    rows = [
        (5, "対象日", f"=実績!$D${JIS_FIRST}", "yyyy/mm/dd", "実績シートの先頭行の作業日"),
        (6, "時刻の扱い", "候補・時刻要確認", None, "元帳の作業時間が基地発・基地着か未確認のため、同時稼働は候補"),
        (8, "■ 実績（元帳）", None, None, None),
        (9, "作業件数", f'=COUNTIF({J("C")},"○")', None, "元帳の1作業（タグ別行ではない）"),
        (10, "延べタグ（タグ欄の行数）", f'=COUNTA({J("R")})', None, "取消し記載の行も含む"),
        (11, "うち取消しの記載がある行", f'=COUNTIF({J("U")},"○")', None, "同時稼働には含めない"),
        (12, "うち時刻がない行", f'=COUNTIF({J("V")},"時刻なし*")', None, "同時稼働には含めない"),
        (13, "同時稼働に含めた行", f'=COUNTIF({J("AE")},"○")', None, "取消し・時刻なし・通船等を除く"),
        (14, "最大同時稼働隻数（候補）", f"=MAX({T})", None, "同じ船は1隻。複数本船の作業を重ねた5分刻みの最大"),
        (15, "最初にその隻数となった時刻", f"=INDEX({B},MATCH(C14,{T},0))", "hh:mm", None),
        (16, "その時刻に作業中の本船作業数", f"=INDEX({JB},MATCH(C14,{T},0))", None, "1本船だけで3隻以上か、複数本船の重なりかを区別"),
        (17, "同時3隻以上の時間（分）", f'=COUNTIF({T},">=3")*5', None, None),
        (18, "　うち複数本船の重なり（分）", f"=SUMPRODUCT(({T}>=3)*({JB}>=2))*5", None, None),
        (19, "　うち1本船のみ（分）", f"=SUMPRODUCT(({T}>=3)*({JB}<2))*5", None, None),
        (20, "同時4隻以上の時間（分）", f'=COUNTIF({T},">=4")*5', None, None),
        (21, "同時5隻以上の時間（分）", f'=COUNTIF({T},">=5")*5', None, None),
        (22, "製鉄曳船（会社要確認）の最大同時", f"=MAX({ST})", None, "マスターで会社区分が「製鉄曳船」の船"),
        (23, "西日本海運の最大同時", f"=MAX({NW})", None, None),
        (24, "重複配船の警告がある行", f'=COUNTIF({J("AH")},"同じ*")', None, "同じ船が同じ時間に別作業"),
        (25, "マスター未登録の行", f'=COUNTIF({J("X")},"マスター未登録")', None, "マスターへ追加すると解消"),
        (26, "船名・会社が確認済みでない行", f'=COUNTA({J("R")})-COUNTIF({J("Z")},"確認済み")', None, "未確認として残している"),
        (28, "■ 会社区分別（実績）", None, None, None),
    ]
    for r, label, f, fmt, note in rows:
        ws.cell(r, 2, label).font = font(bold=label.startswith("■"))
        if label.startswith("■"):
            ws.cell(r, 2).font = font(size=11, bold=True, color="1F4E79")
            continue
        c = ws.cell(r, 3, f)
        paint(ws, r, r, 3, 3, "auto", fmt)
        if note:
            ws.cell(r, 4, note).font = font(color="595959")
    header(ws, 29, 2, "会社区分【自動】", "auto")
    header(ws, 29, 3, "延べタグ【自動】", "auto")
    header(ws, 29, 4, "稼働した船数【自動】", "auto")
    g1, g2 = col(GRID_COL0), col(GRID_LASTCOL)
    for n, k in enumerate(["製鉄曳船", "西日本海運", "他社", "未確認", "マスター未登録"]):
        r = 30 + n
        ws.cell(r, 2, k)
        ws.cell(r, 3, f'=COUNTIF({J("Y")},B{r})')
        ws.cell(r, 4, f'=SUMIF({calc}!${g1}$2:${g2}$2,B{r},{calc}!${g1}$4:${g2}$4)' if k != "マスター未登録" else "—")
        paint(ws, r, r, 2, 4, "auto")
    ws.cell(35, 2, "安瀬候補のタグ（按分配船・応援に含めない）")
    ws.cell(35, 3, f'=COUNTIFS({J("Y")},"他社",{J("AB")},"安瀬*")+COUNTIFS({J("Y")},"未確認",{J("AB")},"安瀬*")')
    ws.cell(36, 2, "他社・会社未確認のタグ（安瀬候補を除く）")
    ws.cell(36, 3, f'=COUNTIFS({J("Y")},"他社",{J("AB")},"")+COUNTIFS({J("Y")},"未確認",{J("AB")},"")')
    ws.cell(37, 2, "※ 安瀬候補は品名「支給炭」による仮の判定（要確認）。他社が実績で参加していないことを「応援不可」とは扱いません。").font = font(color="595959")
    paint(ws, 35, 36, 2, 3, "auto")

    section_rows = [
        (39, "■ 予定・出勤手配・同時稼働（それぞれ別の数値）", None, None, None),
        (40, "前日確定の照合状況", f"={Z}!$C$10", None, None),
        (41, "前日最終の予定時刻を入力した作業数", f"=COUNT({Z}!$J${YJ_FIRST}:$J${YJ_LAST})", None, None),
        (42, "予定のみ（実績なし）の作業数", f'=COUNTIF({Z}!$A${YJ_FIRST}:$A${YJ_LAST},"予定のみ")', None, None),
        (43, "作業予定タグ数（前日確定・延べ）", f"=SUM({Z}!$Q${YJ_FIRST}:$Q${YJ_LAST})", None, None),
        (44, "翌日の出勤手配隻数", f"={Z}!$C$22", None, None),
        (45, "実績の最大同時稼働（候補）", "=C14", None, None),
        (46, "同時に必要な隻数", "未入力（代表事例で別途確認）", None, "予定上の4隻出勤は4隻同時必要の証明ではない"),
        (48, "■ 変更・キャンセル", None, None, None),
        (49, "対象日の変更確認", f'=IFERROR(IF(INDEX({H}!$B${DAY_FIRST}:$B${DAY_LAST},MATCH(C5,{H}!$A${DAY_FIRST}:$A${DAY_LAST},0))="","未確認",INDEX({H}!$B${DAY_FIRST}:$B${DAY_LAST},MATCH(C5,{H}!$A${DAY_FIRST}:$A${DAY_LAST},0))),"未入力（未確認）")', None, None),
        (50, "全船休航", f'=IFERROR(IF(INDEX({H}!$C${DAY_FIRST}:$C${DAY_LAST},MATCH(C5,{H}!$A${DAY_FIRST}:$A${DAY_LAST},0))="○","○",""),"")', None, None),
        (51, "変更記録の行数", f"=COUNTA({H}!$A${CH_FIRST}:$A${CH_LAST})", None, None),
        (52, "変更のあった作業数", f'=SUMPRODUCT(({H}!$B${CH_FIRST}:$B${CH_LAST}<>"")/COUNTIF({H}!$B${CH_FIRST}:$B${CH_LAST},{H}!$B${CH_FIRST}:$B${CH_LAST}&""))', None, None),
        (53, "取消しになった作業数", f'=COUNTIF({Z}!$W${YJ_FIRST}:$W${YJ_LAST},"○")', None, None),
        (54, "うち基地発後の取消し（記録数）", f'=COUNTIFS({H}!$R${CH_FIRST}:$R${CH_LAST},"○",{H}!$S${CH_FIRST}:$S${CH_LAST},"当日・基地発後")', None, "実拘束あり。実施作業へは数えない"),
        (55, "時間調整ありの記録", f'=COUNTIF({H}!$V${CH_FIRST}:$V${CH_LAST},"あり")', None, None),
        (56, "休船からの呼出回数（同一呼出は1回）", f'=COUNTA({H}!$A${YB_FIRST}:$A${YB_LAST})-COUNTIF({H}!$H${YB_FIRST}:$H${YB_LAST},"同一呼出*")', None, "呼出実績＝常駐船腹の不足とは限らない"),
        (57, "　うち乗組員の出勤手配の不足による呼出", f'=COUNTIFS({H}!$E${YB_FIRST}:$E${YB_LAST},"乗組員の出勤手配の不足",{H}!$H${YB_FIRST}:$H${YB_LAST},"")', None, "船体の不足と区別"),
        (58, "呼出の対象外の会社の記録", f'=COUNTIF({H}!$I${YB_FIRST}:$I${YB_LAST},"製鉄曳船*")', None, None),
        (60, "■ 使えない船・拘束・他社配置（対象日）", None, None, None),
        (61, "入渠で使えない船数", '=COUNTIF(入渠!$G$6:$G$55,"○")', None, None),
        (62, "他港作業で拘束された船数", '=COUNTIF(他港作業!$I$6:$I$55,"○")', None, None),
        (63, "他社配置の記録数", '=COUNTIF(他社配置!$I$6:$I$55,"○")', None, "配置の記録であり応援の確約ではない"),
    ]
    for r, label, f, fmt, note in section_rows:
        if label.startswith("■"):
            ws.cell(r, 2, label).font = font(size=11, bold=True, color="1F4E79")
            continue
        ws.cell(r, 2, label).font = font()
        ws.cell(r, 3, f)
        paint(ws, r, r, 3, 3, "auto", fmt)
        if note:
            ws.cell(r, 4, note).font = font(color="595959")

    # 時間帯別
    ws.cell(8, 6, "■ 時間帯別の最大同時稼働（候補）").font = font(size=11, bold=True, color="1F4E79")
    for i, t in enumerate(["時台", "全体【自動】", "本船作業数【自動】", "製鉄曳船（要確認）【自動】", "西日本海運【自動】"]):
        header(ws, 9, 6 + i, t, "auto")
    for h in range(24):
        r = 10 + h
        a, b = SLOT_FIRST + h * 12, SLOT_FIRST + h * 12 + 11
        ws.cell(r, 6, f"{h}時")
        for i, cc in enumerate((COL_TOTAL, COL_JOBS, COL_STC, COL_NW)):
            L = col(cc)
            ws.cell(r, 7 + i, f"=MAX({calc}!{L}{a}:{L}{b})")
        paint(ws, r, r, 6, 10, "auto")

    # 会社別（マスターの会社一覧）
    ws.cell(8, 12, "■ 会社別（マスターの会社一覧から自動）").font = font(size=11, bold=True, color="1F4E79")
    for i, t in enumerate(["会社【自動】", "会社区分【自動】", "延べタグ【自動】", "稼働した船数【自動】"]):
        header(ws, 9, 12 + i, t, "auto")
    for n in range(C_LAST - C_FIRST + 1):
        r = 10 + n
        mr = C_FIRST + n
        ws.cell(r, 12, f'=IF(マスター!$P${mr}="","",マスター!$P${mr})')
        ws.cell(r, 13, f'=IF(L{r}="","",マスター!$Q${mr})')
        ws.cell(r, 14, f'=IF(L{r}="","",COUNTIF({J("X")},L{r}))')
        ws.cell(r, 15, f'=IF(L{r}="","",SUMIF({calc}!${g1}$1:${g2}$1,L{r},{calc}!${g1}$4:${g2}$4))')
        paint(ws, r, r, 12, 15, "auto")
    r = 10 + C_LAST - C_FIRST + 1
    ws.cell(r, 12, "会社未確認")
    ws.cell(r, 14, f'=COUNTIF({J("X")},"未確認")')
    ws.cell(r, 15, f'=SUMIF({calc}!${g1}$1:${g2}$1,"未確認",{calc}!${g1}$4:${g2}$4)')
    paint(ws, r, r, 12, 15, "auto")
    for c, w in zip("ABCDEFGHIJKLMNO", (2, 34, 14, 40, 2, 7, 9, 9, 11, 9, 2, 24, 10, 9, 10)):
        ws.column_dimensions[c].width = w
    select(ws, "C14")
    return ws


# ---------------------------------------------------------------- はじめに
def build_intro(wb):
    ws = wb.create_sheet("はじめに", 0)
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 100
    ws["B1"] = "八幡地区 常駐体制検証 入力試作 v01"
    ws["B1"].font = font(size=16, bold=True)
    lines = [
        ("目的", "製鉄曳船の3隻と4隻で対応の差を説明するための基礎データを集めます。この試作は、入力と集計のつながりを試すためのものです。4隻の必要性を証明するものではありません。"),
        ("対象", f"元帳R7年度のうち、同時3隻以上の候補日 {TARGET:%Y年%m月%d日} の1日分（実績17行・9作業）。元帳の時刻は「候補・時刻要確認」です。"),
        ("最初の操作", "下の青い文字を押すと、最初の入力欄（前日確定シートC7：根拠の予定表）へ移動します。"),
        ("入力の順番", "①前日確定 → ②変更・キャンセル → ③入渠・他港作業・他社配置 → ④標準時間。船名や会社が選択肢にないときは、先にマスターへ追加します。"),
        ("担当者", "前日確定・変更・キャンセル・製鉄曳船の入渠・八幡タグの標準時間：上田さん／西日本海運の他港作業・入渠・標準時間：西日本海運担当者／他社配置：分かる範囲で上田さん"),
        ("色の意味", "薄い黄色＝入力する欄（見出しに【入力】）／薄い灰色＝自動表示（見出しに【自動】）。灰色の欄は入力しません。"),
        ("選び方", "船名・会社・理由などはセルを選ぶと出る▼の一覧から選びます。○の欄は▼から○を選びます。ボタンやマクロは使っていません。"),
        ("未確認の意味", "資料で確かめられていないことです。推測で埋めず「未確認」のまま残します。確認できたら「確認済み」に変えます。"),
        ("再入力しない", "本船名・タグ名・実績時刻は元帳から自動表示します。作業番号の入力は不要で、作業は▼の一覧から選びます。"),
        ("前日確定と変更", "前日最終の予定は「前日確定」に入れ、変更があっても上書きしません。変更は「変更・キャンセル」に1回ごとに1行追加します。"),
        ("実績のない取消し", "予定表にあって元帳にない作業は、「前日確定」の下の空き行に本船名などを入れて追加し、「変更・キャンセル」で選んで取消しを入力します。"),
        ("安瀬", "安瀬は若松との按分配船です。応援とは別に数えます（品名「支給炭」を安瀬候補として仮に表示。要確認）。"),
        ("集計", "「集計」シートで、作業予定タグ数・出勤手配隻数・実績の同時稼働を別々に表示します。同時に必要な隻数は代表事例で別に確認します。"),
    ]
    for n, (k, v) in enumerate(lines):
        r = 3 + n
        ws.cell(r, 2, k).font = font(bold=True)
        ws.cell(r, 3, v).font = font()
        ws.cell(r, 3).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(r, 2).alignment = Alignment(vertical="top")
        ws.row_dimensions[r].height = 30
    r = 3 + len(lines) + 1
    link = ws.cell(r, 2, "▶ 最初の入力へ移動（前日確定 C7）")
    link.hyperlink = Hyperlink(ref=link.coordinate, location="'前日確定'!C7", display=link.value)
    link.font = font(size=12, bold=True, color="0563C1", underline="single")
    ws.cell(r, 3, "対象日の前日最終の予定表のファイル名・版を入力します（2025年度分は資料一式に含まれていないため、上田さんの照合で入力）。").font = font()
    r += 2
    ws.cell(r, 2, "色見本").font = font(bold=True)
    ws.cell(r, 3, "入力する欄【入力】").fill = F_IN
    ws.cell(r + 1, 3, "自動表示の欄【自動】").fill = F_AUTO
    r += 3
    ws.cell(r, 2, "シートの役割").font = font(bold=True)
    for n, (s, d) in enumerate([("実績", "元帳の取り込み（見るだけ）"), ("前日確定", "前日最終の予定・予定タグ・翌日の出勤手配"),
                                 ("変更・キャンセル", "日ごとの変更確認、休船からの呼出、変更・取消しの記録"),
                                 ("入渠", "作業に使えない期間"), ("他港作業", "西日本海運の全他港作業"), ("他社配置", "若松・関門等の他社の体制"),
                                 ("標準時間", "基地発～基地着の見込み時間"), ("マスター", "会社・船名・基地港・選択肢の追加"),
                                 ("集計", "対比結果と不足情報（自動）")]):
        c = ws.cell(r + 1 + n, 2, s)
        c.hyperlink = Hyperlink(ref=c.coordinate, location=f"'{s}'!A1", display=s)
        c.font = font(color="0563C1", underline="single")
        ws.cell(r + 1 + n, 3, d).font = font()
    select(ws, f"B{3 + len(lines) + 1}")


def main():
    rows = read_target_rows()
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    build_jisseki_ws, jobs = build_jisseki(wb, rows)
    build_zenjitsu(wb, jobs)
    build_henkou(wb)
    build_others(wb)
    build_master(wb)
    build_summary(wb)
    build_calc(wb)
    build_intro(wb)
    order = ["はじめに", "実績", "前日確定", "変更・キャンセル", "入渠", "他港作業", "他社配置", "標準時間", "マスター", "集計", "_計算"]
    wb._sheets = [wb[n] for n in order]
    wb.active = 0
    for ws in wb.worksheets:
        ws.sheet_view.tabSelected = ws.title == "はじめに"
    wb.calculation.fullCalcOnLoad = True
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    wb.save(OUT)
    print("saved", OUT, "rows", len(rows), "jobs", len(jobs))


if __name__ == "__main__":
    main()
