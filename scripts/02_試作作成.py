"""八幡地区_検証試作_v01.xlsx を作成する（実データ1日分の小規模試作）。

- 元帳（input/01_元資料/作業実績…xlsx のシート「R7年度」）から対象日1日分を読み取り、「実績」シートへ取り込む。
- 元資料は読み取りのみ。入力欄は空欄（または「未確認」）で作成し、推測で埋めない。
- 試験専用コピーは scripts/03_試作検証.py が build() に架空の実績行を渡して作る（納品ファイルには入れない）。
- 計算条件は output/計算条件_v01.md に記載。

使い方: python3 scripts/02_試作作成.py
"""
import datetime as dt
import glob
import os

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import column_index_from_string as cidx
from openpyxl.utils import get_column_letter as col
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.hyperlink import Hyperlink

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "output", "八幡地区_検証試作_v01.xlsx")

# 対象日：R7年度で同時3隻以上の候補日のうち、複数本船（LNG入港を含む）と他社タグが重なる日。
# 上田資料「最大同時稼働隻数」に記載があり、資料の値（9隻）と今回計算（8隻）が異なるため照合対象として選んだ。
TARGET = dt.date(2025, 12, 6)

# ---------------------------------------------------------------- 容量（行数）
JIS_FIRST, JIS_LAST = 6, 505          # 実績
M_FIRST, M_LAST = 6, 125              # マスター（船）
C_FIRST, C_LAST = 6, 45               # マスター（会社）
L_FIRST, L_LAST = 6, 25               # マスター（基地港・品名・選択肢）
SLOT_FIRST, SLOTS = 6, 288            # _計算：5分刻み×24時間
SLOT_LAST = SLOT_FIRST + SLOTS - 1
GRID_COL0 = 3                         # _計算：船の列の開始（C列）
NSHIP = M_LAST - M_FIRST + 1
GRID_LASTCOL = GRID_COL0 + NSHIP - 1
COL_TOTAL, COL_JOBS, COL_STC, COL_NW = (GRID_LASTCOL + i for i in range(1, 5))
YJ_FIRST, YJ_LAST = 26, 45            # 前日確定：作業行
CH_FIRST, CH_LAST = 26, 85            # 変更・キャンセル：変更行
YB_FIRST, YB_LAST = 15, 22            # 変更・キャンセル：休船呼出
DAY_FIRST, DAY_LAST = 7, 11           # 変更・キャンセル：日ごとの確認
OT_FIRST, OT_LAST = 6, 55             # 入渠・他港作業・他社配置・標準時間
SUP_FIRST = 22                        # 集計：供給可能船の表

# ---------------------------------------------------------------- 列の位置
# 実績（A～Vは元帳からの取込値、W以降は式）
J_KEY, J_LABEL, J_FIRSTROW, J_DATE, J_ROWNO = "A", "B", "C", "D", "E"
J_IO, J_HIN, J_START, J_END, J_CROSS, J_TUG, J_CANCEL, J_TIMING = "J", "M", "O", "P", "Q", "R", "U", "V"
J_NAME, J_COMP, J_KUBUN, J_CONF, J_KIND, J_AZE, J_SAGYO, J_KOURO = "W", "X", "Y", "Z", "AA", "AB", "AC", "AD"
J_SDT, J_EDT, J_PREV, J_WARN = "AE", "AF", "AG", "AH"
J_SMIN, J_EMIN, J_INC, J_INC2, J_JSMIN, J_JEMIN = "AI", "AJ", "AK", "AL", "AM", "AN"
# マスター（船）
M_SEL, M_FORMAL, M_COMP, M_KIND, M_BASE, M_SUPPLY = "A", "B", "C", "D", "E", "F"
M_FROM, M_TO, M_SRC, M_CONF, M_NAME, M_KUBUN, M_VALID, M_KEY, M_HEAD = "G", "H", "I", "J", "K", "L", "M", "N", "O"
# マスター（会社・基地港・品名）
MC_NAME, MC_KUBUN, MC_SRC, MC_CONF = "Q", "R", "S", "T"
MB_PORT = "V"
MH_NAME, MH_KOURO, MH_AZE, MH_SRC, MH_CONF = "X", "Y", "Z", "AA", "AB"
CHOICE_COL0 = cidx("AD")

FONT = "游ゴシック"
F_IN = PatternFill("solid", fgColor="FFF9C4")    # 薄黄色＝入力
F_AUTO = PatternFill("solid", fgColor="EDEDED")  # 薄灰色＝自動
F_HEAD_IN = PatternFill("solid", fgColor="FFE97F")
F_HEAD_AUTO = PatternFill("solid", fgColor="BFBFBF")
F_TITLE = PatternFill("solid", fgColor="DDEBF7")
THIN = Side(style="thin", color="A6A6A6")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
FMT_DT = "[<1]hh:mm;m/d hh:mm"   # 時刻だけなら hh:mm、日付つきなら m/d hh:mm


def font(**kw):
    return Font(name=FONT, size=kw.pop("size", 10), **kw)


def fy_date(m, d):
    return dt.date(2025 if m >= 4 else 2026, m, d)


def rng(sheet, c, first, last):
    return f"{sheet}!${c}${first}:${c}${last}"


def JR(c):
    return rng("実績", c, JIS_FIRST, JIS_LAST)


def MR(c):
    return rng("マスター", c, M_FIRST, M_LAST)


def lookup_master(key_expr, c, default):
    """選択名（対象日に有効な行）で船マスターを引く。"""
    return f'IFERROR(INDEX({MR(c)},MATCH({key_expr},{MR(M_KEY)},0)),{default})'


# ---------------------------------------------------------------- 元帳の読み取り
def read_target_rows(target=TARGET):
    path = glob.glob(os.path.join(ROOT, "input", "01_元資料", "作業実績*"))[0]
    ws = openpyxl.load_workbook(path, read_only=True)["R7年度"]
    rows, prev_cross = [], []
    for i, r in enumerate(ws.iter_rows(values_only=True), 1):
        if i == 1 or r[0] is None or r[17] is None:
            continue
        d = fy_date(r[0], r[1])
        if d == target:
            rows.append((i, r))
        elif d == target - dt.timedelta(1):
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
    ws["A3"] = dest
    ws["A2"].font = font()
    ws["A3"].font = font(color="595959")
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


def section(ws, row, text, c=1):
    ws.cell(row, c, text).font = font(size=11, bold=True, color="1F4E79")


def add_list(ws, ref, source, prompt=None):
    dv = DataValidation(type="list", formula1=source, allow_blank=True, showDropDown=False)
    dv.error = "一覧から選んでください（一覧にない場合は「マスター」シートへ追加してください）"
    dv.errorTitle = "選択してください"
    if prompt:
        dv.prompt = prompt
        dv.showInputMessage = True
    ws.add_data_validation(dv)
    dv.add(ref)


def circle(ws, ref):
    add_list(ws, ref, '"○"', "○を選びます（空欄＝該当なし）")


def select(ws, ref):
    ws.sheet_view.selection[0].activeCell = ref
    ws.sheet_view.selection[0].sqref = ref


def table(ws, row, heads, first, last):
    """heads: [(見出し, 'in'|'auto', 幅, 書式)]"""
    for i, h in enumerate(heads, 1):
        header(ws, row, i, h[0], h[1], h[2])
        paint(ws, first, last, i, i, h[1], h[3] if len(h) > 3 else None)
    ws.row_dimensions[row].height = 60


# ---------------------------------------------------------------- マスター
MASTER_SHIPS = [
    # 選択名, 正式船名, 会社, 種類, 基地港, 供給数に数える, 情報元, 確認状況
    ("八幡丸", "八幡丸", "製鉄曳船", "タグ", "未確認", "", "00_作成条件.md", "要確認（基地港・元帳記号との対応）"),
    ("鐵豊丸", "鐵豊丸", "製鉄曳船", "タグ", "未確認", "", "00_作成条件.md", "要確認（基地港・元帳記号との対応）"),
    ("新豊丸", "新豊丸", "製鉄曳船", "タグ", "未確認", "", "00_作成条件.md", "要確認（基地港・元帳記号との対応）"),
    ("八豊丸", "八豊丸", "製鉄曳船", "タグ", "未確認", "", "00_作成条件.md", "要確認（基地港・元帳記号との対応）"),
    ("松豊丸", "松豊丸", "製鉄曳船", "タグ", "未確認", "", "00_作成条件.md", "要確認（基地港・元帳記号との対応）"),
    ("T", "", "製鉄曳船", "タグ", "未確認", "○", "元帳R7年度の記号。会社は既存集計（西日本G 2,744）との照合上の扱い", "要確認（正式船名・会社）"),
    ("S", "", "製鉄曳船", "タグ", "未確認", "○", "同上", "要確認（正式船名・会社）"),
    ("Y", "", "製鉄曳船", "タグ", "未確認", "○", "同上", "要確認（正式船名・会社）"),
    ("M", "", "製鉄曳船", "タグ", "未確認", "○", "同上", "要確認（正式船名・会社）"),
    ("P", "", "製鉄曳船", "タグ", "未確認", "○", "同上", "要確認（正式船名・会社）"),
    ("関豊丸", "関豊丸", "西日本海運", "タグ", "門司港", "", "00_作成条件.md／要目表2026.8", "要確認（元帳記号との対応）"),
    ("長豊丸", "長豊丸", "西日本海運", "タグ", "門司港", "", "00_作成条件.md／要目表2026.8", "要確認（元帳記号との対応）"),
    ("関", "", "西日本海運", "タグ", "未確認", "○", "元帳R7年度の記号。会社は既存集計との照合上の扱い", "要確認（正式船名・会社）"),
    ("長", "", "西日本海運", "タグ", "未確認", "○", "同上", "要確認（正式船名・会社）"),
    ("かざし丸", "かざし丸", "シーゲートコーポレーション", "タグ", "門司港", "", "要目表2026.8（2025年度の所属は未確認）", "要確認（2025年度の所属）"),
    ("鷲羽丸", "鷲羽丸", "シーゲートコーポレーション", "タグ", "門司港", "", "要目表2026.8（2025年度の所属は未確認）", "要確認（2025年度の所属）"),
    ("むさし丸2", "むさし丸2", "グリーンシッピング", "タグ", "門司港", "", "要目表2026.8（2025年度の所属は未確認）", "要確認（2025年度の所属）"),
    ("かなで", "かなで", "グリーンシッピング", "タグ", "門司港", "", "要目表2026.8（2025年度の所属は未確認）", "要確認（2025年度の所属）"),
    ("かいせい", "かいせい", "日鉄物流", "タグ", "門司港", "", "要目表2026.8（2025年度の所属は未確認）", "要確認（2025年度の所属）"),
    ("春風", "春風", "春風海運", "タグ", "門司港", "", "00_作成条件.md（ユーザー指定）", "確認済み"),
    ("栄春", "栄春", "春風海運", "タグ", "門司港", "", "00_作成条件.md（ユーザー指定）", "確認済み"),
    ("双美", "双美", "春風海運", "タグ", "門司港", "", "00_作成条件.md（ユーザー指定）", "確認済み"),
    ("春幸", "春幸", "春風海運", "タグ", "門司港", "", "00_作成条件.md（ユーザー指定）", "確認済み"),
    ("さくら", "", "", "タグ", "未確認", "", "元帳R7年度（S列「洞」は参考。会社は推定しない）", "要確認（正式船名・会社）"),
    ("みさき", "", "", "タグ", "未確認", "", "元帳R7年度（S列「洞」は参考。会社は推定しない）", "要確認（正式船名・会社）"),
    ("竜山", "", "", "タグ", "未確認", "", "元帳R7年度（S列「矢」は参考。会社は推定しない）", "要確認（正式船名・会社）"),
    ("竜峰", "", "", "タグ", "未確認", "", "元帳R7年度（S列「矢」は参考。会社は推定しない）", "要確認（正式船名・会社）"),
    ("竜昇", "", "", "タグ", "未確認", "", "元帳R7年度（S列「矢」は参考。会社は推定しない）", "要確認（正式船名・会社）"),
    ("竜海", "", "", "タグ", "未確認", "", "元帳R7年度（S列「矢」は参考。会社は推定しない）", "要確認（正式船名・会社）"),
    ("慶鳳丸", "", "", "タグ", "未確認", "", "元帳R7年度", "要確認（正式船名・会社）"),
    ("秀豊", "", "", "タグ", "未確認", "", "元帳R7年度（秀豊丸と同一船かは未確認）", "要確認（正式船名・会社）"),
    ("秀豊丸", "", "", "タグ", "未確認", "", "元帳R7年度（秀豊と同一船かは未確認）", "要確認（正式船名・会社）"),
    ("まんじゅ", "", "", "未確認", "未確認", "", "元帳R7年度。LNGエスコートのみ・時刻空欄。既存集計では除外とみられる", "要確認（種類・会社）"),
    ("せいじゅ", "", "", "未確認", "未確認", "", "元帳R7年度。LNGエスコートのみ・時刻空欄。既存集計では除外とみられる", "要確認（種類・会社）"),
]
MASTER_COMPANIES = [
    ("製鉄曳船", "製鉄曳船", "00_作成条件.md", "確認済み"),
    ("西日本海運", "西日本海運", "00_作成条件.md", "確認済み"),
    ("シーゲートコーポレーション", "他社", "00_作成条件.md／要目表2026.8", "未確認"),
    ("グリーンシッピング", "他社", "00_作成条件.md／要目表2026.8", "未確認"),
    ("日鉄物流", "他社", "00_作成条件.md／要目表2026.8", "未確認"),
    ("春風海運", "他社", "00_作成条件.md", "確認済み"),
    ("矢野海運", "他社", "00_作成条件.md（船名は未確認）", "未確認"),
    ("洞海マリン", "他社", "00_作成条件.md（船名は未確認）", "未確認"),
    ("福島海運", "他社", "00_作成条件.md（船名は未確認）", "未確認"),
]
BASE_PORTS = ["門司港", "戸畑", "八幡", "若松", "未確認"]
HINMEI = [  # 品名, 高炉原料区分, 安瀬, 情報元
    ("鋼材", "高炉原料以外", "", "初期値（要確認）"),
    ("LNG", "高炉原料以外", "", "初期値（要確認）"),
    ("鉱石", "高炉原料", "", "初期値（要確認）"),
    ("石炭", "高炉原料", "", "初期値（要確認）。発電用等の区別は未確認"),
    ("支給炭", "未確認", "○", "既存集計の安瀬34と支給炭入港34が一致（要確認）"),
    ("副原料", "未確認", "", "将来の取扱い未確認"),
    ("内航", "未確認", "", "未確認"),
    ("その他", "未確認", "", "未確認"),
]
CHOICES = [  # 見出し, 名前, 値
    ("確認状況", "選択_確認状況", ["未確認", "確認済み", "要確認（正式船名・会社）", "要確認（基地港・元帳記号との対応）", "要確認（元帳記号との対応）", "要確認（2025年度の所属）", "要確認（種類・会社）"]),
    ("照合の状況", "選択_照合", ["未確認", "予定表と照合済み", "予定表に記載なし", "予定表なし"]),
    ("取消しの時点", "選択_取消時点", ["前日", "当日・基地発前", "当日・基地発後", "不明"]),
    ("変更・取消しの理由", "選択_理由", ["強風・荒天", "視界不良", "本船都合", "荷役都合", "代理店・荷主都合", "バース都合", "タグの都合", "その他", "不明"]),
    ("時間調整", "選択_時間調整", ["あり", "なし", "不明"]),
    ("呼出の理由", "選択_呼出理由", ["常駐船が他作業で不足", "入渠中の船の代わり", "他港作業中の船の代わり", "乗組員の出勤手配の不足", "その他", "不明"]),
    ("作業内容", "選択_作業内容", ["入港", "出港", "転錨・シフト", "LNG入港", "LNG出港", "LNG警戒", "LNGエスコート", "その他"]),
    ("予定時刻の基準", "選択_時刻基準", ["本船POB", "着岸", "離岸", "作業開始", "基地発", "未確認"]),
    ("船の種類", "選択_種類", ["タグ", "通船・警戒船", "未確認"]),
    ("時刻の確認状況", "選択_時刻確認", ["実時刻", "予定・推定", "不明"]),
    ("会社区分", "選択_会社区分", ["製鉄曳船", "西日本海運", "他社", "未確認"]),
    ("その日の変更確認", "選択_日確認", ["未確認", "変更なしと確認済み", "変更あり（下の表に入力）"]),
    ("高炉原料区分", "選択_高炉区分", ["高炉原料", "高炉原料以外", "未確認"]),
    ("実際の出勤・待機", "選択_出勤", ["出勤して待機", "出勤して作業", "出勤取消し", "不明"]),
    ("依頼への対応可否", "選択_対応可否", ["可", "不可"]),
]


def build_master(wb):
    ws = wb.create_sheet("マスター")
    sheet_head(ws, "マスター（会社・船名・基地港・品名・選択肢）",
               "誰が：上田・各担当者　何を見て：要目表・元帳・予定表　どこへ：各表の最後の行の下の空き行へ追加（空行を空けずに続けて入力）",
               "集計先：全シートの選択肢（▼）、実績の会社判定、集計の会社別・供給可能船・高炉原料を除く集計。正式船名を入れると同じ船として1隻に数えます。", 50)
    heads = [("選択名（元帳・予定表の表記）【入力】", "in", 14), ("正式船名【入力】", "in", 11), ("会社【入力・選択】", "in", 18),
             ("船の種類【入力・選択】", "in", 9), ("基地港【入力・選択】", "in", 9), ("常駐側の供給数に数える【入力・○】", "in", 9),
             ("適用開始日【入力】", "in", 11, "yyyy/mm/dd"), ("適用終了日【入力】", "in", 11, "yyyy/mm/dd"), ("情報元【入力】", "in", 28),
             ("確認状況【入力・選択】", "in", 18), ("集計船名【自動】", "auto", 10), ("会社区分【自動】", "auto", 9),
             ("対象日に有効【自動】", "auto", 7), ("照合名【自動】", "auto", 7), ("同一船の先頭【自動】", "auto", 7)]
    table(ws, 5, heads, M_FIRST, M_LAST)
    for n, (name, formal, comp, kind, base, sup, src, conf) in enumerate(MASTER_SHIPS):
        r = M_FIRST + n
        for c, v in zip((M_SEL, M_FORMAL, M_COMP, M_KIND, M_BASE, M_SUPPLY, M_SRC, M_CONF),
                        (name, formal, comp, kind, base, sup, src, conf)):
            ws[f"{c}{r}"] = v or None
    for r in range(M_FIRST, M_LAST + 1):
        ws[f"{M_NAME}{r}"] = f'=IF({M_SEL}{r}="","",IF({M_FORMAL}{r}="",{M_SEL}{r},{M_FORMAL}{r}))'
        ws[f"{M_KUBUN}{r}"] = (f'=IF({M_SEL}{r}="","",IF({M_COMP}{r}="","未確認",IFERROR(INDEX({rng("マスター", MC_KUBUN, C_FIRST, C_LAST)},'
                               f'MATCH({M_COMP}{r},{rng("マスター", MC_NAME, C_FIRST, C_LAST)},0)),"未確認")))')
        ws[f"{M_VALID}{r}"] = (f'=IF({M_SEL}{r}="","",IF(AND(OR({M_FROM}{r}="",{M_FROM}{r}<=集計!$C$5),'
                               f'OR({M_TO}{r}="",{M_TO}{r}>=集計!$C$5)),"○",""))')
        ws[f"{M_KEY}{r}"] = f'=IF({M_VALID}{r}="○",{M_SEL}{r},"")'
        ws[f"{M_HEAD}{r}"] = (f'=IF(OR({M_NAME}{r}="",{M_VALID}{r}<>"○"),"",'
                              f'IF(COUNTIFS({M_NAME}${M_FIRST}:{M_NAME}{r},{M_NAME}{r},{M_VALID}${M_FIRST}:{M_VALID}{r},"○")=1,"○",""))')
    # 会社
    c0 = cidx(MC_NAME)
    for i, (t, k, w) in enumerate([("会社名【入力】", "in", 22), ("会社区分【入力・選択】", "in", 9),
                                     ("情報元【入力】", "in", 22), ("確認状況【入力・選択】", "in", 10)]):
        header(ws, 5, c0 + i, t, k, w)
    paint(ws, C_FIRST, C_LAST, c0, c0 + 3, "in")
    for n, row in enumerate(MASTER_COMPANIES):
        for i, v in enumerate(row):
            ws.cell(C_FIRST + n, c0 + i, v)
    # 基地港
    header(ws, 5, cidx(MB_PORT), "基地港【入力】", "in", 9)
    paint(ws, L_FIRST, L_LAST, cidx(MB_PORT), cidx(MB_PORT), "in")
    for n, v in enumerate(BASE_PORTS):
        ws[f"{MB_PORT}{L_FIRST + n}"] = v
    # 品名
    h0 = cidx(MH_NAME)
    for i, (t, w) in enumerate([("品名（元帳の表記）【入力】", 9), ("高炉原料区分【入力・選択】", 11), ("安瀬（按分配船）【入力・○】", 8),
                                 ("情報元【入力】", 22), ("確認状況【入力・選択】", 10)]):
        header(ws, 5, h0 + i, t, "in", w)
    paint(ws, L_FIRST, L_LAST, h0, h0 + 4, "in")
    for n, (h, k, a, src) in enumerate(HINMEI):
        r = L_FIRST + n
        ws[f"{MH_NAME}{r}"], ws[f"{MH_KOURO}{r}"], ws[f"{MH_AZE}{r}"] = h, k, a or None
        ws[f"{MH_SRC}{r}"], ws[f"{MH_CONF}{r}"] = src, "未確認"
    # 選択肢
    for n, (title, _, values) in enumerate(CHOICES):
        c = CHOICE_COL0 + n
        header(ws, 5, c, title + "【入力】", "in", 13)
        paint(ws, L_FIRST, L_LAST, c, c, "in")
        for i, v in enumerate(values):
            ws.cell(L_FIRST + i, c, v)
    ws["A4"] = ("※ 船は正式船名が分かるまで元帳の記号のまま使います。「常駐側の供給数に数える」は、同じ船を二重に数えないため、"
                "現在は元帳で使われている記号（T・S・Y・M・P・関・長）に○を付けています。グリーンシッピングの応援船が4隻あった場合は、船名を確認のうえ空き行へ追加してください。")
    ws["A4"].font = font(color="C00000")
    ws[f"{MC_NAME}4"] = "※ 会社・基地港・品名を追加すると、選択肢と「集計」へ反映されます。品名の高炉原料区分は初期値で要確認です。"
    ws[f"{MC_NAME}4"].font = font(color="C00000")
    ws.freeze_panes = "B6"
    select(ws, f"A{M_FIRST + len(MASTER_SHIPS)}")

    def dyn(name, letter, first, last):
        wb.defined_names[name] = DefinedName(
            name, attr_text=f"OFFSET(マスター!${letter}${first},0,0,MAX(1,COUNTA(マスター!${letter}${first}:${letter}${last})),1)")
    dyn("船名リスト", M_SEL, M_FIRST, M_LAST)
    dyn("会社リスト", MC_NAME, C_FIRST, C_LAST)
    dyn("基地港リスト", MB_PORT, L_FIRST, L_LAST)
    for n, (_, nm, _) in enumerate(CHOICES):
        dyn(nm, col(CHOICE_COL0 + n), L_FIRST, L_LAST)
    add_list(ws, f"{M_COMP}{M_FIRST}:{M_COMP}{M_LAST}", "=会社リスト")
    add_list(ws, f"{M_KIND}{M_FIRST}:{M_KIND}{M_LAST}", "=選択_種類")
    add_list(ws, f"{M_BASE}{M_FIRST}:{M_BASE}{M_LAST}", "=基地港リスト")
    circle(ws, f"{M_SUPPLY}{M_FIRST}:{M_SUPPLY}{M_LAST}")
    add_list(ws, f"{M_CONF}{M_FIRST}:{M_CONF}{M_LAST}", "=選択_確認状況")
    add_list(ws, f"{MC_KUBUN}{C_FIRST}:{MC_KUBUN}{C_LAST}", "=選択_会社区分")
    add_list(ws, f"{MC_CONF}{C_FIRST}:{MC_CONF}{C_LAST}", "=選択_確認状況")
    add_list(ws, f"{MH_KOURO}{L_FIRST}:{MH_KOURO}{L_LAST}", "=選択_高炉区分")
    circle(ws, f"{MH_AZE}{L_FIRST}:{MH_AZE}{L_LAST}")
    add_list(ws, f"{MH_CONF}{L_FIRST}:{MH_CONF}{L_LAST}", "=選択_確認状況")


# ---------------------------------------------------------------- 実績
JIS_HEADS = [
    ("作業キー【自動】", 14), ("作業の表示名【自動】", 34), ("作業の先頭行【自動】", 6), ("作業日【自動】", 11),
    ("元帳の行番号【自動】", 7), ("元帳No【自動】", 6), ("本船名【自動】", 18), ("G/T【自動】", 8),
    ("代理店（参考・会社判定に使わない）【自動】", 10), ("入出転【自動】", 7), ("六/部【自動】", 5), ("バース【自動】", 8),
    ("品名【自動】", 8), ("予定時間（元帳・定義未確認）【自動】", 9), ("作業開始（元帳）【自動】", 8), ("作業終了（元帳）【自動】", 8),
    ("終了が翌日【自動】", 6), ("タグ（元帳の記号）【自動】", 9), ("S列の記号（参考）【自動】", 6), ("備考1（元帳）【自動】", 16),
    ("取消しの記載【自動】", 6), ("時刻の扱い【自動】", 15),
    ("集計船名【自動】", 10), ("会社【自動】", 14), ("会社区分【自動】", 9), ("船名・会社の確認状況【自動】", 16),
    ("船の種類【自動】", 8), ("安瀬の扱い【自動】", 12), ("作業区分（LNG入港・出港・警戒を区別）【自動】", 11), ("高炉原料区分【自動】", 10),
    ("開始日時【自動】", 13), ("終了日時【自動】", 13), ("前日からの継続【自動】", 10), ("重複の警告【自動】", 24),
    ("開始（分）【自動】", 6), ("終了（分）【自動】", 6), ("同時稼働に含む【自動】", 6), ("高炉原料を除く集計に含む【自動】", 6),
    ("作業の開始（分）【自動】", 6), ("作業の終了（分）【自動】", 6),
]


def build_jisseki(wb, rows, target):
    ws = wb.create_sheet("実績")
    sheet_head(ws, "実績（元帳から自動表示・再入力しない）",
               "誰が：全員が見るだけ　何を見て：元帳「作業実績R6.4～」シート「R7年度」　どこへ：入力欄はありません（休船からの呼出は「変更・キャンセル」で最初に対応した作業を選びます）",
               f"対象日 {target:%Y/%m/%d}（前日から日をまたいで続く作業も表示）。時刻は元帳の「作業時間」で、基地発・基地着かは未確認のため「候補・時刻要確認」。集計先：「集計」「前日確定」「変更・キャンセル」", 34)
    for i, (t, w) in enumerate(JIS_HEADS, 1):
        header(ws, 5, i, t, "auto", w)
    ws.row_dimensions[5].height = 70
    last = JIS_FIRST + len(rows) - 1
    if rows:
        paint(ws, JIS_FIRST, last, 1, len(JIS_HEADS), "auto")
    jobs, jobno, key = [], 0, None
    for n, (rowno, r) in enumerate(rows):
        R = JIS_FIRST + n
        d = fy_date(r[0], r[1])
        first = r[4] == 1 or r[2] is not None
        if first or key is None:
            jobno += 1
            key = f"{d:%Y-%m-%d}#{jobno:02d}"
            sched = r[10].strftime("%H:%M") if isinstance(r[10], dt.time) else "--:--"
            label = f"{d:%m/%d} #{jobno:02d} {sched} {r[8] or ''} {r[3]}（{r[14] or ''}）"
            jobs.append({"key": key, "label": label, "ship": r[3], "io": r[8], "berth": r[14], "sched": r[10],
                         "date": d, "tugs": []})
        jobs[-1]["tugs"].append(r[17])
        s, e = r[11], r[12]
        cross, bad = "", False
        if isinstance(e, dt.datetime):
            if (e.year, e.month, e.day) == (1900, 1, 1):
                e, cross = e.time(), "○"
            else:
                bad = True
        elif isinstance(s, dt.time) and isinstance(e, dt.time) and e < s:
            cross = "○"
        timing = "候補・時刻要確認" if isinstance(s, dt.time) and isinstance(e, dt.time) and not bad else "時刻なし（集計対象外）"
        vals = [key, jobs[-1]["label"], "○" if first else "", d, rowno, r[2], r[3], r[5], r[6], r[8], r[9], r[14],
                (r[16] or "").strip() or None, r[10], s if isinstance(s, dt.time) else None,
                e if isinstance(e, dt.time) else None, cross, r[17], r[18], r[24], "○" if is_cancel(r) else "", timing]
        for c, v in enumerate(vals, 1):
            ws.cell(R, c, v)
        ws[f"{J_DATE}{R}"].number_format = "yyyy/mm/dd"
        for c in ("N", J_START, J_END):
            ws[f"{c}{R}"].number_format = "hh:mm"
        key_expr = f"{J_TUG}{R}"
        f = {
            J_NAME: f"={lookup_master(key_expr, M_NAME, key_expr)}",
            J_COMP: f'=IFERROR(IF({lookup_master(key_expr, M_COMP, "NA()")}="","未確認",{lookup_master(key_expr, M_COMP, "NA()")}),"マスター未登録")',
            J_KUBUN: f'={lookup_master(key_expr, M_KUBUN, chr(34) + "マスター未登録" + chr(34))}',
            J_CONF: f'=IFERROR(IF({lookup_master(key_expr, M_CONF, "NA()")}="","未確認",{lookup_master(key_expr, M_CONF, "NA()")}),"マスター未登録")',
            J_KIND: f'=IFERROR(IF({lookup_master(key_expr, M_KIND, "NA()")}="","未確認",{lookup_master(key_expr, M_KIND, "NA()")}),"マスター未登録")',
            J_AZE: (f'=IF(IFERROR(INDEX({rng("マスター", MH_AZE, L_FIRST, L_LAST)},MATCH({J_HIN}{R},'
                    f'{rng("マスター", MH_NAME, L_FIRST, L_LAST)},0)),"")="○","安瀬候補（要確認）","")'),
            J_SAGYO: (f'=IF({J_IO}{R}="","",IF({J_HIN}{R}="LNG",IF({J_IO}{R}="入","LNG入港",IF({J_IO}{R}="出","LNG出港",'
                      f'IF({J_IO}{R}="当直","LNG警戒",IF({J_IO}{R}="エスコート","LNGエスコート","LNG"&{J_IO}{R})))),'
                      f'IF({J_IO}{R}="入","入港",IF({J_IO}{R}="出","出港",IF({J_IO}{R}="転","転錨・シフト",{J_IO}{R})))))'),
            J_KOURO: (f'=IF({J_HIN}{R}="","未確認",IFERROR(INDEX({rng("マスター", MH_KOURO, L_FIRST, L_LAST)},MATCH({J_HIN}{R},'
                      f'{rng("マスター", MH_NAME, L_FIRST, L_LAST)},0))&"","未確認"))'),
            J_SDT: f'=IF({J_START}{R}="","",{J_DATE}{R}+{J_START}{R})',
            J_EDT: f'=IF({J_END}{R}="","",{J_DATE}{R}+{J_END}{R}+IF({J_CROSS}{R}="○",1,0))',
            J_PREV: f'=IF({J_DATE}{R}<集計!$C$5,"前日から継続","")',
            J_SMIN: f'=IF(OR({J_START}{R}="",{J_TIMING}{R}<>"候補・時刻要確認"),"",ROUND(({J_DATE}{R}-集計!$C$5)*1440+{J_START}{R}*1440,0))',
            J_EMIN: (f'=IF(OR({J_END}{R}="",{J_TIMING}{R}<>"候補・時刻要確認"),"",ROUND(({J_DATE}{R}-集計!$C$5)*1440+{J_END}{R}*1440'
                     f'+IF({J_CROSS}{R}="○",1440,0),0))'),
            J_INC: f'=IF(AND({J_SMIN}{R}<>"",{J_EMIN}{R}<>"",{J_CANCEL}{R}<>"○",{J_KIND}{R}<>"通船・警戒船"),"○","")',
            J_INC2: f'=IF(AND({J_INC}{R}="○",{J_KOURO}{R}<>"高炉原料"),"○","")',
            J_JSMIN: (f'=IF({J_FIRSTROW}{R}<>"○","",IF(COUNTIFS({JR(J_KEY)},{J_KEY}{R},{JR(J_INC)},"○")=0,"",'
                      f'_xlfn.MINIFS({JR(J_SMIN)},{JR(J_KEY)},{J_KEY}{R},{JR(J_INC)},"○")))'),
            J_JEMIN: (f'=IF({J_FIRSTROW}{R}<>"○","",IF(COUNTIFS({JR(J_KEY)},{J_KEY}{R},{JR(J_INC)},"○")=0,"",'
                      f'_xlfn.MAXIFS({JR(J_EMIN)},{JR(J_KEY)},{J_KEY}{R},{JR(J_INC)},"○")))'),
            J_WARN: (f'=IF({J_INC}{R}<>"○","",IF(COUNTIFS({JR(J_NAME)},{J_NAME}{R},{JR(J_SMIN)},"<"&{J_EMIN}{R},{JR(J_EMIN)},">"&{J_SMIN}{R},'
                     f'{JR(J_INC)},"○",{JR(J_KEY)},"<>"&{J_KEY}{R})>0,"重複配船の警告（同じ船が同じ時間に別作業）",'
                     f'IF(COUNTIFS({JR(J_NAME)},{J_NAME}{R},{JR(J_SMIN)},"<"&{J_EMIN}{R},{JR(J_EMIN)},">"&{J_SMIN}{R},'
                     f'{JR(J_INC)},"○",{JR(J_KEY)},{J_KEY}{R})>1,"同一記録の重複（1隻として数える）","")))'),
        }
        for c, v in f.items():
            ws[f"{c}{R}"] = v
        ws[f"{J_SDT}{R}"].number_format = "yyyy/mm/dd hh:mm"
        ws[f"{J_EDT}{R}"].number_format = "yyyy/mm/dd hh:mm"
    for c in range(cidx(J_SMIN), cidx(J_JEMIN) + 1):
        ws.column_dimensions[col(c)].hidden = True
    if not rows:
        ws["A6"] = "（この日の実績はありません）"
    ws.freeze_panes = "H6"
    select(ws, "A6")
    return jobs


# ---------------------------------------------------------------- 前日確定
def build_zenjitsu(wb, jobs, target):
    ws = wb.create_sheet("前日確定")
    sheet_head(ws, "前日確定（対象日の前日最終の予定表と翌日の出勤手配）",
               "誰が：上田　何を見て：対象日の前日に翌日の出勤を手配した最終の作業予定表　どこへ：①C7～C10　②出勤手配（B14～）　③作業ごとの予定（J列～T列）。予定表にあって実績にない作業は下の空き行へ追加",
               "集計先：「集計」の予定・出勤手配、「変更・キャンセル」の変更前の時刻・タグ。ここに入れた前日最終の値は変更があっても上書きしません。", 25)
    section(ws, 5, "■ 対象日の予定表")
    labels = [("対象日【自動（取込時に設定）】", "auto"), ("根拠の予定表（ファイル名・版）【入力】", "in"), ("予定表の確定時点（日時）【入力】", "in"),
              ("照合した人【入力】", "in"), ("照合の状況【入力・選択】", "in")]
    for n, (t, k) in enumerate(labels):
        r = 6 + n
        ws.cell(r, 1, t).font = font(bold=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
        paint(ws, r, r, 3, 5, k)
    ws["C6"] = target
    ws["C6"].number_format = "yyyy/mm/dd"
    ws["C8"].number_format = "yyyy/mm/dd hh:mm"
    ws["C10"] = "未確認"
    ws["F7"] = "◀ 最初の入力はここ（対象日の前日最終の予定表のファイル名・版。2025年度分は本一式に含まれていないため、上田さんの照合で入力）"
    ws["F7"].font = font(bold=True, color="C00000")
    add_list(ws, "C10", "=選択_照合")

    section(ws, 12, "■ 翌日の出勤手配（対象日の前日に手配した船）と、実際に出勤・待機したか")
    heads = [("No", "auto", 4), ("船名【入力・選択】", "in", 34), ("会社【自動】", "auto", 14), ("基地港【自動】", "auto", 14),
             ("手配した日時【入力】", "in", 14, "yyyy/mm/dd hh:mm"), ("実際の出勤・待機【入力・選択】", "in", 14), ("情報元【入力】", "in", 16)]
    for i, h in enumerate(heads, 1):
        header(ws, 13, i, h[0], h[1])
        paint(ws, 14, 21, i, i, h[1], h[3] if len(h) > 3 else None)
    for r in range(14, 22):
        ws[f"A{r}"] = r - 13
        ws[f"C{r}"] = f'=IF(B{r}="","",IFERROR(IF({lookup_master(f"B{r}", M_COMP, "NA()")}="","未確認",{lookup_master(f"B{r}", M_COMP, "NA()")}),"マスター未登録"))'
        ws[f"D{r}"] = f'=IF(B{r}="","",IFERROR(IF({lookup_master(f"B{r}", M_BASE, "NA()")}="","未確認",{lookup_master(f"B{r}", M_BASE, "NA()")}),"マスター未登録"))'
    add_list(ws, "B14:B21", "=船名リスト", "マスターの船名から選びます")
    add_list(ws, "F14:F21", "=選択_出勤", "前日の手配とは別に、実際に出勤・待機したかを選びます")
    for r, t, f in [(22, "出勤手配隻数【自動】", '=IF(COUNTA(B14:B21)=0,0,SUMPRODUCT((B14:B21<>"")/COUNTIF(B14:B21,B14:B21&"")))'),
                    (23, "実際に出勤した隻数【自動】", '=IF(COUNTA(B14:B21)=0,0,SUMPRODUCT((B14:B21<>"")*(LEFT(F14:F21,4)="出勤して")/COUNTIF(B14:B21,B14:B21&"")))')]:
        ws[f"A{r}"] = t
        ws[f"A{r}"].font = font(bold=True)
        ws.merge_cells(f"A{r}:B{r}")
        ws[f"C{r}"] = f
        paint(ws, r, r, 3, 3, "auto")

    section(ws, 24, "■ 作業ごとの前日最終予定（実績の作業は自動表示。予定表にだけある作業は下の空き行へ追加）")
    heads = [("区分【自動】", "auto", 8), ("作業の表示名【自動】", "auto", 34), ("作業キー【自動】", "auto", 14),
             ("本船名【自動／予定のみは入力】", "auto", 18), ("入出転【自動／予定のみは入力】", "auto", 7), ("バース【自動／予定のみは入力】", "auto", 8),
             ("実績のタグ（元帳）【自動】", "auto", 14), ("実績タグ数【自動】", "auto", 6), ("元帳の予定時間（参考・前日確定とは限らない）【自動】", "auto", 10, "hh:mm"),
             ("前日最終の予定時刻【入力】（翌日は日付も）", "in", 11, FMT_DT)] + \
            [(f"予定タグ{c}【入力・選択】", "in", 8) for c in "①②③④⑤⑥"] + \
            [("予定タグ数【自動】", "auto", 6), ("予定タグ（まとめ）【自動】", "auto", 16), ("照合の状況【入力・選択】", "in", 13),
             ("備考【入力】", "in", 16), ("変更回数【自動】", "auto", 6), ("変更後の最終日時【自動】", "auto", 11, "m/d hh:mm"),
             ("取消し【自動】", "auto", 6), ("作業日【自動】", "auto", 10, "yyyy/mm/dd"), ("前日確定の日時【自動】", "auto", 11, "m/d hh:mm")]
    table(ws, 25, heads, YJ_FIRST, YJ_LAST)
    for n, j in enumerate(jobs):
        r = YJ_FIRST + n
        for c, v in enumerate(["実績あり", j["label"], j["key"], j["ship"], j["io"], j["berth"], "・".join(j["tugs"])], 1):
            ws.cell(r, c, v)
        ws[f"I{r}"] = j["sched"]
        ws[f"S{r}"] = "未確認"
        ws[f"X{r}"] = j["date"]
    for r in range(YJ_FIRST + len(jobs), YJ_LAST + 1):
        n = r - YJ_FIRST - len(jobs) + 1
        paint(ws, r, r, 4, 6, "in")
        ws[f"A{r}"] = f'=IF(D{r}="","","予定のみ")'
        ws[f"C{r}"] = f'=IF(D{r}="","","{target:%Y-%m-%d}予定のみ#{n:02d}")'
        ws[f"B{r}"] = f'=IF(D{r}="","","{target:%m/%d} 予定のみ#{n:02d} "&IF(J{r}="","--:--",TEXT(J{r},"hh:mm"))&" "&E{r}&" "&D{r})'
        ws[f"G{r}"] = f'=IF(D{r}="","","（実績なし）")'
        ws[f"X{r}"] = f'=IF(D{r}="","",$C$6)'
    if jobs:
        ws.cell(YJ_FIRST + len(jobs) - 1, 26, "▼ ここから下は予定表にだけある作業（実績のない取消し等）を追加する行").font = font(color="C00000")
    H = "'変更・キャンセル'"
    for r in range(YJ_FIRST, YJ_LAST + 1):
        ws[f"H{r}"] = f'=IF(C{r}="","",COUNTIF({JR(J_KEY)},C{r}))'
        ws[f"Q{r}"] = f'=IF(C{r}="","",COUNTA(K{r}:P{r}))'
        ws[f"R{r}"] = f'=_xlfn.TEXTJOIN("・",TRUE,K{r}:P{r})'
        ws[f"U{r}"] = f'=IF(C{r}="","",COUNTIF({H}!$B${CH_FIRST}:$B${CH_LAST},C{r}))'
        ws[f"V{r}"] = f'=IF(OR(C{r}="",U{r}=0),"",INDEX({H}!$AA${CH_FIRST}:$AA${CH_LAST},MATCH(C{r}&"#"&U{r},{H}!$Z${CH_FIRST}:$Z${CH_LAST},0)))'
        ws[f"W{r}"] = f'=IF(C{r}="","",IF(COUNTIFS({H}!$B${CH_FIRST}:$B${CH_LAST},C{r},{H}!$R${CH_FIRST}:$R${CH_LAST},"○")>0,"○",""))'
        ws[f"Y{r}"] = f'=IF(OR(C{r}="",J{r}=""),"",IF(J{r}<1,X{r}+J{r},J{r}))'
    add_list(ws, f"K{YJ_FIRST}:P{YJ_LAST}", "=船名リスト", "予定表の曳船名をマスターの船名から選びます")
    add_list(ws, f"S{YJ_FIRST}:S{YJ_LAST}", "=選択_照合")
    add_list(ws, f"E{YJ_FIRST + len(jobs)}:E{YJ_LAST}", '"入,出,転,エスコート,当直,その他"')
    ws.freeze_panes = "C26"
    select(ws, "C7")


# ---------------------------------------------------------------- 変更・キャンセル
def build_henkou(wb, target):
    ws = wb.create_sheet("変更・キャンセル")
    sheet_head(ws, "変更・キャンセル（全作業。1回の変更を1行。前日確定は上書きしない）",
               "誰が：上田　何を見て：当日の変更連絡・変更分の予定表・配船記録　どこへ：①日ごとの確認（A7～）②休船からの呼出（A15～）③変更の記録（A26～で作業を選ぶ）。実績3隻未満の日・全船休航の日も①へ入力",
               "集計先：「前日確定」の変更回数・最終日時・取消し、「集計」の変更・取消し・呼出。実績のない取消しは「前日確定」の下の空き行へ作業を追加してから選びます。", 25)
    section(ws, 5, "■ 日ごとの変更確認（変更なしと確認できた日と未確認の日を区別）")
    for i, (t, k) in enumerate([("日付【入力】", "in"), ("その日の変更確認【入力・選択】", "in"), ("全船休航【入力・○】", "in"),
                                 ("実績の最大同時稼働（候補）【自動】", "auto"), ("情報元【入力】", "in"), ("確認者【入力】", "in")], 1):
        header(ws, 6, i, t, k)
        paint(ws, DAY_FIRST, DAY_LAST, i, i, k)
    ws.cell(DAY_FIRST, 1, target)
    ws.cell(DAY_FIRST, 2, "未確認")
    for r in range(DAY_FIRST, DAY_LAST + 1):
        ws[f"A{r}"].number_format = "yyyy/mm/dd"
        ws[f"D{r}"] = f'=IF(A{r}="","",IF(A{r}=集計!$C$5,集計!$C$14,"この試作では対象日のみ計算"))'
    add_list(ws, f"B{DAY_FIRST}:B{DAY_LAST}", "=選択_日確認")
    circle(ws, f"C{DAY_FIRST}:C{DAY_LAST}")

    section(ws, 13, "■ 休船からの呼出（製鉄曳船・西日本海運。1隻ごとに1行。同じ船の同じ呼出は1回と数える）")
    heads = [("呼出した船【入力・選択】", "in"), ("会社【自動】", "auto"), ("呼出日時【入力】", "in", "yyyy/mm/dd hh:mm"),
             ("最初に対応した作業【入力・選択】", "in"), ("呼出の理由【入力・選択】", "in"), ("情報元【入力】", "in"),
             ("確認状況【入力・選択】", "in"), ("同一呼出の判定【自動】", "auto"), ("対象会社の判定【自動】", "auto"),
             ("その日の従事作業数【自動】", "auto")]
    for i, h in enumerate(heads, 1):
        header(ws, 14, i, h[0], h[1])
        paint(ws, YB_FIRST, YB_LAST, i, i, h[1], h[2] if len(h) > 2 else None)
    for r in range(YB_FIRST, YB_LAST + 1):
        ws[f"B{r}"] = f'=IF(A{r}="","",IFERROR(IF({lookup_master(f"A{r}", M_COMP, "NA()")}="","未確認",{lookup_master(f"A{r}", M_COMP, "NA()")}),"マスター未登録"))'
        ws[f"H{r}"] = f'=IF(A{r}="","",IF(COUNTIFS(A${YB_FIRST}:A{r},A{r},C${YB_FIRST}:C{r},C{r})>1,"同一呼出（1回と数える）",""))'
        ws[f"I{r}"] = f'=IF(A{r}="","",IF(OR(B{r}="製鉄曳船",B{r}="西日本海運"),"","製鉄曳船・西日本海運以外の船です"))'
        nm = lookup_master(f"A{r}", M_NAME, f"A{r}")
        ws[f"J{r}"] = (f'=IF(A{r}="","",SUMPRODUCT(({JR(J_NAME)}={nm})/COUNTIFS({JR(J_KEY)},{JR(J_KEY)}&"",'
                       f'{JR(J_NAME)},{JR(J_NAME)}&"")))')
    add_list(ws, f"A{YB_FIRST}:A{YB_LAST}", "=船名リスト", "マスターの船名から選びます")
    add_list(ws, f"D{YB_FIRST}:D{YB_LAST}", "=作業リスト")
    add_list(ws, f"E{YB_FIRST}:E{YB_LAST}", "=選択_呼出理由")
    add_list(ws, f"G{YB_FIRST}:G{YB_LAST}", "=選択_確認状況")

    section(ws, 24, "■ 変更・キャンセルの記録（作業を選ぶと本船名・前日確定・変更前が自動表示。2回目以降の変更も同じ作業を選んで行を追加）")
    heads = [("作業の選択【入力・選択】", "in", 34), ("作業キー【自動】", "auto", 14), ("本船名【自動】", "auto", 16),
             ("何回目の変更【自動】", "auto", 6), ("前日確定の日時【自動】", "auto", 10, "m/d hh:mm"), ("前日確定のタグ【自動】", "auto", 14),
             ("連絡日時【入力】", "in", 14, "yyyy/mm/dd hh:mm"), ("連絡元【入力】", "in", 10), ("変更前の日時【自動】", "auto", 10, "m/d hh:mm"),
             ("変更後の時刻【入力】（翌日は日付も）", "in", 10, FMT_DT), ("変更前のタグ【自動】", "auto", 14)] + \
            [(f"変更後のタグ{c}【入力・選択】", "in", 8) for c in "①②③④⑤⑥"] + \
            [("取消し【入力・○】", "in", 6), ("取消しの時点【入力・選択】", "in", 11), ("理由【入力・選択】", "in", 11),
             ("理由の詳細【入力】", "in", 14), ("時間調整【入力・選択】", "in", 7), ("時間調整の内容【入力】", "in", 14),
             ("情報元【入力】", "in", 12), ("確認者【入力】", "in", 8),
             ("照合キー【自動】", "auto", 8), ("変更後の日時（有効）【自動】", "auto", 10, "m/d hh:mm"), ("変更後のタグ（有効）【自動】", "auto", 8),
             ("前回の照合キー【自動】", "auto", 8), ("作業日【自動】", "auto", 8, "yyyy/mm/dd")]
    table(ws, 25, heads, CH_FIRST, CH_LAST)
    Z = "'前日確定'"
    for r in range(CH_FIRST, CH_LAST + 1):
        m = f"MATCH(A{r},{Z}!$B${YJ_FIRST}:$B${YJ_LAST},0)"
        idx = lambda c: f"INDEX({Z}!${c}${YJ_FIRST}:${c}${YJ_LAST},{m})"
        ws[f"B{r}"] = f'=IF(A{r}="","",IFERROR({idx("C")},"作業が見つかりません"))'
        ws[f"C{r}"] = f'=IF(A{r}="","",IFERROR({idx("D")},""))'
        ws[f"D{r}"] = f'=IF(B{r}="","",COUNTIF(B${CH_FIRST}:B{r},B{r}))'
        ws[f"E{r}"] = f'=IF(A{r}="","",IFERROR(IF({idx("Y")}="","",{idx("Y")}),""))'
        ws[f"F{r}"] = f'=IF(A{r}="","",IFERROR({idx("R")},""))'
        ws[f"AD{r}"] = f'=IF(A{r}="","",IFERROR({idx("X")},""))'
        if r == CH_FIRST:  # 先頭行は必ず1回目（自分の行を参照しない＝循環参照を避ける）
            ws[f"I{r}"] = f'=IF(A{r}="","",E{r})'
            ws[f"K{r}"] = f'=IF(A{r}="","",F{r})'
        else:
            up = r - 1
            ws[f"I{r}"] = f'=IF(A{r}="","",IF(D{r}=1,E{r},IFERROR(INDEX($AA${CH_FIRST}:$AA${up},MATCH(AC{r},$Z${CH_FIRST}:$Z${up},0)),"")))'
            ws[f"K{r}"] = f'=IF(A{r}="","",IF(D{r}=1,F{r},IFERROR(INDEX($AB${CH_FIRST}:$AB${up},MATCH(AC{r},$Z${CH_FIRST}:$Z${up},0)),"")))'
        ws[f"Z{r}"] = f'=IF(B{r}="","",B{r}&"#"&D{r})'
        ws[f"AA{r}"] = f'=IF(A{r}="","",IF(J{r}="",I{r},IF(J{r}<1,AD{r}+J{r},J{r})))'
        ws[f"AB{r}"] = f'=IF(A{r}="","",IF(COUNTA(L{r}:Q{r})=0,K{r},_xlfn.TEXTJOIN("・",TRUE,L{r}:Q{r})))'
        ws[f"AC{r}"] = f'=IF(B{r}="","",B{r}&"#"&(D{r}-1))'
    add_list(ws, f"A{CH_FIRST}:A{CH_LAST}", "=作業リスト", "「前日確定」の作業から選びます")
    add_list(ws, f"L{CH_FIRST}:Q{CH_LAST}", "=船名リスト")
    circle(ws, f"R{CH_FIRST}:R{CH_LAST}")
    add_list(ws, f"S{CH_FIRST}:S{CH_LAST}", "=選択_取消時点")
    add_list(ws, f"T{CH_FIRST}:T{CH_LAST}", "=選択_理由")
    add_list(ws, f"V{CH_FIRST}:V{CH_LAST}", "=選択_時間調整")
    for c in range(cidx("Z"), cidx("AD") + 1):
        ws.column_dimensions[col(c)].hidden = True
    ws.freeze_panes = "B7"
    select(ws, "B7")
    wb.defined_names["作業リスト"] = DefinedName("作業リスト", attr_text=f"'前日確定'!$B${YJ_FIRST}:$B${YJ_LAST}")


# ---------------------------------------------------------------- 入渠・他港作業・他社配置・標準時間
def simple_sheet(wb, name, title, who, dest, heads):
    ws = wb.create_sheet(name)
    sheet_head(ws, title, who, dest, len(heads) + 1)
    table(ws, 5, heads, OT_FIRST, OT_LAST)
    ws.freeze_panes = "A6"
    select(ws, "A6")
    return ws


def master_cell(r, c_in, c_out, blank="未確認"):
    return (f'=IF({c_in}{r}="","",IFERROR(IF({lookup_master(f"{c_in}{r}", c_out, "NA()")}="","{blank}",'
            f'{lookup_master(f"{c_in}{r}", c_out, "NA()")}),"マスター未登録"))')


def build_others(wb):
    D, DT = "yyyy/mm/dd", "yyyy/mm/dd hh:mm"
    C5 = "集計!$C$5"
    ws = simple_sheet(wb, "入渠", "入渠（作業に使えない期間）",
                      "誰が：上田（製鉄曳船）・西日本海運担当者（西日本海運）、他社は分かる範囲　何を見て：入渠の記録　どこへ：A6から1隻・1回ごとに1行",
                      "集計先：「集計」の入渠で使えない船数・供給可能船の確認",
                      [("船名【入力・選択】", "in", 14), ("会社【自動】", "auto", 16), ("入渠開始日【入力】", "in", 12, D),
                       ("作業復帰日【入力】", "in", 12, D), ("情報元【入力】", "in", 22), ("確認状況【入力・選択】", "in", 12),
                       ("対象日に使えない【自動】", "auto", 9), ("集計船名【自動】", "auto", 10)])
    for r in range(OT_FIRST, OT_LAST + 1):
        ws[f"B{r}"] = master_cell(r, "A", M_COMP)
        ws[f"G{r}"] = f'=IF(OR(A{r}="",C{r}=""),"",IF(AND({C5}>=C{r},OR(D{r}="",{C5}<D{r})),"○",""))'
        ws[f"H{r}"] = f'=IF(A{r}="","",{lookup_master(f"A{r}", M_NAME, f"A{r}")})'
    add_list(ws, f"A{OT_FIRST}:A{OT_LAST}", "=船名リスト")
    add_list(ws, f"F{OT_FIRST}:F{OT_LAST}", "=選択_確認状況")

    ws = simple_sheet(wb, "他港作業", "他港作業（西日本海運の全他港作業。関豊丸・長豊丸に限定しない）",
                      "誰が：西日本海運担当者　何を見て：西日本海運の作業記録　どこへ：A6から1作業ごとに1行。実時刻が分からない場合は「時刻の確認状況」で区別",
                      "集計先：「集計」の他港作業で拘束された船数・供給可能船の確認",
                      [("船名【入力・選択】", "in", 14), ("会社【自動】", "auto", 14), ("作業港【入力】", "in", 10),
                       ("基地発日時【入力】", "in", 15, DT), ("基地着日時【入力】", "in", 15, DT),
                       ("時刻の確認状況【入力・選択】", "in", 10), ("情報元【入力】", "in", 18), ("確認者【入力】", "in", 9),
                       ("対象日に拘束【自動】", "auto", 8), ("会社の確認【自動】", "auto", 22), ("集計船名【自動】", "auto", 9),
                       ("基地発（対象日0時からの分）【自動】", "auto", 8), ("基地着（分）【自動】", "auto", 8)])
    for r in range(OT_FIRST, OT_LAST + 1):
        ws[f"B{r}"] = master_cell(r, "A", M_COMP)
        ws[f"I{r}"] = f'=IF(OR(A{r}="",D{r}=""),"",IF(AND(D{r}<{C5}+1,OR(E{r}="",E{r}>{C5})),"○",""))'
        ws[f"J{r}"] = f'=IF(A{r}="","",IF(B{r}="西日本海運","","西日本海運以外の船です（確認してください）"))'
        ws[f"K{r}"] = f'=IF(A{r}="","",{lookup_master(f"A{r}", M_NAME, f"A{r}")})'
        ws[f"L{r}"] = f'=IF(OR(A{r}="",D{r}=""),"",ROUND((D{r}-{C5})*1440,0))'
        ws[f"M{r}"] = f'=IF(OR(A{r}="",E{r}=""),"",ROUND((E{r}-{C5})*1440,0))'
    add_list(ws, f"A{OT_FIRST}:A{OT_LAST}", "=船名リスト")
    add_list(ws, f"F{OT_FIRST}:F{OT_LAST}", "=選択_時刻確認")

    ws = simple_sheet(wb, "他社配置", "他社配置（若松・関門等の他社の体制）",
                      "誰が：上田（分かる範囲）　何を見て：他社の配置・入渠・他港作業・応援依頼への回答　どこへ：A6から会社・船ごとに1行。依頼への対応可否が分からなければ空欄（＝未確認）",
                      "集計先：「集計」の他社配置・供給可能船の確認。対応可否が空欄の船は供給可能に数えません。2025年の配置は将来の応援確約ではありません。",
                      [("会社【入力・選択】", "in", 18), ("船名【入力・選択】", "in", 12), ("配置港【入力】", "in", 9),
                       ("対象期間の開始日【入力】", "in", 11, D), ("対象期間の終了日【入力】", "in", 11, D),
                       ("依頼への対応可否【入力・選択】（空欄＝未確認）", "in", 10), ("入渠・他港拘束（分かれば）【入力】", "in", 16),
                       ("情報元【入力】", "in", 16), ("確認日【入力】", "in", 11, D),
                       ("対象日に配置【自動】", "auto", 8), ("船名と会社の一致【自動】", "auto", 18), ("集計船名【自動】", "auto", 9)])
    for r in range(OT_FIRST, OT_LAST + 1):
        ws[f"J{r}"] = f'=IF(OR(B{r}="",D{r}=""),"",IF(AND({C5}>=D{r},OR(E{r}="",{C5}<=E{r})),"○",""))'
        ws[f"K{r}"] = (f'=IF(OR(A{r}="",B{r}=""),"",IF(IFERROR({lookup_master(f"B{r}", M_COMP, "NA()")}&"","")=A{r},"",'
                       f'"マスターの会社と異なります"))')
        ws[f"L{r}"] = f'=IF(B{r}="","",{lookup_master(f"B{r}", M_NAME, f"B{r}")})'
    add_list(ws, f"A{OT_FIRST}:A{OT_LAST}", "=会社リスト")
    add_list(ws, f"B{OT_FIRST}:B{OT_LAST}", "=船名リスト")
    add_list(ws, f"F{OT_FIRST}:F{OT_LAST}", "=選択_対応可否", "可・不可を選びます。分からない場合は空欄（未確認）")

    ws = simple_sheet(wb, "標準時間", "標準時間（基地発から基地着までの見込み）",
                      "誰が：八幡タグ＝上田、西日本海運タグ＝西日本海運担当者　何を見て：運航の実態　どこへ：A6から会社・基地港・バース・作業内容ごとに1行（LNG入港・出港・警戒は分ける）",
                      "集計先：今後の標準時間による推計（実測とは区別して表示）。右の上田資料の目安は起点が未確認のため参考表示のみ。",
                      [("会社【入力・選択】", "in", 16), ("基地港【入力・選択】", "in", 9), ("バース【入力】", "in", 8),
                       ("作業内容【入力・選択】", "in", 10), ("予定時刻の基準【入力・選択】", "in", 10),
                       ("基地から現場への移動（分）【入力】", "in", 9), ("作業（分）【入力】", "in", 8), ("帰航（分）【入力】", "in", 8),
                       ("合計（分）【自動】", "auto", 8), ("適用開始日【入力】", "in", 11, D), ("適用終了日【入力】", "in", 11, D),
                       ("確認者【入力】", "in", 9), ("情報元【入力】", "in", 14), ("確認状況【入力・選択】", "in", 11)])
    for r in range(OT_FIRST, OT_LAST + 1):
        ws[f"I{r}"] = f'=IF(COUNT(F{r}:H{r})=0,"",SUM(F{r}:H{r}))'
    add_list(ws, f"A{OT_FIRST}:A{OT_LAST}", "=会社リスト")
    add_list(ws, f"B{OT_FIRST}:B{OT_LAST}", "=基地港リスト")
    add_list(ws, f"D{OT_FIRST}:D{OT_LAST}", "=選択_作業内容")
    add_list(ws, f"E{OT_FIRST}:E{OT_LAST}", "=選択_時刻基準")
    add_list(ws, f"N{OT_FIRST}:N{OT_LAST}", "=選択_確認状況")
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


# ---------------------------------------------------------------- _計算・_計算2（非表示）
def build_calc(wb, name, inc_col, with_extra):
    ws = wb.create_sheet(name)
    ws["A1"] = ("補助計算（非表示）：5分刻みで、マスターの船ごとに対象日の作業中なら1。同じ集計船名は1列にまとめる。"
                + ("取消し・時刻なし・通船等を除く全作業。" if with_extra else "さらに高炉原料区分が「高炉原料」の行を除く。"))
    ws["A2"], ws["A3"], ws["A4"], ws["A5"], ws["B5"] = "会社区分", "集計船名", "当日稼働", "分", "時刻"
    for k in range(NSHIP):
        c = GRID_COL0 + k
        L = col(c)
        mr = M_FIRST + k
        ws.cell(1, c, f'=IF(マスター!${M_HEAD}${mr}<>"○","",IF(マスター!${M_COMP}${mr}="","未確認",マスター!${M_COMP}${mr}))')
        ws.cell(2, c, f'=IF(マスター!${M_HEAD}${mr}<>"○","",マスター!${M_KUBUN}${mr})')
        ws.cell(3, c, f'=IF(マスター!${M_HEAD}${mr}<>"○","",マスター!${M_NAME}${mr})')
        ws.cell(4, c, f'=IF({L}$3="",0,IF(COUNTIFS({JR(J_NAME)},{L}$3,{JR(inc_col)},"○")>0,1,0))')
        for s in range(SLOTS):
            r = SLOT_FIRST + s
            ws.cell(r, c, f'=IF({L}$4=0,0,IF(COUNTIFS({JR(J_NAME)},{L}$3,{JR(J_SMIN)},"<="&$A{r},{JR(J_EMIN)},">"&$A{r},{JR(inc_col)},"○")>0,1,0))')
    g1, g2 = col(GRID_COL0), col(GRID_LASTCOL)
    titles = ["全体の同時隻数", "作業中の本船作業数", "製鉄曳船（会社要確認）", "西日本海運"] if with_extra else ["同時隻数（高炉原料を除く）"]
    for i, t in enumerate(titles):
        ws.cell(5, COL_TOTAL + i, t)
    for s in range(SLOTS):
        r = SLOT_FIRST + s
        ws.cell(r, 1, s * 5)
        ws.cell(r, 2, f"=A{r}/1440").number_format = "hh:mm"
        ws.cell(r, COL_TOTAL, f"=SUM({g1}{r}:{g2}{r})")
        if with_extra:
            ws.cell(r, COL_JOBS, f'=SUMPRODUCT(({JR(J_FIRSTROW)}="○")*ISNUMBER({JR(J_JSMIN)})*({JR(J_JSMIN)}<=$A{r})*({JR(J_JEMIN)}>$A{r}))')
            ws.cell(r, COL_STC, f'=SUMIF(${g1}$2:${g2}$2,"製鉄曳船",{g1}{r}:{g2}{r})')
            ws.cell(r, COL_NW, f'=SUMIF(${g1}$2:${g2}$2,"西日本海運",{g1}{r}:{g2}{r})')
    ws.sheet_state = "hidden"


# ---------------------------------------------------------------- 集計
SUP_STATES = ["供給可能（待機）", "供給可能（他社・対応可）", "作業中", "入渠中", "他港作業中", "対応不可",
              "未確認（他社の対応可否が空欄）", "未確認（会社未確認）", "対象外（記号との対応未確認）", "対象外（通船・警戒船）"]


def build_summary(wb):
    ws = wb.create_sheet("集計")
    sheet_head(ws, "集計（自動表示。入力できるのは C7 の確認する時刻だけ）",
               "誰が：全員が見る　何を見て：各シートの入力結果　どこへ：入力欄は C7（供給可能船を確認する時刻。空欄なら最大同時の時刻）のみ",
               "この試作は入力と集計のつながりを試すものです。3隻と4隻の比較の結論や、4隻の必要性を示すものではありません。", 21)
    T, JB, ST, NW = (f"_計算!${col(c)}${SLOT_FIRST}:${col(c)}${SLOT_LAST}" for c in (COL_TOTAL, COL_JOBS, COL_STC, COL_NW))
    T2 = f"_計算2!${col(COL_TOTAL)}${SLOT_FIRST}:${col(COL_TOTAL)}${SLOT_LAST}"
    B = f"_計算!$B${SLOT_FIRST}:$B${SLOT_LAST}"
    Z, H = "'前日確定'", "'変更・キャンセル'"
    g1, g2 = col(GRID_COL0), col(GRID_LASTCOL)
    S, A = "section", "auto"
    rows = [
        (5, "対象日", f"={Z}!$C$6", "yyyy/mm/dd", "取込時に設定（実績が0件の日も保持）"),
        (6, "時刻の扱い", "候補・時刻要確認", None, "元帳の作業時間が基地発・基地着か未確認のため、同時稼働は候補"),
        (7, "供給可能船を確認する時刻【入力】", None, "hh:mm", "空欄なら最大同時の時刻（右の「供給可能船の確認」へ反映）"),
        (8, "■ 実績（元帳）", S),
        (9, "作業件数", f'=COUNTIF({JR(J_FIRSTROW)},"○")', None, "元帳の1作業（タグ別行ではない）"),
        (10, "延べタグ（タグ欄の行数）", f'=COUNTA({JR(J_TUG)})', None, "取消し記載の行も含む"),
        (11, "うち取消しの記載がある行", f'=COUNTIF({JR(J_CANCEL)},"○")', None, "同時稼働には含めない"),
        (12, "うち時刻がない行", f'=COUNTIF({JR(J_TIMING)},"時刻なし*")', None, "同時稼働には含めない"),
        (13, "同時稼働に含めた行", f'=COUNTIF({JR(J_INC)},"○")', None, "取消し・時刻なし・通船等を除く"),
        (14, "最大同時稼働隻数（候補）", f"=MAX({T})", None, "同じ船は1隻。複数本船の作業を重ねた5分刻みの最大"),
        (15, "最初にその隻数となった時刻", f'=IF(C14=0,"—",INDEX({B},MATCH(C14,{T},0)))', "hh:mm", None),
        (16, "その時刻に作業中の本船作業数", f'=IF(C14=0,0,INDEX({JB},MATCH(C14,{T},0)))', None, "1本船だけで3隻以上か、複数本船の重なりかを区別"),
        (17, "同時3隻以上の時間（分）", f'=COUNTIF({T},">=3")*5', None, None),
        (18, "　うち複数本船の重なり（分）", f"=SUMPRODUCT(({T}>=3)*({JB}>=2))*5", None, None),
        (19, "　うち1本船のみ（分）", f"=SUMPRODUCT(({T}>=3)*({JB}<2))*5", None, None),
        (20, "同時4隻以上の時間（分）", f'=COUNTIF({T},">=4")*5', None, None),
        (21, "同時5隻以上の時間（分）", f'=COUNTIF({T},">=5")*5', None, None),
        (22, "製鉄曳船（会社要確認）の最大同時", f"=MAX({ST})", None, "マスターで会社区分が「製鉄曳船」の船"),
        (23, "西日本海運の最大同時", f"=MAX({NW})", None, None),
        (24, "重複配船の警告がある行", f'=COUNTIF({JR(J_WARN)},"重複配船*")', None, "同じ船が同じ時間に別作業"),
        (25, "同一記録の重複がある行", f'=COUNTIF({JR(J_WARN)},"同一記録*")', None, "1隻として数えている"),
        (26, "マスター未登録の行", f'=COUNTIF({JR(J_COMP)},"マスター未登録")', None, "マスターへ追加すると解消"),
        (27, "船名・会社が確認済みでない行", f'=COUNTA({JR(J_TUG)})-COUNTIF({JR(J_CONF)},"確認済み")', None, "未確認として残している"),
        (29, "■ 高炉原料を除いた集計（全作業の集計とは別に表示）", S),
        (30, "全作業：延べタグ", "=C10", None, None),
        (31, "高炉原料を除く：延べタグ", f'=COUNTIFS({JR(J_TUG)},"<>",{JR(J_KOURO)},"<>高炉原料")', None, "マスターの品名で「高炉原料」の行を除く"),
        (32, "　うち高炉原料区分が未確認の行", f'=COUNTIFS({JR(J_TUG)},"<>",{JR(J_KOURO)},"未確認")', None, "除外していない。将来の扱いは別ケースで確認"),
        (33, "全作業：最大同時稼働（候補）", "=C14", None, None),
        (34, "高炉原料を除く：最大同時稼働（候補）", f"=MAX({T2})", None, None),
        (36, "■ 会社区分別（実績）", S),
        (44, "安瀬候補のタグ（按分配船・応援に含めない）", f'=COUNTIFS({JR(J_AZE)},"安瀬*",{JR(J_KUBUN)},"<>製鉄曳船",{JR(J_KUBUN)},"<>西日本海運")', None, "品名マスターで安瀬に○の品名（要確認）"),
        (45, "応援候補のタグ（他社・会社未確認。安瀬候補を除く）", f'=COUNTIFS({JR(J_AZE)},"",{JR(J_KUBUN)},"他社")+COUNTIFS({JR(J_AZE)},"",{JR(J_KUBUN)},"未確認")', None, "実績で参加していないことを「応援不可」とは扱わない"),
        (47, "■ 予定・出勤手配・同時稼働（それぞれ別の数値）", S),
        (48, "前日確定の照合状況", f"={Z}!$C$10", None, None),
        (49, "前日最終の予定時刻を入力した作業数", f"=COUNT({Z}!$J${YJ_FIRST}:$J${YJ_LAST})", None, None),
        (50, "予定のみ（実績なし）の作業数", f'=COUNTIF({Z}!$A${YJ_FIRST}:$A${YJ_LAST},"予定のみ")', None, None),
        (51, "作業予定タグ数（前日確定・延べ）", f"=SUM({Z}!$Q${YJ_FIRST}:$Q${YJ_LAST})", None, None),
        (52, "予定タグの船数（重複を除く）", f'=IF(COUNTA({Z}!$K${YJ_FIRST}:$P${YJ_LAST})=0,0,SUMPRODUCT(({Z}!$K${YJ_FIRST}:$P${YJ_LAST}<>"")/COUNTIF({Z}!$K${YJ_FIRST}:$P${YJ_LAST},{Z}!$K${YJ_FIRST}:$P${YJ_LAST}&"")))', None, None),
        (53, "翌日の出勤手配隻数（前日）", f"={Z}!$C$22", None, None),
        (54, "実際に出勤・待機した隻数", f"={Z}!$C$23", None, "前日の手配とは別に入力"),
        (55, "実作業の延べタグ（実績）", "=C10-C11", None, "取消し記載の行を除く"),
        (56, "実績の最大同時稼働（候補）", "=C14", None, None),
        (57, "同時に必要な隻数", "未入力（代表事例で別途確認）", None, "予定上の4隻出勤は4隻同時必要の証明ではない"),
        (59, "■ 変更・キャンセル", S),
        (60, "対象日の変更確認", f'=IFERROR(IF(INDEX({H}!$B${DAY_FIRST}:$B${DAY_LAST},MATCH(C5,{H}!$A${DAY_FIRST}:$A${DAY_LAST},0))="","未確認",INDEX({H}!$B${DAY_FIRST}:$B${DAY_LAST},MATCH(C5,{H}!$A${DAY_FIRST}:$A${DAY_LAST},0))),"未入力（未確認）")', None, None),
        (61, "全船休航", f'=IFERROR(IF(INDEX({H}!$C${DAY_FIRST}:$C${DAY_LAST},MATCH(C5,{H}!$A${DAY_FIRST}:$A${DAY_LAST},0))="○","○",""),"")', None, None),
        (62, "変更記録の行数", f"=COUNTA({H}!$A${CH_FIRST}:$A${CH_LAST})", None, None),
        (63, "変更のあった作業数", f'=SUMPRODUCT(({H}!$B${CH_FIRST}:$B${CH_LAST}<>"")/COUNTIF({H}!$B${CH_FIRST}:$B${CH_LAST},{H}!$B${CH_FIRST}:$B${CH_LAST}&""))', None, None),
        (64, "取消しになった作業数", f'=COUNTIF({Z}!$W${YJ_FIRST}:$W${YJ_LAST},"○")', None, None),
        (65, "うち基地発後の取消し（記録数）", f'=COUNTIFS({H}!$R${CH_FIRST}:$R${CH_LAST},"○",{H}!$S${CH_FIRST}:$S${CH_LAST},"当日・基地発後")', None, "実拘束あり。実施作業へは数えない"),
        (66, "時間調整ありの記録", f'=COUNTIF({H}!$V${CH_FIRST}:$V${CH_LAST},"あり")', None, None),
        (67, "休船からの呼出回数（同一呼出は1回）", f'=COUNTA({H}!$A${YB_FIRST}:$A${YB_LAST})-COUNTIF({H}!$H${YB_FIRST}:$H${YB_LAST},"同一呼出*")', None, "呼出実績＝常駐船腹の不足とは限らない"),
        (68, "呼出した船のその日の従事作業数（合計）", f'=SUMIFS({H}!$J${YB_FIRST}:$J${YB_LAST},{H}!$A${YB_FIRST}:$A${YB_LAST},"<>",{H}!$H${YB_FIRST}:$H${YB_LAST},"")', None, "呼出1回で複数作業に従事した場合も作業ごとに数える"),
        (69, "　うち乗組員の出勤手配の不足による呼出", f'=COUNTIFS({H}!$E${YB_FIRST}:$E${YB_LAST},"乗組員の出勤手配の不足",{H}!$H${YB_FIRST}:$H${YB_LAST},"")', None, "船体の不足と区別"),
        (70, "呼出の対象外の会社の記録", f'=COUNTIF({H}!$I${YB_FIRST}:$I${YB_LAST},"製鉄曳船*")', None, None),
        (72, "■ 使えない船・拘束・他社配置（対象日）", S),
        (73, "入渠で使えない船数", f'=COUNTIF(入渠!$G${OT_FIRST}:$G${OT_LAST},"○")', None, None),
        (74, "他港作業で拘束された船数（対象日中）", f'=COUNTIF(他港作業!$I${OT_FIRST}:$I${OT_LAST},"○")', None, None),
        (75, "他社配置の記録数", f'=COUNTIF(他社配置!$J${OT_FIRST}:$J${OT_LAST},"○")', None, "配置の記録であり応援の確約ではない"),
    ]
    for item in rows:
        r, label = item[0], item[1]
        if len(item) == 3:
            section(ws, r, label, 2)
            continue
        f, fmt, note = item[2:]
        ws.cell(r, 2, label).font = font()
        ws.cell(r, 3, f)
        paint(ws, r, r, 3, 3, "in" if "【入力】" in label else A, fmt)
        if note:
            ws.cell(r, 4, note).font = font(color="595959")
    header(ws, 37, 2, "会社区分【自動】", A)
    header(ws, 37, 3, "延べタグ【自動】", A)
    header(ws, 37, 4, "稼働した船数【自動】", A)
    for n, k in enumerate(["製鉄曳船", "西日本海運", "他社", "未確認", "マスター未登録"]):
        r = 38 + n
        ws.cell(r, 2, k)
        ws.cell(r, 3, f'=COUNTIF({JR(J_KUBUN)},B{r})')
        ws.cell(r, 4, f'=SUMIF(_計算!${g1}$2:${g2}$2,B{r},_計算!${g1}$4:${g2}$4)' if k != "マスター未登録" else "—")
        paint(ws, r, r, 2, 4, A)

    # 時間帯別
    section(ws, 8, "■ 時間帯別の最大同時稼働（候補）", 6)
    for i, t in enumerate(["時台", "全体【自動】", "本船作業数【自動】", "製鉄曳船（要確認）【自動】", "西日本海運【自動】", "高炉原料を除く【自動】"]):
        header(ws, 9, 6 + i, t, A)
    for h in range(24):
        r = 10 + h
        a, b = SLOT_FIRST + h * 12, SLOT_FIRST + h * 12 + 11
        ws.cell(r, 6, f"{h}時")
        for i, (sheet, cc) in enumerate((("_計算", COL_TOTAL), ("_計算", COL_JOBS), ("_計算", COL_STC), ("_計算", COL_NW), ("_計算2", COL_TOTAL))):
            L = col(cc)
            ws.cell(r, 7 + i, f"=MAX({sheet}!{L}{a}:{L}{b})")
        paint(ws, r, r, 6, 11, A)

    # 会社別（マスターの会社一覧）
    section(ws, 36, "■ 会社別（マスターの会社一覧から自動）", 6)
    for i, t in enumerate(["会社【自動】", "会社区分【自動】", "延べタグ【自動】", "稼働した船数【自動】"]):
        header(ws, 37, 6 + i, t, A)
    for n in range(C_LAST - C_FIRST + 1):
        r = 38 + n
        mr = C_FIRST + n
        ws.cell(r, 6, f'=IF(マスター!${MC_NAME}${mr}="","",マスター!${MC_NAME}${mr})')
        ws.cell(r, 7, f'=IF(F{r}="","",マスター!${MC_KUBUN}${mr})')
        ws.cell(r, 8, f'=IF(F{r}="","",COUNTIF({JR(J_COMP)},F{r}))')
        ws.cell(r, 9, f'=IF(F{r}="","",SUMIF(_計算!${g1}$1:${g2}$1,F{r},_計算!${g1}$4:${g2}$4))')
        paint(ws, r, r, 6, 9, A)
    r = 38 + C_LAST - C_FIRST + 1
    ws.cell(r, 6, "会社未確認")
    ws.cell(r, 8, f'=COUNTIF({JR(J_COMP)},"未確認")')
    ws.cell(r, 9, f'=SUMIF(_計算!${g1}$1:${g2}$1,"未確認",_計算!${g1}$4:${g2}$4)')
    paint(ws, r, r, 6, 9, A)

    # 供給可能船の確認
    sc = 13  # M列から
    Lc = lambda i: col(sc + i)
    section(ws, 8, "■ 供給可能船の確認（確認する時刻に、入渠中・他港作業中・作業中・未確認の船を除く）", sc)
    header(ws, 9, sc, "確認する時刻【自動】", A)
    ws.cell(9, sc + 1, f'=IF(C7<>"",C7,IF(C14=0,"",INDEX({B},MATCH(C14,{T},0))))').number_format = "hh:mm"
    ws.cell(9, sc + 2, f'=IF({Lc(1)}9="","",ROUND({Lc(1)}9*1440,0))')
    ws.cell(9, sc + 3, f'=IF({Lc(2)}9="",1,MIN({SLOTS},INT({Lc(2)}9/5)+1))')
    paint(ws, 9, 9, sc + 1, sc + 3, A)
    ws.cell(8, sc + 2).value = None
    for n, st in enumerate(SUP_STATES):
        r = 10 + n
        ws.cell(r, sc, st)
        ws.cell(r, sc + 1, f'=COUNTIF(${Lc(3)}${SUP_FIRST}:${Lc(3)}${SUP_FIRST + NSHIP - 1},M{r})')
        paint(ws, r, r, sc, sc + 1, A)
    ws.cell(20, sc, "※ 常駐側は、マスターで「常駐側の供給数に数える」に○の船。他社は「他社配置」で対応可否が「可」の船だけを供給可能に数え、空欄は未確認。乗組員の出勤は別（前日確定）。").font = font(color="595959")
    for i, t in enumerate(["船（集計船名）【自動】", "会社【自動】", "会社区分【自動】", "状態【自動】"]):
        header(ws, 21, sc + i, t, A)
    grid = f"_計算!${g1}${SLOT_FIRST}:${g2}${SLOT_LAST}"
    inyu = lambda c: rng("入渠", c, OT_FIRST, OT_LAST)
    tako = lambda c: rng("他港作業", c, OT_FIRST, OT_LAST)
    tasha = lambda c: rng("他社配置", c, OT_FIRST, OT_LAST)
    N9 = f"${Lc(2)}$9"
    for k in range(NSHIP):
        r = SUP_FIRST + k
        mr = M_FIRST + k
        q, rr, s = f"{Lc(0)}{r}", f"{Lc(1)}{r}", f"{Lc(2)}{r}"
        ws.cell(r, sc, f'=IF(マスター!${M_HEAD}${mr}<>"○","",マスター!${M_NAME}${mr})')
        ws.cell(r, sc + 1, f'=IF({q}="","",IF(マスター!${M_COMP}${mr}="","未確認",マスター!${M_COMP}${mr}))')
        ws.cell(r, sc + 2, f'=IF({q}="","",マスター!${M_KUBUN}${mr})')
        ws.cell(r, sc + 3, (
            f'=IF({q}="","",IF(マスター!${M_KIND}${mr}="通船・警戒船","対象外（通船・警戒船）",'
            f'IF(COUNTIFS({inyu("H")},{q},{inyu("G")},"○")>0,"入渠中",'
            f'IF(AND({N9}<>"",COUNTIFS({tako("K")},{q},{tako("L")},"<="&{N9},{tako("M")},">"&{N9})'
            f'+COUNTIFS({tako("K")},{q},{tako("L")},"<="&{N9},{tako("E")},"")>0),"他港作業中",'
            f'IF(INDEX({grid},${Lc(3)}$9,{k + 1})=1,"作業中",'
            f'IF({s}="未確認","未確認（会社未確認）",'
            f'IF(OR({s}="製鉄曳船",{s}="西日本海運"),IF(マスター!${M_SUPPLY}${mr}="○","供給可能（待機）","対象外（記号との対応未確認）"),'
            f'IF(COUNTIFS({tasha("L")},{q},{tasha("J")},"○",{tasha("F")},"可")>0,"供給可能（他社・対応可）",'
            f'IF(COUNTIFS({tasha("L")},{q},{tasha("J")},"○",{tasha("F")},"不可")>0,"対応不可","未確認（他社の対応可否が空欄）")))))))))'))
        paint(ws, r, r, sc, sc + 3, A)
    for c, w in zip(range(1, sc + 4), (2, 34, 12, 34, 2, 10, 9, 9, 11, 9, 10, 2, 26, 14, 10, 26)):
        ws.column_dimensions[col(c)].width = w
    ws.column_dimensions[col(sc + 2)].hidden = False
    select(ws, "C7")


# ---------------------------------------------------------------- はじめに
def build_intro(wb, target, nrows, njobs):
    ws = wb.create_sheet("はじめに", 0)
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 100
    ws["B1"] = "八幡地区 常駐体制検証 入力試作 v01"
    ws["B1"].font = font(size=16, bold=True)
    lines = [
        ("目的", "製鉄曳船の3隻と4隻で対応の差を説明するための基礎データを集めます。この試作は、入力と集計のつながりを試すためのものです。4隻の必要性を証明するものではありません。"),
        ("対象", f"元帳R7年度のうち、同時3隻以上の候補日 {target:%Y年%m月%d日} の1日分（実績{nrows}行・{njobs}作業）。元帳の時刻は「候補・時刻要確認」です。"),
        ("最初の操作", "下の青い文字を押すと、最初の入力欄（前日確定シートC7：根拠の予定表）へ移動します。"),
        ("入力の順番", "①前日確定 → ②変更・キャンセル → ③入渠・他港作業・他社配置 → ④標準時間。船名や会社が選択肢にないときは、先にマスターへ追加します。"),
        ("担当者", "前日確定・変更・キャンセル・製鉄曳船の入渠・八幡タグの標準時間：上田さん／西日本海運の他港作業・入渠・標準時間：西日本海運担当者／他社配置：分かる範囲で上田さん"),
        ("色の意味", "薄い黄色＝入力する欄（見出しに【入力】）／薄い灰色＝自動表示（見出しに【自動】）。灰色の欄は入力しません。"),
        ("選び方", "船名・会社・理由などはセルを選ぶと出る▼の一覧から選びます。○の欄は▼から○を選びます。ボタンやマクロは使っていません。"),
        ("未確認の意味", "資料で確かめられていないことです。推測で埋めず「未確認」のまま残します。確認できたら「確認済み」に変えます。他社の対応可否が空欄の船も未確認です。"),
        ("再入力しない", "本船名・タグ名・実績時刻は元帳から自動表示します。作業番号の入力は不要で、作業は▼の一覧から選びます。"),
        ("前日確定と変更", "前日最終の予定は「前日確定」に入れ、変更があっても上書きしません。変更は「変更・キャンセル」に1回ごとに1行追加します。翌日にかかる時刻は日付も入れます（例：2025/12/7 0:30）。"),
        ("実績のない取消し", "予定表にあって元帳にない作業は、「前日確定」の下の空き行に本船名などを入れて追加し、「変更・キャンセル」で選んで取消しを入力します。全作業が取り消された日も、前日の手配と実際の出勤を残せます。"),
        ("安瀬・高炉原料", "安瀬は若松との按分配船で、応援とは別に数えます。高炉原料を除いた集計は、全作業の集計とは別に表示します（品名の区分はマスターで設定。初期値は要確認）。"),
        ("集計", "「集計」シートで、作業予定タグ数・出勤手配隻数・実際の出勤・実績の同時稼働・供給可能船を別々に表示します。同時に必要な隻数は代表事例で別に確認します。"),
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
    first_link = f"B{r}"
    r += 2
    ws.cell(r, 2, "色見本").font = font(bold=True)
    ws.cell(r, 3, "入力する欄【入力】").fill = F_IN
    ws.cell(r + 1, 3, "自動表示の欄【自動】").fill = F_AUTO
    r += 3
    ws.cell(r, 2, "シートの役割").font = font(bold=True)
    for n, (s, d) in enumerate([("実績", "元帳の取り込み（見るだけ）"), ("前日確定", "前日最終の予定・予定タグ・翌日の出勤手配と実際の出勤"),
                                 ("変更・キャンセル", "日ごとの変更確認、休船からの呼出、変更・取消しの記録"),
                                 ("入渠", "作業に使えない期間"), ("他港作業", "西日本海運の全他港作業"), ("他社配置", "若松・関門等の他社の体制と依頼への対応可否"),
                                 ("標準時間", "基地発～基地着の見込み時間"), ("マスター", "会社・船名・基地港・品名・選択肢の追加"),
                                 ("集計", "対比結果・供給可能船・不足情報（自動）")]):
        c = ws.cell(r + 1 + n, 2, s)
        c.hyperlink = Hyperlink(ref=c.coordinate, location=f"'{s}'!A1", display=s)
        c.font = font(color="0563C1", underline="single")
        ws.cell(r + 1 + n, 3, d).font = font()
    select(ws, first_link)


def build(rows, target, out):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    jobs = build_jisseki(wb, rows, target)
    build_zenjitsu(wb, jobs, target)
    build_henkou(wb, target)
    build_others(wb)
    build_master(wb)
    build_summary(wb)
    build_calc(wb, "_計算", J_INC, True)
    build_calc(wb, "_計算2", J_INC2, False)
    build_intro(wb, target, len(rows), len(jobs))
    order = ["はじめに", "実績", "前日確定", "変更・キャンセル", "入渠", "他港作業", "他社配置", "標準時間", "マスター", "集計", "_計算", "_計算2"]
    wb._sheets = [wb[n] for n in order]
    wb.active = 0
    for ws in wb.worksheets:
        ws.sheet_view.tabSelected = ws.title == "はじめに"
    wb.calculation.fullCalcOnLoad = True
    os.makedirs(os.path.dirname(out), exist_ok=True)
    wb.save(out)
    return jobs


def main():
    rows = read_target_rows()
    jobs = build(rows, TARGET, OUT)
    print("saved", OUT, "rows", len(rows), "jobs", len(jobs))


if __name__ == "__main__":
    main()
