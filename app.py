import sqlite3
from pathlib import Path
from datetime import date, datetime, timedelta

import pandas as pd
import streamlit as st

LOGIN_ID = "hdmr"
LOGIN_PASSWORD = "rui"

def login_check():

    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False

    if st.session_state.logged_in:
        return True

    st.set_page_config(
        page_title="健康チェック管理システム",
        page_icon="🩺",
        layout="wide"
    )

    st.markdown("""
    <h1 style='text-align:center; color:#2E7D32;'>
    健康チェック管理システム
    </h1>
    """, unsafe_allow_html=True)

    st.write("")

    col1, col2, col3 = st.columns([1,2,1])

    with col2:

        st.markdown("""
        ### ログイン
        利用者様の健康記録を安全に管理するため、
        ID・パスワードを入力してください。
        """)

        input_id = st.text_input("ID")

        input_password = st.text_input(
            "パスワード",
            type="password"
        )

        st.write("")

        if st.button("ログイン", use_container_width=True):

            if (
                input_id == LOGIN_ID
                and input_password == LOGIN_PASSWORD
            ):

                st.session_state.logged_in = True
                st.success("ログインしました。")
                st.rerun()

            else:
                st.error("IDまたはパスワードが違います。")

    return False

# =========================
# ログイン判定
# =========================

if not login_check():
    st.stop()



import pandas as pd
from pathlib import Path
from datetime import date, datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

APP_DIR = Path(__file__).parent
DATA_DIR = APP_DIR / "data"
REPORT_DIR = APP_DIR / "reports"
DATA_FILE = DATA_DIR / "健康チェック入力データ.xlsx"
USER_FILE = DATA_DIR / "利用者マスタ.xlsx"
SHEET_NAME = "入力データ"
USER_SHEET = "利用者マスタ"

DEFAULT_USERS = [
    "さくら様",
    "谷様",
    "磯崎様",
    "川上様",
    "和波様",
    "桜井様",
    "國枝様",
    "中野様",
    "山口様",
]

COLUMNS = [
    "記録日",
    "利用者名",
    "体温",
    "血圧上",
    "血圧下",
    "脈拍",
    "SpO2",
    "体重",
    "家族共有メモ",
    "気になる変化",
    "登録日時",
    "入力者",
]


def ensure_dirs():
    DATA_DIR.mkdir(exist_ok=True)
    REPORT_DIR.mkdir(exist_ok=True)


def ensure_user_file():
    ensure_dirs()
    if not USER_FILE.exists():
        df = pd.DataFrame({"利用者名": DEFAULT_USERS, "表示": ["表示"] * len(DEFAULT_USERS)})
        df.to_excel(USER_FILE, index=False, sheet_name=USER_SHEET)


def load_users(include_hidden=False):
    ensure_user_file()
    try:
        df = pd.read_excel(USER_FILE, sheet_name=USER_SHEET)
    except Exception:
        df = pd.DataFrame({"利用者名": DEFAULT_USERS, "表示": ["表示"] * len(DEFAULT_USERS)})

    if "利用者名" not in df.columns:
        df["利用者名"] = DEFAULT_USERS
    if "表示" not in df.columns:
        df["表示"] = "表示"

    df["利用者名"] = df["利用者名"].astype(str).str.strip()
    df = df[df["利用者名"] != ""].drop_duplicates(subset=["利用者名"], keep="first")

    if not include_hidden:
        df = df[df["表示"].fillna("表示") == "表示"]

    return df.reset_index(drop=True)


def save_users(df):
    ensure_dirs()
    df = df.copy()
    if "利用者名" not in df.columns:
        df["利用者名"] = ""
    if "表示" not in df.columns:
        df["表示"] = "表示"

    df["利用者名"] = df["利用者名"].astype(str).str.strip()
    df = df[df["利用者名"] != ""].drop_duplicates(subset=["利用者名"], keep="first")
    df.to_excel(USER_FILE, index=False, sheet_name=USER_SHEET)


def add_user(user_name):
    user_name = str(user_name).strip()
    if not user_name:
        return False, "利用者名が空欄です。"

    df = load_users(include_hidden=True)
    if user_name in df["利用者名"].tolist():
        df.loc[df["利用者名"] == user_name, "表示"] = "表示"
        save_users(df)
        return True, f"{user_name} を表示に戻しました。"

    new_row = pd.DataFrame([{"利用者名": user_name, "表示": "表示"}])
    df = pd.concat([df, new_row], ignore_index=True)
    save_users(df)
    return True, f"{user_name} を追加しました。"


def hide_user(user_name):
    df = load_users(include_hidden=True)
    if user_name not in df["利用者名"].tolist():
        return False, "対象の利用者が見つかりません。"

    df.loc[df["利用者名"] == user_name, "表示"] = "非表示"
    save_users(df)
    return True, f"{user_name} を入力候補から外しました。過去データは残ります。"


def ensure_data_file():
    ensure_dirs()
    if not DATA_FILE.exists():
        df = pd.DataFrame(columns=COLUMNS)
        df.to_excel(DATA_FILE, index=False, sheet_name=SHEET_NAME)


def load_data():
    ensure_data_file()
    try:
        df = pd.read_excel(DATA_FILE, sheet_name=SHEET_NAME)
    except Exception:
        df = pd.DataFrame(columns=COLUMNS)

    for col in COLUMNS:
        if col not in df.columns:
            df[col] = ""

    df = df[COLUMNS]
    if not df.empty:
        df["記録日"] = pd.to_datetime(df["記録日"], errors="coerce")
    return df


def save_data(df):
    ensure_dirs()
    df.to_excel(DATA_FILE, index=False, sheet_name=SHEET_NAME)


def append_record(record):
    df = load_data()
    new_df = pd.DataFrame([record], columns=COLUMNS)
    df = pd.concat([df, new_df], ignore_index=True)
    save_data(df)


def to_number(series):
    return pd.to_numeric(series, errors="coerce")


def create_family_report(df, user_name, year, month):
    filtered = df.copy()
    filtered["記録日"] = pd.to_datetime(filtered["記録日"], errors="coerce")

    target = filtered[
        (filtered["利用者名"] == user_name)
        & (filtered["記録日"].dt.year == int(year))
        & (filtered["記録日"].dt.month == int(month))
    ].copy()

    target = target.sort_values("記録日")

    wb = Workbook()
    ws = wb.active
    ws.title = "家族向けレポート"

    thin = Side(style="thin", color="CCCCCC")
    border = Border(top=thin, bottom=thin, left=thin, right=thin)

    title_fill = PatternFill("solid", fgColor="EADFCB")
    section_fill = PatternFill("solid", fgColor="DDEFE2")
    note_fill = PatternFill("solid", fgColor="FFF8E7")

    ws.merge_cells("A1:H1")
    ws["A1"] = "ご家族向け 健康・生活レポート"
    ws["A1"].font = Font(name="Meiryo", size=16, bold=True)
    ws["A1"].fill = title_fill
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    ws["A3"] = "利用者名"
    ws["B3"] = user_name
    ws["D3"] = "対象月"
    ws["E3"] = f"{year}年{month}月"

    ws.merge_cells("A5:H5")
    ws["A5"] = "今月の健康状態"
    ws["A5"].font = Font(name="Meiryo", bold=True)
    ws["A5"].fill = section_fill

    metrics = [
        ("体温平均", "体温", "℃"),
        ("血圧上平均", "血圧上", ""),
        ("血圧下平均", "血圧下", ""),
        ("脈拍平均", "脈拍", "回/分"),
        ("SpO2平均", "SpO2", "%"),
        ("体重平均", "体重", "kg"),
    ]

    start_row = 7
    for i, (label, col, unit) in enumerate(metrics):
        row = start_row + i
        ws[f"A{row}"] = label
        if target.empty:
            value = ""
        else:
            value = to_number(target[col]).mean()
            value = "" if pd.isna(value) else round(float(value), 1)
        ws[f"B{row}"] = value
        ws[f"C{row}"] = unit

    ws.merge_cells("A14:H14")
    ws["A14"] = "日付別のご様子（家族共有メモ）"
    ws["A14"].font = Font(name="Meiryo", bold=True)
    ws["A14"].fill = section_fill

    ws["A15"] = "日付"
    ws["B15"] = "家族共有メモ"
    ws["A15"].fill = note_fill
    ws["B15"].fill = note_fill

    row = 16
    memo_rows = target[target["家族共有メモ"].fillna("").astype(str).str.strip() != ""]
    if memo_rows.empty:
        ws[f"A{row}"] = ""
        ws[f"B{row}"] = "記録された家族共有メモはありません。"
        row += 1
    else:
        for _, rec in memo_rows.iterrows():
            ws[f"A{row}"] = rec["記録日"].strftime("%m/%d")
            ws[f"B{row}"] = str(rec["家族共有メモ"])
            row += 1

    row += 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    ws.cell(row=row, column=1).value = "日付別の気になる変化"
    ws.cell(row=row, column=1).font = Font(name="Meiryo", bold=True)
    ws.cell(row=row, column=1).fill = section_fill

    row += 1
    ws.cell(row=row, column=1).value = "日付"
    ws.cell(row=row, column=2).value = "気になる変化"
    ws.cell(row=row, column=1).fill = note_fill
    ws.cell(row=row, column=2).fill = note_fill

    row += 1
    change_rows = target[target["気になる変化"].fillna("").astype(str).str.strip() != ""]
    if change_rows.empty:
        ws.cell(row=row, column=1).value = ""
        ws.cell(row=row, column=2).value = "記録された気になる変化はありません。"
        row += 1
    else:
        for _, rec in change_rows.iterrows():
            ws.cell(row=row, column=1).value = rec["記録日"].strftime("%m/%d")
            ws.cell(row=row, column=2).value = str(rec["気になる変化"])
            row += 1

    row += 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    ws.cell(row=row, column=1).value = "まとめ"
    ws.cell(row=row, column=1).font = Font(name="Meiryo", bold=True)
    ws.cell(row=row, column=1).fill = section_fill

    row += 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row+2, end_column=8)
    ws.cell(row=row, column=1).value = (
        "記録上の内容をもとに、今月のご様子をまとめています。"
        "医療的な判断や診断ではなく、日々の記録に基づく共有です。"
        "引き続き、ご本人の様子を見守ってまいります。"
    )
    ws.cell(row=row, column=1).alignment = Alignment(wrap_text=True, vertical="top")

    for col in range(1, 9):
        ws.column_dimensions[get_column_letter(col)].width = 16
    ws.column_dimensions["B"].width = 70

    for cells in ws.iter_rows():
        for cell in cells:
            cell.font = Font(name="Meiryo", size=10, bold=cell.font.bold if cell.font else False)
            cell.border = border
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    file_name = f"家族向けレポート_{user_name}_{year}年{month}月.xlsx"
    report_path = REPORT_DIR / file_name
    wb.save(report_path)
    return report_path, target


st.set_page_config(
    page_title="健康チェックWebアプリ",
    page_icon="📝",
    layout="wide",
)

st.title("健康チェックWebアプリ")
st.caption("入力はブラウザで、データはExcelへ保存。月別の家族向けレポートを出力できます。")

ensure_data_file()
ensure_user_file()

menu = st.sidebar.radio(
    "メニュー",
    ["健康チェック入力", "入力データ確認", "家族向けレポート作成", "利用者マスタ管理"],
)

active_users = load_users(include_hidden=False)["利用者名"].tolist()
all_users_df = load_users(include_hidden=True)
all_users = all_users_df["利用者名"].tolist()

if not active_users:
    st.warning("表示中の利用者がいません。左メニューの「利用者マスタ管理」から利用者を追加してください。")

if menu == "健康チェック入力":
    st.header("健康チェック入力")

    if not active_users:
        st.stop()

    with st.form("health_form", clear_on_submit=True):
        col1, col2, col3 = st.columns(3)
        with col1:
            record_date = st.date_input("記録日", value=date.today())
        with col2:
            user_name = st.selectbox("利用者名", active_users)
        with col3:
            input_staff = st.text_input("入力者", placeholder="例：藤野")

        col4, col5, col6 = st.columns(3)
        with col4:
            temp = st.number_input("体温", min_value=30.0, max_value=45.0, value=36.5, step=0.1)
        with col5:
            bp_high = st.number_input("血圧上", min_value=50, max_value=250, value=120, step=1)
        with col6:
            bp_low = st.number_input("血圧下", min_value=30, max_value=150, value=75, step=1)

        col7, col8, col9 = st.columns(3)
        with col7:
            pulse = st.number_input("脈拍", min_value=30, max_value=200, value=70, step=1)
        with col8:
            spo2 = st.number_input("SpO2", min_value=70, max_value=100, value=96, step=1)
        with col9:
            weight = st.number_input("体重", min_value=0.0, max_value=200.0, value=50.0, step=0.1)

        family_memo = st.text_area("家族共有メモ", placeholder="ご家族へ共有してよい内容を入力")
        changes = st.text_area("気になる変化", placeholder="食事、睡眠、歩行、表情、体調など")

        submitted = st.form_submit_button("登録する")

    if submitted:
        record = {
            "記録日": record_date,
            "利用者名": user_name,
            "体温": temp,
            "血圧上": bp_high,
            "血圧下": bp_low,
            "脈拍": pulse,
            "SpO2": spo2,
            "体重": weight,
            "家族共有メモ": family_memo,
            "気になる変化": changes,
            "登録日時": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "入力者": input_staff,
        }
        append_record(record)
        st.success("登録しました。Excelデータに保存されています。")

elif menu == "入力データ確認":
    st.header("入力データ確認")
    df = load_data()

    filter_options = ["全員"] + all_users
    col1, col2, col3 = st.columns(3)
    with col1:
        filter_user = st.selectbox("利用者で絞り込み", filter_options)
    with col2:
        year = st.number_input("年", min_value=2024, max_value=2035, value=date.today().year, step=1)
    with col3:
        month = st.number_input("月", min_value=1, max_value=12, value=date.today().month, step=1)

    view = df.copy()
    if not view.empty:
        view["記録日"] = pd.to_datetime(view["記録日"], errors="coerce")
        view = view[(view["記録日"].dt.year == int(year)) & (view["記録日"].dt.month == int(month))]
        if filter_user != "全員":
            view = view[view["利用者名"] == filter_user]

    st.dataframe(view, use_container_width=True)

    with open(DATA_FILE, "rb") as f:
        st.download_button(
            "入力データExcelをダウンロード",
            data=f,
            file_name="健康チェック入力データ.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

elif menu == "家族向けレポート作成":
    st.header("家族向けレポート作成")
    df = load_data()

    if not all_users:
        st.warning("利用者が登録されていません。")
        st.stop()

    col1, col2, col3 = st.columns(3)
    with col1:
        report_user = st.selectbox("利用者名", all_users)
    with col2:
        report_year = st.number_input("対象年", min_value=2024, max_value=2035, value=date.today().year, step=1)
    with col3:
        report_month = st.number_input("対象月", min_value=1, max_value=12, value=date.today().month, step=1)

    if st.button("家族向けレポートを作成"):
        report_path, target = create_family_report(df, report_user, report_year, report_month)

        if target.empty:
            st.warning("指定した利用者・年月のデータがありません。空のレポートを作成しました。")
        else:
            st.success("家族向けレポートを作成しました。")

        with open(report_path, "rb") as f:
            st.download_button(
                "作成したレポートをダウンロード",
                data=f,
                file_name=report_path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

elif menu == "利用者マスタ管理":
    st.header("利用者マスタ管理")
    st.caption("利用者の追加・入力候補からの削除ができます。削除しても過去の入力データは残ります。")

    df_users = load_users(include_hidden=True)

    st.subheader("現在の利用者一覧")
    st.dataframe(df_users, use_container_width=True)

    st.subheader("利用者を追加")
    with st.form("add_user_form", clear_on_submit=True):
        new_user = st.text_input("追加する利用者名", placeholder="例：田中様")
        add_submit = st.form_submit_button("追加する")

    if add_submit:
        ok, msg = add_user(new_user)
        if ok:
            st.success(msg)
            st.rerun()
        else:
            st.error(msg)

    st.subheader("利用者を入力候補から外す")
    visible_users = load_users(include_hidden=False)["利用者名"].tolist()
    if visible_users:
        target_user = st.selectbox("対象利用者", visible_users)
        st.warning("この操作は、入力画面の候補から外すだけです。過去データは削除されません。")
        if st.button("入力候補から外す"):
            ok, msg = hide_user(target_user)
            if ok:
                st.success(msg)
                st.rerun()
            else:
                st.error(msg)
    else:
        st.info("現在、表示中の利用者はいません。")

    st.subheader("非表示の利用者を戻す")
    hidden_df = df_users[df_users["表示"] == "非表示"]
    if not hidden_df.empty:
        restore_user = st.selectbox("表示に戻す利用者", hidden_df["利用者名"].tolist())
        if st.button("表示に戻す"):
            ok, msg = add_user(restore_user)
            if ok:
                st.success(msg)
                st.rerun()
            else:
                st.error(msg)
    else:
        st.info("非表示の利用者はいません。")

    with open(USER_FILE, "rb") as f:
        st.download_button(
            "利用者マスタExcelをダウンロード",
            data=f,
            file_name="利用者マスタ.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

































































































from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

APP_DIR = Path(__file__).parent
DATA_DIR = APP_DIR / "data"
REPORT_DIR = APP_DIR / "reports"
DB_FILE = DATA_DIR / "health_check.db"

DEFAULT_USERS = ["さくら様", "谷様", "磯崎様", "川上様", "和波様", "桜井様", "國枝様", "中野様", "山口様"]

def ensure_dirs():
    DATA_DIR.mkdir(exist_ok=True)
    REPORT_DIR.mkdir(exist_ok=True)

def connect_db():
    ensure_dirs()
    return sqlite3.connect(DB_FILE, check_same_thread=False)

def init_db():
    con = connect_db()
    cur = con.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL,
            active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS health_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            record_date TEXT NOT NULL,
            user_name TEXT NOT NULL,
            temperature REAL,
            bp_high INTEGER,
            bp_low INTEGER,
            pulse INTEGER,
            spo2 INTEGER,
            weight REAL,
            family_memo TEXT,
            changes TEXT,
            input_staff TEXT,
            created_at TEXT,
            updated_at TEXT
        )
    """)

    cur.execute("PRAGMA table_info(health_records)")
    existing_cols = [r[1] for r in cur.fetchall()]
    if "updated_at" not in existing_cols:
        cur.execute("ALTER TABLE health_records ADD COLUMN updated_at TEXT")

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for name in DEFAULT_USERS:
        cur.execute(
            "INSERT OR IGNORE INTO users (name, active, created_at) VALUES (?, 1, ?)",
            (name, now),
        )

    con.commit()
    con.close()

def load_users(active_only=True):
    con = connect_db()
    sql = "SELECT id, name, active, created_at FROM users"
    if active_only:
        sql += " WHERE active = 1"
    sql += " ORDER BY id"
    df = pd.read_sql_query(sql, con)
    con.close()
    return df

def add_user(name):
    name = str(name).strip()
    if not name:
        return False, "利用者名が空欄です。"

    con = connect_db()
    cur = con.cursor()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    try:
        cur.execute("INSERT INTO users (name, active, created_at) VALUES (?, 1, ?)", (name, now))
        msg = f"{name} を追加しました。"
    except sqlite3.IntegrityError:
        cur.execute("UPDATE users SET active = 1 WHERE name = ?", (name,))
        msg = f"{name} を表示に戻しました。"

    con.commit()
    con.close()
    return True, msg

def set_user_active(name, active):
    con = connect_db()
    con.execute("UPDATE users SET active = ? WHERE name = ?", (active, name))
    con.commit()
    con.close()

def insert_record(record):
    con = connect_db()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    con.execute("""
        INSERT INTO health_records (
            record_date, user_name, temperature, bp_high, bp_low, pulse, spo2, weight,
            family_memo, changes, input_staff, created_at, updated_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        record["record_date"],
        record["user_name"],
        record["temperature"],
        record["bp_high"],
        record["bp_low"],
        record["pulse"],
        record["spo2"],
        record["weight"],
        record["family_memo"],
        record["changes"],
        record["input_staff"],
        now,
        "",
    ))
    con.commit()
    con.close()

def update_record(record_id, record):
    con = connect_db()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    con.execute("""
        UPDATE health_records
        SET record_date = ?,
            user_name = ?,
            temperature = ?,
            bp_high = ?,
            bp_low = ?,
            pulse = ?,
            spo2 = ?,
            weight = ?,
            family_memo = ?,
            changes = ?,
            input_staff = ?,
            updated_at = ?
        WHERE id = ?
    """, (
        record["record_date"],
        record["user_name"],
        record["temperature"],
        record["bp_high"],
        record["bp_low"],
        record["pulse"],
        record["spo2"],
        record["weight"],
        record["family_memo"],
        record["changes"],
        record["input_staff"],
        now,
        int(record_id),
    ))
    con.commit()
    con.close()

def delete_record(record_id):
    con = connect_db()
    con.execute("DELETE FROM health_records WHERE id = ?", (int(record_id),))
    con.commit()
    con.close()

def get_record(record_id):
    con = connect_db()
    df = pd.read_sql_query("SELECT * FROM health_records WHERE id = ?", con, params=[int(record_id)])
    con.close()

    if df.empty:
        return None

    row = df.iloc[0].to_dict()
    row["record_date"] = pd.to_datetime(row["record_date"], errors="coerce")
    return row

def search_records(user_name="全員", start_date=None, end_date=None, keyword=""):
    con = connect_db()
    sql = "SELECT * FROM health_records WHERE 1=1"
    params = []

    if user_name and user_name != "全員":
        sql += " AND user_name = ?"
        params.append(user_name)

    if start_date:
        sql += " AND record_date >= ?"
        params.append(start_date.strftime("%Y-%m-%d"))

    if end_date:
        sql += " AND record_date <= ?"
        params.append(end_date.strftime("%Y-%m-%d"))

    keyword = str(keyword).strip()
    if keyword:
        sql += " AND (family_memo LIKE ? OR changes LIKE ? OR input_staff LIKE ?)"
        kw = f"%{keyword}%"
        params.extend([kw, kw, kw])

    sql += " ORDER BY record_date DESC, id DESC"

    df = pd.read_sql_query(sql, con, params=params)
    con.close()

    if not df.empty:
        df["record_date"] = pd.to_datetime(df["record_date"], errors="coerce")

    return df

def monthly_records(user_name, year, month):
    start = date(int(year), int(month), 1)
    if int(month) == 12:
        end = date(int(year) + 1, 1, 1) - timedelta(days=1)
    else:
        end = date(int(year), int(month) + 1, 1) - timedelta(days=1)
    return search_records(user_name=user_name, start_date=start, end_date=end)

def display_df(df):
    if df.empty:
        return df

    out = df.copy()
    out = out.rename(columns={
        "id": "記録ID",
        "record_date": "記録日",
        "user_name": "利用者名",
        "temperature": "体温",
        "bp_high": "血圧上",
        "bp_low": "血圧下",
        "pulse": "脈拍",
        "spo2": "SpO2",
        "weight": "体重",
        "family_memo": "家族共有メモ",
        "changes": "気になる変化",
        "input_staff": "入力者",
        "created_at": "登録日時",
        "updated_at": "更新日時",
    })

    if "記録日" in out.columns:
        out["記録日"] = pd.to_datetime(out["記録日"], errors="coerce").dt.strftime("%Y-%m-%d")

    return out

def create_summary(df):
    if df.empty:
        return pd.DataFrame(columns=["項目", "平均", "最小", "最大", "記録数"])

    metrics = [
        ("体温", "temperature"),
        ("血圧上", "bp_high"),
        ("血圧下", "bp_low"),
        ("脈拍", "pulse"),
        ("SpO2", "spo2"),
        ("体重", "weight"),
    ]

    rows = []
    for label, col in metrics:
        s = pd.to_numeric(df[col], errors="coerce").dropna()
        if s.empty:
            rows.append({"項目": label, "平均": "", "最小": "", "最大": "", "記録数": 0})
        else:
            rows.append({
                "項目": label,
                "平均": round(float(s.mean()), 1),
                "最小": round(float(s.min()), 1),
                "最大": round(float(s.max()), 1),
                "記録数": int(s.count()),
            })

    return pd.DataFrame(rows)

def create_daily_count(df):
    if df.empty:
        return pd.DataFrame(columns=["日付", "記録数"])

    tmp = df.copy()
    tmp["日付"] = pd.to_datetime(tmp["record_date"], errors="coerce").dt.strftime("%Y-%m-%d")
    return tmp.groupby("日付").size().reset_index(name="記録数")

def export_records_excel(df, filename="検索結果.xlsx"):
    REPORT_DIR.mkdir(exist_ok=True)
    path = REPORT_DIR / filename
    display_df(df).to_excel(path, index=False)
    return path

def create_family_report(user_name, year, month):
    df = monthly_records(user_name, year, month)

    wb = Workbook()
    ws = wb.active
    ws.title = "家族向けレポート"

    thin = Side(style="thin", color="CCCCCC")
    border = Border(top=thin, bottom=thin, left=thin, right=thin)
    title_fill = PatternFill("solid", fgColor="EADFCB")
    section_fill = PatternFill("solid", fgColor="DDEFE2")
    note_fill = PatternFill("solid", fgColor="FFF8E7")

    ws.merge_cells("A1:H1")
    ws["A1"] = "ご家族向け 健康・生活レポート"
    ws["A1"].font = Font(name="Meiryo", size=16, bold=True)
    ws["A1"].fill = title_fill
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    ws["A3"] = "利用者名"
    ws["B3"] = user_name
    ws["D3"] = "対象月"
    ws["E3"] = f"{year}年{month}月"

    ws.merge_cells("A5:H5")
    ws["A5"] = "今月の健康状態"
    ws["A5"].font = Font(name="Meiryo", bold=True)
    ws["A5"].fill = section_fill

    summary = create_summary(df)
    r = 7
    for _, row in summary.iterrows():
        ws[f"A{r}"] = f'{row["項目"]} 平均'
        ws[f"B{r}"] = row["平均"]
        ws[f"D{r}"] = "最小"
        ws[f"E{r}"] = row["最小"]
        ws[f"F{r}"] = "最大"
        ws[f"G{r}"] = row["最大"]
        r += 1

    ws.merge_cells("A14:H14")
    ws["A14"] = "日付別のご様子（家族共有メモ）"
    ws["A14"].font = Font(name="Meiryo", bold=True)
    ws["A14"].fill = section_fill

    ws["A15"] = "日付"
    ws["B15"] = "家族共有メモ"
    ws["A15"].fill = note_fill
    ws["B15"].fill = note_fill

    r = 16
    memo_rows = df[df["family_memo"].fillna("").astype(str).str.strip() != ""] if not df.empty else pd.DataFrame()
    if memo_rows.empty:
        ws[f"B{r}"] = "記録された家族共有メモはありません。"
        r += 1
    else:
        for _, rec in memo_rows.sort_values("record_date").iterrows():
            ws[f"A{r}"] = rec["record_date"].strftime("%m/%d")
            ws[f"B{r}"] = str(rec["family_memo"])
            r += 1

    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    ws.cell(row=r, column=1).value = "日付別の気になる変化"
    ws.cell(row=r, column=1).font = Font(name="Meiryo", bold=True)
    ws.cell(row=r, column=1).fill = section_fill

    r += 1
    ws.cell(row=r, column=1).value = "日付"
    ws.cell(row=r, column=2).value = "気になる変化"
    ws.cell(row=r, column=1).fill = note_fill
    ws.cell(row=r, column=2).fill = note_fill

    r += 1
    change_rows = df[df["changes"].fillna("").astype(str).str.strip() != ""] if not df.empty else pd.DataFrame()
    if change_rows.empty:
        ws.cell(row=r, column=2).value = "記録された気になる変化はありません。"
        r += 1
    else:
        for _, rec in change_rows.sort_values("record_date").iterrows():
            ws.cell(row=r, column=1).value = rec["record_date"].strftime("%m/%d")
            ws.cell(row=r, column=2).value = str(rec["changes"])
            r += 1

    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    ws.cell(row=r, column=1).value = "まとめ"
    ws.cell(row=r, column=1).font = Font(name="Meiryo", bold=True)
    ws.cell(row=r, column=1).fill = section_fill

    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r+2, end_column=8)
    ws.cell(row=r, column=1).value = (
        "記録上の内容をもとに、今月のご様子をまとめています。"
        "医療的な判断や診断ではなく、日々の記録に基づく共有です。"
        "引き続き、ご本人の様子を見守ってまいります。"
    )

    for col in range(1, 9):
        ws.column_dimensions[get_column_letter(col)].width = 16
    ws.column_dimensions["B"].width = 70

    for rows in ws.iter_rows():
        for cell in rows:
            cell.font = Font(name="Meiryo", size=10, bold=cell.font.bold if cell.font else False)
            cell.border = border
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    REPORT_DIR.mkdir(exist_ok=True)
    path = REPORT_DIR / f"家族向けレポート_{user_name}_{year}年{month}月.xlsx"
    wb.save(path)
    return path, df

def make_record_form(prefix, default=None, active_users=None, all_users=None):
    default = default or {}
    active_users = active_users or []
    all_users = all_users or active_users

    default_date = default.get("record_date")
    if pd.isna(default_date) or default_date is None:
        default_date = date.today()
    elif hasattr(default_date, "date"):
        default_date = default_date.date()

    default_user = default.get("user_name") or (active_users[0] if active_users else "")
    choices = all_users if default else active_users
    idx = choices.index(default_user) if default_user in choices else 0

    c1, c2, c3 = st.columns(3)
    record_date = c1.date_input("記録日", value=default_date, key=f"{prefix}_date")
    user_name = c2.selectbox("利用者名", choices, index=idx, key=f"{prefix}_user")
    staff = c3.text_input("入力者", value=default.get("input_staff") or "", key=f"{prefix}_staff")

    c4, c5, c6 = st.columns(3)
    temp = c4.number_input("体温", 30.0, 45.0, float(default.get("temperature") or 36.5), step=0.1, key=f"{prefix}_temp")
    bp_high = c5.number_input("血圧上", 50, 250, int(default.get("bp_high") or 120), step=1, key=f"{prefix}_bph")
    bp_low = c6.number_input("血圧下", 30, 150, int(default.get("bp_low") or 75), step=1, key=f"{prefix}_bpl")

    c7, c8, c9 = st.columns(3)
    pulse = c7.number_input("脈拍", 30, 200, int(default.get("pulse") or 70), step=1, key=f"{prefix}_pulse")
    spo2 = c8.number_input("SpO2", 70, 100, int(default.get("spo2") or 96), step=1, key=f"{prefix}_spo2")
    weight = c9.number_input("体重", 0.0, 200.0, float(default.get("weight") or 50.0), step=0.1, key=f"{prefix}_weight")

    memo = st.text_area("家族共有メモ", value=default.get("family_memo") or "", key=f"{prefix}_memo")
    changes = st.text_area("気になる変化", value=default.get("changes") or "", key=f"{prefix}_changes")

    return {
        "record_date": record_date.strftime("%Y-%m-%d"),
        "user_name": user_name,
        "temperature": temp,
        "bp_high": bp_high,
        "bp_low": bp_low,
        "pulse": pulse,
        "spo2": spo2,
        "weight": weight,
        "family_memo": memo,
        "changes": changes,
        "input_staff": staff,
    }

st.set_page_config(page_title="健康チェックWebアプリ", page_icon="📝", layout="wide")
init_db()

st.title("健康チェックWebアプリ")
st.caption("過去データの登録・更新・削除・検索・サマリーに対応。保存はSQLite、出力はExcelです。")

menu = st.sidebar.radio(
    "メニュー",
    ["健康チェック入力", "過去データ登録", "検索・更新・削除", "サマリー", "家族向けレポート作成", "利用者マスタ管理", "バックアップ"],
)

active_users = load_users(True)["name"].tolist()
all_users = load_users(False)["name"].tolist()

if menu == "健康チェック入力":
    st.header("健康チェック入力")
    if not active_users:
        st.warning("表示中の利用者がいません。利用者マスタ管理から追加してください。")
        st.stop()

    with st.form("new_record_form", clear_on_submit=True):
        record = make_record_form("new", active_users=active_users, all_users=all_users)
        submitted = st.form_submit_button("登録する")

    if submitted:
        insert_record(record)
        st.success("登録しました。")

elif menu == "過去データ登録":
    st.header("過去データ登録")
    st.caption("過去の日付を指定して、あとから記録を追加できます。")

    if not active_users:
        st.warning("表示中の利用者がいません。")
        st.stop()

    default = {"record_date": date.today() - timedelta(days=1)}
    with st.form("past_record_form", clear_on_submit=True):
        record = make_record_form("past", default=default, active_users=active_users, all_users=all_users)
        submitted = st.form_submit_button("過去データとして登録する")

    if submitted:
        insert_record(record)
        st.success("過去データを登録しました。")

elif menu == "検索・更新・削除":
    st.header("検索・更新・削除")

    c1, c2, c3, c4 = st.columns(4)
    filter_user = c1.selectbox("利用者", ["全員"] + all_users)
    start_date = c2.date_input("開始日", value=date.today().replace(day=1))
    end_date = c3.date_input("終了日", value=date.today())
    keyword = c4.text_input("キーワード", placeholder="メモ・変化・入力者")

    df = search_records(filter_user, start_date, end_date, keyword)
    st.subheader("検索結果")
    st.dataframe(display_df(df), use_container_width=True)

    if not df.empty:
        export_path = export_records_excel(df, "検索結果.xlsx")
        with open(export_path, "rb") as f:
            st.download_button("検索結果をExcelでダウンロード", f, file_name=export_path.name)

        st.divider()
        st.subheader("記録の更新")
        selected_id = st.selectbox("更新する記録ID", df["id"].tolist())
        rec = get_record(selected_id)

        if rec:
            with st.form("edit_record_form"):
                updated = make_record_form("edit", default=rec, active_users=active_users, all_users=all_users)
                saved = st.form_submit_button("この内容で更新する")

            if saved:
                update_record(selected_id, updated)
                st.success(f"記録ID {selected_id} を更新しました。")
                st.rerun()

            st.divider()
            st.subheader("記録の削除")
            st.warning("削除すると元に戻せません。入力ミスの削除用として使ってください。")
            confirm = st.checkbox(f"記録ID {selected_id} を削除することを確認しました")
            if st.button("この記録を削除する", disabled=not confirm):
                delete_record(selected_id)
                st.success(f"記録ID {selected_id} を削除しました。")
                st.rerun()
    else:
        st.info("該当するデータがありません。")

elif menu == "サマリー":
    st.header("サマリー")

    c1, c2, c3 = st.columns(3)
    target_user = c1.selectbox("利用者", ["全員"] + all_users)
    year = c2.number_input("年", 2024, 2035, date.today().year)
    month = c3.number_input("月", 1, 12, date.today().month)

    df = monthly_records(target_user, year, month)
    summary = create_summary(df)
    daily_count = create_daily_count(df)

    st.subheader("月間サマリー")
    st.dataframe(summary, use_container_width=True)

    st.subheader("記録件数")
    st.metric("対象期間の記録件数", len(df))

    st.subheader("日別記録件数")
    st.dataframe(daily_count, use_container_width=True)

    st.subheader("メモ・気になる変化一覧")
    memo_cols = ["record_date", "user_name", "family_memo", "changes", "input_staff"]
    memo_df = df[memo_cols].copy() if not df.empty else pd.DataFrame(columns=memo_cols)
    st.dataframe(display_df(memo_df), use_container_width=True)

elif menu == "家族向けレポート作成":
    st.header("家族向けレポート作成")

    if not all_users:
        st.warning("利用者が登録されていません。")
        st.stop()

    c1, c2, c3 = st.columns(3)
    report_user = c1.selectbox("利用者名", all_users)
    report_year = c2.number_input("対象年", 2024, 2035, date.today().year)
    report_month = c3.number_input("対象月", 1, 12, date.today().month)

    if st.button("家族向けレポートを作成"):
        path, df = create_family_report(report_user, report_year, report_month)
        if df.empty:
            st.warning("指定した利用者・年月のデータがありません。空のレポートを作成しました。")
        else:
            st.success("家族向けレポートを作成しました。")

        with open(path, "rb") as f:
            st.download_button("作成したレポートをダウンロード", f, file_name=path.name)

elif menu == "利用者マスタ管理":
    st.header("利用者マスタ管理")

    dfu = load_users(False).copy()
    dfu_display = dfu.copy()
    dfu_display["active"] = dfu_display["active"].map({1: "表示", 0: "非表示"})
    st.dataframe(dfu_display.rename(columns={"id": "ID", "name": "利用者名", "active": "表示", "created_at": "登録日時"}), use_container_width=True)

    st.subheader("利用者を追加")
    with st.form("add_user_form", clear_on_submit=True):
        new_user = st.text_input("追加する利用者名", placeholder="例：田中様")
        add_submitted = st.form_submit_button("追加する")

    if add_submitted:
        ok, msg = add_user(new_user)
        if ok:
            st.success(msg)
            st.rerun()
        else:
            st.error(msg)

    visible = load_users(True)["name"].tolist()
    if visible:
        st.subheader("入力候補から外す")
        target = st.selectbox("対象利用者", visible)
        if st.button("入力候補から外す"):
            set_user_active(target, 0)
            st.success(f"{target} を入力候補から外しました。過去データは残ります。")
            st.rerun()

    hidden = load_users(False)
    hidden_names = hidden[hidden["active"] == 0]["name"].tolist()
    if hidden_names:
        st.subheader("非表示の利用者を戻す")
        restore = st.selectbox("表示に戻す利用者", hidden_names)
        if st.button("表示に戻す"):
            set_user_active(restore, 1)
            st.success(f"{restore} を表示に戻しました。")
            st.rerun()

elif menu == "バックアップ":
    st.header("バックアップ")

    if DB_FILE.exists():
        with open(DB_FILE, "rb") as f:
            st.download_button("SQLiteデータベースをダウンロード", f, file_name="health_check.db")

    all_df = search_records()
    path = export_records_excel(all_df, "全入力データ.xlsx")
    with open(path, "rb") as f:
        st.download_button("全データをExcelでダウンロード", f, file_name=path.name)
