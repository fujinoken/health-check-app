import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import date, datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter


# =========================
# ページ設定
# =========================
st.set_page_config(
    page_title="健康チェックWebアプリ",
    page_icon="🩺",
    layout="wide",
)


# =========================
# 簡易ログイン設定
# =========================
LOGIN_ID = "hdmr"
LOGIN_PASSWORD = "rui"


def login_check():
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False

    if st.session_state.logged_in:
        return True

    st.markdown(
        """
        <h1 style='text-align:center; color:#2E7D32;'>
        健康チェック管理システム
        </h1>
        """,
        unsafe_allow_html=True,
    )

    col1, col2, col3 = st.columns([1, 2, 1])

    with col2:
        st.markdown(
            """
            ### ログイン
            利用者様の健康記録を安全に管理するため、  
            ID・パスワードを入力してください。
            """
        )

        input_id = st.text_input("ID")
        input_password = st.text_input("パスワード", type="password")

        if st.button("ログイン", use_container_width=True):
            if input_id == LOGIN_ID and input_password == LOGIN_PASSWORD:
                st.session_state.logged_in = True
                st.rerun()
            else:
                st.error("IDまたはパスワードが違います。")

    return False


if not login_check():
    st.stop()


# =========================
# 保存先設定
# =========================
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


# =========================
# 基本関数
# =========================
def ensure_dirs():
    DATA_DIR.mkdir(exist_ok=True)
    REPORT_DIR.mkdir(exist_ok=True)


def ensure_user_file():
    ensure_dirs()
    if not USER_FILE.exists():
        df = pd.DataFrame(
            {
                "利用者名": DEFAULT_USERS,
                "表示": ["表示"] * len(DEFAULT_USERS),
            }
        )
        df.to_excel(USER_FILE, index=False, sheet_name=USER_SHEET)


def load_users(include_hidden=False):
    ensure_user_file()

    try:
        df = pd.read_excel(USER_FILE, sheet_name=USER_SHEET)
    except Exception:
        df = pd.DataFrame(
            {
                "利用者名": DEFAULT_USERS,
                "表示": ["表示"] * len(DEFAULT_USERS),
            }
        )

    if "利用者名" not in df.columns:
        df["利用者名"] = DEFAULT_USERS

    if "表示" not in df.columns:
        df["表示"] = "表示"

    df["利用者名"] = df["利用者名"].astype(str).str.strip()
    df = df[df["利用者名"] != ""].drop_duplicates(
        subset=["利用者名"],
        keep="first",
    )

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
    df = df[df["利用者名"] != ""].drop_duplicates(
        subset=["利用者名"],
        keep="first",
    )

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

    new_row = pd.DataFrame(
        [
            {
                "利用者名": user_name,
                "表示": "表示",
            }
        ]
    )

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


# =========================
# 家族向けレポート作成
# =========================
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

    memo_rows = target[
        target["家族共有メモ"].fillna("").astype(str).str.strip() != ""
    ]

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

    change_rows = target[
        target["気になる変化"].fillna("").astype(str).str.strip() != ""
    ]

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
    ws.merge_cells(start_row=row, start_column=1, end_row=row + 2, end_column=8)
    ws.cell(row=row, column=1).value = (
        "記録上の内容をもとに、今月のご様子をまとめています。"
        "医療的な判断や診断ではなく、日々の記録に基づく共有です。"
        "引き続き、ご本人の様子を見守ってまいります。"
    )
    ws.cell(row=row, column=1).alignment = Alignment(
        wrap_text=True,
        vertical="top",
    )

    for col in range(1, 9):
        ws.column_dimensions[get_column_letter(col)].width = 16

    ws.column_dimensions["B"].width = 70

    for cells in ws.iter_rows():
        for cell in cells:
            cell.font = Font(
                name="Meiryo",
                size=10,
                bold=cell.font.bold if cell.font else False,
            )
            cell.border = border
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    file_name = f"家族向けレポート_{user_name}_{year}年{month}月.xlsx"
    report_path = REPORT_DIR / file_name
    wb.save(report_path)

    return report_path, target


# =========================
# アプリ本体
# =========================
st.title("健康チェックWebアプリ")
st.caption("入力はブラウザで、データはExcelへ保存。月別の家族向けレポートを出力できます。")

ensure_data_file()
ensure_user_file()


# =========================
# サイドバーメニュー
# ※ここは1回だけ
# =========================
menu = st.sidebar.radio(
    "メニュー",
    [
        "健康チェック入力",
        "入力データ確認",
        "家族向けレポート作成",
        "利用者マスタ管理",
    ],
)


active_users = load_users(include_hidden=False)["利用者名"].tolist()
all_users_df = load_users(include_hidden=True)
all_users = all_users_df["利用者名"].tolist()


if not active_users:
    st.warning("表示中の利用者がいません。左メニューの「利用者マスタ管理」から利用者を追加してください。")


# =========================
# 健康チェック入力
# =========================
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
            temp = st.number_input(
                "体温",
                min_value=30.0,
                max_value=45.0,
                value=36.5,
                step=0.1,
            )

        with col5:
            bp_high = st.number_input(
                "血圧上",
                min_value=50,
                max_value=250,
                value=120,
                step=1,
            )

        with col6:
            bp_low = st.number_input(
                "血圧下",
                min_value=30,
                max_value=150,
                value=75,
                step=1,
            )

        col7, col8, col9 = st.columns(3)

        with col7:
            pulse = st.number_input(
                "脈拍",
                min_value=30,
                max_value=200,
                value=70,
                step=1,
            )

        with col8:
            spo2 = st.number_input(
                "SpO2",
                min_value=70,
                max_value=100,
                value=96,
                step=1,
            )

        with col9:
            weight = st.number_input(
                "体重",
                min_value=0.0,
                max_value=200.0,
                value=50.0,
                step=0.1,
            )

        family_memo = st.text_area(
            "家族共有メモ",
            placeholder="ご家族へ共有してよい内容を入力",
        )

        changes = st.text_area(
            "気になる変化",
            placeholder="食事、睡眠、歩行、表情、体調など",
        )

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


# =========================
# 入力データ確認
# =========================
elif menu == "入力データ確認":
    st.header("入力データ確認")

    df = load_data()

    filter_options = ["全員"] + all_users

    col1, col2, col3 = st.columns(3)

    with col1:
        filter_user = st.selectbox("利用者で絞り込み", filter_options)

    with col2:
        year = st.number_input(
            "年",
            min_value=2024,
            max_value=2035,
            value=date.today().year,
            step=1,
        )

    with col3:
        month = st.number_input(
            "月",
            min_value=1,
            max_value=12,
            value=date.today().month,
            step=1,
        )

    view = df.copy()

    if not view.empty:
        view["記録日"] = pd.to_datetime(view["記録日"], errors="coerce")
        view = view[
            (view["記録日"].dt.year == int(year))
            & (view["記録日"].dt.month == int(month))
        ]

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


# =========================
# 家族向けレポート作成
# =========================
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
        report_year = st.number_input(
            "対象年",
            min_value=2024,
            max_value=2035,
            value=date.today().year,
            step=1,
        )

    with col3:
        report_month = st.number_input(
            "対象月",
            min_value=1,
            max_value=12,
            value=date.today().month,
            step=1,
        )

    if st.button("家族向けレポートを作成"):
        report_path, target = create_family_report(
            df,
            report_user,
            report_year,
            report_month,
        )

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


# =========================
# 利用者マスタ管理
# =========================
elif menu == "利用者マスタ管理":
    st.header("利用者マスタ管理")
    st.caption("利用者の追加・入力候補からの削除ができます。削除しても過去の入力データは残ります。")

    df_users = load_users(include_hidden=True)

    st.subheader("現在の利用者一覧")
    st.dataframe(df_users, use_container_width=True)

    st.subheader("利用者を追加")

    with st.form("add_user_form", clear_on_submit=True):
        new_user = st.text_input(
            "追加する利用者名",
            placeholder="例：田中様",
        )
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
        restore_user = st.selectbox(
            "表示に戻す利用者",
            hidden_df["利用者名"].tolist(),
        )

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
