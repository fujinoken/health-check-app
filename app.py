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
# ログイン設定
# =========================
USERS = {
    "kanri": {
        "password": "rui",
        "role": "admin",
        "label": "管理者",
    },
    "staff": {
        "password": "rui",
        "role": "staff",
        "label": "職員",
    },
}


def login_check():
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False

    if "role" not in st.session_state:
        st.session_state.role = None

    if "user_label" not in st.session_state:
        st.session_state.user_label = ""

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
            user = USERS.get(input_id)

            if user and input_password == user["password"]:
                st.session_state.logged_in = True
                st.session_state.role = user["role"]
                st.session_state.user_label = user["label"]
                st.rerun()
            else:
                st.error("IDまたはパスワードが違います。")

    return False


def logout_button():
    with st.sidebar:
        st.caption(f"ログイン中：{st.session_state.user_label}")
        if st.button("ログアウト"):
            st.session_state.logged_in = False
            st.session_state.role = None
            st.session_state.user_label = ""
            st.rerun()


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
    df = df.copy()
    for col in COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df = df[COLUMNS]
    df.to_excel(DATA_FILE, index=False, sheet_name=SHEET_NAME)


def append_record(record):
    df = load_data()
    new_df = pd.DataFrame([record], columns=COLUMNS)
    df = pd.concat([df, new_df], ignore_index=True)
    save_data(df)


def to_number(series):
    return pd.to_numeric(series, errors="coerce")


def safe_float(value, default):
    try:
        if pd.isna(value) or value == "":
            return default
        return float(value)
    except Exception:
        return default


def safe_int(value, default):
    try:
        if pd.isna(value) or value == "":
            return default
        return int(float(value))
    except Exception:
        return default


def safe_text(value):
    if pd.isna(value):
        return ""
    return str(value)


def get_month_data(df, user_name, year, month):
    work = df.copy()
    work["記録日"] = pd.to_datetime(work["記録日"], errors="coerce")
    return work[
        (work["利用者名"] == user_name)
        & (work["記録日"].dt.year == int(year))
        & (work["記録日"].dt.month == int(month))
    ].sort_values("記録日")


def create_family_summary_text(target, user_name, year, month):
    """家族向けレポートに入れる、やわらかい自然文を作成する。
    医療判断・診断・改善断定は避け、記録に基づく共有に限定する。
    """
    if target.empty:
        return (
            f"{user_name}の{year}年{month}月分の記録は、現時点では登録されていません。"
            "今後の記録をもとに、ご様子を継続して確認していきます。"
        )

    lines = []
    record_count = len(target)
    lines.append(
        f"{user_name}の{year}年{month}月の記録は、{record_count}件確認されています。"
        "このレポートは医療的な判断ではなく、日々の健康チェック記録をもとにした共有です。"
    )

    temp_mean = to_number(target["体温"]).mean()
    spo2_mean = to_number(target["SpO2"]).mean()
    weight_mean = to_number(target["体重"]).mean()

    health_parts = []
    if not pd.isna(temp_mean):
        health_parts.append(f"体温は平均{round(float(temp_mean), 1)}℃")
    if not pd.isna(spo2_mean):
        health_parts.append(f"SpO2は平均{round(float(spo2_mean), 1)}％")
    if not pd.isna(weight_mean):
        health_parts.append(f"体重は平均{round(float(weight_mean), 1)}kg")

    if health_parts:
        lines.append(
            "記録上、" + "、".join(health_parts) + "として確認されています。"
            "数値は日々の状態を振り返るための目安として扱っています。"
        )

    # 注意が必要そうな記録を、断定せずに拾う
    alerts = []
    temp_alerts = target[to_number(target["体温"]) >= 37.5]
    spo2_alerts = target[to_number(target["SpO2"]) <= 93]
    bp_alerts = target[to_number(target["血圧上"]) >= 160]

    if not temp_alerts.empty:
        dates = "、".join(temp_alerts["記録日"].dt.strftime("%m/%d").tolist()[:5])
        alerts.append(f"{dates}に体温が37.5℃以上の記録")
    if not spo2_alerts.empty:
        dates = "、".join(spo2_alerts["記録日"].dt.strftime("%m/%d").tolist()[:5])
        alerts.append(f"{dates}にSpO2が93％以下の記録")
    if not bp_alerts.empty:
        dates = "、".join(bp_alerts["記録日"].dt.strftime("%m/%d").tolist()[:5])
        alerts.append(f"{dates}に血圧上が160以上の記録")

    if alerts:
        lines.append(
            "今月は、" + "、".join(alerts) + "がありました。"
            "一時的な変動の可能性もあるため、引き続き経過を見ながら確認していきます。"
        )
    else:
        lines.append(
            "記録上、設定した注意目安に該当する大きな変化は目立っていません。"
            "今後も日々の様子を継続して確認していきます。"
        )

    memo_rows = target[target["家族共有メモ"].fillna("").astype(str).str.strip() != ""]
    change_rows = target[target["気になる変化"].fillna("").astype(str).str.strip() != ""]

    if not memo_rows.empty:
        first = memo_rows.iloc[0]
        lines.append(
            f"ご様子として、{first['記録日'].strftime('%m/%d')}の記録に"
            f"「{str(first['家族共有メモ'])[:80]}」とあります。"
            "日々の関わりの中で、ご本人の様子を確認しています。"
        )

    if not change_rows.empty:
        first = change_rows.iloc[0]
        lines.append(
            f"また、{first['記録日'].strftime('%m/%d')}に"
            f"「{str(first['気になる変化'])[:80]}」という記録があります。"
            "必要に応じて職員間で共有しながら見守っています。"
        )

    lines.append(
        "今後も、数値だけでなく表情や生活の様子も含めて、安心して過ごせるよう見守ってまいります。"
    )

    return "\n\n".join(lines)


def get_today_dashboard(df, active_users):
    today = pd.Timestamp(date.today())

    if df.empty:
        today_df = pd.DataFrame(columns=COLUMNS)
    else:
        work = df.copy()
        work["記録日"] = pd.to_datetime(work["記録日"], errors="coerce")
        today_df = work[work["記録日"].dt.date == today.date()].copy()

    entered_users = today_df["利用者名"].dropna().astype(str).unique().tolist()
    missing_users = [user for user in active_users if user not in entered_users]

    temp_alert = today_df[to_number(today_df["体温"]) >= 37.5] if not today_df.empty else today_df
    spo2_alert = today_df[to_number(today_df["SpO2"]) <= 93] if not today_df.empty else today_df
    bp_alert = today_df[to_number(today_df["血圧上"]) >= 160] if not today_df.empty else today_df

    alert_rows = pd.concat([temp_alert, spo2_alert, bp_alert]).drop_duplicates() if not today_df.empty else today_df

    memo_count = 0
    if not today_df.empty:
        memo_count = int(
            (
                today_df["家族共有メモ"].fillna("").astype(str).str.strip() != ""
            ).sum()
        )

    return {
        "today_df": today_df,
        "entered_users": entered_users,
        "missing_users": missing_users,
        "alert_rows": alert_rows,
        "memo_count": memo_count,
    }


def show_dashboard(active_users):
    st.header("本日の管理者ダッシュボード")
    st.caption("今日の入力状況・未入力・注意記録を一画面で確認できます。")

    df = load_data()
    dashboard = get_today_dashboard(df, active_users)

    today_df = dashboard["today_df"]
    entered_count = len(dashboard["entered_users"])
    total_count = len(active_users)
    missing_users = dashboard["missing_users"]
    alert_rows = dashboard["alert_rows"]
    memo_count = dashboard["memo_count"]

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric("本日の入力済み", f"{entered_count} / {total_count} 名")

    with col2:
        st.metric("未入力", f"{len(missing_users)} 名")

    with col3:
        st.metric("注意記録", f"{len(alert_rows)} 件")

    with col4:
        st.metric("家族共有メモあり", f"{memo_count} 件")

    st.divider()

    col_left, col_right = st.columns(2)

    with col_left:
        st.subheader("未入力の利用者様")
        if missing_users:
            st.warning("、".join(missing_users))
        else:
            st.success("本日の入力は全員分そろっています。")

    with col_right:
        st.subheader("本日の記録件数")
        if today_df.empty:
            st.info("本日の記録はまだありません。")
        else:
            count_df = today_df.groupby("利用者名").size().reset_index(name="記録件数")
            st.dataframe(count_df, use_container_width=True, hide_index=True)

    st.subheader("注意記録の確認")
    st.caption("目安：体温37.5℃以上、SpO2 93％以下、血圧上160以上。医療判断ではなく、見落とし防止のための確認欄です。")

    if alert_rows.empty:
        st.success("本日、設定した注意目安に該当する記録はありません。")
    else:
        show_cols = [
            "記録日",
            "利用者名",
            "体温",
            "血圧上",
            "血圧下",
            "脈拍",
            "SpO2",
            "体重",
            "気になる変化",
            "入力者",
        ]
        st.dataframe(alert_rows[show_cols], use_container_width=True, hide_index=True)

    with st.expander("本日の全記録を表示する"):
        if today_df.empty:
            st.info("本日の記録はまだありません。")
        else:
            st.dataframe(today_df, use_container_width=True, hide_index=True)


# =========================
# 家族向けレポート作成
# =========================
def create_family_report(df, user_name, year, month):
    target = get_month_data(df, user_name, year, month)

    wb = Workbook()
    ws = wb.active
    ws.title = "家族向けレポート"

    thin = Side(style="thin", color="CCCCCC")
    border = Border(top=thin, bottom=thin, left=thin, right=thin)

    title_fill = PatternFill("solid", fgColor="EADFCB")
    section_fill = PatternFill("solid", fgColor="DDEFE2")
    note_fill = PatternFill("solid", fgColor="FFF8E7")
    summary_fill = PatternFill("solid", fgColor="F4F0E8")

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
    ws["A14"] = "今月のまとめ"
    ws["A14"].font = Font(name="Meiryo", bold=True)
    ws["A14"].fill = section_fill

    summary_text = create_family_summary_text(target, user_name, year, month)
    ws.merge_cells("A15:H20")
    ws["A15"] = summary_text
    ws["A15"].fill = summary_fill
    ws["A15"].alignment = Alignment(wrap_text=True, vertical="top")

    row = 22
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    ws.cell(row=row, column=1).value = "日付別のご様子（家族共有メモ）"
    ws.cell(row=row, column=1).font = Font(name="Meiryo", bold=True)
    ws.cell(row=row, column=1).fill = section_fill

    row += 1
    ws.cell(row=row, column=1).value = "日付"
    ws.cell(row=row, column=2).value = "家族共有メモ"
    ws.cell(row=row, column=1).fill = note_fill
    ws.cell(row=row, column=2).fill = note_fill

    row += 1
    memo_rows = target[target["家族共有メモ"].fillna("").astype(str).str.strip() != ""]

    if memo_rows.empty:
        ws.cell(row=row, column=1).value = ""
        ws.cell(row=row, column=2).value = "記録された家族共有メモはありません。"
        row += 1
    else:
        for _, rec in memo_rows.iterrows():
            ws.cell(row=row, column=1).value = rec["記録日"].strftime("%m/%d")
            ws.cell(row=row, column=2).value = str(rec["家族共有メモ"])
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
    ws.merge_cells(start_row=row, start_column=1, end_row=row + 2, end_column=8)
    ws.cell(row=row, column=1).value = (
        "※このレポートは、施設内の健康チェック記録をもとにした共有資料です。"
        "医療的な診断・治療効果の判断を行うものではありません。"
    )
    ws.cell(row=row, column=1).alignment = Alignment(wrap_text=True, vertical="top")

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

    return report_path, target, summary_text


# =========================
# アプリ本体
# =========================
logout_button()

st.title("健康チェックWebアプリ")
st.caption("入力はブラウザで、データはExcelへ保存。管理者ダッシュボード、過去データ管理、家族向けレポート出力ができます。")

ensure_data_file()
ensure_user_file()


# =========================
# サイドバーメニュー
# 管理者と職員で表示を分ける
# =========================
if st.session_state.role == "admin":
    st.sidebar.success("管理者モード")
    menu_items = [
        "管理者ダッシュボード",
        "健康チェック入力",
        "過去データ管理",
        "入力データ確認",
        "家族向けレポート作成",
        "利用者マスタ管理",
    ]
else:
    st.sidebar.info("職員モード")
    menu_items = [
        "健康チェック入力",
        "過去データ管理",
    ]

menu = st.sidebar.radio(
    "メニュー",
    menu_items,
)


active_users = load_users(include_hidden=False)["利用者名"].tolist()
all_users_df = load_users(include_hidden=True)
all_users = all_users_df["利用者名"].tolist()


if not active_users:
    st.warning("表示中の利用者がいません。管理者に利用者マスタの確認を依頼してください。")

if st.session_state.role == "staff":
    st.info("お疲れ様です。今日の健康チェック入力をお願いします。入力後は、必要に応じて過去データ管理から確認・修正できます。")
elif st.session_state.role == "admin":
    st.success("管理者としてログインしています。入力状況・注意記録・レポート作成を確認できます。")


# =========================
# 管理者ダッシュボード
# =========================
if menu == "管理者ダッシュボード":
    show_dashboard(active_users)


# =========================
# 健康チェック入力
# =========================
elif menu == "健康チェック入力":
    st.header("健康チェック入力")

    if st.session_state.role == "staff":
        st.markdown("### お疲れ様です。")
        st.write("利用者様の今日の健康状態を、わかる範囲で落ち着いて入力してください。")

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
# 過去データ管理
# 検索・更新・削除
# =========================
elif menu == "過去データ管理":
    st.header("過去データ管理")
    st.caption("過去に登録した健康チェック記録を検索し、修正・削除できます。")

    if st.session_state.role == "staff":
        st.info("お疲れ様です。入力内容の確認・修正が必要な場合はこちらから行えます。削除は慎重に行ってください。")

    df = load_data()

    if df.empty:
        st.info("まだ登録データがありません。")
        st.stop()

    df["記録日"] = pd.to_datetime(df["記録日"], errors="coerce")

    # 元のExcel上の行番号を管理IDとして保持
    df = df.reset_index(drop=False)
    df = df.rename(columns={"index": "管理ID"})

    st.subheader("検索条件")

    col1, col2, col3 = st.columns(3)

    with col1:
        search_user = st.selectbox(
            "利用者名",
            ["全員"] + all_users,
        )

    with col2:
        search_year = st.number_input(
            "年",
            min_value=2024,
            max_value=2035,
            value=date.today().year,
            step=1,
        )

    with col3:
        search_month = st.number_input(
            "月",
            min_value=1,
            max_value=12,
            value=date.today().month,
            step=1,
        )

    search_text = st.text_input(
        "メモ検索",
        placeholder="家族共有メモ・気になる変化・入力者から検索できます",
    )

    result = df.copy()

    result = result[
        (result["記録日"].dt.year == int(search_year))
        & (result["記録日"].dt.month == int(search_month))
    ]

    if search_user != "全員":
        result = result[result["利用者名"] == search_user]

    if search_text.strip():
        keyword = search_text.strip()
        result = result[
            result["家族共有メモ"].fillna("").astype(str).str.contains(keyword, case=False, na=False)
            | result["気になる変化"].fillna("").astype(str).str.contains(keyword, case=False, na=False)
            | result["入力者"].fillna("").astype(str).str.contains(keyword, case=False, na=False)
        ]

    st.subheader("検索結果")
    st.dataframe(result, use_container_width=True)

    if result.empty:
        st.warning("該当するデータがありません。")
        st.stop()

    st.divider()

    st.subheader("データの更新・削除")

    selected_id = st.selectbox(
        "修正・削除する管理IDを選択",
        result["管理ID"].tolist(),
    )

    selected_row = df[df["管理ID"] == selected_id].iloc[0]

    with st.expander("選択中のデータを確認する", expanded=True):
        st.dataframe(
            pd.DataFrame([selected_row]),
            use_container_width=True,
        )

    with st.form("edit_record_form"):
        edit_date = st.date_input(
            "記録日",
            value=selected_row["記録日"].date()
            if pd.notna(selected_row["記録日"])
            else date.today(),
        )

        edit_user = st.selectbox(
            "利用者名",
            all_users,
            index=all_users.index(selected_row["利用者名"])
            if selected_row["利用者名"] in all_users
            else 0,
        )

        col4, col5, col6 = st.columns(3)

        with col4:
            edit_temp = st.number_input(
                "体温",
                min_value=30.0,
                max_value=45.0,
                value=safe_float(selected_row["体温"], 36.5),
                step=0.1,
            )

        with col5:
            edit_bp_high = st.number_input(
                "血圧上",
                min_value=50,
                max_value=250,
                value=safe_int(selected_row["血圧上"], 120),
                step=1,
            )

        with col6:
            edit_bp_low = st.number_input(
                "血圧下",
                min_value=30,
                max_value=150,
                value=safe_int(selected_row["血圧下"], 75),
                step=1,
            )

        col7, col8, col9 = st.columns(3)

        with col7:
            edit_pulse = st.number_input(
                "脈拍",
                min_value=30,
                max_value=200,
                value=safe_int(selected_row["脈拍"], 70),
                step=1,
            )

        with col8:
            edit_spo2 = st.number_input(
                "SpO2",
                min_value=70,
                max_value=100,
                value=safe_int(selected_row["SpO2"], 96),
                step=1,
            )

        with col9:
            edit_weight = st.number_input(
                "体重",
                min_value=0.0,
                max_value=200.0,
                value=safe_float(selected_row["体重"], 50.0),
                step=0.1,
            )

        edit_family_memo = st.text_area(
            "家族共有メモ",
            value=safe_text(selected_row["家族共有メモ"]),
        )

        edit_changes = st.text_area(
            "気になる変化",
            value=safe_text(selected_row["気になる変化"]),
        )

        edit_staff = st.text_input(
            "入力者",
            value=safe_text(selected_row["入力者"]),
        )

        update_submit = st.form_submit_button("この内容で更新する")

    if update_submit:
        original_df = load_data()

        original_df.loc[selected_id, "記録日"] = edit_date
        original_df.loc[selected_id, "利用者名"] = edit_user
        original_df.loc[selected_id, "体温"] = edit_temp
        original_df.loc[selected_id, "血圧上"] = edit_bp_high
        original_df.loc[selected_id, "血圧下"] = edit_bp_low
        original_df.loc[selected_id, "脈拍"] = edit_pulse
        original_df.loc[selected_id, "SpO2"] = edit_spo2
        original_df.loc[selected_id, "体重"] = edit_weight
        original_df.loc[selected_id, "家族共有メモ"] = edit_family_memo
        original_df.loc[selected_id, "気になる変化"] = edit_changes
        original_df.loc[selected_id, "入力者"] = edit_staff
        original_df.loc[selected_id, "登録日時"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        save_data(original_df)

        st.success("データを更新しました。")
        st.rerun()

    st.divider()

    st.subheader("データ削除")
    st.warning("削除すると元に戻せません。必要に応じて先にExcelをダウンロードしてください。")

    delete_check = st.checkbox("このデータを削除することを確認しました")

    if st.button("選択したデータを削除する"):
        if not delete_check:
            st.error("削除する場合は確認チェックを入れてください。")
        else:
            original_df = load_data()
            original_df = original_df.drop(index=selected_id).reset_index(drop=True)
            save_data(original_df)

            st.success("データを削除しました。")
            st.rerun()


# =========================
# 入力データ確認
# =========================
elif menu == "入力データ確認":
    if st.session_state.role != "admin":
        st.error("この画面は管理者専用です。")
        st.stop()

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
    if st.session_state.role != "admin":
        st.error("この画面は管理者専用です。")
        st.stop()

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
        report_path, target, summary_text = create_family_report(
            df,
            report_user,
            report_year,
            report_month,
        )

        if target.empty:
            st.warning("指定した利用者・年月のデータがありません。空のレポートを作成しました。")
        else:
            st.success("家族向けレポートを作成しました。")

        st.subheader("レポート文章プレビュー")
        st.info(summary_text)

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
    if st.session_state.role != "admin":
        st.error("この画面は管理者専用です。")
        st.stop()

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
