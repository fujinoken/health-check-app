import streamlit as st
import pandas as pd
import json
import hashlib
from pathlib import Path
from datetime import date, datetime
from io import BytesIO

try:
    import matplotlib.pyplot as plt
except Exception:
    plt = None

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
except Exception:
    colors = None


# =========================
# ページ設定
# =========================
st.set_page_config(
    page_title="ひだまり 健康チェック管理システム",
    page_icon="🌿",
    layout="wide",
)


# =========================
# ログイン設定
# =========================
USERS = {
    "kanri": {"password": "rui", "role": "admin", "label": "管理者"},
    "staff": {"password": "rui", "role": "staff", "label": "職員"},
}


# =========================
# ファイル設定
# =========================
DATA_DIR = Path("data")
REPORT_DIR = Path("reports")

HEALTH_FILE = DATA_DIR / "health_data.xlsx"
EXCRETION_FILE = DATA_DIR / "excretion_data.xlsx"
USER_FILE = DATA_DIR / "user_master.xlsx"

HEALTH_SHEET = "健康チェック"
EXCRETION_SHEET = "排泄チェック"
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

ASSESSMENT_COLUMNS = [
    "基本情報",
    "主訴",
    "生活状況",
    "ADL",
    "IADL",
    "認知機能",
    "健康状態",
    "課題",
    "支援内容",
]

HEALTH_COLUMNS = [
    "記録日",
    "利用者名",
    "体温",
    "血圧上",
    "血圧下",
    "脈拍",
    "SpO2",
    "体重",
    "朝食摂取率",
    "昼食摂取率",
    "夕食摂取率",
    "家族共有メモ",
    "気になる変化",
    "登録日時",
    "入力者",
]

EXCRETION_SLOTS = [
    ("午前", "9時〜12時"),
    ("午後", "12時〜15時"),
    ("夕方", "15時〜17時"),
    ("夜", "18時〜22時"),
    ("深夜", "22時〜5時"),
    ("朝方", "5時〜8時"),
]

URINE_AMOUNT_OPTIONS = ["なし", "少", "中", "大"]
URINE_TYPE_OPTIONS = ["なし", "普通尿", "濃縮尿"]
STOOL_AMOUNT_OPTIONS = ["なし", "少", "中", "大"]
STOOL_TYPE_OPTIONS = ["なし", "普通便", "下痢便", "水様便"]

EXCRETION_COLUMNS = [
    "記録日",
    "利用者名",
    "時間帯",
    "時間帯目安",
    "尿量",
    "尿性状",
    "便量",
    "便性状",
    "排泄メモ",
    "入力者",
    "登録日時",
]

USER_COLUMNS = ["利用者名", "表示"] + ASSESSMENT_COLUMNS


# =========================
# 共通関数
# =========================
def ensure_dirs():
    DATA_DIR.mkdir(exist_ok=True)
    REPORT_DIR.mkdir(exist_ok=True)


def clean_text(value, default=""):
    try:
        if pd.isna(value):
            return default
    except Exception:
        pass

    text = str(value).strip()
    if text.lower() in ["nan", "none", "nat"]:
        return default
    return text


def safe_float(value, default=0.0):
    try:
        if pd.isna(value) or value == "":
            return default
        return float(value)
    except Exception:
        return default


def safe_int(value, default=0):
    try:
        if pd.isna(value) or value == "":
            return default
        return int(float(value))
    except Exception:
        return default


def to_number(series):
    return pd.to_numeric(series, errors="coerce")


def make_date_user_key(record_date, user_name):
    d = pd.to_datetime(record_date, errors="coerce")
    if pd.isna(d):
        return ""
    return f"{d.strftime('%Y-%m-%d')}__{clean_text(user_name)}"


def make_excretion_key(record_date, user_name, slot):
    d = pd.to_datetime(record_date, errors="coerce")
    if pd.isna(d):
        return ""
    return f"{d.strftime('%Y-%m-%d')}__{clean_text(user_name)}__{clean_text(slot)}"


def get_option_index(options, value, default="なし"):
    value = clean_text(value, default)
    if value in options:
        return options.index(value)
    if default in options:
        return options.index(default)
    return 0


def ensure_excel_file(path, sheet_name, columns):
    ensure_dirs()
    if not path.exists():
        pd.DataFrame(columns=columns).to_excel(path, index=False, sheet_name=sheet_name)


# =========================
# 利用者マスタ
# =========================
def ensure_user_file():
    ensure_dirs()
    if not USER_FILE.exists():
        data = {"利用者名": DEFAULT_USERS, "表示": ["表示"] * len(DEFAULT_USERS)}
        for col in ASSESSMENT_COLUMNS:
            data[col] = [""] * len(DEFAULT_USERS)
        pd.DataFrame(data, columns=USER_COLUMNS).to_excel(USER_FILE, index=False, sheet_name=USER_SHEET)


def load_users(include_hidden=False):
    ensure_user_file()

    try:
        df = pd.read_excel(USER_FILE, sheet_name=USER_SHEET)
    except Exception:
        df = pd.DataFrame(columns=USER_COLUMNS)

    for col in USER_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    df = df[USER_COLUMNS].copy()
    df["利用者名"] = df["利用者名"].astype(str).str.strip()
    df = df[df["利用者名"] != ""].drop_duplicates(subset=["利用者名"], keep="first")

    if not include_hidden:
        df = df[df["表示"].fillna("表示") == "表示"]

    return df.reset_index(drop=True)


def save_users(df):
    ensure_dirs()
    df = df.copy()

    for col in USER_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    df = df[USER_COLUMNS]
    df["利用者名"] = df["利用者名"].astype(str).str.strip()
    df = df[df["利用者名"] != ""].drop_duplicates(subset=["利用者名"], keep="first")
    df.to_excel(USER_FILE, index=False, sheet_name=USER_SHEET)


def add_user(user_name):
    user_name = clean_text(user_name)

    if not user_name:
        return False, "利用者名を入力してください。"

    df = load_users(include_hidden=True)

    if user_name in df["利用者名"].tolist():
        df.loc[df["利用者名"] == user_name, "表示"] = "表示"
        save_users(df)
        return True, f"{user_name}を表示に戻しました。"

    row = {"利用者名": user_name, "表示": "表示"}
    for col in ASSESSMENT_COLUMNS:
        row[col] = ""

    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    save_users(df)

    return True, f"{user_name}を追加しました。"


def hide_user(user_name):
    df = load_users(include_hidden=True)

    if user_name not in df["利用者名"].tolist():
        return False, "対象の利用者が見つかりません。"

    df.loc[df["利用者名"] == user_name, "表示"] = "非表示"
    save_users(df)

    return True, f"{user_name}を入力候補から外しました。"


def get_user_assessment(user_name):
    df = load_users(include_hidden=True)
    row = df[df["利用者名"] == user_name]

    if row.empty:
        return {}

    row = row.iloc[0]

    return {
        col: clean_text(row.get(col, ""))
        for col in ASSESSMENT_COLUMNS
        if clean_text(row.get(col, ""))
    }


def build_assessment_context_text(user_name):
    data = get_user_assessment(user_name)

    if not data:
        return ""

    order = ["主訴", "生活状況", "ADL", "IADL", "認知機能", "健康状態", "課題", "支援内容"]
    lines = []

    for col in order:
        if data.get(col):
            lines.append(f"{col}：{data[col]}")

    return "\n".join(lines)


# =========================
# 健康チェックデータ
# =========================
def ensure_health_file():
    ensure_excel_file(HEALTH_FILE, HEALTH_SHEET, HEALTH_COLUMNS)


def load_health_data():
    ensure_health_file()

    try:
        df = pd.read_excel(HEALTH_FILE, sheet_name=HEALTH_SHEET)
    except Exception:
        df = pd.DataFrame(columns=HEALTH_COLUMNS)

    for col in HEALTH_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    df = df[HEALTH_COLUMNS].copy()

    if not df.empty:
        df["記録日"] = pd.to_datetime(df["記録日"], errors="coerce")
        df["利用者名"] = df["利用者名"].astype(str).str.strip()

    return df.astype("object")


def save_health_data(df):
    ensure_dirs()
    df = df.copy()

    for col in HEALTH_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    df = df[HEALTH_COLUMNS].astype("object")

    if not df.empty:
        df["記録日"] = pd.to_datetime(df["記録日"], errors="coerce")
        df["利用者名"] = df["利用者名"].astype(str).str.strip()
        df["_key"] = df.apply(lambda row: make_date_user_key(row["記録日"], row["利用者名"]), axis=1)
        df = df[df["_key"] != ""]
        df = df.drop_duplicates(subset=["_key"], keep="last")
        df = df.drop(columns=["_key"])

    df.to_excel(HEALTH_FILE, index=False, sheet_name=HEALTH_SHEET)


def find_health_index(df, record_date, user_name):
    if df.empty:
        return None

    work = df.copy()
    work["記録日"] = pd.to_datetime(work["記録日"], errors="coerce")
    work["利用者名"] = work["利用者名"].astype(str).str.strip()

    target_date = pd.to_datetime(record_date, errors="coerce")
    if pd.isna(target_date):
        return None

    mask = (work["記録日"].dt.date == target_date.date()) & (work["利用者名"] == clean_text(user_name))
    matches = work.index[mask].tolist()

    if not matches:
        return None

    return matches[0]


def upsert_health_record(record):
    df = load_health_data()
    df = df.astype("object")

    idx = find_health_index(df, record["記録日"], record["利用者名"])

    if idx is None:
        new_df = pd.DataFrame([record], columns=HEALTH_COLUMNS).astype("object")
        df = pd.concat([df, new_df], ignore_index=True)
        action = "登録"
    else:
        for col in HEALTH_COLUMNS:
            df.at[idx, col] = record.get(col, "")
        action = "更新"

    save_health_data(df)

    return action


def get_month_health_data(df, user_name, year, month):
    if df.empty:
        return df

    work = df.copy()
    work["記録日"] = pd.to_datetime(work["記録日"], errors="coerce")

    return work[
        (work["利用者名"] == user_name)
        & (work["記録日"].dt.year == int(year))
        & (work["記録日"].dt.month == int(month))
    ].sort_values("記録日")


# =========================
# 排泄チェックデータ
# =========================
def ensure_excretion_file():
    ensure_excel_file(EXCRETION_FILE, EXCRETION_SHEET, EXCRETION_COLUMNS)


def normalize_excretion_record(record):
    urine_amount = clean_text(record.get("尿量", "なし"), "なし")
    urine_type = clean_text(record.get("尿性状", "なし"), "なし")
    stool_amount = clean_text(record.get("便量", "なし"), "なし")
    stool_type = clean_text(record.get("便性状", "なし"), "なし")

    if urine_amount == "":
        urine_amount = "なし"
    if urine_type == "":
        urine_type = "なし"
    if stool_amount == "":
        stool_amount = "なし"
    if stool_type == "":
        stool_type = "なし"

    if urine_amount == "なし":
        urine_type = "なし"

    if stool_amount == "なし":
        stool_type = "なし"

    record["尿量"] = urine_amount
    record["尿性状"] = urine_type
    record["便量"] = stool_amount
    record["便性状"] = stool_type

    return record


def load_excretion_data():
    ensure_excretion_file()

    try:
        df = pd.read_excel(EXCRETION_FILE, sheet_name=EXCRETION_SHEET)
    except Exception:
        df = pd.DataFrame(columns=EXCRETION_COLUMNS)

    for col in EXCRETION_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    df = df[EXCRETION_COLUMNS].copy()

    if not df.empty:
        df["記録日"] = pd.to_datetime(df["記録日"], errors="coerce")
        for col in ["利用者名", "時間帯", "時間帯目安", "尿量", "尿性状", "便量", "便性状", "排泄メモ", "入力者", "登録日時"]:
            df[col] = df[col].fillna("").astype(str)

    return df.astype("object")


def save_excretion_data(df):
    ensure_dirs()
    df = df.copy()

    for col in EXCRETION_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    records = []
    for _, row in df.iterrows():
        rec = row.to_dict()
        rec = normalize_excretion_record(rec)
        records.append(rec)

    df = pd.DataFrame(records, columns=EXCRETION_COLUMNS).astype("object")

    if not df.empty:
        df["記録日"] = pd.to_datetime(df["記録日"], errors="coerce")
        df["利用者名"] = df["利用者名"].astype(str).str.strip()
        df["時間帯"] = df["時間帯"].astype(str).str.strip()
        df["_key"] = df.apply(lambda row: make_excretion_key(row["記録日"], row["利用者名"], row["時間帯"]), axis=1)
        df = df[df["_key"] != ""]
        df = df.drop_duplicates(subset=["_key"], keep="last")
        df = df.drop(columns=["_key"])

    df = df[EXCRETION_COLUMNS]
    df.to_excel(EXCRETION_FILE, index=False, sheet_name=EXCRETION_SHEET)


def find_excretion_index(df, record_date, user_name, slot):
    if df.empty:
        return None

    work = df.copy()
    work["記録日"] = pd.to_datetime(work["記録日"], errors="coerce")
    work["利用者名"] = work["利用者名"].astype(str).str.strip()
    work["時間帯"] = work["時間帯"].astype(str).str.strip()

    target_date = pd.to_datetime(record_date, errors="coerce")
    if pd.isna(target_date):
        return None

    mask = (
        (work["記録日"].dt.date == target_date.date())
        & (work["利用者名"] == clean_text(user_name))
        & (work["時間帯"] == clean_text(slot))
    )

    matches = work.index[mask].tolist()

    if not matches:
        return None

    return matches[0]


def get_excretion_row(df, record_date, user_name, slot):
    idx = find_excretion_index(df, record_date, user_name, slot)

    if idx is None:
        return None

    return df.loc[idx]


def upsert_excretion_record(record):
    record = normalize_excretion_record(record)

    df = load_excretion_data()
    idx = find_excretion_index(
        df,
        record["記録日"],
        record["利用者名"],
        record["時間帯"],
    )

    if idx is None:
        df = pd.concat(
            [df, pd.DataFrame([record], columns=EXCRETION_COLUMNS).astype("object")],
            ignore_index=True,
        )
        action = "登録"
    else:
        for col in EXCRETION_COLUMNS:
            df.at[idx, col] = record.get(col, "")
        action = "更新"

    save_excretion_data(df)

    return action


def get_day_excretion_data(df, record_date, user_name=None):
    if df.empty:
        return df

    work = df.copy()
    work["記録日"] = pd.to_datetime(work["記録日"], errors="coerce")

    target_date = pd.to_datetime(record_date, errors="coerce")
    if pd.isna(target_date):
        return pd.DataFrame(columns=EXCRETION_COLUMNS)

    work = work[work["記録日"].dt.date == target_date.date()]

    if user_name and user_name != "全員":
        work = work[work["利用者名"] == user_name]

    slot_order = {slot: i for i, (slot, _) in enumerate(EXCRETION_SLOTS)}
    work["_slot_order"] = work["時間帯"].map(slot_order).fillna(99)
    work = work.sort_values(["利用者名", "_slot_order"]).drop(columns=["_slot_order"])

    return work


def get_month_excretion_data(df, user_name, year, month):
    if df.empty:
        return df

    work = df.copy()
    work["記録日"] = pd.to_datetime(work["記録日"], errors="coerce")

    return work[
        (work["利用者名"] == user_name)
        & (work["記録日"].dt.year == int(year))
        & (work["記録日"].dt.month == int(month))
    ].sort_values(["記録日", "時間帯"])


def summarize_excretion(df):
    if df.empty:
        return {
            "排尿回数": 0,
            "排便回数": 0,
            "濃縮尿": 0,
            "下痢便": 0,
            "水様便": 0,
            "排便なし枠": 0,
        }

    return {
        "排尿回数": int((df["尿量"].fillna("なし") != "なし").sum()),
        "排便回数": int((df["便量"].fillna("なし") != "なし").sum()),
        "濃縮尿": int((df["尿性状"].fillna("") == "濃縮尿").sum()),
        "下痢便": int((df["便性状"].fillna("") == "下痢便").sum()),
        "水様便": int((df["便性状"].fillna("") == "水様便").sum()),
        "排便なし枠": int((df["便量"].fillna("なし") == "なし").sum()),
    }


def build_excretion_text(df):
    if df.empty:
        return "排泄記録はありません。"

    lines = []

    for _, row in df.iterrows():
        lines.append(
            f"{row['記録日'].strftime('%m/%d') if pd.notna(row['記録日']) else ''} "
            f"{row['時間帯']}：尿 {row['尿量']}・{row['尿性状']} ／ 便 {row['便量']}・{row['便性状']}"
        )

    return "\n".join(lines)


# =========================
# レポート系
# =========================
def create_family_summary_text(health_df, excretion_df, user_name, year, month):
    target = get_month_health_data(health_df, user_name, year, month)
    ex_target = get_month_excretion_data(excretion_df, user_name, year, month)

    lines = []

    if target.empty:
        lines.append(f"{user_name}の{year}年{month}月分の健康チェック記録は、現時点では登録されていません。")
    else:
        lines.append(
            f"{user_name}の{year}年{month}月の健康チェック記録は、{len(target)}件確認されています。"
            "この文章は医療的な判断ではなく、日々の記録をもとにした共有です。"
        )

        temp_mean = to_number(target["体温"]).mean()
        spo2_mean = to_number(target["SpO2"]).mean()
        weight_mean = to_number(target["体重"]).mean()

        health_parts = []
        if not pd.isna(temp_mean):
            health_parts.append(f"体温平均{round(float(temp_mean), 1)}℃")
        if not pd.isna(spo2_mean):
            health_parts.append(f"SpO2平均{round(float(spo2_mean), 1)}％")
        if not pd.isna(weight_mean):
            health_parts.append(f"体重平均{round(float(weight_mean), 1)}kg")

        if health_parts:
            lines.append("記録上、" + "、".join(health_parts) + "として確認されています。")

        meal_parts = []
        for label in ["朝食摂取率", "昼食摂取率", "夕食摂取率"]:
            mean = to_number(target[label]).mean()
            if not pd.isna(mean):
                meal_parts.append(f"{label.replace('摂取率', '')}平均{round(float(mean), 1)}％")

        if meal_parts:
            lines.append("食事摂取率は、" + "、".join(meal_parts) + "でした。")

        memo_rows = target[target["家族共有メモ"].fillna("").astype(str).str.strip() != ""]
        change_rows = target[target["気になる変化"].fillna("").astype(str).str.strip() != ""]

        if not memo_rows.empty:
            first = memo_rows.iloc[0]
            lines.append(
                f"ご様子として、{first['記録日'].strftime('%m/%d')}の記録に"
                f"「{str(first['家族共有メモ'])[:80]}」とあります。"
            )

        if not change_rows.empty:
            first = change_rows.iloc[0]
            lines.append(
                f"また、{first['記録日'].strftime('%m/%d')}に"
                f"「{str(first['気になる変化'])[:80]}」という記録があります。"
                "必要に応じて職員間で共有しながら見守っています。"
            )

    assessment = build_assessment_context_text(user_name)
    if assessment:
        lines.append("アセスメント情報もふまえ、生活全体の様子を確認しています。\n" + assessment)

    if ex_target.empty:
        lines.append("排泄記録は、対象月にはまだ登録されていません。")
    else:
        ex_sum = summarize_excretion(ex_target)
        lines.append(
            "排泄状況は、"
            f"排尿記録{ex_sum['排尿回数']}回、排便記録{ex_sum['排便回数']}回、"
            f"濃縮尿{ex_sum['濃縮尿']}回、下痢便{ex_sum['下痢便']}回、水様便{ex_sum['水様便']}回として記録されています。"
        )

    lines.append("今後も、数値だけでなく表情や生活の様子も含めて、安心して過ごせるよう見守ってまいります。")

    return "\n\n".join(lines)


def create_hidamari_pdf(health_df, excretion_df, user_name, year, month):
    if colors is None:
        raise RuntimeError("reportlab が利用できません。requirements.txt に reportlab を追加してください。")

    pdfmetrics.registerFont(UnicodeCIDFont("HeiseiKakuGo-W5"))
    pdfmetrics.registerFont(UnicodeCIDFont("HeiseiMin-W3"))

    file_path = REPORT_DIR / f"ひだまりレポート_{user_name}_{year}年{month}月.pdf"

    doc = SimpleDocTemplate(
        str(file_path),
        pagesize=A4,
        rightMargin=17 * mm,
        leftMargin=17 * mm,
        topMargin=14 * mm,
        bottomMargin=14 * mm,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "jp_title",
        parent=styles["Title"],
        fontName="HeiseiKakuGo-W5",
        fontSize=22,
        leading=28,
        alignment=1,
        textColor=colors.HexColor("#2F3437"),
    )
    h2_style = ParagraphStyle(
        "jp_h2",
        parent=styles["Heading2"],
        fontName="HeiseiKakuGo-W5",
        fontSize=13,
        leading=18,
        textColor=colors.HexColor("#2F3437"),
    )
    body_style = ParagraphStyle(
        "jp_body",
        parent=styles["BodyText"],
        fontName="HeiseiMin-W3",
        fontSize=10,
        leading=16,
    )
    small_style = ParagraphStyle(
        "jp_small",
        parent=styles["BodyText"],
        fontName="HeiseiMin-W3",
        fontSize=8,
        leading=12,
        textColor=colors.HexColor("#666666"),
    )

    story = []
    story.append(Paragraph("ひだまりレポート", title_style))
    story.append(Spacer(1, 6))
    story.append(Paragraph(f"{user_name}　{year}年{month}月", body_style))
    story.append(Spacer(1, 12))

    summary = create_family_summary_text(health_df, excretion_df, user_name, year, month)
    story.append(Paragraph("今月のまとめ", h2_style))
    for para in summary.split("\n\n"):
        story.append(Paragraph(para.replace("\n", "<br/>"), body_style))
        story.append(Spacer(1, 5))

    story.append(PageBreak())
    story.append(Paragraph("排泄記録", h2_style))

    ex_target = get_month_excretion_data(excretion_df, user_name, year, month)
    if ex_target.empty:
        story.append(Paragraph("対象月の排泄記録はありません。", body_style))
    else:
        table_data = [["日付", "時間帯", "尿", "便", "メモ"]]
        for _, row in ex_target.iterrows():
            table_data.append([
                row["記録日"].strftime("%m/%d") if pd.notna(row["記録日"]) else "",
                row["時間帯"],
                f"{row['尿量']}・{row['尿性状']}",
                f"{row['便量']}・{row['便性状']}",
                str(row.get("排泄メモ", ""))[:40],
            ])

        table = Table(table_data, colWidths=[20*mm, 24*mm, 35*mm, 35*mm, 55*mm])
        table.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, -1), "HeiseiMin-W3"),
            ("FONTNAME", (0, 0), (-1, 0), "HeiseiKakuGo-W5"),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F7F4EE")),
            ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#D9D9D9")),
            ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#D9D9D9")),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]))
        story.append(table)

    story.append(Spacer(1, 12))
    story.append(Paragraph("※このレポートは施設内の記録をもとにした共有資料です。医療的な診断・治療効果の判断を行うものではありません。", small_style))

    doc.build(story)
    return file_path


def create_handover_text(health_df, excretion_df, target_date):
    lines = [
        f"{target_date.strftime('%Y/%m/%d')}の申し送りまとめです。",
        "医療的な判断ではなく、記録内容をもとにした共有用メモです。",
        "",
    ]

    h = health_df.copy()
    if not h.empty:
        h["記録日"] = pd.to_datetime(h["記録日"], errors="coerce")
        h = h[h["記録日"].dt.date == target_date]

        for _, row in h.iterrows():
            notes = []
            if clean_text(row.get("気になる変化", "")):
                notes.append(f"気になる変化：{row.get('気になる変化')}")
            if clean_text(row.get("家族共有メモ", "")):
                notes.append(f"家族共有メモ：{row.get('家族共有メモ')}")

            vital_alerts = []
            if safe_float(row.get("体温"), 0) >= 37.5:
                vital_alerts.append("体温高め")
            if safe_int(row.get("SpO2"), 100) <= 93:
                vital_alerts.append("SpO2低め")
            if safe_int(row.get("血圧上"), 0) >= 160:
                vital_alerts.append("血圧上高め")

            if vital_alerts:
                notes.append("確認目安：" + "、".join(vital_alerts))

            if notes:
                lines.append(f"■ {row.get('利用者名')}")
                lines.extend([f"・{x}" for x in notes])
                lines.append("")

    e = get_day_excretion_data(excretion_df, target_date, None)
    if not e.empty:
        for user in e["利用者名"].dropna().unique():
            user_ex = e[e["利用者名"] == user]
            alerts = []

            for _, row in user_ex.iterrows():
                if row["尿性状"] == "濃縮尿":
                    alerts.append(f"{row['時間帯']}に濃縮尿")
                if row["便性状"] in ["下痢便", "水様便"]:
                    alerts.append(f"{row['時間帯']}に{row['便性状']}")

            if alerts:
                lines.append(f"■ {user} 排泄確認")
                lines.append("・" + "、".join(alerts))
                lines.append("")

    if len(lines) <= 3:
        lines.append("記録上、特に申し送り対象となるメモや注意目安はありません。")

    lines.append("引き続き、普段との違いがないかを確認しながら見守ります。")
    return "\n".join(lines)


# =========================
# ログイン・デザイン
# =========================
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
        <div style='text-align:center; padding:24px;'>
            <h1 style='color:#2E7D32;'>ひだまり 健康チェック管理システム</h1>
            <p style='color:#666;'>利用者様の健康記録・排泄記録を安全に管理します</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    col1, col2, col3 = st.columns([1, 2, 1])

    with col2:
        st.markdown("### ログイン")
        input_id = st.text_input("ID")
        input_password = st.text_input("パスワード", type="password")

        if st.button("ログイン", use_container_width=True):
            login_id = clean_text(input_id).lower()
            login_password = clean_text(input_password)
            user = USERS.get(login_id)

            if user and login_password == user["password"]:
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


def apply_design():
    if st.session_state.role == "staff":
        bg = "#FFFDF7"
        accent = "#D97A6A"
        box = "#FFF1E8"
    else:
        bg = "#F4F6F9"
        accent = "#1F4E79"
        box = "#EAF1F8"

    st.markdown(
        f"""
        <style>
        .stApp {{
            background-color: {bg};
        }}
        h1, h2, h3 {{
            color: {accent};
        }}
        .info-box {{
            background: {box};
            padding: 14px 18px;
            border-radius: 14px;
            border: 1px solid rgba(0,0,0,0.08);
            margin-bottom: 12px;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


if not login_check():
    st.stop()

apply_design()
logout_button()

st.title("健康チェックWebアプリ")
st.caption("管理者支援・職員入力・家族共有を一体化した健康チェックシステムです。")

if st.session_state.role == "admin":
    st.success("管理者モード")
else:
    st.info("お疲れ様です。今日の健康チェック入力・排泄チェック入力をお願いします。")


# =========================
# メニュー
# =========================
users_df = load_users(include_hidden=False)
active_users = users_df["利用者名"].tolist()
all_users = active_users

if st.session_state.role == "admin":
    menu = st.sidebar.radio(
        "メニュー",
        [
            "管理者ダッシュボード",
            "健康チェック入力",
            "排泄チェック入力",
            "過去データ管理",
            "排泄詳細管理",
            "家族向けレポート作成",
            "ひだまりレポートPDF",
            "管理者支援",
            "利用者マスタ管理",
        ],
    )
else:
    menu = st.sidebar.radio(
        "メニュー",
        [
            "健康チェック入力",
            "排泄チェック入力",
            "過去データ管理",
        ],
    )


# =========================
# 管理者ダッシュボード
# =========================
if menu == "管理者ダッシュボード":
    if st.session_state.role != "admin":
        st.error("この画面は管理者専用です。")
        st.stop()

    st.header("管理者ダッシュボード")

    health_df = load_health_data()
    ex_df = load_excretion_data()
    today = date.today()

    today_excretion = get_day_excretion_data(ex_df, today, None)

    h_today = health_df.copy()
    if not h_today.empty:
        h_today["記録日"] = pd.to_datetime(h_today["記録日"], errors="coerce")
        h_today = h_today[h_today["記録日"].dt.date == today]

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("本日の健康記録", len(h_today))
    col2.metric("本日の排泄記録", len(today_excretion))
    col3.metric("利用者数", len(active_users))

    ex_sum = summarize_excretion(today_excretion)
    col4.metric("本日の排便記録", ex_sum["排便回数"])

    st.subheader("本日の申し送り支援")
    st.text_area(
        "申し送りメモ",
        value=create_handover_text(health_df, ex_df, today),
        height=320,
    )

    st.subheader("本日の排泄状況")
    if today_excretion.empty:
        st.info("本日の排泄記録はまだありません。")
    else:
        st.dataframe(today_excretion, use_container_width=True, hide_index=True)

        if ex_sum["濃縮尿"] or ex_sum["下痢便"] or ex_sum["水様便"]:
            st.warning(
                f"確認項目：濃縮尿 {ex_sum['濃縮尿']}件、"
                f"下痢便 {ex_sum['下痢便']}件、水様便 {ex_sum['水様便']}件"
            )
        else:
            st.success("本日の排泄状況で大きな注意記録はありません。")


# =========================
# 健康チェック入力
# =========================
elif menu == "健康チェック入力":
    st.header("健康チェック入力")

    if st.session_state.role == "staff":
        st.markdown("### お疲れ様です。")
        st.write("利用者様の今日の健康状態を入力してください。")

    if not active_users:
        st.warning("利用者マスタに表示中の利用者がいません。")
        st.stop()

    col1, col2, col3 = st.columns(3)
    with col1:
        record_date = st.date_input("記録日", value=date.today(), key="health_date")
    with col2:
        user_name = st.selectbox("利用者名", active_users, key="health_user")
    with col3:
        input_staff = st.text_input("入力者", placeholder="例：藤野", key="health_staff")

    health_df = load_health_data()
    idx = find_health_index(health_df, record_date, user_name)

    if idx is None:
        existing_row = None
        st.markdown(
            """
            <div style='background:#EAF4FF; border:1px solid #9CC7F0; color:#174A7C; padding:12px 14px; border-radius:10px; margin:8px 0 12px 0;'>
                <b>この記録日・利用者名の健康チェックデータはありません。</b><br>
                登録すると新規データとして保存されます。
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        existing_row = health_df.loc[idx]
        st.markdown(
            """
            <div style='background:#FFF3E0; border:1px solid #F0B36A; color:#8A4B00; padding:12px 14px; border-radius:10px; margin:8px 0 12px 0;'>
                <b>この記録日・利用者名の健康チェックデータは既にあります。</b><br>
                登録すると上書き更新されます。既存データを初期表示しています。
            </div>
            """,
            unsafe_allow_html=True,
        )

    def row_float(col, default):
        if existing_row is None:
            return default
        return safe_float(existing_row.get(col), default)

    def row_int(col, default):
        if existing_row is None:
            return default
        return safe_int(existing_row.get(col), default)

    def row_text(col, default=""):
        if existing_row is None:
            return default
        return clean_text(existing_row.get(col), default)

    with st.form("health_form", clear_on_submit=False):
        st.subheader("バイタル")

        c1, c2, c3 = st.columns(3)
        with c1:
            temp = st.number_input("体温", min_value=0.0, max_value=45.0, value=row_float("体温", 0.0), step=0.1)
        with c2:
            bp_high = st.number_input("血圧上", min_value=0, max_value=250, value=row_int("血圧上", 0), step=1)
        with c3:
            bp_low = st.number_input("血圧下", min_value=0, max_value=150, value=row_int("血圧下", 0), step=1)

        c4, c5, c6 = st.columns(3)
        with c4:
            pulse = st.number_input("脈拍", min_value=0, max_value=200, value=row_int("脈拍", 0), step=1)
        with c5:
            spo2 = st.number_input("SpO2", min_value=0, max_value=100, value=row_int("SpO2", 0), step=1)
        with c6:
            weight = st.number_input("体重", min_value=0.0, max_value=200.0, value=row_float("体重", 0.0), step=0.1)

        st.divider()
        st.subheader("食事摂取率")

        m1, m2, m3 = st.columns(3)
        with m1:
            breakfast = st.slider("朝食", 0, 100, row_int("朝食摂取率", 80), step=10)
        with m2:
            lunch = st.slider("昼食", 0, 100, row_int("昼食摂取率", 80), step=10)
        with m3:
            dinner = st.slider("夕食", 0, 100, row_int("夕食摂取率", 80), step=10)

        family_memo = st.text_area("家族共有メモ", value=row_text("家族共有メモ"), placeholder="ご家族へ共有してよい内容を入力")
        changes = st.text_area("気になる変化", value=row_text("気になる変化"), placeholder="食事、睡眠、歩行、表情、体調など")

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
            "朝食摂取率": breakfast,
            "昼食摂取率": lunch,
            "夕食摂取率": dinner,
            "家族共有メモ": family_memo,
            "気になる変化": changes,
            "登録日時": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "入力者": input_staff,
        }
        action = upsert_health_record(record)
        st.success(f"健康チェックを{action}しました。")
        st.rerun()


# =========================
# 排泄チェック入力
# =========================
elif menu == "排泄チェック入力":
    st.header("排泄チェック入力")
    st.caption("排泄記録は健康チェックとは別データとして保存します。キーは「記録日＋利用者名＋時間帯」です。")

    if st.session_state.role == "staff":
        st.markdown("### お疲れ様です。")
        st.write("排泄状況を時間帯ごとに入力してください。")

    if not active_users:
        st.warning("利用者マスタに表示中の利用者がいません。")
        st.stop()

    col1, col2, col3 = st.columns(3)
    with col1:
        record_date = st.date_input("記録日", value=date.today(), key="ex_input_date")
    with col2:
        user_name = st.selectbox("利用者名", active_users, key="ex_input_user")
    with col3:
        input_staff = st.text_input("入力者", placeholder="例：藤野", key="ex_input_staff")

    ex_df = load_excretion_data()
    day_data = get_day_excretion_data(ex_df, record_date, user_name)

    if day_data.empty:
        st.markdown(
            """
            <div style='background:#EAF4FF; border:1px solid #9CC7F0; color:#174A7C; padding:12px 14px; border-radius:10px; margin:8px 0 12px 0;'>
                <b>この記録日・利用者名の排泄データはありません。</b><br>
                登録すると新規データとして保存されます。
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            """
            <div style='background:#FFF3E0; border:1px solid #F0B36A; color:#8A4B00; padding:12px 14px; border-radius:10px; margin:8px 0 12px 0;'>
                <b>この記録日・利用者名の排泄データは既にあります。</b><br>
                登録すると時間帯ごとに上書き更新されます。
            </div>
            """,
            unsafe_allow_html=True,
        )

    with st.form("excretion_form", clear_on_submit=False):
        records_to_save = []

        st.markdown("#### ☀️ 日中帯（9時〜17時）")
        day_cols = st.columns(3)

        for col, (slot, time_label) in zip(day_cols, EXCRETION_SLOTS[:3]):
            existing = get_excretion_row(ex_df, record_date, user_name, slot)
            sig = "new" if existing is None else hashlib.md5(str(existing.to_dict()).encode("utf-8")).hexdigest()[:8]
            key_base = f"ex_{record_date}_{user_name}_{slot}_{sig}"

            with col:
                st.markdown(
                    f"""
                    <div style='background:#FFF7EC; padding:12px; border-radius:14px; border:1px solid #E5D5BF; margin-bottom:10px;'>
                        <b style='font-size:16px;'>{slot}</b><br>
                        <span style='font-size:12px; color:#666;'>{time_label}</span>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                urine_amount = st.selectbox(
                    f"{slot} 尿量",
                    URINE_AMOUNT_OPTIONS,
                    index=get_option_index(URINE_AMOUNT_OPTIONS, existing.get("尿量", "なし") if existing is not None else "なし"),
                    key=f"{key_base}_urine_amount",
                )
                urine_type = st.selectbox(
                    f"{slot} 尿性状",
                    URINE_TYPE_OPTIONS,
                    index=get_option_index(URINE_TYPE_OPTIONS, existing.get("尿性状", "なし") if existing is not None else "なし"),
                    key=f"{key_base}_urine_type",
                )
                stool_amount = st.selectbox(
                    f"{slot} 便量",
                    STOOL_AMOUNT_OPTIONS,
                    index=get_option_index(STOOL_AMOUNT_OPTIONS, existing.get("便量", "なし") if existing is not None else "なし"),
                    key=f"{key_base}_stool_amount",
                )
                stool_type = st.selectbox(
                    f"{slot} 便性状",
                    STOOL_TYPE_OPTIONS,
                    index=get_option_index(STOOL_TYPE_OPTIONS, existing.get("便性状", "なし") if existing is not None else "なし"),
                    key=f"{key_base}_stool_type",
                )
                memo = st.text_area(
                    f"{slot} メモ",
                    value=existing.get("排泄メモ", "") if existing is not None else "",
                    key=f"{key_base}_memo",
                    height=80,
                )

                if urine_amount == "なし":
                    urine_type = "なし"
                if stool_amount == "なし":
                    stool_type = "なし"

                records_to_save.append({
                    "記録日": record_date,
                    "利用者名": user_name,
                    "時間帯": slot,
                    "時間帯目安": time_label,
                    "尿量": urine_amount,
                    "尿性状": urine_type,
                    "便量": stool_amount,
                    "便性状": stool_type,
                    "排泄メモ": memo,
                    "入力者": input_staff,
                    "登録日時": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                })

        st.markdown("#### 🌙 夜間帯（18時〜翌8時）")
        night_cols = st.columns(3)

        for col, (slot, time_label) in zip(night_cols, EXCRETION_SLOTS[3:]):
            existing = get_excretion_row(ex_df, record_date, user_name, slot)
            sig = "new" if existing is None else hashlib.md5(str(existing.to_dict()).encode("utf-8")).hexdigest()[:8]
            key_base = f"ex_{record_date}_{user_name}_{slot}_{sig}"

            with col:
                st.markdown(
                    f"""
                    <div style='background:#EEF4FA; padding:12px; border-radius:14px; border:1px solid #C9D8E6; margin-bottom:10px;'>
                        <b style='font-size:16px;'>{slot}</b><br>
                        <span style='font-size:12px; color:#666;'>{time_label}</span>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                urine_amount = st.selectbox(
                    f"{slot} 尿量",
                    URINE_AMOUNT_OPTIONS,
                    index=get_option_index(URINE_AMOUNT_OPTIONS, existing.get("尿量", "なし") if existing is not None else "なし"),
                    key=f"{key_base}_urine_amount",
                )
                urine_type = st.selectbox(
                    f"{slot} 尿性状",
                    URINE_TYPE_OPTIONS,
                    index=get_option_index(URINE_TYPE_OPTIONS, existing.get("尿性状", "なし") if existing is not None else "なし"),
                    key=f"{key_base}_urine_type",
                )
                stool_amount = st.selectbox(
                    f"{slot} 便量",
                    STOOL_AMOUNT_OPTIONS,
                    index=get_option_index(STOOL_AMOUNT_OPTIONS, existing.get("便量", "なし") if existing is not None else "なし"),
                    key=f"{key_base}_stool_amount",
                )
                stool_type = st.selectbox(
                    f"{slot} 便性状",
                    STOOL_TYPE_OPTIONS,
                    index=get_option_index(STOOL_TYPE_OPTIONS, existing.get("便性状", "なし") if existing is not None else "なし"),
                    key=f"{key_base}_stool_type",
                )
                memo = st.text_area(
                    f"{slot} メモ",
                    value=existing.get("排泄メモ", "") if existing is not None else "",
                    key=f"{key_base}_memo",
                    height=80,
                )

                if urine_amount == "なし":
                    urine_type = "なし"
                if stool_amount == "なし":
                    stool_type = "なし"

                records_to_save.append({
                    "記録日": record_date,
                    "利用者名": user_name,
                    "時間帯": slot,
                    "時間帯目安": time_label,
                    "尿量": urine_amount,
                    "尿性状": urine_type,
                    "便量": stool_amount,
                    "便性状": stool_type,
                    "排泄メモ": memo,
                    "入力者": input_staff,
                    "登録日時": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                })

        submitted = st.form_submit_button("排泄チェックを登録・更新する")

    if submitted:
        actions = []
        for record in records_to_save:
            actions.append(upsert_excretion_record(record))

        st.success("排泄チェックを保存しました。時間帯ごとに登録・更新されています。")
        st.rerun()

    st.subheader("この日の排泄記録")
    day_data = get_day_excretion_data(load_excretion_data(), record_date, user_name)
    if day_data.empty:
        st.info("この日の排泄記録はまだありません。")
    else:
        st.dataframe(day_data, use_container_width=True, hide_index=True)


# =========================
# 過去データ管理
# =========================
elif menu == "過去データ管理":
    st.header("過去データ管理")
    st.caption("健康チェックデータを、記録日＋利用者名で検索・更新・削除します。")

    health_df = load_health_data()

    if health_df.empty:
        st.info("まだ健康チェックデータがありません。")
        st.stop()

    col1, col2 = st.columns(2)
    with col1:
        key_date = st.date_input("記録日", value=date.today(), key="past_health_date")
    with col2:
        key_user = st.selectbox("利用者名", all_users, key="past_health_user")

    idx = find_health_index(health_df, key_date, key_user)

    if idx is None:
        st.info("この記録日・利用者名の健康チェックデータはありません。")
    else:
        st.success("該当データが見つかりました。")
        row = health_df.loc[idx]

        with st.form("health_update_form"):
            c1, c2, c3 = st.columns(3)
            with c1:
                temp = st.number_input("体温", value=safe_float(row.get("体温"), 0.0), step=0.1)
            with c2:
                bp_high = st.number_input("血圧上", value=safe_int(row.get("血圧上"), 0), step=1)
            with c3:
                bp_low = st.number_input("血圧下", value=safe_int(row.get("血圧下"), 0), step=1)

            c4, c5, c6 = st.columns(3)
            with c4:
                pulse = st.number_input("脈拍", value=safe_int(row.get("脈拍"), 0), step=1)
            with c5:
                spo2 = st.number_input("SpO2", value=safe_int(row.get("SpO2"), 0), step=1)
            with c6:
                weight = st.number_input("体重", value=safe_float(row.get("体重"), 0.0), step=0.1)

            m1, m2, m3 = st.columns(3)
            with m1:
                breakfast = st.slider("朝食", 0, 100, safe_int(row.get("朝食摂取率"), 80), step=10)
            with m2:
                lunch = st.slider("昼食", 0, 100, safe_int(row.get("昼食摂取率"), 80), step=10)
            with m3:
                dinner = st.slider("夕食", 0, 100, safe_int(row.get("夕食摂取率"), 80), step=10)

            family_memo = st.text_area("家族共有メモ", value=clean_text(row.get("家族共有メモ", "")))
            changes = st.text_area("気になる変化", value=clean_text(row.get("気になる変化", "")))
            staff = st.text_input("入力者", value=clean_text(row.get("入力者", "")))

            update_submit = st.form_submit_button("更新する")

        if update_submit:
            record = {
                "記録日": key_date,
                "利用者名": key_user,
                "体温": temp,
                "血圧上": bp_high,
                "血圧下": bp_low,
                "脈拍": pulse,
                "SpO2": spo2,
                "体重": weight,
                "朝食摂取率": breakfast,
                "昼食摂取率": lunch,
                "夕食摂取率": dinner,
                "家族共有メモ": family_memo,
                "気になる変化": changes,
                "登録日時": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "入力者": staff,
            }
            action = upsert_health_record(record)
            st.success(f"{action}しました。")
            st.rerun()

        st.warning("削除すると元に戻せません。")
        delete_check = st.checkbox("この健康チェックデータを削除する")

        if st.button("削除する"):
            if not delete_check:
                st.error("削除する場合は確認チェックを入れてください。")
            else:
                health_df = health_df.drop(index=idx).reset_index(drop=True)
                save_health_data(health_df)
                st.success("削除しました。")
                st.rerun()

    st.divider()
    st.subheader("一覧検索")

    health_df = load_health_data()
    if not health_df.empty:
        health_df["記録日"] = pd.to_datetime(health_df["記録日"], errors="coerce")
        year = st.number_input("年", min_value=2024, max_value=2035, value=date.today().year, step=1)
        month = st.number_input("月", min_value=1, max_value=12, value=date.today().month, step=1)
        user_filter = st.selectbox("利用者で絞り込み", ["全員"] + all_users)

        result = health_df[
            (health_df["記録日"].dt.year == int(year))
            & (health_df["記録日"].dt.month == int(month))
        ]
        if user_filter != "全員":
            result = result[result["利用者名"] == user_filter]

        st.dataframe(result.sort_values(["記録日", "利用者名"]), use_container_width=True, hide_index=True)


# =========================
# 排泄詳細管理
# =========================
elif menu == "排泄詳細管理":
    if st.session_state.role != "admin":
        st.error("この画面は管理者専用です。")
        st.stop()

    st.header("排泄詳細管理")
    st.caption("排泄チェックデータを、記録日＋利用者名＋時間帯で管理します。")

    ex_df = load_excretion_data()

    if ex_df.empty:
        st.info("まだ排泄チェックデータがありません。")
        st.stop()

    col1, col2, col3 = st.columns(3)
    with col1:
        ex_user = st.selectbox("利用者", ["全員"] + all_users, key="ex_admin_user")
    with col2:
        start_date = st.date_input("開始日", value=date.today(), key="ex_admin_start")
    with col3:
        end_date = st.date_input("終了日", value=date.today(), key="ex_admin_end")

    work = ex_df.copy()
    work["記録日"] = pd.to_datetime(work["記録日"], errors="coerce")
    work = work[
        (work["記録日"].dt.date >= start_date)
        & (work["記録日"].dt.date <= end_date)
    ]

    if ex_user != "全員":
        work = work[work["利用者名"] == ex_user]

    st.subheader("排泄サマリー")
    if work.empty:
        st.warning("該当する排泄データがありません。")
    else:
        summary_rows = []
        for user in work["利用者名"].dropna().unique():
            user_df = work[work["利用者名"] == user]
            s = summarize_excretion(user_df)
            summary_rows.append({
                "利用者名": user,
                "記録数": len(user_df),
                "排尿回数": s["排尿回数"],
                "排便回数": s["排便回数"],
                "濃縮尿": s["濃縮尿"],
                "下痢便": s["下痢便"],
                "水様便": s["水様便"],
                "排便なし枠": s["排便なし枠"],
            })
        st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)

        st.subheader("注意して確認したい排泄記録")
        alert = work[
            (work["尿性状"] == "濃縮尿")
            | (work["便性状"].isin(["下痢便", "水様便"]))
        ]

        if alert.empty:
            st.success("指定期間内に、濃縮尿・下痢便・水様便の記録はありません。")
        else:
            st.warning("確認したい排泄記録があります。")
            st.dataframe(alert, use_container_width=True, hide_index=True)

        st.subheader("時系列の排泄詳細")
        slot_order = {slot: i for i, (slot, _) in enumerate(EXCRETION_SLOTS)}
        work["_slot_order"] = work["時間帯"].map(slot_order).fillna(99)
        work = work.sort_values(["記録日", "利用者名", "_slot_order"]).drop(columns=["_slot_order"])
        st.dataframe(work, use_container_width=True, hide_index=True)

        csv = work.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "排泄詳細CSVをダウンロード",
            data=csv,
            file_name="排泄詳細データ.csv",
            mime="text/csv",
        )

        st.subheader("管理者向け確認メモ")
        memo_lines = [
            "排泄詳細データをもとにした管理者確認メモです。",
            "医療判断ではなく、職員間の共有と見守り方針の整理に使用してください。",
            "",
        ]

        for _, row in pd.DataFrame(summary_rows).iterrows():
            memo_lines.append(
                f"■ {row['利用者名']}：排尿{row['排尿回数']}回、排便{row['排便回数']}回、"
                f"濃縮尿{row['濃縮尿']}回、下痢便{row['下痢便']}回、水様便{row['水様便']}回。"
            )

        st.text_area("確認メモ", value="\n".join(memo_lines), height=260)

    st.divider()
    st.subheader("排泄データの更新・削除")

    c1, c2, c3 = st.columns(3)
    with c1:
        key_date = st.date_input("更新対象日", value=date.today(), key="ex_edit_date")
    with c2:
        key_user = st.selectbox("更新対象利用者", all_users, key="ex_edit_user")
    with c3:
        key_slot = st.selectbox("時間帯", [slot for slot, _ in EXCRETION_SLOTS], key="ex_edit_slot")

    current = get_excretion_row(load_excretion_data(), key_date, key_user, key_slot)

    if current is None:
        st.info("このキーの排泄データはありません。")
    else:
        st.success("該当する排泄データがあります。")

    with st.form("ex_edit_form"):
        time_label = dict(EXCRETION_SLOTS).get(key_slot, "")

        urine_amount = st.selectbox(
            "尿量",
            URINE_AMOUNT_OPTIONS,
            index=get_option_index(URINE_AMOUNT_OPTIONS, current.get("尿量", "なし") if current is not None else "なし"),
        )
        urine_type = st.selectbox(
            "尿性状",
            URINE_TYPE_OPTIONS,
            index=get_option_index(URINE_TYPE_OPTIONS, current.get("尿性状", "なし") if current is not None else "なし"),
        )
        stool_amount = st.selectbox(
            "便量",
            STOOL_AMOUNT_OPTIONS,
            index=get_option_index(STOOL_AMOUNT_OPTIONS, current.get("便量", "なし") if current is not None else "なし"),
        )
        stool_type = st.selectbox(
            "便性状",
            STOOL_TYPE_OPTIONS,
            index=get_option_index(STOOL_TYPE_OPTIONS, current.get("便性状", "なし") if current is not None else "なし"),
        )
        memo = st.text_area("排泄メモ", value=current.get("排泄メモ", "") if current is not None else "")
        staff = st.text_input("入力者", value=current.get("入力者", "") if current is not None else "")

        submit = st.form_submit_button("登録・更新する")

    if submit:
        if urine_amount == "なし":
            urine_type = "なし"
        if stool_amount == "なし":
            stool_type = "なし"

        record = {
            "記録日": key_date,
            "利用者名": key_user,
            "時間帯": key_slot,
            "時間帯目安": time_label,
            "尿量": urine_amount,
            "尿性状": urine_type,
            "便量": stool_amount,
            "便性状": stool_type,
            "排泄メモ": memo,
            "入力者": staff,
            "登録日時": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        action = upsert_excretion_record(record)
        st.success(f"排泄データを{action}しました。")
        st.rerun()

    if current is not None:
        delete_check = st.checkbox("この排泄データを削除する")
        if st.button("排泄データを削除する"):
            if not delete_check:
                st.error("削除する場合は確認チェックを入れてください。")
            else:
                ex_df = load_excretion_data()
                idx = find_excretion_index(ex_df, key_date, key_user, key_slot)
                if idx is not None:
                    ex_df = ex_df.drop(index=idx).reset_index(drop=True)
                    save_excretion_data(ex_df)
                    st.success("削除しました。")
                    st.rerun()


# =========================
# 家族向けレポート作成
# =========================
elif menu == "家族向けレポート作成":
    if st.session_state.role != "admin":
        st.error("この画面は管理者専用です。")
        st.stop()

    st.header("家族向けレポート作成")

    if not all_users:
        st.warning("利用者が登録されていません。")
        st.stop()

    col1, col2, col3 = st.columns(3)
    with col1:
        report_user = st.selectbox("利用者", all_users, key="family_report_user")
    with col2:
        report_year = st.number_input("対象年", min_value=2024, max_value=2035, value=date.today().year, step=1)
    with col3:
        report_month = st.number_input("対象月", min_value=1, max_value=12, value=date.today().month, step=1)

    health_df = load_health_data()
    ex_df = load_excretion_data()

    report_text = create_family_summary_text(health_df, ex_df, report_user, report_year, report_month)
    st.text_area("家族向け文章", value=report_text, height=420)

    ex_target = get_month_excretion_data(ex_df, report_user, report_year, report_month)
    with st.expander("排泄記録を確認する"):
        if ex_target.empty:
            st.info("対象月の排泄記録はありません。")
        else:
            st.dataframe(ex_target, use_container_width=True, hide_index=True)


# =========================
# ひだまりレポートPDF
# =========================
elif menu == "ひだまりレポートPDF":
    if st.session_state.role != "admin":
        st.error("この画面は管理者専用です。")
        st.stop()

    st.header("ひだまりレポートPDF")

    if not all_users:
        st.warning("利用者が登録されていません。")
        st.stop()

    col1, col2, col3 = st.columns(3)
    with col1:
        pdf_user = st.selectbox("利用者", all_users, key="pdf_user")
    with col2:
        pdf_year = st.number_input("対象年", min_value=2024, max_value=2035, value=date.today().year, step=1, key="pdf_year")
    with col3:
        pdf_month = st.number_input("対象月", min_value=1, max_value=12, value=date.today().month, step=1, key="pdf_month")

    if st.button("ひだまりレポートPDFを作成する"):
        try:
            path = create_hidamari_pdf(load_health_data(), load_excretion_data(), pdf_user, pdf_year, pdf_month)
            with open(path, "rb") as f:
                st.download_button(
                    "PDFをダウンロード",
                    data=f,
                    file_name=path.name,
                    mime="application/pdf",
                )
            st.success("PDFを作成しました。")
        except Exception as e:
            st.error(f"PDF作成中にエラーが発生しました：{e}")


# =========================
# 管理者支援
# =========================
elif menu == "管理者支援":
    if st.session_state.role != "admin":
        st.error("この画面は管理者専用です。")
        st.stop()

    st.header("管理者支援")
    health_df = load_health_data()
    ex_df = load_excretion_data()

    tab1, tab2, tab3, tab4 = st.tabs(["AI家族レポート", "バイタル推移グラフ", "ChatGPT連携", "申し送り支援"])

    with tab1:
        st.subheader("AI家族レポート自動文章")
        col1, col2, col3 = st.columns(3)
        with col1:
            ai_user = st.selectbox("利用者", all_users, key="ai_user")
        with col2:
            ai_year = st.number_input("対象年", min_value=2024, max_value=2035, value=date.today().year, step=1, key="ai_year")
        with col3:
            ai_month = st.number_input("対象月", min_value=1, max_value=12, value=date.today().month, step=1, key="ai_month")

        summary = create_family_summary_text(health_df, ex_df, ai_user, ai_year, ai_month)
        st.text_area("家族向け文章", value=summary, height=360)

    with tab2:
        st.subheader("バイタル推移グラフ")
        if health_df.empty:
            st.info("データがありません。")
        else:
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                graph_user = st.selectbox("利用者", all_users, key="graph_user")
            with col2:
                graph_item = st.selectbox("項目", ["体温", "血圧上", "血圧下", "脈拍", "SpO2", "体重", "朝食摂取率", "昼食摂取率", "夕食摂取率"], key="graph_item")
            with col3:
                graph_year = st.number_input("年", min_value=2024, max_value=2035, value=date.today().year, step=1, key="graph_year")
            with col4:
                graph_month = st.number_input("月", min_value=1, max_value=12, value=date.today().month, step=1, key="graph_month")

            target = get_month_health_data(health_df, graph_user, graph_year, graph_month)
            if target.empty:
                st.warning("対象データがありません。")
            else:
                chart_df = target[["記録日", graph_item]].copy()
                chart_df[graph_item] = pd.to_numeric(chart_df[graph_item], errors="coerce")
                chart_df = chart_df.dropna()
                chart_df = chart_df.set_index("記録日")
                st.line_chart(chart_df)

    with tab3:
        st.subheader("ChatGPT連携用プロンプト")
        col1, col2, col3 = st.columns(3)
        with col1:
            prompt_user = st.selectbox("利用者", all_users, key="prompt_user")
        with col2:
            prompt_year = st.number_input("年", min_value=2024, max_value=2035, value=date.today().year, step=1, key="prompt_year")
        with col3:
            prompt_month = st.number_input("月", min_value=1, max_value=12, value=date.today().month, step=1, key="prompt_month")

        target_h = get_month_health_data(health_df, prompt_user, prompt_year, prompt_month)
        target_e = get_month_excretion_data(ex_df, prompt_user, prompt_year, prompt_month)

        prompt = f"""あなたは介護施設の家族向けレポートを整える文章整理係です。
以下の健康チェック記録・排泄記録・アセスメント情報をもとに、ご家族へ渡す月間レポート文を作成してください。

【重要ルール】
・医療判断、診断、治療効果の断定はしない。
・「問題ありません」「改善しました」「安心です」と断定しない。
・記録に基づく表現にする。
・不安を煽らず、やわらかく丁寧な文章にする。

【利用者】
{prompt_user}

【アセスメント情報】
{build_assessment_context_text(prompt_user)}

【健康チェック記録】
{target_h.to_string(index=False)}

【排泄記録】
{target_e.to_string(index=False)}
"""
        st.text_area("プロンプト", value=prompt, height=520)

    with tab4:
        st.subheader("申し送り支援")
        target_date = st.date_input("対象日", value=date.today(), key="handover_date")
        handover = create_handover_text(health_df, ex_df, target_date)
        st.text_area("申し送り案", value=handover, height=360)


# =========================
# 利用者マスタ管理
# =========================
elif menu == "利用者マスタ管理":
    if st.session_state.role != "admin":
        st.error("この画面は管理者専用です。")
        st.stop()

    st.header("利用者マスタ管理")
    st.caption("利用者の追加・非表示・アセスメント情報の登録管理ができます。")

    df_users = load_users(include_hidden=True)

    st.subheader("現在の利用者一覧")
    st.dataframe(df_users, use_container_width=True, hide_index=True)

    st.divider()
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

    st.divider()
    st.subheader("アセスメント情報の登録・更新")

    if df_users.empty:
        st.info("利用者が登録されていません。")
    else:
        selected_user = st.selectbox("アセスメントを編集する利用者", df_users["利用者名"].tolist(), key="assessment_user")
        selected = df_users[df_users["利用者名"] == selected_user].iloc[0]

        with st.form("assessment_form"):
            values = {}
            values["基本情報"] = st.text_area("基本情報（氏名・住所など）", value=clean_text(selected.get("基本情報", "")), height=80)
            values["主訴"] = st.text_area("主訴（本人・家族の希望や困りごと）", value=clean_text(selected.get("主訴", "")), height=100)
            values["生活状況"] = st.text_area("生活状況（1日の流れ）", value=clean_text(selected.get("生活状況", "")), height=120)
            values["ADL"] = st.text_area("ADL（日常生活動作）", value=clean_text(selected.get("ADL", "")), height=100)
            values["IADL"] = st.text_area("IADL（生活関連動作）", value=clean_text(selected.get("IADL", "")), height=100)
            values["認知機能"] = st.text_area("認知機能（判断・記憶）", value=clean_text(selected.get("認知機能", "")), height=100)
            values["健康状態"] = st.text_area("健康状態（疾患・服薬）", value=clean_text(selected.get("健康状態", "")), height=100)
            values["課題"] = st.text_area("課題（支援が必要な問題点）", value=clean_text(selected.get("課題", "")), height=100)
            values["支援内容"] = st.text_area("支援内容（具体的な対応）", value=clean_text(selected.get("支援内容", "")), height=100)
            assessment_submit = st.form_submit_button("アセスメント情報を保存する")

        if assessment_submit:
            df_save = load_users(include_hidden=True)
            mask = df_save["利用者名"] == selected_user
            for col, value in values.items():
                df_save.loc[mask, col] = value
            save_users(df_save)
            st.success("アセスメント情報を保存しました。")
            st.rerun()

    st.divider()
    st.subheader("利用者を入力候補から外す")

    visible_users = load_users(include_hidden=False)["利用者名"].tolist()

    if visible_users:
        target_user = st.selectbox("対象利用者", visible_users, key="hide_user")
        st.warning("この操作は、入力画面の候補から外すだけです。過去データとアセスメント情報は削除されません。")

        if st.button("入力候補から外す"):
            ok, msg = hide_user(target_user)
            if ok:
                st.success(msg)
                st.rerun()
            else:
                st.error(msg)

    hidden_df = load_users(include_hidden=True)
    hidden_df = hidden_df[hidden_df["表示"] == "非表示"]

    st.subheader("非表示の利用者を戻す")

    if not hidden_df.empty:
        restore_user = st.selectbox("表示に戻す利用者", hidden_df["利用者名"].tolist(), key="restore_user")
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
