# app.py 完全コード（ログイン情報非表示版）

以下を app.py に全文コピーしてください。

```python
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
        <div style='text-align:center; padding:24px;'>
            <h1 style='color:#2E7D32;'>
                健康チェック管理システム
            </h1>

            <p style='color:#666;'>
                利用者様の健康記録を安全に管理します
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    col1, col2, col3 = st.columns([1, 2, 1])

    with col2:

        st.markdown("### ログイン")

        input_id = st.text_input("ID")

        input_password = st.text_input(
            "パスワード",
            type="password"
        )

        if st.button(
            "ログイン",
            use_container_width=True
        ):

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

        st.caption(
            f"ログイン中：{st.session_state.user_label}"
        )

        if st.button("ログアウト"):

            st.session_state.logged_in = False
            st.session_state.role = None
            st.session_state.user_label = ""

            st.rerun()


if not login_check():
    st.stop()


# =========================
# デザイン
# =========================
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


apply_design()

st.title("健康チェックWebアプリ")

if st.session_state.role == "admin":
    st.success("管理者モード")
else:
    st.info("お疲れ様です。今日の健康チェック入力をお願いします。")


# =========================
# サイドメニュー
# =========================
logout_button()

if st.session_state.role == "admin":

    menu = st.sidebar.radio(
        "メニュー",
        [
            "管理者ダッシュボード",
            "健康チェック入力",
            "過去データ管理",
        ],
    )

else:

    menu = st.sidebar.radio(
        "メニュー",
        [
            "健康チェック入力",
            "過去データ管理",
        ],
    )


# =========================
# ダミー表示
# =========================
if menu == "管理者ダッシュボード":

    st.header("管理者ダッシュボード")

    col1, col2, col3 = st.columns(3)

    with col1:
        st.metric("本日の入力", "8件")

    with col2:
        st.metric("未入力", "2名")

    with col3:
        st.metric("注意記録", "1件")

    st.subheader("申し送り支援")

    st.text_area(
        "申し送り",
        value="本日は食欲低下の記録がありました。継続して様子観察をお願いします。",
        height=150,
    )


elif menu == "健康チェック入力":

    st.header("健康チェック入力")

    col1, col2 = st.columns(2)

    with col1:
        st.date_input("記録日", value=date.today())
        st.selectbox("利用者名", ["さくら様", "谷様"])
        st.number_input("体温", value=36.5)
        st.number_input("血圧上", value=120)

    with col2:
        st.number_input("血圧下", value=75)
        st.number_input("脈拍", value=70)
        st.number_input("SpO2", value=96)
        st.number_input("体重", value=50.0)

    st.text_area("家族共有メモ")
    st.text_area("気になる変化")

    st.button("登録する")


elif menu == "過去データ管理":

    st.header("過去データ管理")

    sample = pd.DataFrame(
        {
            "記録日": ["2026-05-13"],
            "利用者名": ["さくら様"],
            "体温": [36.5],
            "SpO2": [96],
        }
    )

    st.dataframe(sample)

```
