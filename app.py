# =========================
{user}様の記録では、
体温は平均{round(avg_temp,1)}℃、
SpO2は平均{round(avg_spo2,1)}％でした。

記録上、大きな変化は目立っていません。
一方で、日々の表情や生活リズムも含め、
引き続き様子を見守っています。

数値だけではなく、
安心して過ごせることを大切にしながら、
日々の関わりを続けています。
            """

            st.text_area("AIレポート", report, height=300)


# =========================
# 申し送り支援
# =========================
elif menu == "申し送り支援":
    st.header("AI申し送り支援")

    df = load_data()

    if not df.empty:
        latest = df.tail(10)

        summary = []

        for _, row in latest.iterrows():
            summary.append(
                f"{row['利用者名']}：{row['気になる変化']}"
            )

        result = "\n".join(summary)

        st.text_area(
            "申し送りまとめ",
            result,
            height=300,
        )


# =========================
# バイタル推移グラフ
# =========================
if st.session_state.role == "admin":

    st.sidebar.divider()
    st.sidebar.markdown("### バイタル推移")

    graph_user = st.sidebar.selectbox(
        "グラフ利用者",
        DEFAULT_USERS,
    )

    graph_item = st.sidebar.selectbox(
        "項目",
        ["体温", "血圧上", "SpO2", "体重"],
    )

    df = load_data()

    if not df.empty:
        target = df[df["利用者名"] == graph_user]

        if not target.empty:
            target["記録日"] = pd.to_datetime(target["記録日"])

            fig, ax = plt.subplots(figsize=(8,4))

            ax.plot(
                target["記録日"],
                pd.to_numeric(target[graph_item], errors="coerce"),
                marker="o",
            )

            ax.set_title(f"{graph_user} {graph_item}推移")
            ax.grid(True)

            st.sidebar.pyplot(fig)
