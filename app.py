import streamlit as st
import pandas as pd
import os
import io

EXCEL_FILE = "商品マスタ統合_raw.xlsx"

# =============================
# データ読み込み
# =============================
def load_data():
    df = pd.read_excel(EXCEL_FILE)

    # 選択列がなければ追加
    if "選択" not in df.columns:
        df["選択"] = ""

    # メモ列がなければ追加
    if "メモ" not in df.columns:
        df["メモ"] = ""

    # 受注締日 → 日付文字列（数字文字列も Excel シリアルとして扱う）
    def convert_date(x):
        if pd.isna(x):
            return ""
        s = str(x).strip()
        if s == "":
            return ""
        if s.isdigit():
            try:
                v = int(s)
                d = pd.to_datetime(v, unit="d", origin="1899-12-30")
                return d.strftime("%Y-%m-%d")
            except:
                return ""
        try:
            d = pd.to_datetime(s)
            return d.strftime("%Y-%m-%d")
        except:
            return ""

    df["締め日表示"] = df["受注締日"].apply(convert_date)

    return df


# =============================
# session_state に保持
# =============================
if "df" not in st.session_state:
    st.session_state["df"] = load_data()

df = st.session_state["df"]

# =============================
# UI
# =============================
st.title("🛒 展示会カタログ（Excel画像版）")

st.subheader("検索")

col1, col2, col3 = st.columns([2,1,1])

with col1:
    keyword = st.text_input("キーワード検索（品番 / 商品名 / カテゴリ）", "")

with col2:
    締め日候補 = ["すべて"] + sorted([d for d in df["締め日表示"].unique() if d])
    selected_締め日 = st.selectbox("受注締日", 締め日候補)

with col3:
    取扱候補 = ["すべて"] + sorted(df["ZETT取扱"].fillna("").unique().tolist())
    selected_取扱 = st.selectbox("ZETT取扱", 取扱候補)

# =============================
# フィルタリング
# =============================
filtered = df.copy()

if keyword:
    kw = keyword.lower()
    def match(row):
        text = (
            str(row["Article_No"]) +
            str(row["Model_No"]) +
            str(row["Model_Name_Local_Zenkaku"]) +
            str(row["Category_Code"]) +
            str(row["Range1"])
        ).lower()
        return kw in text
    filtered = filtered[filtered.apply(match, axis=1)]

if selected_締め日 != "すべて":
    filtered = filtered[filtered["締め日表示"] == selected_締め日]

if selected_取扱 != "すべて":
    filtered = filtered[filtered["ZETT取扱"].fillna("") == selected_取扱]

st.write(f"検索結果：{len(filtered)} 件")
st.divider()

# =============================
# 商品カード（画像左・情報右）
# =============================
for idx, row in filtered.iterrows():
    with st.container(border=True):
        # 画像を左、情報を右
        c1, c2 = st.columns([2, 3])

        # ---------- 画像（左） ----------
        with c1:
            raw_path = str(row["画像パス"])
            filename = os.path.basename(raw_path)
            img_path = os.path.join("images", filename)

            if os.path.exists(img_path):
                st.image(img_path, width=260)   # ← 大きめ画像
            else:
                st.write("画像なし")

        # ---------- 商品情報（右） ----------
        with c2:
            st.markdown(f"### {row['Article_No']}")
            st.write(f"**Model_No**：{row['Model_No']}")
            st.write(f"**商品名**：{row['Model_Name_Local_Zenkaku']}")
            st.write(f"**上代**：{row['Retail_Price']}")
            st.write(f"**受注締日**：{row['締め日表示']}")
            st.write(f"**ZETT取扱**：{row['ZETT取扱']}")
            st.write(f"**カテゴリ**：{row['Category_Code']}")
            st.write(f"**Range1**：{row['Range1']}")

            # トグル式選択
            b1, b2, b3 = st.columns(3)
            current = st.session_state["df"].at[idx, "選択"]

            if b1.button("BUY", key=f"buy_{idx}"):
                st.session_state["df"].at[idx, "選択"] = "" if current == "BUY" else "BUY"
            if b2.button("HOLD", key=f"hold_{idx}"):
                st.session_state["df"].at[idx, "選択"] = "" if current == "HOLD" else "HOLD"
            if b3.button("NO", key=f"no_{idx}"):
                st.session_state["df"].at[idx, "選択"] = "" if current == "NO" else "NO"

            st.write(f"**選択中**：{st.session_state['df'].at[idx, '選択']}")

            # メモ欄
            memo = st.text_area("メモ", value=row["メモ"], key=f"memo_{idx}")
            st.session_state["df"].at[idx, "メモ"] = memo

st.divider()

# =============================
# BUY の Excel 出力
# =============================
buy_df = st.session_state["df"][st.session_state["df"]["選択"] == "BUY"]

if st.button("📥 BUY の商品を Excel でダウンロード"):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        buy_df.to_excel(writer, index=False, sheet_name="BUY")
    st.download_button(
        "Excel をダウンロード",
        data=output.getvalue(),
        file_name="BUY商品一覧.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )