import streamlit as st
import pandas as pd

def main():
    st.title("Excelファイル読み込みテスト")

    uploaded_file = st.file_uploader("Excelファイルをアップロードしてください", type=["xlsx", "xls", "xlsm"])
    if uploaded_file is None:
        st.info("ファイルをアップロードすると内容を表示します。")
        return

    # アップロードファイルを一時保存（問題切り分け用）
    with open("temp_uploaded_file.xlsx", "wb") as f:
        f.write(uploaded_file.getbuffer())

    try:
        # engine指定なしで読み込みを試みる
        df = pd.read_excel(uploaded_file)
        st.success("ファイルを正常に読み込みました。")
        st.dataframe(df.head())
    except Exception as e:
        st.error(f"Excelファイルの読み込みに失敗しました: {e}")

if __name__ == "__main__":
    main()
