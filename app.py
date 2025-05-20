import streamlit as st
import pandas as pd
import io

def main():
    st.title("Excelファイル読み込みテスト")

    uploaded_file = st.file_uploader("Excelファイルをアップロードしてください", type=["xlsx", "xls", "xlsm"])
    if uploaded_file is None:
        st.info("ファイルをアップロードすると内容を表示します。")
        return

    try:
        # ファイルのバイナリデータをBytesIOで包んでPandasに渡す
        df = pd.read_excel(io.BytesIO(uploaded_file.read()))
        st.success("ファイルを正常に読み込みました。")
        st.dataframe(df.head())
    except Exception as e:
        st.error(f"Excelファイルの読み込みに失敗しました: {e}")

if __name__ == "__main__":
    main()
