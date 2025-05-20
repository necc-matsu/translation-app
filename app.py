import streamlit as st
import pandas as pd
import re
import json
from deep_translator import DeepL
from unidecode import unidecode
import os

# キャッシュファイル名（ローカル保存用。クラウド環境では永続化されない点ご注意）
CACHE_FILE = "translation_cache.json"

@st.cache_resource
def load_cache():
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            cache = json.load(f)
        return cache.get("manual", {}), cache.get("auto", {})
    else:
        return {}, {}

def save_cache(manual_cache, auto_cache):
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump({"manual": manual_cache, "auto": auto_cache}, f, ensure_ascii=False, indent=2)

def normalize_brackets(text):
    return text.replace("(", "（").replace(")", "）")

def japanese_to_romaji(text):
    return unidecode(text)

def extract_japanese_parts(text):
    return re.findall(r'[\u3040-\u30FF\u4E00-\u9FFF]+', text)

def translate_text(text, translator, manual_cache, auto_cache):
    if pd.isna(text):
        return ""
    text = str(text)
    text = normalize_brackets(text)
    text = text.replace("℃", "C")

    # 手動キャッシュによる置換
    for jp, en in manual_cache.items():
        text = text.replace(jp, en)

    japanese_parts = extract_japanese_parts(text)

    for jp in japanese_parts:
        if jp in auto_cache:
            en = auto_cache[jp]
        else:
            try:
                en = translator.translate(jp)
                auto_cache[jp] = en
            except Exception:
                en = japanese_to_romaji(jp)
                auto_cache[jp] = en
        text = text.replace(jp, en)

    # 残った未翻訳の日本語をローマ字に置換
    remaining = extract_japanese_parts(text)
    for jp in remaining:
        text = text.replace(jp, japanese_to_romaji(jp))

    return text

def main():
    st.title("Excel日本語→英語 翻訳アプリ (DeepL API使用)")

    manual_cache, auto_cache = load_cache()

    # SecretsからAPIキー取得
    DEEPL_API_KEY = st.secrets.get("DEEPL_API_KEY")
    if not DEEPL_API_KEY:
        st.error("DEEPL_API_KEYがSecretsに登録されていません。設定してください。")
        st.stop()

    uploaded_file = st.file_uploader("翻訳するExcelファイルをアップロードしてください", type=["xlsx", "xls", "xlsm"])
    if not uploaded_file:
        st.info("ファイルをアップロードすると翻訳処理を開始します。")
        return

    # Excel読み込み
    try:
        df = pd.read_excel(uploaded_file, engine="openpyxl")
    except Exception as e:
        st.error(f"Excelファイルの読み込みに失敗しました: {e}")
        return

    st.write("アップロードされたデータのプレビュー")
    st.dataframe(df.head())

    target_col = "サンプル名"  # 必要に応じてここを変更または選択UI追加

    if target_col not in df.columns:
        st.error(f"列 '{target_col}' がファイルに存在しません。")
        return

    texts_to_translate = df[target_col].dropna().unique().tolist()
    st.write(f"翻訳対象テキスト（{target_col}列）")
    st.write(texts_to_translate)

    if st.button("翻訳を実行"):
        st.info("翻訳処理中...しばらくお待ちください。")

        translator = DeepL(api_key=DEEPL_API_KEY, source="JA", target="EN")

        translated_map = {}
        for text in texts_to_translate:
            translated_map[text] = translate_text(text, translator, manual_cache, auto_cache)

        # 翻訳結果をDataFrameに反映
        df["英語名"] = df[target_col].map(translated_map)

        # キャッシュ保存
        save_cache(manual_cache, auto_cache)

        st.success("翻訳が完了しました。")
        st.dataframe(df.head())

        # ダウンロード用にExcelファイルを作成
        output_file = uploaded_file.name.rsplit(".", 1)[0] + "_translated.xlsx"
        df.to_excel(output_file, index=False, engine="openpyxl")

        with open(output_file, "rb") as f:
            st.download_button(
                label="翻訳済みExcelファイルをダウンロード",
                data=f,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
