import streamlit as st
import pandas as pd
import re
import json
from unidecode import unidecode
import os
import deepl

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
    if pd.isna(text) or text.strip() == "":
        return ""
    text = str(text)
    text = normalize_brackets(text)
    text = text.replace("℃", "C")

    for jp, en in manual_cache.items():
        text = text.replace(jp, en)

    japanese_parts = extract_japanese_parts(text)

    for jp in japanese_parts:
        if jp in auto_cache:
            en = auto_cache[jp]
        else:
            if len(jp) > 2000:  # DeepLの文字数制限(約5000文字)をかなり余裕持って短縮
                st.warning(f"文字列が長すぎるため省略します: {jp[:50]}...")
                en = japanese_to_romaji(jp)
                auto_cache[jp] = en
                continue
            try:
                en = translator.translate_text(jp, source_lang="JA", target_lang="EN-US").text
                auto_cache[jp] = en
            except Exception as e:
                error_msg = f"DeepL翻訳エラー: '{jp}' の翻訳に失敗しました。例外: {e}"
                st.error(error_msg)
                print(error_msg)
                en = japanese_to_romaji(jp)
                auto_cache[jp] = en
        text = text.replace(jp, en)

    remaining = extract_japanese_parts(text)
    for jp in remaining:
        text = text.replace(jp, japanese_to_romaji(jp))

    return text

def main():
    st.title("Excel日本語→英語 翻訳アプリ (DeepL API使用)")

    manual_cache, auto_cache = load_cache()

    DEEPL_API_KEY = st.secrets.get("DEEPL_API_KEY")
    if not DEEPL_API_KEY:
        st.error("DEEPL_API_KEYがSecretsに登録されていません。設定してください。")
        st.stop()

    translator = deepl.Translator(DEEPL_API_KEY)

    uploaded_file = st.file_uploader("翻訳するExcelファイルをアップロードしてください", type=["xlsx", "xls", "xlsm"])
    if not uploaded_file:
        st.info("ファイルをアップロードすると翻訳処理を開始します。")
        return

    try:
        df = pd.read_excel(uploaded_file, engine="openpyxl")
    except Exception as e:
        st.error(f"Excelファイルの読み込みに失敗しました: {e}")
        return

    st.write("アップロードされたデータのプレビュー")
    st.dataframe(df.head())

    target_col = "サンプル名"
    if target_col not in df.columns:
        st.error(f"列 '{target_col}' がファイルに存在しません。")
        return

    texts_to_translate = df[target_col].dropna().unique().tolist()
    st.write(f"翻訳対象テキスト（{target_col}列）")
    st.write(texts_to_translate)

    if st.button("翻訳を実行"):
        st.info("翻訳処理中...しばらくお待ちください。")

        translated_map = {}
        for text in texts_to_translate:
            translated_map[text] = translate_text(text, translator, manual_cache, auto_cache)

        df["英語名"] = df[target_col].map(translated_map)

        save_cache(manual_cache, auto_cache)

        st.success("翻訳が完了しました。")
        st.dataframe(df.head())

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
