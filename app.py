import streamlit as st
import pandas as pd
import re
import json
from unidecode import unidecode
import os
import io
import deepl

MANUAL_CACHE_FILE = "manual_cache.json"
AUTO_CACHE_FILE = "auto_cache.json"

def load_cache():
    manual_cache = {}
    auto_cache = {}
    if os.path.exists(MANUAL_CACHE_FILE):
        with open(MANUAL_CACHE_FILE, "r", encoding="utf-8") as f:
            manual_cache = json.load(f)
    if os.path.exists(AUTO_CACHE_FILE):
        with open(AUTO_CACHE_FILE, "r", encoding="utf-8") as f:
            auto_cache = json.load(f)
    return manual_cache, auto_cache

def save_cache(manual_cache, auto_cache):
    with open(MANUAL_CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(manual_cache, f, ensure_ascii=False, indent=2)
    with open(AUTO_CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(auto_cache, f, ensure_ascii=False, indent=2)

def normalize_brackets(text):
    return text.replace("(", "（").replace(")", "）")

def japanese_to_romaji(text):
    return unidecode(text)

def clean_text(text):
    text = re.sub(r'[\x00-\x1F\x7F]', '', text)
    text = re.sub(r'[\r\n\t]', ' ', text)
    text = re.sub(r'[^\u3040-\u30FF\u4E00-\u9FFFa-zA-Z0-9\s（）.,\-＋()・]', '', text)
    return text.strip()

def translate_text(text, translator, manual_cache, auto_cache):
    if pd.isna(text) or text.strip() == "":
        return ""
    text = str(text)
    text = normalize_brackets(text)
    text = text.replace("℃", "C")

    if text in manual_cache:
        return manual_cache[text]

    for jp, en in manual_cache.items():
        if jp in text:
            text = text.replace(jp, en)

    japanese_parts = re.findall(r'[\u3040-\u30FF\u4E00-\u9FFF]+', text)

    for jp in japanese_parts:
        jp = clean_text(jp)
        if not jp:
            continue

        if len(jp.encode("utf-8")) > 4500:
            st.warning(f"翻訳スキップ（長すぎ）: {jp[:50]}...")
            en = japanese_to_romaji(jp)
            auto_cache[jp] = en
            continue

        if jp in auto_cache:
            en = auto_cache[jp]
        else:
            try:
                result = translator.translate_text(jp, source_lang="JA", target_lang="EN-US")
                en = result.text
                auto_cache[jp] = en
            except Exception as e:
                st.error(f"DeepL翻訳エラー: '{jp}' → 例外: {e}")
                en = japanese_to_romaji(jp)
                auto_cache[jp] = en

        text = text.replace(jp, en)

    remaining = re.findall(r'[\u3040-\u30FF\u4E00-\u9FFF]+', text)
    for jp in remaining:
        text = text.replace(jp, japanese_to_romaji(jp))

    return text

def manual_cache_editor(manual_cache):
    st.sidebar.header("🔧 手動翻訳キャッシュ編集")
    with st.sidebar.form("manual_cache_form"):
        key = st.text_input("日本語語句（キー）")
        val = st.text_input("英語訳（値）")
        submitted = st.form_submit_button("追加 / 更新")

        if submitted:
            if key.strip() and val.strip():
                manual_cache[key.strip()] = val.strip()
                save_cache(manual_cache, {})  # 自動キャッシュは更新なし
                st.sidebar.success(f"キャッシュを更新しました: '{key}' → '{val}'")

    st.sidebar.markdown("---")
    st.sidebar.subheader("現在の手動キャッシュ")
    st.sidebar.write(manual_cache)

def main():
    st.title("Excelサンプル名列の日本語→英語 翻訳アプリ（DeepL + 手動キャッシュ対応）")

    manual_cache, auto_cache = load_cache()

    DEEPL_API_KEY = st.secrets.get("DEEPL_API_KEY")
    if not DEEPL_API_KEY:
        st.error("DEEPL_API_KEYがSecretsに登録されていません。")
        st.stop()

    translator = deepl.Translator(DEEPL_API_KEY)

    manual_cache_editor(manual_cache)

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

    if "サンプル名" not in df.columns:
        st.error("Excelファイルに「サンプル名」列が見つかりません。列名を確認してください。")
        return

    if st.button("翻訳を実行"):
        st.info("翻訳処理中...しばらくお待ちください。")

        texts_to_translate = df["サンプル名"].dropna().unique().tolist()
        translated_map = {}

        for text in texts_to_translate:
            translated_map[text] = translate_text(text, translator, manual_cache, auto_cache)

        df["英語訳"] = df["サンプル名"].map(translated_map)

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
