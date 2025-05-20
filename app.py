import streamlit as st
import pandas as pd
import re
import json
from unidecode import unidecode
import os
import deepl
import io
from pathlib import Path

# キャッシュファイルのパス（ユーザーのデスクトップ）
CACHE_FILE = Path(r"C:\Users\1117\Desktop\python勉強\translation-app\translation_cache.json")

# キャッシュファイルを毎回読み込む（@st.cache_resource は外す）
def load_cache():
    if CACHE_FILE.exists():
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                cache = json.load(f)
            return cache.get("manual", {}), cache.get("auto", {})
        except Exception as e:
            st.error(f"キャッシュファイルの読み込みでエラー: {e}")
            return {}, {}
    else:
        return {}, {}

def save_cache(manual_cache, auto_cache):
    try:
        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump({"manual": manual_cache, "auto": auto_cache}, f, ensure_ascii=False, indent=2)
    except Exception as e:
        st.error(f"キャッシュファイルの保存でエラー: {e}")

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

    # 手動キャッシュで置換（部分一致置換）
    for jp, en in manual_cache.items():
        if jp in text:
            text = text.replace(jp, en)

    # 日本語部分のみ抽出して翻訳 or キャッシュ利用
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

    # まだ残る日本語をローマ字に置換
    remaining = re.findall(r'[\u3040-\u30FF\u4E00-\u9FFF]+', text)
    for jp in remaining:
        text = text.replace(jp, japanese_to_romaji(jp))

    return text

def main():
    st.title("Excel日本語→英語 翻訳アプリ (DeepL API使用)")

    # キャッシュを毎回読み込む
    manual_cache, auto_cache = load_cache()
    st.write("手動キャッシュ例:", list(manual_cache.items())[:3])  # 読み込み確認用表示

    DEEPL_API_KEY = st.secrets.get("DEEPL_API_KEY")
    if not DEEPL_API_KEY:
        st.error("DEEPL_API_KEYがSecretsに登録されていません。")
        st.stop()

    translator = deepl.Translator(DEEPL_API_KEY)

    uploaded_file = st.file_uploader("翻訳するExcelファイルをアップロードしてください", type=["xlsx", "xls", "xlsm"])
    if not uploaded_file:
        st.info("ファイルをアップロードすると翻訳処理を開始します。")
        return

    try:
        df = pd.read_excel(io.BytesIO(uploaded_file.read()), engine="openpyxl")
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
