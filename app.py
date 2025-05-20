import streamlit as st
import pandas as pd
import re
import json
import os
import requests
from unidecode import unidecode

# DeepL API設定（ここにあなたのAPIキーを入力してください）
DEEPL_API_KEY = st.secrets["DEEPL_API_KEY"]
DEEPL_API_URL = "https://api-free.deepl.com/v2/translate"

# キャッシュファイル名
CACHE_FILE = "translation_cache.json"

# キャッシュ読み込み
def load_cache():
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    else:
        return {"manual": {}, "auto": {}}

# キャッシュ保存
def save_cache(cache):
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)

# 括弧正規化
def normalize_brackets(text):
    return text.replace("(", "（").replace(")", "）")

# ローマ字変換
def japanese_to_romaji(text):
    return unidecode(text)

# DeepL翻訳関数
def deepl_translate(text):
    try:
        params = {
            "auth_key": DEEPL_API_KEY,
            "text": text,
            "source_lang": "JA",
            "target_lang": "EN"
        }
        response = requests.post(DEEPL_API_URL, data=params)
        response.raise_for_status()
        return response.json()["translations"][0]["text"]
    except Exception as e:
        st.error(f"DeepL翻訳エラー: {e}")
        return japanese_to_romaji(text)

# 翻訳処理
def translate_text(text, cache):
    if pd.isna(text) or text.strip() == "":
        return ""

    text = str(text)
    text = normalize_brackets(text)
    text = text.replace("℃", "C")

    for jp, en in cache["manual"].items():
        text = text.replace(jp, en)

    japanese_parts = re.findall(r'[\u3040-\u30FF\u4E00-\u9FFF]+', text)

    for part in japanese_parts:
        if part in cache["auto"]:
            translated = cache["auto"][part]
        else:
            translated = deepl_translate(part)
            cache["auto"][part] = translated
        text = text.replace(part, translated)

    # 未翻訳の残りをローマ字に
    remaining = re.findall(r'[\u3040-\u30FF\u4E00-\u9FFF]+', text)
    for part in remaining:
        text = text.replace(part, japanese_to_romaji(part))

    return text

# Streamlit UI
st.title("Excel翻訳アプリ（DeepL API対応）")
uploaded_file = st.file_uploader("Excelファイルをアップロードしてください", type=["xlsx", "xls", "xlsm"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    cache = load_cache()

    if "サンプル名" not in df.columns:
        st.error("Excelに「サンプル名」列が見つかりません。")
    else:
        if st.button("翻訳実行"):
            with st.spinner("翻訳中..."):
                df["英語名"] = df["サンプル名"].apply(lambda x: translate_text(x, cache))
                save_cache(cache)

            st.success("翻訳完了！")
            st.dataframe(df)

            output = df.to_excel(index=False, engine="openpyxl")
            st.download_button(
                label="翻訳結果をダウンロード",
                data=output,
                file_name="translated.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
