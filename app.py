import streamlit as st
import pandas as pd
import re
import json
from unidecode import unidecode
import deepl
import io

# キャッシュ初期化（アップロードで上書き）
manual_cache = {}
auto_cache = {}

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

    # 手動キャッシュ置換
    for jp, en in manual_cache.items():
        text = text.replace(jp, en)

    # 日本語部分だけ抽出して翻訳 or キャッシュ参照
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

    # 残った日本語をローマ字に変換
    remaining = re.findall(r'[\u3040-\u30FF\u4E00-\u9FFF]+', text)
    for jp in remaining:
        text = text.replace(jp, japanese_to_romaji(jp))

    return text

def main():
    st.title("Excel日本語→英語 翻訳アプリ (DeepL API使用)")

    st.markdown("### 1. キャッシュファイルをアップロードしてください（無ければ空で進めます）")
    uploaded_cache_file = st.file_uploader("キャッシュファイル（JSON）", type=["json"], key="cache_uploader")

    global manual_cache, auto_cache
    if uploaded_cache_file is not None:
        try:
            cache_json = json.load(uploaded_cache_file)
            manual_cache = cache_json.get("manual", {})
            auto_cache = cache_json.get("auto", {})
            st.success(f"キャッシュ読み込み成功: manual {len(manual_cache)}件, auto {len(auto_cache)}件")
        except Exception as e:
            st.error(f"キャッシュファイルの読み込みエラー: {e}")
    else:
        st.info("キャッシュファイルをアップロードしてください。")

    st.markdown("### 2. 翻訳するExcelファイルをアップロードしてください")
    uploaded_file = st.file_uploader("Excelファイル", type=["xlsx", "xls", "xlsm"], key="excel_uploader")
    if not uploaded_file:
        st.info("Excelファイルをアップロードすると翻訳できます。")
        return

    try:
        df = pd.read_excel(io.BytesIO(uploaded_file.read()), engine="openpyxl")
    except Exception as e:
        st.error(f"Excel読み込みエラー: {e}")
        return

    st.write("アップロードされたデータのプレビュー")
    st.dataframe(df.head())

    target_col = "サンプル名"
    if target_col not in df.columns:
        st.error(f"列 '{target_col}' が存在しません。")
        return

    DEEPL_API_KEY = st.secrets.get("DEEPL_API_KEY")
    if not DEEPL_API_KEY:
        st.error("DEEPL_API_KEYが設定されていません。")
        return

    translator = deepl.Translator(DEEPL_API_KEY)

    if st.button("翻訳を実行"):
        st.info("翻訳処理中...しばらくお待ちください。")

        texts_to_translate = df[target_col].dropna().unique().tolist()
        translated_map = {}

        for text in texts_to_translate:
            translated_map[text] = translate_text(text, translator, manual_cache, auto_cache)

        df["英語名"] = df[target_col].map(translated_map)

        st.success("翻訳が完了しました。")
        st.dataframe(df.head())

        # 更新キャッシュをJSONとしてダウンロードできるように準備
        cache_json_str = json.dumps({"manual": manual_cache, "auto": auto_cache}, ensure_ascii=False, indent=2)
        st.download_button(
            label="更新されたキャッシュファイルをダウンロード",
            data=cache_json_str,
            file_name="translation_cache_updated.json",
            mime="application/json"
        )

        # 翻訳結果Excelファイルを保存してダウンロード
        output_file = uploaded_file.name.rsplit(".", 1)[0] + "_translated.xlsx"
        with io.BytesIO() as output_bytes:
            with pd.ExcelWriter(output_bytes, engine="openpyxl") as writer:
                df.to_excel(writer, index=False)
            data = output_bytes.getvalue()
            st.download_button(
                label="翻訳済みExcelファイルをダウンロード",
                data=data,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
