import streamlit as st
import pandas as pd
import re
import json
from unidecode import unidecode
import io
from pathlib import Path
import zipfile
import deepl

# キャッシュファイルのパス（初期値は空、アップロードファイルから読み込み）
manual_cache = {}
auto_cache = {}

def load_cache_from_file(file) -> tuple[dict, dict]:
    try:
        cache = json.load(file)
        manual = cache.get("manual", {})
        auto = cache.get("auto", {})
        return manual, auto
    except Exception as e:
        st.error(f"キャッシュファイルの読み込みに失敗しました: {e}")
        return {}, {}

def save_cache(manual_cache, auto_cache, path: Path):
    with open(path, "w", encoding="utf-8") as f:
        json.dump({"manual": manual_cache, "auto": auto_cache}, f, ensure_ascii=False, indent=2)

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
    if text is None or (isinstance(text, float) and pd.isna(text)):
        return ""
    text = str(text).strip()
    if text == "":
        return ""

    text = normalize_brackets(text)
    text = text.replace("℃", "C")

    for jp, en in manual_cache.items():
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

    remaining = re.findall(r'[\u3040-\u30FF\u9FFF]+', text)
    for jp in remaining:
        text = text.replace(jp, japanese_to_romaji(jp))

    text = text.replace("/", "_")

    maru_map = {
        "①": "_1", "②": "_2", "③": "_3", "④": "_4", "⑤": "_5",
        "⑥": "_6", "⑦": "_7", "⑧": "_8", "⑨": "_9", "⑩": "_10",
        "⑪": "_11", "⑫": "_12", "⑬": "_13", "⑭": "_14", "⑮": "_15"
    }
    for maru, repl in maru_map.items():
        text = text.replace(maru, repl)

    return text

def main():
    st.title("サンプル名変換 (日本語→英語)")
    st.write("※ファイル名は英数字のみにしてください。「サンプル名」を含む列のみ変換します。")

    global manual_cache, auto_cache

    st.sidebar.info("最初にキャッシュ(JSONファイル)をアップロードしてください。")
    cache_file = st.sidebar.file_uploader("キャッシュファイル(JSON)をアップロード", type=["json"])
    if cache_file is not None:
        manual_cache, auto_cache = load_cache_from_file(cache_file)
        st.sidebar.success("キャッシュファイルを読み込みました。")

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
        translated_map = {}
        for text in texts_to_translate:
            translated_map[text] = translate_text(text, translator, manual_cache, auto_cache)

        df["英語名"] = df[target_col].map(translated_map)

        output_excel_name = uploaded_file.name.rsplit(".", 1)[0] + "_translated.xlsx"

        # Excelバイト列作成
        excel_bytes_io = io.BytesIO()
        df.to_excel(excel_bytes_io, index=False, engine="openpyxl")
        excel_bytes_io.seek(0)

        # キャッシュJSON文字列作成
        cache_json_str = json.dumps({"manual": manual_cache, "auto": auto_cache}, ensure_ascii=False, indent=2)
        cache_bytes_io = io.BytesIO(cache_json_str.encode("utf-8"))

        # キャッシュ保存（任意ファイル保存）
        cache_save_path = Path(output_excel_name).parent / "translation_cache.json"
        try:
            save_cache(manual_cache, auto_cache, cache_save_path)
        except Exception as e:
            st.warning(f"キャッシュの保存に失敗しました: {e}")

        st.toast("翻訳が完了しました。", icon="✅")

        st.write("翻訳後のExcelプレビュー")
        st.dataframe(df)

        # 横並びにボタン配置
        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                label="翻訳済みExcelをダウンロード",
                data=excel_bytes_io,
                file_name=output_excel_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with col2:
            st.download_button(
                label="キャッシュファイルをダウンロード",
                data=cache_bytes_io,
                file_name="translation_cache.json",
                mime="application/json"
            )

        # 英語名コピー機能はそのまま
        if "英語名" in df.columns:
            english_names = df["英語名"].dropna().astype(str).tolist()
            english_text = "\n".join(english_names).replace("`", "\\`").replace("\\", "\\\\")
            st.markdown("#### 英語名リストのコピー")
            st.text_area("コピー対象", english_text, height=200)

            copy_button_html = f"""
            <button onclick=\"navigator.clipboard.writeText(`{english_text}`); alert('英語名をコピーしました！');\" 
                style=\"
                    background-color:#4CAF50;
                    color:white;
                    padding:10px 16px;
                    font-size:16px;
                    border:none;
                    border-radius:6px;
                    cursor:pointer;
                    margin-top:10px;\">
                \u2705 クリックして英語名をコピー
            </button>
            """
            st.markdown(copy_button_html, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
