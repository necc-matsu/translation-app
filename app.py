import streamlit as st
import deepl

def main():
    DEEPL_API_KEY = st.secrets.get("DEEPL_API_KEY")
    if not DEEPL_API_KEY:
        st.error("DEEPL_API_KEYがSecretsに設定されていません。")
        st.stop()
    translator = deepl.Translator(DEEPL_API_KEY)
    try:
        result = translator.translate_text("こんにちは", source_lang="JA", target_lang="EN-US")
        st.write("翻訳結果:", result.text)
    except Exception as e:
        st.error(f"翻訳エラー: {e}")

if __name__ == "__main__":
    main()
