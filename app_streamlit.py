import streamlit as st
import tempfile, os
from agent import run_agent
from schema import ExtractedInfo
from image_enhancer import PRESETS

st.set_page_config(page_title="AI Document Extractor", layout="wide")
st.title("AI Document Information Extractor")
st.markdown("Upload a **PDF** or **DOCX** - supports native text and scanned files.")

with st.sidebar:
    st.header("Settings")
    ocr_preset = st.selectbox("OCR Preset", options=list(PRESETS.keys()), index=0)
    lang = st.text_input("OCR Language", value="eng")
    st.markdown("---")
    st.markdown("**Preset Guide**\n- `scanner` - flat-bed scan\n- `phone_camera` - photo/glare\n- `faded` - old/low-contrast\n- `fax` - dot-matrix/fax")

uploaded = st.file_uploader("Drop your file here", type=["pdf", "docx"])

if uploaded:
    st.info(f"File: {uploaded.name} ({uploaded.size / 1024:.1f} KB)")
    if st.button("Extract Information", type="primary"):
        suffix = os.path.splitext(uploaded.name)[-1]
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(uploaded.read())
            tmp_path = tmp.name
        with st.spinner("Analyzing document..."):
            try:
                result: ExtractedInfo = run_agent(tmp_path, ocr_preset=ocr_preset, lang=lang)
                os.unlink(tmp_path)
                col1, col2 = st.columns(2)
                with col1:
                    st.subheader("Title")
                    st.write(result.title)
                    st.subheader("Summary")
                    st.write(result.summary)
                    st.subheader("Sentiment")
                    emoji = {"positive": "🟢", "neutral": "🟡", "negative": "🔴"}.get(
                        (result.sentiment or "neutral").lower(), "⚪")
                    st.write(f"{emoji} {(result.sentiment or 'N/A').capitalize()}")
                with col2:
                    st.subheader("Key Topics")
                    for t in result.key_topics:
                        st.markdown(f"- {t}")
                    st.subheader("Named Entities")
                    st.write(", ".join(result.entities) or "None found")
                    st.subheader("Dates")
                    st.write(", ".join(result.dates) if result.dates else "None found")
                if result.action_items:
                    st.subheader("Action Items")
                    for item in result.action_items:
                        st.checkbox(item, value=False, key=item)
                st.divider()
                st.download_button("Download JSON", data=result.model_dump_json(indent=2),
                                   file_name="extracted_info.json", mime="application/json")
            except Exception as e:
                st.error(f"Error: {str(e)}")
                if os.path.exists(tmp_path):
                    os.unlink(tmp_path)