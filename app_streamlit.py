# -*- coding: utf-8 -*-
import streamlit as st
import tempfile, os, io
from dotenv import load_dotenv, set_key
from agent import run_agent, AVAILABLE_MODELS, PROVIDER_ENV_KEYS
from schema import ExtractedInfo
from image_enhancer import PRESETS
from translations import TRANSLATIONS

load_dotenv()
ENV_FILE = ".env"

st.set_page_config(
    page_title="AI Document Extractor",
    layout="wide",
    initial_sidebar_state="expanded"
)



OCR_LANGUAGES = {
    "English":"eng","Chinese Simplified":"chi_sim","Chinese Traditional":"chi_tra",
    "French":"fra","German":"deu","Spanish":"spa","Japanese":"jpn","Korean":"kor",
    "Arabic":"ara","Portuguese":"por","Russian":"rus","Italian":"ita",
    "Dutch":"nld","Hindi":"hin","Vietnamese":"vie",
}
PROVIDER_COLORS = {
    "OpenAI":"green","Anthropic":"orange","Google":"blue",
    "DeepSeek":"purple","xAI":"gray","Alibaba":"red",
}
PROVIDER_ICONS = {
    "OpenAI":"🟢","Anthropic":"🟠","Google":"🔵",
    "DeepSeek":"🟣","xAI":"⚫","Alibaba":"🔴",
}

# ── Helpers ───────────────────────────────────────────────────────────────────
def save_api_key(env_var, value):
    if not os.path.exists(ENV_FILE): open(ENV_FILE,"w").close()
    set_key(ENV_FILE, env_var, value)
    os.environ[env_var] = value

def get_masked_key(env_var):
    val = os.getenv(env_var,"")
    if not val: return ""
    return val[:6]+"•"*(len(val)-10)+val[-4:] if len(val)>10 else "••••••••"

def export_json(r): return r.model_dump_json(indent=2).encode("utf-8")

def export_txt(r):
    lines = ["AI DOCUMENT EXTRACTION REPORT","="*50,"",
             "[ TITLE ]", r.title,"",
             "[ SUMMARY ]", r.summary,"",
             "[ KEY TOPICS ]",*[f"  - {t}" for t in r.key_topics],"",
             "[ NAMED ENTITIES ]", f"  {', '.join(r.entities) or 'None'}","",
             "[ DATES ]", f"  {', '.join(r.dates) if r.dates else 'None'}",""]
    if r.action_items: lines += ["[ ACTION ITEMS ]",*[f"  - {a}" for a in r.action_items],""]
    lines += ["[ SENTIMENT ]", f"  {r.sentiment or 'N/A'}"]
    return "\n".join(lines).encode("utf-8")

def export_csv(r):
    import csv; out=io.StringIO(); w=csv.writer(out)
    w.writerow(["Section","Value"])
    w.writerow(["Title",r.title]); w.writerow(["Summary",r.summary])
    [w.writerow(["Key Topic",t]) for t in r.key_topics]
    [w.writerow(["Entity",e]) for e in r.entities]
    [w.writerow(["Date",d]) for d in (r.dates or [])]
    [w.writerow(["Action Item",a]) for a in (r.action_items or [])]
    w.writerow(["Sentiment",r.sentiment or "N/A"])
    return out.getvalue().encode("utf-8")

def export_docx(r):
    from docx import Document
    from docx.shared import Pt, RGBColor, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    import datetime
    doc=Document()
    for s in doc.sections:
        s.top_margin=Inches(1); s.bottom_margin=Inches(1)
        s.left_margin=Inches(1.2); s.right_margin=Inches(1.2)
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    run=p.add_run("AI DOCUMENT EXTRACTION REPORT")
    run.bold=True; run.font.size=Pt(18); run.font.color.rgb=RGBColor(0x4B,0x4B,0xFF)
    doc.add_paragraph(f"Generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}").alignment=WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("-"*60)
    def sec(title,lines):
        h=doc.add_heading(title,level=1); h.runs[0].font.color.rgb=RGBColor(0x4B,0x4B,0xFF)
        for l in lines: doc.add_paragraph(l); doc.add_paragraph()
    sec("Title",[r.title]); sec("Summary",[r.summary])
    sec("Key Topics",[f"- {t}" for t in r.key_topics])
    sec("Named Entities",[", ".join(r.entities) or "None"])
    sec("Dates",[", ".join(r.dates) if r.dates else "None"])
    if r.action_items: sec("Action Items",[f"- {a}" for a in r.action_items])
    sec("Sentiment",[r.sentiment or "N/A"])
    buf=io.BytesIO(); doc.save(buf); return buf.getvalue()

def export_excel(r):
    import openpyxl, datetime
    from openpyxl.styles import Font,PatternFill,Alignment
    wb=openpyxl.Workbook(); ws=wb.active; ws.title="Summary"
    hf=Font(bold=True,color="FFFFFF",size=11); hfill=PatternFill("solid",fgColor="4B4BFF")
    alt=PatternFill("solid",fgColor="F5F5FF")
    ws["A1"]="AI Document Extraction Report"; ws["A1"].font=Font(bold=True,size=13,color="4B4BFF")
    ws["C1"]=datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    ws.append([])
    ws.append(["Field","Value"])
    for c in ws[3]: c.font=hf; c.fill=hfill; c.alignment=Alignment(horizontal="center")
    rows=[("Title",r.title),("Summary",r.summary),("Sentiment",r.sentiment or "N/A"),
          ("Key Topics"," | ".join(r.key_topics)),
          ("Named Entities"," | ".join(r.entities) or "None"),
          ("Dates"," | ".join(r.dates) if r.dates else "None"),
          ("Action Items"," | ".join(r.action_items) if r.action_items else "None")]
    for i,(f,v) in enumerate(rows,4):
        ws.cell(i,1,f).font=Font(bold=True); ws.cell(i,2,v).alignment=Alignment(wrap_text=True,vertical="top")
        if i%2==0:
            ws.cell(i,1).fill=alt; ws.cell(i,2).fill=alt
    ws.column_dimensions["A"].width=18; ws.column_dimensions["B"].width=90
    ws2=wb.create_sheet("Key Topics"); ws2.append(["#","Topic"])
    for c in ws2[1]: c.font=hf; c.fill=hfill
    for i,t in enumerate(r.key_topics,1): ws2.append([i,t]); ws2.cell(i+1,2).alignment=Alignment(wrap_text=True)
    ws2.column_dimensions["B"].width=50
    ws3=wb.create_sheet("Entities & Dates"); ws3.append(["Named Entity","Date"])
    for c in ws3[1]: c.font=hf; c.fill=hfill
    for i in range(max(len(r.entities),len(r.dates or []))):
        ws3.append([r.entities[i] if i<len(r.entities) else "",
                    r.dates[i] if r.dates and i<len(r.dates) else ""])
    ws3.column_dimensions["A"].width=40; ws3.column_dimensions["B"].width=30
    if r.action_items:
        ws4=wb.create_sheet("Action Items"); ws4.append(["#","Action Item","Status"])
        for c in ws4[1]: c.font=hf; c.fill=hfill
        for i,a in enumerate(r.action_items,1): ws4.append([i,a,"Pending"])
        ws4.column_dimensions["B"].width=70
    buf=io.BytesIO(); wb.save(buf); return buf.getvalue()

def build_preview(r,T):
    sent=(r.sentiment or "neutral").lower()
    sent_label={"positive":T["positive"],"neutral":T["neutral"],"negative":T["negative"]}.get(sent,sent)
    sent_class={"positive":"sent-positive","neutral":"sent-neutral","negative":"sent-negative"}.get(sent,"sent-neutral")
    tags=lambda lst:"".join(f'<span class="preview-tag">{x}</span>' for x in lst) or "None"
    actions=""
    if r.action_items:
        actions=f'<div class="preview-section"><div class="preview-label">{T["action_items"]}</div><div class="preview-value">{"".join(f"<div>- {a}</div>" for a in r.action_items)}</div></div>'
    return f"""<div class="preview-box">
    <div class="preview-section"><div class="preview-label">{T["title"]}</div><div class="preview-value" style="font-size:1.1rem;font-weight:600">{r.title}</div></div>
    <div class="preview-section"><div class="preview-label">{T["sentiment"]}</div><div class="preview-value"><span class="sentiment-badge {sent_class}">{sent_label}</span></div></div>
    <div class="preview-section"><div class="preview-label">{T["summary"]}</div><div class="preview-value">{r.summary}</div></div>
    <div class="preview-section"><div class="preview-label">{T["key_topics"]}</div><div class="preview-value">{tags(r.key_topics)}</div></div>
    <div class="preview-section"><div class="preview-label">{T["entities"]}</div><div class="preview-value">{tags(r.entities)}</div></div>
    <div class="preview-section"><div class="preview-label">{T["dates"]}</div><div class="preview-value">{tags(r.dates or [])}</div></div>
    {actions}</div>"""

# Defaults — prevent NameError on re-render
T = TRANSLATIONS["English"]
provider_choice = sorted(set(v[0] for v in AVAILABLE_MODELS.values()))[0]
model_choice = list(AVAILABLE_MODELS.keys())[0]
ocr_preset = list(PRESETS.keys())[0]
ocr_lang = "eng"

# ── SIDEBAR ───────────────────────────────────────────────────────────────────
with st.sidebar:
    ui_lang = st.selectbox("UI Language", options=list(TRANSLATIONS.keys()), index=0, key="ui_language")
    T = TRANSLATIONS[ui_lang]
    st.markdown("---")
    st.header("Settings")

    st.subheader("Model")
    providers = sorted(set(v[0] for v in AVAILABLE_MODELS.values()))
    provider_choice = st.selectbox(T["provider"], options=providers,
        format_func=lambda p: f"{PROVIDER_ICONS.get(p, '')} {p}")
    provider_models = {k:v for k,v in AVAILABLE_MODELS.items() if v[0]==provider_choice}
    model_choice = st.selectbox(T["model_label"], options=list(provider_models.keys()),
        format_func=lambda m: f"{m}  —  {AVAILABLE_MODELS[m][2]}")
    _,cost,desc = AVAILABLE_MODELS[model_choice]
    st.info(f"Cost: {cost}\n\n{desc}")

    st.markdown("---")
    st.subheader("API Keys")
    env_key = PROVIDER_ENV_KEYS[provider_choice]
    is_set = bool(os.getenv(env_key,""))
    if is_set:
        st.success(f"Set: `{env_key}`\n\n`{get_masked_key(env_key)}`")
        if st.button(f"Remove {provider_choice} Key", use_container_width=True):
            save_api_key(env_key,""); os.environ.pop(env_key,None); st.rerun()
    else:
        st.warning(f"Not set: `{env_key}`")
    with st.expander(f"{'Update' if is_set else 'Add'} {provider_choice} API Key", expanded=not is_set):
        new_key = st.text_input("Paste API key", type="password", placeholder="e.g. sk-...", key=f"key_{provider_choice}")
        c1,c2=st.columns(2)
        with c1:
            if st.button("Save Key", type="primary", use_container_width=True):
                if new_key.strip(): save_api_key(env_key,new_key.strip()); st.success("Saved!"); st.rerun()
                else: st.error("Key cannot be empty.")
        with c2:
            links={"OpenAI":"https://platform.openai.com/api-keys","Anthropic":"https://console.anthropic.com/",
                   "Google":"https://aistudio.google.com/app/apikey","DeepSeek":"https://platform.deepseek.com/",
                   "xAI":"https://console.x.ai/","Alibaba":"https://dashscope.console.aliyun.com/"}
            st.link_button("Get Key", url=links.get(provider_choice,"#"), use_container_width=True)
    with st.expander("All API Keys Status"):
        for p,ev in PROVIDER_ENV_KEYS.items():
            v=os.getenv(ev,""); st.markdown(f"{'OK' if v else '--'} {PROVIDER_ICONS.get(p, '')} **{p}** — `{get_masked_key(ev) if v else 'not set'}`")

    st.markdown("---")
    st.subheader("OCR")
    ocr_preset = st.selectbox(T["ocr_preset"], options=list(PRESETS.keys()), index=0)
    ocr_lang_name = st.selectbox(T["ocr_language"], options=list(OCR_LANGUAGES.keys()), index=0)
    ocr_lang = OCR_LANGUAGES[ocr_lang_name]
    if st.toggle(T["combine_eng"], value=False, help=T["combine_help"]):
        if ocr_lang!="eng": ocr_lang=f"{ocr_lang}+eng"
    st.caption(f"{T['ocr_code']} `{ocr_lang}`")
    st.markdown("---")
    st.markdown("**OCR Preset Guide**")
    st.markdown(f"- `scanner` — {T['preset_scanner']}\n- `phone_camera` — {T['preset_phone']}\n- `faded` — {T['preset_faded']}\n- `fax` — {T['preset_fax']}")

# ── MAIN ──────────────────────────────────────────────────────────────────────
st.title(T["app_title"])
st.markdown(T["app_subtitle"])

uploaded = st.file_uploader(T["drop_file"], type=["pdf","docx"])
if uploaded:
    st.info(f"File: `{uploaded.name}` — {uploaded.size/1024:.1f} KB")
    if st.button(T["extract_btn"], type="primary"):
        if not os.getenv(PROVIDER_ENV_KEYS[provider_choice]):
            st.error(T["no_api_key"].format(provider_choice)); st.stop()
        suffix=os.path.splitext(uploaded.name)[-1]
        with tempfile.NamedTemporaryFile(delete=False,suffix=suffix) as tmp:
            tmp.write(uploaded.read()); tmp_path=tmp.name
        with st.spinner(f"{T['analyzing']} `{model_choice}`..."):
            try:
                result=run_agent(tmp_path,ocr_preset=ocr_preset,lang=ocr_lang,model=model_choice)
                os.unlink(tmp_path)
                st.session_state["result"]=result
                st.session_state["model_used"]=model_choice
                st.session_state["provider_used"]=provider_choice
            except Exception as e:
                st.error(f"Error: {e}")
                if os.path.exists(tmp_path): os.unlink(tmp_path)

if "result" in st.session_state:
    r=st.session_state["result"]
    mu=st.session_state["model_used"]; pu=st.session_state["provider_used"]
    st.success(f"{T['extracted_with']} `{mu}` ({PROVIDER_ICONS.get(pu,'')} {pu})")

    tab1,tab2,tab3=st.tabs(["Results","Preview Export","Download"])

    with tab1:
        col1,col2=st.columns(2)
        with col1:
            st.subheader(T["title"]); st.write(r.title)
            st.subheader(T["summary"]); st.write(r.summary)
            st.subheader(T["sentiment"])
            sent=(r.sentiment or "neutral").lower()
            emoji={"positive":"green","neutral":"yellow","negative":"red"}.get(sent,"gray")
            label={"positive":T["positive"],"neutral":T["neutral"],"negative":T["negative"]}.get(sent,sent)
            st.write(f":{emoji}[{label}]")
        with col2:
            st.subheader(T["key_topics"])
            for t in r.key_topics: st.markdown(f"- {t}")
            st.subheader(T["entities"]); st.write(", ".join(r.entities) or T["none_found"])
            st.subheader(T["dates"]); st.write(", ".join(r.dates) if r.dates else T["none_found"])
        if r.action_items:
            st.subheader(T["action_items"])
            for a in r.action_items: st.checkbox(a,value=False,key=f"act_{a}")

    with tab2:
        st.markdown("### Export Preview")
        st.markdown(build_preview(r,T),unsafe_allow_html=True)
        st.markdown("---")
        st.markdown("**JSON Preview**"); st.json(r.model_dump())
        st.markdown("**TXT Preview**"); st.code(export_txt(r).decode("utf-8"),language="text")

    with tab3:
        st.markdown("### Download")
        st.caption("All formats contain the same information, structured differently.")
        c1,c2,c3,c4,c5=st.columns(5)
        with c1: st.download_button("JSON",data=export_json(r),file_name="extracted.json",mime="application/json",use_container_width=True)
        with c2: st.download_button("TXT",data=export_txt(r),file_name="extracted.txt",mime="text/plain",use_container_width=True)
        with c3: st.download_button("CSV",data=export_csv(r),file_name="extracted.csv",mime="text/csv",use_container_width=True)
        with c4: st.download_button("DOCX",data=export_docx(r),file_name="extracted.docx",mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",use_container_width=True)
        with c5: st.download_button("Excel",data=export_excel(r),file_name="extracted.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
        st.markdown("---\n- **JSON** — raw structured data\n- **TXT** — readable flat report\n- **CSV** — one row per item\n- **DOCX** — styled Word document\n- **Excel** — 4-sheet workbook")