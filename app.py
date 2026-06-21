"""
Guideline Reference Extractor — App Streamlit
Modos:
  1. Subir PDF → extraer referencias
  2. Buscar guía por nombre → CrossRef devuelve DOI → pipeline completo
  3. Introducir DOI directamente
"""

import streamlit as st
import tempfile, os, time, re, requests
from pipeline import extract_references_from_pdf, enrich_reference, classify_reference, export_to_excel

st.set_page_config(page_title="Guideline Reference Extractor", page_icon="📚", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@300;400;600&family=IBM+Plex+Mono&display=swap');
html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif; }
.hero { background: linear-gradient(135deg, #0A2342 0%, #1B4F8A 100%); border-radius: 12px; padding: 2rem; margin-bottom: 1.5rem; }
.hero h1 { font-size: 1.8rem; font-weight: 600; margin: 0 0 0.3rem 0; color: white; }
.hero p  { font-size: 0.95rem; opacity: 0.85; margin: 0; color: white; }
.search-result { background: white; border: 1px solid #E3EAF4; border-radius: 8px;
    padding: 0.9rem 1.1rem; margin-bottom: 0.5rem; cursor: pointer;
    transition: border-color 0.15s; }
.search-result:hover { border-color: #1B4F8A; }
.search-result .sr-title { font-weight: 600; font-size: 0.9rem; color: #0A2342; }
.search-result .sr-meta  { font-size: 0.78rem; color: #6B7A99; margin-top: 2px; }
.search-result .sr-doi   { font-size: 0.75rem; color: #1565C0; font-family: 'IBM Plex Mono', monospace; }
.selected-guide { background: #E8F0FB; border: 1.5px solid #1B4F8A; border-radius: 10px; padding: 1rem 1.2rem; margin-bottom: 1rem; }
.selected-guide .sg-title { font-weight: 600; font-size: 0.95rem; color: #0A2342; }
.selected-guide .sg-doi   { font-size: 0.8rem; color: #1565C0; font-family: 'IBM Plex Mono', monospace; margin-top: 3px; }
.badge-rct   { background:#E8F5E9; color:#2E7D32; border:1px solid #A5D6A7; border-radius:6px; padding:2px 8px; font-size:0.78rem; font-weight:600; }
.badge-rct2  { background:#FFF8E1; color:#F57F17; border:1px solid #FFE082; border-radius:6px; padding:2px 8px; font-size:0.78rem; font-weight:600; }
.badge-meta  { background:#E3F2FD; color:#1565C0; border:1px solid #90CAF9; border-radius:6px; padding:2px 8px; font-size:0.78rem; font-weight:600; }
.badge-reg   { background:#FCE4EC; color:#880E4F; border:1px solid #F48FB1; border-radius:6px; padding:2px 8px; font-size:0.78rem; font-weight:600; }
.badge-guide { background:#EDE7F6; color:#4527A0; border:1px solid #B39DDB; border-radius:6px; padding:2px 8px; font-size:0.78rem; font-weight:600; }
.badge-other { background:#F5F5F5; color:#616161; border:1px solid #BDBDBD; border-radius:6px; padding:2px 8px; font-size:0.78rem; font-weight:600; }
.ref-row    { background: white; border: 1px solid #E8EDF5; border-radius: 8px; padding: 1rem 1.2rem; margin-bottom: 0.5rem; }
.ref-number { font-family: 'IBM Plex Mono', monospace; font-size: 0.75rem; color: #9AA3B5; margin-bottom: 2px; }
.ref-title  { font-weight: 600; font-size: 0.9rem; color: #0A2342; margin-bottom: 3px; }
.ref-meta   { font-size: 0.8rem; color: #6B7A99; }
.ref-links a { font-size: 0.78rem; color: #1565C0; text-decoration: none; margin-right: 12px; }
#MainMenu {visibility:hidden;} footer {visibility:hidden;} header {visibility:hidden;}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
  <h1>📚 Guideline Reference Extractor</h1>
  <p>Extrae, enriquece y clasifica todas las referencias de cualquier guía clínica.<br>
  Obtén PMID, DOI, autores, año y tipo de estudio en un Excel listo para usar.</p>
</div>
""", unsafe_allow_html=True)

# ── Constantes ─────────────────────────────────────────────────────────────────

BADGE_MAP = {
    "RCT_primario":           '<span class="badge-rct">ECA primario</span>',
    "RCT_secundario":         '<span class="badge-rct2">ECA secundario</span>',
    "meta-analisis":          '<span class="badge-meta">Meta-análisis</span>',
    "registro_observacional": '<span class="badge-reg">Registro/Cohorte</span>',
    "guia_clinica":           '<span class="badge-guide">Guía clínica</span>',
    "otro/no_clasificado":    '<span class="badge-other">Otro</span>',
}
COLOR_MAP = {
    "RCT_primario":"#E8F5E9","RCT_secundario":"#FFF8E1","meta-analisis":"#E3F2FD",
    "registro_observacional":"#FCE4EC","guia_clinica":"#EDE7F6","otro/no_clasificado":"#F5F5F5",
}
TYPE_LABELS = {
    "RCT_primario":"ECAs primarios","RCT_secundario":"ECAs secundarios",
    "meta-analisis":"Meta-análisis","registro_observacional":"Registros",
    "guia_clinica":"Guías","otro/no_clasificado":"Otros",
}
CLASSIFICATION_RULES = {
    "RCT_primario":           [r'\brandomis[ei]d\b',r'\bplacebo.controlled\b',r'\bblind(ed)?\b',
                                r'\brandom(ized|ised)\s+(clinical|controlled)\s+trial\b',r'\bRCT\b'],
    "RCT_secundario":         [r'\bsubgroup\s+anal',r'\bpost.hoc\b',r'\bsecondary\s+anal',r'\bsub-?study\b'],
    "meta-analisis":          [r'\bmeta.anal',r'\bsystematic\s+review\b',r'\bpooled\s+anal'],
    "registro_observacional": [r'\bregist(ry|er|ro)\b',r'\bcohort\b',r'\bobservational\b',r'\bretrospective\b'],
    "guia_clinica":           [r'\bguideline\b',r'\brecommendation\b',r'\bconsensus\s+(statement|document)\b'],
}

def classify(title, pub_type=""):
    text = (title+" "+pub_type).lower()
    for st_type in ["RCT_secundario","meta-analisis","guia_clinica","registro_observacional","RCT_primario"]:
        for pat in CLASSIFICATION_RULES[st_type]:
            if re.search(pat, text, re.IGNORECASE):
                return st_type
    return "otro/no_clasificado"

# ── CrossRef helpers ───────────────────────────────────────────────────────────

CROSSREF_HEADERS = {"User-Agent": "GuidelineRefExtractor/1.0 (research tool; mailto:researcher@example.com)"}

def search_guidelines_crossref(query: str) -> list[dict]:
    """Busca guías clínicas en CrossRef por texto libre."""
    try:
        r = requests.get(
            "https://api.crossref.org/works",
            params={
                "query.bibliographic": query,
                "rows": 8,
                "select": "DOI,title,author,published,container-title,type,publisher",
                "filter": "type:journal-article",
            },
            headers=CROSSREF_HEADERS, timeout=15
        )
        items = r.json().get("message", {}).get("items", [])
        results = []
        for item in items:
            title = (item.get("title") or [""])[0]
            if not title: continue
            pub = item.get("published", {}).get("date-parts", [[""]])
            year = str(pub[0][0]) if pub and pub[0] else ""
            journal = (item.get("container-title") or [""])[0]
            authors = item.get("author", [])
            first_author = f"{authors[0].get('family','')} et al." if authors else ""
            results.append({
                "doi": item.get("DOI",""),
                "title": title,
                "year": year,
                "journal": journal,
                "first_author": first_author,
            })
        return results
    except Exception as e:
        return []

def get_refs_crossref(doi: str) -> list[dict]:
    try:
        r = requests.get(f"https://api.crossref.org/works/{doi}",
                         headers=CROSSREF_HEADERS, timeout=30)
        return r.json().get("message", {}).get("reference", [])
    except:
        return []

def enrich_doi_crossref(doi: str) -> dict:
    try:
        r = requests.get(f"https://api.crossref.org/works/{doi}",
                         headers=CROSSREF_HEADERS, timeout=10)
        item = r.json().get("message", {})
        al = item.get("author", [])
        authors = ", ".join([
            f"{a.get('family','')} {a.get('given','')[0]}." if a.get('given') else a.get('family','')
            for a in al[:3]
        ])
        if len(al) > 3: authors += " et al."
        pub = item.get("published",{}).get("date-parts",[[""]])
        year = str(pub[0][0]) if pub and pub[0] else ""
        return {
            "title":       (item.get("title") or [""])[0],
            "authors":     authors,
            "year":        year,
            "journal":     (item.get("container-title") or [""])[0],
            "pub_type_raw": item.get("type",""),
        }
    except:
        return {}

def search_pmid(title: str, year: str = "", api_key: str = "") -> str:
    words = re.findall(r'\b[A-Za-z]{4,}\b', title)
    query = " ".join(words[:8])
    if year: query += f"[Title/Abstract] AND {year}[PDAT]"
    params = {"db":"pubmed","term":query,"retmax":1,"retmode":"json"}
    if api_key: params["api_key"] = api_key
    try:
        r = requests.get("https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi",
                         params=params, timeout=10)
        ids = r.json().get("esearchresult",{}).get("idlist",[])
        return ids[0] if ids else ""
    except:
        return ""

def process_crossref_refs(raw_refs, guideline_name, ncbi_key, progress_bar, status_text):
    records = []
    total = len(raw_refs)
    for idx, ref in enumerate(raw_refs, 1):
        progress_bar.progress(int(idx/total*90))
        record = {
            "ref_number": idx,
            "authors": "",
            "year":    ref.get("year",""),
            "title":   ref.get("article-title","") or ref.get("unstructured","")[:150],
            "journal": ref.get("journal-title","") or ref.get("volume-title",""),
            "doi":     ref.get("DOI",""),
            "pmid": "", "pubmed_url": "", "doi_url": "",
            "study_type_auto": "", "study_type": "",
            "pub_type_raw": "", "notes": "",
            "ref_raw": ref.get("unstructured",""),
        }
        status_text.markdown(
            f'<p style="font-size:0.8rem;color:#6B7A99;">[{idx}/{total}] {record["title"][:80]}...</p>',
            unsafe_allow_html=True)
        if record["doi"]:
            time.sleep(0.2)
            enriched = enrich_doi_crossref(record["doi"])
            for k in ["title","authors","year","journal","pub_type_raw"]:
                if enriched.get(k) and not record.get(k):
                    record[k] = enriched[k]
            record["doi_url"] = f"https://doi.org/{record['doi']}"
        if record["title"] and len(record["title"]) > 15:
            time.sleep(0.35)
            pmid = search_pmid(record["title"], record.get("year",""), ncbi_key)
            if pmid:
                record["pmid"] = pmid
                record["pubmed_url"] = f"https://pubmed.ncbi.nlm.nih.gov/{pmid}/"
        record["study_type_auto"] = classify(record["title"], record.get("pub_type_raw",""))
        records.append(record)
    return records

# ── Render resultados ──────────────────────────────────────────────────────────

def build_excel_bytes(records):
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp_path = tmp.name
    export_to_excel(records, tmp_path)
    with open(tmp_path,"rb") as f: data = f.read()
    os.unlink(tmp_path)
    return data

def render_results(records, file_label):
    excel_bytes = build_excel_bytes(records)
    counts = {}
    for r in records:
        t = r.get("study_type_auto","otro/no_clasificado")
        counts[t] = counts.get(t,0)+1

    st.download_button("⬇️  Descargar Excel completo", data=excel_bytes,
        file_name=f"referencias_{file_label}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True, type="primary")
    st.markdown("---")
    st.markdown("#### Resumen")
    cols = st.columns(min(len(counts),4))
    for i,(t,c) in enumerate(counts.items()):
        with cols[i%4]:
            st.markdown(f"""<div style="background:{COLOR_MAP.get(t,'#eee')};border-radius:8px;
                padding:1rem;text-align:center;border-top:3px solid #ccc;margin-bottom:8px;">
                <div style="font-size:2rem;font-weight:600;color:#0A2342;">{c}</div>
                <div style="font-size:0.8rem;color:#6B7A99;">{TYPE_LABELS.get(t,t)}</div>
                </div>""", unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("#### Referencias")
    c1,c2 = st.columns([2,1])
    with c1: q = st.text_input("Buscar...", placeholder="título, autor, año", label_visibility="collapsed", key=f"search_{file_label}")
    with c2: tf = st.selectbox("Tipo",["Todos"]+list(counts.keys()), label_visibility="collapsed", key=f"type_{file_label}")
    filtered = records
    if tf != "Todos": filtered = [r for r in filtered if r.get("study_type_auto")==tf]
    if q:
        ql = q.lower()
        filtered = [r for r in filtered if ql in (r.get("title","")).lower()
                    or ql in (r.get("authors","")).lower() or ql in (r.get("year","")).lower()]
    st.markdown(f'<p style="font-size:0.82rem;color:#9AA3B5;">{len(filtered)} referencias</p>', unsafe_allow_html=True)
    for r in filtered:
        badge = BADGE_MAP.get(r.get("study_type_auto",""),"")
        title = r.get("title") or r.get("ref_raw","")[:100]
        authors = r.get("authors",""); year = r.get("year",""); journal = r.get("journal","")
        meta = " · ".join(filter(None,[authors[:60]+("..." if len(authors)>60 else ""),year,journal[:40]]))
        pmid = r.get("pmid",""); doi = r.get("doi","")
        links = ""
        if pmid: links += f'<a href="https://pubmed.ncbi.nlm.nih.gov/{pmid}/" target="_blank">PubMed →</a>'
        if doi:  links += f'<a href="https://doi.org/{doi}" target="_blank">DOI →</a>'
        st.markdown(f"""<div class="ref-row">
            <div class="ref-number">#{r.get('ref_number','')} &nbsp;{badge}</div>
            <div class="ref-title">{title}</div>
            <div class="ref-meta">{meta}</div>
            {"<div class='ref-links'>"+links+"</div>" if links else ""}
            </div>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════════════════════

tab1, tab2 = st.tabs(["📄  Subir PDF", "🔍  Buscar guía por nombre o DOI"])

# ── TAB 1: PDF ────────────────────────────────────────────────────────────────
with tab1:
    col_l, col_r = st.columns([1,2], gap="large")
    with col_l:
        st.markdown("#### Subir guía clínica en PDF")
        uploaded = st.file_uploader("PDF", type=["pdf"], label_visibility="collapsed")
        st.markdown('<p style="font-size:0.82rem;color:#6B7A99;">Funciona con cualquier guía en PDF.<br>Layouts de 1 y 2 columnas.</p>', unsafe_allow_html=True)
    with col_r:
        if uploaded is None:
            st.markdown("""<div style="background:#F8FAFD;border:2px dashed #C5D4E8;border-radius:12px;
                padding:3rem 2rem;text-align:center;color:#6B7A99;">
                <div style="font-size:3rem;margin-bottom:1rem;">📄</div>
                <div style="font-size:1.1rem;font-weight:600;color:#0A2342;margin-bottom:0.5rem;">Sube un PDF para empezar</div>
                <div style="font-size:0.85rem;">El extractor detectará la sección de referencias automáticamente.</div>
                </div>""", unsafe_allow_html=True)
        else:
            if "pdf_records" not in st.session_state or st.session_state.get("pdf_file") != uploaded.name:
                if st.button("🔍 Extraer y enriquecer referencias", type="primary", use_container_width=True, key="btn_pdf"):
                    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                        tmp.write(uploaded.read()); tmp_path = tmp.name
                    prog = st.progress(0); stat = st.empty()
                    stat.markdown("**Extrayendo referencias del PDF...**")
                    raw_refs = extract_references_from_pdf(tmp_path)
                    prog.progress(10)
                    records = []
                    for i, ref_text in enumerate(raw_refs):
                        prog.progress(10+int(i/len(raw_refs)*85))
                        stat.markdown(f'<p style="font-size:0.8rem;color:#6B7A99;">[{i+1}/{len(raw_refs)}] {ref_text[:70]}...</p>', unsafe_allow_html=True)
                        record = enrich_reference(ref_text, i+1)
                        record["study_type_auto"] = classify_reference(record)
                        records.append(record)
                    os.unlink(tmp_path); prog.progress(100); stat.empty()
                    st.session_state["pdf_records"] = records
                    st.session_state["pdf_file"] = uploaded.name
                    st.rerun()
            if "pdf_records" in st.session_state and st.session_state.get("pdf_file") == uploaded.name:
                render_results(st.session_state["pdf_records"], uploaded.name.replace(".pdf",""))
                if st.button("🔄 Procesar otro PDF", key="reset_pdf"):
                    del st.session_state["pdf_records"], st.session_state["pdf_file"]
                    st.rerun()

# ── TAB 2: BÚSQUEDA / DOI ─────────────────────────────────────────────────────
with tab2:
    col_l2, col_r2 = st.columns([1,2], gap="large")

    with col_l2:
        # ── Modo de entrada ────────────────────────────────────────────────────
        mode = st.radio("Modo de búsqueda", ["🔍 Buscar por nombre", "🔗 Introducir DOI directamente"],
                        label_visibility="collapsed")
        st.markdown("---")

        doi_final = ""
        nombre_final = ""

        if mode == "🔍 Buscar por nombre":
            st.markdown("#### Buscar guía")
            search_query = st.text_input("Nombre de la guía",
                placeholder="Ej: ESC heart failure 2021, ACC AHA hypertension 2023...",
                label_visibility="collapsed")

            if search_query and len(search_query) > 5:
                if st.button("Buscar", use_container_width=True):
                    with st.spinner("Buscando en CrossRef..."):
                        results = search_guidelines_crossref(search_query)
                    st.session_state["search_results"] = results
                    st.session_state["search_query"] = search_query
                    if "selected_doi" in st.session_state:
                        del st.session_state["selected_doi"]

            # Mostrar resultados de búsqueda
            if "search_results" in st.session_state and st.session_state.get("search_query") == search_query:
                results = st.session_state["search_results"]
                if not results:
                    st.warning("No se encontraron resultados. Prueba con otro término o usa el DOI directamente.")
                else:
                    st.markdown(f'<p style="font-size:0.82rem;color:#6B7A99;margin-bottom:8px;">{len(results)} resultados — selecciona la guía correcta:</p>', unsafe_allow_html=True)
                    for i, res in enumerate(results):
                        label = f"**{res['title'][:70]}{'...' if len(res['title'])>70 else ''}**\n\n{res.get('first_author','')} · {res.get('year','')} · {res.get('journal','')[:40]}\n\n`{res['doi']}`"
                        if st.button(f"{res['title'][:55]}{'...' if len(res['title'])>55 else ''} ({res.get('year','')})",
                                     key=f"sel_{i}", use_container_width=True):
                            st.session_state["selected_doi"] = res["doi"]
                            st.session_state["selected_nombre"] = res["title"]
                            st.rerun()

            # Guía seleccionada
            if "selected_doi" in st.session_state:
                doi_final = st.session_state["selected_doi"]
                nombre_final = st.session_state.get("selected_nombre", doi_final)
                st.markdown(f"""<div class="selected-guide">
                    <div style="font-size:0.75rem;color:#1B4F8A;font-weight:600;margin-bottom:4px;">✓ GUÍA SELECCIONADA</div>
                    <div class="sg-title">{nombre_final[:80]}</div>
                    <div class="sg-doi">{doi_final}</div>
                    </div>""", unsafe_allow_html=True)
                if st.button("✕ Cambiar selección", use_container_width=False):
                    del st.session_state["selected_doi"]
                    if f"doi_records_{doi_final}" in st.session_state:
                        del st.session_state[f"doi_records_{doi_final}"]
                    st.rerun()

        else:  # DOI directo
            st.markdown("#### DOI de la guía")
            doi_input = st.text_input("DOI", placeholder="10.1093/eurheartj/ehad191",
                                       label_visibility="collapsed")
            nombre_input = st.text_input("Nombre (opcional)", placeholder="ESC Heart Failure 2023",
                                          label_visibility="collapsed")
            doi_final = doi_input.strip().lstrip("https://doi.org/")
            nombre_final = nombre_input.strip() or doi_final

        st.markdown("---")
        st.markdown("#### API Key NCBI *(opcional)*")
        ncbi_key = st.text_input("NCBI API Key", type="password",
            placeholder="Deja vacío si no tienes", label_visibility="collapsed",
            help="Obtener gratis en ncbi.nlm.nih.gov/account · Acelera ×3 la búsqueda de PMIDs")
        st.markdown('<p style="font-size:0.78rem;color:#9AA3B5;">Sin key: ~20 min/guía · Con key: ~5 min</p>', unsafe_allow_html=True)

    # ── Panel derecho: procesamiento y resultados ──────────────────────────────
    with col_r2:
        cache_key = f"doi_records_{doi_final}"

        if not doi_final:
            st.markdown("""<div style="background:#F8FAFD;border:2px dashed #C5D4E8;border-radius:12px;
                padding:3rem 2rem;text-align:center;color:#6B7A99;">
                <div style="font-size:3rem;margin-bottom:1rem;">🔍</div>
                <div style="font-size:1.1rem;font-weight:600;color:#0A2342;margin-bottom:0.5rem;">
                Busca una guía o introduce su DOI</div>
                <div style="font-size:0.85rem;">
                Funciona con cualquier guía indexada en CrossRef:<br>
                ESC, ACC/AHA, AHA, NICE, SIGN, y muchas más.</div>
                </div>""", unsafe_allow_html=True)

        elif cache_key not in st.session_state:
            st.markdown(f"""<div class="selected-guide" style="margin-bottom:1.5rem;">
                <div style="font-size:0.75rem;color:#1B4F8A;font-weight:600;margin-bottom:4px;">GUÍA A PROCESAR</div>
                <div class="sg-title">{nombre_final[:100]}</div>
                <div class="sg-doi">{doi_final}</div>
                </div>""", unsafe_allow_html=True)

            if st.button(f"🚀 Extraer todas las referencias", type="primary", use_container_width=True, key="btn_doi"):
                prog2 = st.progress(0); stat2 = st.empty()
                stat2.markdown("**Conectando con CrossRef...**")
                raw_refs = get_refs_crossref(doi_final)
                if not raw_refs:
                    st.error(f"CrossRef no encontró referencias para el DOI `{doi_final}`.\n\n"
                             "Posibles causas:\n- DOI incorrecto\n- La guía no está indexada en CrossRef con referencias\n\n"
                             "En ese caso usa el modo **Subir PDF**.")
                else:
                    prog2.progress(5)
                    stat2.markdown(f"**{len(raw_refs)} referencias encontradas. Enriqueciendo...**")
                    records = process_crossref_refs(raw_refs, nombre_final, ncbi_key, prog2, stat2)
                    prog2.progress(100); stat2.empty()
                    st.session_state[cache_key] = records
                    st.rerun()

        else:
            records = st.session_state[cache_key]
            render_results(records, doi_final.replace("/","_"))
            if st.button("🔄 Procesar otra guía", key="reset_doi"):
                del st.session_state[cache_key]
                for k in ["selected_doi","selected_nombre","search_results","search_query"]:
                    if k in st.session_state: del st.session_state[k]
                st.rerun()
