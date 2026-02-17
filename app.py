import streamlit as st
import pandas as pd
import pdfplumber
import re
import io
import os
import unicodedata
from streamlit_gsheets import GSheetsConnection
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Border, Side, Font

# --- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Tecama Hub Industrial", layout="wide", page_icon="🏗️")

# --- 2. CSS PARA VISUAL v6.6 ---
st.markdown("""
    <style>
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] label { font-size: 22px !important; font-weight: 600 !important; color: #333 !important; }
    h1 { color: #FF5722 !important; font-family: 'Segoe UI', sans-serif; }
    .home-link .stButton > button {
        background-color: transparent !important; color: #FF5722 !important; border: none !important;
        font-size: 24px !important; font-weight: bold !important; text-align: left !important;
        padding: 0 !important; height: auto !important; text-decoration: underline !important;
    }
    .stButton > button {
        background-color: #FF5722; color: white; width: 100%; border-radius: 12px;
        font-weight: bold; height: 3.5em; font-size: 16px; border: none;
    }
    </style>
    """, unsafe_allow_html=True)

conn = st.connection("gsheets", type=GSheetsConnection)

# --- 3. FUNÇÕES AUXILIARES ---
def norm(t):
    """Limpeza profunda: remove acentos, quebras de linha e espaços duplos"""
    if t is None or pd.isna(t): return ""
    t = unicodedata.normalize("NFD", str(t).upper()).encode("ascii", "ignore").decode("utf-8")
    return " ".join(t.split()).strip()

# --- 4. NAVEGAÇÃO ---
if 'nav' not in st.session_state: st.session_state.nav = "🏠 Início"

with st.sidebar:
    if os.path.exists("logo_tecama.png"): st.image("logo_tecama.png", use_container_width=True)
    opcao = st.radio("NAVEGAÇÃO", ["🏠 Início", "🌲 Marcenaria", "⚙️ Metalurgia"], 
                     index=["🏠 Início", "🌲 Marcenaria", "⚙️ Metalurgia"].index(st.session_state.nav))
    st.session_state.nav = opcao
    st.caption("Tecama Hub Industrial v7.9")

# ==========================================
# PÁGINA: INÍCIO (v6.6)
# ==========================================
if st.session_state.nav == "🏠 Início":
    st.title("Tecama Hub Industrial")
    st.markdown("### Bem-vindo ao Sistema Unificado de Produção")
    st.write("Esta plataforma centraliza as operações das divisões integradas ao sistema **Pontta**.")
    st.markdown("---")
    st.markdown('<div class="home-link">', unsafe_allow_html=True)
    if st.button("🌲 Divisão de Marcenaria"):
        st.session_state.nav = "🌲 Marcenaria"; st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)
    st.write("Processamento de arquivos CSV (Pontta) com cálculo automático de pesos.")
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="home-link">', unsafe_allow_html=True)
    if st.button("⚙️ Divisão de Metalurgia"):
        st.session_state.nav = "⚙️ Metalurgia"; st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)
    st.write("Levantamento automático de peso através do relatório PDF (Pontta).")

# ==========================================
# PÁGINA: METALURGIA
# ==========================================
elif st.session_state.nav == "⚙️ Metalurgia":
    st.header("⚙️ Metalurgia")
    aba_calc, aba_db = st.tabs(["📋 Calculadora PDF", "🛠️ Gerenciar Tabelas"])
    
    # Carregamento de dados com TTL baixo para teste
    try:
        db_map = conn.read(worksheet="MAPEAMENTO_TIPO", ttl=2).to_dict('records')
        db_metro = conn.read(worksheet="PESO_POR_METRO", ttl=2)
        db_conj = conn.read(worksheet="PESO_CONJUNTO", ttl=2).to_dict('records')
        
        dict_metro = dict(zip(db_metro['secao'].apply(norm), db_metro['peso_kg_m']))
    except:
        st.error("Erro ao conectar com as tabelas do Google Sheets.")

    with aba_calc:
        up_pdf = st.file_uploader("Suba o PDF Pontta", type="pdf")
        if up_pdf:
            itens = []
            with pdfplumber.open(up_pdf) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        for r in table:
                            if r and len(r) > 3 and str(r[0]).strip().replace('.','').isdigit():
                                itens.append({"QTD": r[0], "DESCRIÇÃO": r[1], "MEDIDA": r[3], "COR": r[2]})
            
            df_edit = st.data_editor(pd.DataFrame(itens), num_rows="dynamic", use_container_width=True)
            
            if st.button("🚀 Calcular Pesos"):
                res = []
                for _, r in df_edit.iterrows():
                    desc_bruta = str(r.get('DESCRIÇÃO', ''))
                    desc_limpa = norm(desc_bruta)
                    qtd = float(r.get('QTD', 0)) if r.get('QTD') else 0.0
                    tipo = "DESCONHECIDO"
                    
                    # 1. Identifica o TIPO
                    for regra in db_map:
                        txt_regra = norm(regra.get('texto_contido', ''))
                        if txt_regra and txt_regra in desc_limpa:
                            tipo = str(regra.get('tipo', 'DESCONHECIDO'))
                            break
                    
                    if tipo == "IGNORAR": continue

                    # 2. Calcula o PESO UNITÁRIO
                    p_unit = 0.0
                    if tipo == "CONJUNTO":
                        # Busca o peso do conjunto por correspondência de texto
                        for c_regra in db_conj:
                            nome_cadastrado = norm(c_regra.get('nome_conjunto', ''))
                            if nome_cadastrado and nome_cadastrado in desc_limpa:
                                p_unit = float(c_regra.get('peso_unit_kg', 0))
                                break
                    elif "TUBO" in tipo:
                        medida = 0.0
                        try:
                            med_str = str(r.get('MEDIDA', '0')).lower().replace('mm','').replace(',','.').strip()
                            medida = float(med_str)
                        except: medida = 0.0
                        sec_key = norm(tipo.replace('TUBO ', '').strip())
                        p_unit = (medida / 1000) * dict_metro.get(sec_key, 0.0)
                    
                    res.append({
                        "QTD": qtd, "DESCRIÇÃO": desc_bruta, "MEDIDA": r.get('MEDIDA', ''),
                        "TIPO": tipo, "PESO UNIT.": round(p_unit, 3), 
                        "PESO TOTAL": round(p_unit * qtd, 3)
                    })
                
                res_df = pd.DataFrame(res)
                st.metric("Total Geral", f"{res_df['PESO TOTAL'].sum():.2f} kg")
                st.dataframe(res_df, use_container_width=True)

    with aba_db:
        st.write("Gerencie as tabelas base abaixo:")
        # (Restante da aba de gerenciamento igual à v7.4)
