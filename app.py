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

# --- 1. CONFIGURAÇÃO ---
st.set_page_config(page_title="Tecama Hub Industrial", layout="wide", page_icon="🏗️")

# --- 2. CSS v6.6 ORIGINAL ---
st.markdown("""
    <style>
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] label { font-size: 22px !important; font-weight: 600 !important; color: #333 !important; }
    h1 { color: #FF5722 !important; font-family: 'Segoe UI', sans-serif; }
    .stButton > button {
        background-color: #FF5722; color: white; width: 100%; border-radius: 12px;
        font-weight: bold; height: 3.5em; font-size: 16px; border: none;
    }
    .home-link .stButton > button {
        background-color: transparent !important; color: #FF5722 !important; border: none !important;
        font-size: 24px !important; font-weight: bold !important; text-align: left !important;
        padding: 0 !important; height: auto !important; text-decoration: underline !important;
    }
    </style>
    """, unsafe_allow_html=True)

conn = st.connection("gsheets", type=GSheetsConnection)

def norm(t):
    if t is None or pd.isna(t): return ""
    t = unicodedata.normalize("NFD", str(t).upper()).encode("ascii", "ignore").decode("utf-8")
    return " ".join(t.split()).strip()

# --- 3. NAVEGAÇÃO ---
if 'nav' not in st.session_state: st.session_state.nav = "🏠 Início"

with st.sidebar:
    if os.path.exists("logo_tecama.png"): st.image("logo_tecama.png", use_container_width=True)
    opcao = st.radio("NAVEGAÇÃO", ["🏠 Início", "🌲 Marcenaria", "⚙️ Metalurgia"], 
                     index=["🏠 Início", "🌲 Marcenaria", "⚙️ Metalurgia"].index(st.session_state.nav))
    st.session_state.nav = opcao

# ==========================================
# PÁGINA: INÍCIO
# ==========================================
if st.session_state.nav == "🏠 Início":
    st.title("Tecama Hub Industrial")
    st.markdown("### Bem-vindo ao Sistema Unificado de Produção")
    st.write("Esta plataforma foi desenvolvida para centralizar as operações das divisões de **Marcenaria** e **Metalurgia**.")
    st.markdown("---")
    st.markdown('<div class="home-link">', unsafe_allow_html=True)
    if st.button("🌲 Divisão de Marcenaria"): 
        st.session_state.nav = "🌲 Marcenaria"
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="home-link">', unsafe_allow_html=True)
    if st.button("⚙️ Divisão de Metalurgia"): 
        st.session_state.nav = "⚙️ Metalurgia"
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

# ==========================================
# PÁGINA: MARCENARIA
# ==========================================
elif st.session_state.nav == "🌲 Marcenaria":
    st.header("🌲 Marcenaria")
    aba1, aba2 = st.tabs(["📋 Conversor", "🎨 Cores"])
    with aba1:
        up = st.file_uploader("Arquivo CSV")
        if up: st.success("Arquivo pronto para conversão")
    with aba2:
        df_c = conn.read(worksheet="CORES_MARCENARIA", ttl=0)
        st.data_editor(df_c, use_container_width=True)

# ==========================================
# PÁGINA: METALURGIA
# ==========================================
elif st.session_state.nav == "⚙️ Metalurgia":
    st.header("⚙️ Metalurgia")
    m1, m2 = st.tabs(["📋 Calculadora", "🛠️ Tabelas Base"])
    
    with m1:
        up_pdf = st.file_uploader("Relatório PDF")
        if up_pdf: st.info("PDF carregado")
        
    with m2:
        # Recuperação total dos botões de tabela
        col1, col2, col3 = st.columns(3)
        if 'tab_met' not in st.session_state: st.session_state.tab_met = "MAPEAMENTO_TIPO"
        
        if col1.button("📋 Mapeamento"): st.session_state.tab_met = "MAPEAMENTO_TIPO"
        if col2.button("⚖️ Tubos"): st.session_state.tab_met = "PESO_POR_METRO"
        if col3.button("📦 Conjuntos"): st.session_state.tab_met = "PESO_CONJUNTO"
        
        df_v = conn.read(worksheet=st.session_state.tab_met, ttl=0)
        st.subheader(f"Editando: {st.session_state.tab_met}")
        st.data_editor(df_v, num_rows="dynamic", use_container_width=True)
