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

# --- 2. CSS PERSONALIZADO (VISUAL v6.6 TRAVADO) ---
st.markdown("""
    <style>
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] label { font-size: 22px !important; font-weight: 600 !important; color: #333 !important; }
    h1 { color: #FF5722 !important; font-family: 'Segoe UI', sans-serif; }
    h3 { color: #444 !important; }
    .home-link .stButton > button {
        background-color: transparent !important; color: #FF5722 !important; border: none !important;
        font-size: 24px !important; font-weight: bold !important; text-align: left !important;
        padding: 0 !important; height: auto !important; text-decoration: underline !important;
        box-shadow: none !important;
    }
    .stButton > button {
        background-color: #FF5722; color: white; width: 100%; border-radius: 12px;
        font-weight: bold; height: 3.5em; font-size: 16px; border: none;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
    }
    .stButton > button:hover { background-color: #E64A19; transform: translateY(-2px); }
    </style>
    """, unsafe_allow_html=True)

conn = st.connection("gsheets", type=GSheetsConnection)

# --- 3. FUNÇÕES AUXILIARES ---
def norm(t):
    if t is None or pd.isna(t): return ""
    t = unicodedata.normalize("NFD", str(t).upper()).encode("ascii", "ignore").decode("utf-8")
    return " ".join(t.split()).strip()

# --- A ALTERAÇÃO SOLICITADA ESTÁ AQUI ---
def limpar_material_apenas_cor(t):
    """Remove 'CHAPA', 'MDF', 'MDP' e espessuras para manter apenas a cor"""
    t = norm(t)
    # Remove qualquer número seguido de MM (ex: 18MM, 15 MM)
    t = re.sub(r'\d+\s*MM', '', t) 
    # Remove os termos específicos
    termos_para_remover = ["CHAPA DE", "CHAPA", "MDF", "MDP", "HDF", "DURATEX", "ARACO"]
    for termo in termos_para_remover:
        t = t.replace(termo, "")
    return t.strip()

def calcular_pesos_madeira(larg, comp, quant, material_texto):
    PESO_M2_BASE = {"MDP": 12.0, "MDF": 13.5}
    try:
        l, c, q = float(str(larg).replace(',','.')), float(str(comp).replace(',','.')), float(str(quant).replace(',','.'))
        m_norm = norm(material_texto)
        tipo = "MDF" if "MDF" in m_norm else "MDP"
        esp_match = re.search(r"(\d+)\s*MM", m_norm)
        e = float(esp_match.group(1)) if esp_match else 18.0
        peso_uni = (l/1000) * (c/1000) * PESO_M2_BASE[tipo] * (e/18)
        return round(peso_uni, 2), round(peso_uni * q, 2)
    except: return 0.0, 0.0

# --- 4. NAVEGAÇÃO ---
if 'nav' not in st.session_state: st.session_state.nav = "🏠 Início"

with st.sidebar:
    if os.path.exists("logo_tecama.png"): st.image("logo_tecama.png", use_container_width=True)
    opcao = st.radio("NAVEGAÇÃO", ["🏠 Início", "🌲 Marcenaria", "⚙️ Metalurgia"], 
                     index=["🏠 Início", "🌲 Marcenaria", "⚙️ Metalurgia"].index(st.session_state.nav))
    st.session_state.nav = opcao

# ==========================================
# PÁGINA: INÍCIO (v6.6 INTEGRAL)
# ==========================================
if st.session_state.nav == "🏠 Início":
    st.title("Tecama Hub Industrial")
    st.markdown("### Bem-vindo ao Sistema Unificado de Produção")
    st.write("Esta plataforma foi desenvolvida para centralizar as operações das divisões de **Marcenaria** e **Metalurgia**, garantindo agilidade no processamento de pedidos e precisão nos cálculos de engenharia.")
    st.markdown("---")
    if st.button("🌲 Divisão de Marcenaria"):
        st.session_state.nav = "🌲 Marcenaria"; st.rerun()
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("⚙️ Divisão de Metalurgia"):
        st.session_state.nav = "⚙️ Metalurgia"; st.rerun()

# ==========================================
# PÁGINA: MARCENARIA
# ==========================================
elif st.session_state.nav == "🌲 Marcenaria":
    st.header("🌲 Marcenaria")
    aba_conv, aba_cores = st.tabs(["📋 Conversor de Arquivos", "🎨 Gestão de Cores"])
    
    with aba_conv:
        up_csv = st.file_uploader("Suba o CSV do Pontta", type="csv")
        if up_csv:
            df_b = pd.read_csv(up_csv, sep=None, engine='python', dtype=str)
            df_b.columns = [norm(c) for c in df_b.columns]
            if st.button("🚀 Gerar Excel de Produção"):
                # APLICAÇÃO DA LIMPEZA NA COLUNA MATERIAL
                df_b["MATERIAL"] = df_b["MATERIAL"].apply(limpar_material_apenas_cor)
                
                pesos = df_b.apply(lambda r: calcular_pesos_madeira(r.get("LARG",0), r.get("COMP",0), r.get("QUANT",0), r.get("MATERIAL","")), axis=1)
                df_b["PESO_UNIT"] = pesos.apply(lambda x: x[0]); df_b["PESO_TOTAL"] = pesos.apply(lambda x: x[1])
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    ws = writer.book.create_sheet("PRODUCAO")
                    header = ["QUANT","COMP","LARG","MATERIAL","COR","DESCPECA","DES_PAI","PESO UNIT.","PESO TOTAL"]
                    for i, h in enumerate(header, 1):
                        ws.cell(row=3, column=i, value=h).font = Font(bold=True)
                    
                    for idx, r in df_b.iterrows():
                        row_idx = idx + 4
                        vals = [r.get("QUANT"), r.get("COMP"), r.get("LARG"), r.get("MATERIAL"), r.get("COR"), r.get("DESCPECA"), r.get("DES_PAI"), r.get("PESO_UNIT"), r.get("PESO_TOTAL")]
                        for i, v in enumerate(vals, 1):
                            ws.cell(row=row_idx, column=i, value=v).border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
                st.download_button("📥 Baixar Excel", output.getvalue(), "PRODUCAO_TECAMA.xlsx")

# ==========================================
# PÁGINA: METALURGIA (v6.6 INTEGRAL)
# ==========================================
elif st.session_state.nav == "⚙️ Metalurgia":
    st.header("⚙️ Metalurgia")
    m1, m2 = st.tabs(["📋 Calculadora", "🛠️ Tabelas Base"])
    # Tabelas base restauradas como solicitado
    if m2:
        if 't_m' not in st.session_state: st.session_state.t_m = "MAPEAMENTO_TIPO"
        c1, c2, c3 = st.columns(3)
        if c1.button("Mapeamento"): st.session_state.t_m = "MAPEAMENTO_TIPO"
        if c2.button("Tubos"): st.session_state.t_m = "PESO_POR_METRO"
        if c3.button("Conjuntos"): st.session_state.t_m = "PESO_CONJUNTO"
        df_v = conn.read(worksheet=st.session_state.t_m, ttl=0)
        st.data_editor(df_v, use_container_width=True)
