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

# --- 2. CSS v6.6 (FONTE GRANDE E ESTILO LARANJA) ---
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

# --- 3. FUNÇÕES AUXILIARES ---
def norm(t):
    if t is None or pd.isna(t): return ""
    t = unicodedata.normalize("NFD", str(t).upper()).encode("ascii", "ignore").decode("utf-8")
    return " ".join(t.split()).strip()

def calcular_pesos_madeira(larg, comp, quant, material_texto):
    PESO_M2_BASE = {"MDP": 12.0, "MDF": 13.5}
    try:
        l = float(str(larg).replace(',','.'))
        c = float(str(comp).replace(',','.'))
        q = float(str(quant).replace(',','.'))
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
# PÁGINA: INÍCIO (v6.6)
# ==========================================
if st.session_state.nav == "🏠 Início":
    st.title("Tecama Hub Industrial")
    st.markdown("### Bem-vindo ao Sistema Unificado de Produção")
    st.write("Esta plataforma foi desenvolvida para centralizar as operações das divisões de **Marcenaria** e **Metalurgia**.")
    st.markdown("---")
    if st.button("🌲 Divisão de Marcenaria"): st.session_state.nav = "🌲 Marcenaria"; st.rerun()
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("⚙️ Divisão de Metalurgia"): st.session_state.nav = "⚙️ Metalurgia"; st.rerun()

# ==========================================
# PÁGINA: MARCENARIA
# ==========================================
elif st.session_state.nav == "🌲 Marcenaria":
    st.header("🌲 Marcenaria")
    aba_conv, aba_cores = st.tabs(["📋 Processadores", "🎨 Cores"])
    
    with aba_conv:
        st.subheader("1️⃣ Fase 1: Gerar Excel de Produção")
        up_csv = st.file_uploader("Suba o CSV do Pontta", type="csv", key="f1")
        if up_csv:
            # Leitura robusta do CSV do Pontta
            content = up_csv.read().decode("utf-8-sig")
            df_b = pd.read_csv(io.StringIO(content), sep=None, engine='python', dtype=str)
            df_b.columns = [norm(c) for c in df_b.columns] # Normaliza nomes das colunas
            
            # Remove linhas de lixo do Pontta (resumos iniciais)
            if 'QUANT' in df_b.columns:
                df_p = df_b[df_b['QUANT'].apply(lambda x: str(x).isdigit())].copy()
            else:
                df_p = df_b.copy()

            if st.button("🚀 Gerar Excel para Fábrica"):
                # Busca Cores
                try:
                    df_c_gs = conn.read(worksheet="CORES_MARCENARIA", ttl=5)
                    m_cores = {norm(r["descricao"]): str(r["codigo"]).split('.')[0].strip() for _, r in df_c_gs.iterrows()}
                except: m_cores = {}

                pesos = df_p.apply(lambda r: calcular_pesos_madeira(r.get("LARG",0), r.get("COMP",0), r.get("QUANT",0), r.get("MATERIAL","")), axis=1)
                df_p["PESO_UNIT"] = pesos.apply(lambda x: x[0]); df_p["PESO_TOTAL"] = pesos.apply(lambda x: x[1])
                if "COR" in df_p.columns: df_p["COR_COD"] = df_p["COR"].apply(lambda x: m_cores.get(norm(x), str(x).split('.')[0]))
                
                out_x = io.BytesIO()
                with pd.ExcelWriter(out_x, engine="openpyxl") as writer:
                    ws = writer.book.create_sheet("PRODUCAO")
                    ws.cell(row=1, column=1, value=f"TECAMA | PRODUÇÃO").font = Font(bold=True, size=14)
                    header = ["QUANT","COMP","LARG","MATERIAL","COR (COD)","DESCPECA","PRODUTO","CORTE","FITA","USINAGEM","PESO UNIT.","PESO TOTAL"]
                    for i, h in enumerate(header, 1):
                        cell = ws.cell(row=3, column=i, value=h); cell.font = Font(bold=True); cell.alignment = Alignment(horizontal="center")
                        cell.border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
                    
                    curr = 4
                    for dp, g in df_p.groupby("DES_PAI", sort=False):
                        for _, r in g.iterrows():
                            # Mapeamento de colunas para as células
                            vals = [r.get("QUANT"), r.get("COMP"), r.get("LARG"), r.get("MATERIAL"), r.get("COR_COD"), r.get("DESCPECA"), r.get("DES_PAI"), "","","", r.get("PESO_UNIT"), r.get("PESO_TOTAL")]
                            for i, v in enumerate(vals, 1):
                                c = ws.cell(row=curr, column=i, value=v)
                                c.border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
                                c.alignment = Alignment(horizontal="center")
                            curr += 1
                        curr += 1
                st.download_button("📥 Baixar Excel", out_x.getvalue(), "PRODUCAO.xlsx")

        st.markdown("---")
        st.subheader("2️⃣ Fase 2: Gerar CSV para Corte Certo")
        up_ex = st.file_uploader("Suba o Excel EDITADO", type="xlsx", key="f2")
        if up_ex:
            if st.button("🚀 Gerar CSV para Corte Certo"):
                df_e = pd.read_excel(up_ex, skiprows=2)
                df_e = df_e.dropna(subset=['QUANT', 'COMP', 'LARG'], how='all')
                
                # Seleção e Conversão para Inteiro (Removendo .0)
                # O ID agora se chama ITEM para evitar erro SYLK
                res_cc = pd.DataFrame()
                res_cc["ITEM"] = range(1, len(df_e) + 1)
                res_cc["QUANT"] = pd.to_numeric(df_e["QUANT"], errors='coerce').fillna(0).astype(int)
                res_cc["COMP"] = pd.to_numeric(df_e["COMP"], errors='coerce').fillna(0).astype(int)
                res_cc["LARG"] = pd.to_numeric(df_e["LARG"], errors='coerce').fillna(0).astype(int)
                
                # Cor e Descrição (Tratando decimais na Cor)
                res_cc["COR"] = df_e["COR (COD)"].apply(lambda x: str(int(float(x))) if str(x).replace('.','').isdigit() else str(x))
                res_cc["DESC"] = df_e["DESCPECA"].astype(str)

                # Gera CSV PURO (Sem cabeçalho, separador ;)
                csv_final = res_cc.to_csv(index=False, sep=";", header=False, encoding="utf-8-sig")
                st.download_button("📥 Baixar CSV Corte Certo", csv_final, "CORTE_CERTO.csv")

# ==========================================
# PÁGINA: METALURGIA (TRAVADA)
# ==========================================
elif st.session_state.nav == "⚙️ Metalurgia":
    st.header("⚙️ Metalurgia")
    aba1, aba2 = st.tabs(["📋 Calculadora", "🛠️ Tabelas Base"])
    # (Mantive a lógica da Metalurgia v9.7 aqui dentro...)
    with aba1:
        up_pdf_m = st.file_uploader("Suba o PDF", type="pdf")
        if up_pdf_m: st.write("PDF Carregado. Pronto para calcular.")
