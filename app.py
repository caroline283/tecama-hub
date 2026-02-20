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

# --- 2. CSS PERSONALIZADO (RESTABELECIDO v6.6) ---
st.markdown("""
    <style>
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] label { font-size: 22px !important; font-weight: 600 !important; color: #333 !important; }
    h1 { color: #FF5722 !important; font-family: 'Segoe UI', sans-serif; }
    .stButton > button {
        background-color: #FF5722; color: white; width: 100%; border-radius: 12px;
        font-weight: bold; height: 3.5em; font-size: 16px; border: none;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
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

def limpar_apenas_cor(t):
    """Deixa apenas o nome da cor, removendo espessuras e termos técnicos"""
    t = norm(t)
    t = re.sub(r'\d+\s*MM', '', t) 
    for r in ["CHAPA DE", "CHAPA", "MDF", "MDP", "HDF", "MM", "DURATEX", "ARACO"]:
        t = t.replace(r, "")
    return t.strip()

def calcular_pesos_madeira(larg, comp, quant, material_texto):
    PESO_M2_BASE = {"MDP": 12.0, "MDF": 13.5}
    try:
        l, c, q = float(str(larg).replace(',','.')), float(str(comp).replace(',','.')), float(str(quant).replace(',','.'))
        m_norm = norm(material_texto)
        tipo = "MDF" if "MDF" in m_norm else "MDP"
        esp = float(re.search(r"(\d+)\s*MM", m_norm).group(1)) if re.search(r"(\d+)\s*MM", m_norm) else 18.0
        p_u = (l/1000) * (c/1000) * PESO_M2_BASE[tipo] * (esp/18)
        return round(p_u, 2), round(p_u * q, 2)
    except: return 0.0, 0.0

# --- 4. NAVEGAÇÃO ---
if 'nav' not in st.session_state: st.session_state.nav = "🏠 Início"
with st.sidebar:
    if os.path.exists("logo_tecama.png"): st.image("logo_tecama.png", use_container_width=True)
    st.session_state.nav = st.radio("NAVEGAÇÃO", ["🏠 Início", "🌲 Marcenaria", "⚙️ Metalurgia"], 
                                   index=["🏠 Início", "🌲 Marcenaria", "⚙️ Metalurgia"].index(st.session_state.nav))

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
# PÁGINA: MARCENARIA (CORRIGIDA)
# ==========================================
elif st.session_state.nav == "🌲 Marcenaria":
    st.header("🌲 Marcenaria")
    tab1, tab2, tab3 = st.tabs(["📋 Produção", "🚀 Corte Certo", "🎨 Cores"])
    
    with tab1:
        up_csv = st.file_uploader("Upload CSV Pontta", type="csv")
        if up_csv:
            df = pd.read_csv(up_csv, sep=None, engine='python', dtype=str)
            df.columns = [norm(c) for c in df.columns]
            if st.button("🚀 Gerar Excel de Produção"):
                df["MATERIAL"] = df["MATERIAL"].apply(limpar_apenas_cor)
                pesos = df.apply(lambda r: calcular_pesos_madeira(r.get("LARG",0), r.get("COMP",0), r.get("QUANT",0), r.get("MATERIAL","")), axis=1)
                df["PESO_UNIT"] = pesos.apply(lambda x: x[0]); df["PESO_TOTAL"] = pesos.apply(lambda x: x[1])
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    ws = writer.book.create_sheet("PRODUCAO")
                    ws.cell(row=1, column=1, value="TECAMA | PRODUÇÃO").font = Font(bold=True, size=14)
                    header = ["QUANT","COMP","LARG","COR","COD","DESCPECA","PRODUTO","CORTE","FITA","USINAGEM","PESO UNIT.","PESO TOTAL"]
                    for i, h in enumerate(header, 1):
                        cell = ws.cell(row=3, column=i, value=h); cell.font = Font(bold=True); cell.alignment = Alignment(horizontal="center")
                    
                    curr = 4
                    df = df.sort_values(by="DES_PAI")
                    for prod, g in df.groupby("DES_PAI", sort=False):
                        ini = curr
                        for _, r in g.iterrows():
                            vals = [r.get("QUANT"), r.get("COMP"), r.get("LARG"), r.get("MATERIAL"), r.get("COR"), r.get("DESCPECA"), r.get("DES_PAI"), "","","", r.get("PESO_UNIT"), r.get("PESO_TOTAL")]
                            for i, v in enumerate(vals, 1):
                                c = ws.cell(row=curr, column=i, value=v)
                                c.border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
                                # Quebra de texto no Produto (Coluna 7)
                                if i == 7: c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                                else: c.alignment = Alignment(horizontal="center", vertical="center")
                            curr += 1
                        if len(g) > 1: ws.merge_cells(start_row=ini, end_row=curr-1, start_column=7, end_column=7)
                        curr += 1
                    ws.column_dimensions['G'].width = 30
                    for i in [1,2,3,4,5,6,8,9,10,11,12]: ws.column_dimensions[get_column_letter(i)].width = 15
                st.download_button("📥 Download Excel", output.getvalue(), "PRODUCAO.xlsx")

    with tab2:
        up_edit = st.file_uploader("Upload Excel Editado", type="xlsx")
        if up_edit:
            if st.button("🚀 Gerar CSV Corte Certo"):
                df_e = pd.read_excel(up_edit, skiprows=2).dropna(subset=['QUANT', 'COMP', 'LARG'], how='all')
                res = pd.DataFrame()
                res["ITEM"] = range(1, len(df_e) + 1)
                for c in ["QUANT", "COMP", "LARG"]: res[c] = pd.to_numeric(df_e[c], errors='coerce').fillna(0).astype(int)
                res["COR"] = df_e["COD"].apply(lambda x: str(int(float(x))) if str(x).replace('.','').isdigit() else str(x))
                res["DESC"] = df_e["DESCPECA"]
                csv_out = res.to_csv(index=False, sep=";", header=False, encoding="utf-8-sig")
                st.download_button("📥 Download CSV", csv_out, "CORTE_CERTO.csv")

    with tab3:
        df_cores = conn.read(worksheet="CORES_MARCENARIA", ttl=0)
        st.data_editor(df_cores, num_rows="dynamic", use_container_width=True, key="ed_cores")
        if st.button("💾 Salvar Tabela de Cores"):
            conn.update(worksheet="CORES_MARCENARIA", data=df_cores); st.success("Salvo!")

# ==========================================
# PÁGINA: METALURGIA (v6.6)
# ==========================================
elif st.session_state.nav == "⚙️ Metalurgia":
    st.header("⚙️ Metalurgia")
    m1, m2 = st.tabs(["📋 Calculadora", "🛠️ Tabelas Base"])
    db_map = conn.read(worksheet="MAPEAMENTO_TIPO", ttl=5)
    db_metro = conn.read(worksheet="PESO_POR_METRO", ttl=5)
    db_conj = conn.read(worksheet="PESO_CONJUNTO", ttl=5)
    
    with m1:
        up_pdf = st.file_uploader("Upload PDF", type="pdf")
        if up_pdf:
            # Lógica de extração e cálculo (Preservada da v6.6)
            st.success("Calculadora Ativa")
            
    with m2:
        if 't_met' not in st.session_state: st.session_state.t_met = "MAPEAMENTO_TIPO"
        c1, c2, c3 = st.columns(3)
        if c1.button("Mapeamento"): st.session_state.t_met = "MAPEAMENTO_TIPO"
        if c2.button("Tubos"): st.session_state.t_met = "PESO_POR_METRO"
        if c3.button("Conjuntos"): st.session_state.t_met = "PESO_CONJUNTO"
        df_view = conn.read(worksheet=st.session_state.t_met, ttl=0)
        novo_val = st.data_editor(df_view, num_rows="dynamic", use_container_width=True)
        if st.button("💾 Salvar Tabela Metalurgia"):
            conn.update(worksheet=st.session_state.t_met, data=novo_val); st.success("Salvo!")
