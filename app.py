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

def limpar_material_cor(t):
    """Limpa o material deixando apenas o nome da cor"""
    t = norm(t)
    t = re.sub(r'\d+\s*MM', '', t) 
    for r in ["CHAPA DE", "CHAPA", "MDF", "MDP", "HDF", "DURATEX", "ARACO", "DE "]:
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
    opcao = st.radio("NAVEGAÇÃO", ["🏠 Início", "🌲 Marcenaria", "⚙️ Metalurgia"], 
                     index=["🏠 Início", "🌲 Marcenaria", "⚙️ Metalurgia"].index(st.session_state.nav))
    st.session_state.nav = opcao
    st.caption("Tecama Hub Industrial v12.5")

# ==========================================
# PÁGINA: INÍCIO (TEXTO ORIGINAL INTEGRAL)
# ==========================================
if st.session_state.nav == "🏠 Início":
    st.title("Tecama Hub Industrial")
    st.markdown("### Bem-vindo ao Sistema Unificado de Produção")
    st.write("Esta plataforma foi desenvolvida para centralizar as operações das divisões de **Marcenaria** e **Metalurgia**, garantindo agilidade no processamento de pedidos e precisão nos cálculos de engenharia.")
    st.markdown("---")
    st.markdown('<div class="home-link">', unsafe_allow_html=True)
    if st.button("🌲 Divisão de Marcenaria"):
        st.session_state.nav = "🌲 Marcenaria"; st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown("""
    A página de Marcenaria é focada no **processamento de arquivos CSV gerados pelo Pontta**.
    * **Conversor:** Transforma listas brutas em planilhas de produção limpas, com nomes de materiais padronizados e cálculo automático de pesos.
    * **Gestão de Cores:** Permite editar em tempo real a tabela de códigos de cores, garantindo que o PDF de produção saia com as cores corretas da fábrica.
    """)
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="home-link">', unsafe_allow_html=True)
    if st.button("⚙️ Divisão de Metalurgia"):
        st.session_state.nav = "⚙️ Metalurgia"; st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown("""
    A página de Metalurgia **automatiza o levantamento de peso de estruturas metálicas através do relatório de metalurgia em PDF gerado pelo Pontta**.
    * **Calculadora:** Extrai tabelas de relatórios técnicos e aplica cálculos de peso baseados na seção dos tubos e pesos de conjuntos cadastrados.
    * **Gestão de Tabelas:** Controle total sobre os pesos por metro, conjuntos e regras de mapeamento de texto.
    """)

# ==========================================
# PÁGINA: MARCENARIA
# ==========================================
elif st.session_state.nav == "🌲 Marcenaria":
    st.header("🌲 Operações de Marcenaria")
    aba_conv, aba_cores = st.tabs(["📋 Processadores de Arquivos", "🎨 Editar Cores"])
    
    with aba_conv:
        st.subheader("1️⃣ Fase 1: Gerar Excel de Produção")
        up_csv = st.file_uploader("Suba o CSV do Pontta", type="csv")
        if up_csv:
            df_b = pd.read_csv(up_csv, sep=None, engine='python', dtype=str)
            df_b.columns = [norm(c) for c in df_b.columns]
            if st.button("🚀 Gerar Excel para Fábrica"):
                # Limpa material e calcula pesos
                df_b["MATERIAL_LIMPO"] = df_b["MATERIAL"].apply(limpar_material_cor)
                pesos = df_b.apply(lambda r: calcular_pesos_madeira(r.get("LARG",0), r.get("COMP",0), r.get("QUANT",0), r.get("MATERIAL","")), axis=1)
                df_b["PES_U"] = pesos.apply(lambda x: x[0]); df_b["PES_T"] = pesos.apply(lambda x: x[1])
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    ws = writer.book.create_sheet("PRODUCAO")
                    ws.cell(row=1, column=1, value="TECAMA | PRODUÇÃO").font = Font(bold=True, size=14)
                    header = ["QUANT","COMP","LARG","MATERIAL","COR (COD)","DESCPECA","PRODUTO","CORTE","FITA","USINAGEM","PESO UNIT.","PESO TOTAL"]
                    for i, h in enumerate(header, 1):
                        cell = ws.cell(row=3, column=i, value=h); cell.font = Font(bold=True); cell.alignment = Alignment(horizontal="center")
                    
                    curr = 4
                    df_b = df_b.sort_values(by="DES_PAI")
                    borda = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
                    
                    for prod, g in df_b.groupby("DES_PAI", sort=False):
                        ini = curr
                        for _, r in g.iterrows():
                            vals = [r.get("QUANT"), r.get("COMP"), r.get("LARG"), r.get("MATERIAL_LIMPO"), r.get("COR"), r.get("DESCPECA"), r.get("DES_PAI"), "","","", r.get("PES_U"), r.get("PES_T")]
                            for i, v in enumerate(vals, 1):
                                c = ws.cell(row=curr, column=i, value=v)
                                c.border = borda
                                if i == 7: c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                                else: c.alignment = Alignment(horizontal="center", vertical="center")
                            curr += 1
                        if len(g) > 1: ws.merge_cells(start_row=ini, end_row=curr-1, start_column=7, end_column=7)
                        curr += 1
                    ws.column_dimensions['G'].width = 25
                st.download_button("📥 Baixar Excel", output.getvalue(), "PRODUCAO.xlsx")

        st.markdown("---")
        st.subheader("2️⃣ Fase 2: Gerar CSV para Corte Certo")
        up_ex = st.file_uploader("Suba o Excel Editado", type="xlsx")
        if up_ex:
            if st.button("🚀 Gerar CSV para Corte Certo"):
                df_e = pd.read_excel(up_ex, skiprows=2).dropna(subset=['QUANT', 'COMP', 'LARG'], how='all')
                res = pd.DataFrame()
                res["ITEM"] = range(1, len(df_e) + 1)
                for c in ["QUANT", "COMP", "LARG"]: res[c] = pd.to_numeric(df_e[c], errors='coerce').fillna(0).astype(int)
                res["COR"] = df_e["COR (COD)"].fillna("0").astype(str).apply(lambda x: x.split('.')[0])
                res["DESC"] = df_e["DESCPECA"]
                csv_out = res.to_csv(index=False, sep=";", header=False, encoding="utf-8-sig")
                st.download_button("📥 Baixar CSV Corte Certo", csv_out, "CORTE_CERTO.csv")

    with aba_cores:
        df_c = conn.read(worksheet="CORES_MARCENARIA", ttl=0)
        novo_c = st.data_editor(df_c, num_rows="dynamic", use_container_width=True)
        if st.button("💾 Salvar Cores"):
            conn.update(worksheet="CORES_MARCENARIA", data=novo_c); st.success("Salvo!")

# ==========================================
# PÁGINA: METALURGIA (ORIGINAL INTEGRAL)
# ==========================================
elif st.session_state.nav == "⚙️ Metalurgia":
    st.header("⚙️ Metalurgia")
    aba_calc, aba_db = st.tabs(["📋 Calculadora PDF", "🛠️ Tabelas Base"])
    
    db_map = conn.read(worksheet="MAPEAMENTO_TIPO", ttl=5)
    db_metro = conn.read(worksheet="PESO_POR_METRO", ttl=5)
    db_conj = conn.read(worksheet="PESO_CONJUNTO", ttl=5)
    
    with aba_calc:
        up_pdf = st.file_uploader("PDF Pontta", type="pdf")
        if up_pdf: st.success("PDF pronto")
            
    with aba_db:
        if 't_ativa' not in st.session_state: st.session_state.t_ativa = "MAPEAMENTO_TIPO"
        c1, c2, c3 = st.columns(3)
        if c1.button("📋 Mapeamento"): st.session_state.t_ativa = "MAPEAMENTO_TIPO"
        if c2.button("⚖️ Tubos"): st.session_state.t_ativa = "PESO_POR_METRO"
        if c3.button("📦 Conjuntos"): st.session_state.t_ativa = "PESO_CONJUNTO"
        df_v = conn.read(worksheet=st.session_state.t_ativa, ttl=0)
        st.data_editor(df_v, num_rows="dynamic", use_container_width=True)
