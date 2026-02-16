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

# --- 1. CONFIGURAÇÃO DO HUB TECAMA ---
st.set_page_config(page_title="Tecama Hub Industrial", layout="wide", page_icon="🏗️")

# --- 2. CSS PERSONALIZADO ---
st.markdown("""
    <style>
    h1 { color: #FF5722; }
    .stButton>button { background-color: #FF5722; color: white; width: 100%; border-radius: 8px; font-weight: bold; }
    div[data-testid="stMetric"] { background-color: #F8F9FA; border-left: 5px solid #FF5722; padding: 15px; border-radius: 5px; }
    .stTabs [data-baseweb="tab-list"] { gap: 24px; }
    </style>
    """, unsafe_allow_html=True)

# --- 3. CONEXÃO COM GOOGLE SHEETS ---
conn = st.connection("gsheets", type=GSheetsConnection)

# --- 4. FUNÇÕES DE AUXÍLIO ---
def norm(t):
    if not t or pd.isna(t): return ""
    t = unicodedata.normalize("NFD", str(t).upper()).encode("ascii", "ignore").decode("utf-8")
    return " ".join(t.split()).strip()

def limpa_material(t):
    t = norm(t)
    t = re.sub(rf'\d+\s*MM', '', t)
    t = re.sub(rf'\d+', '', t)
    for r in ["CHAPA DE", "CHAPA", "MDF", "MDP", "HDF", "MM"]:
        t = re.sub(rf'\b{r}\b', '', t)
    return t.strip()

# --- 5. MENU LATERAL ---
with st.sidebar:
    st.markdown("<h1 style='text-align: center;'>🏗️ TECAMA</h1>", unsafe_allow_html=True)
    opcao = st.radio("Selecione a Divisão:", ["🏠 Início", "🪵 Marcenaria (CSV)", "⚙️ Metalurgia (PDF)"])
    st.markdown("---")
    st.caption("Tecama Hub v5.0")

# ==========================================
# DIVISÃO 1: MARCENARIA (CONVERSOR CSV)
# ==========================================
if opcao == "🪵 Marcenaria (CSV)":
    st.header("🪵 Divisão de Marcenaria")
    st.subheader("Conversor de Produção CSV")

    # Carregar cores do Google Sheets em vez de CSV local
    try:
        df_cores_gs = conn.read(worksheet="CORES_MARCENARIA", ttl=5)
        m_cores = {norm(r["descricao"]): str(r["codigo"]).strip() 
                   for _, r in df_cores_gs.iterrows() if "descricao" in df_cores_gs.columns}
    except:
        st.warning("⚠️ Não foi possível carregar a aba 'CORES_MARCENARIA' do Google Sheets.")
        m_cores = {}

    up_csv = st.file_uploader("Suba o ficheiro CSV gerado pelo sistema", type="csv")
    
    if up_csv:
        df_b = pd.read_csv(up_csv, sep=None, engine='python', dtype=str)
        nome_f = up_csv.name.replace(".csv", "").upper()
        
        # Identificar título do pedido
        l_teste = pd.to_numeric(df_b.iloc[0].get('LARG', ''), errors='coerce')
        if pd.isna(l_teste):
            info_l = " - ".join([str(v) for v in df_b.iloc[0].dropna() if str(v).strip() != ""])
            tit = f"{nome_f} | {info_l}"
            df = df_b.iloc[1:].copy()
        else:
            tit = nome_f
            df = df_b.copy()

        if st.button("🚀 Gerar Excel de Produção"):
            df.columns = [norm(c) for c in df.columns]
            
            # Aplicar Mapa de Cores e Limpeza
            if "COR" in df.columns:
                df["COR"] = df["COR"].apply(lambda x: m_cores.get(norm(x), x))
            if "MATERIAL" in df.columns:
                df["MATERIAL"] = df["MATERIAL"].apply(limpa_material)
            
            # Colunas Extras para oficina
            for c in ["CORTE", "FITA", "USINAGEM"]: df[c] = ""
            
            # Criar Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="PRODUCAO", startrow=3, index=False)
                ws = writer.sheets["PRODUCAO"]
                ws.cell(row=1, column=1, value=f"TECAMA | PEDIDO: {tit}").font = Font(bold=True, size=14)
                
                # Ajuste automático de colunas
                for col in ws.columns:
                    ws.column_dimensions[col[0].column_letter].width = 20
            
            st.success("✅ Conversão concluída!")
            st.download_button("📥 Baixar Planilha Marcenaria", output.getvalue(), f"PROD_{nome_f}.xlsx")

# ==========================================
# DIVISÃO 2: METALURGIA (PDF)
# ==========================================
elif opcao == "⚙️ Metalurgia (PDF)":
    st.header("⚙️ Metalurgia System 3.0")
    
    # Abas internas da Metalurgia
    aba_calc, aba_db = st.tabs(["📋 Calculadora", "🛠️ Configurações (Nuvem)"])

    with aba_calc:
        uploaded_pdf = st.file_uploader("Suba o Relatório PDF do Pedido Metálico", type="pdf")
        if uploaded_pdf:
            # Aqui entra a tua lógica original de extração de tabelas do PDF
            st.info("Processando extração de dados do PDF...")
            # (Insere aqui as tuas funções: pdfplumber -> extract_tables)

    with aba_db:
        st.subheader("Base de Dados (Google Sheets)")
        if st.button("🔄 Sincronizar Tabelas"):
            st.rerun()
        # Mostra as tabelas de Mapeamento, Pesos Tubos, etc.

elif opcao == "🏠 Início":
    st.title("Bem-vindo ao Tecama Hub Industrial")
    st.markdown("""
    Este é o centro de operações digital da **Tecama**.
    
    - **Marcenaria:** Converte CSVs de projeto em listas de corte limpas com cores codificadas.
    - **Metalurgia:** Extrai dados de PDFs e calcula pesos de estruturas metálicas.
    """)
    st.image("https://via.placeholder.com/800x300?text=TECAMA+INDUSTRIAL+HUB", use_container_width=True)
