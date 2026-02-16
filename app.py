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

# --- CONFIGURAÇÃO DO HUB TECAMA ---
st.set_page_config(page_title="Tecama Hub Industrial", layout="wide", page_icon="🏗️")

# --- CSS ESTILIZADO ---
st.markdown("""
    <style>
    h1 { color: #FF5722; }
    .stButton>button { background-color: #FF5722; color: white; width: 100%; border-radius: 8px; }
    div[data-testid="stMetric"] { background-color: #F0F2F6; border-left: 5px solid #FF5722; padding: 15px; border-radius: 5px; }
    </style>
    """, unsafe_allow_html=True)

# --- MENU LATERAL ---
with st.sidebar:
    st.title("🏗️ TECAMA HUB")
    opcao = st.radio("Selecione a Divisão:", ["🏠 Início", "🪵 Marcenaria (CSV)", "⚙️ Metalurgia (PDF)"])
    st.markdown("---")
    st.caption("Versão Hub 5.0")

# ==========================================
# DIVISÃO 1: MARCENARIA (CONVERSOR CSV)
# ==========================================
if opcao == "🪵 Marcenaria (CSV)":
    st.header("🪵 Conversor de Produção - Marcenaria")
    
    def norm(t):
        if not t or pd.isna(t): return ""
        t = unicodedata.normalize("NFD", str(t).upper()).encode("ascii", "ignore").decode("utf-8")
        return " ".join(t.split()).strip()

    def limpa_mat(t):
        t = norm(t)
        t = re.sub(rf'\d+\s*MM', '', t); t = re.sub(rf'\d+', '', t)
        for r in ["CHAPA DE", "CHAPA", "MDF", "MDP", "HDF", "MM"]: t = re.sub(rf'\b{r}\b', '', t)
        return t.strip()

    try:
        df_cores = pd.read_csv("cores.csv")
        m_cores = {norm(r["descricao"]): str(r["codigo"]).strip() for _, r in df_cores.iterrows()}
    except:
        st.error("Atenção: cores.csv não encontrado no repositório.")
        m_cores = {}

    up_csv = st.file_uploader("Suba o arquivo CSV", type="csv")
    if up_csv:
        df_b = pd.read_csv(up_csv, sep=None, engine='python', dtype=str)
        nome_f = up_csv.name.replace(".csv", "").upper()
        
        # Título Inteligente
        l_teste = pd.to_numeric(df_b.iloc[0].get('LARG', ''), errors='coerce')
        if pd.isna(l_teste):
            info_l = " - ".join([str(v) for v in df_b.iloc[0].dropna() if str(v).strip() != ""])
            tit = f"{nome_f} | {info_l}"
            df = df_b.iloc[1:].copy()
        else:
            tit = nome_f; df = df_b.copy()

        if st.button("🚀 Processar Excel da Marcenaria"):
            df.columns = [norm(c) for c in df.columns]
            if "COR" in df.columns: df["COR"] = df["COR"].apply(lambda x: m_cores.get(norm(x), x))
            df["MATERIAL"] = df["MATERIAL"].apply(limpa_mat)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as w:
                ws = w.book.create_sheet("PRODUCAO")
                ws.cell(row=1, column=1, value=f"TECAMA | CLIENTE: {tit}").font = Font(bold=True, size=12)
                
                # Ajuste de colunas e bordas automático aqui...
                df.to_excel(w, sheet_name="PRODUCAO", startrow=2, index=False)
            
            st.success("Excel Gerado!")
            st.download_button("📥 Baixar Arquivo", output.getvalue(), f"Tecama_{nome_f}.xlsx")

# ==========================================
# DIVISÃO 2: METALURGIA (PDF)
# ==========================================
elif opcao == "⚙️ Metalurgia (PDF)":
    st.header("⚙️ Metalurgia System")
    # --- COLOQUE AQUI O RESTANTE DO SEU CÓDIGO DE METALURGIA ---
    st.info("Aqui entrará o código que usa o pdfplumber e GSheets.")

elif opcao == "🏠 Início":
    st.title("Bem-vindo ao Tecama Hub Industrial")
    st.write("Selecione uma ferramenta na barra lateral para começar.")
