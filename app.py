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

# --- 2. CSS PARA VISUAL ---
st.markdown("""
    <style>
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] label { font-size: 22px !important; font-weight: 600 !important; color: #333 !important; }
    h1 { color: #FF5722 !important; font-family: 'Segoe UI', sans-serif; }
    .stButton > button {
        background-color: #FF5722; color: white; width: 100%; border-radius: 12px;
        font-weight: bold; height: 3.5em; font-size: 16px; border: none;
    }
    </style>
    """, unsafe_allow_html=True)

conn = st.connection("gsheets", type=GSheetsConnection)

# --- 3. FUNÇÕES AUXILIARES ---
def norm(t):
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
    st.caption("Tecama Hub Industrial v8.0")

# ==========================================
# PÁGINA: METALURGIA
# ==========================================
if st.session_state.nav == "⚙️ Metalurgia":
    st.header("⚙️ Metalurgia")
    aba_calc, aba_db = st.tabs(["📋 Calculadora PDF (Pontta)", "🛠️ Gerenciar Tabelas Base"])
    
    # Carregamento robusto das tabelas do Sheets
    try:
        db_map = conn.read(worksheet="MAPEAMENTO_TIPO", ttl=10)
        db_metro = conn.read(worksheet="PESO_POR_METRO", ttl=10)
        db_conj = conn.read(worksheet="PESO_CONJUNTO", ttl=10)
        
        # Dicionários para busca rápida
        dict_metro = dict(zip(db_metro['secao'].apply(norm), db_metro['peso_kg_m']))
        list_map = db_map.to_dict('records')
        list_conj = db_conj.to_dict('records')
    except Exception as e:
        st.error(f"Erro ao carregar tabelas do Google Sheets: {e}")

    with aba_calc:
        up_pdf = st.file_uploader("Suba o PDF do Pontta", type="pdf")
        if up_pdf:
            itens_extraidos = []
            with pdfplumber.open(up_pdf) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        for r in table:
                            # Filtra linhas que começam com número (Quantidade)
                            if r and len(r) > 3 and str(r[0]).strip().replace('.','').isdigit():
                                itens_extraidos.append({
                                    "QTD": r[0], 
                                    "DESCRIÇÃO": r[1], 
                                    "MEDIDA": r[3], 
                                    "COR": r[2]
                                })
            
            # Editor de dados para conferência manual antes do cálculo
            df_editor = st.data_editor(pd.DataFrame(itens_extraidos), num_rows="dynamic", use_container_width=True)
            
            if st.button("🚀 CALCULAR E GERAR EXCEL"):
                res_final = []
                for _, r in df_editor.iterrows():
                    desc_bruta = str(r.get('DESCRIÇÃO', ''))
                    desc_limpa = norm(desc_bruta)
                    qtd = float(str(r.get('QTD', 0)).replace(',','.')) if r.get('QTD') else 0.0
                    
                    # 1. Identifica o TIPO via Mapeamento
                    tipo_encontrado = "DESCONHECIDO"
                    for regra in list_map:
                        txt_regra = norm(regra.get('texto_contido', ''))
                        if txt_regra and txt_regra in desc_limpa:
                            tipo_encontrado = str(regra.get('tipo', 'DESCONHECIDO')).upper()
                            break
                    
                    if tipo_encontrado == "IGNORAR": continue

                    # 2. Lógica de Cálculo de Peso
                    p_unit = 0.0
                    
                    # Se for CONJUNTO
                    if tipo_encontrado == "CONJUNTO":
                        for c in list_conj:
                            nome_conj = norm(c.get('nome_conjunto', ''))
                            if nome_conj and nome_conj in desc_limpa:
                                p_unit = float(c.get('peso_unit_kg', 0))
                                break
                    
                    # Se for TUBO (ou qualquer outro que dependa da medida/metro)
                    elif "TUBO" in tipo_encontrado or tipo_encontrado in dict_metro:
                        medida_val = 0.0
                        try:
                            # Limpa a medida (ex: "1200 mm" -> 1200.0)
                            m_raw = str(r.get('MEDIDA', '0')).lower().replace('mm','').replace(',','.').strip()
                            medida_val = float(m_raw)
                        except: medida_val = 0.0
                        
                        # Tenta achar o peso/metro pela seção (ex: 20X20)
                        secao_key = norm(tipo_encontrado.replace('TUBO ', '').strip())
                        peso_m = dict_metro.get(secao_key, 0.0)
                        p_unit = (medida_val / 1000) * peso_m

                    res_final.append({
                        "QTD": qtd,
                        "DESCRIÇÃO": desc_bruta,
                        "MEDIDA": r.get('MEDIDA', ''),
                        "TIPO": tipo_encontrado,
                        "PESO UNIT (kg)": round(p_unit, 3),
                        "PESO TOTAL (kg)": round(p_unit * qtd, 3)
                    })
                
                df_res = pd.DataFrame(res_final)
                st.metric("PESO TOTAL DO PEDIDO", f"{df_res['PESO TOTAL (kg)'].sum():.2f} kg")
                st.dataframe(df_res, use_container_width=True)

                # --- GERAR EXCEL PARA DOWNLOAD ---
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df_res.to_excel(writer, index=False, sheet_name="METALURGIA", startrow=1)
                    ws = writer.sheets["METALURGIA"]
                    
                    # Cabeçalho Personalizado
                    ws.cell(row=1, column=1, value=f"RELATÓRIO DE PESOS - METALURGIA").font = Font(bold=True)
                    
                    # Ajuste de Colunas e bordas
                    for i in range(1, 7):
                        ws.column_dimensions[get_column_letter(i)].width = 25
                    
                    # Linha de Total no final
                    last_row = len(df_res) + 3
                    ws.cell(row=last_row, column=5, value="TOTAL GERAL:").font = Font(bold=True)
                    ws.cell(row=last_row, column=6, value=f"{df_res['PESO TOTAL (kg)'].sum():.2f} kg").font = Font(bold=True)
                
                st.download_button(
                    label="📥 BAIXAR PLANILHA CALCULADA",
                    data=output.getvalue(),
                    file_name=f"PESOS_METAL_{up_pdf.name.replace('.pdf', '')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    with aba_db:
        st.info("Para atualizar os pesos, edite as tabelas no Google Sheets e aguarde alguns segundos.")
        # Seletor de tabelas para edição rápida se necessário
        tab_selecionada = st.selectbox("Escolha a tabela para visualizar/editar:", ["MAPEAMENTO_TIPO", "PESO_POR_METRO", "PESO_CONJUNTO"])
        df_view = conn.read(worksheet=tab_selecionada, ttl=0)
        novo_df = st.data_editor(df_view, num_rows="dynamic", use_container_width=True)
        if st.button(f"💾 Salvar Alterações em {tab_selecionada}"):
            conn.update(worksheet=tab_selecionada, data=novo_df)
            st.success("Tabela atualizada com sucesso!")

# (Mantive a lógica da Marcenaria e Início ocultas aqui para o código não ficar gigante, mas elas devem permanecer no seu arquivo principal)
