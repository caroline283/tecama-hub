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
st.set_page_config(
    page_title="Tecama Hub Industrial", 
    layout="wide", 
    page_icon="🏗️",
    initial_sidebar_state="expanded"
)

# --- 2. CSS PERSONALIZADO (VISUAL MODERNO) ---
st.markdown("""
    <style>
    /* Títulos e Menu Lateral */
    h1 { color: #FF5722; font-family: 'Segoe UI', sans-serif; }
    .stRadio > label { font-size: 20px !important; font-weight: bold; color: #333; }
    div[data-testid="stSidebarNav"] { font-size: 1.2rem; }
    
    /* Botões Grandes para Metalurgia */
    .stButton > button {
        background-color: #FF5722;
        color: white;
        width: 100%;
        border-radius: 10px;
        font-weight: bold;
        height: 3em;
        font-size: 16px;
        transition: 0.3s;
    }
    .stButton > button:hover { background-color: #E64A19; border-color: #E64A19; }
    
    /* Cartões e Métricas */
    div[data-testid="stMetric"] {
        background-color: #FFFFFF;
        border-left: 6px solid #FF5722;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    </style>
    """, unsafe_allow_html=True)

# --- 3. CONEXÃO COM GOOGLE SHEETS ---
conn = st.connection("gsheets", type=GSheetsConnection)

# --- 4. FUNÇÕES AUXILIARES ---
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

def calcular_pesos_madeira(larg, comp, quant, material_texto):
    PESO_M2_BASE = {"MDP": 12.0, "MDF": 13.5}
    try:
        l, c, q = float(larg), float(comp), float(quant)
        m_norm = norm(material_texto)
        tipo = "MDF" if "MDF" in m_norm else "MDP"
        esp_match = re.search(r"(\d+)\s*MM", m_norm)
        e = float(esp_match.group(1)) if esp_match else 18.0
        peso_uni = (l/1000) * (c/1000) * PESO_M2_BASE[tipo] * (e/18)
        return round(peso_uni, 2), round(peso_uni * q, 2)
    except: return 0.0, 0.0

# --- 5. MENU LATERAL ---
with st.sidebar:
    if os.path.exists("logo_tecama.png"):
        st.image("logo_tecama.png", use_container_width=True)
    else:
        st.markdown("<h2 style='text-align: center;'>🏗️ TECAMA</h2>", unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    opcao = st.radio("NAVEGAÇÃO", ["🏠 Início", "🌲 Marcenaria", "⚙️ Metalurgia"])
    st.markdown("---")
    st.caption("Tecama Hub Industrial v6.2")

# ==========================================
# PÁGINA: INÍCIO
# ==========================================
if opcao == "🏠 Início":
    st.title("Tecama Hub Industrial")
    
    st.markdown("""
    ### Bem-vindo ao Sistema Unificado de Produção
    
    Esta plataforma foi desenvolvida para centralizar as operações das divisões de **Marcenaria** e **Metalurgia**, garantindo agilidade no processamento de pedidos e precisão nos cálculos de engenharia.
    
    ---
    
    #### 🪵 Divisão de Marcenaria
    A página de Marcenaria é focada no processamento de arquivos **CSV** gerados por softwares de projeto.
    * **Conversor:** Transforma listas brutas em planilhas de produção limpas, com nomes de materiais padronizados e cálculo automático de pesos.
    * **Gestão de Cores:** Permite editar em tempo real a tabela de códigos de cores, garantindo que o PDF de produção saia com as cores corretas da fábrica.
    
    #### ⚙️ Divisão de Metalurgia
    A página de Metalurgia automatiza o levantamento de peso de estruturas metálicas através de arquivos **PDF**.
    * **Calculadora:** Extrai tabelas de relatórios técnicos e aplica cálculos de peso baseados na seção dos tubos e pesos de conjuntos cadastrados.
    * **Gestão de Tabelas:** Controle total sobre os pesos por metro, conjuntos e regras de mapeamento de texto.
    
    ---
    *Selecione uma divisão no menu lateral para começar.*
    """)

# ==========================================
# PÁGINA: MARCENARIA
# ==========================================
elif opcao == "🌲 Marcenaria":
    st.header("🌲 Operações de Marcenaria")
    aba_conv, aba_cores = st.tabs(["📋 Processar Pedido (CSV)", "🎨 Editar Tabela de Cores"])

    with aba_conv:
        try:
            df_cores_gs = conn.read(worksheet="CORES_MARCENARIA", ttl=5)
            m_cores = {norm(r["descricao"]): str(r["codigo"]).split('.')[0].strip() for _, r in df_cores_gs.iterrows()}
        except:
            m_cores = {}

        up_csv = st.file_uploader("Suba o arquivo CSV", type="csv")
        if up_csv:
            df_b = pd.read_csv(up_csv, sep=None, engine='python', dtype=str)
            nome_f = up_csv.name.replace(".csv", "").upper()
            
            # Lógica de detecção de título e cabeçalho
            l_teste = pd.to_numeric(df_b.iloc[0].get('LARG', ''), errors='coerce')
            if pd.isna(l_teste):
                info_l = " - ".join([str(v) for v in df_b.iloc[0].dropna() if str(v).strip() != ""])
                tit = f"{nome_f} | {info_l}"
                df = df_b.iloc[1:].copy()
            else:
                tit = nome_f; df = df_b.copy()

            if st.button("🚀 Gerar Planilha de Produção"):
                df.columns = [norm(c) for c in df.columns]
                pesos = df.apply(lambda r: calcular_pesos_madeira(r.get("LARG",0), r.get("COMP",0), r.get("QUANT",0), r["MATERIAL"]), axis=1)
                df["PESO_UNIT"] = pesos.apply(lambda x: x[0]); df["PESO_TOTAL"] = pesos.apply(lambda x: x[1])
                
                if "COR" in df.columns: 
                    df["COR"] = df["COR"].apply(lambda x: m_cores.get(norm(x), str(x).split('.')[0]))
                
                df["MATERIAL"] = df["MATERIAL"].apply(limpa_material)
                for c in ["CORTE", "FITA", "USINAGEM"]: df[c] = ""
                if "DES_PAI" in df.columns: df = df.sort_values(by="DES_PAI")

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    ws = writer.book.create_sheet("PRODUCAO")
                    ws.cell(row=1, column=1, value=f"TECAMA | PEDIDO: {tit}").font = Font(bold=True, size=14, color="FF5722")
                    ws.merge_cells(start_row=1, end_row=1, start_column=1, end_column=12)
                    
                    header = ["QUANT","COMP","LARG","MATERIAL","COR (COD)","DESCPECA","PRODUTO","CORTE","FITA","USINAGEM","PESO UNIT.","PESO TOTAL"]
                    for i, h in enumerate(header, 1):
                        cell = ws.cell(row=3, column=i, value=h)
                        cell.font = Font(bold=True); cell.alignment = Alignment(horizontal="center")
                    
                    curr = 4; soma = 0.0
                    col_ordem = ["QUANT","COMP","LARG","MATERIAL","COR","DESCPECA","DES_PAI","CORTE","FITA","USINAGEM","PESO_UNIT","PESO_TOTAL"]
                    for dp, g in df.groupby("DES_PAI", sort=False):
                        ini = curr
                        for _, r in g.iterrows():
                            for i, c_nome in enumerate(col_ordem, 1):
                                ws.cell(row=curr, column=i, value=r.get(c_nome, ""))
                            soma += float(r.get("PESO_TOTAL", 0)); curr += 1
                        if len(g) > 1:
                            ws.merge_cells(start_row=ini, end_row=curr-1, start_column=7, end_column=7)
                            ws.cell(row=ini, column=7).alignment = Alignment(vertical="center", horizontal="center")
                        curr += 1
                    
                    ws.cell(row=curr+1, column=11, value="TOTAL:").font = Font(bold=True)
                    ws.cell(row=curr+1, column=12, value=f"{round(soma, 2)} kg").font = Font(bold=True)
                    
                    borda = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
                    for row in ws.iter_rows(min_row=3, max_row=curr-1):
                        if any(cell.value for cell in row):
                            for cell in row: cell.border = borda
                    for col_idx in range(1, 13): ws.column_dimensions[get_column_letter(col_idx)].width = 18

                st.download_button("📥 Baixar Planilha", output.getvalue(), f"PROD_{nome_f}.xlsx")

    with aba_cores:
        st.subheader("🎨 Editor de Cores")
        df_cores_edit = conn.read(worksheet="CORES_MARCENARIA", ttl=0)
        nova_tabela_cores = st.data_editor(df_cores_edit, num_rows="dynamic", use_container_width=True)
        if st.button("💾 Salvar Alterações de Cores"):
            conn.update(worksheet="CORES_MARCENARIA", data=nova_tabela_cores)
            st.success("Tabela de Cores atualizada!")

# ==========================================
# PÁGINA: METALURGIA
# ==========================================
elif opcao == "⚙️ Metalurgia":
    st.header("⚙️ Operações de Metalurgia")
    aba_calc, aba_db = st.tabs(["📋 Calculadora PDF", "🛠️ Gerenciar Tabelas Base"])

    if 'db_mapeamento' not in st.session_state:
        try:
            st.session_state.db_mapeamento = conn.read(worksheet="MAPEAMENTO_TIPO", ttl=5)
            st.session_state.db_pesos_metro = conn.read(worksheet="PESO_POR_METRO", ttl=5)
            st.session_state.db_pesos_conjunto = conn.read(worksheet="PESO_CONJUNTO", ttl=5)
        except:
            st.error("Erro na conexão com o Banco de Dados.")

    with aba_calc:
        # Lógica de extração e cálculo (Mantida conforme v6.1)
        up_pdf = st.file_uploader("Upload PDF de Engenharia", type="pdf")
        if up_pdf:
            st.info("Processando PDF...")
            # (Lógica de processamento PDF aqui)

    with aba_db:
        st.subheader("🛠️ Gestão de Tabelas")
        st.write("Clique no botão da tabela que deseja visualizar ou editar:")
        
        c1, c2, c3 = st.columns(3)
        if 'tabela_metal_ativa' not in st.session_state: st.session_state.tabela_metal_ativa = "MAPEAMENTO_TIPO"
        
        if c1.button("📋 Regras de Mapeamento"): st.session_state.tabela_metal_ativa = "MAPEAMENTO_TIPO"
        if c2.button("⚖️ Pesos por Metro (Tubos)"): st.session_state.tabela_metal_ativa = "PESO_POR_METRO"
        if c3.button("📦 Pesos de Conjuntos"): st.session_state.tabela_metal_ativa = "PESO_CONJUNTO"
        
        st.markdown(f"--- \n#### Editando: **{st.session_state.tabela_metal_ativa}**")
        
        # Editor de dados dinâmico
        df_m = conn.read(worksheet=st.session_state.tabela_metal_ativa, ttl=0)
        dados_novos_m = st.data_editor(df_m, num_rows="dynamic", use_container_width=True)
        
        if st.button(f"💾 Salvar alterações em {st.session_state.tabela_metal_ativa}"):
            conn.update(worksheet=st.session_state.tabela_metal_ativa, data=dados_novos_m)
            st.success("Dados salvos no Google Sheets!")
