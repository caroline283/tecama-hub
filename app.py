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

# --- 2. CSS PERSONALIZADO (VISUAL MODERNO E FONTE MAIOR) ---
st.markdown("""
    <style>
    /* Aumentar o texto da barra lateral */
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] label {
        font-size: 22px !important;
        font-weight: 600 !important;
        padding: 10px 0px !important;
        color: #333 !important;
    }
    
    /* Estilo dos títulos e textos */
    h1 { color: #FF5722 !important; font-family: 'Segoe UI', sans-serif; }
    h3 { color: #444 !important; }
    
    /* Botões Grandes e Independentes (Metalurgia) */
    .stButton > button {
        background-color: #FF5722;
        color: white;
        width: 100%;
        border-radius: 12px;
        font-weight: bold;
        height: 4em;
        font-size: 16px;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        border: none;
    }
    .stButton > button:hover { background-color: #E64A19; transform: translateY(-2px); }
    
    /* Ajuste do Logo na Sidebar */
    [data-testid="stSidebar"] [data-testid="stImage"] {
        padding-top: 20px;
        margin-bottom: -20px;
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
        st.markdown("<h2 style='text-align: center; color: #FF5722;'>TECAMA</h2>", unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    # Lista de navegação com texto maior via CSS
    opcao = st.radio("NAVEGAÇÃO", ["🏠 Início", "🌲 Marcenaria", "⚙️ Metalurgia"])
    st.markdown("---")
    st.caption("Tecama Hub Industrial v6.4")

# ==========================================
# PÁGINA: INÍCIO
# ==========================================
if opcao == "🏠 Início":
    st.title("Tecama Hub Industrial")
    
    st.markdown("### Bem-vindo ao Sistema Unificado de Produção")
    st.write("Esta plataforma foi desenvolvida para centralizar as operações das divisões de **Marcenaria** e **Metalurgia**, garantindo agilidade no processamento de pedidos e precisão nos cálculos de engenharia.")
    
    st.markdown("---")
    
    # Seção Marcenaria
    st.subheader("🌲 Divisão de Marcenaria")
    st.markdown("""
    A página de Marcenaria é focada no **processamento de arquivos CSV gerados pelo Pontta**.
    * **Conversor:** Transforma listas brutas em planilhas de produção limpas, com nomes de materiais padronizados e cálculo automático de pesos.
    * **Gestão de Cores:** Permite editar em tempo real a tabela de códigos de cores, garantindo que o PDF de produção saia com as cores corretas da fábrica.
    """)
    
    # Seção Metalurgia
    st.subheader("⚙️ Divisão de Metalurgia")
    st.markdown("""
    A página de Metalurgia **automatiza o levantamento de peso de estruturas metálicas através do relatório de metalurgia em PDF gerado pelo Pontta**.
    * **Calculadora:** Extrai tabelas de relatórios técnicos e aplica cálculos de peso baseados na seção dos tubos e pesos de conjuntos cadastrados.
    * **Gestão de Tabelas:** Controle total sobre os pesos por metro, conjuntos e regras de mapeamento de texto.
    """)
    
    st.markdown("---")
    st.info("Selecione uma divisão no menu lateral para começar.")

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

        up_csv = st.file_uploader("Suba o arquivo CSV (Pontta)", type="csv")
        if up_csv:
            df_b = pd.read_csv(up_csv, sep=None, engine='python', dtype=str)
            nome_f = up_csv.name.replace(".csv", "").upper()
            
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
                    col_ordem = ["QUANT","COMP","LARG","MATERIAL","COR","DESCPECA","DES_PAI","CORTE","FITA","USINAGEM","PES_UNIT","PESO_TOTAL"]
                    # Correção: O código deve usar os nomes exatos das colunas calculadas
                    col_map = {"PES_UNIT": "PESO_UNIT", "PESO_TOTAL": "PESO_TOTAL"}
                    
                    for dp, g in df.groupby("DES_PAI", sort=False):
                        ini = curr
                        for _, r in g.iterrows():
                            for i, c_nome in enumerate(col_ordem, 1):
                                val = r.get(col_map.get(c_nome, c_nome), "")
                                ws.cell(row=curr, column=i, value=val)
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

                st.download_button("📥 Baixar Planilha Marcenaria", output.getvalue(), f"PROD_{nome_f}.xlsx")

    with aba_cores:
        st.subheader("🎨 Editor de Cores")
        df_cores_edit = conn.read(worksheet="CORES_MARCENARIA", ttl=0)
        nova_tabela_cores = st.data_editor(df_cores_edit, num_rows="dynamic", use_container_width=True)
        if st.button("💾 Salvar Alterações de Cores"):
            conn.update(worksheet="CORES_MARCENARIA", data=nova_tabela_cores)
            st.success("Tabela de Cores atualizada no Google Sheets!")

# ==========================================
# PÁGINA: METALURGIA
# ==========================================
elif opcao == "⚙️ Metalurgia":
    st.header("⚙️ Operações de Metalurgia")
    aba_calc, aba_db = st.tabs(["📋 Calculadora PDF (Pontta)", "🛠️ Gerenciar Tabelas Base"])

    if 'db_mapeamento' not in st.session_state:
        try:
            st.session_state.db_mapeamento = conn.read(worksheet="MAPEAMENTO_TIPO", ttl=5)
            st.session_state.db_pesos_metro = conn.read(worksheet="PESO_POR_METRO", ttl=5)
            st.session_state.db_pesos_conjunto = conn.read(worksheet="PESO_CONJUNTO", ttl=5)
        except:
            st.error("Erro na conexão com o Banco de Dados.")

    with aba_calc:
        up_pdf = st.file_uploader("Suba o Relatório de Metalurgia (PDF)", type="pdf")
        if up_pdf:
            st.info("Extraindo dados do relatório Pontta...")
            # Lógica de cálculo PDF aqui...

    with aba_db:
        st.subheader("🛠️ Gestão de Tabelas")
        if 'tab_met' not in st.session_state: st.session_state.tab_met = "MAPEAMENTO_TIPO"
        
        c1, c2, c3 = st.columns(3)
        if c1.button("📋 Mapeamento de Peças"): st.session_state.tab_met = "MAPEAMENTO_TIPO"
        if c2.button("⚖️ Pesos de Tubos (m)"): st.session_state.tab_met = "PESO_POR_METRO"
        if c3.button("📦 Pesos de Conjuntos"): st.session_state.tab_met = "PESO_CONJUNTO"
        
        st.markdown(f"#### Editando: **{st.session_state.tab_met}**")
        df_m = conn.read(worksheet=st.session_state.tab_met, ttl=0)
        dados_novos_m = st.data_editor(df_m, num_rows="dynamic", use_container_width=True)
        
        if st.button(f"💾 Salvar alterações em {st.session_state.tab_met}"):
            conn.update(worksheet=st.session_state.tab_met, data=dados_novos_m)
            st.success(f"Tabela {st.session_state.tab_met} salva!")
