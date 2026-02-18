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

# --- 2. CSS PERSONALIZADO (VISUAL v6.6) ---
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

# --- 4. NAVEGAÇÃO ---
if 'nav' not in st.session_state: st.session_state.nav = "🏠 Início"

with st.sidebar:
    if os.path.exists("logo_tecama.png"): st.image("logo_tecama.png", use_container_width=True)
    opcao = st.radio("NAVEGAÇÃO", ["🏠 Início", "🌲 Marcenaria", "⚙️ Metalurgia"], 
                     index=["🏠 Início", "🌲 Marcenaria", "⚙️ Metalurgia"].index(st.session_state.nav))
    st.session_state.nav = opcao
    st.caption("Tecama Hub Industrial v9.3")

# ==========================================
# PÁGINA: INÍCIO (TEXTO v6.6)
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
        # FASE 1: PRODUÇÃO
        st.subheader("1️⃣ Fase 1: Gerar Excel de Produção")
        try:
            df_cores_gs = conn.read(worksheet="CORES_MARCENARIA", ttl=5)
            m_cores = {norm(r["descricao"]): str(r["codigo"]).split('.')[0].strip() for _, r in df_cores_gs.iterrows()}
        except: m_cores = {}
        
        up_csv_fase1 = st.file_uploader("Suba o CSV original do Pontta", type="csv", key="f1")
        if up_csv_fase1:
            df_b = pd.read_csv(up_csv_fase1, sep=None, engine='python', dtype=str)
            nome_f = up_csv_fase1.name.replace(".csv", "").upper()
            l_teste = pd.to_numeric(df_b.iloc[0].get('LARG', ''), errors='coerce')
            df = df_b.iloc[1:].copy() if pd.isna(l_teste) else df_b.copy()

            if st.button("🚀 Gerar Excel para Fábrica"):
                df.columns = [norm(c) for c in df.columns]
                pesos = df.apply(lambda r: calcular_pesos_madeira(r.get("LARG",0), r.get("COMP",0), r.get("QUANT",0), r["MATERIAL"]), axis=1)
                df["PESO_UNIT"] = pesos.apply(lambda x: x[0]); df["PESO_TOTAL"] = pesos.apply(lambda x: x[1])
                if "COR" in df.columns: df["COR"] = df["COR"].apply(lambda x: m_cores.get(norm(x), str(x).split('.')[0]))
                df["MATERIAL"] = df["MATERIAL"].apply(limpa_material)
                for c in ["CORTE", "FITA", "USINAGEM"]: df[c] = ""
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    ws = writer.book.create_sheet("PRODUCAO")
                    ws.cell(row=1, column=1, value=f"TECAMA | PEDIDO: {nome_f}").font = Font(bold=True, size=14)
                    ws.merge_cells(start_row=1, end_row=1, start_column=1, end_column=12)
                    header = ["QUANT","COMP","LARG","MATERIAL","COR (COD)","DESCPECA","PRODUTO","CORTE","FITA","USINAGEM","PESO UNIT.","PESO TOTAL"]
                    for i, h in enumerate(header, 1):
                        cell = ws.cell(row=3, column=i, value=h); cell.font = Font(bold=True); cell.alignment = Alignment(horizontal="center")
                    
                    curr = 4
                    borda_fin = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
                    for dp, g in df.sort_values(by="DES_PAI").groupby("DES_PAI", sort=False):
                        ini = curr
                        for _, r in g.iterrows():
                            for i, c_nome in enumerate(["QUANT","COMP","LARG","MATERIAL","COR","DESCPECA","DES_PAI","CORTE","FITA","USINAGEM","PESO_UNIT","PESO_TOTAL"], 1):
                                cell = ws.cell(row=curr, column=i, value=r.get(c_nome, ""))
                                cell.border = borda_fin
                                cell.alignment = Alignment(horizontal="center", vertical="center")
                            curr += 1
                        ws.merge_cells(start_row=ini, end_row=curr-1, start_column=7, end_column=7)
                        curr += 1
                    for i in range(1, 13): ws.column_dimensions[get_column_letter(i)].width = 20
                st.download_button("📥 Baixar Excel de Produção", output.getvalue(), f"PRODUCAO_{nome_f}.xlsx")

        st.markdown("---")
        # FASE 2: CORTE CERTO
        st.subheader("2️⃣ Fase 2: Gerar CSV para Corte Certo")
        up_excel_f2 = st.file_uploader("Suba o Excel que você editou", type="xlsx", key="f2")
        if up_excel_f2:
            if st.button("🚀 Gerar CSV para Corte Certo"):
                df_edit = pd.read_excel(up_excel_f2, skiprows=2)
                df_edit = df_edit.dropna(subset=['QUANT', 'COMP', 'LARG'], how='all')
                col_cc = ["QUANT", "COMP", "LARG", "COR (COD)", "DESCPECA"]
                df_cc = df_edit[col_cc].copy()
                df_cc.insert(0, "ID", range(1, len(df_cc) + 1))
                csv_out = df_cc.to_csv(index=False, sep=";", encoding="utf-8-sig")
                st.download_button("📥 Baixar CSV Corte Certo", csv_out, f"CORTE_CERTO_{nome_f}.csv")

    with aba_cores:
        df_c = conn.read(worksheet="CORES_MARCENARIA", ttl=0)
        st.data_editor(df_c, num_rows="dynamic", use_container_width=True, key="ed_cores")
        if st.button("💾 Salvar Cores"):
            conn.update(worksheet="CORES_MARCENARIA", data=df_c); st.success("Salvo!")

# ==========================================
# PÁGINA: METALURGIA
# ==========================================
elif st.session_state.nav == "⚙️ Metalurgia":
    st.header("⚙️ Metalurgia")
    aba_calc, aba_db = st.tabs(["📋 Calculadora PDF", "🛠️ Gerenciar Tabelas Base"])
    try:
        db_map = conn.read(worksheet="MAPEAMENTO_TIPO", ttl=5)
        db_metro = conn.read(worksheet="PESO_POR_METRO", ttl=5)
        db_conj = conn.read(worksheet="PESO_CONJUNTO", ttl=5)
        dict_m = dict(zip(db_metro['secao'].apply(norm), db_metro['peso_kg_m']))
        list_m = db_map.to_dict('records')
        list_c = db_conj.to_dict('records')
    except: st.error("Erro nas tabelas.")

    with aba_calc:
        up_pdf = st.file_uploader("Suba o PDF Pontta", type="pdf")
        if up_pdf:
            itens = []
            with pdfplumber.open(up_pdf) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        for r in table:
                            if r and len(r) > 3 and str(r[0]).strip().replace('.','').isdigit():
                                itens.append({"QTD": r[0], "DESCRIÇÃO": r[1], "MEDIDA": r[3], "COR": r[2]})
            df_ed = st.data_editor(pd.DataFrame(itens), num_rows="dynamic", use_container_width=True)
            if st.button("🚀 Calcular e Exportar Metalurgia"):
                res = []
                for _, r in df_ed.iterrows():
                    desc_l = norm(str(r.get('DESCRIÇÃO')))
                    qtd = float(str(r.get('QTD', 0)).replace(',','.')) if r.get('QTD') else 0.0
                    tipo = "DESCONHECIDO"
                    for regra in list_m:
                        if norm(regra.get('texto_contido')) in desc_l:
                            tipo = str(regra.get('tipo', 'DESCONHECIDO')).upper(); break
                    if tipo == "IGNORAR": continue
                    p_u = 0.0
                    if tipo == "CONJUNTO":
                        for c in list_c:
                            if norm(c.get('nome_conjunto')) in desc_l: p_u = float(c.get('peso_unit_kg', 0)); break
                    elif "TUBO" in tipo or tipo in dict_m:
                        m_r = str(r.get('MEDIDA', '0')).lower().replace('mm','').replace(',','.').strip()
                        med = float(m_r) if m_r else 0.0
                        sec = norm(tipo.replace('TUBO ', '').strip())
                        p_u = (med / 1000) * dict_m.get(sec, 0.0)
                    res.append({"QTD": qtd, "DESCRIÇÃO": r.get('DESCRIÇÃO'), "MEDIDA": r.get('MEDIDA'), "TIPO": tipo, "PESO UNIT.": round(p_u, 3), "PESO TOTAL": round(p_u * qtd, 3)})
                df_res = pd.DataFrame(res)
                st.metric("Total", f"{df_res['PESO TOTAL'].sum():.2f} kg")
                st.dataframe(df_res)
                out_m = io.BytesIO()
                with pd.ExcelWriter(out_m, engine="openpyxl") as writer:
                    df_res.to_excel(writer, index=False, sheet_name="METALURGIA", startrow=1)
                    ws = writer.sheets["METALURGIA"]
                    for i in range(1, 7): ws.column_dimensions[get_column_letter(i)].width = 25
                st.download_button("📥 Baixar Excel Metalurgia", out_m.getvalue(), f"METAL_{up_pdf.name}.xlsx")

    with aba_db:
        st.subheader("🛠️ Gestão de Tabelas")
        if 't_ativa' not in st.session_state: st.session_state.t_ativa = "MAPEAMENTO_TIPO"
        c1, c2, c3 = st.columns(3)
        if c1.button("📋 Mapeamento"): st.session_state.t_ativa = "MAPEAMENTO_TIPO"
        if c2.button("⚖️ Tubos"): st.session_state.t_ativa = "PESO_POR_METRO"
        if c3.button("📦 Conjuntos"): st.session_state.t_ativa = "PESO_CONJUNTO"
        df_v = conn.read(worksheet=st.session_state.t_ativa, ttl=0)
        st.data_editor(df_v, num_rows="dynamic", use_container_width=True)
