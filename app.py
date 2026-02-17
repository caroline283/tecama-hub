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
    
    /* Títulos clicáveis na Home */
    .home-link .stButton > button {
        background-color: transparent !important;
        color: #FF5722 !important;
        border: none !important;
        font-size: 24px !important;
        font-weight: bold !important;
        text-align: left !important;
        padding: 0 !important;
        height: auto !important;
        text-decoration: underline !important;
        box-shadow: none !important;
    }
    .home-link .stButton > button:hover { color: #E64A19 !important; text-decoration: none !important; }

    /* Botões Laranja */
    .stButton > button {
        background-color: #FF5722;
        color: white;
        width: 100%;
        border-radius: 12px;
        font-weight: bold;
        height: 3.5em;
        font-size: 16px;
        border: none;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
    }
    .stButton > button:hover { background-color: #E64A19; transform: translateY(-2px); }
    </style>
    """, unsafe_allow_html=True)

conn = st.connection("gsheets", type=GSheetsConnection)

# --- 3. FUNÇÕES AUXILIARES ---
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

# --- 4. MENU LATERAL ---
with st.sidebar:
    if os.path.exists("logo_tecama.png"): st.image("logo_tecama.png", use_container_width=True)
    if 'nav' not in st.session_state: st.session_state.nav = "🏠 Início"
    opcao = st.radio("NAVEGAÇÃO", ["🏠 Início", "🌲 Marcenaria", "⚙️ Metalurgia"], 
                     index=["🏠 Início", "🌲 Marcenaria", "⚙️ Metalurgia"].index(st.session_state.nav))
    st.session_state.nav = opcao
    st.caption("Tecama Hub Industrial v7.4")

# ==========================================
# PÁGINA: INÍCIO (TEXTO v6.6 INTEGRAL)
# ==========================================
if st.session_state.nav == "🏠 Início":
    st.title("Tecama Hub Industrial")
    st.markdown("### Bem-vindo ao Sistema Unificado de Produção")
    st.write("Esta plataforma foi desenvolvida para centralizar as operações das divisões de **Marcenaria** e **Metalurgia**, garantindo agilidade no processamento de pedidos e precisão nos cálculos de engenharia.")
    st.markdown("---")
    
    st.markdown('<div class="home-link">', unsafe_allow_html=True)
    if st.button("🌲 Divisão de Marcenaria"):
        st.session_state.nav = "🌲 Marcenaria"
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown("""
    A página de Marcenaria é focada no **processamento de arquivos CSV gerados pelo Pontta**.
    * **Conversor:** Transforma listas brutas em planilhas de produção limpas, com nomes de materiais padronizados e cálculo automático de pesos.
    * **Gestão de Cores:** Permite editar em tempo real a tabela de códigos de cores, garantindo que o PDF de produção saia com as cores corretas da fábrica.
    """)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    st.markdown('<div class="home-link">', unsafe_allow_html=True)
    if st.button("⚙️ Divisão de Metalurgia"):
        st.session_state.nav = "⚙️ Metalurgia"
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)
    
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
elif st.session_state.nav == "🌲 Marcenaria":
    st.header("🌲 Operações de Marcenaria")
    aba_conv, aba_cores = st.tabs(["📋 Processar Pedido (CSV)", "🎨 Editar Tabela de Cores"])
    with aba_conv:
        try:
            df_cores_gs = conn.read(worksheet="CORES_MARCENARIA", ttl=5)
            m_cores = {norm(r["descricao"]): str(r["codigo"]).split('.')[0].strip() for _, r in df_cores_gs.iterrows()}
        except: m_cores = {}
        up_csv = st.file_uploader("Suba o arquivo CSV (Pontta)", type="csv")
        if up_csv:
            df_b = pd.read_csv(up_csv, sep=None, engine='python', dtype=str)
            nome_f = up_csv.name.replace(".csv", "").upper()
            l_teste = pd.to_numeric(df_b.iloc[0].get('LARG', ''), errors='coerce')
            df = df_b.iloc[1:].copy() if pd.isna(l_teste) else df_b.copy()
            if st.button("🚀 Gerar Planilha de Produção"):
                df.columns = [norm(c) for c in df.columns]
                pesos = df.apply(lambda r: calcular_pesos_madeira(r.get("LARG",0), r.get("COMP",0), r.get("QUANT",0), r["MATERIAL"]), axis=1)
                df["PESO_UNIT"] = pesos.apply(lambda x: x[0]); df["PESO_TOTAL"] = pesos.apply(lambda x: x[1])
                if "COR" in df.columns: df["COR"] = df["COR"].apply(lambda x: m_cores.get(norm(x), str(x).split('.')[0]))
                df["MATERIAL"] = df["MATERIAL"].apply(limpa_material)
                for c in ["CORTE", "FITA", "USINAGEM"]: df[c] = ""
                if "DES_PAI" in df.columns: df = df.sort_values(by="DES_PAI")
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    ws = writer.book.create_sheet("PRODUCAO")
                    ws.cell(row=1, column=1, value=f"TECAMA | PEDIDO: {nome_f}").font = Font(bold=True, size=14, color="FF5722")
                    ws.merge_cells(start_row=1, end_row=1, start_column=1, end_column=12)
                    header = ["QUANT","COMP","LARG","MATERIAL","COR (COD)","DESCPECA","PRODUTO","CORTE","FITA","USINAGEM","PESO UNIT.","PESO TOTAL"]
                    for i, h in enumerate(header, 1):
                        cell = ws.cell(row=3, column=i, value=h); cell.font = Font(bold=True); cell.alignment = Alignment(horizontal="center")
                    curr = 4; soma = 0.0
                    col_ordem = ["QUANT","COMP","LARG","MATERIAL","COR","DESCPECA","DES_PAI","CORTE","FITA","USINAGEM","PESO_UNIT","PESO_TOTAL"]
                    for dp, g in df.groupby("DES_PAI", sort=False):
                        ini = curr
                        for _, r in g.iterrows():
                            for i, c_nome in enumerate(col_ordem, 1):
                                cell = ws.cell(row=curr, column=i, value=r.get(c_nome, ""))
                                if c_nome == "DES_PAI": cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
                            soma += float(r.get("PESO_TOTAL", 0)); curr += 1
                        if len(g) > 1: ws.merge_cells(start_row=ini, end_row=curr-1, start_column=7, end_column=7)
                        curr += 1
                    ws.cell(row=curr+1, column=11, value="TOTAL:").font = Font(bold=True)
                    ws.cell(row=curr+1, column=12, value=f"{round(soma, 2)} kg").font = Font(bold=True)
                    borda = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
                    for row in ws.iter_rows(min_row=3, max_row=curr-1):
                        if any(cell.value for cell in row):
                            for cell in row: cell.border = borda
                    for i, col in enumerate(ws.columns, 1):
                        letter = get_column_letter(i)
                        if letter == 'G': ws.column_dimensions[letter].width = 35
                        else:
                            max_length = 0
                            for cell in col:
                                try:
                                    if cell.value: max_length = max(max_length, len(str(cell.value)))
                                except: pass
                            ws.column_dimensions[letter].width = max_length + 5
                st.download_button("📥 Baixar Planilha Marcenaria", output.getvalue(), f"PROD_{nome_f}.xlsx")
    with aba_cores:
        df_cores_edit = conn.read(worksheet="CORES_MARCENARIA", ttl=0)
        nova_tabela_cores = st.data_editor(df_cores_edit, num_rows="dynamic", use_container_width=True)
        if st.button("💾 Salvar Cores"):
            conn.update(worksheet="CORES_MARCENARIA", data=nova_tabela_cores)
            st.success("Salvo!")

# ==========================================
# PÁGINA: METALURGIA
# ==========================================
elif st.session_state.nav == "⚙️ Metalurgia":
    st.header("⚙️ Metalurgia")
    aba_calc, aba_db = st.tabs(["📋 Calculadora PDF (Pontta)", "🛠️ Gerenciar Tabelas Base"])
    if 'db_mapeamento' not in st.session_state:
        st.session_state.db_mapeamento = conn.read(worksheet="MAPEAMENTO_TIPO", ttl=5)
        st.session_state.db_pesos_metro = conn.read(worksheet="PESO_POR_METRO", ttl=5)
        st.session_state.db_pesos_conjunto = conn.read(worksheet="PESO_CONJUNTO", ttl=5)
    with aba_calc:
        def calcular_metal(df_input):
            map_rules = st.session_state.db_mapeamento.to_dict('records')
            dict_metro = dict(zip(st.session_state.db_pesos_metro['secao'].apply(norm), st.session_state.db_pesos_metro['peso_kg_m']))
            dict_conjunto = dict(zip(st.session_state.db_pesos_conjunto['nome_conjunto'].apply(norm), st.session_state.db_pesos_conjunto['peso_unit_kg']))
            res = []
            for _, r in df_input.iterrows():
                desc_limpa = norm(str(r['DESCRIÇÃO']))
                qtd = float(r['QTD']) if r['QTD'] else 0.0
                tipo = "DESCONHECIDO"
                for regra in map_rules:
                    if norm(regra['texto_contido']) in desc_limpa:
                        tipo = regra['tipo']; break
                if tipo == "IGNORAR": continue
                medida = 0.0
                try: medida = float(str(r['MEDIDA']).lower().replace('mm','').replace(',','.').strip())
                except: pass
                p_unit = 0.0
                if tipo == 'CONJUNTO':
                    for n_conj, p_val in dict_conjunto.items():
                        if n_conj in desc_limpa: p_unit = p_val; break
                elif 'tubo' in tipo.lower():
                    sec = norm(tipo.lower().replace('tubo ', '').strip())
                    p_unit = (medida/1000) * dict_metro.get(sec, 0.0)
                res.append({"QTD": qtd, "DESCRIÇÃO": str(r['DESCRIÇÃO']), "MEDIDA": r['MEDIDA'], "TIPO": tipo, "PESO UNIT.": round(p_unit, 3), "PESO TOTAL": round(p_unit * qtd, 3)})
            return pd.DataFrame(res)
        up_pdf = st.file_uploader("Suba o PDF Pontta", type="pdf")
        if up_pdf:
            itens = []
            with pdfplumber.open(up_pdf) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        for r in table:
                            if len(r) > 3 and str(r[0]).strip().replace('.','').isdigit():
                                itens.append({"QTD": r[0], "DESCRIÇÃO": r[1], "MEDIDA": r[3], "COR": r[2]})
            df_edit = st.data_editor(pd.DataFrame(itens), num_rows="dynamic", use_container_width=True)
            if st.button("🚀 Calcular e Gerar Excel Detalhado"):
                res_met = calcular_metal(df_edit)
                st.metric("Peso Total Estimado", f"{res_met['PESO TOTAL'].sum():.2f} kg")
                st.dataframe(res_met, use_container_width=True)
                output_met = io.BytesIO()
                with pd.ExcelWriter(output_met, engine="openpyxl") as writer:
                    res_met.to_excel(writer, index=False, sheet_name="METALURGIA", startrow=1)
                    ws_met = writer.sheets["METALURGIA"]
                    last_row = len(res_met) + 3
                    ws_met.cell(row=last_row, column=5, value="TOTAL GERAL:").font = Font(bold=True)
                    ws_met.cell(row=last_row, column=6, value=f"{res_met['PESO TOTAL'].sum():.2f} kg").font = Font(bold=True)
                    for col_idx in range(1, len(res_met.columns) + 1):
                        max_len = 0
                        letter = get_column_letter(col_idx)
                        for cell in ws_met[letter]:
                            try:
                                if cell.value: max_len = max(max_len, len(str(cell.value)))
                            except: pass
                        ws_met.column_dimensions[letter].width = max_len + 5
                st.download_button("📥 Baixar Excel Metalurgia", output_met.getvalue(), f"METAL_{up_pdf.name}.xlsx")
    with aba_db:
        if 'tab_m' not in st.session_state: st.session_state.tab_m = "MAPEAMENTO_TIPO"
        c1, c2, c3 = st.columns(3)
        if c1.button("📋 Mapeamento"): st.session_state.tab_m = "MAPEAMENTO_TIPO"
        if c2.button("⚖️ Tubos"): st.session_state.tab_m = "PESO_POR_METRO"
        if c3.button("📦 Conjuntos"): st.session_state.tab_m = "PESO_CONJUNTO"
        df_m = conn.read(worksheet=st.session_state.tab_m, ttl=0)
        dados_novos = st.data_editor(df_m, num_rows="dynamic", use_container_width=True)
        if st.button(f"💾 Salvar {st.session_state.tab_m}"):
            conn.update(worksheet=st.session_state.tab_m, data=dados_novos)
            st.success("Salvo!")
