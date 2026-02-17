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
    h1 { color: #FF5722 !important; }
    .stButton > button { background-color: #FF5722; color: white; width: 100%; border-radius: 12px; font-weight: bold; height: 3.5em; }
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
    opcao = st.radio("NAVEGAÇÃO", ["🏠 Início", "🌲 Marcenaria", "⚙️ Metalurgia"])
    st.caption("Tecama Hub v6.7")

# ==========================================
# PÁGINA: INÍCIO
# ==========================================
if opcao == "🏠 Início":
    st.title("Tecama Hub Industrial")
    st.markdown("### Bem-vindo ao Sistema Unificado de Produção")
    st.write("Plataforma de processamento de pedidos integrada ao sistema **Pontta**.")
    st.markdown("---")
    st.subheader("🌲 Divisão de Marcenaria")
    st.write("Processamento de arquivos CSV com padronização de materiais e cálculo de pesos.")
    st.subheader("⚙️ Divisão de Metalurgia")
    st.write("Levantamento de peso através de relatórios PDF com filtragem inteligente de peças.")

# ==========================================
# PÁGINA: MARCENARIA
# ==========================================
elif opcao == "🌲 Marcenaria":
    st.header("🌲 Marcenaria")
    aba_conv, aba_cores = st.tabs(["📋 Processar Pedido", "🎨 Cores"])

    with aba_conv:
        try:
            df_cores_gs = conn.read(worksheet="CORES_MARCENARIA", ttl=5)
            m_cores = {norm(r["descricao"]): str(r["codigo"]).split('.')[0].strip() for _, r in df_cores_gs.iterrows()}
        except: m_cores = {}

        up_csv = st.file_uploader("Suba o CSV Pontta", type="csv")
        if up_csv:
            df_b = pd.read_csv(up_csv, sep=None, engine='python', dtype=str)
            nome_f = up_csv.name.replace(".csv", "").upper()
            l_teste = pd.to_numeric(df_b.iloc[0].get('LARG', ''), errors='coerce')
            df = df_b.iloc[1:].copy() if pd.isna(l_teste) else df_b.copy()

            if st.button("🚀 Gerar Planilha"):
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
                                if c_nome == "DES_PAI": # Coluna Produto
                                    cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
                            soma += float(r.get("PESO_TOTAL", 0)); curr += 1
                        if len(g) > 1: ws.merge_cells(start_row=ini, end_row=curr-1, start_column=7, end_column=7)
                        curr += 1
                    
                    # AutoAjuste e Configuração de Coluna Produto
                    for i, col in enumerate(ws.columns, 1):
                        letter = get_column_letter(i)
                        if letter == 'G': # Coluna Produto (DES_PAI)
                            ws.column_dimensions[letter].width = 30
                        else:
                            max_len = 0
                            for cell in col:
                                try: max_len = max(max_len, len(str(cell.value)))
                                except: pass
                            ws.column_dimensions[letter].width = max_len + 4
                            
                st.download_button("📥 Baixar Excel Marcenaria", output.getvalue(), f"PROD_{nome_f}.xlsx")

# ==========================================
# PÁGINA: METALURGIA
# ==========================================
elif opcao == "⚙️ Metalurgia":
    st.header("⚙️ Metalurgia")
    aba_calc, aba_db = st.tabs(["📋 Calculadora", "🛠️ Gerenciar"])

    if 'db_mapeamento' not in st.session_state:
        st.session_state.db_mapeamento = conn.read(worksheet="MAPEAMENTO_TIPO", ttl=5)
        st.session_state.db_pesos_metro = conn.read(worksheet="PESO_POR_METRO", ttl=5)
        st.session_state.db_pesos_conjunto = conn.read(worksheet="PESO_CONJUNTO", ttl=5)

    with aba_calc:
        def calcular_metal(df_input):
            map_rules = st.session_state.db_mapeamento.to_dict('records')
            dict_metro = dict(zip(st.session_state.db_pesos_metro['secao'], st.session_state.db_pesos_metro['peso_kg_m']))
            dict_conjunto = dict(zip(st.session_state.db_pesos_conjunto['nome_conjunto'], st.session_state.db_pesos_conjunto['peso_unit_kg']))
            res = []
            for _, r in df_input.iterrows():
                desc = str(r['DESCRIÇÃO']); qtd = float(r['QTD']) if r['QTD'] else 0.0
                tipo = "DESCONHECIDO"
                for regra in map_rules:
                    if str(regra['texto_contido']).upper() in desc.upper(): tipo = regra['tipo']; break
                
                # FILTRO DE IGNORAR
                if tipo == "IGNORAR": continue

                medida = 0.0
                try: medida = float(str(r['MEDIDA']).lower().replace('mm','').strip())
                except: pass
                p_unit = 0.0
                if tipo == 'CONJUNTO':
                    for n, p in dict_conjunto.items():
                        if n.upper() in desc.upper(): p_unit = p; break
                elif 'tubo' in tipo.lower():
                    sec = tipo.lower().replace('tubo ', '').strip()
                    p_unit = (medida/1000) * dict_metro.get(sec, 0.0)
                res.append({"QTD": qtd, "DESCRIÇÃO": desc, "MEDIDA": r['MEDIDA'], "TIPO": tipo, "PESO_TOTAL": round(p_unit * qtd, 3)})
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
            if st.button("🚀 Calcular e Gerar Excel"):
                res_met = calcular_metal(df_edit)
                st.metric("Peso Total", f"{res_met['PESO_TOTAL'].sum():.2f} kg")
                
                output_met = io.BytesIO()
                with pd.ExcelWriter(output_met, engine="openpyxl") as writer:
                    res_met.to_excel(writer, index=False, sheet_name="METALURGIA")
                    ws_met = writer.sheets["METALURGIA"]
                    # AutoAjuste Metalurgia
                    for col in ws_met.columns:
                        max_len = 0
                        column = col[0].column_letter
                        for cell in col:
                            try: max_len = max(max_len, len(str(cell.value)))
                            except: pass
                        ws_met.column_dimensions[column].width = max_len + 4
                
                st.download_button("📥 Baixar Excel Metalurgia", output_met.getvalue(), f"METAL_{up_pdf.name}.xlsx")

    with aba_db:
        # Lógica de edição permanece igual...
        tabela_sel = st.selectbox("Escolha a tabela:", ["MAPEAMENTO_TIPO", "PESO_POR_METRO", "PESO_CONJUNTO"])
        df_m = conn.read(worksheet=tabela_sel, ttl=0)
        dados_novos = st.data_editor(df_m, num_rows="dynamic", use_container_width=True)
        if st.button(f"💾 Salvar {tabela_sel}"):
            conn.update(worksheet=tabela_sel, data=dados_novos)
            st.success("Salvo!")
