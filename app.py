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

# --- 2. CSS PARA VISUAL v6.6 (FONTE GRANDE) ---
st.markdown("""
    <style>
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] label { font-size: 22px !important; font-weight: 600 !important; color: #333 !important; }
    h1 { color: #FF5722 !important; font-family: 'Segoe UI', sans-serif; }
    h3 { color: #444 !important; }
    .home-link .stButton > button {
        background-color: transparent !important; color: #FF5722 !important; border: none !important;
        font-size: 24px !important; font-weight: bold !important; text-align: left !important;
        padding: 0 !important; height: auto !important; text-decoration: underline !important;
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
    """Limpa textos removendo acentos, quebras de linha e espaços extras"""
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
    st.caption("Tecama Hub Industrial v8.6")

# ==========================================
# PÁGINA: INÍCIO (v6.6)
# ==========================================
if st.session_state.nav == "🏠 Início":
    st.title("Tecama Hub Industrial")
    st.markdown("### Bem-vindo ao Sistema Unificado de Produção")
    st.write("Plataforma centralizada para Marcenaria e Metalurgia (Sistema Pontta).")
    st.markdown("---")
    st.markdown('<div class="home-link">', unsafe_allow_html=True)
    if st.button("🌲 Divisão de Marcenaria"):
        st.session_state.nav = "🌲 Marcenaria"; st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)
    st.write("Processamento de arquivos CSV (Pontta) com cálculo automático de pesos.")
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="home-link">', unsafe_allow_html=True)
    if st.button("⚙️ Divisão de Metalurgia"):
        st.session_state.nav = "⚙️ Metalurgia"; st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)
    st.write("Levantamento automático de peso através do relatório PDF (Pontta).")

# ==========================================
# PÁGINA: MARCENARIA
# ==========================================
elif st.session_state.nav == "🌲 Marcenaria":
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
                    ws.cell(row=1, column=1, value=f"TECAMA | PEDIDO: {nome_f}").font = Font(bold=True, size=14)
                    ws.merge_cells(start_row=1, end_row=1, start_column=1, end_column=12)
                    
                    header = ["QUANT","COMP","LARG","MATERIAL","COR (COD)","DESCPECA","PRODUTO","CORTE","FITA","USINAGEM","PESO UNIT.","PESO TOTAL"]
                    for i, h in enumerate(header, 1):
                        cell = ws.cell(row=3, column=i, value=h); cell.font = Font(bold=True); cell.alignment = Alignment(horizontal="center")
                    
                    curr = 4; soma = 0.0
                    borda_fin = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
                    
                    for dp, g in df.groupby("DES_PAI", sort=False):
                        ini = curr
                        for _, r in g.iterrows():
                            for i, c_nome in enumerate(["QUANT","COMP","LARG","MATERIAL","COR","DESCPECA","DES_PAI","CORTE","FITA","USINAGEM","PESO_UNIT","PESO_TOTAL"], 1):
                                val = r.get(c_nome, "")
                                cell = ws.cell(row=curr, column=i, value=val)
                                cell.border = borda_fin
                                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True if i == 7 else False)
                            soma += float(r.get("PESO_TOTAL", 0)); curr += 1
                        if len(g) > 1: ws.merge_cells(start_row=ini, end_row=curr-1, start_column=7, end_column=7)
                        curr += 1
                    
                    ws.cell(row=curr, column=11, value="TOTAL GERAL:").font = Font(bold=True)
                    ws.cell(row=curr, column=12, value=f"{round(soma, 2)} kg").font = Font(bold=True)
                    
                    for i in range(1, 13):
                        ws.column_dimensions[get_column_letter(i)].width = 35 if i == 7 else 18
                st.download_button("📥 Baixar Excel Marcenaria", output.getvalue(), f"PROD_{nome_f}.xlsx")

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
        dict_metro = dict(zip(db_metro['secao'].apply(norm), db_metro['peso_kg_m']))
        list_map = db_map.to_dict('records')
        list_conj = db_conj.to_dict('records')
    except: st.error("Erro ao carregar tabelas do Sheets.")

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
            df_editor = st.data_editor(pd.DataFrame(itens), num_rows="dynamic", use_container_width=True)
            if st.button("🚀 Calcular e Gerar Excel Metalurgia"):
                res = []
                for _, r in df_editor.iterrows():
                    desc_bruta = str(r.get('DESCRIÇÃO', ''))
                    desc_limpa = norm(desc_bruta)
                    qtd = float(str(r.get('QTD', 0)).replace(',','.')) if r.get('QTD') else 0.0
                    tipo_val = "DESCONHECIDO"
                    for regra in list_map:
                        txt_regra = norm(regra.get('texto_contido', ''))
                        if txt_regra and txt_regra in desc_limpa:
                            tipo_val = str(regra.get('tipo', 'DESCONHECIDO')).upper(); break
                    if tipo_val == "IGNORAR": continue
                    
                    p_unit = 0.0
                    if tipo_val == "CONJUNTO":
                        for c in list_conj:
                            nome_conj = norm(c.get('nome_conjunto', ''))
                            if nome_conj and nome_conj in desc_limpa:
                                p_unit = float(c.get('peso_unit_kg', 0)); break
                    elif tipo_val and 'TUBO' in tipo_val:
                        m_raw = str(r.get('MEDIDA', '0')).lower().replace('mm','').replace(',','.').strip()
                        medida = float(m_raw) if m_raw else 0.0
                        secao = norm(tipo_val.replace('TUBO ', '').strip())
                        p_unit = (medida / 1000) * dict_metro.get(secao, 0.0)
                    
                    res.append({"QTD": qtd, "DESCRIÇÃO": desc_bruta, "MEDIDA": r.get('MEDIDA', ''), "TIPO": tipo_val, "PESO UNIT.": round(p_unit, 3), "PESO TOTAL": round(p_unit * qtd, 3)})
                
                df_res = pd.DataFrame(res)
                st.metric("Total", f"{df_res['PESO TOTAL'].sum():.2f} kg")
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df_res.to_excel(writer, index=False, sheet_name="METALURGIA", startrow=1)
                    ws = writer.sheets["METALURGIA"]
                    borda_fin = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
                    
                    for row in ws.iter_rows(min_row=2, max_row=len(df_res)+2):
                        for cell in row:
                            cell.border = borda_fin
                            cell.alignment = Alignment(horizontal="center", vertical="center")
                    
                    for i in range(1, 7):
                        ws.column_dimensions[get_column_letter(i)].width = 25
                        
                    last_row = len(df_res) + 4
                    ws.cell(row=last_row, column=5, value="TOTAL GERAL:").font = Font(bold=True)
                    ws.cell(row=last_row, column=6, value=f"{df_res['PESO TOTAL'].sum():.2f} kg").font = Font(bold=True)
                st.download_button("📥 Baixar Excel Metalurgia", output.getvalue(), f"METAL_{up_pdf.name}.xlsx")
