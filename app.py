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

# --- 2. CSS PERSONALIZADO ---
st.markdown("""
    <style>
    h1 { color: #FF5722; }
    .stButton > button {
        background-color: #FF5722;
        color: white;
        width: 100%;
        border-radius: 8px;
        font-weight: bold;
    }
    div[data-testid="stMetric"] {
        background-color: #F8F9FA;
        border-left: 5px solid #FF5722;
        padding: 15px;
        border-radius: 5px;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 3. CONEXÃO COM GOOGLE SHEETS ---
conn = st.connection("gsheets", type=GSheetsConnection)

# --- 4. FUNÇÕES MARCENARIA ---
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
    # Mostra o logo se o arquivo estiver no GitHub
    if os.path.exists("logo_tecama.png"):
        st.image("logo_tecama.png", use_container_width=True)
    else:
        st.markdown("<h1 style='text-align: center;'>🏗️ TECAMA</h1>", unsafe_allow_html=True)
    
    opcao = st.radio("Selecione a Divisão:", ["🏠 Início", "🪵 Marcenaria (CSV)", "⚙️ Metalurgia (PDF)"])
    st.markdown("---")
    st.caption("Tecama Hub v5.7")

# ==========================================
# DIVISÃO 1: MARCENARIA
# ==========================================
if opcao == "🪵 Marcenaria (CSV)":
    st.header("🪵 Marcenaria")
    aba_csv, aba_cores = st.tabs(["📋 Conversor", "🛠️ Cores"])

    with aba_csv:
        try:
            df_cores_gs = conn.read(worksheet="CORES_MARCENARIA", ttl=5)
            m_cores = {norm(r["descricao"]): str(r["codigo"]).split('.')[0].strip() for _, r in df_cores_gs.iterrows()}
        except:
            m_cores = {}

        up_csv = st.file_uploader("Upload CSV", type="csv")
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

            if st.button("🚀 Gerar Excel"):
                df.columns = [norm(c) for c in df.columns]
                pesos = df.apply(lambda r: calcular_pesos_madeira(r.get("LARG",0), r.get("COMP",0), r.get("QUANT",0), r["MATERIAL"]), axis=1)
                df["PESO_UNIT"] = pesos.apply(lambda x: x[0])
                df["PESO_TOTAL"] = pesos.apply(lambda x: x[1])
                
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
                    for col_idx in range(1, 13):
                        ws.column_dimensions[get_column_letter(col_idx)].width = 18

                st.download_button("📥 Baixar Excel", output.getvalue(), f"PROD_{nome_f}.xlsx")

    with aba_cores:
        st.subheader("🛠️ Gestão de Cores")
        try:
            df_v = conn.read(worksheet="CORES_MARCENARIA", ttl=0)
            st.dataframe(df_v, use_container_width=True)
            st.markdown(f'<a href="https://docs.google.com/spreadsheets/d/{st.secrets["connections"]["gsheets"]["spreadsheet"]}/edit" target="_blank"><button style="background-color: #217346; color: white; padding: 10px; width: 100%; border: none; border-radius: 5px;">📝 Abrir Planilha Google</button></a>', unsafe_allow_html=True)
        except:
            st.warning("Verifique a aba CORES_MARCENARIA.")

# ==========================================
# DIVISÃO 2: METALURGIA
# ==========================================
elif opcao == "⚙️ Metalurgia (PDF)":
    st.header("⚙️ Metalurgia")
    
    if 'db_mapeamento' not in st.session_state:
        try:
            st.session_state.db_mapeamento = conn.read(worksheet="MAPEAMENTO_TIPO", ttl=5)
            st.session_state.db_pesos_metro = conn.read(worksheet="PESO_POR_METRO", ttl=5)
            st.session_state.db_pesos_conjunto = conn.read(worksheet="PESO_CONJUNTO", ttl=5)
        except:
            st.error("Erro nas tabelas de Metalurgia.")

    def calcular_metal(df_input):
        map_rules = st.session_state.db_mapeamento.to_dict('records')
        dict_metro = dict(zip(st.session_state.db_pesos_metro['secao'], st.session_state.db_pesos_metro['peso_kg_m']))
        dict_conjunto = dict(zip(st.session_state.db_pesos_conjunto['nome_conjunto'], st.session_state.db_pesos_conjunto['peso_unit_kg']))
        resultados = []
        for _, row in df_input.iterrows():
            desc = str(row['DESCRIÇÃO']); qtd = float(row['QTD']) if row['QTD'] else 0.0
            tipo_final = "DESCONHECIDO"
            for regra in map_rules:
                if str(regra['texto_contido']).upper() in desc.upper():
                    tipo_final = regra['tipo']; break
            
            medida_mm = 0.0
            try: medida_mm = float(str(row['MEDIDA']).lower().replace('mm','').strip())
            except: pass

            peso_unit = 0.0
            if tipo_final == 'CONJUNTO':
                for nome, p in dict_conjunto.items():
                    if nome.upper() in desc.upper(): peso_unit = p; break
            elif 'tubo' in tipo_final.lower():
                secao = tipo_final.lower().replace('tubo ', '').strip()
                peso_m = dict_metro.get(secao, 0.0)
                peso_unit = (medida_mm/1000) * peso_m
            
            resultados.append({"QTD": qtd, "DESCRIÇÃO": desc, "MEDIDA": row['MEDIDA'], "TIPO": tipo_final, "PESO_TOTAL": round(peso_unit * qtd, 3)})
        return pd.DataFrame(resultados)

    aba_calc, aba_db = st.tabs(["📋 Calculadora", "🛠️ Base de Dados"])

    with aba_calc:
        up_pdf = st.file_uploader("Upload PDF", type="pdf")
        if up_pdf:
            itens = []
            with pdfplumber.open(up_pdf) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        for row in table:
                            if len(row) > 3 and str(row[0]).strip().replace('.','').isdigit():
                                itens.append({"QTD": row[0], "DESCRIÇÃO": row[1], "MEDIDA": row[3], "COR": row[2]})
            
            df_edit = st.data_editor(pd.DataFrame(itens), num_rows="dynamic", use_container_width=True)
            if st.button("🚀 Calcular"):
                res_met = calcular_metal(df_edit)
                st.metric("Total", f"{res_met['PESO_TOTAL'].sum():.2f} kg")
                st.dataframe(res_met, use_container_width=True)

    with aba_db:
        st.info("Arquivo: base_metalurgia")
        st.markdown(f'<a href="https://docs.google.com/spreadsheets/d/{st.secrets["connections"]["gsheets"]["spreadsheet"]}/edit" target="_blank"><button style="background-color: #217346; color: white; padding: 10px; width: 100%; border: none; border-radius: 5px;">📂 Abrir Planilha Metalurgia</button></a>', unsafe_allow_html=True)

elif opcao == "🏠 Início":
    st.title("Tecama Hub")
    st.info("Selecione Marcenaria ou Metalurgia no menu lateral.")
