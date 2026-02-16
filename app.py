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
    </style>
    """, unsafe_allow_html=True)

# --- 3. CONEXÃO COM GOOGLE SHEETS ---
conn = st.connection("gsheets", type=GSheetsConnection)

# --- 4. FUNÇÕES DE AUXÍLIO MARCENARIA ---
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
    st.markdown("<h1 style='text-align: center;'>🏗️ TECAMA</h1>", unsafe_allow_html=True)
    opcao = st.radio("Selecione a Divisão:", ["🏠 Início", "🪵 Marcenaria (CSV)", "⚙️ Metalurgia (PDF)"])
    st.markdown("---")
    st.caption("Tecama Hub v5.2")

# ==========================================
# DIVISÃO 1: MARCENARIA (CONVERSOR CSV)
# ==========================================
if opcao == "🪵 Marcenaria (CSV)":
    st.header("🪵 Divisão de Marcenaria")
    
    try:
        df_cores_gs = conn.read(worksheet="CORES_MARCENARIA", ttl=5)
        # Garante que o código da cor seja lido sem .0
        m_cores = {norm(r["descricao"]): str(r["codigo"]).split('.')[0].strip() for _, r in df_cores_gs.iterrows()}
    except:
        st.error("Erro: Aba 'CORES_MARCENARIA' não encontrada no Sheets.")
        m_cores = {}

    up_csv = st.file_uploader("Suba o arquivo CSV", type="csv")
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

        if st.button("🚀 Gerar Excel de Produção"):
            df.columns = [norm(c) for c in df.columns]
            
            # Pesos e Limpeza
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
                ws.cell(row=1, column=1, value=f"TECAMA | PEDIDO: {tit}").font = Font(bold=True, size=14, color="F97316")
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
                    
                    # Pula linha sem adicionar bordas nela
                    curr += 1
                
                ws.cell(row=curr+1, column=11, value="TOTAL:").font = Font(bold=True)
                ws.cell(row=curr+1, column=12, value=f"{round(soma, 2)} kg").font = Font(bold=True)
                
                # Bordas seletivas (Apenas se a linha tiver conteúdo)
                borda = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
                for row in ws.iter_rows(min_row=3, max_row=curr-1):
                    # Verifica se a linha não é a separadora (testando se a primeira célula tem valor)
                    if any(cell.value for cell in row):
                        for cell in row:
                            cell.border = borda

                # AutoFit (Ignora células mescladas para evitar erros)
                for col_idx in range(1, 13):
                    max_l = 0
                    for row_idx in range(3, curr):
                        cell = ws.cell(row=row_idx, column=col_idx)
                        if cell.value and not any(cell.coordinate in rng for rng in ws.merged_cells.ranges):
                            max_l = max(max_l, len(str(cell.value)))
                    ws.column_dimensions[get_column_letter(col_idx)].width = max_l + 5

            st.success(f"✅ Excel formatado com sucesso! Peso: {round(soma, 2)} kg")
            st.download_button("📥 Baixar Planilha", output.getvalue(), f"PROD_{nome_f}.xlsx")

# ==========================================
# DIVISÃO 2: METALURGIA (PDF)
# ==========================================
elif opcao == "⚙️ Metalurgia (PDF)":
    st.header("⚙️ Metalurgia System")
    st.info("Área de cálculo de estruturas metálicas ativa conforme base de dados do GSheets.")
    # Cole aqui sua lógica PDF se necessário

elif opcao == "🏠 Início":
    st.title("Hub Industrial Tecama")
    st.write("Portal unificado para processamento de marcenaria e metalurgia.")
