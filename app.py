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

# --- 2. CSS ---
st.markdown("""
<style>
[data-testid="stSidebar"] .stRadio div[role="radiogroup"] label { font-size: 22px; font-weight: 600; }
h1 { color: #FF5722; }
.stButton > button {
    background-color: #FF5722; color: white; width: 100%;
    border-radius: 12px; font-weight: bold; height: 3.5em;
}
</style>
""", unsafe_allow_html=True)

conn = st.connection("gsheets", type=GSheetsConnection)

# --- 3. FUNÇÕES ---
def norm(t):
    if t is None or pd.isna(t):
        return ""
    t = unicodedata.normalize("NFD", str(t).upper()).encode("ascii", "ignore").decode("utf-8")
    return " ".join(t.split()).strip()

def calcular_pesos_madeira(larg, comp, quant, material_texto):
    PESO_M2_BASE = {"MDP": 12.0, "MDF": 13.5}
    try:
        l, c, q = float(larg), float(comp), float(quant)
        m_norm = norm(material_texto)
        tipo = "MDF" if "MDF" in m_norm else "MDP"
        esp = re.search(r"(\d+)\s*MM", m_norm)
        e = float(esp.group(1)) if esp else 18.0
        peso_uni = (l / 1000) * (c / 1000) * PESO_M2_BASE[tipo] * (e / 18)
        return round(peso_uni, 2), round(peso_uni * q, 2)
    except Exception:
        return 0.0, 0.0

# --- 4. NAVEGAÇÃO ---
if "nav" not in st.session_state:
    st.session_state.nav = "🏠 Início"

with st.sidebar:
    if os.path.exists("logo_tecama.png"):
        st.image("logo_tecama.png", use_container_width=True)
    st.session_state.nav = st.radio(
        "NAVEGAÇÃO",
        ["🏠 Início", "🌲 Marcenaria", "⚙️ Metalurgia"],
        index=["🏠 Início", "🌲 Marcenaria", "⚙️ Metalurgia"].index(st.session_state.nav)
    )
    st.caption("Tecama Hub Industrial v9.8")

# ===============================
# INÍCIO
# ===============================
if st.session_state.nav == "🏠 Início":
    st.title("Tecama Hub Industrial")
    st.markdown("### Bem-vindo ao Sistema Unificado de Produção")

# ===============================
# MARCENARIA
# ===============================
elif st.session_state.nav == "🌲 Marcenaria":
    st.header("🌲 Operações de Marcenaria")
    aba_conv, aba_cores = st.tabs(["📋 Processadores", "🎨 Editar Cores"])

    with aba_conv:
        try:
            df_cores = conn.read(worksheet="CORES_MARCENARIA", ttl=5)
            mapa_cores = {norm(r["descricao"]): str(r["codigo"]).split(".")[0] for _, r in df_cores.iterrows()}
        except Exception:
            mapa_cores = {}

        up_csv = st.file_uploader("CSV do Pontta", type="csv")
        if up_csv:
            df_raw = pd.read_csv(up_csv, sep=None, engine="python", dtype=str)
            df_raw.columns = [norm(c) for c in df_raw.columns]

            if "LARG" in df_raw.columns:
                df_p = df_raw[df_raw["LARG"].apply(lambda x: str(x).replace(",", ".").replace(".", "").isdigit())].copy()
            else:
                df_p = df_raw.copy()

            if "DES_PAI" not in df_p.columns:
                df_p["DES_PAI"] = ""

            if st.button("🚀 Gerar Excel"):
                pesos = df_p.apply(
                    lambda r: calcular_pesos_madeira(
                        r.get("LARG", 0),
                        r.get("COMP", 0),
                        r.get("QUANT", 0),
                        r.get("MATERIAL", "")
                    ),
                    axis=1
                )

                df_p["PESO_UNIT"], df_p["PESO_TOTAL"] = zip(*pesos)

                if "COR" in df_p.columns:
                    df_p["COR (COD)"] = df_p["COR"].apply(lambda x: mapa_cores.get(norm(x), x))
                elif "COR (COD)" not in df_p.columns:
                    df_p["COR (COD)"] = ""

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df_p.head(0).to_excel(writer, sheet_name="PRODUCAO", index=False)
                    ws = writer.sheets["PRODUCAO"]

                    ws.cell(1, 1, f"TECAMA | PEDIDO: {up_csv.name}").font = Font(bold=True, size=14)
                    ws.merge_cells(start_row=1, end_row=1, start_column=1, end_column=12)

                    headers = [
                        "QUANT", "COMP", "LARG", "MATERIAL", "COR (COD)",
                        "DESCPECA", "DES_PAI", "CORTE", "FITA",
                        "USINAGEM", "PESO_UNIT", "PESO_TOTAL"
                    ]

                    for i, h in enumerate(headers, 1):
                        ws.cell(3, i, h).font = Font(bold=True)

                    row = 4
                    borda = Border(*(Side(style="thin"),) * 4)

                    for _, r in df_p.iterrows():
                        for i, c in enumerate(headers, 1):
                            cell = ws.cell(row, i, r.get(c, ""))
                            cell.border = borda
                            cell.alignment = Alignment(horizontal="center")
                        row += 1

                    for i in range(1, 13):
                        ws.column_dimensions[get_column_letter(i)].width = 20

                st.download_button("📥 Baixar Excel", output.getvalue(), "PRODUCAO.xlsx")

    with aba_cores:
        df_c = conn.read(worksheet="CORES_MARCENARIA", ttl=0)
        novo = st.data_editor(df_c, num_rows="dynamic", use_container_width=True)
        if st.button("💾 Salvar Cores"):
            conn.update(worksheet="CORES_MARCENARIA", data=novo)
            st.success("Salvo")

# ===============================
# METALURGIA
# ===============================
elif st.session_state.nav == "⚙️ Metalurgia":
    st.header("⚙️ Metalurgia")

    try:
        db_map = conn.read(worksheet="MAPEAMENTO_TIPO", ttl=5)
        db_metro = conn.read(worksheet="PESO_POR_METRO", ttl=5)
        db_conj = conn.read(worksheet="PESO_CONJUNTO", ttl=5)
    except Exception as e:
        st.error(f"Erro ao carregar tabelas: {e}")
        st.stop()

    dict_m = dict(zip(db_metro["secao"].apply(norm), db_metro["peso_kg_m"]))
    regras = db_map.to_dict("records")
    conjuntos = db_conj.to_dict("records")

    up_pdf = st.file_uploader("PDF Pontta", type="pdf")
    if up_pdf:
        itens = []
        with pdfplumber.open(up_pdf) as pdf:
            for p in pdf.pages:
                for t in p.extract_tables():
                    for r in t:
                        if r and len(r) > 3 and str(r[0]).replace(".", "").isdigit():
                            itens.append({
                                "QTD": r[0],
                                "DESCRIÇÃO": r[1],
                                "COR": r[2],
                                "MEDIDA": r[3]
                            })

        df = st.data_editor(pd.DataFrame(itens), num_rows="dynamic")

        if st.button("🚀 Calcular"):
            res = []
            for _, r in df.iterrows():
                desc = norm(r["DESCRIÇÃO"])
                qtd = float(str(r["QTD"]).replace(",", ".")) if r["QTD"] else 0

                tipo = "DESCONHECIDO"
                for regra in regras:
                    if norm(regra["texto_contido"]) in desc:
                        tipo = regra["tipo"]
                        break

                if tipo == "IGNORAR":
                    continue

                peso_u = 0
                if tipo == "CONJUNTO":
                    for c in conjuntos:
                        if norm(c["nome_conjunto"]) in desc:
                            peso_u = float(c["peso_unit_kg"])
                            break
                else:
                    try:
                        med = float(str(r["MEDIDA"]).lower().replace("mm", "").replace(",", "."))
                        sec = norm(tipo.replace("TUBO", ""))
                        peso_u = (med / 1000) * dict_m.get(sec, 0)
                    except Exception:
                        peso_u = 0

                res.append({
                    "QTD": qtd,
                    "DESCRIÇÃO": r["DESCRIÇÃO"],
                    "TIPO": tipo,
                    "PESO UNIT.": round(peso_u, 3),
                    "PESO TOTAL": round(peso_u * qtd, 3)
                })

            df_res = pd.DataFrame(res)
            st.metric("Peso Total", f"{df_res['PESO TOTAL'].sum():.2f} kg")
            st.dataframe(df_res, use_container_width=True)
