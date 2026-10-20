import io
import os
import re
import sqlite3
import base64
from datetime import datetime
import openpyxl
import pandas as pd
import streamlit as st

# ==============================================================================
# 1. CAMADA DE BASE DE DADOS LOCAL (SQLITE)
# ==============================================================================
DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "historico_ovos_117.db")

class DatabaseManager:
    CHAVES_UNICIDADE = {
        "df_reclass": ["Data_Processamento", "Data_Postura", "Lote", "Origem"],
        "df_prod": ["Data Produção", "Nome Fazenda", "Nº Lote", "Nome Produto"],
        "df_vendas": ["Data_Expedicao", "Cliente", "Destino", "Data_Postura", "Lote", "Tipo_Ovo"],
        "df_transf": ["Data_Expedicao", "Aba_Destino", "Palete", "Lote", "Data_Postura", "Categoria"],
        "df_rec_mtech": ["Data Transação", "Data Produção", "Nº Lote", "Nome Produto", "Cód Trans Ovo"],
        "df_desc_mtech": ["Data Transação", "Data Produção", "Nº Lote", "Produto"],
        "df_est_mtech": ["Produto_Original"],
        "df_est_sankhya": ["Produto_Original"],
        "df_est_cpo": ["Categoria_Original"]
    }

    @staticmethod
    def _get_conn():
        return sqlite3.connect(DB_PATH, check_same_thread=False)

    @staticmethod
    def inicializar():
        tabelas = [
            "df_prod", "df_reclass", "df_vendas", "df_transf", 
            "df_rec_mtech", "df_desc_mtech", "df_est_mtech", 
            "df_est_sankhya", "df_est_cpo"
        ]
        with DatabaseManager._get_conn() as conn:
            for tab in tabelas:
                try:
                    df_db = pd.read_sql(f"SELECT * FROM {tab}", conn)
                    st.session_state[tab] = df_db
                except Exception:
                    st.session_state[tab] = pd.DataFrame()

    @staticmethod
    def salvar(nome_tabela: str, df_novo: pd.DataFrame):
        if df_novo is None or df_novo.empty:
            return 0, 0, []
        
        df_atual = DatabaseManager.obter(nome_tabela)
        chaves = DatabaseManager.CHAVES_UNICIDADE.get(nome_tabela, [])
        chaves_validas = [c for c in chaves if c in df_novo.columns and c in df_atual.columns]
        
        duplicados = []
        if not df_atual.empty and chaves_validas:
            dups = df_novo.merge(df_atual[chaves_validas].drop_duplicates(), on=chaves_validas, how='inner')
            if not dups.empty:
                duplicados = dups[chaves_validas].drop_duplicates().to_dict(orient='records')

        if not df_atual.empty:
            df_combinado = pd.concat([df_atual, df_novo], ignore_index=True)
        else:
            df_combinado = df_novo.copy()

        if chaves_validas:
            df_combinado = df_combinado.drop_duplicates(subset=chaves_validas, keep="last").reset_index(drop=True)
        else:
            df_combinado = df_combinado.drop_duplicates(keep="last").reset_index(drop=True)

        with DatabaseManager._get_conn() as conn:
            df_combinado.to_sql(nome_tabela, conn, if_exists="replace", index=False)
        
        st.session_state[nome_tabela] = df_combinado
        return len(df_novo), len(duplicados), duplicados

    @staticmethod
    def obter(nome_tabela: str) -> pd.DataFrame:
        if nome_tabela not in st.session_state or st.session_state[nome_tabela].empty:
            with DatabaseManager._get_conn() as conn:
                try:
                    st.session_state[nome_tabela] = pd.read_sql(f"SELECT * FROM {nome_tabela}", conn)
                except Exception:
                    st.session_state[nome_tabela] = pd.DataFrame()
        return st.session_state.get(nome_tabela, pd.DataFrame())

    @staticmethod
    def limpar_tabela(nome_tabela: str):
        with DatabaseManager._get_conn() as conn:
            conn.execute(f"DROP TABLE IF EXISTS {nome_tabela}")
        st.session_state[nome_tabela] = pd.DataFrame()

DatabaseManager.inicializar()

# ==============================================================================
# 2. SISTEMA DE AVISOS VISUAIS
# ==============================================================================
def show_warning(msg):
    st.markdown(f'<div style="background-color: #FFFBEB; border: 1px solid #FEF3C7; border-left: 6px solid #F59E0B; padding: 14px 18px; border-radius: 8px; color: #92400E; font-size: 15px; font-weight: 600; margin-bottom: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">⚠️ &nbsp; {msg}</div>', unsafe_allow_html=True)

def show_info(msg):
    st.markdown(f'<div style="background-color: #EFF6FF; border: 1px solid #DBEAFE; border-left: 6px solid #3B82F6; padding: 14px 18px; border-radius: 8px; color: #1E3A8A; font-size: 15px; font-weight: 600; margin-bottom: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">ℹ️ &nbsp; {msg}</div>', unsafe_allow_html=True)

def show_success(msg):
    st.markdown(f'<div style="background-color: #F0FDF4; border: 1px solid #DCFCE7; border-left: 6px solid #22C55E; padding: 14px 18px; border-radius: 8px; color: #166534; font-size: 15px; font-weight: 600; margin-bottom: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">✅ &nbsp; {msg}</div>', unsafe_allow_html=True)

def show_error(msg):
    st.markdown(f'<div style="background-color: #FEF2F2; border: 1px solid #FEE2E2; border-left: 6px solid #EF4444; padding: 14px 18px; border-radius: 8px; color: #991B1B; font-size: 15px; font-weight: 600; margin-bottom: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">❌ &nbsp; {msg}</div>', unsafe_allow_html=True)

# ==============================================================================
# 3. ESTILO VISUAL RESPONSIVO & PAINEL LATERAL
# ==============================================================================
st.set_page_config(page_title="Hub Operacional de Ovos", page_icon="🥚", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"], .stApp {
        font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif !important;
        font-size: 15px !important;
    }

    header[data-testid="stHeader"] { background-color: transparent !important; box-shadow: none !important; }

    .block-container {
        padding-top: clamp(1rem, 2vw, 1.5rem) !important; 
        padding-bottom: 3rem !important;
        padding-left: clamp(0.8rem, 2.5vw, 2rem) !important;
        padding-right: clamp(0.8rem, 2.5vw, 2rem) !important;
        max-width: 100% !important;
        width: 100% !important;
    }

    .stSelectbox label p, .stNumberInput label p, .stFileUploader label p {
        font-size: 15.5px !important; font-weight: 700 !important; color: var(--text-color) !important;
    }
    div[data-baseweb="select"] span, input { font-size: 14.5px !important; }

    section[data-testid="stSidebar"] {
        background-color: #132517 !important;
        border-right: 1px solid #0E1B11 !important;
    }
    section[data-testid="stSidebar"] * { color: #E2E8F0 !important; }
    
    .brand-box {
        background: #132517;
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 10px;
        padding: 14px 16px;
        margin-top: -20px !important;
        margin-bottom: 25px;
        display: flex;
        align-items: center;
        gap: 14px;
        width: 100%;
        box-sizing: border-box;
    }
    .brand-title { font-size: 1.05rem; font-weight: 800; color: #FFFFFF !important; margin: 0; }
    .brand-sub { font-size: 0.78rem; color: #94A3B8 !important; font-weight: 600; text-transform: uppercase; }

    .nav-label {
        font-size: 0.82rem; text-transform: uppercase; letter-spacing: 0.08em; font-weight: 800;
        color: #84A98C !important; margin: 18px 4px 12px 4px;
    }

    section[data-testid="stSidebar"] div[data-testid="stRadio"] { width: 100% !important; }
    section[data-testid="stSidebar"] div[data-testid="stRadio"] div[role="radiogroup"] {
        display: flex !important; flex-direction: column !important; align-items: stretch !important; width: 100% !important; gap: 10px !important;
    }
    section[data-testid="stSidebar"] div[data-testid="stRadio"] div[role="radiogroup"] label div:first-child { display: none !important; }
    section[data-testid="stSidebar"] div[data-testid="stRadio"] div[role="radiogroup"] label {
        background-color: #1A3620 !important; border: 1px solid #1A3620 !important; border-radius: 8px !important;
        min-height: 48px !important; height: auto !important; padding: 10px 16px !important; margin: 0 !important;
        width: 100% !important; display: flex !important; align-items: center !important; cursor: pointer !important;
    }
    section[data-testid="stSidebar"] div[data-testid="stRadio"] div[role="radiogroup"] label:hover { background-color: #25442A !important; border: 1px solid #4ADE80 !important; }
    section[data-testid="stSidebar"] div[data-testid="stRadio"] div[role="radiogroup"] label:has(input:checked) { background-color: #204428 !important; border-left: 5px solid #4ADE80 !important; }
    section[data-testid="stSidebar"] div[data-testid="stRadio"] div[role="radiogroup"] label p {
        color: #E2E8F0 !important; font-size: 0.95rem !important; font-weight: 600 !important; margin: 0 !important; width: 100% !important;
    }

    section[data-testid="stSidebar"] input { background-color: #1B3621 !important; color: #FFFFFF !important; border: 1px solid #2B5433 !important; border-radius: 6px !important; padding: 10px !important; width: 100%; box-sizing: border-box; }

    .cotacao-sidebar-card { background: #1A3620; border: 1px solid #2B5433; border-radius: 10px; padding: 14px 16px; margin-top: 18px; width: 100%; box-sizing: border-box; }
    .cotacao-sidebar-tag { font-size: 0.70rem; font-weight: 700; text-transform: uppercase; color: #84A98C !important; }
    .cotacao-sidebar-val { font-size: 1.35rem; font-weight: 800; color: #FFFFFF !important; margin-top: 4px; }

    .saas-panel { 
        background-color: var(--background-color) !important; 
        border: 1px solid var(--secondary-background-color) !important; 
        border-radius: 12px !important; 
        padding: clamp(14px, 2vw, 24px) !important; 
        box-shadow: 0 4px 6px rgba(0,0,0,0.05) !important; 
        margin-bottom: 24px !important; 
        box-sizing: border-box !important;
        width: 100% !important;
    }

    div[data-testid="stDataFrame"] div, div[data-testid="stDataFrame"] span, div[data-testid="stDataFrame"] p { 
        font-size: 14.5px !important; 
        color: #0F172A !important; 
    }
    th.col_heading, th.row_heading { color: #0F172A !important; font-weight: 800 !important; }

    .kpi-container { 
        display: grid !important; 
        grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)) !important; 
        gap: clamp(12px, 1.8vw, 18px) !important; 
        margin-bottom: 24px !important; 
        width: 100% !important; 
        box-sizing: border-box !important;
    }

    .kpi-card { 
        background: var(--background-color); 
        border: 1px solid var(--secondary-background-color); 
        border-radius: 12px; 
        padding: 16px 18px; 
        box-shadow: 0 3px 5px rgba(0,0,0,0.04); 
        display: flex; 
        align-items: center; 
        gap: 14px; 
        width: 100%; 
        box-sizing: border-box; 
        min-width: 0;
    }
    
    /* Novo estilo de bolha neutra e limpa para os ícones */
    .kpi-icon-bubble { 
        width: 48px; 
        height: 48px; 
        border-radius: 12px; 
        display: flex; 
        align-items: center; 
        justify-content: center; 
        font-size: 1.4rem; 
        flex-shrink: 0; 
        background-color: #F8FAFC; 
        border: 1px solid #E2E8F0;
        color: #475569;
    }
    
    .kpi-title { font-size: 0.78rem; font-weight: 700; text-transform: uppercase; color: var(--text-color); opacity: 0.7; margin: 0; }
    .kpi-value { font-size: clamp(1.25rem, 2.2vw, 1.55rem); font-weight: 800; color: var(--text-color); margin: 2px 0 0 0; word-break: break-word; }
    .kpi-sub { font-size: 0.80rem; font-weight: 600; margin: 4px 0 0 0; }

    .top-header-bar { 
        display: flex; 
        flex-wrap: wrap; 
        justify-content: space-between; 
        align-items: center; 
        gap: 12px;
        margin-bottom: 22px; 
        padding-bottom: 16px; 
        border-bottom: 2px solid var(--secondary-background-color); 
    }
    .top-header-title h1 { 
        font-size: clamp(1.3rem, 2.5vw, 1.7rem) !important; 
        font-weight: 800 !important; 
        color: var(--text-color) !important; 
        margin: 0 !important; 
    }
    .top-header-title span { font-size: 0.85rem !important; font-weight: 700 !important; color: #527958 !important; text-transform: uppercase !important; }
    
    .table-container-responsive {
        overflow-x: auto; 
        width: 100%; 
        border-radius: 10px; 
        border: 1px solid #CBD5E1; 
        margin-bottom: 22px;
        -webkit-overflow-scrolling: touch;
    }
    </style>
""", unsafe_allow_html=True)

# ==============================================================================
# 4. REGRAS DE NEGÓCIO E PADRONIZAÇÃO (MAPAS)
# ==============================================================================

MAP_SANKHYA = {
    "OVO FERTIL - NN": "01-OVO NINHO NORMAL",
    "OVO FERTIL - NS": "02-OVO NINHO SUJO",
    "OVO FERTIL - CM": "03-OVO DE CAMA",
    "04-OVO DUAS GEMAS GJA": "04-OVO DUAS GEMAS GJA",
    "05-OVO TRINCADO GJA": "05-OVO TRINCADO GJA",
    "08-OVO COMERCIO SJ GJA": "06-OVO SUJO GJA",
    "08-OVO INDUSTRIAL GJA": "08-OVO INDUSTRIAL GJA",
    "09-OVO PERDIDO": "09-OVO PERDIDO",
    "12- OVO DEFORMADO GJA": "12-OVO DEFORMADO GJA",
    "OVO LIMPO GJA": "13-OVO LIMPO GJA"
}

MAP_MTECH = {
    "01-OVO NINHO NORMAL": "01-OVO NINHO NORMAL",
    "02-OVO NINHO SUJO": "02-OVO NINHO SUJO",
    "03-OVO DE CAMA": "03-OVO DE CAMA",
    "04-OVO DUAS GEMAS GJA": "04-OVO DUAS GEMAS GJA",
    "05-OVO TRINCADO GJA": "05-OVO TRINCADO GJA",
    "06-OVO SUJO GJA": "06-OVO SUJO GJA",
    "08-OVO INDUSTRIAL GJA": "08-OVO INDUSTRIAL GJA",
    "09-OVO PERDIDO": "09-OVO PERDIDO",
    "12-OVO DEFORMADO GJA": "12-OVO DEFORMADO GJA",
    "13-OVO LIMPO GJA": "13-OVO LIMPO GJA"
}

MAP_CPO = {
    "NN": "01-OVO NINHO NORMAL",
    "NS": "02-OVO NINHO SUJO",
    "CM": "03-OVO DE CAMA",
    "DG": "04-OVO DUAS GEMAS GJA",
    "TR": "05-OVO TRINCADO GJA",
    "SJ": "06-OVO SUJO GJA",
    "ID": "08-OVO INDUSTRIAL GJA",
    "DF": "12-OVO DEFORMADO GJA",
    "LP": "13-OVO LIMPO GJA"
}

PADRAO_OBI = [
    "01-OVO NINHO NORMAL", 
    "02-OVO NINHO SUJO", 
    "03-OVO DE CAMA"
]

PADRAO_COMERCIAL = [
    "04-OVO DUAS GEMAS GJA", 
    "05-OVO TRINCADO GJA", 
    "06-OVO SUJO GJA", 
    "08-OVO INDUSTRIAL GJA", 
    "12-OVO DEFORMADO GJA", 
    "13-OVO LIMPO GJA"
]

PADRAO_ORDEM = PADRAO_OBI + PADRAO_COMERCIAL + ["09-OVO PERDIDO"]


def encurtar_nome_coluna(col_name):
    texto = str(col_name)
    match = re.search(r'(\d+)\s*[-_]\s*(\d+)$', texto)
    if match:
        return f"NP{match.group(1)}-{match.group(2)}"
    return texto

def normalizar_data_str(val, ano_padrao="2026"):
    if val is None or pd.isna(val): return ""
    if isinstance(val, (pd.Timestamp, datetime)): return val.strftime("%d/%m/%Y")
    s = str(val).strip().lower()
    m_full = re.match(r'^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})$', s)
    if m_full:
        d, m, y = m_full.groups()
        if len(y) == 2: y = "20" + y
        return f"{int(d):02d}/{int(m):02d}/{y}"
    meses = {"jan": "01", "fev": "02", "mar": "03", "abr": "04", "mai": "05", "jun": "06", "jul": "07", "ago": "08", "set": "09", "out": "10", "nov": "11", "dez": "12"}
    for mes_txt, mes_num in meses.items():
        if mes_txt in s:
            dia_match = re.search(r'\b(\d{1,2})\b', s)
            if dia_match: return f"{int(dia_match.group(1)):02d}/{mes_num}/{ano_padrao}"
    try:
        dt = pd.to_datetime(s, dayfirst=True)
        return dt.strftime("%d/%m/%Y")
    except Exception: return str(val).strip()

def mapear_categoria_mtech(texto):
    if not texto or pd.isna(texto): return "Outros"
    t = str(texto).upper().strip()

    if "DUAS GEMAS" in t or re.search(r'\bDG\b', t) or t.startswith("DG"): return "Duas Gemas"
    if "DEFORMADO" in t or re.search(r'\bDF\b', t) or t.startswith("DF"): return "Deformado"
    if "TRINCADO" in t or "MICRO" in t or re.search(r'\bTR\b', t) or t.startswith("TR"): return "Trincado"
    if "INDUSTRIAL" in t or re.search(r'\bID\b', t) or t.startswith("ID"): return "Industrial"
    if "NINHO SUJO" in t or re.search(r'\bNS\b', t) or t.startswith("NS"): return "Ninho Sujo"
    if "SUJO" in t or re.search(r'\bSJ\b', t) or t.startswith("SJ"): return "Sujo"
    if "CAMA" in t or re.search(r'\bCM\b', t) or t.startswith("CM"): return "Cama"
    if t == "NN" or re.search(r'\bNN\b', t) or "NINHO NORMAL" in t: return "Ninho Normal"
    if "LIMPO" in t or re.search(r'\bLP\b', t) or t.startswith("LP"): return "Ovo Limpo"
    if "PERDIDO" in t or "QUEBRA" in t or t == "PR" or re.search(r'\bPR\b', t) or t.startswith("PR"): return "Perda Galpão"

    return "Outros"

def safe_int(val):
    if val is None or val == "": return 0
    try: return int(float(val))
    except (ValueError, TypeError): return 0

def parse_reclassificacao_cpo(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws_name = next((s for s in wb.sheetnames if "classifica" in s.lower()), wb.sheetnames[0])
    ws = wb[ws_name]
    
    data_proc = normalizar_data_str(ws["C2"].value)
    resp = ws["G2"].value or ws["H2"].value or "Não informado"
    linhas = []
    secao_atual = "NN"

    for r in range(4, ws.max_row + 1):
        col_a = ws.cell(r, 1).value
        col_b = ws.cell(r, 2).value

        if col_a and "LOTES VELHOS" in str(col_a).upper(): secao_atual = "LOTES_VELHOS"; continue
        if str(col_a).strip().upper() in ["DATA POSTURA", "DATA PRODUÇÃO", "TOTAL =", "TOTAL"] or str(col_b).strip().upper() in ["LOTE", "LOTE:"]: continue

        if col_a is not None and col_b is not None:
            try: dt_post = normalizar_data_str(col_a); lote_num = int(float(col_b))
            except Exception: continue

            base_dict = {
                "Data_Processamento": data_proc, "Data_Postura": dt_post, "Lote": str(lote_num), 
                "Responsavel": str(resp), "P_NN": 0, "P_Cama": 0, "P_NS": 0, "P_Def": 0, 
                "P_Sujo": 0, "P_Ind": 0, "P_DG": 0, "P_Trinc": 0, "Trincado_Baixa": 0
            }

            if secao_atual == "NN":
                saldo_ini = safe_int(ws.cell(r, 3).value)
                if saldo_ini > 0:
                    df_val = safe_int(ws.cell(r, 4).value); ns_val = safe_int(ws.cell(r, 5).value)
                    id_val = safe_int(ws.cell(r, 6).value); dg_val = safe_int(ws.cell(r, 7).value)
                    sj_val = safe_int(ws.cell(r, 8).value); lp_val = safe_int(ws.cell(r, 9).value)
                    micro = safe_int(ws.cell(r, 10).value); trinc = safe_int(ws.cell(r, 11).value)
                    tot_trinc = safe_int(ws.cell(r, 12).value) or (micro + trinc)
                    perd_nn = safe_int(ws.cell(r, 13).value)
                    tot_destinos = df_val + ns_val + id_val + dg_val + sj_val + lp_val + tot_trinc
                    tot_class = safe_int(ws.cell(r, 14).value) or (tot_destinos + perd_nn)
                    saldo_fim = safe_int(ws.cell(r, 15).value) or (saldo_ini - tot_class)

                    trinc_descarte = trinc if trinc > 0 else tot_trinc

                    nn_dict = base_dict.copy()
                    nn_dict.update({
                        "Origem": "Ninho Normal", "Saldo_Inicial": saldo_ini, "Sujo": sj_val, 
                        "Deformado": df_val, "Ninho_Sujo": ns_val, "Industrial": id_val, 
                        "Duas_Gemas": dg_val, "Ovo_Limpo": lp_val, "Trincado": tot_trinc, 
                        "Microtrincado": micro, "Total_Destinos": tot_destinos, 
                        "Total_Classificado": tot_class, "Saldo_Final": saldo_fim, 
                        "P_NN": perd_nn, "Trincado_Baixa": trinc_descarte
                    })
                    linhas.append(nn_dict)

            elif secao_atual == "LOTES_VELHOS":
                ns_ini = safe_int(ws.cell(r, 3).value); cm_ini = safe_int(ws.cell(r, 4).value)
                if ns_ini > 0:
                    ns_dict = base_dict.copy(); ns_dict.update({"Origem": "Ninho Sujo", "Saldo_Inicial": ns_ini, "Sujo": ns_ini, "Deformado": 0, "Ninho_Sujo": 0, "Industrial": 0, "Duas_Gemas": 0, "Ovo_Limpo": 0, "Trincado": 0, "Microtrincado": 0, "Total_Destinos": ns_ini, "Total_Classificado": ns_ini, "Saldo_Final": 0, "Trincado_Baixa": 0})
                    linhas.append(ns_dict)
                if cm_ini > 0:
                    cm_dict = base_dict.copy(); cm_dict.update({"Origem": "Cama", "Saldo_Inicial": cm_ini, "Sujo": cm_ini, "Deformado": 0, "Ninho_Sujo": 0, "Industrial": 0, "Duas_Gemas": 0, "Ovo_Limpo": 0, "Trincado": 0, "Microtrincado": 0, "Total_Destinos": cm_ini, "Total_Classificado": cm_ini, "Saldo_Final": 0, "Trincado_Baixa": 0})
                    linhas.append(cm_dict)

    aba_perdidos = next((s for s in wb.sheetnames if "perdido" in s.lower() or "quebra" in s.lower()), None)
    if aba_perdidos:
        ws_p = wb[aba_perdidos]
        for r in range(2, ws_p.max_row + 1):
            val_a = ws_p.cell(r, 1).value
            if not val_a: continue
            str_a = str(val_a).strip().upper()
            if str_a in ["DATA POSTURA", "DATA DE POSTURA", "TOTAL", "TOTAL =", "DATA"]: continue
            try:
                dt_post = normalizar_data_str(val_a)
                lote_num = "-" 
                v_nn = safe_int(ws_p.cell(r, 3).value); v_cama = safe_int(ws_p.cell(r, 4).value); v_ns = safe_int(ws_p.cell(r, 5).value)
                v_def = safe_int(ws_p.cell(r, 6).value); v_sujo = safe_int(ws_p.cell(r, 7).value); v_ind = safe_int(ws_p.cell(r, 8).value); v_dg = safe_int(ws_p.cell(r, 9).value); v_trinc = safe_int(ws_p.cell(r, 10).value)
                if (v_nn + v_cama + v_ns + v_def + v_sujo + v_ind + v_dg + v_trinc) > 0:
                    linhas.append({"Data_Processamento": data_proc, "Data_Postura": dt_post, "Lote": lote_num, "Origem": "Ovos Perdidos (Aba)", "Saldo_Inicial": 0, "Sujo": 0, "Deformado": 0, "Ninho_Sujo": 0, "Industrial": 0, "Duas_Gemas": 0, "Ovo_Limpo": 0, "Trincado": 0, "Microtrincado": 0, "Total_Destinos": 0, "Total_Classificado": 0, "Saldo_Final": 0, "Responsavel": str(resp), "P_NN": v_nn, "P_Cama": v_cama, "P_NS": v_ns, "P_Def": v_def, "P_Sujo": v_sujo, "P_Ind": v_ind, "P_DG": v_dg, "P_Trinc": v_trinc, "Trincado_Baixa": 0})
            except Exception: pass

    return pd.DataFrame(linhas)

def parse_producao_erp(file_bytes):
    df = pd.read_excel(io.BytesIO(file_bytes))
    df.columns = [str(c).strip() for c in df.columns]
    
    faz_col = next((c for c in df.columns if "FAZENDA" in c.upper()), "Nome Fazenda")
    lote_col = next((c for c in df.columns if "LOTE" in c.upper()), "Nº Lote")
    prod_col = next((c for c in df.columns if "PRODUTO" in c.upper()), "Nome Produto")
    ovos_col = next((c for c in df.columns if "OVO" in c.upper()), "Ovos")
    dt_col = next((c for c in df.columns if "PRODU" in c.upper() or "DATA" in c.upper()), "Data Produção")

    df = df.dropna(subset=[lote_col, ovos_col])
    df["Nome Fazenda"] = df[faz_col].astype(str) if faz_col in df.columns else "Balsas 117"
    df["Nº Lote"] = df[lote_col].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    df["Nome Produto"] = df[prod_col].astype(str)
    df["Ovos"] = pd.to_numeric(df[ovos_col], errors="coerce").fillna(0).astype(int)
    df["Data Produção"] = df[dt_col].apply(normalizar_data_str)
    
    cols_agrupar = ["Nome Fazenda", "Nº Lote", "Nome Produto", "Data Produção"]
    df = df.groupby(cols_agrupar, dropna=False, as_index=False)["Ovos"].sum()
    return df

def parse_reclassificacao_mtech(file_bytes):
    df = pd.read_excel(io.BytesIO(file_bytes))
    df.columns = [str(c).strip() for c in df.columns]
    
    lote_col = next((c for c in df.columns if "LOTE" in c.upper()), "Nº Lote")
    nome_prod_col = next((c for c in df.columns if "NOME PRODUTO" in c.upper() or "PRODUTO" in c.upper()), "Nome Produto")
    cod_trans_col = next((c for c in df.columns if "CÓD TRANS" in c.upper() or "COD TRANS" in c.upper()), "Cód Trans Ovo")
    ovos_col = next((c for c in df.columns if "OVO" in c.upper() and "TRANS" not in c.upper() and "COD" not in c.upper() and "CÓD" not in c.upper()), "Ovos")
    prod_dt_col = next((c for c in df.columns if "DATA PRODU" in c.upper() or "DATA POSTURA" in c.upper() or ("PRODU" in c.upper() and "DATA" in c.upper())), "Data Produção")
    trans_dt_col = next((c for c in df.columns if ("DATA" in c.upper() and "TRANS" in c.upper()) or ("DATA" in c.upper() and "MOVIM" in c.upper())), "Data Transação")

    df = df.dropna(subset=[lote_col, ovos_col])
    
    df["Nº Lote"] = df[lote_col].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    df["Ovos"] = pd.to_numeric(df[ovos_col], errors='coerce').fillna(0).astype(int)
    df["Nome Produto"] = df[nome_prod_col].astype(str) if nome_prod_col in df.columns else ""
    df["Cód Trans Ovo"] = df[cod_trans_col].astype(str) if cod_trans_col in df.columns else ""
    df["Data Produção"] = df[prod_dt_col].apply(normalizar_data_str) if prod_dt_col in df.columns else ""
    df["Data Transação"] = df[trans_dt_col].apply(normalizar_data_str) if trans_dt_col in df.columns else ""
    
    cols_agrupar = ["Nº Lote", "Nome Produto", "Cód Trans Ovo", "Data Produção", "Data Transação"]
    df = df.groupby(cols_agrupar, dropna=False, as_index=False)["Ovos"].sum()
    return df

def parse_descarte_mtech(file_bytes):
    df = pd.read_excel(io.BytesIO(file_bytes))
    df.columns = [str(c).strip() for c in df.columns]
    
    lote_col = next((c for c in df.columns if "LOTE" in c.upper()), "Nº Lote")
    prod_col = next((c for c in df.columns if "PRODUTO" in c.upper() or "ITEM" in c.upper() or "DESCR" in c.upper()), "Produto")
    ovos_col = next((c for c in df.columns if "OVO" in c.upper() and "TRANS" not in c.upper() and "COD" not in c.upper() and "CÓD" not in c.upper()), "Ovos")
    prod_dt_col = next((c for c in df.columns if "DATA PRODU" in c.upper() or "DATA POSTURA" in c.upper() or ("PRODU" in c.upper() and "DATA" in c.upper())), "Data Produção")
    trans_dt_col = next((c for c in df.columns if ("DATA" in c.upper() and "TRANS" in c.upper()) or ("DATA" in c.upper() and "MOVIM" in c.upper())), "Data Transação")

    df = df.dropna(subset=[lote_col, ovos_col])
    
    df["Nº Lote"] = df[lote_col].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    df["Produto"] = df[prod_col].astype(str) if prod_col in df.columns else ""
    df["Ovos"] = pd.to_numeric(df[ovos_col], errors='coerce').fillna(0).astype(int)
    df["Data Produção"] = df[prod_dt_col].apply(normalizar_data_str) if prod_dt_col in df.columns else ""
    df["Data Transação"] = df[trans_dt_col].apply(normalizar_data_str) if trans_dt_col in df.columns else ""
    
    cols_agrupar = ["Nº Lote", "Produto", "Data Produção", "Data Transação"]
    df = df.groupby(cols_agrupar, dropna=False, as_index=False)["Ovos"].sum()
    return df

def parse_vendas_romaneios(file_bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    vendas = []
    for cliente_aba in xls.sheet_names:
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=cliente_aba)
        if len(df) > 8 and df.shape[1] >= 8:
            data_exp = normalizar_data_str(df.iloc[0, 8] if df.shape[1] > 8 else "")
            cliente_nome = df.iloc[1, 1]
            destino = df.iloc[3, 1]
            for r in range(9, len(df)):
                dt_post = df.iloc[r, 0]; lote = df.iloc[r, 1]; tipo = df.iloc[r, 4]; caixas = df.iloc[r, 6]; ovos = df.iloc[r, 7]
                if pd.notna(dt_post) and pd.notna(lote) and pd.notna(ovos):
                    try:
                        if int(float(ovos)) > 0: vendas.append({"Data_Expedicao": data_exp, "Cliente": str(cliente_nome), "Destino": str(destino), "Data_Postura": normalizar_data_str(dt_post), "Lote": str(int(float(lote))), "Tipo_Ovo": str(tipo), "Caixas_360": float(caixas), "Quantidade_Ovos": int(float(ovos))})
                    except Exception: pass
    return pd.DataFrame(vendas)

def parse_transferencias_incubatorio(file_bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    transf = []
    for aba in xls.sheet_names:
        if aba.upper() in ["DICAS", "VEICULOS"]: continue
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=aba)
        if len(df) > 8 and df.shape[1] >= 8:
            origem = df.iloc[0, 3]; destino = df.iloc[1, 3]
            data_exp = normalizar_data_str(df.iloc[4, 3])
            for r in range(9, len(df)):
                palete = df.iloc[r, 1]; lote = df.iloc[r, 2]; dt_post = df.iloc[r, 3]; qtd = df.iloc[r, 4]; categoria = df.iloc[r, 7]; obs = df.iloc[r, 8] if df.shape[1] > 8 else ""
                if pd.notna(lote) and pd.notna(qtd):
                    try:
                        if int(float(qtd)) > 0: transf.append({"Aba_Destino": aba, "Origem": str(origem), "Destino": str(destino), "Data_Expedicao": data_exp, "Palete": str(palete), "Lote": str(int(float(lote))), "Data_Postura": normalizar_data_str(dt_post), "Categoria": str(categoria), "Quantidade_Ovos": int(float(qtd)), "Observacao": str(obs) if pd.notna(obs) else ""})
                    except Exception: pass
    return pd.DataFrame(transf)

def parse_estoque_mtech(file_bytes):
    df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None)
    header_idx = None
    for i in range(min(20, len(df_raw))):
        row_vals = [str(x).upper() for x in df_raw.iloc[i].values]
        if any("NOME PRODUTO" in v for v in row_vals) and any("OVOS" in v for v in row_vals):
            header_idx = i
            break
    df = pd.read_excel(io.BytesIO(file_bytes), skiprows=header_idx) if header_idx is not None else pd.read_excel(io.BytesIO(file_bytes))
    df.columns = [str(c).strip().upper() for c in df.columns]
    prod_col = next((c for c in df.columns if "NOME PRODUTO" in c or "PRODUTO" in c), "NOME PRODUTO")
    qtd_col = next((c for c in df.columns if "OVOS" in c or "QTD" in c), "OVOS")
    
    res = pd.DataFrame()
    res["Produto_Original"] = df[prod_col].astype(str).str.strip()
    res["Quantidade"] = pd.to_numeric(df[qtd_col], errors="coerce").fillna(0).astype(int)
    return res[res["Quantidade"] != 0]

def parse_sankhya(file_bytes):
    df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None)
    header_idx = None
    for i in range(min(20, len(df_raw))):
        row_vals = [str(x).upper() for x in df_raw.iloc[i].values]
        if any("SKU" in v for v in row_vals) and any("ESTOQUE" in v for v in row_vals):
            header_idx = i
            break
    df = pd.read_excel(io.BytesIO(file_bytes), skiprows=header_idx) if header_idx is not None else pd.read_excel(io.BytesIO(file_bytes))
    df.columns = [str(c).strip().upper() for c in df.columns]
    prod_col = next((c for c in df.columns if "PRODUTO" in c or "DESC" in c), "PRODUTO")
    qtd_col = next((c for c in df.columns if "ESTOQUE" in c or "QTD" in c), "ESTOQUE")
    
    res = pd.DataFrame()
    res["Produto_Original"] = df[prod_col].astype(str).str.strip()
    res["Quantidade"] = pd.to_numeric(df[qtd_col], errors="coerce").fillna(0).astype(int)
    return res[res["Quantidade"] != 0]

def parse_estoque_cpo(file_bytes, resumo_sheet):
    df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=resumo_sheet, header=None)
    header_idx = None
    for i in range(min(20, len(df_raw))):
        row_vals = [str(x).upper() for x in df_raw.iloc[i].values]
        if any("CATEGORIA" in v for v in row_vals) and any("QUANTIDADE" in v or "TOTAL" in v for v in row_vals):
            header_idx = i
            break
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=resumo_sheet, skiprows=header_idx) if header_idx is not None else pd.read_excel(io.BytesIO(file_bytes), sheet_name=resumo_sheet)
    df.columns = [str(c).strip().upper() for c in df.columns]
    cat_col = next((c for c in df.columns if "CATEGORIA" in c), df.columns[0])
    qtd_col = next((c for c in df.columns if "QUANTIDADE" in c or "TOTAL" in c), df.columns[1])
    
    res = pd.DataFrame()
    res["Categoria_Original"] = df[cat_col].astype(str).str.strip()
    res["Quantidade"] = pd.to_numeric(df[qtd_col], errors="coerce").fillna(0).astype(int)
    res = res[~res["Categoria_Original"].str.upper().str.contains("TOTAL")]
    return res[res["Quantidade"] != 0]

def inspecionar_e_processar_conteudo(uploaded_file):
    nome_arquivo = getattr(uploaded_file, "name", "").lower()
    file_bytes = uploaded_file.getvalue()
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    abas_lower = [s.lower().strip() for s in xls.sheet_names]

    resumo_sheet = next((s for s in xls.sheet_names if s.upper().strip() == "RESUMO"), None)
    if resumo_sheet and not any("classifica" in s for s in abas_lower):
        df = parse_estoque_cpo(file_bytes, resumo_sheet)
        if not df.empty:
            DatabaseManager.limpar_tabela("df_est_cpo")
            n, dups, _ = DatabaseManager.salvar("df_est_cpo", df)
            return "Controle Físico de Estoque (CPO)", n, dups

    if any("classifica" in s for s in abas_lower) or "cpo" in nome_arquivo:
        df = parse_reclassificacao_cpo(file_bytes)
        n, dups, _ = DatabaseManager.salvar("df_reclass", df)
        return "Ficha de Reclassificação CPO", n, dups

    if any(s in ["dicas", "veiculos", "veículos"] for s in abas_lower):
        df = parse_transferencias_incubatorio(file_bytes)
        n, dups, _ = DatabaseManager.salvar("df_transf", df)
        return "Romaneio de Transferência OBI", n, dups

    df_amostra = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0)
    cols_upper = [str(c).strip().upper() for c in df_amostra.columns]
    texto_interno = df_amostra.iloc[:10].to_string().upper() if not df_amostra.empty else ""

    if "SKU" in texto_interno and ("ESTOQUE" in texto_interno or "CÓD. EMP." in texto_interno):
        df = parse_sankhya(file_bytes)
        DatabaseManager.limpar_tabela("df_est_sankhya")
        n, dups, _ = DatabaseManager.salvar("df_est_sankhya", df)
        return "Estoque Fiscal (Sankhya)", n, dups

    if "BIM ESTOQUE SALA OVOS" in texto_interno or ("NOME CLASSE OVO" in texto_interno and "IDADE MATRIZ" in texto_interno):
        df = parse_estoque_mtech(file_bytes)
        DatabaseManager.limpar_tabela("df_est_mtech")
        n, dups, _ = DatabaseManager.salvar("df_est_mtech", df)
        return "Estoque Origem (MTech)", n, dups

    is_reclassificacao_mtech = (
        any("TRANS OVO" in c for c in cols_upper) or
        "DÉBITO" in texto_interno or "CRÉDITO" in texto_interno or
        "DEBITO" in texto_interno or "CREDITO" in texto_interno
    )

    if is_reclassificacao_mtech:
        df = parse_reclassificacao_mtech(file_bytes)
        n, dups, _ = DatabaseManager.salvar("df_rec_mtech", df)
        return "Lançamento de Reclassificação", n, dups

    tem_lote = any("LOTE" in c for c in cols_upper)
    tem_fazenda = any("FAZENDA" in c for c in cols_upper)
    tem_transacao = any("TRANS" in c or "MOVIM" in c for c in cols_upper)

    indica_descarte = (
        any(k in nome_arquivo for k in ["descarte", "baixa", "perda", "quebra", "desc"]) or
        any(k in s for s in abas_lower for k in ["descarte", "baixa", "perda"]) or
        any("DESCARTE" in c or "BAIXA" in c for c in cols_upper) or
        "DESCARTE" in texto_interno
    )

    indica_producao = (
        any(k in nome_arquivo for k in ["produ", "coleta", "galp"]) or
        any(k in s for s in abas_lower for k in ["produ", "coleta", "galp"])
    )

    if indica_descarte or (tem_lote and tem_transacao and not indica_producao):
        df = parse_descarte_mtech(file_bytes)
        n, dups, _ = DatabaseManager.salvar("df_desc_mtech", df)
        return "Lançamento de Descartes", n, dups

    if tem_fazenda and tem_lote:
        df = parse_producao_erp(file_bytes)
        n, dups, _ = DatabaseManager.salvar("df_prod", df)
        return "Produção Diária (Núcleo)", n, dups

    if "PALETE" in texto_interno or any("PALETE" in c for c in cols_upper):
        df = parse_transferencias_incubatorio(file_bytes)
        n, dups, _ = DatabaseManager.salvar("df_transf", df)
        return "Romaneio de Transferência OBI", n, dups
    
    if "CLIENTE" in texto_interno or "CAIXAS" in texto_interno or any("CLIENTE" in c for c in cols_upper):
        df = parse_vendas_romaneios(file_bytes)
        n, dups, _ = DatabaseManager.salvar("df_vendas", df)
        return "Romaneio de Venda Externa", n, dups

    raise ValueError(f"Estrutura de ficheiro não reconhecida. Colunas: {cols_upper}")

# ==============================================================================
# 5. RENDERIZADORES DE TABELA HTML FORMATADA (COM SCROLL RESPONSIVO)
# ==============================================================================
def render_reclass_table_html(df):
    rows_html = []
    for idx, r in df.iterrows():
        bg = "#F9FBFA" if idx % 2 == 0 else "#FFFFFF"
        origem_color = "#1E40AF" if r["Origem"] == "Ninho Normal" else ("#B45309" if r["Origem"] == "Cama" else "#065F46")
        row_str = (
            f'<tr style="background-color: {bg}; border-bottom: 1px solid #E2E8F0; color: #1E293B; font-size: 14.5px;">'
            f'<td style="padding: 10px 12px; border-right: 1px solid #E2E8F0;">{r["Data_Postura"]}</td>'
            f'<td style="padding: 10px 12px; font-weight: 700; border-right: 2px solid #CBD5E1;">{r["Lote"]}</td>'
            f'<td style="padding: 10px 12px; font-weight: 700; color: {origem_color}; border-right: 1px solid #E2E8F0;">{r["Origem"]}</td>'
            f'<td style="padding: 10px 12px; font-weight: 700; text-align: right; border-right: 2px solid #CBD5E1;">{r["Saldo_Inicial"]:,}'.replace(",", ".") + f'</td>'
            f'<td style="padding: 10px 12px; text-align: right; border-right: 1px solid #E2E8F0; font-weight: 600;">{r["Sujo"]:,}'.replace(",", ".") + f'</td>'
            f'<td style="padding: 10px 12px; text-align: right; border-right: 1px solid #E2E8F0; font-weight: 600;">{r["Deformado"]:,}'.replace(",", ".") + f'</td>'
            f'<td style="padding: 10px 12px; text-align: right; border-right: 1px solid #E2E8F0; font-weight: 600;">{r["Ninho_Sujo"]:,}'.replace(",", ".") + f'</td>'
            f'<td style="padding: 10px 12px; text-align: right; border-right: 1px solid #E2E8F0; font-weight: 600;">{r["Industrial"]:,}'.replace(",", ".") + f'</td>'
            f'<td style="padding: 10px 12px; text-align: right; border-right: 1px solid #E2E8F0; font-weight: 600;">{r["Duas_Gemas"]:,}'.replace(",", ".") + f'</td>'
            f'<td style="padding: 10px 12px; text-align: right; border-right: 1px solid #E2E8F0; font-weight: 600;">{r["Ovo_Limpo"]:,}'.replace(",", ".") + f'</td>'
            f'<td style="padding: 10px 12px; font-weight: 700; text-align: right; border-right: 2px solid #CBD5E1; color: #1E293B;">{r["Trincado"]:,}'.replace(",", ".") + f'</td>'
            f'<td style="padding: 10px 12px; font-weight: 800; text-align: right; background-color: #F1F5F9; border-right: 1px solid #E2E8F0;">{r["Total_Destinos"]:,}'.replace(",", ".") + f'</td>'
            f'<td style="padding: 10px 12px; font-weight: 800; text-align: right; background-color: ' + ("#FEF2F2" if r["Saldo_Final"] > 0 else "#ECFDF5") + '; color: ' + ("#991B1B" if r["Saldo_Final"] > 0 else "#065F46") + f';">{r["Saldo_Final"]:,}'.replace(",", ".") + f'</td>'
            f'</tr>'
        )
        rows_html.append(row_str)

    t_saldo, t_sujo, t_def, t_ns = df["Saldo_Inicial"].sum(), df["Sujo"].sum(), df["Deformado"].sum(), df["Ninho_Sujo"].sum()
    t_ind, t_dg, t_limpo, t_trinc = df["Industrial"].sum(), df["Duas_Gemas"].sum(), df["Ovo_Limpo"].sum(), df["Trincado"].sum()
    t_dest, t_fim = df["Total_Destinos"].sum(), df["Saldo_Final"].sum()

    row_total = (
        f'<tr style="background-color: #1E293B; color: #FFFFFF; font-weight: 800; font-size: 15px;">'
        f'<td colspan="3" style="padding: 12px; text-align: right; border-right: 2px solid #CBD5E1;">TOTAL GERAL DA SELEÇÃO</td>'
        f'<td style="padding: 12px; text-align: right; border-right: 2px solid #CBD5E1;">{t_saldo:,}</td>'
        f'<td style="padding: 12px; text-align: right; border-right: 1px solid #475569;">{t_sujo:,}</td>'
        f'<td style="padding: 12px; text-align: right; border-right: 1px solid #475569;">{t_def:,}</td>'
        f'<td style="padding: 12px; text-align: right; border-right: 1px solid #475569;">{t_ns:,}</td>'
        f'<td style="padding: 12px; text-align: right; border-right: 1px solid #475569;">{t_ind:,}</td>'
        f'<td style="padding: 12px; text-align: right; border-right: 1px solid #475569;">{t_dg:,}</td>'
        f'<td style="padding: 12px; text-align: right; border-right: 1px solid #475569;">{t_limpo:,}</td>'
        f'<td style="padding: 12px; text-align: right; border-right: 2px solid #CBD5E1;">{t_trinc:,}</td>'
        f'<td style="padding: 12px; text-align: right; border-right: 1px solid #475569;">{t_dest:,}</td>'
        f'<td style="padding: 12px; text-align: right;">{t_fim:,}</td>'
        f'</tr>'
    ).replace(",", ".")

    table_html = (
        '<div class="table-container-responsive">'
        '<table style="min-width: 1250px; width: 100%; border-collapse: collapse; text-align: center;">'
        '<thead>'
        '<tr style="color: white; font-weight: 700; font-size: 13px;">'
        '<th colspan="2" style="background-color: #334155; padding: 10px; border-right: 2px solid white;">IDENTIFICAÇÃO</th>'
        '<th style="background-color: #1E3A8A; padding: 10px; border-right: 2px solid white;">ORIGEM</th>'
        '<th style="background-color: #1E3A8A; padding: 10px; border-right: 2px solid white;">SALDO INI</th>'
        '<th colspan="7" style="background-color: #14532D; padding: 10px; border-right: 2px solid white;">DESTINO COMERCIAL</th>'
        '<th colspan="2" style="background-color: #475569; padding: 10px;">CONFERÊNCIA</th>'
        '</tr>'
        '<tr style="color: white; font-weight: 600; font-size: 14px;">'
        '<th style="background-color: #475569; padding: 10px;">Data Postura</th>'
        '<th style="background-color: #475569; padding: 10px; border-right: 2px solid white;">Lote</th>'
        '<th style="background-color: #2563EB; padding: 10px;">Origem</th>'
        '<th style="background-color: #2563EB; padding: 10px; border-right: 2px solid white;">Saldo Inicial</th>'
        '<th style="background-color: #166534; padding: 10px;">Sujo</th>'
        '<th style="background-color: #166534; padding: 10px;">Deformado</th>'
        '<th style="background-color: #166534; padding: 10px;">Ninho Sujo</th>'
        '<th style="background-color: #166534; padding: 10px;">Industrial</th>'
        '<th style="background-color: #166534; padding: 10px;">Duas Gemas</th>'
        '<th style="background-color: #166534; padding: 10px;">Ovo Limpo</th>'
        '<th style="background-color: #166534; padding: 10px; border-right: 2px solid white;">Trincado</th>'
        '<th style="background-color: #475569; padding: 10px;">Total Destinos</th>'
        '<th style="background-color: #475569; padding: 10px;">Saldo Final</th>'
        '</tr>'
        '</thead>'
        f'<tbody>{"".join(rows_html) + row_total}</tbody>'
        '</table></div>'
    )
    return table_html

def render_descarte_table_html(df):
    rows_html = []
    for idx, r in df.iterrows():
        bg = "#F9FBFA" if idx % 2 == 0 else "#FFFFFF"
        rows_html.append(
            f'<tr style="background-color: {bg}; border-bottom: 1px solid #FEE2E2; color: #1E293B; font-size: 14.5px;">'
            f'<td style="padding: 10px 12px; border-right: 1px solid #FEE2E2;">{r["Data_Postura"]}</td>'
            f'<td style="padding: 10px 12px; font-weight: 700; border-right: 2px solid #CBD5E1;">{r["Lote"]}</td>'
            f'<td style="padding: 10px 12px; text-align: right; border-right: 1px solid #FEE2E2; font-weight: 800; color: #1E3A8A;">{r["Ninho Normal"]:,}</td>'
            f'<td style="padding: 10px 12px; text-align: right; border-right: 1px solid #FEE2E2; font-weight: 600;">{r["Cama"]:,}</td>'
            f'<td style="padding: 10px 12px; text-align: right; border-right: 1px solid #FEE2E2; font-weight: 600;">{r["Ninho Sujo"]:,}</td>'
            f'<td style="padding: 10px 12px; text-align: right; border-right: 1px solid #FEE2E2; font-weight: 600;">{r["Deformado"]:,}</td>'
            f'<td style="padding: 10px 12px; text-align: right; border-right: 1px solid #FEE2E2; font-weight: 600;">{r["Sujo Granja"]:,}</td>'
            f'<td style="padding: 10px 12px; text-align: right; border-right: 1px solid #FEE2E2; font-weight: 600;">{r["Industrial"]:,}</td>'
            f'<td style="padding: 10px 12px; text-align: right; border-right: 1px solid #FEE2E2; font-weight: 600;">{r["Duas Gemas"]:,}</td>'
            f'<td style="padding: 10px 12px; text-align: right; border-right: 1px solid #FEE2E2; font-weight: 800; color: #1E293B;">{r["Trincado"]:,}</td>'
            f'<td style="padding: 10px 12px; font-weight: 800; text-align: right; color: #991B1B; border-left: 2px solid #CBD5E1;">{r["TOTAL BAIXA"]:,}</td>'
            f'</tr>'.replace(",", ".")
        )
    
    t_nn, t_cama, t_ns = df["Ninho Normal"].sum(), df["Cama"].sum(), df["Ninho Sujo"].sum()
    t_def, t_sujog, t_ind = df["Deformado"].sum(), df["Sujo Granja"].sum(), df["Industrial"].sum()
    t_dg, t_trinc, t_total = df["Duas Gemas"].sum(), df["Trincado"].sum(), df["TOTAL BAIXA"].sum()
    
    row_total = (
        f'<tr style="background-color: #334155; color: #FFFFFF; font-weight: 800; font-size: 15px;">'
        f'<td colspan="2" style="padding: 12px; text-align: right; border-right: 2px solid #CBD5E1;">TOTAIS GERAIS DA SELEÇÃO</td>'
        f'<td style="padding: 12px; text-align: right; border-right: 1px solid #64748B;">{t_nn:,}</td>'
        f'<td style="padding: 12px; text-align: right; border-right: 1px solid #64748B;">{t_cama:,}</td>'
        f'<td style="padding: 12px; text-align: right; border-right: 1px solid #64748B;">{t_ns:,}</td>'
        f'<td style="padding: 12px; text-align: right; border-right: 1px solid #64748B;">{t_def:,}</td>'
        f'<td style="padding: 12px; text-align: right; border-right: 1px solid #64748B;">{t_sujog:,}</td>'
        f'<td style="padding: 12px; text-align: right; border-right: 1px solid #64748B;">{t_ind:,}</td>'
        f'<td style="padding: 12px; text-align: right; border-right: 1px solid #64748B;">{t_dg:,}</td>'
        f'<td style="padding: 12px; text-align: right; border-right: 1px solid #64748B; color: #FCA5A5;">{t_trinc:,}</td>'
        f'<td style="padding: 12px; text-align: right; border-left: 2px solid #CBD5E1;">{t_total:,}</td>'
        f'</tr>'.replace(",", ".")
    )
    
    table_html = (
        '<div class="table-container-responsive">'
        '<table style="min-width: 1000px; width: 100%; border-collapse: collapse; text-align: center;">'
        '<thead>'
        '<tr style="color: white; font-weight: 700; font-size: 13px; background-color: #334155;">'
        '<th style="padding: 10px; border-right: 1px solid #64748B;">Data Postura</th>'
        '<th style="padding: 10px; border-right: 2px solid white;">Lote</th>'
        '<th style="padding: 10px; border-right: 1px solid #64748B;">Ninho Normal</th>'
        '<th style="padding: 10px; border-right: 1px solid #64748B;">Cama</th>'
        '<th style="padding: 10px; border-right: 1px solid #64748B;">Ninho Sujo</th>'
        '<th style="padding: 10px; border-right: 1px solid #64748B;">Deformado</th>'
        '<th style="padding: 10px; border-right: 1px solid #64748B;">Sujo Granja</th>'
        '<th style="padding: 10px; border-right: 1px solid #64748B;">Industrial</th>'
        '<th style="padding: 10px; border-right: 1px solid #64748B;">Duas Gemas</th>'
        '<th style="padding: 10px; border-right: 2px solid white; background-color: #166534;">Trincado</th>'
        '<th style="padding: 10px; background-color: #991B1B;">TOTAL BAIXA</th>'
        '</tr>'
        '</thead>'
        f'<tbody>{"".join(rows_html) + row_total}</tbody>'
        '</table></div>'
    )
    return table_html

def render_estoque_sankhya_mtech_html(df):
    rows_html = []
    for idx, r in df.iterrows():
        bg = "#F9FBFA" if idx % 2 == 0 else "#FFFFFF"
        div_val = r["Divergencia Qtd"]
        div_color = "#166534" if div_val == 0 else "#991B1B"
        div_bg = "transparent" if div_val == 0 else "#FEF2F2"
        
        row_str = (
            f'<tr style="background-color: {bg}; border-bottom: 1px solid #E2E8F0; color: #1E293B; font-size: 14.5px;">'
            f'<td style="padding: 10px 12px; font-weight: 700; border-right: 1px solid #E2E8F0; text-align: left;">{r["Nome Padrao"]}</td>'
            f'<td style="padding: 10px 12px; text-align: right; border-right: 1px solid #E2E8F0;">{r["Qtd MTech"]:,}</td>'
            f'<td style="padding: 10px 12px; text-align: right; border-right: 2px solid #CBD5E1;">{r["Qtd Sankhya"]:,}</td>'
            f'<td style="padding: 10px 12px; font-weight: 800; text-align: right; background-color: {div_bg}; color: {div_color};">{div_val:,}</td>'
            f'</tr>'
        )
        rows_html.append(row_str.replace(",", "."))
        
    t_mtech = df["Qtd MTech"].sum()
    t_sank = df["Qtd Sankhya"].sum()
    t_div = df["Divergencia Qtd"].sum()
    
    row_total = (
        f'<tr style="background-color: #1E293B; color: #FFFFFF; font-weight: 800; font-size: 15px;">'
        f'<td style="padding: 12px; text-align: left; border-right: 1px solid #CBD5E1;">Total</td>'
        f'<td style="padding: 12px; text-align: right; border-right: 1px solid #475569;">{t_mtech:,}</td>'
        f'<td style="padding: 12px; text-align: right; border-right: 2px solid #CBD5E1;">{t_sank:,}</td>'
        f'<td style="padding: 12px; text-align: right;">{t_div:,}</td>'
        f'</tr>'
    ).replace(",", ".")

    table_html = (
        '<div class="table-container-responsive">'
        '<table style="min-width: 600px; width: 100%; border-collapse: collapse; text-align: center;">'
        '<thead>'
        '<tr style="color: white; font-weight: 700; font-size: 13px;">'
        '<th style="background-color: #334155; padding: 10px; border-right: 2px solid white; text-align: left;">Nome Padrao</th>'
        '<th style="background-color: #1E3A8A; padding: 10px; border-right: 1px solid white;">Qtd Ovos MTech</th>'
        '<th style="background-color: #1E3A8A; padding: 10px; border-right: 2px solid white;">Qtd Sankhya</th>'
        '<th style="background-color: #475569; padding: 10px;">Divergencia Qtd</th>'
        '</tr>'
        '</thead>'
        f'<tbody>{"".join(rows_html) + row_total}</tbody>'
        '</table></div>'
    )
    return table_html

def render_estoque_cpo_sankhya_html(df):
    rows_html = []
    for idx, r in df.iterrows():
        bg = "#F9FBFA" if idx % 2 == 0 else "#FFFFFF"
        div_val = r["Divergencia CPO"]
        div_color = "#166534" if div_val == 0 else "#991B1B"
        div_bg = "transparent" if div_val == 0 else "#FEF2F2"
        
        row_str = (
            f'<tr style="background-color: {bg}; border-bottom: 1px solid #E2E8F0; color: #1E293B; font-size: 14.5px;">'
            f'<td style="padding: 10px 12px; font-weight: 700; border-right: 1px solid #E2E8F0; text-align: left;">{r["Nome Padrao"]}</td>'
            f'<td style="padding: 10px 12px; text-align: right; border-right: 1px solid #E2E8F0;">{r["Qtd Sankhya"]:,}</td>'
            f'<td style="padding: 10px 12px; text-align: right; border-right: 2px solid #CBD5E1;">{r["Qtd CPO"]:,}</td>'
            f'<td style="padding: 10px 12px; font-weight: 800; text-align: right; background-color: {div_bg}; color: {div_color};">{div_val:,}</td>'
            f'</tr>'
        )
        rows_html.append(row_str.replace(",", "."))
        
    t_sank = df["Qtd Sankhya"].sum()
    t_cpo = df["Qtd CPO"].sum()
    t_div = df["Divergencia CPO"].sum()
    
    row_total = (
        f'<tr style="background-color: #1E293B; color: #FFFFFF; font-weight: 800; font-size: 15px;">'
        f'<td style="padding: 12px; text-align: left; border-right: 1px solid #CBD5E1;">Total</td>'
        f'<td style="padding: 12px; text-align: right; border-right: 1px solid #475569;">{t_sank:,}</td>'
        f'<td style="padding: 12px; text-align: right; border-right: 2px solid #CBD5E1;">{t_cpo:,}</td>'
        f'<td style="padding: 12px; text-align: right;">{t_div:,}</td>'
        f'</tr>'
    ).replace(",", ".")

    table_html = (
        '<div class="table-container-responsive">'
        '<table style="min-width: 600px; width: 100%; border-collapse: collapse; text-align: center;">'
        '<thead>'
        '<tr style="color: white; font-weight: 700; font-size: 13px;">'
        '<th style="background-color: #334155; padding: 10px; border-right: 2px solid white; text-align: left;">Nome Padrao</th>'
        '<th style="background-color: #1E3A8A; padding: 10px; border-right: 1px solid white;">Qtd Sankhya</th>'
        '<th style="background-color: #1E3A8A; padding: 10px; border-right: 2px solid white;">Qtd Ovos CPO</th>'
        '<th style="background-color: #475569; padding: 10px;">Divergencia CPO</th>'
        '</tr>'
        '</thead>'
        f'<tbody>{"".join(rows_html) + row_total}</tbody>'
        '</table></div>'
    )
    return table_html

# ==============================================================================
# 6. MENU LATERAL EXECUTIVO
# ==============================================================================
def obter_logo_b64():
    return "https://media.licdn.com/dms/image/v2/D4D0BAQE8aP53G2wHSA/company-logo_200_200/company-logo_200_200/0/1691157833055/natto_alimentos_logo?e=2147483647&v=beta&t=HXY2xK7O4JzVz0j5Z_WjB8OqC0fWc_8X9K3Lz8EwPjQ"

st.sidebar.markdown(f"""
<div class="brand-box">
    <img src="{obter_logo_b64()}" style="width: 48px; height: 48px; border-radius: 8px; object-fit: cover; border: 1px solid rgba(255,255,255,0.2);">
    <div>
        <p class="brand-title">Granja - 117</p>
        <span class="brand-sub">Controle Operacional</span>
    </div>
</div>
""", unsafe_allow_html=True)

st.sidebar.markdown('<div class="nav-label">Menu de Navegação</div>', unsafe_allow_html=True)

tela = st.sidebar.radio(
    "Navegação:",
    [
        "📊 Dashboard",
        "📥 Central de Carga de Arquivos",
        "🏭 Produção Diária (Núcleo)",
        "📋 Validação de Reclassificação",
        "🗑️ Validação de Descartes",
        "🚚 Validação Estoque (MTech x Sankhya x CPO)"
    ]
)

st.sidebar.markdown('<div class="nav-label">Cotação & Parâmetros</div>', unsafe_allow_html=True)

preco_unitario = st.sidebar.number_input("Preço por Unidade (R$):", min_value=0.000, value=2.380, step=0.010, format="%.3f")
valor_caixa_calc = preco_unitario * 360.0
val_cx_str = f"R$ {valor_caixa_calc:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

st.sidebar.markdown(f"""
<div class="cotacao-sidebar-card">
    <div class="cotacao-sidebar-tag">Caixa Fechada (360 und)</div>
    <div class="cotacao-sidebar-val">{val_cx_str}</div>
</div>
""", unsafe_allow_html=True)

# ==============================================================================
# 7. CONTROLADORES DOS ECRÃS
# ==============================================================================

if tela == "📊 Dashboard":
    st.markdown("""
    <div class="top-header-bar">
        <div class="top-header-title">
            <span>Visão Consolidada • Base SQLite Ativa</span>
            <h1>Painel Geral de Operações & Faturamento</h1>
        </div>
        <div class="top-status-badge">● Base Local Conectada</div>
    </div>
    """, unsafe_allow_html=True)

    df_prod = DatabaseManager.obter("df_prod")
    df_vendas = DatabaseManager.obter("df_vendas")
    df_transf = DatabaseManager.obter("df_transf")
    df_sankhya = DatabaseManager.obter("df_est_sankhya")
    df_cpo = DatabaseManager.obter("df_est_cpo")
    df_reclass = DatabaseManager.obter("df_reclass")

    # Média de Produção e Lotes Ativos
    if not df_prod.empty:
        df_prod["Núcleo_Lote"] = df_prod["Nome Fazenda"].astype(str) + "-" + df_prod["Nº Lote"].astype(str)
        lotes_ativos = df_prod["Núcleo_Lote"].nunique()
        media_prod_diaria = df_prod.groupby("Data Produção")["Ovos"].sum().mean()
    else:
        lotes_ativos = 0
        media_prod_diaria = 0

    # Lógica de Filtro Comercial/OBI com Base Padrão Universal para calcular o Estoque Total
    saldo_comercial = 0
    saldo_obi = 0
    sankhya_nn = 0

    if not df_sankhya.empty:
        df_s_temp = df_sankhya.copy()
        df_s_temp["Nome Padrao"] = df_s_temp["Produto_Original"].str.strip().map(MAP_SANKHYA).fillna("Outros")
        
        # Puxa o NN estritamente do Sankhya para a Projeção
        sankhya_nn = df_s_temp[df_s_temp["Nome Padrao"] == "01-OVO NINHO NORMAL"]["Quantidade"].sum()
        
        if df_cpo.empty:
            saldo_obi = df_s_temp[df_s_temp["Nome Padrao"].isin(PADRAO_OBI)]["Quantidade"].sum()
            saldo_comercial = df_s_temp[df_s_temp["Nome Padrao"].isin(PADRAO_COMERCIAL)]["Quantidade"].sum()

    if not df_cpo.empty:
        df_c_temp = df_cpo.copy()
        df_c_temp["Nome Padrao"] = df_c_temp["Categoria_Original"].str.strip().map(MAP_CPO).fillna("Outros")
        saldo_obi = df_c_temp[df_c_temp["Nome Padrao"].isin(PADRAO_OBI)]["Quantidade"].sum()
        saldo_comercial = df_c_temp[df_c_temp["Nome Padrao"].isin(PADRAO_COMERCIAL)]["Quantidade"].sum()

    estoque_total = saldo_comercial + saldo_obi

    kpis_top = f"""
    <div class="kpi-container" style="grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));">
        <div class="kpi-card">
            <div class="kpi-icon-bubble">📦</div>
            <div>
                <p class="kpi-title">Estoque Total</p>
                <p class="kpi-value">{estoque_total:,.0f}</p>
                <p class="kpi-sub" style="color: #64748B;">Comercial + OBI</p>
            </div>
        </div>
        <div class="kpi-card">
            <div class="kpi-icon-bubble">🛒</div>
            <div>
                <p class="kpi-title">Saldo (Comercial)</p>
                <p class="kpi-value">{saldo_comercial:,.0f}</p>
                <p class="kpi-sub" style="color: #64748B;">Filtro Aplicado</p>
            </div>
        </div>
        <div class="kpi-card">
            <div class="kpi-icon-bubble">🐣</div>
            <div>
                <p class="kpi-title">Saldo (OBI)</p>
                <p class="kpi-value">{saldo_obi:,.0f}</p>
                <p class="kpi-sub" style="color: #64748B;">Filtro Aplicado</p>
            </div>
        </div>
        <div class="kpi-card">
            <div class="kpi-icon-bubble">📈</div>
            <div>
                <p class="kpi-title">Média Prod. Diária</p>
                <p class="kpi-value">{media_prod_diaria:,.0f}</p>
                <p class="kpi-sub" style="color: #64748B;">Ovos / Dia</p>
            </div>
        </div>
        <div class="kpi-card">
            <div class="kpi-icon-bubble">🏷️</div>
            <div>
                <p class="kpi-title">Núcleos Ativos</p>
                <p class="kpi-value">{lotes_ativos}</p>
                <p class="kpi-sub" style="color: #64748B;">Em Produção</p>
            </div>
        </div>
    </div>
    """.replace(",", "X").replace(".", ",").replace("X", ".")
    st.markdown(kpis_top, unsafe_allow_html=True)

    # 2. Tabela Matriz de Produção
    st.markdown('<div class="saas-panel">', unsafe_allow_html=True)
    st.subheader("📊 Detalhamento de Produção por Núcleo e Lote (Matriz Dinâmica)")
    if not df_prod.empty:
        df_pivot = pd.pivot_table(
            df_prod, 
            values="Ovos", 
            index="Nome Produto", 
            columns="Núcleo_Lote", 
            aggfunc="sum", 
            fill_value=0,
            margins=True,
            margins_name="Total"
        )
        df_pivot = df_pivot.fillna(0).astype(int)
        
        df_pivot.columns = [encurtar_nome_coluna(col) if col != 'Total' else col for col in df_pivot.columns]
        
        styled_df = df_pivot.style.set_properties(**{'color': '#0F172A', 'font-weight': '600'}) \
                          .apply(lambda x: ['font-weight: 800; color: #0F172A; background-color: #E2E8F0' if x.name == 'Total' else '' for i in x], axis=1) \
                          .apply(lambda x: ['font-weight: 800; color: #0F172A; background-color: #F8FAFC' if i == len(x)-1 else '' for i, v in enumerate(x)], axis=0) \
                          .format(lambda x: f"{x:,.0f}".replace(",", "."))
        
        st.dataframe(styled_df, use_container_width=True)
    else:
        show_info("Nenhum dado de produção diária gravado no banco.")
    st.markdown('</div>', unsafe_allow_html=True)

    # 3. Projeções Sankhya
    caixas_nn = sankhya_nn // 360
    ovos_embalados = caixas_nn * 360
    faturamento_previsto = ovos_embalados * preco_unitario

    st.markdown('<div class="saas-panel">', unsafe_allow_html=True)
    st.subheader("📈 Projeções Baseadas no Estoque Fiscal (Sankhya)")
    kpis_proj = f"""
    <div class="kpi-container" style="grid-template-columns: 1fr 1fr;">
        <div class="kpi-card">
            <div class="kpi-icon-bubble">📦</div>
            <div>
                <p class="kpi-title">Projeção P/ Venda (NN)</p>
                <p class="kpi-value">{caixas_nn:,.0f} caixas</p>
                <p class="kpi-sub" style="color: #64748B;">{sankhya_nn:,.0f} ovos NN -> Cx 360</p>
            </div>
        </div>
        <div class="kpi-card">
            <div class="kpi-icon-bubble">📈</div>
            <div>
                <p class="kpi-title">Previsão Faturamento (NN)</p>
                <p class="kpi-value">R$ {faturamento_previsto:,.2f}</p>
                <p class="kpi-sub" style="color: #64748B;">Base: {ovos_embalados:,.0f} ovos embalados</p>
            </div>
        </div>
    </div>
    """.replace(",", "X").replace(".", ",").replace("X", ".")
    st.markdown(kpis_proj, unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # 4. Saídas Registradas
    st.markdown('<div class="saas-panel">', unsafe_allow_html=True)
    st.subheader("📤 Saídas Registradas")
    
    tipo_saida = st.radio("Selecione a Visão de Saídas:", ["Comercial (Vendas Externas)", "Incubatório (Transferências OBI)"], horizontal=True, label_visibility="collapsed")
    
    if tipo_saida == "Comercial (Vendas Externas)":
        tot_ovos_vendidos = df_vendas["Quantidade_Ovos"].sum() if not df_vendas.empty else 0
        tot_caixas_venda = df_vendas["Caixas_360"].sum() if not df_vendas.empty else 0
        faturamento_vendas = tot_ovos_vendidos * preco_unitario
        
        kpis_saida = f"""
        <div class="kpi-container" style="grid-template-columns: 1fr 1fr;">
            <div class="kpi-card">
                <div class="kpi-icon-bubble">🚚</div>
                <div>
                    <p class="kpi-title">Ovos Vendidos</p>
                    <p class="kpi-value">{tot_ovos_vendidos:,.0f}</p>
                    <p class="kpi-sub" style="color: #64748B;">{tot_caixas_venda:,.1f} caixas (360 un)</p>
                </div>
            </div>
            <div class="kpi-card">
                <div class="kpi-icon-bubble">💰</div>
                <div>
                    <p class="kpi-title">Faturamento Realizado</p>
                    <p class="kpi-value">R$ {faturamento_vendas:,.2f}</p>
                    <p class="kpi-sub" style="color: #64748B;">Preço Und: R$ {preco_unitario:.3f}</p>
                </div>
            </div>
        </div>
        """.replace(",", "X").replace(".", ",").replace("X", ".")
        st.markdown(kpis_saida, unsafe_allow_html=True)
        
        if not df_vendas.empty:
            venda_cli = df_vendas.groupby("Cliente")[["Caixas_360", "Quantidade_Ovos"]].sum().reset_index()
            venda_cli["Faturamento (R$)"] = venda_cli["Quantidade_Ovos"] * preco_unitario
            st.dataframe(venda_cli, use_container_width=True, hide_index=True)
        else:
            show_info("Nenhuma venda registada no histórico.")
            
    else:
        tot_ovos_obi = df_transf["Quantidade_Ovos"].sum() if not df_transf.empty else 0
        tot_caixas_obi = tot_ovos_obi / 360
        
        kpis_saida = f"""
        <div class="kpi-container" style="grid-template-columns: 1fr 1fr;">
            <div class="kpi-card">
                <div class="kpi-icon-bubble">🔄</div>
                <div>
                    <p class="kpi-title">Ovos Transferidos (OBI)</p>
                    <p class="kpi-value">{tot_ovos_obi:,.0f}</p>
                    <p class="kpi-sub" style="color: #64748B;">Centros de Incubação</p>
                </div>
            </div>
            <div class="kpi-card">
                <div class="kpi-icon-bubble">📦</div>
                <div>
                    <p class="kpi-title">Volume de Transferência</p>
                    <p class="kpi-value">{tot_caixas_obi:,.1f} caixas</p>
                    <p class="kpi-sub" style="color: #64748B;">Lotes Ninho Normal</p>
                </div>
            </div>
        </div>
        """.replace(",", "X").replace(".", ",").replace("X", ".")
        st.markdown(kpis_saida, unsafe_allow_html=True)
        
        if not df_transf.empty:
            transf_dest = df_transf.groupby("Destino")[["Quantidade_Ovos"]].sum().reset_index()
            st.dataframe(transf_dest, use_container_width=True, hide_index=True)
        else:
            show_info("Nenhuma transferência registada no histórico.")
            
    st.markdown('</div>', unsafe_allow_html=True)

elif tela == "📥 Central de Carga de Arquivos":
    st.markdown("""
    <div class="top-header-bar">
        <div class="top-header-title">
            <span>Leitura por Conteúdo Interno & Persistência SQLite</span>
            <h1>Central de Carga & Banco de Dados</h1>
        </div>
        <div class="top-status-badge">💾 SQLite: historico_ovos_117.db</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="saas-panel" style="border-top: 4px solid #16A34A;">', unsafe_allow_html=True)
    st.subheader("🚀 Importador Inteligente (Leitura Automática)")
    st.markdown(
        '<div style="background-color: #F0FDF4; border: 1px solid #DCFCE7; padding: 12px 16px; border-radius: 8px; color: #166534; font-size: 14.5px; margin-bottom: 16px;">'
        '<b>Arraste qualquer folha de cálculo abaixo:</b> O sistema identifica automaticamente o tipo do documento e guarda na base de dados de forma segura sem duplicar registos.'
        '</div>',
        unsafe_allow_html=True
    )

    arquivos_unificados = st.file_uploader(
        "Selecione ou arraste todas as planilhas (Fichas CPO, Romaneios, ERP Produção, Reclassificação ou Descartes):",
        type=["xls", "xlsx"],
        accept_multiple_files=True,
        key="uploader_inteligente"
    )

    if arquivos_unificados:
        for arq in arquivos_unificados:
            try:
                tipo_identificado, qtd_reg, qtd_dups = inspecionar_e_processar_conteudo(arq)
                if qtd_dups > 0:
                    show_warning(f"⚠️ **Alerta de Registros Repetidos:** Na planilha de **{tipo_identificado}**, foram identificadas {qtd_dups} linha(s) com Lote/Data que já constavam no banco de dados. Verifique se esta ficha já havia sido lançada anteriormente!")
                else:
                    show_success(f"Conteúdo identificado como **{tipo_identificado}** — {qtd_reg} registos sincronizados na base!")
            except Exception as e:
                show_error(f"Não foi possível identificar a estrutura interna do ficheiro: {e}")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="saas-panel">', unsafe_allow_html=True)
    st.subheader("🗄️ Estado do Histórico Guardado na Base Local (SQLite)")
    
    resumo_db = [
        {"Tabela na Base": "Fichas de Reclassificação & Perdas (CPO)", "Código": "df_reclass", "Registros Gravados": len(DatabaseManager.obter("df_reclass"))},
        {"Tabela na Base": "MTech - Lançamentos de Reclassificação", "Código": "df_rec_mtech", "Registros Gravados": len(DatabaseManager.obter("df_rec_mtech"))},
        {"Tabela na Base": "MTech - Lançamentos de Descartes", "Código": "df_desc_mtech", "Registros Gravados": len(DatabaseManager.obter("df_desc_mtech"))},
        {"Tabela na Base": "MTech - Produção Diária (Núcleo)", "Código": "df_prod", "Registros Gravados": len(DatabaseManager.obter("df_prod"))},
        {"Tabela na Base": "Romaneios de Venda Externa", "Código": "df_vendas", "Registros Gravados": len(DatabaseManager.obter("df_vendas"))},
        {"Tabela na Base": "Romaneios de Transferência OBI", "Código": "df_transf", "Registros Gravados": len(DatabaseManager.obter("df_transf"))},
        {"Tabela na Base": "Fotografia de Estoque Físico (CPO)", "Código": "df_est_cpo", "Registros Gravados": len(DatabaseManager.obter("df_est_cpo"))},
        {"Tabela na Base": "Fotografia de Estoque Fiscal (Sankhya)", "Código": "df_est_sankhya", "Registros Gravados": len(DatabaseManager.obter("df_est_sankhya"))},
        {"Tabela na Base": "Fotografia de Estoque Origem (MTech)", "Código": "df_est_mtech", "Registros Gravados": len(DatabaseManager.obter("df_est_mtech"))},
    ]
    st.dataframe(pd.DataFrame(resumo_db), use_container_width=True, hide_index=True)

    with st.expander("🛠️ Manutenção da Base de Dados (Limpar Tabelas ou Baixar Cópia de Segurança)"):
        c_db1, c_db2 = st.columns(2)
        with c_db1:
            tab_para_limpar = st.selectbox(
                "Selecione uma base caso deseje limpar os seus registos:",
                ["Selecione...", "df_reclass", "df_rec_mtech", "df_desc_mtech", "df_prod", "df_vendas", "df_transf", "df_est_cpo", "df_est_sankhya", "df_est_mtech", "TODAS AS TABELAS"]
            )
            if st.button("🗑️ Limpar Base Selecionada"):
                if tab_para_limpar == "TODAS AS TABELAS":
                    for t in ["df_reclass", "df_rec_mtech", "df_desc_mtech", "df_prod", "df_vendas", "df_transf", "df_est_cpo", "df_est_sankhya", "df_est_mtech"]:
                        DatabaseManager.limpar_tabela(t)
                    show_success("Todas as tabelas foram reinicializadas a zeros na base local.")
                    st.rerun()
                elif tab_para_limpar != "Selecione...":
                    DatabaseManager.limpar_tabela(tab_para_limpar)
                    show_success(f"Tabela {tab_para_limpar} limpa com sucesso!")
                    st.rerun()
        with c_db2:
            if os.path.exists(DB_PATH):
                with open(DB_PATH, "rb") as f_db:
                    st.download_button(
                        "💾 Descarregar Cópia de Segurança (.db)",
                        data=f_db.read(),
                        file_name="historico_ovos_117.db",
                        mime="application/octet-stream"
                    )
    st.markdown('</div>', unsafe_allow_html=True)

elif tela == "🏭 Produção Diária (Núcleo)":
    st.markdown('<div class="top-header-bar"><div class="top-header-title"><span>Visão Núcleos</span><h1>Apontamentos de Produção Diária (MTech ERP)</h1></div></div>', unsafe_allow_html=True)
    df_p = DatabaseManager.obter("df_prod")
    if df_p.empty: show_warning("Nenhum dado de produção diária gravado no banco.")
    else:
        st.markdown('<div class="saas-panel">', unsafe_allow_html=True)
        c_f1, c_f2 = st.columns(2)
        with c_f1: filtro_nuc = st.selectbox("Filtrar por Núcleo/Fazenda:", ["Todos"] + sorted(list(df_p["Nome Fazenda"].unique())))
        with c_f2: filtro_lot = st.selectbox("Filtrar por Lote:", ["Todos"] + sorted(list(df_p["Nº Lote"].unique())))
        df_filtered = df_p.copy()
        if filtro_nuc != "Todos": df_filtered = df_filtered[df_filtered["Nome Fazenda"] == filtro_nuc]
        if filtro_lot != "Todos": df_filtered = df_filtered[df_filtered["Nº Lote"] == filtro_lot]
        cols_exib = [c for c in ["Data Produção", "Nome Fazenda", "Nº Lote", "Nome Produto", "Ovos", "Cód Ref", "Nº Rastreamento"] if c in df_filtered.columns]
        st.dataframe(df_filtered[cols_exib], use_container_width=True, hide_index=True)
        st.markdown('</div>', unsafe_allow_html=True)

elif tela == "📋 Validação de Reclassificação":
    st.markdown("""
    <div class="top-header-bar">
        <div class="top-header-title">
            <span>Procedimento Operacional & Validação</span>
            <h1>Validação de Reclassificação</h1>
        </div>
    </div>
    """, unsafe_allow_html=True)

    df_r_original = DatabaseManager.obter("df_reclass")
    df_p_original = DatabaseManager.obter("df_prod")
    df_rec_mtech_orig = DatabaseManager.obter("df_rec_mtech")

    if df_r_original.empty:
        show_warning("Carregue a ficha de reclassificação na aba 'Central de Carga de Arquivos' para visualizar.")
    else:
        st.markdown("""
        <div style="background-color: #F0F9FF; border: 1px solid #BAE6FD; border-left: 6px solid #0284C7; border-radius: 8px; padding: 18px 22px; margin-bottom: 24px;">
            <h4 style="margin: 0 0 10px 0; color: #0369A1; font-size: 16px; font-weight: 800;">⚙️ Painel de Filtros (Histórico do Banco de Dados)</h4>
            <p style="margin: 0; font-size: 14.5px; color: #0C4A6E;">A <b>Data de Transação (Processamento)</b> sincroniza os lotes auditados no MTech com a ficha física correspondente gravada no banco.</p>
        </div>
        """, unsafe_allow_html=True)

        col_f1, col_f2, _ = st.columns([2, 2, 3])
        
        with col_f1:
            unicos_proc = [d for d in df_r_original["Data_Processamento"].unique() if str(d).strip() != ""]
            datas_proc = ["Todos os Processamentos"] + sorted(list(unicos_proc), reverse=True)
            filtro_proc = st.selectbox("📅 Data de Transação (Processamento):", datas_proc)

        with col_f2:
            if filtro_proc != "Todos os Processamentos":
                datas_post_validas = df_r_original[df_r_original["Data_Processamento"] == filtro_proc]["Data_Postura"].dropna().unique()
            else:
                datas_post_validas = df_r_original["Data_Postura"].dropna().unique()
                
            datas_post = ["Todas as Datas"] + sorted(list(datas_post_validas), reverse=True)
            filtro_dt = st.selectbox("🥚 Data de Postura (Produção):", datas_post)

        df_r = df_r_original.copy()
        if filtro_proc != "Todos os Processamentos":
            df_r = df_r[df_r["Data_Processamento"] == filtro_proc]
        if filtro_dt != "Todas as Datas":
            df_r = df_r[df_r["Data_Postura"] == filtro_dt]

        if df_r.empty:
            show_warning("Nenhum dado encontrado para esta combinação de filtros.")
            st.stop()

        # ======================================================================
        # CONFERÊNCIA DE SALDO INICIAL: PRODUÇÃO MTECH (CANÔNICO) vs FICHA CPO
        # ======================================================================
        if not df_p_original.empty:
            st.markdown('<div class="saas-panel" style="border-left: 6px solid #10B981;">', unsafe_allow_html=True)
            st.subheader("🔎 Auditoria de Saldo Inicial vs Produção MTech (Núcleo)")
            
            linhas_conf_saldo = []
            df_nn_sub = df_r[df_r["Origem"] == "Ninho Normal"]
            
            for _, r_nn in df_nn_sub.iterrows():
                dt_p = r_nn["Data_Postura"]
                lt = str(r_nn["Lote"]).strip()
                saldo_cpo = int(r_nn["Saldo_Inicial"])
                class_hoje = int(r_nn["Total_Classificado"])
                
                prod_mtech_galpao = df_p_original[
                    (df_p_original["Data Produção"] == dt_p) & 
                    (df_p_original["Nº Lote"] == lt)
                ]["Ovos"].sum()
                
                if prod_mtech_galpao > 0:
                    dif_saldo_ini = saldo_cpo - prod_mtech_galpao
                    saldo_restante_real = max(0, prod_mtech_galpao - class_hoje)
                    status_saldo = "✅ Bate com Produção" if dif_saldo_ini == 0 else f"⚠️ Saldo Ficha difere ({dif_saldo_ini:+,} un)"
                    
                    linhas_conf_saldo.append({
                        "Data Postura": dt_p,
                        "Lote": lt,
                        "Saldo Inicial (Ficha)": saldo_cpo,
                        "Produção Núcleo (MTech)": int(prod_mtech_galpao),
                        "Classificado Hoje": class_hoje,
                        "Saldo Restante Núcleo": int(saldo_restante_real),
                        "Status Auditoria": status_saldo
                    })
            
            if linhas_conf_saldo:
                df_conf_disp = pd.DataFrame(linhas_conf_saldo)
                st.dataframe(df_conf_disp, use_container_width=True, hide_index=True)
                st.caption("ℹ️ *A coluna 'Produção Núcleo (MTech)' é a referência canónica do lote coletado no aviário. O 'Saldo Restante' indica o que ainda resta para classificar dessa postura.*")
            st.markdown('</div>', unsafe_allow_html=True)

        # ======================================================================
        # 1. LANÇAMENTO DE RECLASSIFICAÇÃO
        # ======================================================================
        st.markdown('<div class="saas-panel">', unsafe_allow_html=True)
        st.subheader("1. Lançamento de Reclassificação — Destinos Comerciais")
        
        df_comercial = df_r[df_r["Origem"].isin(["Ninho Normal", "Ninho Sujo", "Cama"])].copy()
        df_comercial["Sujo"] = df_comercial.apply(lambda row: row["Saldo_Inicial"] if row["Origem"] in ["Ninho Sujo", "Cama"] else row["Sujo"], axis=1)

        colunas_espelho = ["Data_Postura", "Lote", "Origem", "Saldo_Inicial", "Sujo", "Deformado", "Ninho_Sujo", "Industrial", "Duas_Gemas", "Ovo_Limpo", "Trincado", "Total_Destinos", "Saldo_Final"]
        df_comercial_limpo = df_comercial[colunas_espelho].copy()
        
        if not df_comercial_limpo.empty:
            df_totais_comercial = pd.DataFrame([["TOTAIS GERAIS", "-", "-", df_comercial_limpo["Saldo_Inicial"].sum(), df_comercial_limpo["Sujo"].sum(), df_comercial_limpo["Deformado"].sum(), df_comercial_limpo["Ninho_Sujo"].sum(), df_comercial_limpo["Industrial"].sum(), df_comercial_limpo["Duas_Gemas"].sum(), df_comercial_limpo["Ovo_Limpo"].sum(), df_comercial_limpo["Trincado"].sum(), df_comercial_limpo["Total_Destinos"].sum(), df_comercial_limpo["Saldo_Final"].sum()]], columns=colunas_espelho)
            df_comercial_display = pd.concat([df_comercial_limpo, df_totais_comercial], ignore_index=True)
        else:
            df_comercial_display = df_comercial_limpo
        
        tab_v1, tab_v2 = st.tabs(["🎨 Visão Formatada (Com Totais)", "📊 Tabela Interativa (Planilha Limpa)"])
        with tab_v1: st.markdown(render_reclass_table_html(df_comercial), unsafe_allow_html=True)
        with tab_v2: st.dataframe(df_comercial_display.style.apply(lambda x: ['background-color: #f1f5f9; font-weight: bold' if x.name == df_comercial_display.index[-1] else '' for i in x], axis=1) if not df_comercial_display.empty else df_comercial_display, use_container_width=True, hide_index=True)
        st.markdown('</div>', unsafe_allow_html=True)

        # ======================================================================
        # 2. RESUMOS OPERACIONAIS
        # ======================================================================
        st.markdown('<div class="saas-panel">', unsafe_allow_html=True)
        st.subheader("2. Resumo de Fechamento por Data de Postura")
        col_res1, col_res2 = st.columns(2)
        with col_res1:
            st.markdown("**🥚 Resumo de Microtrincados para Faturamento**")
            res_trinc = df_r[df_r["Origem"] != "Ovos Perdidos (Aba)"].groupby("Data_Postura").agg({"Microtrincado": "sum"}).reset_index()
            if not res_trinc.empty: res_trinc.loc["Total"] = ["TOTAIS GERAIS", res_trinc["Microtrincado"].sum()]
            st.dataframe(res_trinc.style.apply(lambda x: ['font-weight: bold; background-color: #f1f5f9' if x.name == 'Total' else '' for i in x], axis=1) if not res_trinc.empty else res_trinc, use_container_width=True, hide_index=True)
        with col_res2:
            st.markdown("**🧹 Resumo Consolidado de Ovos Sujos**")
            res_sujo = df_r[df_r["Origem"] != "Ovos Perdidos (Aba)"].groupby("Data_Postura").agg({"Sujo": "sum"}).reset_index()
            res_sujo.rename(columns={"Sujo": "Total Sujo"}, inplace=True)
            if not res_sujo.empty: res_sujo.loc["Total"] = ["TOTAIS GERAIS", res_sujo["Total Sujo"].sum()]
            st.dataframe(res_sujo.style.apply(lambda x: ['font-weight: bold; background-color: #f1f5f9' if x.name == 'Total' else '' for i in x], axis=1) if not res_sujo.empty else res_sujo, use_container_width=True, hide_index=True)
        st.markdown('</div>', unsafe_allow_html=True)

        # ======================================================================
        # 3. VALIDAÇÃO EXCLUSIVA: RECLASSIFICAÇÃO (FICHA CPO vs MTECH DÉBITO)
        # ======================================================================
        st.markdown('<div class="saas-panel" style="border-top: 4px solid #1E3A8A;">', unsafe_allow_html=True)
        st.subheader("3. Validação de Reclassificação (Destinos Comerciais)")
        
        if df_rec_mtech_orig.empty:
            show_info("Para validar a reclassificação, solte a planilha de Reclassificação do MTech na Central de Carga.")
        elif filtro_proc == "Todos os Processamentos":
            show_warning("Selecione uma <b>Data de Transação (Processamento)</b> específica no painel de filtros para ativar a validação do lote contábil exato.")
        else:
            df_rec_m_sub = df_rec_mtech_orig.copy()
            df_rec_m_sub = df_rec_m_sub[df_rec_m_sub["Data Transação"] == filtro_proc]
            
            if filtro_dt != "Todas as Datas": df_rec_m_sub = df_rec_m_sub[df_rec_m_sub["Data Produção"] == filtro_dt]

            col_nome = "Nome Produto" if "Nome Produto" in df_rec_m_sub.columns else ("Produto" if "Produto" in df_rec_m_sub.columns else "")
            if col_nome:
                df_rec_m_sub["Categoria_Normalizada"] = df_rec_m_sub[col_nome].apply(mapear_categoria_mtech)
            else:
                df_rec_m_sub["Categoria_Normalizada"] = "Outros"

            tot_saida_nn = df_r[df_r["Origem"] == "Ninho Normal"]["Total_Classificado"].sum()
            tot_saida_ns = df_r[df_r["Origem"] == "Ninho Sujo"]["Saldo_Inicial"].sum()
            tot_saida_cm = df_r[df_r["Origem"] == "Cama"]["Saldo_Inicial"].sum()

            mtech_saidas = df_rec_m_sub[df_rec_m_sub["Cód Trans Ovo"].astype(str).str.contains("Cr[é|e]dito", case=False, na=False)]
            mtech_saida_nn = mtech_saidas[mtech_saidas["Categoria_Normalizada"] == "Ninho Normal"]["Ovos"].sum()
            mtech_saida_ns = mtech_saidas[mtech_saidas["Categoria_Normalizada"] == "Ninho Sujo"]["Ovos"].sum()
            mtech_saida_cm = mtech_saidas[mtech_saidas["Categoria_Normalizada"] == "Cama"]["Ovos"].sum()

            erros_origem = []
            if tot_saida_nn != mtech_saida_nn: erros_origem.append(f"• Ninho Normal (Ficha CPO: {int(tot_saida_nn):,} | MTech Crédito: {int(mtech_saida_nn):,})".replace(",", "."))
            if tot_saida_ns != mtech_saida_ns: erros_origem.append(f"• Ninho Sujo Velho (Ficha CPO: {int(tot_saida_ns):,} | MTech Crédito: {int(mtech_saida_ns):,})".replace(",", "."))
            if tot_saida_cm != mtech_saida_cm: erros_origem.append(f"• Cama Velha (Ficha CPO: {int(tot_saida_cm):,} | MTech Crédito: {int(mtech_saida_cm):,})".replace(",", "."))
                
            if erros_origem: show_error("Divergência detectada nas Origens (Saídas/Créditos) do MTech:<br><br>" + "<br>".join(erros_origem))

            categorias_alvo = ["Sujo", "Deformado", "Ninho Sujo", "Industrial", "Duas Gemas", "Ovo Limpo", "Trincado"]
            tot_dest_ficha = df_comercial[["Sujo", "Deformado", "Ninho_Sujo", "Industrial", "Duas_Gemas", "Ovo_Limpo", "Trincado"]].sum().sum()
            
            tot_cred_m = df_rec_m_sub[
                (df_rec_m_sub["Cód Trans Ovo"].astype(str).str.contains("D[é|e]bito", case=False, na=False)) &
                (df_rec_m_sub["Categoria_Normalizada"].isin(categorias_alvo))
            ]["Ovos"].sum()
            
            cats_comerciais = [
                ("Sujo", "Sujo", df_comercial["Sujo"].sum()),
                ("Deformado", "Deformado", df_comercial["Deformado"].sum()),
                ("Ninho Sujo", "Ninho Sujo", df_comercial["Ninho_Sujo"].sum()),
                ("Industrial", "Industrial", df_comercial["Industrial"].sum()),
                ("Duas Gemas", "Duas Gemas", df_comercial["Duas_Gemas"].sum()),
                ("Ovo Limpo", "Ovo Limpo", df_comercial["Ovo_Limpo"].sum()),
                ("Trincado", "Trincado", df_comercial["Trincado"].sum()),
            ]

            linhas_validacao_reclass = []
            for nome_exib, cat_chave, qtd_f in cats_comerciais:
                qtd_m = df_rec_m_sub[
                    (df_rec_m_sub["Cód Trans Ovo"].astype(str).str.contains("D[é|e]bito", case=False, na=False)) & 
                    (df_rec_m_sub["Categoria_Normalizada"] == cat_chave)
                ]["Ovos"].sum()
                
                diferenca = qtd_f - qtd_m
                linhas_validacao_reclass.append({
                    "Categoria do Ovo": nome_exib,
                    "Ficha CPO (Físico)": int(qtd_f),
                    "MTech (Sistema)": int(qtd_m),
                    "Divergência": int(diferenca),
                    "Status": "✅ OK" if diferenca == 0 else f"❌ Erro ({diferenca:+,})"
                })

            df_val_reclass = pd.DataFrame(linhas_validacao_reclass)
            tot_div_abs_reclass = df_val_reclass["Divergência"].abs().sum()

            c_val1, c_val2 = st.columns(2)
            with c_val1:
                st.markdown(f"""
                <div class="kpi-card" style="border-left: 5px solid {'#16A34A' if tot_div_abs_reclass == 0 else '#DC2626'}; min-height: 105px;">
                    <div class="kpi-icon-bubble">
                        {'✓' if tot_div_abs_reclass == 0 else '✕'}
                    </div>
                    <div>
                        <p class="kpi-title">Balanço Total de Reclassificação</p>
                        <p class="kpi-value" style="color: {'#16A34A' if tot_div_abs_reclass == 0 else '#DC2626'}; font-size: 1.4rem;">
                            {'✅ 100% Validado' if tot_div_abs_reclass == 0 else f'❌ Divergência Absoluta: {tot_div_abs_reclass:,} ovos'}
                        </p>
                        <p class="kpi-sub" style="color: var(--text-color); opacity: 0.8;">
                            Ficha CPO: {tot_dest_ficha:,.0f} | MTech (Entrada): {tot_cred_m:,.0f}
                        </p>
                    </div>
                </div>
                """.replace(",", "."), unsafe_allow_html=True)
            with c_val2:
                st.markdown(f"""
                <div class="kpi-card" style="border-left: 5px solid #0284C7; min-height: 105px;">
                    <div class="kpi-icon-bubble">📋</div>
                    <div>
                        <p class="kpi-title">Transação Auditada</p>
                        <p class="kpi-value" style="font-size: 1.3rem;">{filtro_proc}</p>
                        <p class="kpi-sub" style="color: var(--text-color); opacity: 0.8;">
                            {'Todas as posturas consolidadas' if filtro_dt == 'Todas as Datas' else f'Postura: {filtro_dt}'}
                        </p>
                    </div>
                </div>
                """, unsafe_allow_html=True)

            df_val_reclass.loc[len(df_val_reclass)] = {
                "Categoria do Ovo": "TOTAIS GERAIS",
                "Ficha CPO (Físico)": df_val_reclass["Ficha CPO (Físico)"].sum(),
                "MTech (Sistema)": df_val_reclass["MTech (Sistema)"].sum(),
                "Divergência": tot_div_abs_reclass,
                "Status": ""
            }

            st.dataframe(df_val_reclass.style.apply(lambda x: ['font-weight: bold; background-color: #f1f5f9' if x.name == df_val_reclass.index[-1] else '' for i in x], axis=1), use_container_width=True, hide_index=True)

        st.markdown('</div>', unsafe_allow_html=True)

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df_comercial_limpo.to_excel(writer, sheet_name="Reclassificacao", index=False)
            res_trinc.to_excel(writer, sheet_name="Resumo_Trincados", index=False)
            res_sujo.to_excel(writer, sheet_name="Resumo_Sujos", index=False)

        st.download_button("📥 Baixar Ficha de Reclassificação (.xlsx)", data=buf.getvalue(), file_name="Ficha_Reclassificacao_Validada.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif tela == "🗑️ Validação de Descartes":
    st.markdown("""
    <div class="top-header-bar">
        <div class="top-header-title">
            <span style="color: #DC2626 !important;">Auditoria de Baixas & Perdas</span>
            <h1>Validação de Descartes</h1>
        </div>
    </div>
    """, unsafe_allow_html=True)

    df_r_original = DatabaseManager.obter("df_reclass")
    df_p_original = DatabaseManager.obter("df_prod")
    df_desc_mtech_orig = DatabaseManager.obter("df_desc_mtech")

    if df_r_original.empty:
        show_warning("Carregue a ficha de reclassificação na aba 'Central de Carga de Arquivos' para auditar os descartes.")
    else:
        col_f1, col_f2, _ = st.columns([2, 2, 3])
        with col_f1:
            unicos_proc = [d for d in df_r_original["Data_Processamento"].unique() if str(d).strip() != ""]
            datas_proc = ["Todos os Processamentos"] + sorted(list(unicos_proc), reverse=True)
            filtro_proc = st.selectbox("📅 Data de Transação (Processamento):", datas_proc, key="desc_filtro_proc")

        with col_f2:
            if filtro_proc != "Todos os Processamentos":
                datas_post_validas = df_r_original[df_r_original["Data_Processamento"] == filtro_proc]["Data_Postura"].dropna().unique()
            else:
                datas_post_validas = df_r_original["Data_Postura"].dropna().unique()
                
            datas_post = ["Todas as Datas"] + sorted(list(datas_post_validas), reverse=True)
            filtro_dt = st.selectbox("🥚 Data de Postura (Produção):", datas_post, key="desc_filtro_dt")

        df_r = df_r_original.copy()
        if filtro_proc != "Todos os Processamentos":
            df_r = df_r[df_r["Data_Processamento"] == filtro_proc]
        if filtro_dt != "Todas as Datas":
            df_r = df_r[df_r["Data_Postura"] == filtro_dt]

        # 1. Montagem da Baixa FÍSICA (Aba principal envia NN e Trincado)
        cols_base_desc = ["Data_Postura", "Lote", "P_NN", "Trincado_Baixa"]
        df_baixa_class = df_r[df_r["Origem"].isin(["Ninho Normal", "Ninho Sujo", "Cama"])].copy()
        for col_chk in cols_base_desc:
            if col_chk not in df_baixa_class.columns:
                df_baixa_class[col_chk] = 0
        df_baixa_class = df_baixa_class[cols_base_desc].copy()
        df_baixa_class = df_baixa_class.rename(columns={"P_NN": "Ninho Normal", "Trincado_Baixa": "Trincado"})
        
        # Ovos da aba secundária são consolidados ignorando o lote (Lote = "-")
        cols_perd = ["P_NN", "P_Cama", "P_NS", "P_Def", "P_Sujo", "P_Ind", "P_DG", "P_Trinc"]
        df_baixa_perd = df_r[df_r["Origem"] == "Ovos Perdidos (Aba)"].copy()
        for col_chk in cols_perd:
            if col_chk not in df_baixa_perd.columns:
                df_baixa_perd[col_chk] = 0
        df_baixa_perd = df_baixa_perd[["Data_Postura", "Lote"] + cols_perd].copy()
        df_baixa_perd["Lote"] = "-"
        df_baixa_perd = df_baixa_perd.rename(columns={"P_NN": "Ninho Normal", "P_Cama": "Cama", "P_NS": "Ninho Sujo", "P_Def": "Deformado", "P_Sujo": "Sujo Granja", "P_Ind": "Industrial", "P_DG": "Duas Gemas", "P_Trinc": "Trincado"})
        
        df_desc_final = pd.concat([df_baixa_class, df_baixa_perd], ignore_index=True).fillna(0)
        
        # Consolida descartes
        all_cols = ["Ninho Normal", "Cama", "Ninho Sujo", "Deformado", "Sujo Granja", "Industrial", "Duas Gemas", "Trincado"]
        for c in all_cols:
            if c not in df_desc_final.columns: 
                df_desc_final[c] = 0
            df_desc_final[c] = pd.to_numeric(df_desc_final[c], errors="coerce").fillna(0).astype(int)

        df_desc_final = df_desc_final.groupby(["Data_Postura", "Lote"])[all_cols].sum().reset_index()
        df_desc_final["TOTAL BAIXA"] = df_desc_final[all_cols].sum(axis=1).fillna(0).astype(int)
        df_desc_final = df_desc_final[df_desc_final["TOTAL BAIXA"] > 0].sort_values(by=["Data_Postura", "Lote"], ascending=[False, True])

        # ======================================================================
        # 1. LANÇAMENTO FÍSICO DE BAIXA & DESCARTE (CPO)
        # ======================================================================
        st.markdown('<div class="saas-panel">', unsafe_allow_html=True)
        st.subheader("1. Lançamento Físico de Baixa & Descarte (CPO)")
        
        if not df_desc_final.empty:
            df_totais_desc = pd.DataFrame([["TOTAIS GERAIS", "-"] + [df_desc_final[c].sum() for c in all_cols] + [df_desc_final["TOTAL BAIXA"].sum()]], columns=["Data_Postura", "Lote"] + all_cols + ["TOTAL BAIXA"])
            df_desc_display = pd.concat([df_desc_final, df_totais_desc], ignore_index=True)
            
            tab_d1, tab_d2 = st.tabs(["🔴 Visão Formatada (Detalhamento Completo)", "📊 Tabela Interativa"])
            with tab_d1: st.markdown(render_descarte_table_html(df_desc_final), unsafe_allow_html=True)
            with tab_d2: st.dataframe(df_desc_display.style.apply(lambda x: ['background-color: #fca5a5; font-weight: bold; color: #991B1B' if x.name == df_desc_display.index[-1] else '' for i in x], axis=1), use_container_width=True, hide_index=True)
        else:
            show_info("Nenhuma baixa/descarte para registrar nas abas físicas CPO.")
        st.markdown('</div>', unsafe_allow_html=True)

        # 2. Filtragem dos Descartes do MTech para Comparativo
        df_desc_m_sub = df_desc_mtech_orig.copy()
        if not df_desc_m_sub.empty:
            if filtro_proc != "Todos os Processamentos":
                df_desc_m_sub = df_desc_m_sub[df_desc_m_sub["Data Transação"] == filtro_proc]
            if filtro_dt != "Todas as Datas":
                df_desc_m_sub = df_desc_m_sub[df_desc_m_sub["Data Produção"] == filtro_dt]

            col_nome_desc = "Produto" if "Produto" in df_desc_m_sub.columns else ("Nome Produto" if "Nome Produto" in df_desc_m_sub.columns else "")
            if col_nome_desc:
                df_desc_m_sub["Categoria_Normalizada"] = df_desc_m_sub[col_nome_desc].apply(mapear_categoria_mtech)
            else:
                df_desc_m_sub["Categoria_Normalizada"] = "Outros"

        tot_desc_esperado_cpo = df_desc_final["TOTAL BAIXA"].sum() if not df_desc_final.empty else 0
        
        nucl_desc = {"Trincado": 0, "Perda Galpão": 0}
        if not df_p_original.empty:
            df_p_sub = df_p_original.copy()
            if filtro_dt != "Todas as Datas": df_p_sub = df_p_sub[df_p_sub["Data Produção"] == filtro_dt]
            col_nome_prod = "Nome Produto" if "Nome Produto" in df_p_sub.columns else ("Produto" if "Produto" in df_p_sub.columns else "")
            if col_nome_prod:
                df_p_sub["Cat_Prod"] = df_p_sub[col_nome_prod].apply(mapear_categoria_mtech)
                df_nucl_desc = df_p_sub[df_p_sub["Cat_Prod"].isin(["Trincado", "Perda Galpão"])]
                nucl_dict = df_nucl_desc.groupby("Cat_Prod")["Ovos"].sum().to_dict()
                nucl_desc["Trincado"] = nucl_dict.get("Trincado", 0)
                nucl_desc["Perda Galpão"] = nucl_dict.get("Perda Galpão", 0)

        tot_desc_galpao = sum(nucl_desc.values())
        tot_esperado_geral = tot_desc_esperado_cpo + tot_desc_galpao
        
        categorias_descarte_ativas = ["Ninho Normal", "Cama", "Ninho Sujo", "Deformado", "Sujo", "Industrial", "Duas Gemas", "Trincado", "Perda Galpão"]
        tot_desc_m = df_desc_m_sub[df_desc_m_sub["Categoria_Normalizada"].isin(categorias_descarte_ativas)]["Ovos"].sum() if not df_desc_m_sub.empty else 0

        cats_descartes = [
            ("Ninho Normal", "Ninho Normal", "Ninho Normal"),
            ("Cama", "Cama", "Cama"),
            ("Ninho Sujo", "Ninho Sujo", "Ninho Sujo"),
            ("Deformado", "Deformado", "Deformado"),
            ("Sujo Granja", "Sujo", "Sujo Granja"),
            ("Industrial", "Industrial", "Industrial"),
            ("Duas Gemas", "Duas Gemas", "Duas Gemas"),
            ("Trincado", "Trincado", "Trincado"),
            ("Ovo Perdido", "Perda Galpão", None)
        ]

        linhas_validacao_desc = []
        for nome_exib, cat_mtech, col_cpo in cats_descartes:
            qtd_cpo = df_desc_final[col_cpo].sum() if col_cpo and col_cpo in df_desc_final.columns else 0
            qtd_nucl = nucl_desc.get(cat_mtech, 0)
            tot_esperado = qtd_cpo + qtd_nucl
            
            qtd_mtech = df_desc_m_sub[df_desc_m_sub["Categoria_Normalizada"] == cat_mtech]["Ovos"].sum() if not df_desc_m_sub.empty else 0
            
            diferenca = tot_esperado - qtd_mtech
            
            linhas_validacao_desc.append({
                "Tipo de Ovo / Categoria": nome_exib,
                "Ficha CPO (Físico)": int(qtd_cpo),
                "Núcleo (Produção)": int(qtd_nucl),
                "Total Esperado": int(tot_esperado),
                "MTech (Sistema)": int(qtd_mtech),
                "Divergência": int(diferenca),
                "Status": "✅ OK" if diferenca == 0 else f"❌ Erro ({diferenca:+,})"
            })

        df_val_desc = pd.DataFrame(linhas_validacao_desc)
        tot_div_abs_desc = df_val_desc["Divergência"].abs().sum()

        st.markdown('<div class="saas-panel" style="border-top: 4px solid #DC2626;">', unsafe_allow_html=True)
        st.subheader("2. Comparativo Detalhado por Tipo de Ovo (Descartes & Perdas)")

        st.markdown(f"""
        <div class="kpi-container">
            <div class="kpi-card" style="border-left: 5px solid {'#16A34A' if tot_div_abs_desc == 0 else '#DC2626'};">
                <div class="kpi-icon-bubble">
                    {'✓' if tot_div_abs_desc == 0 else '✕'}
                </div>
                <div>
                    <p class="kpi-title">Balanço Total de Descartes</p>
                    <p class="kpi-value" style="color: {'#16A34A' if tot_div_abs_desc == 0 else '#DC2626'}; font-size: 1.35rem;">
                        {'✅ 100% Validado' if tot_div_abs_desc == 0 else f'❌ Divergência Real: {tot_div_abs_desc:,} ovos'}
                    </p>
                    <p class="kpi-sub" style="color: var(--text-color); opacity: 0.8;">
                        Total Esperado: {tot_esperado_geral:,.0f} | MTech: {tot_desc_m:,.0f}
                    </p>
                </div>
            </div>
            <div class="kpi-card">
                <div class="kpi-icon-bubble">🥚</div>
                <div>
                    <p class="kpi-title">Descartes CPO (Fichas)</p>
                    <p class="kpi-value" style="font-size: 1.25rem;">{tot_desc_esperado_cpo:,.0f}</p>
                    <p class="kpi-sub" style="color: var(--text-color); opacity: 0.8;">NN Perdidos + Trincados + Outros da Aba Secundária</p>
                </div>
            </div>
            <div class="kpi-card">
                <div class="kpi-icon-bubble">🚜</div>
                <div>
                    <p class="kpi-title">Descartes Núcleo (Produção)</p>
                    <p class="kpi-value" style="font-size: 1.25rem;">{tot_desc_galpao:,.0f}</p>
                    <p class="kpi-sub" style="color: var(--text-color); opacity: 0.8;">Trincados + Perdidos na Coleta</p>
                </div>
            </div>
        </div>
        """.replace(",", "."), unsafe_allow_html=True)

        df_val_desc.loc[len(df_val_desc)] = {
            "Tipo de Ovo / Categoria": "TOTAIS GERAIS",
            "Ficha CPO (Físico)": df_val_desc["Ficha CPO (Físico)"].sum(),
            "Núcleo (Produção)": df_val_desc["Núcleo (Produção)"].sum(),
            "Total Esperado": df_val_desc["Total Esperado"].sum(),
            "MTech (Sistema)": df_val_desc["MTech (Sistema)"].sum(),
            "Divergência": tot_div_abs_desc,
            "Status": ""
        }

        st.dataframe(df_val_desc.style.apply(lambda x: ['font-weight: bold; background-color: #E2E8F0; color: #1E293B' if x.name == df_val_desc.index[-1] else '' for i in x], axis=1), use_container_width=True, hide_index=True)
        st.markdown('</div>', unsafe_allow_html=True)

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            if not df_desc_final.empty:
                df_desc_final.to_excel(writer, sheet_name="Descartes", index=False)
            df_val_desc.to_excel(writer, sheet_name="Auditoria_MTech", index=False)

        st.download_button("📥 Baixar Ficha de Descartes Validada (.xlsx)", data=buf.getvalue(), file_name="Ficha_Descartes_Validada.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif tela == "🚚 Validação Estoque (MTech x Sankhya x CPO)":
    st.markdown("""
    <div class="top-header-bar">
        <div class="top-header-title">
            <span>Visão Consolidada de Inventário Fotográfico</span>
            <h1>Validação Estoque (MTech x Sankhya x CPO)</h1>
        </div>
    </div>
    """, unsafe_allow_html=True)

    df_m = DatabaseManager.obter("df_est_mtech")
    df_s = DatabaseManager.obter("df_est_sankhya")
    df_c = DatabaseManager.obter("df_est_cpo")

    if df_s.empty or df_m.empty:
        show_warning("Para o comparativo funcionar, envie os relatórios de **Estoque do MTech** e **Estoque do Sankhya** na Central de Carga.")
    else:
        df_m["Nome Padrao"] = df_m["Produto_Original"].map(MAP_MTECH).fillna("Outros")
        mtech_agg = df_m.groupby("Nome Padrao")["Quantidade"].sum().reset_index()
        mtech_agg.rename(columns={"Quantidade": "Qtd MTech"}, inplace=True)

        df_s["Nome Padrao"] = df_s["Produto_Original"].str.strip().map(MAP_SANKHYA).fillna("Outros")
        sank_agg = df_s.groupby("Nome Padrao")["Quantidade"].sum().reset_index()
        sank_agg.rename(columns={"Quantidade": "Qtd Sankhya"}, inplace=True)

        if not df_c.empty:
            df_c["Nome Padrao"] = df_c["Categoria_Original"].str.strip().map(MAP_CPO).fillna("Outros")
            cpo_agg = df_c.groupby("Nome Padrao")["Quantidade"].sum().reset_index()
            cpo_agg.rename(columns={"Quantidade": "Qtd CPO"}, inplace=True)
        else:
            cpo_agg = pd.DataFrame(columns=["Nome Padrao", "Qtd CPO"])

        df_est = pd.DataFrame({"Nome Padrao": PADRAO_ORDEM})
        df_est = df_est.merge(mtech_agg, on="Nome Padrao", how="left")
        df_est = df_est.merge(sank_agg, on="Nome Padrao", how="left")
        df_est = df_est.merge(cpo_agg, on="Nome Padrao", how="left")

        df_est.fillna(0, inplace=True)
        
        df_est["Qtd MTech"] = pd.to_numeric(df_est["Qtd MTech"], errors="coerce").fillna(0).astype(int)
        df_est["Qtd Sankhya"] = pd.to_numeric(df_est["Qtd Sankhya"], errors="coerce").fillna(0).astype(int)
        df_est["Qtd CPO"] = pd.to_numeric(df_est["Qtd CPO"], errors="coerce").fillna(0).astype(int)

        df_est["Divergencia Qtd"] = df_est["Qtd Sankhya"] - df_est["Qtd MTech"]
        df_est["Divergencia CPO"] = df_est["Qtd CPO"] - df_est["Qtd Sankhya"]

        st.markdown('<div class="saas-panel">', unsafe_allow_html=True)
        st.subheader("1. Comparativo Fiscal vs Origem (SANKHYA - MTECH)")
        st.markdown(render_estoque_sankhya_mtech_html(df_est), unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="saas-panel" style="border-top: 4px solid #1E3A8A;">', unsafe_allow_html=True)
        st.subheader("2. Comparativo Físico vs Fiscal (CPO - SANKHYA)")
        if df_c.empty:
            show_info("Suba a planilha de Estoque do CPO (Aba RESUMO) para liberar este comparativo.")
        else:
            st.markdown(render_estoque_cpo_sankhya_html(df_est), unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)