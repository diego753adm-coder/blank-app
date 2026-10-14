import io
import os
import re
import sqlite3
from datetime import datetime
import openpyxl
import pandas as pd
import streamlit as st

# ==============================================================================
# 1. CAMADA DE BANCO DE DADOS LOCAL (SQLITE - PRONTO PARA MIGRAÇÃO FUTURA)
# ==============================================================================
DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "historico_ovos_117.db")

class DatabaseManager:
    CHAVES_UNICIDADE = {
        "df_reclass": ["Data_Processamento", "Data_Postura", "Lote", "Origem"],
        "df_prod": ["Data Produção", "Nome Fazenda", "Nº Lote", "Nome Produto"],
        "df_vendas": ["Data_Expedicao", "Cliente", "Destino", "Data_Postura", "Lote", "Tipo_Ovo"],
        "df_transf": ["Data_Expedicao", "Aba_Destino", "Palete", "Lote", "Data_Postura", "Categoria"],
        "df_rec_mtech": ["Data Transação", "Data Produção", "Nº Lote", "Nome Produto", "Cód Trans Ovo"],
        "df_desc_mtech": ["Data Transação", "Data Produção", "Nº Lote", "Produto"]
    }

    @staticmethod
    def _get_conn():
        return sqlite3.connect(DB_PATH, check_same_thread=False)

    @staticmethod
    def inicializar():
        tabelas = ["df_prod", "df_reclass", "df_vendas", "df_transf", "df_rec_mtech", "df_desc_mtech"]
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
            return
        
        df_atual = DatabaseManager.obter(nome_tabela)
        if not df_atual.empty:
            df_combinado = pd.concat([df_atual, df_novo], ignore_index=True)
        else:
            df_combinado = df_novo.copy()

        chaves = DatabaseManager.CHAVES_UNICIDADE.get(nome_tabela, [])
        chaves_validas = [c for c in chaves if c in df_combinado.columns]
        if chaves_validas:
            df_combinado = df_combinado.drop_duplicates(subset=chaves_validas, keep="last").reset_index(drop=True)
        else:
            df_combinado = df_combinado.drop_duplicates(keep="last").reset_index(drop=True)

        with DatabaseManager._get_conn() as conn:
            df_combinado.to_sql(nome_tabela, conn, if_exists="replace", index=False)
        
        st.session_state[nome_tabela] = df_combinado

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
# 2. CONFIGURAÇÃO GERAL E ESTILO VISUAL SAAS (NÚMEROS MAIORES & RESPONSIVO)
# ==============================================================================
st.set_page_config(
    page_title="Hub Operacional de Ovos • Balsas 117",
    page_icon="🥚",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"], .stApp {
        font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif !important;
        font-size: 14px !important;
        background-color: #EDF5EE !important;
        color: #1F2937 !important;
    }

    header[data-testid="stHeader"] {
        background-color: transparent !important;
        box-shadow: none !important;
    }

    .block-container {
        padding-top: 1rem !important; 
        padding-bottom: 3rem !important;
        padding-left: 2rem !important;
        padding-right: 2rem !important;
        max-width: 100% !important;
        width: 100% !important;
    }

    section[data-testid="stSidebar"] {
        background-color: #132517 !important;
        border-right: 1px solid #0E1B11 !important;
    }
    section[data-testid="stSidebar"] * {
        color: #E2E8F0 !important;
    }
    
    [data-testid="stSidebarCollapseButton"], [data-testid="collapsedControl"] {
        opacity: 1 !important;
        visibility: visible !important;
        color: #84A98C !important;
        background-color: transparent !important;
    }

    .brand-box {
        background: rgba(255, 255, 255, 0.05);
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 10px;
        padding: 12px 14px;
        margin-top: -25px !important;
        margin-bottom: 20px;
        display: flex;
        align-items: center;
        gap: 12px;
    }
    .brand-icon {
        font-size: 1.4rem;
        background: #25442A;
        width: 38px;
        height: 38px;
        display: flex;
        align-items: center;
        justify-content: center;
        border-radius: 8px;
        border: 1px solid #37643E;
    }
    .brand-title {
        font-size: 0.95rem;
        font-weight: 800;
        color: #FFFFFF !important;
        margin: 0;
        letter-spacing: -0.01em;
    }
    .brand-sub {
        font-size: 0.72rem;
        color: #94A3B8 !important;
        font-weight: 500;
        text-transform: uppercase;
        letter-spacing: 0.04em;
    }

    .nav-label {
        font-size: 0.75rem;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        font-weight: 700;
        color: #84A98C !important;
        margin: 16px 4px 10px 4px;
    }

    section[data-testid="stSidebar"] div[data-testid="stRadio"] {
        width: 100% !important;
    }
    section[data-testid="stSidebar"] div[data-testid="stRadio"] div[role="radiogroup"] {
        display: flex !important;
        flex-direction: column !important;
        align-items: stretch !important;
        width: 100% !important;
        gap: 10px !important;
    }
    section[data-testid="stSidebar"] div[data-testid="stRadio"] div[role="radiogroup"] label div:first-child { 
        display: none !important; 
    }
    section[data-testid="stSidebar"] div[data-testid="stRadio"] div[role="radiogroup"] label {
        background-color: #1A3620 !important;
        border: 1px solid #1A3620 !important;
        border-radius: 8px !important;
        min-height: 48px !important;
        height: 48px !important;
        padding: 0 16px !important;
        margin: 0 !important;
        width: 100% !important;
        min-width: 100% !important;
        box-sizing: border-box !important;
        display: flex !important;
        align-items: center !important;
        cursor: pointer !important;
        transition: all 0.2s ease-in-out !important;
    }
    section[data-testid="stSidebar"] div[data-testid="stRadio"] div[role="radiogroup"] label:hover {
        background-color: #25442A !important;
        border: 1px solid #4ADE80 !important;
    }
    section[data-testid="stSidebar"] div[data-testid="stRadio"] div[role="radiogroup"] label:has(input:checked) {
        background-color: #204428 !important;
        border-left: 5px solid #4ADE80 !important;
    }
    section[data-testid="stSidebar"] div[data-testid="stRadio"] div[role="radiogroup"] label p {
        color: #E2E8F0 !important;
        font-size: 0.92rem !important;
        font-weight: 600 !important;
        margin: 0 !important;
        width: 100% !important;
    }

    section[data-testid="stSidebar"] input {
        background-color: #1B3621 !important;
        color: #FFFFFF !important;
        border: 1px solid #2B5433 !important;
        border-radius: 6px !important;
        padding: 10px !important;
        width: 100% !important;
        box-sizing: border-box !important;
    }

    .cotacao-sidebar-card {
        background: rgba(245, 158, 11, 0.08);
        border: 1px solid rgba(245, 158, 11, 0.35);
        border-radius: 10px;
        padding: 14px 16px;
        margin-top: 18px;
    }
    .cotacao-sidebar-tag {
        font-size: 0.70rem;
        font-weight: 700;
        text-transform: uppercase;
        color: #FCD34D !important;
        letter-spacing: 0.04em;
    }
    .cotacao-sidebar-val {
        font-size: 1.35rem;
        font-weight: 800;
        color: #FBBF24 !important;
        margin-top: 4px;
    }

    .saas-panel {
        background-color: #FFFFFF !important;
        border: 1px solid #D6E4D7 !important;
        border-radius: 12px !important;
        padding: 22px 24px !important;
        box-shadow: 0 2px 4px rgba(19, 37, 23, 0.03) !important;
        margin-bottom: 22px !important;
    }

    div[data-testid="stDataFrame"] div {
        font-size: 14.5px !important;
    }

    .kpi-container {
        display: grid;
        grid-template-columns: repeat(4, 1fr);
        gap: 16px;
        margin-bottom: 22px;
    }
    .kpi-card {
        background: #FFFFFF;
        border: 1px solid #D6E4D7;
        border-radius: 12px;
        padding: 16px 18px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.02);
        display: flex;
        align-items: center;
        gap: 16px;
    }
    .kpi-icon-bubble {
        width: 48px;
        height: 48px;
        border-radius: 10px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 1.4rem;
        flex-shrink: 0;
    }
    .kpi-title {
        font-size: 0.74rem;
        font-weight: 600;
        text-transform: uppercase;
        color: #64748B;
        letter-spacing: 0.03em;
        margin: 0;
    }
    .kpi-value {
        font-size: 1.45rem;
        font-weight: 800;
        color: #0F172A;
        letter-spacing: -0.02em;
        margin: 2px 0 0 0;
    }
    .kpi-sub {
        font-size: 0.76rem;
        font-weight: 600;
        margin: 2px 0 0 0;
    }

    .bubble-green { background: #E8F5E9; border: 1px solid #C8E6C9; color: #2E7D32; }
    .bubble-teal { background: #E0F2F1; border: 1px solid #B2DFDB; color: #00796B; }
    .bubble-yellow { background: #FEF9C3; border: 1px solid #FEF08A; color: #B45309; }
    .bubble-red { background: #FEE2E2; border: 1px solid #FECACA; color: #DC2626; }

    .destaque-verde {
        background-color: #F0FDF4 !important;
        border-left: 4px solid #16A34A !important;
        border-radius: 8px !important;
        padding: 10px 14px !important;
        color: #14532D !important;
        font-size: 0.86rem !important;
        margin-bottom: 14px !important;
    }

    .top-header-bar {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 18px;
        padding-bottom: 12px;
        border-bottom: 1px solid #D6E4D7;
    }
    .top-header-title h1 {
        font-size: 1.45rem !important;
        font-weight: 800 !important;
        color: #132517 !important;
        margin: 0 !important;
        letter-spacing: -0.02em !important;
    }
    .top-header-title span {
        font-size: 0.74rem !important;
        font-weight: 600 !important;
        color: #527958 !important;
        text-transform: uppercase !important;
        letter-spacing: 0.05em !important;
    }
    .top-status-badge {
        background: #FFFFFF;
        border: 1px solid #CBD5E1;
        border-radius: 20px;
        padding: 5px 12px;
        font-size: 0.78rem;
        font-weight: 600;
        color: #15803D;
        box-shadow: 0 1px 2px rgba(0,0,0,0.03);
    }
    </style>
""", unsafe_allow_html=True)

# ==============================================================================
# 3. MOTOR DE LEITURA DE CONTEÚDO INTERNO & PADRONIZAÇÃO
# ==============================================================================
def normalizar_data_str(val, ano_padrao="2026"):
    if val is None or pd.isna(val):
        return ""
    if isinstance(val, (pd.Timestamp, datetime)):
        return val.strftime("%d/%m/%Y")
    
    s = str(val).strip().lower()
    m_full = re.match(r'^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})$', s)
    if m_full:
        d, m, y = m_full.groups()
        if len(y) == 2: y = "20" + y
        return f"{int(d):02d}/{int(m):02d}/{y}"
        
    meses = {
        "jan": "01", "fev": "02", "mar": "03", "abr": "04", "mai": "05", "jun": "06",
        "jul": "07", "ago": "08", "set": "09", "out": "10", "nov": "11", "dez": "12"
    }
    for mes_txt, mes_num in meses.items():
        if mes_txt in s:
            dia_match = re.search(r'\b(\d{1,2})\b', s)
            if dia_match:
                dia = int(dia_match.group(1))
                return f"{dia:02d}/{mes_num}/{ano_padrao}"
    try:
        dt = pd.to_datetime(s, dayfirst=True)
        return dt.strftime("%d/%m/%Y")
    except Exception:
        return str(val).strip()

def mapear_categoria_mtech(texto):
    if not texto or pd.isna(texto):
        return "Outros"
    
    t = str(texto).upper().strip()
    numeros = re.findall(r'\b0?([24568]|12|13)\b', t)
    
    if "DEFORMADO" in t or "DF" in t.split() or "12" in numeros:
        return "Deformado"
    if any(x in t for x in ["DUAS GEMAS", "DUAS_GEMAS", "2 GEMAS", "GEMA"]) or "DG" in t.split() or "04" in numeros or "4" in numeros:
        return "Duas Gemas"
    if any(x in t for x in ["TRINCADO", "MICRO", "TRINCA"]) or "TR" in t.split() or "05" in numeros or "5" in numeros:
        return "Trincado"
    if "INDUSTRIAL" in t or "IND" in t.split() or "ID" in t.split() or "08" in numeros or "8" in numeros:
        return "Industrial"
    if "LIMPO" in t or "LP" in t.split() or "13" in numeros:
        return "Ovo Limpo"
    if "NINHO SUJO" in t or "N.SUJO" in t or "NS" in t.split():
        return "Ninho Sujo"
    if "NINHO NORMAL" in t or "N.NORMAL" in t or "NN" in t.split():
        return "Ninho Normal"
    if "CAMA" in t or "CM" in t.split():
        return "Cama"
    if "SUJO" in t or "SJ" in t.split() or "06" in numeros or "6" in numeros or "02" in numeros or "2" in numeros:
        return "Sujo"
        
    return "Outros"

def safe_int(val):
    if val is None or val == "": return 0
    try: return int(float(val))
    except (ValueError, TypeError): return 0

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
    return df

def parse_reclassificacao_cpo(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws_name = next((s for s in wb.sheetnames if "classifica" in s.lower()), wb.sheetnames[0])
    ws = wb[ws_name]
    
    data_proc_val = ws["C2"].value
    data_proc = normalizar_data_str(data_proc_val)
    resp = ws["G2"].value or ws["H2"].value or "Não informado"

    linhas = []
    secao_atual = "NN"

    for r in range(4, ws.max_row + 1):
        col_a = ws.cell(r, 1).value
        col_b = ws.cell(r, 2).value

        if col_a and "LOTES VELHOS" in str(col_a).upper():
            secao_atual = "LOTES_VELHOS"
            continue

        if str(col_a).strip().upper() in ["DATA POSTURA", "DATA PRODUÇÃO", "TOTAL =", "TOTAL"] or str(col_b).strip().upper() in ["LOTE", "LOTE:"]:
            continue

        if col_a is not None and col_b is not None:
            try:
                dt_post = normalizar_data_str(col_a)
                lote_num = int(float(col_b))
            except Exception:
                continue

            base_dict = {
                "Data_Processamento": data_proc,
                "Data_Postura": dt_post, "Lote": str(lote_num), "Responsavel": str(resp),
                "P_NN": 0, "P_Cama": 0, "P_NS": 0, "P_Def": 0, "P_Sujo": 0, "P_Ind": 0, "P_DG": 0, "P_Trinc": 0
            }

            if secao_atual == "NN":
                saldo_ini = safe_int(ws.cell(r, 3).value)
                if saldo_ini > 0:
                    df_val = safe_int(ws.cell(r, 4).value)
                    ns_val = safe_int(ws.cell(r, 5).value)
                    id_val = safe_int(ws.cell(r, 6).value)
                    dg_val = safe_int(ws.cell(r, 7).value)
                    sj_val = safe_int(ws.cell(r, 8).value)
                    lp_val = safe_int(ws.cell(r, 9).value)
                    micro = safe_int(ws.cell(r, 10).value)
                    trinc = safe_int(ws.cell(r, 11).value)
                    tot_trinc = safe_int(ws.cell(r, 12).value) or (micro + trinc)
                    
                    perd_nn = safe_int(ws.cell(r, 13).value)
                    tot_destinos = df_val + ns_val + id_val + dg_val + sj_val + lp_val + tot_trinc
                    tot_class = safe_int(ws.cell(r, 14).value) or (tot_destinos + perd_nn)
                    saldo_fim = safe_int(ws.cell(r, 15).value) or (saldo_ini - tot_class)

                    nn_dict = base_dict.copy()
                    nn_dict.update({
                        "Origem": "Ninho Normal", "Saldo_Inicial": saldo_ini, "Sujo": sj_val, "Deformado": df_val,
                        "Ninho_Sujo": ns_val, "Industrial": id_val, "Duas_Gemas": dg_val, "Ovo_Limpo": lp_val, 
                        "Trincado": tot_trinc, "Microtrincado": micro, "Total_Destinos": tot_destinos,
                        "Total_Classificado": tot_class, "Saldo_Final": saldo_fim,
                        "P_NN": perd_nn
                    })
                    linhas.append(nn_dict)

            elif secao_atual == "LOTES_VELHOS":
                ns_ini = safe_int(ws.cell(r, 3).value)
                cm_ini = safe_int(ws.cell(r, 4).value)

                if ns_ini > 0:
                    ns_dict = base_dict.copy()
                    ns_dict.update({
                        "Origem": "Ninho Sujo", "Saldo_Inicial": ns_ini, "Sujo": ns_ini, "Deformado": 0, "Ninho_Sujo": 0,
                        "Industrial": 0, "Duas_Gemas": 0, "Ovo_Limpo": 0, "Trincado": 0, "Microtrincado": 0,
                        "Total_Destinos": ns_ini, "Total_Classificado": ns_ini, "Saldo_Final": 0
                    })
                    linhas.append(ns_dict)

                if cm_ini > 0:
                    cm_dict = base_dict.copy()
                    cm_dict.update({
                        "Origem": "Cama", "Saldo_Inicial": cm_ini, "Sujo": cm_ini, "Deformado": 0, "Ninho_Sujo": 0,
                        "Industrial": 0, "Duas_Gemas": 0, "Ovo_Limpo": 0, "Trincado": 0, "Microtrincado": 0,
                        "Total_Destinos": cm_ini, "Total_Classificado": cm_ini, "Saldo_Final": 0
                    })
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
                val_b = ws_p.cell(r, 2).value
                lote_num = str(int(float(val_b))) if val_b and str(val_b).replace('.', '').replace(',', '').isnumeric() else "-"
                
                v_nn = safe_int(ws_p.cell(r, 3).value)
                v_cama = safe_int(ws_p.cell(r, 4).value)
                v_ns = safe_int(ws_p.cell(r, 5).value)
                v_def = safe_int(ws_p.cell(r, 6).value)
                v_sujo = safe_int(ws_p.cell(r, 7).value)
                v_ind = safe_int(ws_p.cell(r, 8).value)
                v_dg = safe_int(ws_p.cell(r, 9).value)
                v_trinc = safe_int(ws_p.cell(r, 10).value)
                    
                if (v_nn + v_cama + v_ns + v_def + v_sujo + v_ind + v_dg + v_trinc) > 0:
                    linhas.append({
                        "Data_Processamento": data_proc,
                        "Data_Postura": dt_post, "Lote": lote_num, "Origem": "Ovos Perdidos (Aba)",
                        "Saldo_Inicial": 0, "Sujo": 0, "Deformado": 0, "Ninho_Sujo": 0,
                        "Industrial": 0, "Duas_Gemas": 0, "Ovo_Limpo": 0, "Trincado": 0,
                        "Microtrincado": 0, "Total_Destinos": 0, "Total_Classificado": 0, "Saldo_Final": 0, "Responsavel": str(resp),
                        "P_NN": v_nn, "P_Cama": v_cama, "P_NS": v_ns, "P_Def": v_def, 
                        "P_Sujo": v_sujo, "P_Ind": v_ind, "P_DG": v_dg, "P_Trinc": v_trinc
                    })
            except Exception:
                pass

    return pd.DataFrame(linhas)

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
                dt_post = df.iloc[r, 0]
                lote = df.iloc[r, 1]
                tipo = df.iloc[r, 4]
                caixas = df.iloc[r, 6]
                ovos = df.iloc[r, 7]
                if pd.notna(dt_post) and pd.notna(lote) and pd.notna(ovos):
                    try:
                        qtd_ovos = int(float(ovos))
                        if qtd_ovos > 0:
                            vendas.append({
                                "Data_Expedicao": data_exp,
                                "Cliente": str(cliente_nome), "Destino": str(destino),
                                "Data_Postura": normalizar_data_str(dt_post),
                                "Lote": str(int(float(lote))), "Tipo_Ovo": str(tipo),
                                "Caixas_360": float(caixas), "Quantidade_Ovos": qtd_ovos
                            })
                    except (ValueError, TypeError):
                        pass
    return pd.DataFrame(vendas)

def parse_transferencias_incubatorio(file_bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    transf = []
    for aba in xls.sheet_names:
        if aba.upper() in ["DICAS", "VEICULOS"]: continue
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=aba)
        if len(df) > 8 and df.shape[1] >= 8:
            origem = df.iloc[0, 3]
            destino = df.iloc[1, 3]
            data_exp = normalizar_data_str(df.iloc[4, 3])
            for r in range(9, len(df)):
                palete = df.iloc[r, 1]
                lote = df.iloc[r, 2]
                dt_post = df.iloc[r, 3]
                qtd = df.iloc[r, 4]
                categoria = df.iloc[r, 7]
                obs = df.iloc[r, 8] if df.shape[1] > 8 else ""
                if pd.notna(lote) and pd.notna(qtd):
                    try:
                        qtd_ovos = int(float(qtd))
                        if qtd_ovos > 0:
                            transf.append({
                                "Aba_Destino": aba, "Origem": str(origem), "Destino": str(destino),
                                "Data_Expedicao": data_exp,
                                "Palete": str(palete), "Lote": str(int(float(lote))),
                                "Data_Postura": normalizar_data_str(dt_post),
                                "Categoria": str(categoria), "Quantidade_Ovos": qtd_ovos,
                                "Observacao": str(obs) if pd.notna(obs) else ""
                            })
                    except (ValueError, TypeError):
                        pass
    return pd.DataFrame(transf)

def parse_reclassificacao_mtech(file_bytes):
    df = pd.read_excel(io.BytesIO(file_bytes))
    df.columns = [str(c).strip() for c in df.columns]
    
    lote_col = next((c for c in df.columns if "LOTE" in c.upper()), "Nº Lote")
    ovos_col = next((c for c in df.columns if "OVO" in c.upper() and "TRANS" not in c.upper() and "COD" not in c.upper()), "Ovos")
    prod_dt_col = next((c for c in df.columns if "PRODU" in c.upper() or "POSTURA" in c.upper()), "Data Produção")
    trans_dt_col = next((c for c in df.columns if "TRANS" in c.upper() or "MOVIM" in c.upper()), "Data Transação")

    df = df.dropna(subset=[lote_col, ovos_col])
    df["Nº Lote"] = df[lote_col].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    df["Ovos"] = pd.to_numeric(df[ovos_col], errors='coerce').fillna(0).astype(int)
    df["Data Produção"] = df[prod_dt_col].apply(normalizar_data_str)
    df["Data Transação"] = df[trans_dt_col].apply(normalizar_data_str)
    return df

def parse_descarte_mtech(file_bytes):
    df = pd.read_excel(io.BytesIO(file_bytes))
    df.columns = [str(c).strip() for c in df.columns]
    
    lote_col = next((c for c in df.columns if "LOTE" in c.upper()), "Nº Lote")
    ovos_col = next((c for c in df.columns if "OVO" in c.upper() and "TRANS" not in c.upper() and "COD" not in c.upper()), "Ovos")
    prod_dt_col = next((c for c in df.columns if "PRODU" in c.upper() or "POSTURA" in c.upper()), "Data Produção")
    trans_dt_col = next((c for c in df.columns if "TRANS" in c.upper() or "MOVIM" in c.upper()), "Data Transação")
    prod_col = next((c for c in df.columns if "PRODUTO" in c.upper() or "ITEM" in c.upper() or "DESCR" in c.upper()), "Produto")

    df = df.dropna(subset=[ovos_col])
    if lote_col in df.columns:
        df["Nº Lote"] = df[lote_col].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    else:
        df["Nº Lote"] = "-"
        
    df["Produto"] = df[prod_col]
    df["Ovos"] = pd.to_numeric(df[ovos_col], errors='coerce').fillna(0).astype(int)
    df["Data Produção"] = df[prod_dt_col].apply(normalizar_data_str)
    df["Data Transação"] = df[trans_dt_col].apply(normalizar_data_str)
    return df

def inspecionar_e_processar_conteudo(uploaded_file):
    file_bytes = uploaded_file.getvalue()
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    abas_lower = [s.lower().strip() for s in xls.sheet_names]

    if any("classifica" in s for s in abas_lower):
        df = parse_reclassificacao_cpo(file_bytes)
        DatabaseManager.salvar("df_reclass", df)
        return "Ficha de Reclassificação CPO", len(df)

    if any(s in ["dicas", "veiculos", "veículos"] for s in abas_lower):
        df = parse_transferencias_incubatorio(file_bytes)
        DatabaseManager.salvar("df_transf", df)
        return "Romaneio de Transferência OBI", len(df)

    df_amostra = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0)
    cols_upper = [str(c).strip().upper() for c in df_amostra.columns]

    if any("TRANS OVO" in c for c in cols_upper):
        df = parse_reclassificacao_mtech(file_bytes)
        DatabaseManager.salvar("df_rec_mtech", df)
        return "MTech - Lançamento de Reclassificação", len(df)

    if any("TRANSA" in c for c in cols_upper) or (any(c == "PRODUTO" for c in cols_upper) and any("OVO" in c for c in cols_upper)):
        df = parse_descarte_mtech(file_bytes)
        DatabaseManager.salvar("df_desc_mtech", df)
        return "MTech - Lançamento de Descartes", len(df)

    if any("FAZENDA" in c for c in cols_upper) and any("LOTE" in c for c in cols_upper):
        df = parse_producao_erp(file_bytes)
        DatabaseManager.salvar("df_prod", df)
        return "MTech - Produção Diária dos Galpões", len(df)

    texto_interno = df_amostra.iloc[:10].to_string().upper() if not df_amostra.empty else ""
    if "PALETE" in texto_interno:
        df = parse_transferencias_incubatorio(file_bytes)
        DatabaseManager.salvar("df_transf", df)
        return "Romaneio de Transferência OBI", len(df)
    
    if "CLIENTE" in texto_interno or "CAIXAS" in texto_interno:
        df = parse_vendas_romaneios(file_bytes)
        DatabaseManager.salvar("df_vendas", df)
        return "Romaneio de Venda Externa", len(df)

    raise ValueError(f"Colunas internas não reconhecidas: {list(df_amostra.columns)}")

# ==============================================================================
# 4. RENDERIZADORES DE TABELA FORMATADA (NÚMEROS EXPANDIDOS & NÍTIDOS)
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
            f'<td style="padding: 10px 12px; font-weight: 600; color: {origem_color}; border-right: 1px solid #E2E8F0;">{r["Origem"]}</td>'
            f'<td style="padding: 10px 12px; font-weight: 700; text-align: right; border-right: 2px solid #CBD5E1;">{r["Saldo_Inicial"]:,}'.replace(",", ".") + f'</td>'
            f'<td style="padding: 10px 12px; text-align: right; border-right: 1px solid #E2E8F0; font-weight: 600;">{r["Sujo"]:,}'.replace(",", ".") + f'</td>'
            f'<td style="padding: 10px 12px; text-align: right; border-right: 1px solid #E2E8F0; font-weight: 600;">{r["Deformado"]:,}'.replace(",", ".") + f'</td>'
            f'<td style="padding: 10px 12px; text-align: right; border-right: 1px solid #E2E8F0; font-weight: 600;">{r["Ninho_Sujo"]:,}'.replace(",", ".") + f'</td>'
            f'<td style="padding: 10px 12px; text-align: right; border-right: 1px solid #E2E8F0; font-weight: 600;">{r["Industrial"]:,}'.replace(",", ".") + f'</td>'
            f'<td style="padding: 10px 12px; text-align: right; border-right: 1px solid #E2E8F0; font-weight: 600;">{r["Duas_Gemas"]:,}'.replace(",", ".") + f'</td>'
            f'<td style="padding: 10px 12px; text-align: right; border-right: 1px solid #E2E8F0; font-weight: 600;">{r["Ovo_Limpo"]:,}'.replace(",", ".") + f'</td>'
            f'<td style="padding: 10px 12px; font-weight: 700; text-align: right; border-right: 2px solid #CBD5E1;">{r["Trincado"]:,}'.replace(",", ".") + f'</td>'
            f'<td style="padding: 10px 12px; font-weight: 700; text-align: right; background-color: #F1F5F9; border-right: 1px solid #E2E8F0;">{r["Total_Destinos"]:,}'.replace(",", ".") + f'</td>'
            f'<td style="padding: 10px 12px; font-weight: 700; text-align: right; background-color: ' + ("#FEF2F2" if r["Saldo_Final"] > 0 else "#ECFDF5") + '; color: ' + ("#991B1B" if r["Saldo_Final"] > 0 else "#065F46") + f';">{r["Saldo_Final"]:,}'.replace(",", ".") + f'</td>'
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

    tbody_content = "".join(rows_html) + row_total

    table_html = (
        '<div style="overflow-x: auto; width: 100%; border-radius: 10px; border: 1px solid #CBD5E1; margin-bottom: 22px;">'
        '<table style="min-width: 1300px; width: 100%; border-collapse: collapse; text-align: center;">'
        '<thead>'
        '<tr style="color: white; font-weight: 700; font-size: 13px;">'
        '<th colspan="2" style="background-color: #334155; padding: 10px; border-right: 2px solid white;">IDENTIFICAÇÃO</th>'
        '<th style="background-color: #1E3A8A; padding: 10px; border-right: 2px solid white;">ORIGEM</th>'
        '<th style="background-color: #1E3A8A; padding: 10px; border-right: 2px solid white;">SALDO INI</th>'
        '<th colspan="7" style="background-color: #14532D; padding: 10px; border-right: 2px solid white;">DESTINO COMERCIAL</th>'
        '<th colspan="2" style="background-color: #475569; padding: 10px;">CONFERÊNCIA</th>'
        '</tr>'
        '<tr style="color: white; font-weight: 600; font-size: 13.5px;">'
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
        f'<tbody>{tbody_content}</tbody>'
        '</table></div>'
    )
    return table_html

def render_descarte_table_html(df):
    rows_html = []
    for idx, r in df.iterrows():
        bg = "#FFF5F5" if idx % 2 == 0 else "#FFFFFF"
        rows_html.append(
            f'<tr style="background-color: {bg}; border-bottom: 1px solid #FEE2E2; color: #1E293B; font-size: 14.5px;">'
            f'<td style="padding: 10px 12px; border-right: 1px solid #FEE2E2;">{r["Data_Postura"]}</td>'
            f'<td style="padding: 10px 12px; font-weight: 700; border-right: 2px solid #CBD5E1;">{r["Lote"]}</td>'
            f'<td style="padding: 10px 12px; text-align: right; border-right: 1px solid #FEE2E2; font-weight: 700; color: #1E3A8A;">{r["Ninho Normal"]:,}</td>'
            f'<td style="padding: 10px 12px; text-align: right; border-right: 1px solid #FEE2E2; font-weight: 600;">{r["Cama"]:,}</td>'
            f'<td style="padding: 10px 12px; text-align: right; border-right: 1px solid #FEE2E2; font-weight: 600;">{r["Ninho Sujo"]:,}</td>'
            f'<td style="padding: 10px 12px; text-align: right; border-right: 1px solid #FEE2E2; font-weight: 600;">{r["Deformado"]:,}</td>'
            f'<td style="padding: 10px 12px; text-align: right; border-right: 1px solid #FEE2E2; font-weight: 600;">{r["Sujo Granja"]:,}</td>'
            f'<td style="padding: 10px 12px; text-align: right; border-right: 1px solid #FEE2E2; font-weight: 600;">{r["Industrial"]:,}</td>'
            f'<td style="padding: 10px 12px; text-align: right; border-right: 1px solid #FEE2E2; font-weight: 600;">{r["Duas Gemas"]:,}</td>'
            f'<td style="padding: 10px 12px; text-align: right; border-right: 1px solid #FEE2E2; font-weight: 600;">{r["Trincado"]:,}</td>'
            f'<td style="padding: 10px 12px; font-weight: 800; text-align: right; color: #991B1B; border-left: 2px solid #CBD5E1;">{r["TOTAL BAIXA"]:,}</td>'
            f'</tr>'.replace(",", ".")
        )
    
    t_nn, t_cama, t_ns = df["Ninho Normal"].sum(), df["Cama"].sum(), df["Ninho Sujo"].sum()
    t_def, t_sujog, t_ind = df["Deformado"].sum(), df["Sujo Granja"].sum(), df["Industrial"].sum()
    t_dg, t_trinc, t_total = df["Duas Gemas"].sum(), df["Trincado"].sum(), df["TOTAL BAIXA"].sum()
    
    row_total = (
        f'<tr style="background-color: #991B1B; color: #FFFFFF; font-weight: 800; font-size: 15px;">'
        f'<td colspan="2" style="padding: 12px; text-align: right; border-right: 2px solid #CBD5E1;">TOTAL BAIXAS DA SELEÇÃO</td>'
        f'<td style="padding: 12px; text-align: right; border-right: 1px solid #FCA5A5;">{t_nn:,}</td>'
        f'<td style="padding: 12px; text-align: right; border-right: 1px solid #FCA5A5;">{t_cama:,}</td>'
        f'<td style="padding: 12px; text-align: right; border-right: 1px solid #FCA5A5;">{t_ns:,}</td>'
        f'<td style="padding: 12px; text-align: right; border-right: 1px solid #FCA5A5;">{t_def:,}</td>'
        f'<td style="padding: 12px; text-align: right; border-right: 1px solid #FCA5A5;">{t_sujog:,}</td>'
        f'<td style="padding: 12px; text-align: right; border-right: 1px solid #FCA5A5;">{t_ind:,}</td>'
        f'<td style="padding: 12px; text-align: right; border-right: 1px solid #FCA5A5;">{t_dg:,}</td>'
        f'<td style="padding: 12px; text-align: right; border-right: 1px solid #FCA5A5;">{t_trinc:,}</td>'
        f'<td style="padding: 12px; text-align: right; border-left: 2px solid #CBD5E1;">{t_total:,}</td>'
        f'</tr>'.replace(",", ".")
    )
    
    tbody_content = "".join(rows_html) + row_total
    
    table_html = (
        '<div style="overflow-x: auto; width: 100%; border-radius: 10px; border: 1px solid #FECACA; margin-bottom: 22px;">'
        '<table style="min-width: 1050px; width: 100%; border-collapse: collapse; text-align: center;">'
        '<thead>'
        '<tr style="color: white; font-weight: 700; font-size: 13.5px; background-color: #334155;">'
        '<th style="padding: 10px; border-right: 1px solid #64748B;">Data Postura</th>'
        '<th style="padding: 10px; border-right: 2px solid white;">Lote</th>'
        '<th style="padding: 10px; border-right: 1px solid #64748B;">Perdido (NN) / Ninho Normal</th>'
        '<th style="padding: 10px; border-right: 1px solid #64748B;">Cama</th>'
        '<th style="padding: 10px; border-right: 1px solid #64748B;">Ninho Sujo</th>'
        '<th style="padding: 10px; border-right: 1px solid #64748B;">Deformado</th>'
        '<th style="padding: 10px; border-right: 1px solid #64748B;">Sujo Granja</th>'
        '<th style="padding: 10px; border-right: 1px solid #64748B;">Industrial</th>'
        '<th style="padding: 10px; border-right: 1px solid #64748B;">Duas Gemas</th>'
        '<th style="padding: 10px; border-right: 2px solid white;">Trincado</th>'
        '<th style="padding: 10px; background-color: #991B1B;">TOTAL BAIXA</th>'
        '</tr>'
        '</thead>'
        f'<tbody>{tbody_content}</tbody>'
        '</table></div>'
    )
    return table_html

# ==============================================================================
# 5. BARRA LATERAL EXECUTIVA
# ==============================================================================
st.sidebar.markdown("""
<div class="brand-box">
    <div class="brand-icon">🥚</div>
    <div>
        <p class="brand-title">Balsas 117</p>
        <span class="brand-sub">Controle Operacional</span>
    </div>
</div>
""", unsafe_allow_html=True)

st.sidebar.markdown('<div class="nav-label">Menu de Navegação</div>', unsafe_allow_html=True)

tela = st.sidebar.radio(
    "Navegação:",
    [
        "📊 Dashboard Executivo & Faturamento",
        "📥 Central de Carga de Arquivos",
        "📋 Validação de Reclassificação & Descartes",
        "🚚 Validação de Expedição (Vendas & Transferências)",
        "🏭 Produção Diária (MTech)",
        "🗑️ Descartes & Trincados"
    ]
)

st.sidebar.markdown('<div class="nav-label">Cotação & Parâmetros</div>', unsafe_allow_html=True)

preco_unitario = st.sidebar.number_input(
    "Preço por Unidade (R$):",
    min_value=0.000,
    value=2.380,
    step=0.010,
    format="%.3f"
)

valor_caixa_calc = preco_unitario * 360.0
val_cx_str = f"R$ {valor_caixa_calc:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

st.sidebar.markdown(f"""
<div class="cotacao-sidebar-card">
    <div class="cotacao-sidebar-tag">Caixa Fechada (360 und)</div>
    <div class="cotacao-sidebar-val">{val_cx_str}</div>
</div>
""", unsafe_allow_html=True)

# ==============================================================================
# 6. CONTROLADORES DAS TELAS
# ==============================================================================

if tela == "📊 Dashboard Executivo & Faturamento":
    st.markdown("""
    <div class="top-header-bar">
        <div class="top-header-title">
            <span>Visão Consolidada • Banco SQLite Ativo</span>
            <h1>Painel Geral de Operações & Faturamento</h1>
        </div>
        <div class="top-status-badge">● Banco Local Conectado</div>
    </div>
    """, unsafe_allow_html=True)

    df_prod = DatabaseManager.obter("df_prod")
    df_reclass = DatabaseManager.obter("df_reclass")
    df_vendas = DatabaseManager.obter("df_vendas")
    df_transf = DatabaseManager.obter("df_transf")

    tot_prod = df_prod["Ovos"].sum() if not df_prod.empty else 0
    tot_saldo_nn = df_reclass[df_reclass["Origem"] == "Ninho Normal"]["Saldo_Final"].sum() if not df_reclass.empty else 0
    tot_ovos_vendidos = df_vendas["Quantidade_Ovos"].sum() if not df_vendas.empty else 0
    tot_caixas_venda = df_vendas["Caixas_360"].sum() if not df_vendas.empty else 0
    faturamento_vendas = tot_ovos_vendidos * preco_unitario

    kpis_html = f"""
    <div class="kpi-container">
        <div class="kpi-card">
            <div class="kpi-icon-bubble bubble-green">🐔</div>
            <div>
                <p class="kpi-title">Produção Galpões</p>
                <p class="kpi-value">{tot_prod:,.0f}</p>
                <p class="kpi-sub" style="color: #16A34A;">Acumulado no Banco</p>
            </div>
        </div>
        <div class="kpi-card">
            <div class="kpi-icon-bubble bubble-teal">🥚</div>
            <div>
                <p class="kpi-title">Saldo NN Disponível</p>
                <p class="kpi-value">{tot_saldo_nn:,.0f}</p>
                <p class="kpi-sub" style="color: #0284C7;">Ovo Ninho Normal</p>
            </div>
        </div>
        <div class="kpi-card">
            <div class="kpi-icon-bubble bubble-yellow">📦</div>
            <div>
                <p class="kpi-title">Ovos Vendidos</p>
                <p class="kpi-value">{tot_ovos_vendidos:,.0f}</p>
                <p class="kpi-sub" style="color: #D97706;">{tot_caixas_venda:,.1f} caixas (360 un)</p>
            </div>
        </div>
        <div class="kpi-card">
            <div class="kpi-icon-bubble bubble-yellow">💰</div>
            <div>
                <p class="kpi-title">Faturamento Total</p>
                <p class="kpi-value">R$ {faturamento_vendas:,.2f}</p>
                <p class="kpi-sub" style="color: #B45309;">R$ {preco_unitario:.3f} / ovo</p>
            </div>
        </div>
    </div>
    """.replace(",", ".")
    st.markdown(kpis_html, unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="saas-panel">', unsafe_allow_html=True)
        st.subheader("📦 Vendas Externas por Cliente")
        if not df_vendas.empty:
            venda_cli = df_vendas.groupby("Cliente")[["Caixas_360", "Quantidade_Ovos"]].sum().reset_index()
            venda_cli["Faturamento (R$)"] = venda_cli["Quantidade_Ovos"] * preco_unitario
            st.dataframe(venda_cli, use_container_width=True, hide_index=True)
        else:
            st.info("Nenhuma venda gravada no histórico.")
        st.markdown('</div>', unsafe_allow_html=True)

    with c2:
        st.markdown('<div class="saas-panel">', unsafe_allow_html=True)
        st.subheader("🔁 Transferências OBI por Incubatório")
        if not df_transf.empty:
            transf_dest = df_transf.groupby("Destino")[["Quantidade_Ovos"]].sum().reset_index()
            st.dataframe(transf_dest, use_container_width=True, hide_index=True)
        else:
            st.info("Nenhuma transferência gravada no histórico.")
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
    st.subheader("🚀 Importador Inteligente (Leitura Automática pelo Conteúdo Interno)")
    st.markdown(
        '<div class="destaque-verde">'
        '<b>Solte qualquer planilha abaixo:</b> O sistema <b>não depende do nome do arquivo</b>. Ele abre cada planilha, '
        'inspeciona as abas, células e cabeçalhos internos, identifica automaticamente o tipo do documento e grava no banco de dados sem duplicar registros.'
        '</div>',
        unsafe_allow_html=True
    )

    arquivos_unificados = st.file_uploader(
        "Selecione ou arraste todas as planilhas (Fichas CPO, Romaneios, MTech Produção, Reclassificação ou Descartes):",
        type=["xls", "xlsx"],
        accept_multiple_files=True,
        key="uploader_inteligente"
    )

    if arquivos_unificados:
        for arq in arquivos_unificados:
            try:
                tipo_identificado, qtd_reg = inspecionar_e_processar_conteudo(arq)
                st.success(f"✅ Conteúdo identificado como **{tipo_identificado}** — {qtd_reg} registros sincronizados no banco!")
            except Exception as e:
                st.error(f"❌ Não foi possível identificar a estrutura interna de um dos arquivos: {e}")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="saas-panel">', unsafe_allow_html=True)
    st.subheader("🗄️ Status do Histórico Salvo no Banco Local (SQLite)")
    
    resumo_db = [
        {"Tabela no Banco": "Fichas de Reclassificação & Perdas (CPO)", "Código": "df_reclass", "Registros Gravados": len(DatabaseManager.obter("df_reclass"))},
        {"Tabela no Banco": "MTech - Lançamentos de Reclassificação", "Código": "df_rec_mtech", "Registros Gravados": len(DatabaseManager.obter("df_rec_mtech"))},
        {"Tabela no Banco": "MTech - Lançamentos de Descartes", "Código": "df_desc_mtech", "Registros Gravados": len(DatabaseManager.obter("df_desc_mtech"))},
        {"Tabela no Banco": "MTech - Produção Diária Galpões", "Código": "df_prod", "Registros Gravados": len(DatabaseManager.obter("df_prod"))},
        {"Tabela no Banco": "Romaneios de Venda Externa", "Código": "df_vendas", "Registros Gravados": len(DatabaseManager.obter("df_vendas"))},
        {"Tabela no Banco": "Romaneios de Transferência OBI", "Código": "df_transf", "Registros Gravados": len(DatabaseManager.obter("df_transf"))},
    ]
    st.dataframe(pd.DataFrame(resumo_db), use_container_width=True, hide_index=True)

    with st.expander("🛠️ Manutenção do Banco de Dados (Limpar Tabelas ou Baixar Backup)"):
        c_db1, c_db2 = st.columns(2)
        with c_db1:
            tab_para_limpar = st.selectbox(
                "Selecione uma base caso deseje zerar seus registros:",
                ["Selecione...", "df_reclass", "df_rec_mtech", "df_desc_mtech", "df_prod", "df_vendas", "df_transf", "TODAS AS TABELAS"]
            )
            if st.button("🗑️ Limpar Base Selecionada"):
                if tab_para_limpar == "TODAS AS TABELAS":
                    for t in ["df_reclass", "df_rec_mtech", "df_desc_mtech", "df_prod", "df_vendas", "df_transf"]:
                        DatabaseManager.limpar_tabela(t)
                    st.success("Todas as tabelas foram zeradas no banco local.")
                    st.rerun()
                elif tab_para_limpar != "Selecione...":
                    DatabaseManager.limpar_tabela(tab_para_limpar)
                    st.success(f"Tabela {tab_para_limpar} limpa com sucesso!")
                    st.rerun()
        with c_db2:
            if os.path.exists(DB_PATH):
                with open(DB_PATH, "rb") as f_db:
                    st.download_button(
                        "💾 Baixar Backup do Banco SQLite (.db)",
                        data=f_db.read(),
                        file_name="historico_ovos_117.db",
                        mime="application/octet-stream"
                    )
    st.markdown('</div>', unsafe_allow_html=True)

# ==============================================================================
# TELA: VALIDAÇÃO DE RECLASSIFICAÇÃO & DESCARTES
# ==============================================================================
elif tela == "📋 Validação de Reclassificação & Descartes":
    st.markdown("""
    <div class="top-header-bar">
        <div class="top-header-title">
            <span>Procedimento Operacional & Validação</span>
            <h1>Validação de Reclassificação & Descartes</h1>
        </div>
    </div>
    """, unsafe_allow_html=True)

    df_r_original = DatabaseManager.obter("df_reclass")
    df_p_original = DatabaseManager.obter("df_prod")
    df_rec_mtech_orig = DatabaseManager.obter("df_rec_mtech")
    df_desc_mtech_orig = DatabaseManager.obter("df_desc_mtech")

    if df_r_original.empty:
        st.warning("⚠️ Carregue a ficha de reclassificação na aba 'Central de Carga de Arquivos' para visualizar.")
    else:
        st.markdown("""
        <div style="background-color: #F0F9FF; border: 1px solid #BAE6FD; border-left: 6px solid #0284C7; border-radius: 8px; padding: 16px 20px; margin-bottom: 24px;">
            <h4 style="margin: 0 0 10px 0; color: #0369A1; font-size: 15px; font-weight: 700;">⚙️ Painel de Filtros (Histórico do Banco de Dados)</h4>
            <p style="margin: 0; font-size: 13.5px; color: #0C4A6E;">A <b>Data de Transação (Processamento)</b> sincroniza os lotes auditados no MTech com a ficha física correspondente gravada no banco.</p>
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
            st.warning("Nenhum dado encontrado para esta combinação de filtros.")
            st.stop()

        # 1. LANÇAMENTO DE RECLASSIFICAÇÃO
        st.markdown('<div class="saas-panel">', unsafe_allow_html=True)
        st.subheader("1. Lançamento de Reclassificação — Destinos Comerciais")
        
        df_comercial = df_r[df_r["Origem"].isin(["Ninho Normal", "Ninho Sujo", "Cama"])].copy()
        colunas_espelho = [
            "Data_Postura", "Lote", "Origem", "Saldo_Inicial", "Sujo", "Deformado", 
            "Ninho_Sujo", "Industrial", "Duas_Gemas", "Ovo_Limpo", "Trincado", 
            "Total_Destinos", "Saldo_Final"
        ]
        df_comercial_limpo = df_comercial[colunas_espelho]
        
        tab_v1, tab_v2 = st.tabs(["🎨 Visão Formatada (Com Totais)", "📊 Tabela Interativa (Planilha Limpa)"])
        with tab_v1:
            st.markdown(render_reclass_table_html(df_comercial), unsafe_allow_html=True)
        with tab_v2:
            st.dataframe(df_comercial_limpo, use_container_width=True, hide_index=True)
        st.markdown('</div>', unsafe_allow_html=True)

        # 2. LANÇAMENTO DE DESCARTES COM LOTE REAL E PERDIDAS (NN)
        st.markdown('<div class="saas-panel">', unsafe_allow_html=True)
        st.subheader("2. Lançamento de Baixa & Descarte (MTech)")
        
        cols_base_desc = ["Data_Postura", "Lote", "P_NN", "Deformado", "Trincado"]
        df_baixa_class = df_r[df_r["Origem"].isin(["Ninho Normal", "Ninho Sujo", "Cama"])][cols_base_desc].copy()
        df_baixa_class = df_baixa_class.rename(columns={"P_NN": "Ninho Normal"})
        
        cols_perd = ["P_NN", "P_Cama", "P_NS", "P_Def", "P_Sujo", "P_Ind", "P_DG", "P_Trinc"]
        df_baixa_perd = df_r[df_r["Origem"] == "Ovos Perdidos (Aba)"][["Data_Postura", "Lote"] + cols_perd].copy()
        df_baixa_perd = df_baixa_perd.rename(columns={
            "P_NN": "Ninho Normal", "P_Cama": "Cama", "P_NS": "Ninho Sujo",
            "P_Def": "Deformado", "P_Sujo": "Sujo Granja", "P_Ind": "Industrial",
            "P_DG": "Duas Gemas", "P_Trinc": "Trincado"
        })
        
        df_desc_final = pd.concat([df_baixa_class, df_baixa_perd], ignore_index=True).fillna(0)
        
        all_cols = ["Ninho Normal", "Cama", "Ninho Sujo", "Deformado", "Sujo Granja", "Industrial", "Duas Gemas", "Trincado"]
        for c in all_cols:
            if c not in df_desc_final.columns:
                df_desc_final[c] = 0
            df_desc_final[c] = df_desc_final[c].astype(int)

        df_desc_final = df_desc_final.groupby(["Data_Postura", "Lote"])[all_cols].sum().reset_index()
        df_desc_final["TOTAL BAIXA"] = df_desc_final[all_cols].sum(axis=1)
        df_desc_final = df_desc_final[df_desc_final["TOTAL BAIXA"] > 0]
        df_desc_final = df_desc_final.sort_values(by=["Data_Postura", "Lote"], ascending=[False, True])

        if not df_desc_final.empty:
            tab_d1, tab_d2 = st.tabs(["🔴 Visão Formatada (Detalhamento Completo)", "📊 Tabela Interativa"])
            with tab_d1:
                st.markdown(render_descarte_table_html(df_desc_final), unsafe_allow_html=True)
            with tab_d2:
                st.dataframe(df_desc_final, use_container_width=True, hide_index=True)
        else:
            st.info("Nenhuma baixa/descarte para registrar nas abas (Comercial ou Ovos Perdidos).")
            
        st.markdown('</div>', unsafe_allow_html=True)

        # 3. RESUMOS OPERACIONAIS
        st.markdown('<div class="saas-panel">', unsafe_allow_html=True)
        st.subheader("3. Resumo de Fechamento por Data de Postura")
        col_res1, col_res2 = st.columns(2)
        with col_res1:
            st.markdown("**🥚 Resumo de Trincados para Faturamento**")
            res_trinc = df_r[df_r["Origem"] != "Ovos Perdidos (Aba)"].groupby("Data_Postura").agg({"Microtrincado": "sum"}).reset_index()
            st.dataframe(res_trinc, use_container_width=True, hide_index=True)

        with col_res2:
            st.markdown("**🧹 Resumo Consolidado de Ovos Sujos**")
            res_sujo = df_r[df_r["Origem"] != "Ovos Perdidos (Aba)"].groupby("Data_Postura").agg({"Sujo": "sum"}).reset_index()
            res_sujo.rename(columns={"Sujo": "Total Sujo"}, inplace=True)
            st.dataframe(res_sujo, use_container_width=True, hide_index=True)
        st.markdown('</div>', unsafe_allow_html=True)

        # ======================================================================
        # 4.1 VALIDAÇÃO EXCLUSIVA: RECLASSIFICAÇÃO (FICHA CPO vs MTECH DÉBITO)
        # ======================================================================
        st.markdown('<div class="saas-panel" style="border-top: 4px solid #1E3A8A;">', unsafe_allow_html=True)
        st.subheader("4.1 Validação de Reclassificação (Destinos Comerciais)")
        
        if df_rec_mtech_orig.empty:
            st.info("ℹ️ Para validar a reclassificação, solte a planilha de Reclassificação do MTech na Central de Carga.")
        elif filtro_proc == "Todos os Processamentos":
            st.warning("⚠️ Selecione uma **Data de Transação (Processamento)** específica no painel de filtros para ativar a validação do lote contábil exato.")
        else:
            df_rec_m_sub = df_rec_mtech_orig.copy()
            df_rec_m_sub = df_rec_m_sub[df_rec_m_sub["Data Transação"] == filtro_proc]
            
            if filtro_dt != "Todas as Datas":
                df_rec_m_sub = df_rec_m_sub[df_rec_m_sub["Data Produção"] == filtro_dt]

            df_rec_m_sub["Categoria_Normalizada"] = df_rec_m_sub["Nome Produto"].apply(mapear_categoria_mtech)

            categorias_alvo = ["Sujo", "Deformado", "Ninho Sujo", "Industrial", "Duas Gemas", "Ovo Limpo", "Trincado"]
            tot_dest_ficha = df_comercial[["Sujo", "Deformado", "Ninho_Sujo", "Industrial", "Duas_Gemas", "Ovo_Limpo", "Trincado"]].sum().sum()
            
            # ATENÇÃO: No MTech, o que entra no comercial é DÉBITO de reclassificação!
            tot_cred_m = df_rec_m_sub[
                (df_rec_m_sub["Cód Trans Ovo"].astype(str).str.contains("D[é|e]bito", case=False, na=False)) &
                (df_rec_m_sub["Categoria_Normalizada"].isin(categorias_alvo))
            ]["Ovos"].sum()
            
            dif_reclass = tot_dest_ficha - tot_cred_m

            st.markdown(f"""
            <div class="kpi-container" style="grid-template-columns: repeat(2, 1fr); margin-top: 10px;">
                <div class="kpi-card" style="border-left: 5px solid {'#16A34A' if dif_reclass == 0 else '#DC2626'};">
                    <div class="kpi-icon-bubble {'bubble-green' if dif_reclass == 0 else 'bubble-red'}">
                        {'✓' if dif_reclass == 0 else '✕'}
                    </div>
                    <div>
                        <p class="kpi-title">Balanço Total de Reclassificação</p>
                        <p class="kpi-value" style="color: {'#16A34A' if dif_reclass == 0 else '#DC2626'}; font-size: 1.35rem;">
                            {'✅ 100% Validado' if dif_reclass == 0 else f'❌ Divergência: {dif_reclass:+,} ovos'}
                        </p>
                        <p class="kpi-sub" style="color: #64748B;">
                            Ficha CPO: {tot_dest_ficha:,.0f} | MTech (Entrada): {tot_cred_m:,.0f}
                        </p>
                    </div>
                </div>
                <div class="kpi-card" style="border-left: 5px solid #0284C7;">
                    <div class="kpi-icon-bubble bubble-teal">📋</div>
                    <div>
                        <p class="kpi-title">Transação Auditada</p>
                        <p class="kpi-value" style="font-size: 1.25rem;">{filtro_proc}</p>
                        <p class="kpi-sub" style="color: #64748B;">
                            {'Todas as posturas consolidadas' if filtro_dt == 'Todas as Datas' else f'Postura: {filtro_dt}'}
                        </p>
                    </div>
                </div>
            </div>
            """.replace(",", "."), unsafe_allow_html=True)

            cats_comerciais = [
                ("Sujo (02 / 06)", "Sujo", df_comercial["Sujo"].sum()),
                ("Deformado (12)", "Deformado", df_comercial["Deformado"].sum()),
                ("Ninho Sujo (02)", "Ninho Sujo", df_comercial["Ninho_Sujo"].sum()),
                ("Industrial (08)", "Industrial", df_comercial["Industrial"].sum()),
                ("Duas Gemas (04)", "Duas Gemas", df_comercial["Duas_Gemas"].sum()),
                ("Ovo Limpo (13)", "Ovo Limpo", df_comercial["Ovo_Limpo"].sum()),
                ("Trincado / Micro (05)", "Trincado", df_comercial["Trincado"].sum()),
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

            st.dataframe(pd.DataFrame(linhas_validacao_reclass), use_container_width=True, hide_index=True)

            # VALIDAÇÃO DE ORIGEM (SAÍDA / CRÉDITO DO ESTOQUE)
            st.markdown("<hr>", unsafe_allow_html=True)
            st.markdown("**Validação de Saída da Reclassificação (Origens MTech)**")
            
            # Ovos que saíram do estoque na Tabela 1 e Tabela 2
            tot_saida_nn = df_r[df_r["Origem"] == "Ninho Normal"]["Total_Classificado"].sum()
            tot_saida_ns = df_r[df_r["Origem"] == "Ninho Sujo"]["Saldo_Inicial"].sum()
            tot_saida_cm = df_r[df_r["Origem"] == "Cama"]["Saldo_Inicial"].sum()

            linhas_validacao_origem = []
            for nome_origem, cat_origem, qtd_saida_ficha in [
                ("Saída de Ninho Normal (NN)", "Ninho Normal", tot_saida_nn),
                ("Saída de Ninho Sujo Velho", "Ninho Sujo", tot_saida_ns),
                ("Saída de Cama", "Cama", tot_saida_cm)
            ]:
                qtd_saida_mtech = df_rec_m_sub[
                    (df_rec_m_sub["Cód Trans Ovo"].astype(str).str.contains("Cr[é|e]dito", case=False, na=False)) & 
                    (df_rec_m_sub["Categoria_Normalizada"] == cat_origem)
                ]["Ovos"].sum()
                
                diferenca = qtd_saida_ficha - qtd_saida_mtech
                linhas_validacao_origem.append({
                    "Origem Baixada": nome_origem,
                    "Ficha CPO (Saída Esperada)": int(qtd_saida_ficha),
                    "MTech Crédito (Saída Lançada)": int(qtd_saida_mtech),
                    "Divergência": int(diferenca),
                    "Status": "✅ OK" if diferenca == 0 else f"❌ Erro ({diferenca:+,})"
                })

            st.dataframe(pd.DataFrame(linhas_validacao_origem), use_container_width=True, hide_index=True)

        st.markdown('</div>', unsafe_allow_html=True)

        # ======================================================================
        # 4.2 VALIDAÇÃO EXCLUSIVA: DESCARTES POR TIPO DE OVO (COM NN PERDIDO)
        # ======================================================================
        st.markdown('<div class="saas-panel" style="border-top: 4px solid #DC2626;">', unsafe_allow_html=True)
        st.subheader("4.2 Validação de Baixas & Descartes por Tipo de Ovo")
        
        if df_desc_mtech_orig.empty:
            st.info("ℹ️ Para validar os descartes, solte a planilha de Descarte do MTech na Central de Carga.")
        elif filtro_proc == "Todos os Processamentos":
            st.warning("⚠️ Selecione uma **Data de Transação (Processamento)** específica no painel de filtros para ativar a validação do lote contábil exato.")
        else:
            df_desc_m_sub = df_desc_mtech_orig.copy()
            df_desc_m_sub = df_desc_m_sub[df_desc_m_sub["Data Transação"] == filtro_proc]
            
            if filtro_dt != "Todas as Datas":
                df_desc_m_sub = df_desc_m_sub[df_desc_m_sub["Data Produção"] == filtro_dt]

            df_desc_m_sub["Categoria_Normalizada"] = df_desc_m_sub["Produto"].apply(mapear_categoria_mtech)

            tot_desc_esperado_cpo = df_desc_final["TOTAL BAIXA"].sum() if not df_desc_final.empty else 0
            
            tot_desc_galpao = 0
            if not df_p_original.empty:
                df_p_sub = df_p_original.copy()
                if filtro_dt != "Todas as Datas":
                    df_p_sub = df_p_sub[df_p_sub["Data Produção"] == filtro_dt]
                df_p_sub["Cat_Prod"] = df_p_sub["Nome Produto"].apply(mapear_categoria_mtech)
                tot_desc_galpao = df_p_sub[df_p_sub["Cat_Prod"].isin(["Trincado", "Ninho Normal"]) & df_p_sub["Nome Produto"].astype(str).str.contains("TRINCADO|PERDIDO|QUEBRA", case=False, na=False)]["Ovos"].sum()

            tot_esperado_geral = tot_desc_esperado_cpo + tot_desc_galpao
            tot_desc_m = df_desc_m_sub["Ovos"].sum()
            dif_desc_geral = tot_esperado_geral - tot_desc_m

            st.markdown(f"""
            <div class="kpi-container" style="grid-template-columns: repeat(3, 1fr); margin-top: 10px;">
                <div class="kpi-card" style="border-left: 5px solid {'#16A34A' if dif_desc_geral == 0 else '#DC2626'};">
                    <div class="kpi-icon-bubble {'bubble-green' if dif_desc_geral == 0 else 'bubble-red'}">
                        {'✓' if dif_desc_geral == 0 else '✕'}
                    </div>
                    <div>
                        <p class="kpi-title">Balanço Total de Descartes</p>
                        <p class="kpi-value" style="color: {'#16A34A' if dif_desc_geral == 0 else '#DC2626'}; font-size: 1.35rem;">
                            {'✅ 100% Validado' if dif_desc_geral == 0 else f'❌ Divergência: {dif_desc_geral:+,} ovos'}
                        </p>
                        <p class="kpi-sub" style="color: #64748B;">
                            Total Esperado: {tot_esperado_geral:,.0f} | MTech: {tot_desc_m:,.0f}
                        </p>
                    </div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-icon-bubble bubble-yellow">🥚</div>
                    <div>
                        <p class="kpi-title">Descartes do CPO (Fichas)</p>
                        <p class="kpi-value" style="font-size: 1.25rem;">{tot_desc_esperado_cpo:,.0f}</p>
                        <p class="kpi-sub" style="color: #64748B;">NN Perdidos + Avarias</p>
                    </div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-icon-bubble bubble-teal">🚜</div>
                    <div>
                        <p class="kpi-title">Descartes Galpão (Produção)</p>
                        <p class="kpi-value" style="font-size: 1.25rem;">{tot_desc_galpao:,.0f}</p>
                        <p class="kpi-sub" style="color: #64748B;">Baixas Diretas na Coleta</p>
                    </div>
                </div>
            </div>
            """.replace(",", "."), unsafe_allow_html=True)

            st.markdown("**Comparativo Detalhado por Tipo de Ovo (Descartes & Perdas)**")
            
            cats_descartes = [
                ("Perdido (NN) / Ninho Normal", "Ninho Normal", df_desc_final["Ninho Normal"].sum() if "Ninho Normal" in df_desc_final.columns else 0),
                ("Cama (CM)", "Cama", df_desc_final["Cama"].sum() if "Cama" in df_desc_final.columns else 0),
                ("Ninho Sujo (NS)", "Ninho Sujo", df_desc_final["Ninho Sujo"].sum() if "Ninho Sujo" in df_desc_final.columns else 0),
                ("Deformado (DF)", "Deformado", df_desc_final["Deformado"].sum() if "Deformado" in df_desc_final.columns else 0),
                ("Sujo Granja (SJ)", "Sujo", df_desc_final["Sujo Granja"].sum() if "Sujo Granja" in df_desc_final.columns else 0),
                ("Industrial (ID)", "Industrial", df_desc_final["Industrial"].sum() if "Industrial" in df_desc_final.columns else 0),
                ("Duas Gemas (DG)", "Duas Gemas", df_desc_final["Duas Gemas"].sum() if "Duas Gemas" in df_desc_final.columns else 0),
                ("Trincado / Micro (TR)", "Trincado", df_desc_final["Trincado"].sum() if "Trincado" in df_desc_final.columns else 0),
            ]

            linhas_validacao_desc = []
            for nome_exib, cat_chave, qtd_cpo in cats_descartes:
                qtd_m = df_desc_m_sub[df_desc_m_sub["Categoria_Normalizada"] == cat_chave]["Ovos"].sum()
                diferenca = qtd_cpo - qtd_m
                linhas_validacao_desc.append({
                    "Tipo de Ovo / Categoria": nome_exib,
                    "Ficha CPO (Físico)": int(qtd_cpo),
                    "MTech (Sistema)": int(qtd_m),
                    "Divergência": int(diferenca),
                    "Status": "✅ OK" if diferenca == 0 else f"❌ Erro ({diferenca:+,})"
                })

            st.dataframe(pd.DataFrame(linhas_validacao_desc), use_container_width=True, hide_index=True)

        st.markdown('</div>', unsafe_allow_html=True)

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df_comercial_limpo.to_excel(writer, sheet_name="Reclassificacao", index=False)
            if not df_desc_final.empty:
                df_desc_final.to_excel(writer, sheet_name="Descartes", index=False)
            res_trinc.to_excel(writer, sheet_name="Resumo_Trincados", index=False)
            res_sujo.to_excel(writer, sheet_name="Resumo_Sujos", index=False)

        st.download_button(
            "📥 Baixar Ficha Formatada Completa (.xlsx)",
            data=buf.getvalue(),
            file_name="Ficha_Reclassificacao_Validada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ==============================================================================
# TELA: VALIDAÇÃO DE EXPEDIÇÃO (VENDAS & TRANSFERÊNCIAS)
# ==============================================================================
elif tela == "🚚 Validação de Expedição (Vendas & Transferências)":
    st.markdown("""
    <div class="top-header-bar">
        <div class="top-header-title">
            <span>Saída de Cargas & Romaneios</span>
            <h1>Validação de Expedição (Vendas & Transferências)</h1>
        </div>
    </div>
    """, unsafe_allow_html=True)

    df_vendas = DatabaseManager.obter("df_vendas")
    df_transf = DatabaseManager.obter("df_transf")

    st.markdown('<div class="saas-panel">', unsafe_allow_html=True)
    tab1, tab2 = st.tabs(["Romaneios de Venda Externa", "Transferências OBI"])
    with tab1:
        if df_vendas.empty:
            st.warning("Nenhum romaneio de venda gravado no banco.")
        else:
            st.dataframe(df_vendas, use_container_width=True, hide_index=True)
    with tab2:
        if df_transf.empty:
            st.warning("Nenhum romaneio de transferência gravado no banco.")
        else:
            st.dataframe(df_transf, use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

elif tela == "🏭 Produção Diária (MTech)":
    st.markdown("""
    <div class="top-header-bar">
        <div class="top-header-title">
            <span>Visão Galpões</span>
            <h1>Apontamentos de Produção Diária (MTech ERP)</h1>
        </div>
    </div>
    """, unsafe_allow_html=True)

    df_p = DatabaseManager.obter("df_prod")
    if df_p.empty:
        st.warning("Nenhum dado de produção diária gravado no banco.")
    else:
        st.markdown('<div class="saas-panel">', unsafe_allow_html=True)
        c_f1, c_f2 = st.columns(2)
        with c_f1:
            nucleos = ["Todos"] + sorted(list(df_p["Nome Fazenda"].unique()))
            filtro_nuc = st.selectbox("Filtrar por Núcleo/Fazenda:", nucleos)
        with c_f2:
            lotes = ["Todos"] + sorted(list(df_p["Nº Lote"].unique()))
            filtro_lot = st.selectbox("Filtrar por Lote:", lotes)

        df_filtered = df_p.copy()
        if filtro_nuc != "Todos":
            df_filtered = df_filtered[df_filtered["Nome Fazenda"] == filtro_nuc]
        if filtro_lot != "Todos":
            df_filtered = df_filtered[df_filtered["Nº Lote"] == filtro_lot]

        cols_exib = [c for c in ["Data Produção", "Nome Fazenda", "Nº Lote", "Nome Produto", "Ovos", "Cód Ref", "Nº Rastreamento"] if c in df_filtered.columns]
        st.dataframe(df_filtered[cols_exib], use_container_width=True, hide_index=True)
        st.markdown('</div>', unsafe_allow_html=True)

elif tela == "🗑️ Descartes & Trincados":
    st.markdown("""
    <div class="top-header-bar">
        <div class="top-header-title">
            <span style="color: #DC2626 !important;">Auditoria de Perdas</span>
            <h1>Central de Perdas, Descartes & Avarias</h1>
        </div>
    </div>
    """, unsafe_allow_html=True)

    df_reclass = DatabaseManager.obter("df_reclass")
    if df_reclass.empty:
        st.warning("Carregue a planilha de reclassificação na Central de Carga para visualizar.")
    else:
        df_desc = df_reclass[(df_reclass["Origem"] == "Ovos Perdidos (Aba)") | (df_reclass["Microtrincado"] > 0) | (df_reclass["P_NN"] > 0)].copy()
        tot_micro = df_reclass['Microtrincado'].sum()
        
        cols_perda = ["P_NN", "P_Cama", "P_NS", "P_Def", "P_Sujo", "P_Ind", "P_DG", "P_Trinc"]
        tot_quebras = df_reclass[cols_perda].sum().sum()

        descartes_kpi_html = f"""
        <div class="kpi-container" style="grid-template-columns: repeat(2, 1fr);">
            <div class="kpi-card">
                <div class="kpi-icon-bubble bubble-yellow">🔍</div>
                <div>
                    <p class="kpi-title">Microtrincados (05)</p>
                    <p class="kpi-value">{tot_micro:,.0f}</p>
                    <p class="kpi-sub" style="color: #D97706;">Passível de Comercialização</p>
                </div>
            </div>
            <div class="kpi-card">
                <div class="kpi-icon-bubble bubble-red">🗑️</div>
                <div>
                    <p class="kpi-title">Quebras Físicas / Perdidos (Inclui NN)</p>
                    <p class="kpi-value" style="color: #DC2626;">{tot_quebras:,.0f}</p>
                    <p class="kpi-sub" style="color: #DC2626;">Perda Total Consolidada</p>
                </div>
            </div>
        </div>
        """.replace(",", ".")
        st.markdown(descartes_kpi_html, unsafe_allow_html=True)

        st.markdown('<div class="saas-panel">', unsafe_allow_html=True)
        st.subheader("Detalhamento por Data e Lote")
        if not df_desc.empty:
            df_desc_resumo = df_desc.groupby(["Data_Postura", "Lote"]).sum(numeric_only=True).reset_index()
            st.dataframe(df_desc_resumo, use_container_width=True, hide_index=True)
        else:
            st.info("Nenhuma avaria registrada.")
        st.markdown('</div>', unsafe_allow_html=True)