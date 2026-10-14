import streamlit as st
import pandas as pd
from datetime import date

st.set_page_config(
    page_title="Controlo e Auditoria de Ração",
    page_icon="🌾",
    layout="wide"
)

# Inicialização de bases em sessão (preparado para ligar ao Neon PostgreSQL)
if "entradas" not in st.session_state:
    st.session_state.entradas = pd.DataFrame(columns=["Data", "Tipo Racao", "Quantidade_Kg", "Documento"])

if "transferencias" not in st.session_state:
    st.session_state.transferencias = pd.DataFrame(columns=["Data", "Origem", "Destino", "Lote", "Tipo Racao", "Quantidade_Kg"])

if "consumos" not in st.session_state:
    st.session_state.consumos = pd.DataFrame(columns=["Data", "Nucleo", "Lote", "Tipo Racao", "Sexo", "Quantidade_Kg"])

st.title("🌾 Sistema Integrado de Gestão e Auditoria de Ração")

# --- BARRA LATERAL: FILTROS TEMPORAIS E GERAIS ---
with st.sidebar:
    st.header("🔍 Filtros de Período")
    data_inicio = st.date_input("Data Inicial", value=date(2026, 10, 1))
    data_fim = st.date_input("Data Final", value=date(2026, 10, 7))
    st.divider()
    st.info("Os filtros abaixo afetam as análises e cruzamentos de auditoria.")

# --- SEPARADORES PRINCIPAIS ---
tab_lancamentos, tab_auditoria, tab_sistemas = st.tabs([
    "📝 Lançamentos Operacionais",
    "⚖️ Auditoria: Transferido vs. Consumido",
    "📂 Conciliação MTech x Sankhya"
])

# ----------------------------------------------------------------------
# TAB 1: LANÇAMENTOS (Entradas da fábrica, Transferências e Consumos)
# ----------------------------------------------------------------------
with tab_lancamentos:
    tipo_operacao = st.radio(
        "Selecione o tipo de registo a realizar:",
        ["1. Entrada de Fábrica (Armazém)", "2. Transferência para Núcleo", "3. Consumo de Campo (Ficha)"],
        horizontal=True
    )
    
    if tipo_operacao == "1. Entrada de Fábrica (Armazém)":
        with st.form("form_entrada"):
            st.subheader("Recebimento de Ração nos Silos Centrais")
            c1, c2, c3, c4 = st.columns(4)
            d_ent = c1.date_input("Data Recebimento", value=date.today())
            racao_ent = c2.selectbox("Tipo de Ração", ["PRE-INICIAL", "INICIAL", "CRESCIMENTO", "GALO", "PRE-POSTURA", "POSTURA 1", "POSTURA 2", "POSTURA 3", "POSTURA 4", "DESCARTE"])
            qtd_ent = c3.number_input("Quantidade (Kg)", min_value=0.0, step=1000.0)
            doc_ent = c4.text_input("Nº NF / Carga")
            
            if st.form_submit_button("Gravar Entrada"):
                nova_linha = pd.DataFrame([{"Data": d_ent, "Tipo Racao": racao_ent, "Quantidade_Kg": qtd_ent, "Documento": doc_ent}])
                st.session_state.entradas = pd.concat([st.session_state.entradas, nova_linha], ignore_index=True)
                st.success("Entrada registada com sucesso!")

    elif tipo_operacao == "2. Transferência para Núcleo":
        with st.form("form_transf"):
            st.subheader("Transferência Interna (Camião de Distribuição)")
            c1, c2, c3 = st.columns(3)
            d_tr = c1.date_input("Data Transferência", value=date.today())
            origem = c2.selectbox("Origem", ["ARMAZEM", "NP5", "NP6", "NP7", "NP8", "NR1", "NR2"])
            destino = c3.selectbox("Destino (Núcleo)", ["NP5", "NP6", "NP7", "NP8", "NR1", "NR2"])
            
            c4, c5, c6 = st.columns(3)
            lote_tr = c4.text_input("Lote (ex: 57, 54)")
            racao_tr = c5.selectbox("Tipo de Ração", ["PRE-INICIAL", "INICIAL", "CRESCIMENTO", "GALO", "PRE-POSTURA", "POSTURA 1", "POSTURA 2", "POSTURA 3", "POSTURA 4", "DESCARTE"])
            qtd_tr = c6.number_input("Quantidade Transferida (Kg)", min_value=0.0, step=500.0)
            
            if st.form_submit_button("Gravar Transferência"):
                nova_tr = pd.DataFrame([{"Data": d_tr, "Origem": origem, "Destino": destino, "Lote": lote_tr, "Tipo Racao": racao_tr, "Quantidade_Kg": qtd_tr}])
                st.session_state.transferencias = pd.concat([st.session_state.transferencias, nova_tr], ignore_index=True)
                st.success("Transferência registada com sucesso!")

    elif tipo_operacao == "3. Consumo de Campo (Ficha)":
        with st.form("form_consumo"):
            st.subheader("Consumo Diário das Aves")
            c1, c2, c3 = st.columns(3)
            d_cs = c1.date_input("Data Consumo", value=date.today())
            nucleo_cs = c2.selectbox("Núcleo", ["NP5", "NP6", "NP7", "NP8", "NR1", "NR2"])
            lote_cs = c3.text_input("Lote Composto")
            
            c4, c5, c6 = st.columns(3)
            racao_cs = c4.selectbox("Tipo Ração Consumida", ["PRE-INICIAL", "INICIAL", "CRESCIMENTO", "GALO", "PRE-POSTURA", "POSTURA 1", "POSTURA 2", "POSTURA 3", "POSTURA 4", "DESCARTE"])
            sexo_cs = c5.selectbox("Sexo", ["Fêmea", "Macho", "Misto"])
            qtd_cs = c6.number_input("Quantidade Consumida (Kg)", min_value=0.0, step=100.0)
            
            if st.form_submit_button("Gravar Consumo"):
                novo_cs = pd.DataFrame([{"Data": d_cs, "Nucleo": nucleo_cs, "Lote": lote_cs, "Tipo Racao": racao_cs, "Sexo": sexo_cs, "Quantidade_Kg": qtd_cs}])
                st.session_state.consumos = pd.concat([st.session_state.consumos, novo_cs], ignore_index=True)
                st.success("Consumo registado com sucesso!")

# ----------------------------------------------------------------------
# TAB 2: AUDITORIA (Cruzamento Transferido vs. Consumido)
# ----------------------------------------------------------------------
with tab_auditoria:
    st.subheader("📊 Cruzamento de Auditoria: Ração Transferida vs. Consumida")
    
    if st.session_state.transferencias.empty and st.session_state.consumos.empty:
        st.info("Ainda não existem registos de transferências ou consumos para calcular a auditoria.")
    else:
        # Agrupamento das transferências por Destino e Tipo de Ração
        df_tr_grp = st.session_state.transferencias.groupby(["Destino", "Tipo Racao"])["Quantidade_Kg"].sum().reset_index()
        df_tr_grp.rename(columns={"Destino": "Núcleo", "Quantidade_Kg": "Transferido_Kg"}, inplace=True)
        
        # Agrupamento dos consumos por Núcleo e Tipo de Ração
        df_cs_grp = st.session_state.consumos.groupby(["Nucleo", "Tipo Racao"])["Quantidade_Kg"].sum().reset_index()
        df_cs_grp.rename(columns={"Nucleo": "Núcleo", "Quantidade_Kg": "Consumido_Kg"}, inplace=True)
        
        # Cruzamento (Merge)
        df_cruzamento = pd.merge(df_tr_grp, df_cs_grp, on=["Núcleo", "Tipo Racao"], how="outer").fillna(0)
        df_cruzamento["Diferença (Saldo Silo)"] = df_cruzamento["Transferido_Kg"] - df_cruzamento["Consumido_Kg"]
        
        # Coluna de Alerta
        def classificar_status(row):
            if row["Diferença (Saldo Silo)"] < 0:
                return "❌ Alerta: Consumo sem Transferência"
            elif row["Diferença (Saldo Silo)"] == 0 and row["Transferido_Kg"] > 0:
                return "✔️ Silo Zerado"
            elif row["Transferido_Kg"] > 0 and row["Consumido_Kg"] == 0:
                return "⚠️ Ração Parada no Silo"
            return "Em Consumo Normal"

        df_cruzamento["Status"] = df_cruzamento.apply(classificar_status, axis=1)
        
        st.dataframe(df_cruzamento, use_container_width=True)

# ----------------------------------------------------------------------
# TAB 3: CONCILIAÇÃO SISTEMAS (MTech vs. Sankhya)
# ----------------------------------------------------------------------
with tab_sistemas:
    st.subheader("📂 Importação e Bate de Ficheiros de Fecho")
    col_up1, col_up2 = st.columns(2)
    with col_up1:
        f_mtech = st.file_uploader("Base MTech (.xlsx)", type=["xlsx", "csv"], key="mtech_audit")
    with col_up2:
        f_sankhya = st.file_uploader("Base Sankhya (.xlsx)", type=["xlsx", "csv"], key="sankhya_audit")


















        