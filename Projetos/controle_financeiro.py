import streamlit as st
import pandas as pd
from datetime import date, datetime

# 1. Configuração Global da Página
st.set_page_config(
    page_title="Controle Financeiro Pessoal",
    page_icon="💰",
    layout="wide"
)

# 2. Funções Utilitárias
def formata_brl(valor: float) -> str:
    """Formata valores numéricos para o padrão monetário brasileiro (R$)."""
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# 3. Inicialização do Estado (Session State)
if "transacoes" not in st.session_state:
    st.session_state.transacoes = pd.DataFrame([
        {"Data": date(2026, 10, 1), "Descrição": "Salário", "Categoria": "Salário", "Tipo": "Receita", "Valor": 5000.00},
        {"Data": date(2026, 10, 2), "Descrição": "Mercado", "Categoria": "Alimentação", "Tipo": "Despesa", "Valor": 850.00},
        {"Data": date(2026, 10, 2), "Descrição": "Combustível", "Categoria": "Transporte", "Tipo": "Despesa", "Valor": 300.00},
    ])

# 4. Título do Sistema
st.title("💼 Controle Financeiro Pessoal")
st.markdown("Gerencie suas receitas, despesas e acompanhe seu saldo em tempo real.")
st.divider()

# 5. Cálculos dos Indicadores (KPIs)
df = st.session_state.transacoes

total_receitas = df[df["Tipo"] == "Receita"]["Valor"].sum() if not df.empty else 0.0
total_despesas = df[df["Tipo"] == "Despesa"]["Valor"].sum() if not df.empty else 0.0
saldo_atual = total_receitas - total_despesas

# 6. Painel de Indicadores
col1, col2, col3 = st.columns(3)

with col1:
    st.metric(label="Total de Receitas", value=formata_brl(total_receitas))

with col2:
    st.metric(
        label="Total de Despesas", 
        value=formata_brl(total_despesas), 
        delta=f"-{formata_brl(total_despesas)}", 
        delta_color="inverse"
    )

with col3:
    st.metric(label="Saldo Atual", value=formata_brl(saldo_atual))

st.divider()

# 7. Formulário e Tabela
col_form, col_tabela = st.columns([1, 2])

with col_form:
    st.subheader("➕ Nova Transação")
    
    with st.form("form_transacao", clear_on_submit=True):
        descricao = st.text_input("Descrição", placeholder="Ex: Supermercado, Salário")
        valor = st.number_input("Valor (R$)", min_value=0.01, format="%.2f", step=10.0)
        tipo = st.selectbox("Tipo de Movimentação", ["Despesa", "Receita"])
        categoria = st.selectbox(
            "Categoria",
            ["Alimentação", "Transporte", "Moradia", "Lazer", "Salário", "Investimentos", "Outros"]
        )
        data_movimentacao = st.date_input("Data da Operação", value=date.today())
        
        botao_salvar = st.form_submit_button("Salvar Registro", use_container_width=True)
        
        if botao_salvar:
            if not descricao.strip():
                st.error("Preencha a descrição antes de salvar!")
            else:
                novo_registro = pd.DataFrame([{
                    "Data": data_movimentacao,  # Mantido como objeto date
                    "Descrição": descricao.strip(),
                    "Categoria": categoria,
                    "Tipo": tipo,
                    "Valor": float(valor)
                }])
                st.session_state.transacoes = pd.concat([st.session_state.transacoes, novo_registro], ignore_index=True)
                st.success(f"Transação '{descricao.strip()}' registrada!")
                st.rerun()

with col_tabela:
    st.subheader("📋 Histórico de Lançamentos")
    
    # data_editor permite selecionar linhas e deletar pressionando Delete no teclado
    tabela_editada = st.data_editor(
        st.session_state.transacoes,
        column_config={
            "Valor": st.column_config.NumberColumn(
                "Valor",
                format="R$ %.2f"
            ),
            "Data": st.column_config.DateColumn(
                "Data",
                format="DD/MM/YYYY"
            )
        },
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True
    )
    
    # Atualiza o estado se linhas forem excluídas ou alteradas diretamente na tabela
    if not tabela_editada.equals(st.session_state.transacoes):
        st.session_state.transacoes = tabela_editada
        st.rerun()