import pandas as pd
import streamlit as st
import plotly.express as px
from datetime import datetime

# Carregar dados
file_path = r"./data/SISTEMA DE GESTÃO DE AVARIAS PREVENÇÃO - FRAGA MAIA.xlsm"
xls = pd.ExcelFile(file_path)

# Folhas disponíveis
folhas = ["Avarias Padaria", "Avarias Salgados", "Avarias Rotisseria"]

# Ler dados da folha selecionada
def carregar_dados(nome_folha):
    df = pd.read_excel(xls, sheet_name=nome_folha, skiprows=1)
    
    # Limpar e pré-processar os dados
    df['QTD'] = pd.to_numeric(df['QTD'], errors='coerce')

    # Função para limpar colunas de moeda e convertê-las para numérico (float64)
    def limpar_coluna_moeda(coluna):
        coluna = coluna.replace({r'R\$ ': '', r',': '.'}, regex=True)
        return pd.to_numeric(coluna, errors='coerce')

    df['VLR. UNIT. VENDA'] = limpar_coluna_moeda(df['VLR. UNIT. VENDA'])
    df['VLR. UNIT. CUSTO'] = limpar_coluna_moeda(df['VLR. UNIT. CUSTO'])
    df['VLR. TOT. VENDA'] = limpar_coluna_moeda(df['VLR. TOT. VENDA'])
    df['VLR. TOT. CUSTO'] = limpar_coluna_moeda(df['VLR. TOT. CUSTO'])
    
    # Remover linhas com quantidades inválidas ou zero
    df = df[(df['QTD'] > 0)].dropna(subset=['QTD'])
    
    return df

# Processar datas e semanas
def processar_datas(df):
    df['DATA'] = pd.to_datetime(df['DATA'], format='%d/%m/%Y', errors='coerce')
    df['mês'] = df['DATA'].dt.month
    df['dia'] = df['DATA'].dt.day
    return df

# Filtrar por período (Mês, Semana)
def filtrar_por_periodo(df, tipo_periodo, valor_periodo, meses):
    if tipo_periodo == 'Mês':
        df = df[df['mês'] == valor_periodo]
    elif tipo_periodo == 'Semana':
        inicio, fim = valor_periodo.split('-')
        inicio = int(inicio.replace('Dia ', ''))
        fim = int(fim)
        df = df[(df['dia'] >= inicio) & (df['dia'] <= fim) & (df['mês'] == meses.index(st.session_state.mes_selecionado) + 1)]
    return df

# Top 10 produtos por QTD
def top_10_por_qtd(df):
    top_10_qtd = df.groupby('DESCRIÇÃO').agg({'QTD': 'sum'}).sort_values(by='QTD', ascending=False).head(10)
    return top_10_qtd

# Top 10 produtos por valor total de venda
def top_10_por_valor_venda(df):
    df['VLR. TOT. VENDA'] = df['QTD'] * df['VLR. UNIT. VENDA']
    top_10_vendas = df.groupby('DESCRIÇÃO').agg({'VLR. TOT. VENDA': 'sum'}).sort_values(by='VLR. TOT. VENDA', ascending=False).head(10)
    return top_10_vendas

# Top 10 produtos por valor total de custo
def top_10_por_valor_custo(df):
    df['VLR. TOT. CUSTO'] = df['QTD'] * df['VLR. UNIT. CUSTO']
    top_10_custo = df.groupby('DESCRIÇÃO').agg({'VLR. TOT. CUSTO': 'sum'}).sort_values(by='VLR. TOT. CUSTO', ascending=False).head(10)
    return top_10_custo

# Nova função para resumo de avarias
def resumo_avarias(df):
    # Agrupar por 'DESCRIÇÃO', somando QTD e valores totais, e pegando o primeiro 'Código Interno'
    resumo = df.groupby('DESCRIÇÃO').agg({
        'QTD': 'sum',
        'VLR. TOT. VENDA': 'sum',
        'VLR. TOT. CUSTO': 'sum',
        'CÓD. INT.': 'first'  # Assume que 'Código Interno' é o nome da coluna
    }).reset_index()
    return resumo

# Interface do Streamlit
def app():
    st.title("Dashboard de Avarias")
    
    # Lista de meses em português
    meses = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 
             'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
    
    # Filtros na barra lateral
    setor = st.sidebar.selectbox('Escolha o setor', folhas)
    tipo_periodo = st.sidebar.selectbox('Escolha o período', ['Mês', 'Semana'])
    
    if tipo_periodo == 'Mês':
        mes_selecionado = st.sidebar.selectbox('Escolha o mês', meses)
        st.session_state.mes_selecionado = mes_selecionado
        valor_periodo = meses.index(mes_selecionado) + 1
    else:
        mes_selecionado = st.sidebar.selectbox('Escolha o mês para as semanas', meses)
        st.session_state.mes_selecionado = mes_selecionado
        semanas = ['Dia 1-7', 'Dia 8-14', 'Dia 15-21', 'Dia 22-31']
        valor_periodo = st.sidebar.selectbox('Escolha a semana', semanas)

    # Carregar e processar os dados
    df = carregar_dados(setor)
    df = processar_datas(df)
    df_filtrado = filtrar_por_periodo(df, tipo_periodo, valor_periodo, meses)
        
    # Exibir total de vendas e custos
    total_vendas = df_filtrado['VLR. TOT. VENDA'].sum()
    total_custo = df_filtrado['VLR. TOT. CUSTO'].sum()
    
    st.markdown("### Total de Vendas e Custos")
    st.metric("Total de Vendas", f"R$ {total_vendas:,.2f}")
    st.metric("Total de Custo", f"R$ {total_custo:,.2f}")
    
    # Top 10 produtos por QTD (quantidade perdida)
    top_10_qtd = top_10_por_qtd(df_filtrado)
    
    if not top_10_qtd.empty:
        st.markdown("### Top 10 Produtos Mais Perdidos (por Quantidade)")
        fig_qtd = px.bar(
            top_10_qtd.reset_index(),
            x='DESCRIÇÃO',
            y='QTD',
            title="Top 10 Produtos por Quantidade Perdida",
            category_orders={'DESCRIÇÃO': top_10_qtd.index.tolist()}
        )
        st.plotly_chart(fig_qtd)
    else:
        st.markdown("### Nenhum dado disponível para Quantidade")
    
    # Top 10 produtos por valor total de venda
    top_10_vendas = top_10_por_valor_venda(df_filtrado)
    
    if not top_10_vendas.empty:
        st.markdown("### Top 10 Produtos Mais Perdidos (por Valor Total de Venda)")
        fig_vendas = px.bar(
            top_10_vendas.reset_index(),
            x='DESCRIÇÃO',
            y='VLR. TOT. VENDA',
            title="Top 10 Produtos por Valor Total de Venda Perdido",
            category_orders={'DESCRIÇÃO': top_10_vendas.index.tolist()}
        )
        st.plotly_chart(fig_vendas)
    else:
        st.markdown("### Nenhum dado disponível para Vendas Totais")
    
    # Top 10 produtos por valor total de custo
    top_10_custo = top_10_por_valor_custo(df_filtrado)
    
    if not top_10_custo.empty:
        st.markdown("### Top 10 Produtos Mais Perdidos (por Valor Total de Custo)")
        fig_custo = px.bar(
            top_10_custo.reset_index(),
            x='DESCRIÇÃO',
            y='VLR. TOT. CUSTO',
            title="Top 10 Produtos por Valor Total de Custo Perdido",
            category_orders={'DESCRIÇÃO': top_10_custo.index.tolist()}
        )
        st.plotly_chart(fig_custo)
    else:
        st.markdown("### Nenhum dado disponível para Custo Total")
    
    # Exibir tabela de avarias detalhada
    st.markdown("### Tabela de Avarias Detalhada")
    df_filtrado_exibicao = df_filtrado[['DATA', 'DESCRIÇÃO', 'QTD', 'VLR. UNIT. VENDA', 'VLR. UNIT. CUSTO', 'VLR. TOT. VENDA', 'VLR. TOT. CUSTO']].copy()

    # Criar novas colunas para exibição formatada
    df_filtrado_exibicao['VLR. UNIT. VENDA (R$)'] = df_filtrado_exibicao['VLR. UNIT. VENDA'].apply(lambda x: f"R$ {x:,.2f}" if pd.notna(x) else "R$ 0,00")
    df_filtrado_exibicao['VLR. UNIT. CUSTO (R$)'] = df_filtrado_exibicao['VLR. UNIT. CUSTO'].apply(lambda x: f"R$ {x:,.2f}" if pd.notna(x) else "R$ 0,00")
    df_filtrado_exibicao['VLR. TOT. VENDA (R$)'] = df_filtrado_exibicao['VLR. TOT. VENDA'].apply(lambda x: f"R$ {x:,.2f}" if pd.notna(x) else "R$ 0,00")
    df_filtrado_exibicao['VLR. TOT. CUSTO (R$)'] = df_filtrado_exibicao['VLR. TOT. CUSTO'].apply(lambda x: f"R$ {x:,.2f}" if pd.notna(x) else "R$ 0,00")

    # Exibir tabela detalhada
    st.dataframe(df_filtrado_exibicao[['DATA', 'DESCRIÇÃO', 'QTD', 'VLR. UNIT. VENDA (R$)', 'VLR. UNIT. CUSTO (R$)', 'VLR. TOT. VENDA (R$)', 'VLR. TOT. CUSTO (R$)']])

    # Exibir tabela de resumo
    st.markdown("### Tabela de Avarias - Resumo")
    df_resumo = resumo_avarias(df_filtrado)
    
    # Formatando valores monetários para exibição
    df_resumo['VLR. TOT. VENDA (R$)'] = df_resumo['VLR. TOT. VENDA'].apply(lambda x: f"R$ {x:,.2f}" if pd.notna(x) else "R$ 0,00")
    df_resumo['VLR. TOT. CUSTO (R$)'] = df_resumo['VLR. TOT. CUSTO'].apply(lambda x: f"R$ {x:,.2f}" if pd.notna(x) else "R$ 0,00")

    # Exibir a tabela de resumo com 'Código Interno' como terceira coluna
    st.dataframe(df_resumo[['DESCRIÇÃO', 'QTD', 'CÓD. INT.', 'VLR. TOT. VENDA (R$)', 'VLR. TOT. CUSTO (R$)']])

if __name__ == "__main__":
    app()