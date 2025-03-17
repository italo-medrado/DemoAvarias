import pandas as pd
import streamlit as st
import plotly.express as px
from datetime import datetime

# Configuração da página
st.set_page_config(
    page_title="Dashboard Prevenção",
    page_icon="👮🏻‍♂️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Carregar dados
file_path = "./data/SISTEMA GERAL PREVENÇÃO - FRAGA MAIA3.xlsm"
xls = pd.ExcelFile(file_path)

# Folhas disponíveis
folhas = ["Recuperação de Avarias", "Furtos Recuperados", "Quebra Mês", "Quebra Deg"]

# Ler dados da folha selecionada
def carregar_dados(nome_folha):
    df = pd.read_excel(xls, sheet_name=nome_folha, skiprows=1, usecols="A:H")
    
    # Definir nomes das colunas
    column_names = ['DATA', 'CÓDIGO BARRAS', 'CÓDIGO INTERNO', 'DESCRIÇÃO', 'QTD', 'VLR. UNI.', 'TOTAL', 'PREV.']
    df.columns = column_names
    
    # Limpar e pré-processar os dados
    df['QTD'] = pd.to_numeric(df['QTD'], errors='coerce')
    
    # Função para limpar colunas de moeda
    def limpar_coluna_moeda(coluna):
        # Converter para string e remover espaços e "R$"
        coluna = coluna.astype(str).str.replace('R\$', '', regex=True).str.strip()
        
        # Função auxiliar para processar cada valor
        def processar_valor(valor):
            # Remover qualquer formatação estranha, mas preservar o último ponto ou vírgula como decimal
            valor = valor.replace(' ', '')
            
            # Se houver vírgula, tratamos como decimal
            if ',' in valor:
                # Dividir em parte inteira e decimal
                partes = valor.rsplit(',', 1)
                inteiro = partes[0].replace('.', '')  # Remover pontos de milhar
                decimal = partes[1] if len(partes) > 1 else '00'
                valor = f"{inteiro}.{decimal}"
            else:
                # Se não houver vírgula, tratamos o último ponto como decimal
                partes = valor.rsplit('.', 1)
                if len(partes) > 1:
                    inteiro = partes[0].replace('.', '')  # Remover pontos de milhar
                    decimal = partes[1]
                    valor = f"{inteiro}.{decimal}"
            
            # Converter para numérico
            return pd.to_numeric(valor, errors='coerce')
        
        # Aplicar a função a cada valor na coluna
        return coluna.apply(processar_valor)
    
    df['VLR. UNI.'] = limpar_coluna_moeda(df['VLR. UNI.'])
    df['TOTAL'] = limpar_coluna_moeda(df['TOTAL'])
        
    # Remover linhas com quantidades inválidas ou zero
    df = df[(df['QTD'] > 0)].dropna(subset=['QTD', 'TOTAL'])
    
    return df

# Processar datas e períodos
def processar_dates(df):
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

# Top 5 prevenções por total recuperado
def top_5_prevencao(df):
    top_5 = df.groupby('PREV.')['TOTAL'].sum().nlargest(5).reset_index()
    return top_5

# Top 5 produtos por valor total
def top_5_por_valor(df):
    top_5 = df.groupby('DESCRIÇÃO')['TOTAL'].sum().nlargest(5).reset_index()
    return top_5

# Top 5 produtos por quantidade
def top_5_por_quantidade(df):
    top_5 = df.groupby('DESCRIÇÃO')['QTD'].sum().nlargest(5).reset_index()
    return top_5

# Resumo de prevenções
def resumo_prevencoes(df):
    resumo = df.groupby('DESCRIÇÃO').agg({
        'QTD': 'sum',
        'TOTAL': 'sum',
        'CÓDIGO INTERNO': 'first'
    }).reset_index()
    return resumo

# Função para formatar valores monetários no formato brasileiro
def formatar_moeda(valor):
    if pd.isna(valor):
        return "R$ 0,00"
    # Converter o valor para string com 2 casas decimais
    valor_str = f"{valor:.2f}"
    # Substituir o ponto decimal por vírgula e adicionar ponto como separador de milhares
    partes = valor_str.split('.')
    inteiro = partes[0]
    decimal = partes[1]
    # Adicionar pontos como separador de milhares
    inteiro_com_pontos = ''
    for i, digito in enumerate(reversed(inteiro)):
        if i > 0 and i % 3 == 0:
            inteiro_com_pontos = '.' + inteiro_com_pontos
        inteiro_com_pontos = digito + inteiro_com_pontos
    return f"R$ {inteiro_com_pontos},{decimal}"

# Interface do Streamlit
def app():
    st.title("Dashboard Prevenção 👮🏻‍♂️")
    
    # Lista de meses em português
    meses = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 
             'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
    
    # Filtros na barra lateral
    with st.sidebar:
        st.title('👮🏻‍♂️ Dashboard Prevenção')
        setor = st.selectbox('Escolha o setor', folhas)
        tipo_periodo = st.selectbox('Escolha o período', ['Mês', 'Semana'])
        
        if tipo_periodo == 'Mês':
            mes_selecionado = st.selectbox('Escolha o mês', meses)
            st.session_state.mes_selecionado = mes_selecionado
            valor_periodo = meses.index(mes_selecionado) + 1
        else:
            mes_selecionado = st.selectbox('Escolha o mês para as semanas', meses)
            st.session_state.mes_selecionado = mes_selecionado
            semanas = ['Dia 1-7', 'Dia 8-14', 'Dia 15-21', 'Dia 22-31']
            valor_periodo = st.selectbox('Escolha a semana', semanas)
        
        prevention_filter = st.multiselect("Escolha o Prevenção", carregar_dados(setor)['PREV.'].unique())

    # Carregar e processar os dados
    df = carregar_dados(setor)
    df = processar_dates(df)
    df_filtrado = filtrar_por_periodo(df, tipo_periodo, valor_periodo, meses)
    
    # Aplicar filtro de PREV.
    if prevention_filter:
        df_filtrado = df_filtrado[df_filtrado['PREV.'].isin(prevention_filter)]
    
    # Exibir total recuperado
    total_recuperado = df_filtrado['TOTAL'].sum()
    st.markdown("### Total Recuperado")
    st.metric("Total Recuperado", formatar_moeda(total_recuperado))
        
    # Top 5 Prevenções que mais recuperaram
    top_5_prev = top_5_prevencao(df_filtrado)
    if not top_5_prev.empty:
        st.markdown("### Top 5 Prevenções que Mais Recuperaram")
        fig_prev = px.bar(
            top_5_prev, x='PREV.', y='TOTAL',
            title="Top 5 Prevenções que Mais Recuperaram",
            labels={'PREV.': 'Prevenção', 'TOTAL': 'Recuperado (R$)'}
        )
        st.plotly_chart(fig_prev)
    else:
        st.markdown("### Nenhum dado disponível para Prevenções")
    
    # Top 5 Produtos por Valor
    top_5_valor = top_5_por_valor(df_filtrado)
    if not top_5_valor.empty:
        st.markdown("### Top 5 Produtos (por Valor)")
        fig_valor = px.bar(
            top_5_valor, x='DESCRIÇÃO', y='TOTAL',
            title="Top 5 Produtos por Valor",
            labels={'DESCRIÇÃO': 'Produto', 'TOTAL': 'Valor Total (R$)'}
        )
        st.plotly_chart(fig_valor)
    else:
        st.markdown("### Nenhum dado disponível para Valor")
    
    # Top 5 Produtos por Quantidade
    top_5_qtd = top_5_por_quantidade(df_filtrado)
    if not top_5_qtd.empty:
        st.markdown("### Top 5 Produtos (por Quantidade)")
        fig_qtd = px.bar(
            top_5_qtd, x='DESCRIÇÃO', y='QTD',
            title="Top 5 Produtos por Quantidade",
            labels={'DESCRIÇÃO': 'Produto', 'QTD': 'Quantidade Total'}
        )
        st.plotly_chart(fig_qtd)
    else:
        st.markdown("### Nenhum dado disponível para Quantidade")
    
    # Tabela de dados filtrados
    st.markdown("### Tabela de Prevenções Detalhada")
    df_exibicao = df_filtrado[['DATA', 'DESCRIÇÃO', 'QTD', 'VLR. UNI.', 'TOTAL', 'PREV.']].copy()
    df_exibicao['VLR. UNI. (R$)'] = df_exibicao['VLR. UNI.'].apply(formatar_moeda)
    df_exibicao['TOTAL (R$)'] = df_exibicao['TOTAL'].apply(formatar_moeda)
    st.dataframe(df_exibicao[['DATA', 'DESCRIÇÃO', 'QTD', 'VLR. UNI. (R$)', 'TOTAL (R$)', 'PREV.']])
    
    # Tabela de resumo
    st.markdown("### Tabela de Prevenções - Resumo")
    df_resumo = resumo_prevencoes(df_filtrado)
    df_resumo['TOTAL (R$)'] = df_resumo['TOTAL'].apply(formatar_moeda)
    st.dataframe(df_resumo[['DESCRIÇÃO', 'QTD', 'CÓDIGO INTERNO', 'TOTAL (R$)']])

if __name__ == "__main__":
    app()