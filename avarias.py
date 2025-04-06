import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime

# Carregar dados
file_path = r"./data/SISTEMA DE GESTÃO DE AVARIAS PREVENÇÃO - FRAGA MAIA.xlsx"
xls = pd.ExcelFile(file_path)

# Folhas disponíveis
folhas = ["Avarias Padaria", "Avarias Salgados", "Avarias Rotisseria"]

# Ler dados da folha selecionada
def carregar_dados(nome_folha):
    df = pd.read_excel(xls, sheet_name=nome_folha, skiprows=1)
    
    df['QTD'] = pd.to_numeric(df['QTD'], errors='coerce')

    def limpar_coluna_moeda(coluna):
        coluna = coluna.replace({r'R\$ ': '', r',': '.'}, regex=True)
        return pd.to_numeric(coluna, errors='coerce')

    df['VLR. UNIT. VENDA'] = limpar_coluna_moeda(df['VLR. UNIT. VENDA'])
    df['VLR. UNIT. CUSTO'] = limpar_coluna_moeda(df['VLR. UNIT. CUSTO'])
    df['VLR. TOT. VENDA'] = limpar_coluna_moeda(df['VLR. TOT. VENDA'])
    df['VLR. TOT. CUSTO'] = limpar_coluna_moeda(df['VLR. TOT. CUSTO'])
    
    df = df[(df['QTD'] > 0)].dropna(subset=['QTD'])
    
    return df

# Processar datas com ano e semana
def processar_datas(df):
    df['DATA'] = pd.to_datetime(df['DATA'], format='%d/%m/%Y', errors='coerce')
    df['mês'] = df['DATA'].dt.month
    df['dia'] = df['DATA'].dt.day
    df['ano'] = df['DATA'].dt.year
    df['semana'] = df['DATA'].dt.isocalendar().week
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

# Resumo de avarias
def resumo_avarias(df):
    resumo = df.groupby('DESCRIÇÃO').agg({
        'QTD': 'sum',
        'VLR. TOT. VENDA': 'sum',
        'VLR. TOT. CUSTO': 'sum',
        'CÓD. INT.': 'first'
    }).reset_index()
    return resumo

def app():
    st.title("Dashboard de Avarias")
    
    meses = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 
             'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
    
    # Filtros na barra lateral
    setor = st.sidebar.selectbox('Escolha o setor', folhas)
    tipo_periodo = st.sidebar.selectbox('Escolha o período', ['Geral', 'Mês', 'Semana'])
    
    # Carregar e processar os dados básicos
    df = carregar_dados(setor)
    df = processar_datas(df)
    
    if tipo_periodo == 'Mês':
        mes_selecionado = st.sidebar.selectbox('Escolha o mês', meses)
        st.session_state.mes_selecionado = mes_selecionado
        valor_periodo = meses.index(mes_selecionado) + 1
        df_filtrado = filtrar_por_periodo(df, tipo_periodo, valor_periodo, meses)
    elif tipo_periodo == 'Semana':
        mes_selecionado = st.sidebar.selectbox('Escolha o mês para as semanas', meses)
        st.session_state.mes_selecionado = mes_selecionado
        semanas = ['Dia 1-7', 'Dia 8-14', 'Dia 15-21', 'Dia 22-31']
        valor_periodo = st.sidebar.selectbox('Escolha a semana', semanas)
        df_filtrado = filtrar_por_periodo(df, tipo_periodo, valor_periodo, meses)
    else:  # Geral
        df_filtrado = df.copy()
        df_all = pd.concat([processar_datas(carregar_dados(folha)).assign(CATEGORIA=folha) for folha in folhas])
        
        # Gráfico de linha - Valor Total de Venda por mês com média móvel
        vendas_por_mes = df.groupby('mês').agg({'VLR. TOT. VENDA': 'sum'}).reset_index()
        vendas_por_mes['mês_nome'] = vendas_por_mes['mês'].apply(lambda x: meses[x-1])
        vendas_por_mes['Média Móvel (3 meses)'] = vendas_por_mes['VLR. TOT. VENDA'].rolling(window=3, min_periods=1).mean()
        
        fig_vendas = go.Figure()
        fig_vendas.add_trace(go.Scatter(
            x=vendas_por_mes['mês_nome'],
            y=vendas_por_mes['VLR. TOT. VENDA'],
            mode='lines+markers',
            name='Valor Real',
            text=[f'R$ {x:,.2f}' for x in vendas_por_mes['VLR. TOT. VENDA']],
            hovertemplate='%{text}<br>Variação: %{customdata:.2f}%<extra></extra>',
            customdata=[0] + list(((vendas_por_mes['VLR. TOT. VENDA'][1:].values - 
                                  vendas_por_mes['VLR. TOT. VENDA'][:-1].values) / 
                                  vendas_por_mes['VLR. TOT. VENDA'][:-1].values * 100))
        ))
        fig_vendas.add_trace(go.Scatter(
            x=vendas_por_mes['mês_nome'],
            y=vendas_por_mes['Média Móvel (3 meses)'],
            mode='lines',
            name='Média Móvel (3 meses)',
            line=dict(dash='dash'),
            text=[f'R$ {x:,.2f}' for x in vendas_por_mes['Média Móvel (3 meses)']],
            hovertemplate='%{text}<extra></extra>'
        ))
        fig_vendas.update_layout(
            title="Valor Total de Venda por Mês (com Média Móvel)",
            xaxis_title="Mês",
            yaxis_title="Valor (R$)",
            legend=dict(x=0, y=1.1, orientation="h")
        )
        st.plotly_chart(fig_vendas)
        
        # Gráfico de linha - Valor Total de Custo por mês com média móvel
        custo_por_mes = df.groupby('mês').agg({'VLR. TOT. CUSTO': 'sum'}).reset_index()
        custo_por_mes['mês_nome'] = custo_por_mes['mês'].apply(lambda x: meses[x-1])
        custo_por_mes['Média Móvel (3 meses)'] = custo_por_mes['VLR. TOT. CUSTO'].rolling(window=3, min_periods=1).mean()
        
        fig_custo = go.Figure()
        fig_custo.add_trace(go.Scatter(
            x=custo_por_mes['mês_nome'],
            y=custo_por_mes['VLR. TOT. CUSTO'],
            mode='lines+markers',
            name='Valor Real',
            text=[f'R$ {x:,.2f}' for x in custo_por_mes['VLR. TOT. CUSTO']],
            hovertemplate='%{text}<br>Variação: %{customdata:.2f}%<extra></extra>',
            customdata=[0] + list(((custo_por_mes['VLR. TOT. CUSTO'][1:].values - 
                                  custo_por_mes['VLR. TOT. CUSTO'][:-1].values) / 
                                  custo_por_mes['VLR. TOT. CUSTO'][:-1].values * 100))
        ))
        fig_custo.add_trace(go.Scatter(
            x=custo_por_mes['mês_nome'],
            y=custo_por_mes['Média Móvel (3 meses)'],
            mode='lines',
            name='Média Móvel (3 meses)',
            line=dict(dash='dash'),
            text=[f'R$ {x:,.2f}' for x in custo_por_mes['Média Móvel (3 meses)']],
            hovertemplate='%{text}<extra></extra>'
        ))
        fig_custo.update_layout(
            title="Valor Total de Custo por Mês (com Média Móvel)",
            xaxis_title="Mês",
            yaxis_title="Valor (R$)",
            legend=dict(x=0, y=1.1, orientation="h")
        )
        st.plotly_chart(fig_custo)
        
        # Gráfico de colunas - Itens mais perdidos por quantidade
        top_qtd = top_10_por_qtd(df_filtrado)
        fig_qtd = px.bar(
            top_qtd.reset_index(),
            x='DESCRIÇÃO',
            y='QTD',
            title="Top 10 Itens Mais Perdidos por Quantidade (Geral)"
        )
        st.plotly_chart(fig_qtd)
        
        # Gráfico de pizza - Perda por categoria (Valor Total de Venda)
        venda_por_categoria = df_all.groupby('CATEGORIA')['VLR. TOT. VENDA'].sum().reset_index()
        fig_pizza_venda = px.pie(
            venda_por_categoria,
            values='VLR. TOT. VENDA',
            names='CATEGORIA',
            title="Perda por Categoria (Valor Total de Venda)"
        )
        st.plotly_chart(fig_pizza_venda)
        
        # Gráfico de pizza - Perda por categoria (Valor Total de Custo)
        custo_por_categoria = df_all.groupby('CATEGORIA')['VLR. TOT. CUSTO'].sum().reset_index()
        fig_pizza_custo = px.pie(
            custo_por_categoria,
            values='VLR. TOT. CUSTO',
            names='CATEGORIA',
            title="Perda por Categoria (Valor Total de Custo)"
        )
        st.plotly_chart(fig_pizza_custo)
        
        # Métricas gerais com impacto na margem de lucro
        total_vendas = df['VLR. TOT. VENDA'].sum()
        total_custo = df['VLR. TOT. CUSTO'].sum()
        lucro_perdido = total_vendas - total_custo
        st.markdown("### Totais Gerais")
        st.metric("Total de Vendas Perdidas", f"R$ {total_vendas:,.2f}")
        st.metric("Total de Custos Perdidos", f"R$ {total_custo:,.2f}")
        st.metric("Lucro Perdido (Margem de Lucro Impactada)", f"R$ {lucro_perdido:,.2f}")
        
        # Comparativo Ano a Ano - Vendas
        vendas_por_ano_mes = df_all.groupby(['ano', 'mês'])['VLR. TOT. VENDA'].sum().reset_index()
        vendas_por_ano_mes['mês_nome'] = vendas_por_ano_mes['mês'].apply(lambda x: meses[x-1])
        fig_vendas_ano = go.Figure()
        for ano in vendas_por_ano_mes['ano'].unique():
            df_ano = vendas_por_ano_mes[vendas_por_ano_mes['ano'] == ano]
            fig_vendas_ano.add_trace(go.Scatter(
                x=df_ano['mês_nome'],
                y=df_ano['VLR. TOT. VENDA'],
                mode='lines+markers',
                name=str(ano),
                text=[f'R$ {x:,.2f}' for x in df_ano['VLR. TOT. VENDA']],
                hovertemplate='%{text}<extra></extra>'
            ))
        fig_vendas_ano.update_layout(
            title="Comparativo de Vendas Perdidas por Ano",
            xaxis_title="Mês",
            yaxis_title="Valor (R$)",
            legend=dict(x=0, y=1.1, orientation="h")
        )
        st.plotly_chart(fig_vendas_ano)
        
        # Comparativo Ano a Ano - Custos
        custo_por_ano_mes = df_all.groupby(['ano', 'mês'])['VLR. TOT. CUSTO'].sum().reset_index()
        custo_por_ano_mes['mês_nome'] = custo_por_ano_mes['mês'].apply(lambda x: meses[x-1])
        fig_custo_ano = go.Figure()
        for ano in custo_por_ano_mes['ano'].unique():
            df_ano = custo_por_ano_mes[custo_por_ano_mes['ano'] == ano]
            fig_custo_ano.add_trace(go.Scatter(
                x=df_ano['mês_nome'],
                y=df_ano['VLR. TOT. CUSTO'],
                mode='lines+markers',
                name=str(ano),
                text=[f'R$ {x:,.2f}' for x in df_ano['VLR. TOT. CUSTO']],
                hovertemplate='%{text}<extra></extra>'
            ))
        fig_custo_ano.update_layout(
            title="Comparativo de Custos Perdidos por Ano",
            xaxis_title="Mês",
            yaxis_title="Valor (R$)",
            legend=dict(x=0, y=1.1, orientation="h")
        )
        st.plotly_chart(fig_custo_ano)
        
        # Padrões Sazonais - Heatmap por Mês e Ano (Vendas)
        vendas_por_mes_ano = df_all.groupby(['ano', 'mês'])['VLR. TOT. VENDA'].sum().reset_index()
        vendas_por_mes_ano['mês_nome'] = vendas_por_mes_ano['mês'].apply(lambda x: meses[x-1])
        fig_heatmap_vendas_mes = px.density_heatmap(
            vendas_por_mes_ano,
            x='mês_nome',
            y='ano',
            z='VLR. TOT. VENDA',
            title="Padrões Sazonais - Vendas por Mês e Ano",
            color_continuous_scale="Viridis"
        )
        st.plotly_chart(fig_heatmap_vendas_mes)
        
        # Padrões Sazonais - Heatmap por Semana e Ano (Vendas)
        vendas_por_semana_ano = df_all.groupby(['ano', 'semana'])['VLR. TOT. VENDA'].sum().reset_index()
        fig_heatmap_vendas_semana = px.density_heatmap(
            vendas_por_semana_ano,
            x='semana',
            y='ano',
            z='VLR. TOT. VENDA',
            title="Padrões Sazonais - Vendas por Semana e Ano",
            color_continuous_scale="Viridis"
        )
        st.plotly_chart(fig_heatmap_vendas_semana)
        
        # Padrões Sazonais - Heatmap por Mês e Ano (Custos)
        custo_por_mes_ano = df_all.groupby(['ano', 'mês'])['VLR. TOT. CUSTO'].sum().reset_index()
        custo_por_mes_ano['mês_nome'] = custo_por_mes_ano['mês'].apply(lambda x: meses[x-1])
        fig_heatmap_custo_mes = px.density_heatmap(
            custo_por_mes_ano,
            x='mês_nome',
            y='ano',
            z='VLR. TOT. CUSTO',
            title="Padrões Sazonais - Custos por Mês e Ano",
            color_continuous_scale="Viridis"
        )
        st.plotly_chart(fig_heatmap_custo_mes)
        
        # Padrões Sazonais - Heatmap por Semana e Ano (Custos)
        custo_por_semana_ano = df_all.groupby(['ano', 'semana'])['VLR. TOT. CUSTO'].sum().reset_index()
        fig_heatmap_custo_semana = px.density_heatmap(
            custo_por_semana_ano,
            x='semana',
            y='ano',
            z='VLR. TOT. CUSTO',
            title="Padrões Sazonais - Custos por Semana e Ano",
            color_continuous_scale="Viridis"
        )
        st.plotly_chart(fig_heatmap_custo_semana)

    # Restante do código para Mês e Semana
    if tipo_periodo != 'Geral':
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

        df_filtrado_exibicao['VLR. UNIT. VENDA (R$)'] = df_filtrado_exibicao['VLR. UNIT. VENDA'].apply(lambda x: f"R$ {x:,.2f}" if pd.notna(x) else "R$ 0,00")
        df_filtrado_exibicao['VLR. UNIT. CUSTO (R$)'] = df_filtrado_exibicao['VLR. UNIT. CUSTO'].apply(lambda x: f"R$ {x:,.2f}" if pd.notna(x) else "R$ 0,00")
        df_filtrado_exibicao['VLR. TOT. VENDA (R$)'] = df_filtrado_exibicao['VLR. TOT. VENDA'].apply(lambda x: f"R$ {x:,.2f}" if pd.notna(x) else "R$ 0,00")
        df_filtrado_exibicao['VLR. TOT. CUSTO (R$)'] = df_filtrado_exibicao['VLR. TOT. CUSTO'].apply(lambda x: f"R$ {x:,.2f}" if pd.notna(x) else "R$ 0,00")

        st.dataframe(df_filtrado_exibicao[['DATA', 'DESCRIÇÃO', 'QTD', 'VLR. UNIT. VENDA (R$)', 'VLR. UNIT. CUSTO (R$)', 'VLR. TOT. VENDA (R$)', 'VLR. TOT. CUSTO (R$)']])

        # Exibir tabela de resumo
        st.markdown("### Tabela de Avarias - Resumo")
        df_resumo = resumo_avarias(df_filtrado)
        
        df_resumo['VLR. TOT. VENDA (R$)'] = df_resumo['VLR. TOT. VENDA'].apply(lambda x: f"R$ {x:,.2f}" if pd.notna(x) else "R$ 0,00")
        df_resumo['VLR. TOT. CUSTO (R$)'] = df_resumo['VLR. TOT. CUSTO'].apply(lambda x: f"R$ {x:,.2f}" if pd.notna(x) else "R$ 0,00")

        st.dataframe(df_resumo[['DESCRIÇÃO', 'QTD', 'CÓD. INT.', 'VLR. TOT. VENDA (R$)', 'VLR. TOT. CUSTO (R$)']])

if __name__ == "__main__":
    app()