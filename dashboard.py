import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

# Configuração da página
st.set_page_config(
    page_title="Dashboard Prevenção",
    page_icon="👮🏻‍♂️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Carregamento dos dados
file_path = './data/SISTEMA GERAL PREVENÇÃO - FRAGA MAIA3.xlsm'
sheet_name = "Recuperação de Avarias"

# Ignorar a primeira linha ao ler o arquivo
df = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=1, usecols="A:H")

# Definir os nomes das colunas manualmente
column_names = ['Data', 'Código Barras', 'Código Interno', 'Descrição', 'QTD.', 'Vlr. Uni.', 'Total', 'PREV.']
df.columns = column_names

# Remoção de linhas com valores faltantes
df = df.dropna()

# Conversão das colunas para float (se necessário)
df['Vlr. Uni.'] = df['Vlr. Uni.'].astype(str).str.replace('R$ ', '').str.replace('.', '').str.replace(',', '.').astype(float)
df['Total'] = df['Total'].astype(str).str.replace('R$ ', '').str.replace('.', '').str.replace(',', '.').astype(float)

# Configuração do sidebar
with st.sidebar:
    st.title('👮🏻‍♂️ Dashboard Prevenção')
    prevention_filter = st.multiselect("Escolha o Prevenção", df["PREV."].unique())
    date_filter = st.radio("Selecione a data", ["Monthly", "Weekly"])

# Filtragem dos dados
if prevention_filter:
    df_filtered = df[df["PREV."].isin(prevention_filter)]
else:
    df_filtered = df.copy()

# Criação da meta e progresso
goal_value = 4000
total_recovered = df_filtered["Total"].sum()
goal_completion_percentage = (total_recovered / goal_value) * 100

# Exibição dos gráficos
st.title("Dashboard Prevenção 👮🏻‍♂️")

# Gráfico: Top 5 Prevenções que mais recuperaram
st.header("Top 5 Prevenções que mais recuperaram")
top_5_prevencao = df_filtered.groupby("PREV.")["Total"].sum().nlargest(5).reset_index()

fig, ax = plt.subplots(figsize=(10, 6))
ax.bar(top_5_prevencao["PREV."], top_5_prevencao["Total"], color='blue')
ax.set_title("Top 5 Prevenções que mais recuperaram")
ax.set_xlabel("Prevenção")
ax.set_ylabel("Recuperado (R$)")
plt.xticks(rotation=45)
st.pyplot(fig)

# Gráfico: Progresso da meta (R$4k recuperados)
st.header("Progresso da meta (R$4k recuperados)")
fig_goal, ax_goal = plt.subplots(figsize=(10, 4))
ax_goal.barh(["Goal Completion"], [goal_completion_percentage], color='green')
ax_goal.axvline(x=100, color='red', linestyle='--', label='Goal (100%)')
ax_goal.set_title("Objetivo (R$4,000.00)")
ax_goal.set_xlabel("Progresso (%)")
ax_goal.legend()
st.pyplot(fig_goal)

# Top 5 Produtos por valor
st.header("Top 5 Produtos (por valor)")
top_5_value = df_filtered.groupby("Descrição")["Total"].sum().nlargest(5).reset_index()

fig_products_value, ax_products_value = plt.subplots(figsize=(10, 6))
ax_products_value.bar(top_5_value["Descrição"], top_5_value["Total"], color='orange')
ax_products_value.set_title("Top 5 Produtos por Valor")
ax_products_value.set_xlabel("Produto")
ax_products_value.set_ylabel("Valor Total (R$)")
plt.xticks(rotation=45)
st.pyplot(fig_products_value)

# Top 5 Produtos por quantidade
st.header("Top 5 Produtos (por quantidade)")
top_5_quantity = df_filtered.groupby("Descrição")["QTD."].sum().nlargest(5).reset_index()

fig_products_quantity, ax_products_quantity = plt.subplots(figsize=(10, 6))
ax_products_quantity.bar(top_5_quantity["Descrição"], top_5_quantity["QTD."], color='purple')
ax_products_quantity.set_title("Top 5 Produtos por Quantidade")
ax_products_quantity.set_xlabel("Produto")
ax_products_quantity.set_ylabel("Quantidade Total")
plt.xticks(rotation=45)
st.pyplot(fig_products_quantity)

# Exibição dos dados filtrados
st.header("Dados Filtrados")
st.write(df_filtered[["Data", "Descrição", "QTD.", "Total", "PREV."]])
