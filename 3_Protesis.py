import pyodbc
import streamlit as st
import pandas as pd
from io import BytesIO
import locale
import plotly.express as px
import dash
import warnings
import subprocess
import os

warnings.filterwarnings('ignore')

st.set_page_config(page_title="Grupo Sanus!!!", page_icon="bar_chart", layout="wide")

st.title(":bar_chart: PROTESIS - ANÁLISE COORDENAÇÃO")
st.markdown('<style>div.block-container{padding-top:1rem;}</style>', unsafe_allow_html=True)

# Texto do botão
button_text = "Power BI"

# URL do seu relatório do Power BI
power_bi_url = "https://app.powerbi.com/view?r=eyJrIjoiMzY3MDY0NmEtMTg4MC00YTg0LWJiMGItOTk2OGRjYmYyOTVjIiwidCI6IjU0NGE1MmMzLTBhNGEtNDVmMi1hMjEzLTFkNTQ1ZjBkMWVjNiJ9"

# Crie um botão com um link
st.markdown(f"[{button_text}]({power_bi_url})")

# Função para reiniciar o aplicativo na página de login
def restart_app():
    current_pid = os.getpid()  # Obtém o ID do processo atual
    os.system(f'taskkill /F /PID {current_pid}')  # Encerra o processo atual
    subprocess.Popen([
        'C:\\Users\\antonio.joseli\\AppData\\Local\\Packages\\PythonSoftwareFoundation.Python.3.11_qbz5n2kfra8p0\\LocalCache\\local-packages\\Python311\\Scripts\\streamlit',
        'run',
        'C:\\Users\\antonio.joseli\\Desktop\\STREAMLIT\\AnaliseBI03\\login\\app.py'
    ])

# Verifique se o botão "Sair" foi pressionado
if st.button("Sair"):
    restart_app()

# Ajustar a tabela inteira na minha página
# st.set_page_config(layout="wide")  # Remove esta linha para corrigir o erro


# Ajustar a tabela inteira na minha página
#st.set_page_config(layout="wide")

# Função para fazer a conexão com o Banco de dados
def connect_to_database():
    dados_conexao = (
        "DRIVER={ODBC Driver 17 for SQL Server};"
        "SERVER=GS-GO-038\\SQLEXPRESS;"
        "DATABASE=BI_Sanus_Logistica;"
        "UID=sa;"
        "PWD=123;"
    )

    conexao = pyodbc.connect(dados_conexao)
    return conexao

# Função para gerar os gráficos com base nos dados filtrados
def generate_plots(data):
    # Agrupe os dados originais por "Fornecedor" e calcule a soma de "Valor_total_item" e a contagem de "Sequencia"
    category_df = data.groupby(by=["Fornecedor"], as_index=False).agg({"Valor_total_item": "sum", "Sequencia": "count"})

    # Renomeie as colunas
    category_df = category_df.rename(columns={"Valor_total_item": "Valor_total_item", "Sequencia": "Nr Nota"})

    return category_df

# Fazendo a conexão com o Banco de dados
conexao = connect_to_database()

# Realizando a consulta SQL
comando = """SELECT * FROM Estab_Protesis_Base_Consignados"""
df = pd.read_sql_query(comando, conexao)
#df["Emissao"] = pd.to_datetime(df["Emissao"].str.split(" ", expand=True)[0], format="%d/%m/%Y")
df['Emissao'] = pd.to_datetime(df['Emissao'], format='%d/%m/%Y')
df = df.sort_values("Emissao")

st.sidebar.header("Escolha seu filtro: ")

# Filtro Ano e Mês
df["Month"] = df["Ano_Emissao"].astype(str) + "-" + df["Nome_Mes_Emissao"]
months = st.sidebar.multiselect("Escolha o Ano e Mês", df["Month"].unique())
if not months:
    df2 = df.copy()
else:
    df2 = df[df["Month"].isin(months)]

# Filtro por Fornecedor
Fornecedor = st.sidebar.multiselect("Escolha o Fornecedor", df["Fornecedor"].unique())
if not Fornecedor:
    df3 = df2.copy()
else:
    df3 = df2[df2["Fornecedor"].isin(Fornecedor)]

# Filtro por Unidade
Unidade = st.sidebar.multiselect("Escolha a Unidade", df3["Unidade"].unique())

# Filtre os dados com base em Mês e Ano, Fornecedor e Unidade
filtered_df = df3
if months:
    filtered_df = filtered_df[filtered_df["Month"].isin(months)]
if Fornecedor:
    filtered_df = filtered_df[filtered_df["Fornecedor"].isin(Fornecedor)]
if Unidade:
    filtered_df = filtered_df[filtered_df["Unidade"].isin(Unidade)]


# Defina a localização para formatar em formato monetário (no caso, pt_BR para o Brasil)
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

# Converter a coluna 'Valor_total_item' para float no DataFrame original
df['Valor_total_item'] = pd.to_numeric(df['Valor_total_item'], errors='coerce')

# Chame a função para gerar os gráficos com base nos dados filtrados
category_df = generate_plots(filtered_df)

# Exibir o data frame com base nos filtros
st.write(filtered_df)

# Exibir os gráficos
col1, col2 = st.columns(2)
with col1:
    st.subheader("Faturamento por Empresa")
    fig = px.bar(category_df, x="Fornecedor", y="Valor_total_item", text=['R${:,.2f}'.format(x) for x in category_df["Valor_total_item"]], template="seaborn")
    fig.update_yaxes(title_text='')
    st.plotly_chart(fig, use_container_width=True)

with col2:
    st.subheader("Faturamento por Empresa em %")
    fig = px.pie(category_df, values="Valor_total_item", names="Fornecedor", hole=0.5)
    st.plotly_chart(fig, use_container_width=True)

locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
# Tabela 1
col1, col2 = st.columns(2)
with col1:
    with st.expander("Visualizar Pagamentos por Fornecedores (Tabela 1) - Filtros"):
        # Selecione apenas as colunas "Fornecedor" e "Valor_total_item"
        filtered_category_df = category_df[["Fornecedor", "Valor_total_item"]]
        # Formate a coluna "Valor_total_item" para moeda
        filtered_category_df["Valor_total_item"] = filtered_category_df["Valor_total_item"].apply(lambda x: locale.currency(x, grouping=True))
        st.dataframe(filtered_category_df, use_container_width=True)
        
        xls_data = BytesIO()
        with pd.ExcelWriter(xls_data, engine="openpyxl", mode="xlsx") as writer:
            filtered_category_df.to_excel(writer, sheet_name="Fornecedor_Tabela1", index=False)
        st.download_button("Baixar Dados (XLS - Tabela 1)", data=xls_data, file_name="Fornecedor_Tabela1.xlsx", key="fornecedor_xls_tabela1",
                        help="Clique aqui para baixar os dados como um arquivo em XLSX")


import locale

# Configure a formatação de moeda para o Brasil
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

# Tabela 1
#col1, col2 = st.columns(2)
with col2:
    with st.expander("Visualizar Pagamentos por Unidades (Tabela 1) - Filtros"):
        region = filtered_df.groupby(by="Unidade", as_index=False)["Valor_total_item"].sum()
        # Formate a coluna "Valor_total_item" para moeda
        region["Valor_total_item"] = region["Valor_total_item"].apply(lambda x: locale.currency(x, grouping=True))
        st.write(region)
        
        xls_data = BytesIO()
        with pd.ExcelWriter(xls_data, engine="openpyxl", mode="xlsx") as writer:
            region.to_excel(writer, sheet_name="Unidade_Tabela1", index=False)
        st.download_button("Download Data (XLS - Tabela 1)", data=xls_data, file_name="Unidade_Tabela1.xlsx", key="unidade_xls_tabela1",
                        help="Clique aqui para baixar os dados como um arquivo em XLSX")


import streamlit as st
import pandas as pd
import plotly.express as px
import locale
import pyodbc

# Ajustar a formatação de moeda para o Brasil
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

def connect_to_database():
    dados_conexao = (
        "DRIVER={ODBC Driver 17 for SQL Server};"
        "SERVER=GS-GO-038\\SQLEXPRESS;"
        "DATABASE=BI_Sanus_Logistica;"
        "UID=sa;"
        "PWD=123;"
    )

    conexao = pyodbc.connect(dados_conexao)
    return conexao

# Função para buscar os dados do banco de dados
def get_data_from_database():
    # Conectar ao banco de dados
    conexao = connect_to_database()

    # Realizar a consulta SQL para obter os dados
    query = """
    SELECT UF, SUM(Valor_total_item) AS Valor_total_item
    FROM Estab_Protesis_Base_Consignados
    GROUP BY UF;
    """
    df = pd.read_sql_query(query, conexao)

    return df

# Buscar os dados do banco de dados
df = get_data_from_database()

# Dados de latitude e longitude para o Brasil
data = {
    "UF": ["AC", "MT", "RR", "GO"],
    "Latitude": [-10.0, -13.0, 1.0, -15.0],  # Latitude central do Brasil
    "Longitude": [-55.0, -57.0, -56.0, -49.0],  # Longitude central do Brasil
}

# Crie um DataFrame com os dados de latitude e longitude
df_coordinates = pd.DataFrame(data)

# Junte os DataFrames de faturamento e coordenadas com base na coluna "UF"
df = pd.merge(df, df_coordinates, on="UF", how="left")

# Aumente o tamanho da bolha multiplicando a coluna "Valor_total_item" por um fator (neste exemplo, 0.0001)
df["Valor_total_item"] = df["Valor_total_item"] * 0.0001

# Crie um gráfico de mapa estilo rodoviário com base nos dados de latitude e longitude
fig = px.scatter_geo(
    df,
    lat="Latitude",  # Use a coluna "Latitude" para as coordenadas de latitude
    lon="Longitude",  # Use a coluna "Longitude" para as coordenadas de longitude
    text="UF",  # Coluna com as informações adicionais a serem exibidas no hover
    size="Valor_total_item",  # Coluna com os tamanhos das bolhas
    projection="equirectangular",  # Estilo de mapa rodoviário
    title="Faturamento por UF no Brasil",  # Título do gráfico
)

# Personalize o layout do gráfico (opcional)
fig.update_geos(
    showcoastlines=True,
    coastlinecolor="Black",
    showland=True,
    landcolor="white",
    showocean=True,
    oceancolor="LightBlue",
    showcountries=True,  # Mostrar as linhas de fronteira do Brasil
    countrycolor="Black",  # Cor das linhas de fronteira
    showframe=False,  # Remover o quadro em volta do mapa
)

# Exiba o gráfico
st.plotly_chart(fig)

#------------------------------------------------------------------------------------

# A função generate_plots permanece a mesma
def generate_plots(data):
    category_df = data.groupby(by=["UF"], as_index=False).agg({"Valor_total_item": "sum"})
    return category_df


# Defina a localização para formatar em formato monetário (no caso, pt_BR para o Brasil)
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

# Aplicação dos filtros e exibição dos dados
col1, col2 = st.columns(2)
with col1:
    with st.expander("Visualizar Pagamentos por Estado (Tabela 2) - Filtros"):
        # Aplicação dos filtros
        filtered_category_df = generate_plots(filtered_df)

        # Formate a coluna "Valor_total_item" para moeda
        filtered_category_df["Valor_total_item"] = filtered_category_df["Valor_total_item"].apply(lambda x: locale.currency(x, grouping=True))

        # Exibir a tabela de dados filtrados
        st.dataframe(filtered_category_df, use_container_width=True)

        # Botão para baixar os dados em XLSX
        xls_data = BytesIO()
        with pd.ExcelWriter(xls_data, engine="openpyxl", mode="xlsx") as writer:
            filtered_category_df.to_excel(writer, sheet_name="UF_Tabela1", index=False)
        st.download_button("Baixar Dados (XLS - Tabela 1)", data=xls_data, file_name="UF_Tabela1.xlsx", key="uf_xls_tabela1",
                        help="Clique aqui para baixar os dados como um arquivo em XLSX")
        

# A função generate_plots permanece a mesma
def generate_plots(data):
    category_df = data.groupby(by=["Pagamento_Vencimento2"], as_index=False).agg({"Valor_total_item": "sum"})
    return category_df


# Defina a localização para formatar em formato monetário (no caso, pt_BR para o Brasil)
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

# Aplicação dos filtros e exibição dos dados
#col1, col2 = st.columns(2)
with col2:
    with st.expander("Visualizar Pagamentos por Vencimento - Filtros"):
        # Aplicação dos filtros
        filtered_category_df = generate_plots(filtered_df)

        # Formate a coluna "Valor_total_item" para moeda
        filtered_category_df["Valor_total_item"] = filtered_category_df["Valor_total_item"].apply(lambda x: locale.currency(x, grouping=True))

        # Exibir a tabela de dados filtrados
        st.dataframe(filtered_category_df, use_container_width=True)

        # Botão para baixar os dados em XLSX
        xls_data = BytesIO()
        with pd.ExcelWriter(xls_data, engine="openpyxl", mode="xlsx") as writer:
            filtered_category_df.to_excel(writer, sheet_name="Pagamento_Tabela1", index=False)
        st.download_button("Baixar Dados (XLS - Tabela 1)", data=xls_data, file_name="Pagamento_Tabela1.xlsx", key="pagamento_xls_tabela1",
                        help="Clique aqui para baixar os dados como um arquivo em XLSX")
        

        


def generate_plots(data):
    # Agrupe os dados originais por "Fornecedor" e "UF" e calcule a soma de "Valor_total_item" e a contagem de "Sequencia"
    category_df2 = data.groupby(by=["UF"], as_index=False).agg({"Valor_total_item": "sum", "Sequencia": "count"})

    # Renomeie as colunas
    category_df2 = category_df2.rename(columns={"Valor_total_item": "Valor_total_item", "Sequencia": "Nr Nota"})

    return category_df2





#Grafico 5
# Defina a localização para formatar em formato monetário (no caso, pt_BR para o Brasil)
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

# Converter a coluna 'Valor_total_item' para float no DataFrame original
df['Valor_total_item'] = pd.to_numeric(df['Valor_total_item'], errors='coerce')

# Chame a função para gerar os gráficos com base nos dados filtrados
category_df2 = generate_plots(filtered_df)

# Exibir o data frame com base nos filtros
#st.write(filtered_df)

# Exibir os gráficos
col1, col2 = st.columns(2)
with col1:
    st.subheader("Pagamento por Estado")
    fig = px.bar(category_df2, x="UF", y="Valor_total_item", text=['R${:,.2f}'.format(x) for x in category_df2["Valor_total_item"]], template="seaborn")
    fig.update_yaxes(title_text='')
    st.plotly_chart(fig, use_container_width=True)

with col2:
    st.subheader("Pagamento por Estado em %")
    fig = px.pie(category_df2, values="Valor_total_item", names="UF", hole=0.5)
    st.plotly_chart(fig, use_container_width=True)