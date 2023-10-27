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

st.title(":bar_chart: GERAL - ANÁLISE COORDENAÇÃO")
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

# Título para a análise Intranet
st.title("Análise Medtrauma")

# Crie uma linha de separação
st.write("___")

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


import locale

# Configure a formatação de moeda para o Brasil
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

# Tabela 2 (corte)
col1, col2 = st.columns(2)
with col1:
    with st.expander("Visualizar Pagamentos e NF por Fornecedores (Tabela 2) - Geral"):
        # Consulta SQL para obter os resultados desejados
        query = """
        SELECT Fornecedor, 
               SUM(Valor_total_item) AS Soma_Valor_total_item,
               COUNT(DISTINCT Sequencia) AS Contagem
        FROM Estab_Medtrauma_Base
        GROUP BY Fornecedor;
        """
        
        # Executa a consulta SQL e cria um DataFrame com os resultados
        result_df = pd.read_sql_query(query, conexao)
        
        # Formate a coluna "Soma_Valor_total_item" para moeda
        result_df["Soma_Valor_total_item"] = result_df["Soma_Valor_total_item"].apply(lambda x: locale.currency(x, grouping=True))

        # Exibe o DataFrame
        st.dataframe(result_df, use_container_width=True)
        
        xls_data = BytesIO()
        with pd.ExcelWriter(xls_data, engine="openpyxl", mode="xlsx") as writer:
            result_df.to_excel(writer, sheet_name="Fornecedor_Tabela2", index=False)  # Especifique um nome de folha
        st.download_button("Baixar Dados (XLS - Tabela 2)", data=xls_data, file_name="Fornecedor_Tabela2.xlsx", key="fornecedor_xls_tabela2",
                        help="Clique aqui para baixar os dados como um arquivo em XLSX")

        

# Tabela 2 (corte)
with col2:
    with st.expander("Visualizar Pagamentos e NF por Unidades (Tabela 2) - Geral"):
        # Consulta SQL para obter os resultados desejados
        query = """
        SELECT Unidade, 
            SUM(Valor_total_item) AS Soma_Valor_total_item,
            COUNT(DISTINCT Sequencia) AS Contagem
        FROM Estab_Medtrauma_Base
        GROUP BY Unidade;
        """

        # Executa a consulta SQL e cria um DataFrame com os resultados
        result_df = pd.read_sql_query(query, conexao)

        # Formate a coluna "Soma_Valor_total_item" para moeda
        result_df["Soma_Valor_total_item"] = result_df["Soma_Valor_total_item"].apply(lambda x: locale.currency(x, grouping=True))

        # Exibe o DataFrame
        st.dataframe(result_df, use_container_width=True)

        xls_data = BytesIO()
        with pd.ExcelWriter(xls_data, engine="openpyxl", mode="xlsx") as writer:
            result_df.to_excel(writer, sheet_name="Unidade_Tabela2", index=False)  # Especifique um nome de folha
        st.download_button("Baixar Dados (XLS - Tabela 2)", data=xls_data, file_name="Unidade_Tabela2.xlsx", key="unidade_xls_tabela2",
                            help="Clique aqui para baixar os dados como um arquivo em XLSX")


import locale

# Configure a formatação de moeda para o Brasil
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

# Tabela 4 (corte)
col3 = st.columns(1)  # Uma única coluna

with col3[0]:
    with st.expander("Visualizar Pagamentos por Fornecedores pela Medtrauma - Geral)"):
        query = """
        SELECT Ano_Emissao, Nome_Mes_Emissao,
            SUM(CASE WHEN Fornecedor = 'ASTRAMED' THEN Valor_total_item ELSE 0 END) AS ASTRAMED,
            SUM(CASE WHEN Fornecedor = 'BIOMED' THEN Valor_total_item ELSE 0 END) AS BIOMED,
            SUM(CASE WHEN Fornecedor = 'DR. MEDIC' THEN Valor_total_item ELSE 0 END) AS [DR. MEDIC],
            SUM(CASE WHEN Fornecedor = 'HEXAGON' THEN Valor_total_item ELSE 0 END) AS HEXAGON,
            SUM(CASE WHEN Fornecedor = 'PROTESIS' THEN Valor_total_item ELSE 0 END) AS PROTESIS,
            SUM(Valor_total_item) AS TOTAL
        FROM Estab_Medtrauma_Base
        GROUP BY Ano_Emissao, Nome_Mes_Emissao;
        """
        result_df = pd.read_sql_query(query, conexao)
        
        # Formate os campos numéricos para moeda
        num_cols = ["ASTRAMED", "BIOMED", "DR. MEDIC", "HEXAGON", "PROTESIS", "TOTAL"]
        for col in num_cols:
            result_df[col] = result_df[col].apply(lambda x: locale.currency(x, grouping=True))

        # Aplica o estilo para todas as colunas da tabela
        styled_df = result_df.style.applymap(lambda x: 'font-size: 12px;', subset=result_df.columns)

        # Escreve a tabela com estilo
        st.write(styled_df, unsafe_allow_html=True)

        xls_data = BytesIO()
        with pd.ExcelWriter(xls_data, engine="openpyxl", mode="xlsx") as writer:
            result_df.to_excel(writer, sheet_name="Geral_Tabela2", index=False)
        st.download_button("Baixar Dados", data=xls_data, file_name="Geral_Tabela2.xlsx", key="geral_xls_tabela2",
                        help="Clique aqui para baixar os dados como um arquivo em XLSX")


#-----------------------------------------------------------------------------------------------------------------

# Título para a análise Intranet
st.title("Análise Intranet")

# Crie uma linha de separação com espaço após o título
st.write("___")

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
comando = """SELECT * FROM Intranet2"""
df = pd.read_sql_query(comando, conexao)
df["Dt_Procedimento"] = pd.to_datetime(df["Dt_Procedimento"], format="%d/%m/%Y")  # Remove .str.split(" ", expand=True)[0]
df = df.sort_values("Dt_Procedimento")

# Exibir o data frame com base nos filtros
#st.write(df)


import locale

# Configure a formatação de moeda para o Brasil
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

# Tabela 4 (corte)
col3 = st.columns(1)  # Uma única coluna

with col3[0]:
    with st.expander("Produção Intranet das Unidades - Geral)"):
        query = """
        SELECT Unidade2,
        Ano_Emissao,
            SUM(CASE WHEN Nome_Mes_Emissao = 'Janeiro' THEN Total ELSE 0 END) AS Janeiro,
            SUM(CASE WHEN Nome_Mes_Emissao = 'Fevereiro' THEN Total ELSE 0 END) AS Fevereiro,
            SUM(CASE WHEN Nome_Mes_Emissao = 'Março' THEN Total ELSE 0 END) AS Março,
            SUM(CASE WHEN Nome_Mes_Emissao = 'Abril' THEN Total ELSE 0 END) AS Abril,
            SUM(CASE WHEN Nome_Mes_Emissao = 'Maio' THEN Total ELSE 0 END) AS Maio,
            SUM(CASE WHEN Nome_Mes_Emissao = 'Junho' THEN Total ELSE 0 END) AS Junho,
            SUM(CASE WHEN Nome_Mes_Emissao = 'Julho' THEN Total ELSE 0 END) AS Julho,
            SUM(CASE WHEN Nome_Mes_Emissao = 'Agosto' THEN Total ELSE 0 END) AS Agosto,
            SUM(CASE WHEN Nome_Mes_Emissao = 'Setembro' THEN Total ELSE 0 END) AS Setembro,
            SUM(CASE WHEN Nome_Mes_Emissao = 'Outubro' THEN Total ELSE 0 END) AS Outubro,
            SUM(CASE WHEN Nome_Mes_Emissao = 'Novembro' THEN Total ELSE 0 END) AS Novembro,
            SUM(CASE WHEN Nome_Mes_Emissao = 'Dezembro' THEN Total ELSE 0 END) AS Dezembro,
            SUM(Total) AS Total
        FROM Intranet2
        GROUP BY Unidade2, Ano_Emissao;
        """
        result_df = pd.read_sql_query(query, conexao)
        
        # Formate os campos numéricos para moeda
        num_cols = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro", "Total"]
        for col in num_cols:
            result_df[col] = result_df[col].apply(lambda x: locale.currency(x, grouping=True))

        # Aplica o estilo para todas as colunas da tabela
        styled_df = result_df.style.applymap(lambda x: 'font-size: 12px;', subset=result_df.columns)

        # Escreve a tabela com estilo
        st.dataframe(styled_df, use_container_width=True)

        xls_data = BytesIO()
        with pd.ExcelWriter(xls_data, engine="openpyxl", mode="xlsx") as writer:
            result_df.to_excel(writer, sheet_name="Geral_Tabela3", index=False)
        st.download_button("Baixar Dados", data=xls_data, file_name="Geral_Tabela3.xlsx", key="geral_xls_tabela3",
                        help="Clique aqui para baixar os dados como um arquivo em XLSX")



# Configure a formatação de moeda para o Brasil
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

# Tabela 4 (corte)
col4 = st.columns(1)  # Uma única coluna

with col4[0]:
    with st.expander("Produção por Procedimento Intranet - Geral"):
        query = """
        SELECT Procedimento, COUNT(*) AS Contagem
            FROM (
                SELECT DISTINCT Procedimento, Dt_Procedimento, Nome
                FROM Intranet2
            ) AS Sub
            GROUP BY Procedimento
            ORDER BY Contagem DESC; -- Classifica por Contagem em ordem decrescente
        """
        # Substitua o código de conexão 'conexao' pelo seu código de conexão ao banco de dados

        # Resultado da consulta SQL
        result_df = pd.read_sql_query(query, conexao)

        # Formate a coluna "Contagem" como número inteiro
        num_cols = ["Contagem"]
        for col in num_cols:
            result_df[col] = result_df[col].apply(lambda x: f"{int(x):,}")

        # Exibe a tabela
        st.dataframe(result_df, use_container_width=True)

        # Salva a tabela como um arquivo XLSX
        xls_data = BytesIO()
        with pd.ExcelWriter(xls_data, engine="xlsxwriter", mode="xlsx", date_format='dd-mm-yyyy') as writer:
            result_df.to_excel(writer, sheet_name="Geral_Tabela4", index=False)

        # Cria o botão para baixar os dados
        st.download_button("Baixar Dados", data=xls_data, file_name="Geral_Tabela4.xlsx", key="geral_xls_tabela4",
                           help="Clique aqui para baixar os dados como um arquivo em XLSX")


# Consulta SQL para contar o número de procedimentos por Unidade2
query_procedimentos = """
SELECT Unidade2, COUNT(*) AS Contagem
FROM (
    SELECT DISTINCT Unidade2, Dt_Procedimento, Nome
    FROM Intranet2
) AS Sub
GROUP BY Unidade2;
"""

# Resultado da consulta SQL
result_procedimentos = pd.read_sql_query(query_procedimentos, conexao)

# Ordenar o DataFrame pelo valor "Contagem" de forma decrescente
result_procedimentos = result_procedimentos.sort_values(by="Contagem", ascending=False)

# Exibir os gráficos um abaixo do outro
# Exibir os gráficos um abaixo do outro
st.subheader("Procedimento por Hospital Barras")
fig = px.bar(result_procedimentos, x="Unidade2", y="Contagem", text=result_procedimentos["Contagem"], template="seaborn")
fig.update_yaxes(title_text='')
st.plotly_chart(fig, use_container_width=True)

st.subheader("Procedimento por Hospital Pizza %")
fig = px.pie(result_procedimentos, values="Contagem", names="Unidade2", hole=0.5)
st.plotly_chart(fig, use_container_width=True)


# Tabela 4 (corte)
col3 = st.columns(1)  # Uma única coluna

with col3[0]:
    with st.expander("Produção Mensal de Implantes - Geral)"):
        query = """
        SELECT
            Unidade2,
            Ano_Emissao,
            SUM(CASE WHEN Nome_Mes_Emissao = 'Janeiro' THEN Qtd_Opme ELSE 0 END) AS Janeiro,
            SUM(CASE WHEN Nome_Mes_Emissao = 'Fevereiro' THEN Qtd_Opme ELSE 0 END) AS Fevereiro,
            SUM(CASE WHEN Nome_Mes_Emissao = 'Março' THEN Qtd_Opme ELSE 0 END) AS Março,
            SUM(CASE WHEN Nome_Mes_Emissao = 'Abril' THEN Qtd_Opme ELSE 0 END) AS Abril,
            SUM(CASE WHEN Nome_Mes_Emissao = 'Maio' THEN Qtd_Opme ELSE 0 END) AS Maio,
            SUM(CASE WHEN Nome_Mes_Emissao = 'Junho' THEN Qtd_Opme ELSE 0 END) AS Junho,
            SUM(CASE WHEN Nome_Mes_Emissao = 'Julho' THEN Qtd_Opme ELSE 0 END) AS Julho,
            SUM(CASE WHEN Nome_Mes_Emissao = 'Agosto' THEN Qtd_Opme ELSE 0 END) AS Agosto,
            SUM(CASE WHEN Nome_Mes_Emissao = 'Setembro' THEN Qtd_Opme ELSE 0 END) AS Setembro,
            SUM(CASE WHEN Nome_Mes_Emissao = 'Outubro' THEN Qtd_Opme ELSE 0 END) AS Outubro,
            SUM(CASE WHEN Nome_Mes_Emissao = 'Novembro' THEN Qtd_Opme ELSE 0 END) AS Novembro,
            SUM(CASE WHEN Nome_Mes_Emissao = 'Dezembro' THEN Qtd_Opme ELSE 0 END) AS Dezembro,
            SUM(Qtd_Opme) AS Qtd_Opme
        FROM Intranet2
        GROUP BY Unidade2, Ano_Emissao;
        """
        result_df = pd.read_sql_query(query, conexao)
        
        # Formate os campos numéricos como números inteiros
        num_cols = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro", "Qtd_Opme"]
        for col in num_cols:
            result_df[col] = result_df[col].astype(int)

        # Aplica o estilo para todas as colunas da tabela
        styled_df = result_df.style.applymap(lambda x: 'font-size: 12px;', subset=result_df.columns)

        # Reordena as colunas para mover Qtd_Opme para o final
        result_df = result_df[["Unidade2", "Ano_Emissao", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro", "Qtd_Opme"]]

        # Escreve a tabela com estilo
        st.dataframe(styled_df, use_container_width=True)

        xls_data = BytesIO()
        with pd.ExcelWriter(xls_data, engine="openpyxl", mode="xlsx") as writer:
            result_df.to_excel(writer, sheet_name="Geral_Tabela6", index=False)
        st.download_button("Baixar Dados", data=xls_data, file_name="Geral_Tabela6.xlsx", key="geral_xls_tabela6",
                        help="Clique aqui para baixar os dados como um arquivo em XLSX")



# Tabela 4 (corte)
col3 = st.columns(1)  # Uma única coluna

with col3[0]:
    with st.expander("Produção Quantidade de OPME por Valor Cobrado - Geral"):
        order_by = st.radio("Ordenar por:", ["Qtde_Opme DESC", "Total DESC"])

        query = f"""
        SELECT Desc_Opme, SUM(Qtd_Opme) AS Qtde_Opme, SUM(Total) AS Total
        FROM Intranet2
        GROUP BY Desc_Opme
        ORDER BY {order_by};
        """
        result_df = pd.read_sql_query(query, conexao)
        
        # Formate os campos numéricos como números inteiros
        num_cols = ["Qtde_Opme"]
        for col in num_cols:
            result_df[col] = result_df[col].astype(int)

        # Formate a coluna "Total" como valores em moeda
        result_df["Total"] = result_df["Total"].apply(lambda x: locale.currency(x, grouping=True))

        # Aplica o estilo para todas as colunas da tabela
        styled_df = result_df.style.applymap(lambda x: 'font-size: 12px;', subset=result_df.columns)

        # Escreve a tabela com estilo
        st.dataframe(styled_df, use_container_width=True)

        xls_data = BytesIO()
        with pd.ExcelWriter(xls_data, engine="openpyxl", mode="xlsx") as writer:
            result_df.to_excel(writer, sheet_name="Geral_Tabela6", index=False)
        st.download_button("Baixar Dados", data=xls_data, file_name="Geral_Tabela6.xlsx", key="geral_xls_tabela6_2",
                    help="Clique aqui para baixar os dados como um arquivo em XLSX")


# Configure a formatação de moeda para o Brasil
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

# Tabela 4 (corte)
col4 = st.columns(1)  # Uma única coluna

with col4[0]:
    with st.expander("Procedimento por Médico Intranet - Geral"):
        query = """
        SELECT Cirurgiao, COUNT(*) AS Contagem
        FROM (
            SELECT DISTINCT Cirurgiao, Dt_Procedimento, Nome
            FROM Intranet2
        ) AS Sub
        GROUP BY Cirurgiao
        ORDER BY Contagem DESC;
        """
        # Substitua o código de conexão 'conexao' pelo seu código de conexão ao banco de dados

        # Resultado da consulta SQL
        # Substitua este bloco de código pela sua consulta e conexão ao banco de dados
        result_df = pd.read_sql_query(query, conexao)

        # Formate a coluna "Contagem" como número inteiro
        num_cols = ["Contagem"]
        for col in num_cols:
            result_df[col] = result_df[col].apply(lambda x: f"{int(x):,}")

        # Exibe a tabela
        st.dataframe(result_df, use_container_width=True)

        # Salva a tabela como um arquivo XLSX
        xls_data = BytesIO()
        with pd.ExcelWriter(xls_data, engine="xlsxwriter", mode="xlsx", date_format='dd-mm-yyyy') as writer:
            result_df.to_excel(writer, sheet_name="Geral_Tabela8", index=False)

        # Cria o botão para baixar os dados
        st.download_button("Baixar Dados", data=xls_data, file_name="Geral_Tabela8.xlsx", key="geral_xls_tabela8",
                           help="Clique aqui para baixar os dados como um arquivo em XLSX")



# Tabela 4 (corte)
col3 = st.columns(1)  # Uma única coluna

with col3[0]:
    with st.expander("Produção Cirurgião de OPME por Valor Cobrado - Geral"):
        order_by = st.radio("Ordenar por:", ["Qtde_Opme DESC", "Total DESC"], key="order_by_radio")  # Adicione a chave "order_by_radio"

        # Query SQL correta
        query = """
        SELECT Cirurgiao, SUM(Qtd_Opme) AS Qtde_Opme, SUM(Total) AS Total
        FROM Intranet2
        GROUP BY Cirurgiao
        ORDER BY Qtde_Opme DESC;
        """
        result_df = pd.read_sql_query(query, conexao)
        
        # Formate os campos numéricos como números inteiros
        num_cols = ["Qtde_Opme"]
        for col in num_cols:
            result_df[col] = result_df[col].astype(int)

        # Formate a coluna "Total" como valores em moeda
        result_df["Total"] = result_df["Total"].apply(lambda x: locale.currency(x, grouping=True))

        # Aplica o estilo para todas as colunas da tabela
        styled_df = result_df.style.applymap(lambda x: 'font-size: 12px;', subset=result_df.columns)

        # Escreve a tabela com estilo
        st.dataframe(styled_df, use_container_width=True)

        xls_data = BytesIO()
        with pd.ExcelWriter(xls_data, engine="openpyxl", mode="xlsx") as writer:
            result_df.to_excel(writer, sheet_name="Geral_Tabela9", index=False)
        st.download_button("Baixar Dados", data=xls_data, file_name="Geral_Tabela9.xlsx", key="geral_xls_tabela9_2",
                    help="Clique aqui para baixar os dados como um arquivo em XLSX")
        














#--------------------------------------------###------------------------------------------------------------------
# Título para a análise Intranet
st.title("Análise Protesis")

# Crie uma linha de separação
st.write("___")

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


import locale

# Configure a formatação de moeda para o Brasil
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

# Tabela 2 (corte)
col1, col2 = st.columns(2)
with col1:
    with st.expander("Visualizar Pagamentos e NF por Fornecedores (Tabela 2) - Geral"):
        # Consulta SQL para obter os resultados desejados
        query = """
        SELECT Fornecedor, 
               SUM(Valor_total_item) AS Soma_Valor_total_item,
               COUNT(DISTINCT Sequencia) AS Contagem
        FROM Estab_Protesis_Base_Consignados
        GROUP BY Fornecedor;
        """
        
        # Executa a consulta SQL e cria um DataFrame com os resultados
        result_df = pd.read_sql_query(query, conexao)
        
        # Formate a coluna "Soma_Valor_total_item" para moeda
        result_df["Soma_Valor_total_item"] = result_df["Soma_Valor_total_item"].apply(lambda x: locale.currency(x, grouping=True))

        # Exibe o DataFrame
        st.dataframe(result_df, use_container_width=True)
        
        xls_data = BytesIO()
        with pd.ExcelWriter(xls_data, engine="openpyxl", mode="xlsx") as writer:
            result_df.to_excel(writer, sheet_name="Fornecedor_Tabela22", index=False)  # Especifique um nome de folha
        st.download_button("Baixar Dados (XLS - Tabela 2)", data=xls_data, file_name="Fornecedor_Tabela22.xlsx", key="fornecedor_xls_tabela22",
                        help="Clique aqui para baixar os dados como um arquivo em XLSX")

        

# Tabela 2 (corte)
with col2:
    with st.expander("Visualizar Pagamentos e NF por Unidades (Tabela 2) - Geral"):
        # Consulta SQL para obter os resultados desejados
        query = """
        SELECT Unidade, 
            SUM(Valor_total_item) AS Soma_Valor_total_item,
            COUNT(DISTINCT Sequencia) AS Contagem
        FROM Estab_Protesis_Base_Consignados
        GROUP BY Unidade;
        """

        # Executa a consulta SQL e cria um DataFrame com os resultados
        result_df = pd.read_sql_query(query, conexao)

        # Formate a coluna "Soma_Valor_total_item" para moeda
        result_df["Soma_Valor_total_item"] = result_df["Soma_Valor_total_item"].apply(lambda x: locale.currency(x, grouping=True))

        # Exibe o DataFrame
        st.dataframe(result_df, use_container_width=True)

        xls_data = BytesIO()
        with pd.ExcelWriter(xls_data, engine="openpyxl", mode="xlsx") as writer:
            result_df.to_excel(writer, sheet_name="Unidade_Tabela23", index=False)  # Especifique um nome de folha
        st.download_button("Baixar Dados (XLS - Tabela 2)", data=xls_data, file_name="Unidade_Tabela23.xlsx", key="unidade_xls_tabela23",
                            help="Clique aqui para baixar os dados como um arquivo em XLSX")


import locale

# Configure a formatação de moeda para o Brasil
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

# Tabela 4 (corte)
col3 = st.columns(1)  # Uma única coluna

with col3[0]:
    with st.expander("Visualizar Pagamentos por Fornecedores pela Medtrauma - Geral)"):
        query = """
        SELECT Ano_Emissao, Nome_Mes_Emissao,
            SUM(CASE WHEN Fornecedor = 'ASTRAMED' THEN Valor_total_item ELSE 0 END) AS ASTRAMED,
            SUM(CASE WHEN Fornecedor = 'BIOMED' THEN Valor_total_item ELSE 0 END) AS BIOMED,
            SUM(CASE WHEN Fornecedor = 'DR. MEDIC' THEN Valor_total_item ELSE 0 END) AS [DR. MEDIC],
            SUM(CASE WHEN Fornecedor = 'HEXAGON' THEN Valor_total_item ELSE 0 END) AS HEXAGON,
            SUM(CASE WHEN Fornecedor = 'PROTESIS' THEN Valor_total_item ELSE 0 END) AS PROTESIS,
            SUM(Valor_total_item) AS TOTAL
        FROM Estab_Protesis_Base_Consignados
        GROUP BY Ano_Emissao, Nome_Mes_Emissao;
        """
        result_df = pd.read_sql_query(query, conexao)
        
        # Formate os campos numéricos para moeda
        num_cols = ["ASTRAMED", "BIOMED", "DR. MEDIC", "HEXAGON", "PROTESIS", "TOTAL"]
        for col in num_cols:
            result_df[col] = result_df[col].apply(lambda x: locale.currency(x, grouping=True))

        # Aplica o estilo para todas as colunas da tabela
        styled_df = result_df.style.applymap(lambda x: 'font-size: 12px;', subset=result_df.columns)

        # Escreve a tabela com estilo
        st.write(styled_df, unsafe_allow_html=True)

        xls_data = BytesIO()
        with pd.ExcelWriter(xls_data, engine="openpyxl", mode="xlsx") as writer:
            result_df.to_excel(writer, sheet_name="Geral_Tabela24", index=False)
        st.download_button("Baixar Dados", data=xls_data, file_name="Geral_Tabela24.xlsx", key="geral_xls_tabela24",
                        help="Clique aqui para baixar os dados como um arquivo em XLSX")

