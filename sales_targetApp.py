import pandas as pd
import streamlit as st
from PIL import Image
# import plotly.graph_objects as go
import calendar
import io
import os
import zipfile
from io import BytesIO
pip install openpyxl

# Criando essas três variáveis que serão usadas ao longo do script. O Ano anterior foi criado somente para
# se comparar com o ano atual e verificar o percentual de crescimento entre eles. Próximo ano é usado para alguns
# nomes.

ano_corrente = 2023
ano_anterior = ano_corrente - 1
proximo_ano = ano_corrente + 1

######################################### Tab Icon ################################

st.set_page_config(
    page_title='Metas 2024 - COUTO',
    page_icon=Image.open(
        'Imagens/coin.png'),
    layout='wide'
)

################################# ETL e cache ##########################

# Fazendo o cache para carregar os dados que vamos utilizar somente uma vez.

@st.cache_resource
def load_dataset():
    # Este load é do csv que traz todas as vendas
    df_function_1 = pd.read_csv(
        'DF.csv')

    # Este load é do csv que traz especificamente as medias mensualizadas dos produtos
    df_function_2 = pd.read_csv('medias_budget_2020_a_2022.csv')

    # Na montagem do DF.csv, há uma manipulação para criar valores zerados, de forma a corrigir inconsistências
    # no power bi. Contudo, essa modificação lá, prejudica aqui. Portanto, para excluir esses valores criados
    # vou filtrar para não pegar 'CASA' ou 'CLIN' na coluna de Subdivisão. (a subdivisão, originalmente,
    # não possui valores exatamente como CASA ou CLIN, mas sim CASA... e CLIN..., portanto, fica fácil de
    # saber quais são inserções falsas).
    df_function_1 = df_function_1[~df_function_1['Subdivisão'].isin([
                                                                    'CASA', 'CLIN'])]

    # puxando somente o ano atual e o anterior, para pesar menos no rerun do streamlit.
    df_function_1 = df_function_1[(df_function_1['No_ano'] == ano_corrente) | (
        df_function_1['No_ano'] == ano_anterior)]

    # O csv estava com string na quantidade e vendas para a importacao no PowerBi
    df_function_1['Quantidade'] = df_function_1['Quantidade'].astype(str).str.replace(
        ',', '.').astype(float)

    # Portanto, to transformando de volta para float, para trabalhar com operações no python
    df_function_1['Valor_vendas'] = df_function_1['Valor_vendas'].astype(str).str.replace(
        ',', '.').astype(float)

    # como os codigos sao só numeros, é necessário transformar em string, pois são códigos e não devem ser usados em operações
    df_function_1['Produto'] = df_function_1['Produto'].astype(str)
    df_function_2['Produto'] = df_function_2['Produto'].astype(str)

    df_function_1.rename(columns={'Valor_vendas': 'Vendas'}, inplace=True)
    df_function_2 = df_function_2.rename(
        columns={'Subdivisão_agrupada': 'Subdivisão'})

    # criando a concatenação de produto com subdivisão, para facilitar o filtro de exclusão, a ser visto mais a frente
    df_function_1['Produto_Subdivisão'] = df_function_1['Produto'] + \
        '_' + df_function_1['Subdivisão_agrupada']
    return df_function_1, df_function_2

#################################### Título ################################

st.markdown('''# Metas 2024 ''')
st.markdown('''---''')

################################### Filtros ##################################
st.markdown(""" **FILTROS** """)
st.text('')

# Carregar o dataset
df = load_dataset()
df1_filters = df[0].copy()

# filtros visíveis:

col1, col2, col3, col4 = st.columns(4)

# filtro area_negocio
area_filter = col1.multiselect(
    key=1,
    label="Área de Negócio",
    options=sorted(df1_filters['Área_negócio'].unique()),
    placeholder='Todos'
)

select_all_area = col1.checkbox("Selecionar Todos", key=2, value=True)

if select_all_area:
    df1_filters = df1_filters.query('Área_negócio == Área_negócio')
else:
    df1_filters = df1_filters.query('Área_negócio == @area_filter')

# filtro gama
gama_filter = col2.multiselect(
    key=3,
    label="Gama",
    options=sorted(df1_filters['Gama'].unique()),
    placeholder='Todos'
)

select_all_gama = col2.checkbox("Selecionar Todos", key=4, value=True)

if select_all_gama:
    df1_filters = df1_filters.query('Gama == Gama')
else:
    df1_filters = df1_filters.query('Gama == @gama_filter')

# filtro descrição
descricao_filter = col3.multiselect(
    key=5,
    label="Descrição",
    options=sorted(df1_filters['Descrição'].unique()),
    placeholder='Todos'
)

select_all_descricao = col3.checkbox("Selecionar Todos", key=6, value=True)

if select_all_descricao:
    df1_filters = df1_filters.query('Descrição == Descrição')
else:
    df1_filters = df1_filters.query('Descrição == @descricao_filter')

# filtro subdivisão
subdiv_filter = col4.multiselect(
    key=7,
    label="Subdivisão",
    options=df1_filters['Subdivisão_agrupada'].unique(),
    placeholder='Todos'
)

select_all_subdiv = col4.checkbox("Selecionar Todos", key=8, value=True)

if select_all_subdiv:
    df1_filters = df1_filters.query(
        'Subdivisão_agrupada == Subdivisão_agrupada')
else:
    df1_filters = df1_filters.query('Subdivisão_agrupada == @subdiv_filter')

# filtro excluir produtos que não quero
exclusao_filter = st.multiselect(
    key=9,
    label="Exclusão",
    options=df1_filters['Produto_Subdivisão'].unique(),
    placeholder='Exclua os produtos que não participarão da criação do Budget'
)
df1_filters = df1_filters.query('Produto_Subdivisão != @exclusao_filter')

st.markdown('''---''')
######################################### Modelando as colunas do Dataframe Budget (1) ##########################################

df2_corrente = df1_filters.query('No_ano == @ano_corrente')
df3_anterior = df1_filters.query('No_ano == @ano_anterior')

# iniciando a montagem do dataframe
meses_jan_set = range(1, 10)  # Filtrando os meses do 1 a 9 (Jan a Set)

# Criando os dois dataframes, com base nos seus respectivos anos, para depois fazer um merge deles e finalmente
# poder criar as colunas de comparação de crescimento.
df4_JanSet_corrente = df2_corrente[df2_corrente['No_mês'].isin(meses_jan_set)].groupby(['Área_negócio', 'Gama', 'Descrição',
                                                                                        'Descrição_detalhada', 'Produto', 'Subdivisão_agrupada'])[['Quantidade', 'Vendas']].sum().reset_index()  # buscando a quantidade vendida de Jan a Set

df5_JanSet_anterior = df3_anterior[df3_anterior['No_mês'].isin(meses_jan_set)].groupby(['Área_negócio', 'Gama', 'Descrição',
                                                                                        'Descrição_detalhada', 'Produto', 'Subdivisão_agrupada'])[['Quantidade', 'Vendas']].sum().reset_index()  # buscando a quantidade vendida de Jan a Set

# Realizando o merge externo para garantir que todas as combinações de produtos estejam presentes
df6_merge = df4_JanSet_corrente.merge(df5_JanSet_anterior, on=['Área_negócio', 'Gama', 'Descrição',
                                                               'Descrição_detalhada', 'Produto', 'Subdivisão_agrupada'], how='outer', suffixes=('_corrente', '_anterior'))

# Preenchendo os valores ausentes (NaN) com 0, para que a divisão que vou fazer funcione corretamente
df6_merge.fillna({'Quantidade_corrente': 0, 'Quantidade_anterior': 0,
                  'Vendas_corrente': 0, 'Vendas_anterior': 0}, inplace=True)

# Criando as colunas com o percentual de crescimento
df6_merge['% N-1 (Qtd)'] = round(((df6_merge['Quantidade_corrente'] -
                                   df6_merge['Quantidade_anterior']) / df6_merge['Quantidade_anterior']) * 100, 1)
df6_merge['% N-1 (€)'] = round(((df6_merge['Vendas_corrente'] -
                                 df6_merge['Vendas_anterior']) / df6_merge['Vendas_anterior']) * 100, 1)

# Adicionando as colunas de percentual de crescimento ao df4_JanSet_corrente, gerando a df final: df7_selectedColumns
df7_selectedColumns = df4_JanSet_corrente.merge(df6_merge[['Área_negócio', 'Gama', 'Descrição',
                                                           'Descrição_detalhada', 'Produto', 'Subdivisão_agrupada', '% N-1 (Qtd)', '% N-1 (€)']], on=['Área_negócio', 'Gama', 'Descrição',
                                                                                                                                                      'Descrição_detalhada', 'Produto', 'Subdivisão_agrupada'], how='left')

# colocando as colunas na ordem mais apropriada para entendimento
df7_selectedColumns = df7_selectedColumns[['Área_negócio', 'Gama', 'Descrição',
                                           'Descrição_detalhada', 'Produto', 'Subdivisão_agrupada', 'Quantidade',
                                           '% N-1 (Qtd)', 'Vendas', '% N-1 (€)']]

# ordenando por ordem alfabética algumas colunas e uma pela ordem descendente
df7_selectedColumns = df7_selectedColumns.sort_values(
    by=['Gama', 'Descrição', 'Descrição_detalhada', 'Subdivisão_agrupada'], ascending=[True, True, True, False])

# criando as próximas colunas

df7_selectedColumns['Prev. Out-Dez (Qtd)'] = round(
    (df7_selectedColumns['Quantidade']/9)*2.5)  # achando a média
df7_selectedColumns['Prev. Out-Dez (€)'] = round(
    (df7_selectedColumns['Vendas']/9)*2.5, 2)  # achando a média
df7_selectedColumns.rename(columns={'Quantidade': 'Jan-Set (Qtd)', 'Vendas': 'Jan-Set (€)', 'Área_negócio': 'Área',
                                    'Descrição_detalhada': 'Descrição detalhada',
                                    'Subdivisão_agrupada': 'Subdivisão'}, inplace=True)  # renomeando as colunas
df7_selectedColumns['Total Ano (Qtd)'] = df7_selectedColumns['Jan-Set (Qtd)'] + \
    df7_selectedColumns['Prev. Out-Dez (Qtd)']  # colocando a previsão do total do ano
df7_selectedColumns['Total Ano (€)'] = df7_selectedColumns['Jan-Set (€)'] + \
    df7_selectedColumns['Prev. Out-Dez (€)']  # colocando a previsão do total do ano

##################### Funções para salvar e carregar o csv das metas incluídas no formulário #####################

# Criando função para carregar


def carregar_metas():
    try:
        df8_metasCsv_function = pd.read_csv(
            "Couto_metas_2024.csv")
        df8_metasCsv_function['Produto'] = df8_metasCsv_function['Produto'].astype(
            'str')
        df8_metasCsv_function['Subdivisão'] = df8_metasCsv_function['Subdivisão'].astype(
            'str')
        df8_metasCsv_function['Meta (%)'] = df8_metasCsv_function['Meta (%)'].astype(
            'float64')
    except FileNotFoundError:
        df8_metasCsv_function = pd.DataFrame(
            columns=["Produto", "Subdivisão", "Meta (%)"])
    return df8_metasCsv_function

# Função para salvar as metas em um arquivo CSV


def salvar_metas():
    df8_metasCsv.to_csv(
        "Couto_metas_2024.csv", index=False)


# Carregando o csv
df8_metasCsv = carregar_metas()

######################################################## Sidebar: Logotipo ####################################################

logotipo = Image.open(
    'Imagens/logo_couto.png')

st.sidebar.image(logotipo, use_column_width=True)
st.sidebar.text('')
st.sidebar.text('')
st.sidebar.text('')
st.sidebar.text('')

########################################################## Sidebar: Formulário Inserir Metas ##########################################

# Montando o formulario de preenchimento da meta
form1 = st.sidebar.form(key='formulario_metas', clear_on_submit=True)

form1.header("Inserir Metas")

codigo_produtos = form1.multiselect(
    'Código do Produto',
    options=sorted(df7_selectedColumns['Produto'].unique()),
    placeholder='Escolha o(s) produto(s)'
)
subdivisões = form1.multiselect(
    'Subdivisão',
    options=df7_selectedColumns['Subdivisão'].unique(),
    placeholder='Escolha CASA e/ou CLIN'
)
meta = form1.number_input("Meta (%)", value=0.0, step=0.1)

form1.text('')


# Botão para adicionar a meta
if form1.form_submit_button('Confirmar'):
    for subdivisão in subdivisões:
        for codigo_produto in codigo_produtos:
            # Verificar se a meta já existe no DataFrame
            filtro = (df8_metasCsv['Produto'] == codigo_produto) & (
                df8_metasCsv['Subdivisão'] == subdivisão)
            if filtro.any():
                # Atualizar a meta existente
                df8_metasCsv.loc[filtro, 'Meta (%)'] = meta
            else:
                # Adicionar uma nova meta
                nova_meta = pd.DataFrame({"Produto": [codigo_produto], "Subdivisão": [
                                         subdivisão], "Meta (%)": [meta]})
                df8_metasCsv = pd.concat(
                    [df8_metasCsv, nova_meta], ignore_index=True)

    # Salvar as metas no arquivo CSV
    salvar_metas()

# Design do formulario
st.markdown("""
    <style>
    .st-emotion-cache-lxzzm9 {
            background-color: rgb(245, 245, 245);
            width: -webkit-fill-available;
    }
    
    .st-emotion-cache-16txtl3 h2 {
            text-align : center;
            color : white;
            font-weight: bold;
            font-size: 1.8rem;
    }
    
    .st-emotion-cache-r421ms {
            border: 1px solid rgba(255, 255, 255, 2);
    }
            
    div[data-testid="stForm"] .st-emotion-cache-16idsys p {
            color: white;
    }
            
    </style>
""",
            unsafe_allow_html=True
            )

############################################# Modelando as colunas do Dataframe Budget (2) ##############################################

# Após clicar no botão, a meta é salva no csv.
# Agora eu faço uma concatenação para pegar a meta do csv e colocar no dataframe principal
# onde o codigo do produto e a subdivisao sejam os mesmos

meta_dict = dict(zip(
    df8_metasCsv['Produto'] + df8_metasCsv['Subdivisão'], df8_metasCsv['Meta (%)']))
df7_selectedColumns['Meta (%)'] = df7_selectedColumns['Produto'] + \
    df7_selectedColumns['Subdivisão']
df7_selectedColumns['Meta (%)'] = df7_selectedColumns['Meta (%)'].map(
    meta_dict)

# Aqui eu crio as ultimas duas colunas, para calcular, com base na meta informada, qual será
# a quantidade/venda final para o proximo ano e o número acrescentado.

df7_selectedColumns['Qtd próx. ano'] = round(
    ((df7_selectedColumns['Meta (%)']/100) * df7_selectedColumns['Total Ano (Qtd)'])) + df7_selectedColumns['Total Ano (Qtd)']
df7_selectedColumns['Dif. (Qtd)'] = df7_selectedColumns['Qtd próx. ano'] - \
    df7_selectedColumns['Total Ano (Qtd)']
df7_selectedColumns['Vendas próx. ano (€)'] = round(
    ((df7_selectedColumns['Meta (%)']/100) * df7_selectedColumns['Total Ano (€)']), 2) + df7_selectedColumns['Total Ano (€)']
df7_selectedColumns['Dif. (€)'] = df7_selectedColumns['Vendas próx. ano (€)'] - \
    df7_selectedColumns['Total Ano (€)']
df7_selectedColumns['Preço médio (€)'] = round(
    (df7_selectedColumns['Vendas próx. ano (€)']/df7_selectedColumns['Qtd próx. ano']), 2)

############################################# Design Dataframe Budget ##############################################

# Design do dataframe, intercalando cores toda vez que a gama, descrição e descrição detalhada mudar
# Se eu tiver só 1 gama no dataframe, ele analisa se a descrição muda. Se eu so tiver 1 gama e uma descrição, ele verifica a detalhada.

prev_data = None
row_styles = []
current_style = None

df9_design = df7_selectedColumns.copy()

if len(df9_design['Gama'].unique()) > 1:
    # Loop through rows and apply row styles
    for index, row in df9_design.iterrows():
        current_data = row['Gama']

        if current_data != prev_data:
            # Change the current style when the product code changes
            current_style = 'background-color: #ffffff' if current_style != 'background-color: #ffffff' else 'background-color: #f2f7f2'

        row_styles.append(current_style)
        prev_data = current_data

else:
    if len(df9_design['Descrição'].unique()) > 1:
        # Loop through rows and apply row styles
        for index, row in df9_design.iterrows():
            current_data = row['Descrição']

            if current_data != prev_data:
                # Change the current style when the product code changes
                current_style = 'background-color: #ffffff' if current_style != 'background-color: #ffffff' else 'background-color: #f2f7f2'

            row_styles.append(current_style)
            prev_data = current_data

    else:
        # Loop through rows and apply row styles
        for index, row in df9_design.iterrows():
            current_data = row['Descrição detalhada']

            if current_data != prev_data:
                # Change the current style when the product code changes
                current_style = 'background-color: #ffffff' if current_style != 'background-color: #ffffff' else 'background-color: #f2f7f2'

            row_styles.append(current_style)
            prev_data = current_data


# Aplicando os estilos acima
df9_design = df9_design.style.apply(lambda x: row_styles, axis=0)

# Configurando os formatos das casas decimas e do separador de milhares
df9_design = df9_design.format({'% N-1 (Qtd)': "{:,.1f}".format,
                                '% N-1 (€)': "{:,.1f}".format,
                                'Jan-Set (Qtd)': "{:,.0f}".format,
                                'Prev. Out-Dez (Qtd)': "{:,.0f}".format,
                                'Total Ano (Qtd)': "{:,.0f}".format,
                                'Meta (%)': "{:,.1f}".format,
                                'Qtd próx. ano': "{:,.0f}".format,
                                'Dif. (Qtd)': "{:,.0f}".format,
                                'Jan-Set (€)': "{:,.2f}".format,
                                'Prev. Out-Dez (€)': "{:,.2f}".format,
                                'Total Ano (€)': "{:,.2f}".format,
                                'Vendas próx. ano (€)': "{:,.2f}".format,
                                'Dif. (€)': "{:,.2f}".format,
                                'Preço médio (€)': "{:,.2f}".format})

# Criando uma formatação de cores dos números, onde será verde se positivo e vermelho se negativo


def color_negative_red(val):
    color = 'red' if val < 0 else 'green'
    return f'color: {color};'


# Aplicando a formatação nas colunas específicas
df9_design = df9_design.applymap(color_negative_red, subset=[
                                 '% N-1 (Qtd)', '% N-1 (€)'])

###################################################### Postando o Dataframe (Budget) #################################################
st.markdown(""" **METAS** """)

# Postando o dataframe final e colocando um balão de informação ao passar o mouse sobre
# os cabeçalhos, para explicar ao usuário.

st.dataframe(
    df9_design,
    column_config={
        "Preço médio (€)": st.column_config.NumberColumn(
            help="Valor resultante da divisão entre 'Vendas próx. ano (€)' e 'Qtd próx. ano'."
        ),
        "% N-1 (Qtd)": st.column_config.NumberColumn(
            help="Mostra o percentual de crescimento da quantidade vendida em comparação com o ano anterior \
                  (somente Jan a Set de cada ano). \
                Valores zerados -> significa que a quantidade vendida ano passado e este ano foram iguais. \
                Valores inf -> significa que ano passado não houve vendas ou o produto não existia."
        ),
        "% N-1 (€)": st.column_config.NumberColumn(
            help="Mostra o percentual de crescimento das vendas em comparação com o ano anterior (somente Jan a Set de cada ano). \
                Valores zerados -> significa que as vendas ano passado e este ano foram iguais. \
                Valores inf -> significa que ano passado não houve vendas ou o produto não existia."
        ),
        "Jan-Set (Qtd)": st.column_config.NumberColumn(
            help=f"Mostra a quantidade vendida de Janeiro a Setembro de {ano_corrente}."
        ),
        "Jan-Set (€)": st.column_config.NumberColumn(
            help=f"Mostra o valor em vendas de Janeiro a Setembro de {ano_corrente}."
        ),
        "Prev. Out-Dez (Qtd)": st.column_config.NumberColumn(
            help=f"Mostra a previsão da quantidade vendida de Outubro a Dezembro de {ano_corrente}, multiplicando-se \
                a média mensal da quantidade vendida de Janeiro a Setembro por 2,5."
        ),
        "Prev. Out-Dez (€)": st.column_config.NumberColumn(
            help=f"Mostra a previsão do valor em vendas de Outubro a Dezembro de {ano_corrente}, multiplicando-se \
                a média mensal das vendas de Janeiro a Setembro por 2,5."
        ),
        "Total Ano (Qtd)": st.column_config.NumberColumn(
            help="Mostra a soma de 'Jan-Set (Qtd)' com 'Prev. Out-Dez (Qtd)'."
        ),
        "Total Ano (€)": st.column_config.NumberColumn(
            help="Mostra a soma de 'Jan-Set (€)' com 'Prev. Out-Dez (€)'."
        ),
        "Qtd próx. ano": st.column_config.NumberColumn(
            help=f"Mostra a quantidade de vendas projetada para o ano de {proximo_ano}, após aplicar a \
            'Meta (%)' sobre 'Total Ano (Qtd)'."
        ),
        "Meta (%)": st.column_config.NumberColumn(
            help=f"Mostra o percentual de crescimento/encolhimento deliberado para o ano de {proximo_ano}."
        ),
        "Vendas próx. ano (€)": st.column_config.NumberColumn(
            help=f"Mostra os valores em vendas projetados para o ano de {proximo_ano}, após aplicar a \
            'Meta (%)' sobre 'Total Ano (€)'."
        ),
        "Dif. (Qtd)": st.column_config.NumberColumn(
            help=f"Mostra a quantidade de vendas acrescida/reduzida para {proximo_ano} em \
                relação a {ano_corrente}."
        ),
        "Dif. (€)": st.column_config.NumberColumn(
            help=f"Mostra os valores em vendas acrescidos/reduzidos para {proximo_ano} em \
                relação a {ano_corrente}."
        )
    },
    use_container_width=True,
    hide_index=True
)

############################### Botão para Download - Dataframe Budget ##############################
# Criando a função para converter o dataframe para excel.


def convert_df_to_excel(df):
    # Create a BytesIO buffer to save the Excel file
    excel_buffer = io.BytesIO()

    # Use Pandas to write the DataFrame to the buffer as an Excel file
    with pd.ExcelWriter(excel_buffer, engine='openpyxl', mode='xlsx') as writer:
        df.to_excel(writer, index=False)

    return excel_buffer


# Aplicando a função
excel_budget = convert_df_to_excel(df9_design)

# Criando o botão de download

colunas_1 = st.columns(6)

colunas_1[4].download_button(
    label="Descarregar Budget",
    data=excel_budget,
    file_name='Metas_2024.xlsx',
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
)

############################### Botão para Resetar o Dataframe Budget ##############################

# Criando a função para deletar o arquivo que é salvo após inserir ao menos uma meta no formulário


def delete_file(path):
    if os.path.exists(path):
        os.remove(path)


file_path = "Couto_metas_2024.csv"

# Para a criação do botão Limpar Metas, é preciso usar o session state, pois aninhar botões no streamlit
# exige vencer as limitações causadas pelo rerun automático.
# Então, eu inicio o clicked na sessão.
# O botão 1 é usado para a exclusão, o botão 2 é usado para confirmar a exclusão,
# e o botão 3 é usado para cancelar a operação. Os 3 foram definidos como falso.
if 'clicked' not in st.session_state:
    st.session_state.clicked = {1: False, 2: False, 3: False}

# Crio a função que recebe o número do botão pressionado. Essa função será chamada
# quando um botão for clicado, mudando o estado da sessão dele para True.


def clicked(button):
    st.session_state.clicked[button] = True


# Criação do botão 'Limpar Metas'. Quando o botão é clicado, ele chama a função clicked
# com o argumento 1, que corresponde a esse botão, transformando o botão em True no estado de sessão.
colunas_1[5].button('Limpar Metas', on_click=clicked, args=[1])

# Portanto, quando clicar no botão 1, ele vai virar True, e aí esse primeiro conditional abaixo vai começar a rodar.
if st.session_state.clicked[1]:
    # Então, se o botão for clicado, vai aparecer uma mensagem de alerta. Em seguida,
    # 2 botões são criados e apresentados (Sim e Não, de forma que se o Sim for clicado, seu
    # estado de sessão vira True. O mesmo vale pro Não.)
    col5, col6, col7 = st.columns([0.4, 0.2, 0.4])
    col6.error(
        "Tem certeza que deseja limpar as metas já incluídas na tabela acima?")
    colunas_2 = st.columns(14)
    colunas_2[6].button('Sim', on_click=clicked, args=[2])
    colunas_2[7].button('Não', on_click=clicked, args=[3])
    # Após clicar no botão de Sim ou Não, o streamlit faz o rerun, e desde lá de cima, ele entende que
    # o session state do botão 1 está true e já adentra no primeiro conditional e depois adentra no
    # segundo ou terceiro conditional, pois o botão 2 ou 3 estará como True.
    # Se entrar no do botão 2, ele vai deletar o arquivo, vai transformar todos os botões como Falso
    # no session state e depois vai fazer um rerun na página, para sumir os botões, já que
    # quando ele passar pelo codigo todo de novo, ele vai se deparar com o primeiro conditional,
    # e como estará falso, ele vai seguir em frente sem criar o alerta e os botões de Sim e Não.
    # Pro caso de ter clicado no Não, ele transforma todos em falso também e depois reinicia a página.
    if st.session_state.clicked[2]:
        delete_file(file_path)
        st.session_state.clicked = {1: False, 2: False, 3: False}
        st.rerun()
    if st.session_state.clicked[3]:
        st.session_state.clicked = {1: False, 2: False, 3: False}
        st.rerun()

# Neste momento, é feito o design dos botões criados até aqui e dos que criarei mais a frente. A parte do stroke se refere
# a um ícone de Help no formulário de Carregar Correções, visto a seguir.
st.markdown("""
    <style>
    .st-emotion-cache-1lc5t6v {
            background-color: rgb(204,223,204);
            width: -webkit-fill-available;
    }
    div.stAlert {
            text-align: -webkit-center;
    }
    .st-emotion-cache-900wwq svg {
            stroke: rgb(255, 255, 255);
    }              
    </style>
""",
            unsafe_allow_html=True
            )

############################################ Montando Dataframe Mensualizado 1 #################################

st.markdown('''---''')
st.markdown(""" **MENSUALIZADO** """)

# Para fazer o mensualizado, vou utilizar como base o do dataframe Budget, deixando somente
# as colunas para identificação do produto e a de quantidade e vendas para o proximo ano, definidas no budget
# e a de preço médio.
df10_base1 = df7_selectedColumns.drop(columns=['Jan-Set (Qtd)', '% N-1 (Qtd)', 'Jan-Set (€)', '% N-1 (€)', 'Prev. Out-Dez (Qtd)',
                                               'Prev. Out-Dez (€)', 'Total Ano (Qtd)', 'Total Ano (€)', 'Meta (%)', 'Dif. (Qtd)', 'Dif. (€)'])

# A segunda base será o segundo argumento do arquivo importado no início do script. Nele, eu tenho para cada
# produto e respectiva subdivisão, a média ponderada de cada mês do ano em sua representação na quantidade
# vendida dos últimos 3 anos.
df11_base2 = df[1].copy()

# Para poder mesclar o 1 com o 2, tenho que criar, para cada produto e subdivisão, 12 meses.
months_data = []
for product_code in df10_base1['Produto'].unique():
    for subdivision in df10_base1['Subdivisão'].unique():
        for month_1 in range(1, 13):
            months_data.append({
                'Produto': product_code,
                'Subdivisão': subdivision,
                'No_mês': month_1,
            })
# Portanto, transformo isso em um dataframe para, em seguida, mesclar com a base 1. Estando apto para depois
# mesclar a base1 com a 2.
df12_monthData = pd.DataFrame(months_data)

df13_monthly = df10_base1.merge(
    df12_monthData, on=['Produto', 'Subdivisão'], how='left')

df13_monthly = df13_monthly.merge(
    df11_base2, on=['Produto', 'Subdivisão', 'No_mês'], how='left')

########################## Sidebar: Formulário Carregar Correção ########################

# Essa pequena interrupção no código é para resolver o problema de ter um produto que foi criado
# no ano atual deste V1, mas não existia nos anos anteriores. Logo, o percentual mensal vai ficar zerado
# para estes. A solução foi criar um botão, chamado Corrigir Zerados, para que, ao ser clicado, seja feito
# um download de uma planilha do excel explicitando os produtos que estão com a soma dos seus
# percentuais mensais do DataFrame Mensualizado zerados. Em seguida, o usuário, sabendo disso, irá preencher os campos
# de 'Jan %' a 'Dez %' com os novos percentuais, salvar e fazer o upload do ficheiro no novo formulário a ser criado
# na sidebar, chamado Carregar Correções. Com isso, o ficheiro carregado será salvo na pasta Corrigir Zerados do servidor
# e alimentará o Mensualizado com as correções feitas.

# Criando a função que vai salvar o ficheiro 'uploaded_file' no diretório correto.


def save_uploaded_file(uploaded_file):

    file_path = os.path.join(
        '', uploaded_file.name)
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())


st.sidebar.markdown('')

# Criando o formulário Carregar Correção. Dentro dele haverá a parte de upload de ficheiro do streamlit e também um botão de enviar.
# O botão de enviar foi uma solução boa para fazer com que um rerun acontecesse na parte de upload, através do argumento clear_on_submit

form2 = st.sidebar.form(
    key='formulario_carregar_correção', clear_on_submit=True)

form2.header("Carregar Correção")

# Essa é a parte de upload do streamlit
uploaded_file = form2.file_uploader(f"Carregue o ficheiro corrigir_zerados_{proximo_ano}.xlsx", type=[
    "xlsx"], help='Ter atenção às instruções do ficheiro ReadMe.txt, descarregado ao clicar em Corrigir Zerados, na parte MENSUALIZADO.')

# Criação do botão para enviar, que vai basicamente efetuar o salvamento do ficheiro na pasta certa.
if form2.form_submit_button('Enviar'):
    save_uploaded_file(uploaded_file)

############################################ Montando Dataframe Mensualizado 2 #################################

# Foi necessário usar o try aqui, pois se ainda não tivesse havido nenhuma correção, haveria um erro.
# Logo, se não houve correção, ele passa em frente sem ler a pasta de Corrigir_zerados.
# Quando ocorre a leitura do ficheiro de correção, é usado um loc, para atualizar o df13_monthly com os
# novos percentuais corrigidos.
try:
    df14_repair = pd.read_excel(
        f'corrigir_zerados_{proximo_ano}.xlsx')
    df14_repair['Produto'] = df14_repair['Produto'].astype(str)

    for index, row in df14_repair.iterrows():
        condition = (df13_monthly['Produto'] == row['Produto']) & (
            df13_monthly['Subdivisão'] == row['Subdivisão'])
        for month_2 in range(1, 13):
            month_name_1 = calendar.month_abbr[month_2]
            df13_monthly.loc[(condition) & (
                df13_monthly['No_mês'] == month_2), 'Percentual_final'] = row[f'{month_name_1} %']
except:
    pass

# Nessa parte, eu trato o dataframe para conseguir formar o ficheiro que será exportado através do botão Corrigir Zerados.
# Com isso, um download é feito de uma tabela com gama, descrições, código e subdivisão dos produtos que se encontram
# com o Percentual_final, de jan a dez, zerados, indicando quais os prováveis produtos que deverão ter seus percentuais
# corrigidos, seja porque houve a criação de um produto totalmente novo ou mudança no código do produto antigo, fazendo
# com que não haja nenhum histórico de vendas deste produto etc. Foi criado logo esse df15_repair aqui, pois antes de prosseguirmos com o
# df16_design, eu já tenho que preparar o arquivo pra ser descarregado no botão Corrigir Zerados.

df15_repair = df13_monthly.groupby(
    ['Gama', 'Descrição', 'Descrição detalhada', 'Produto', 'Subdivisão']).agg({'Percentual_final': sum}).reset_index()
df15_repair = df15_repair[df15_repair['Percentual_final'] == 0]
df15_repair = df15_repair.drop(columns=['Percentual_final'])
for month_3 in range(1, 13):
    month_name_2 = calendar.month_abbr[month_3]
    df15_repair[f'{month_name_2} %'] = 0


# Agora, após a interrupção, vou seguir trabalhando com o df13_monthly, criando 3 colunas para cada mês do ano.
# A primeira receberá o percentual trazido pelo df[1] para cada mês,
# para cada produto e sua respectiva subdivisão. A segunda, será multiplicar a primeira (percentual) pela coluna
# quantidade do proximo ano. A terceira, será multiplicar a primeira pela coluna de vendas do proximo ano.
# OBS: obviamente foi necessário colocar a parte de correções antes desse trecho, pois, se houver correções,
# quero que também apareça aqui.

for month_4 in range(1, 13):
    # Get the abbreviated month name (e.g., 'Jan')
    month_name_3 = calendar.month_abbr[month_4]

    # Create new columns for percentage, quantity, and sales value for each month
    df13_monthly[f'{month_name_3} %'] = df13_monthly.loc[df13_monthly['No_mês']
                                                         == month_4, 'Percentual_final']
    df13_monthly[f'{month_name_3} Qtd'] = round(
        (df13_monthly[f'{month_name_3} %'] * df13_monthly['Qtd próx. ano'])/100)
    df13_monthly[f'{month_name_3} €'] = round(
        (df13_monthly[f'{month_name_3} %'] * df13_monthly['Vendas próx. ano (€)']/100), 2)


# Com o dataframe mensualizado pronto, vou aparar as arestas retirando as colunas que não preciso mais e
# depois vou agrupar e somar elas, colocando na mesma linha do respectivo produto e subdivisão, as 36 colunas criadas
df13_monthly = df13_monthly.drop(['No_mês', 'Percentual_final'], axis=1)

df13_monthly = df13_monthly.groupby(['Área', 'Gama', 'Descrição', 'Descrição detalhada', 'Produto',
                                     'Subdivisão', 'Qtd próx. ano', 'Vendas próx. ano (€)', 'Preço médio (€)'])[['Jan %', 'Jan Qtd', 'Jan €',
                                                                                                                 'Feb %', 'Feb Qtd', 'Feb €',
                                                                                                                 'Mar %', 'Mar Qtd', 'Mar €',
                                                                                                                 'Apr %', 'Apr Qtd', 'Apr €',
                                                                                                                 'May %', 'May Qtd', 'May €',
                                                                                                                 'Jun %', 'Jun Qtd', 'Jun €',
                                                                                                                 'Jul %', 'Jul Qtd', 'Jul €',
                                                                                                                 'Aug %', 'Aug Qtd', 'Aug €',
                                                                                                                 'Sep %', 'Sep Qtd', 'Sep €',
                                                                                                                 'Oct %', 'Oct Qtd', 'Oct €',
                                                                                                                 'Nov %', 'Nov Qtd', 'Nov €',
                                                                                                                 'Dec %', 'Dec Qtd', 'Dec €']].sum().reset_index()

# Nesse momento, eu organizo alfabeticamente o dataframe, exceto a subdivisão, que o CASA deve vir antes do CLIN
df13_monthly = df13_monthly.sort_values(
    by=['Gama', 'Descrição', 'Descrição detalhada', 'Subdivisão'], ascending=[True, True, True, False])


############################# Design do Dataframe Mensualizado e Postando #############################

# O df13_monthly já está pronto, mas agora vou trabalhar no design. Como há muitas colunas antes de iniciar
# a sequência das informações mensais (3 colunas pra cada mês), fica complicado para o usuário saber a qual produto
# se refere aquela linha se ele já estiver lá em Julho. Logo, o ideal é criar um índice.

# Porém, o índice ficaria enorme. Então, vou retirar a área e gama. Depois vou concatenar as outras colunas
# em uma só e transformar ela em índice. Depois, deletarei as colunas usadas para a concatenação.
df16_design = df13_monthly.drop(['Área', 'Gama'], axis=1)
df16_design['Descrição do Produto'] = df16_design['Descrição'] + ' (' + df16_design['Descrição detalhada'] + \
    ') - ' + df16_design['Subdivisão'] + ' - ' + df16_design['Produto']
df16_design = df16_design.set_index(['Descrição do Produto'])
df16_design = df16_design.drop(
    ['Descrição', 'Descrição detalhada', 'Produto', 'Subdivisão'], axis=1)

# Aqui, vou fazer um design diferente do Budget, pois vou previlegiar a separação visual de colunas, não
# linhas. Vou colocar as mesmas cores para cada 3 colunas referentes ao mesmo mês. À medida em que eu mudar o
# mês, vou intercalando as cores.


def color_background(df):
    bg_colors = []
    color = 'background-color: #ffffff'
    for col in df.index:
        if '%' in col:
            color = 'background-color: #f2f7f2' if color == 'background-color: #ffffff' else 'background-color: #ffffff'
        bg_colors.append(color)
    return bg_colors


# Neste momento vou aplicar a função style, para inserir a formatação da função criada acima e depois vou estabelecer
# os formatos das colunas numéricas.
df17_styled = df16_design.style.apply(color_background, axis=1)

df17_styled = df17_styled.format({'Qtd próx. ano': "{:,.0f}".format,
                                  'Vendas próx. ano (€)': "{:,.2f}".format,
                                  'Preço médio (€)': "{:,.2f}".format,
                                  'Jan %': "{:,.1f}".format, 'Jan Qtd': "{:,.0f}".format, 'Jan €': "{:,.2f}".format,
                                  'Feb %': "{:,.1f}".format, 'Feb Qtd': "{:,.0f}".format, 'Feb €': "{:,.2f}".format,
                                  'Mar %': "{:,.1f}".format, 'Mar Qtd': "{:,.0f}".format, 'Mar €': "{:,.2f}".format,
                                  'Apr %': "{:,.1f}".format, 'Apr Qtd': "{:,.0f}".format, 'Apr €': "{:,.2f}".format,
                                  'May %': "{:,.1f}".format, 'May Qtd': "{:,.0f}".format, 'May €': "{:,.2f}".format,
                                  'Jun %': "{:,.1f}".format, 'Jun Qtd': "{:,.0f}".format, 'Jun €': "{:,.2f}".format,
                                  'Jul %': "{:,.1f}".format, 'Jul Qtd': "{:,.0f}".format, 'Jul €': "{:,.2f}".format,
                                  'Aug %': "{:,.1f}".format, 'Aug Qtd': "{:,.0f}".format, 'Aug €': "{:,.2f}".format,
                                  'Sep %': "{:,.1f}".format, 'Sep Qtd': "{:,.0f}".format, 'Sep €': "{:,.2f}".format,
                                  'Oct %': "{:,.1f}".format, 'Oct Qtd': "{:,.0f}".format, 'Oct €': "{:,.2f}".format,
                                  'Nov %': "{:,.1f}".format, 'Nov Qtd': "{:,.0f}".format, 'Nov €': "{:,.2f}".format,
                                  'Dec %': "{:,.1f}".format, 'Dec Qtd': "{:,.0f}".format, 'Dec €': "{:,.2f}".format,
                                  })

# Postando o DataFrame.
st.dataframe(df17_styled,
             column_config={
                 "Qtd próx. ano": st.column_config.NumberColumn(
                     help=f"Mostra a quantidade de vendas projetada para o ano de {proximo_ano}."
                 ),
                 "Vendas próx. ano (€)": st.column_config.NumberColumn(
                     help=f"Mostra os valores em vendas projetados para o ano de {proximo_ano}."
                 ),
                 "Preço médio (€)": st.column_config.NumberColumn(
                     help="Valor resultante da divisão entre 'Vendas próx. ano (€)' e 'Qtd próx. ano'."
                 ),
                 "Jan %": st.column_config.NumberColumn(
                     help=f"Mostra o percentual médio, dos últimos 3 anos, que Janeiro representou frente às vendas do ano inteiro. \
                Sendo a média ponderada com peso 3 para ano passado, peso 2 para o retrasado e 1 para o ano mais distante"
                 ),
                 "Jan Qtd": st.column_config.NumberColumn(
                     help=f"Valor resultante da multiplicação de 'Jan %' pela 'Qtd próx. ano'."
                 ),
                 "Jan €": st.column_config.NumberColumn(
                     help=f"Valor resultante da multiplicação de 'Jan %' pelas 'Vendas próx. ano (€)'."
                 )
             },
             use_container_width=True
             )

############################### Botões (Download e Corrigir Zerados) - Dataframe Mensualizado ##############################

# Como forma de possibilitar um download sem esse índice concatenado e sem ter dropado as colunas
# que identificavam o produto, eu vou reaproveitar o df13 para aplicar a função da colorbackground

df18_styled = df13_monthly.style.apply(color_background, axis=1)
df18_styled = df18_styled.format({'Qtd próx. ano': "{:,.0f}".format,
                                  'Vendas próx. ano (€)': "{:,.2f}".format,
                                  'Preço médio (€)': "{:,.2f}".format,
                                  'Jan %': "{:,.1f}".format, 'Jan Qtd': "{:,.0f}".format, 'Jan €': "{:,.2f}".format,
                                  'Feb %': "{:,.1f}".format, 'Feb Qtd': "{:,.0f}".format, 'Feb €': "{:,.2f}".format,
                                  'Mar %': "{:,.1f}".format, 'Mar Qtd': "{:,.0f}".format, 'Mar €': "{:,.2f}".format,
                                  'Apr %': "{:,.1f}".format, 'Apr Qtd': "{:,.0f}".format, 'Apr €': "{:,.2f}".format,
                                  'May %': "{:,.1f}".format, 'May Qtd': "{:,.0f}".format, 'May €': "{:,.2f}".format,
                                  'Jun %': "{:,.1f}".format, 'Jun Qtd': "{:,.0f}".format, 'Jun €': "{:,.2f}".format,
                                  'Jul %': "{:,.1f}".format, 'Jul Qtd': "{:,.0f}".format, 'Jul €': "{:,.2f}".format,
                                  'Aug %': "{:,.1f}".format, 'Aug Qtd': "{:,.0f}".format, 'Aug €': "{:,.2f}".format,
                                  'Sep %': "{:,.1f}".format, 'Sep Qtd': "{:,.0f}".format, 'Sep €': "{:,.2f}".format,
                                  'Oct %': "{:,.1f}".format, 'Oct Qtd': "{:,.0f}".format, 'Oct €': "{:,.2f}".format,
                                  'Nov %': "{:,.1f}".format, 'Nov Qtd': "{:,.0f}".format, 'Nov €': "{:,.2f}".format,
                                  'Dec %': "{:,.1f}".format, 'Dec Qtd': "{:,.0f}".format, 'Dec €': "{:,.2f}".format,
                                  })

# Para a criação do botão de download, antes se aplica a função que converte dataframe em excel
excel_mensualizado = convert_df_to_excel(df18_styled)

colunas_2 = st.columns(4)

colunas_2[3].download_button(
    label="Descarregar Mensualizado",
    data=excel_mensualizado,
    file_name='Metas_2024_Mensualizado.xlsx',
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
)

# Criação do botão de Corrigir Zerados. Nesta parte, não somente o excel_zerados será descarregado,
# mas também um readme.txt que criei com instruções para que o usuário faça corretamente o upload
# das correções. Logo, como são dois arquivos, o streamlit indica que eles sejam zipados, o que vai
# facilitar o descarregamento, pois será basicamente 1 arquivo só.

# Arquivo em excel:
excel_zerados = convert_df_to_excel(df15_repair)

# Arquivo em txt:
txt_file_path = 'Readme.txt'
with open(txt_file_path, 'r') as file:
    readme_file = file.read()

# Aqui é que crio a conversão em zip file.
zip_buffer = BytesIO()

with zipfile.ZipFile(zip_buffer, 'a', zipfile.ZIP_DEFLATED, False) as zip_file:
    # Add Excel file to the zip archive
    zip_file.writestr(
        f'corrigir_zerados_{proximo_ano}.xlsx', excel_zerados.getvalue())

    # Add readme file to the zip archive
    zip_file.writestr('Readme.txt', readme_file)

zip_buffer.seek(0)

# Com o zip file salvo, vou criar o botão para seu download.
colunas_2[2].download_button(
    label="Corrigir Zerados",
    data=zip_buffer,
    file_name=f'corrigir_zerados_{proximo_ano}.zip',
    mime='application/zip'
)

############################################ Resumo - Montando e postando as métricas 1 #################################

st.markdown('''---''')
st.markdown(""" **RESUMO** """)

col8, col9, col10, col11, col12 = st.columns(
    [0.1, 0.2, 0.2, 0.1, 0.1])

col13, col14, col15, col16, col17 = st.columns(
    [0.1, 0.2, 0.2, 0.1, 0.1])

# montando os valores nas variáveis
total_ano_corrente_1 = df7_selectedColumns['Total Ano (Qtd)'].sum()
total_proximo_ano_1 = df7_selectedColumns['Qtd próx. ano'].sum()
df19_subtotal = df7_selectedColumns[['Subdivisão', 'Total Ano (Qtd)', 'Qtd próx. ano']].groupby(
    'Subdivisão').sum().reset_index()

total_ano_corrente_2 = df7_selectedColumns['Total Ano (€)'].sum()
total_proximo_ano_2 = df7_selectedColumns['Vendas próx. ano (€)'].sum()
df20_subtotal = df7_selectedColumns[['Subdivisão', 'Total Ano (€)', 'Vendas próx. ano (€)']].groupby(
    'Subdivisão').sum().reset_index()

# Para montar o restante dos valores nas variáveis (referente ao ano anterior, agora), preciso adiantar uns steps que iria mostrar na parte 2
# deste resumo: Para montar o Dataframe do resumo , vou usar 2 dataframes base, um para mostrar as somas por área e subdivisão (df24_area_subdiv_final)
# e outro para mostrar as somas por área, que irá servir como Subtotal por área (df26_area_final).
# Neste momento do código (parte 1 das métricas do resumo), vou iniciar a formulação desse primeiro dataframe (df24_area_subdiv_final)

# Primeiro, estou pegando a coluna Total Ano (qtd) e (€), pois eu sei que ela pega de janeiro a setembro do ano corrente e mais
# 2,5x a média mensal de janeiro a setembro. Portanto, eu tenho a ideia do que será o ano corrente em termos de vendas.
# Estou mantendo a coluna de produtos ainda, pois vou precisar mesclar com o segundo dataframe, referente ao ano anterior.
# Isso me permitirá pegar somente a soma do ano anterior nos produtos que também estiveram presentes no ano corrente.
# (atentar que tudo que estou fazendo nesse momento é visando o dataframe da parte 2 do resumo, então, pode parecer que
# estou coletando colunas desnecessárias). As partes estritamente usadas aqui eu vou falando durante este snippet

df21_area_subdiv_corrente = df7_selectedColumns.groupby(['Área', 'Produto', 'Subdivisão'])[
    ['Total Ano (Qtd)', 'Total Ano (€)', 'Qtd próx. ano', 'Vendas próx. ano (€)', 'Dif. (Qtd)', 'Dif. (€)']].sum().reset_index()

# Pegando os dados lá do df3, no início, para, dessa vez, não segregar de janeiro a setembro, mas pegar tudo do ano anterior (até dezembro).
df22_area_subdiv_anterior = df3_anterior.groupby(['Área_negócio', 'Produto', 'Subdivisão_agrupada'])[
    ['Quantidade', 'Vendas']].sum().reset_index()

# Corrigindo os nomes para facilitar o merge.
df22_area_subdiv_anterior.rename(columns={'Quantidade': 'Total Ano (Qtd)_anterior', 'Vendas': 'Total Ano (€)_anterior',
                                          'Área_negócio': 'Área', 'Subdivisão_agrupada': 'Subdivisão'}, inplace=True)

# Fazendo o merge.
df23_area_subdiv = df21_area_subdiv_corrente.merge(df22_area_subdiv_anterior, on=['Área', 'Produto', 'Subdivisão'],
                                                   how='left')

# A partir desse momento, vou começar a focar nas métricas dessa parte 1, deixando o df23_area_subdiv para ser mais manipulado na parte 2 das métricas.

# Eu consigo puxar as métricas referentes ao ano anterior e dar prosseguimento à montagem da métrica
total_ano_anterior_1 = df23_area_subdiv['Total Ano (Qtd)_anterior'].sum()
total_ano_anterior_2 = df23_area_subdiv['Total Ano (€)_anterior'].sum()


# construindo os gráficos de pizza de quantidade e vendas
#fig1 = go.Figure(data=[go.Pie(labels=df19_subtotal['Subdivisão'],
#                              values=df19_subtotal['Qtd próx. ano'])])
#fig1.update_traces(hoverinfo='value', textinfo='label+percent', textfont_size=12,
#                  marker=dict(colors=['#a1c9a4', '#066555'], line=dict(color='#000000', width=0.5)))
#fig1.update_layout(showlegend=False,
#                   width=200,
#                   height=150,
#                   margin=dict(l=20, r=20, b=20, t=30),
#                   paper_bgcolor='#f5f5f5',
#                   font=dict(color='#31333F', size=15)
#                   )

# fig2 = go.Figure(data=[go.Pie(labels=df20_subtotal['Subdivisão'],
#                              values=df20_subtotal['Vendas próx. ano (€)'])])
# fig2.update_traces(hoverinfo='value', textinfo='label+percent', textfont_size=12,
#                   marker=dict(colors=['#a1c9a4', '#066555'], line=dict(color='#000000', width=0.5)))
# fig2.update_layout(showlegend=False,
#                   width=200,
#                   height=150,
#                   margin=dict(l=20, r=20, b=20, t=30),
#                   paper_bgcolor='#f5f5f5',
#                   font=dict(color='#31333F', size=15)
#                   )

# montando e postando as métricas relativas à quantidade
variacao_rel_1 = (
    (total_ano_corrente_1 - total_ano_anterior_1)/total_ano_anterior_1)*100
variacao_abs_1 = (total_ano_corrente_1 - total_ano_anterior_1)
variacao_rel_2 = ((total_proximo_ano_1 - total_ano_corrente_1) /
                  total_ano_corrente_1)*100
variacao_abs_2 = (total_proximo_ano_1 - total_ano_corrente_1)

total_ano_corrente_1 = '{:,.0f}'.format(total_ano_corrente_1)
total_proximo_ano_1 = '{:,.0f}'.format(total_proximo_ano_1)
variacao_rel_1 = '{:.0f}'.format(variacao_rel_1)
variacao_abs_1 = '{:,.0f}'.format(variacao_abs_1)
variacao_rel_2 = '{:.0f}'.format(variacao_rel_2)
variacao_abs_2 = '{:,.0f}'.format(variacao_abs_2)

col9.metric(f'Previsão {ano_corrente} em Quantidade', total_ano_corrente_1,
            delta=f'{variacao_abs_1} ({variacao_rel_1} %)')
col10.metric(f'Meta {proximo_ano} em Quantidade', total_proximo_ano_1,
             delta=f'{variacao_abs_2} ({variacao_rel_2} %)')
# col11.write(fig1)

# montando e postando as métricas relativas à venda
variacao_rel_3 = (
    (total_ano_corrente_2 - total_ano_anterior_2)/total_ano_anterior_2)*100
variacao_abs_3 = (total_ano_corrente_2 - total_ano_anterior_2)
variacao_rel_4 = ((total_proximo_ano_2 - total_ano_corrente_2) /
                  total_ano_corrente_2)*100
variacao_abs_4 = (total_proximo_ano_2 - total_ano_corrente_2)

total_ano_corrente_2 = '€ {:,.2f}'.format(total_ano_corrente_2)
total_proximo_ano_2 = '€ {:,.2f}'.format(total_proximo_ano_2)
variacao_rel_3 = '{:.0f}'.format(variacao_rel_3)
variacao_abs_3 = '{:,.2f}'.format(variacao_abs_3)
variacao_rel_4 = '{:.0f}'.format(variacao_rel_4)
variacao_abs_4 = '{:,.2f}'.format(variacao_abs_4)


col14.metric(f'Previsão {ano_corrente} em Vendas', total_ano_corrente_2,
             delta=f'{variacao_abs_3} ({variacao_rel_3} %)')
col15.metric(f'Meta {proximo_ano} em Vendas', total_proximo_ano_2,
             delta=f'{variacao_abs_4} ({variacao_rel_4} %)')
# col16.write(fig2)

#################################### Resumo - Montando e postando as métricas 2 #################################

# Continuando o que comecei na parte 1, com relação à montagem do dataframe resumo, vou terminar de manipular o df23_area_subdiv,
# criando, depois, as colunas necessárias. E depois de finalizar esse primeiro dataframe (df24_area_subdiv_final), vou iniciar aquele
# segundo que havia falado (voltado só pra área).

# Agora, fazendo o agrupamento pelas colunas que me interessam (área e subdivisão), pois é disso que a tabela de resumo necessita.
# Portanto, somei todas as outras colunas de números.
df23_area_subdiv = df23_area_subdiv.groupby(['Área', 'Subdivisão'])[['Total Ano (Qtd)', 'Total Ano (€)',
                                                                     'Total Ano (Qtd)_anterior', 'Total Ano (€)_anterior',
                                                                     'Qtd próx. ano', 'Vendas próx. ano (€)', 'Dif. (Qtd)', 'Dif. (€)']].sum().reset_index()

# Nesse momento, eu estou criando mais 4 colunas que me interessam. As duas primeiras são comparações do total previsto do ano corrente com
# o total efetivamente ocorrido no ano anterior. Daí, eu acho o crescimento/encolhimento das vendas.
# As duas últimas se referem ao percentual de crescimento/encolhimento para o próximo ano em relação ao ano corrente.
df23_area_subdiv['% N-1 (Qtd)'] = round(((df23_area_subdiv['Total Ano (Qtd)'] -
                                          df23_area_subdiv['Total Ano (Qtd)_anterior']) / df23_area_subdiv['Total Ano (Qtd)_anterior']) * 100, 1)
df23_area_subdiv['% N-1 (€)'] = round(((df23_area_subdiv['Total Ano (€)'] -
                                        df23_area_subdiv['Total Ano (€)_anterior']) / df23_area_subdiv['Total Ano (€)_anterior']) * 100, 1)
df23_area_subdiv['% (Qtd)'] = round(
    (df23_area_subdiv['Dif. (Qtd)']/df23_area_subdiv['Total Ano (Qtd)'])*100, 1)

df23_area_subdiv['% (€)'] = round(
    (df23_area_subdiv['Dif. (€)']/df23_area_subdiv['Total Ano (€)'])*100, 1)

# Agora já posso dropar o que não preciso, partindo para finalizar o primeiro dataframe que falei lá no início. Troquei o nome do dataframe,
# pois, o df23_area_subdiv ainda terá suas colunas usadas para trabalhar no segundo dataframe.
df24_area_subdiv_final = df23_area_subdiv.drop(
    columns=['Total Ano (Qtd)_anterior', 'Total Ano (€)_anterior'], axis=1)

df24_area_subdiv_final = df24_area_subdiv_final.sort_values(
    by=['Área', 'Subdivisão'], ascending=[True, False])

# Reta final: ajustando a ordem das colunas.
df24_area_subdiv_final = df24_area_subdiv_final[['Área', 'Subdivisão', 'Total Ano (Qtd)', '% N-1 (Qtd)',
                                                 'Total Ano (€)', '% N-1 (€)', 'Qtd próx. ano', 'Vendas próx. ano (€)',
                                                 'Dif. (Qtd)', '% (Qtd)', 'Dif. (€)', '% (€)']]

# Aqui inicio o segundo dataframe. Nele, vou me preocupar em agrupar somente por área, de modo que eu tenha um dataframe só com os subtotais por área.
# Ele irá servir para mesclar com o dataframe finalizado acima.
# Atentar que vou seguir o mesmo padrão acima, porém, sem mais agrupar por subdivisão
df25_area = df23_area_subdiv.groupby('Área')[['Total Ano (Qtd)', '% N-1 (Qtd)', 'Total Ano (€)', '% N-1 (€)',
                                              'Total Ano (Qtd)_anterior', 'Total Ano (€)_anterior',
                                              'Qtd próx. ano', 'Vendas próx. ano (€)', 'Dif. (Qtd)', '% (Qtd)', 'Dif. (€)', '% (€)']].sum().reset_index()

# aqui crio as 4 colunas novamente. O intuito é ter 2 dataframes com as mesmas colunas, para mesclar facilmente.
df25_area['% N-1 (Qtd)'] = round(((df25_area['Total Ano (Qtd)'] -
                                   df25_area['Total Ano (Qtd)_anterior']) / df25_area['Total Ano (Qtd)_anterior']) * 100, 1)

df25_area['% N-1 (€)'] = round(((df25_area['Total Ano (€)'] -
                                 df25_area['Total Ano (€)_anterior']) / df25_area['Total Ano (€)_anterior']) * 100, 1)

df25_area['% (Qtd)'] = round(
    (df25_area['Dif. (Qtd)']/df25_area['Total Ano (Qtd)'])*100, 1)

df25_area['% (€)'] = round(
    (df25_area['Dif. (€)']/df25_area['Total Ano (€)'])*100, 1)

# Aqui eu nomeio todos os campos None da coluna Subdivisão com a palavra Subtotal.
df25_area['Subdivisão'] = 'Subtotal'

# Finalizando o segundo dataframe.
df26_area_final = df25_area.drop(
    columns=['Total Ano (Qtd)_anterior', 'Total Ano (€)_anterior'], axis=1)

df26_area_final = df26_area_final[['Área', 'Subdivisão', 'Total Ano (Qtd)', '% N-1 (Qtd)',
                                   'Total Ano (€)', '% N-1 (€)', 'Qtd próx. ano', 'Vendas próx. ano (€)',
                                   'Dif. (Qtd)', '% (Qtd)', 'Dif. (€)', '% (€)']]


# Aqui inicio a concatenação entre esses 2 novos dataframes. Só que não será uma concatenação colocando o segundo dataframe
# embaixo do primeiro, mas sim colocando o subtotal por área embaixo imediatamente da última linha daquela área.

last_group = None

# Inicializando um DataFrame vazio para armazenar o resultado final
df27_selectedColumns = pd.DataFrame(columns=df24_area_subdiv_final.columns)

# Iterando pelas linhas do primeiro dataframe
for index, row in df24_area_subdiv_final.iterrows():
    current_group = row['Área']

    # Se o grupo atual for diferente do último, é hora de adicionar o subtotal
    if current_group != last_group:
        if last_group:
            # Adicionando o subtotal do grupo anterior (df2) ao resultado
            df27_selectedColumns = pd.concat(
                [df27_selectedColumns, df26_area_final[df26_area_final['Área'] == last_group]], ignore_index=True)
        last_group = current_group

    # Adicionando a linha atual ao resultado
    df27_selectedColumns = pd.concat(
        [df27_selectedColumns, df24_area_subdiv_final.loc[[index]]], ignore_index=True)

# Adicionando o último subtotal
df27_selectedColumns = pd.concat(
    [df27_selectedColumns, df26_area_final[df26_area_final['Área'] == last_group]], ignore_index=True)

# Agora vou colocar as cores nas linhas, identificando somente as que possuem subtotal por azul, restando branco para as outras linhas.


def highlight_subtotals(rows):
    styles = []
    for _ in rows:
        style = 'background-color: #f2f7f2' if rows['Subdivisão'] == 'Subtotal' else 'background-color: #ffffff'
        styles.append(style)
    return styles

# Aplicando a função de estilo às linhas do DataFrame
df28_styled = df27_selectedColumns.style.apply(highlight_subtotals, axis=1)

# Configurando os formatos das casas decimas e do separador de milhares
df28_styled = df28_styled.format({'Total Ano (Qtd)': "{:,.0f}".format,
                                  '% N-1 (Qtd)': "{:,.1f}".format,
                                  'Qtd próx. ano': "{:,.0f}".format,
                                  'Dif. (Qtd)': "{:,.0f}".format,
                                  '% (Qtd)': "{:,.1f}".format,
                                  'Total Ano (€)': "{:,.2f}".format,
                                  '% N-1 (€)': "{:,.1f}".format,
                                  'Vendas próx. ano (€)': "{:,.2f}".format,
                                  'Dif. (€)': "{:,.2f}".format,
                                  '% (€)': "{:,.1f}".format
                                  })

# Aplicando a formatação de cores dos números nas colunas específicas
df28_styled = df28_styled.applymap(
    color_negative_red, subset=['% N-1 (Qtd)', '% N-1 (€)', '% (Qtd)', '% (€)'])

# Postando o dataframe e colocando alguns balões de informação para explicar, ao usuário, os cabeçalhos.
st.dataframe(
    df28_styled,
    column_config={
        "Total Ano (Qtd)": st.column_config.NumberColumn(
            help=f"Mostra a previsão para a quantidade vendida de Janeiro a Dezembro de {ano_corrente}."
        ),
        "% N-1 (Qtd)": st.column_config.NumberColumn(
            help="Mostra o percentual de crescimento da quantidade vendida em comparação com o ano anterior \
                  (diferente da tabela Budget, esta é uma comparação de Jan a Dez de cada ano)."
        ),
        "Total Ano (€)": st.column_config.NumberColumn(
            help=f"Mostra a previsão para as vendas de Janeiro a Dezembro de {ano_corrente}."
        ),
        "% N-1 (€)": st.column_config.NumberColumn(
            help="Mostra o percentual de crescimento das vendas em comparação com o ano anterior \
                (diferente da tabela Budget, esta é uma comparação de Jan a Dez de cada ano)."
        ),
        "Qtd próx. ano": st.column_config.NumberColumn(
            help=f"Mostra a quantidade de vendas projetada para o ano de {proximo_ano}."
        ),
        "Vendas próx. ano (€)": st.column_config.NumberColumn(
            help=f"Mostra os valores em vendas projetados para o ano de {proximo_ano}."
        ),
        "Dif. (Qtd)": st.column_config.NumberColumn(
            help=f"Mostra a quantidade de vendas acrescida/reduzida para {proximo_ano} em \
                relação a {ano_corrente}."
        ),
        "% (Qtd)": st.column_config.NumberColumn(
            help=f"Mostra o percentual de crescimento/redução da quantidade de vendas para {proximo_ano} em \
                relação a {ano_corrente}."
        ),
        "Dif. (€)": st.column_config.NumberColumn(
            help=f"Mostra os valores em vendas acrescidos/reduzidos para {proximo_ano} em \
                relação a {ano_corrente}."
        ),
        "% (€)": st.column_config.NumberColumn(
            help=f"Mostra o percentual de crescimento/redução dos valores em vendas para {proximo_ano} em \
                relação a {ano_corrente}."
        )
    },
    use_container_width=True,
    hide_index=True
)

############################### Botão para Download - Dataframe Resumo ##############################

excel_resumo = convert_df_to_excel(df28_styled)

col18, col19, col20 = st.columns([0.3, 0.1, 0.1])

col20.download_button(
    label="Descarregar Resumo",
    data=excel_resumo,
    file_name=f'Metas_2024_resumo.xlsx',
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
)

#################################################### Design da página ########################################################

# Retirando a fita colorida na parte de cima do streamlit e o rodapé escrito Made with Streamlit

hide_st_style = """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    </style>
"""
st.markdown(hide_st_style, unsafe_allow_html=True)
