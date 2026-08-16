import io
import json
from datetime import datetime

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import openpyxl
import requests
import unicodedata
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    Image as RLImage,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)


def money_format(valor):
    valor_string = "{:,}".format(valor)

    valor_string = valor_string.replace('.', '_')
    valor_string = valor_string.replace(',', '.')
    valor_string = valor_string.replace('_', ',')

    return valor_string


def normaliza_sigla(valor):
    # Remove acentos, espaços nas pontas e diferenças de maiúsculas/minúsculas
    # para tornar a comparação com a coluna "Sigla" da planilha de tarifas
    # resistente a pequenas diferenças de formatação (ex.: "Enel CE" vs "ENEL CE").
    texto = str(valor).strip().upper()
    texto = unicodedata.normalize('NFKD', texto).encode('ascii', 'ignore').decode('ascii')
    return texto


@st.cache_data
def carregar_planilha_tarifas(excel_file):
    # Cacheado para não reler o Excel do disco a cada chamada de obter_tarifas()
    banco = pd.read_excel(excel_file)
    banco['Sigla_norm'] = banco['Sigla'].apply(normaliza_sigla)
    return banco


# --- Integração com a API de dados abertos da ANEEL (CKAN) ---------------
# Dataset: "Tarifas das distribuidoras de energia elétrica"
# https://dadosabertos.aneel.gov.br/dataset/tarifas-distribuidoras-energia-eletrica/resource/fcf2906c-7c32-4b9b-a637-054e7a5234f4
ANEEL_API_URL = "https://dadosabertos.aneel.gov.br/api/3/action/datastore_search"
ANEEL_RESOURCE_ID = "fcf2906c-7c32-4b9b-a637-054e7a5234f4"


def _para_float_br(valor):
    # Os campos numéricos da API da ANEEL vêm tipados como "text" e podem
    # usar vírgula como separador decimal (padrão BR). Esta função normaliza
    # para float de forma defensiva, sem quebrar se o valor já vier limpo.
    if valor is None:
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).strip()
    if texto == "":
        return 0.0
    if ',' in texto and '.' in texto:
        texto = texto.replace('.', '').replace(',', '.')
    elif ',' in texto:
        texto = texto.replace(',', '.')
    try:
        return float(texto)
    except ValueError:
        return 0.0


# A API pública da ANEEL parece bloquear clientes sem um User-Agent de
# navegador (o mesmo domínio já recusa robôs via robots.txt em /api/), então
# enviamos cabeçalhos que imitam um navegador comum para reduzir a chance de
# um bloqueio silencioso (403) do lado do servidor.
ANEEL_REQUEST_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json",
}


def _buscar_registros_aneel(filtros, timeout=45, limit=32000):
    # Busca (com paginação) todos os registros do datastore_search da ANEEL
    # que casam com `filtros` (dict aplicado como filtro exato server-side).
    params = {
        "resource_id": ANEEL_RESOURCE_ID,
        "filters": json.dumps(filtros),
        "limit": limit,
        "offset": 0,
    }
    registros = []
    while True:
        resposta = requests.get(ANEEL_API_URL, params=params, headers=ANEEL_REQUEST_HEADERS, timeout=timeout)
        resposta.raise_for_status()
        dados = resposta.json()
        if not dados.get("success"):
            raise RuntimeError("a API retornou uma resposta sem sucesso")
        pagina = dados["result"]["records"]
        registros.extend(pagina)
        total = dados["result"].get("total", len(registros))
        params["offset"] += len(pagina)
        if not pagina or params["offset"] >= total:
            break
    return registros


@st.cache_data(ttl=3600, show_spinner=False)
def consultar_tarifas_api_aneel(sigla, grupo='A4'):
    """
    Consulta em tempo real a API de dados abertos da ANEEL (CKAN) e retorna
    um DataFrame no mesmo formato usado pela planilha local (colunas:
    Sigla_norm, Subgrupo, Modalidade, Detalhe, Base Tarifária, Posto,
    Unidade, TUSD, TE), para que a mesma lógica de extração de tarifas
    (_extrair_tarifas) possa ser reaproveitada por ambas as fontes.

    Estratégia em duas etapas para equilibrar velocidade e robustez:
    1) tenta um filtro exato no servidor por SigAgente + Subgrupo (rápido,
       baixo volume de dados);
    2) se isso não retornar nada (ex.: diferença de maiúsculas/acentos entre
       a Sigla usada aqui e o valor exato de SigAgente na API), busca todos
       os registros do Subgrupo e filtra no lado do cliente usando
       normaliza_sigla — a mesma função usada para a planilha local.
    """
    erro_rede = None
    registros = []
    try:
        registros = _buscar_registros_aneel({"SigAgente": sigla, "DscSubGrupo": grupo})
    except requests.exceptions.RequestException as exc:
        erro_rede = exc
    except (ValueError, KeyError):
        registros = []

    if not registros:
        try:
            registros = _buscar_registros_aneel({"DscSubGrupo": grupo})
        except requests.exceptions.RequestException as exc:
            detalhe = str(exc) or str(erro_rede) or "erro desconhecido de conexão"
            raise RuntimeError(f"não foi possível conectar à API de dados abertos da ANEEL ({detalhe})") from exc

    if not registros:
        raise RuntimeError(f"nenhum registro retornado para o Subgrupo {grupo}")

    banco = pd.DataFrame.from_records(registros)
    banco['Sigla_norm'] = banco['SigAgente'].apply(normaliza_sigla)

    sigla_norm_alvo = normaliza_sigla(sigla)
    banco_conc = banco.loc[banco['Sigla_norm'] == sigla_norm_alvo].copy()
    if banco_conc.empty:
        raise RuntimeError(f"nenhum registro encontrado para a concessionária '{sigla}'")

    # A base pode conter múltiplas vigências históricas para a mesma
    # combinação de modalidade/posto/unidade; ficamos apenas com a mais
    # recente (maior DatInicioVigencia).
    banco_conc['DatInicioVigencia'] = pd.to_datetime(banco_conc['DatInicioVigencia'], errors='coerce')
    banco_conc = banco_conc.sort_values('DatInicioVigencia', ascending=False)
    chave = ['DscModalidadeTarifaria', 'DscDetalhe', 'DscBaseTarifaria', 'NomPostoTarifario',
             'DscUnidadeTerciaria']
    banco_conc = banco_conc.drop_duplicates(subset=chave, keep='first')

    banco_conc['TUSD'] = banco_conc['VlrTUSD'].apply(_para_float_br)
    banco_conc['TE'] = banco_conc['VlrTE'].apply(_para_float_br)

    banco_final = banco_conc.rename(columns={
        'DscSubGrupo': 'Subgrupo',
        'DscModalidadeTarifaria': 'Modalidade',
        'DscDetalhe': 'Detalhe',
        'DscBaseTarifaria': 'Base Tarifária',
        'NomPostoTarifario': 'Posto',
        'DscUnidadeTerciaria': 'Unidade',
    })

    colunas = ['Sigla_norm', 'Subgrupo', 'Modalidade', 'Detalhe', 'Base Tarifária', 'Posto', 'Unidade',
               'TUSD', 'TE']
    return banco_final[colunas]


if "disabled" not in st.session_state:
    st.session_state.disabled = False

# Define the months for the first column
months = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro",
          "Outubro", "Novembro", "Dezembro"]

# Define the column headers
columns = ["Meses", "Demanda na ponta", "Demanda Fora da Ponta", "Consumo na Ponta", "Consumo fora da ponta"]

# Create the data for the DataFrame (the first column contains the months, and the rest are sequential numbers)
data = [[months[i]] + [0.00 for j in range(4)] for i in range(12)]

# Create the DataFrame with the specified columns
dados_entrada = pd.DataFrame(data, columns=columns)

# Entrada Inicial de Estados
estados = ["AC", "AL", "AP", "AM", "BA", "CE", "DF", "ES", "GO", "MA", "MT", "MS", "MG", "PA","PB", "PE", "PI", "PR",
           "RJ", "RN", "RO", "RR", "RS", "SC", "SP", "SE", "TO"]


# Entrada Inicial de Concessionárias
def selecionar_concessionaria(estado):
    if estado == "AC":
        concessionarias = ["Energisa Acre"]
    elif estado == "AL":
        concessionarias = ["Equatorial Alagoas"]
    elif estado == "AP":
        concessionarias = ["CEA"]
    elif estado == "AM":
        concessionarias = ["Amazonas S/A"]
    elif estado == "BA":
        concessionarias = ["COELBA"]
    elif estado == "CE":
        concessionarias = ["Enel CE"]
    elif estado == "DF":
        concessionarias = ["Neoenergia Brasília"]
    elif estado == "ES":
        concessionarias = ["EDP Escelsa", "ELFSM"]
    elif estado == "GO":
        concessionarias = ["Enel GO", "Cia. Hidroelétrica São Patrício"]
    elif estado == "MA":
        concessionarias = ["Equatorial Energia Maranhão"]
    elif estado == "MT":
        concessionarias = ["Energisa Mato Grosso"]
    elif estado == "MS":
        concessionarias = ["Energisa Mato Grosso do Sul"]
    elif estado == "MG":
        concessionarias = ["CEMIG", "Energisa MG", "DME Poços de Caldas"]
    elif estado == "PA":
        concessionarias = ["Equatorial Energia Pará", "CERGAPA"]
    elif estado == "PB":
        concessionarias = ["Energisa Paraíba"]
    elif estado == "PE":
        concessionarias = ["Neoenergia Pernambuco"]
    elif estado == "PI":
        concessionarias = ["Equatorial Piauí"]
    elif estado == "PR":
        concessionarias = ["COPEL", "COCEL", "Forcel", "Cooperativa Castro", "CERAL Arapoti"]
    elif estado == "RJ":
        concessionarias = ["Light", "Enel RJ", "Energisa Nova Friburgo", "CERAL Araruama", "CERCI Papucaia",
                           "CERES"]
    elif estado == "RN":
        concessionarias = ["Neoenergia COSERN"]
    elif estado == "RO":
        concessionarias = ["CERON"]
    elif estado == "RR":
        concessionarias = ["Roraima Energia"]
    elif estado == "RS":
        concessionarias = ["CEEE", "RGE", "DEMEI Ijuí", "Hidropan", "Nova Palma Energia", "Eletrocar",
                           "MuxEnergia", "Cooperativa Centro Jacuí", "CERFOX", "CERGAL", "Ceriluz", "Cermissões",
                           "Certaja", "Certel", "Certhil", "Cooperluz", "Coopernorte", "Coopersul", "Coorsel", "Coprel",
                           "Creluz-D", "Creral"]
    elif estado == "SC":
        concessionarias = ["CELESC", "Cooperaliança", "DCELT", "Força e Luz João Cesa", "EFLUL",
                           "Cooperativa São Ludgero", "Cooperativa Jacinto Machado", "Cooperativa Praia Grande",
                           "Cooperativa Ceraça", "CERBRA Norte", "CEREJ", "CERGRAL", "CERMOFUL", "CERPALO",
                           "CERSAD", "Cersul", "Certrel", "Codesam", "Coopera", "Coopercocal", "Coopermila",
                           "Cooperzem",
                           "Iguaçu Energia"]
    elif estado == "SP":
        concessionarias = ["Enel SP", "CPFL", "CPFL Piratininga", "CPFL Santa Cruz", "Elektro",
                           "Energisa Sul Sudeste", "EDP SP", "CEDRAP", "CEDRI", "CEMIRIM", "CERIM", "CERIPA", "CERMC",
                           "CERNHE", "CERPRO", "CERRP", "Cervam", "Cetril", ]
    elif estado == "SE":
        concessionarias = ["Energisa Sergipe", "Sulgipe", "Cooperativa Centro Sul SE"]
    elif estado == "TO":
        concessionarias = ["Energisa Tocantins"]

    return concessionarias


def valor_ICMS(estado):
    if estado == "AC":
        icms = 0.17

    elif estado == "AL":
        icms = 0.17

    elif estado == "AP":
        icms = 0.18

    elif estado == "AM":
        icms = 0.18

    elif estado == "BA":
        icms = 0.18

    elif estado == "CE":
        icms = 0.18

    elif estado == "DF":
        icms = 0.18

    elif estado == "ES":
        icms = 0.17

    elif estado == "GO":
        icms = 0.17

    elif estado == "MA":
        icms = 0.18

    elif estado == "MT":
        icms = 0.17

    elif estado == "MS":
        icms = 0.17

    elif estado == "MG":
        icms = 0.18

    elif estado == "PA":
        icms = 0.17

    elif estado == "PB":
        icms = 0.18

    elif estado == "PE":
        icms = 0.18

    elif estado == "PI":
        icms = 0.18

    elif estado == "PR":
        icms = 0.18

    elif estado == "RJ":
        icms = 0.18

    elif estado == "RN":
        icms = 0.18

    elif estado == "RO":
        icms = 0.175

    elif estado == "RR":
        icms = 0.17

    elif estado == "RS":
        icms = 0.17

    elif estado == "SC":
        icms = 0.17

    elif estado == "SP":
        icms = 0.18

    elif estado == "SE":
        icms = 0.18

    elif estado == "TO":
        icms = 0.18

    return icms


def definir_sigla(value):
    if value == "Amazonas S/A":
        sigla = "AME"
    elif value == "Cooperativa Castro":
        sigla = "CASTRO-DIS"
    elif value == "CEA":
        sigla = "CEA"
    elif value == "Equatorial Alagoas":
        sigla = "Equatorial AL"
    elif value == "Neoenergia Brasília":
        sigla = "Neoenergia Brasília"
    elif value == "CEDRAP":
        sigla = "Cedrap"
    elif value == "CEDRI":
        sigla = "Cedri"
    elif value == "CEEE":
        sigla = "CEEE-D"
    elif value == "Cooperativa São Ludgero":
        sigla = "Cegero"
    elif value == "Cooperativa Jacinto Machado":
        sigla = "Cejama"
    elif value == "CELESC":
        sigla = "Celesc-DIS"
    elif value == "Cooperativa Centro Jacuí":
        sigla = "CELETRO"
    elif value == "Equatorial Energia Pará":
        sigla = "EQUATORIAL PA"
    elif value == "Neoenergia Pernambuco":
        sigla = "Neoenergia PE"
    elif value == "Equatorial Energia Maranhão":
        sigla = "Equatorial MA"
    elif value == "CEMIG":
        sigla = "Cemig-D"
    elif value == "CEMIRIM":
        sigla = "Cemirim"
    elif value == "Equatorial Piauí":
        sigla = "EQUATORIAL PI"
    elif value == "Cooperativa Praia Grande":
        sigla = "Ceprag"
    elif value == "Cooperativa Ceraça":
        sigla = "Ceraça"
    elif value == "CERAL Araruama":
        sigla = "CERAL ARARUAMA"
    elif value == "CERAL Arapoti":
        sigla = "CERAL-DIS"
    elif value == "CERBRA Norte":
        sigla = "Cerbranorte"
    elif value == "CERCI Papucaia":
        sigla = "CERCI"
    elif value == "Cooperativa Centro Sul SE":
        sigla = "Cercos"
    elif value == "CEREJ":
        sigla = "Cerej"
    elif value == "CERES":
        sigla = "Ceres"
    elif value == "CERFOX":
        sigla = "Cerfox"
    elif value == "CERGAL":
        sigla = "Cergal"
    elif value == "CERGAPA":
        sigla = "Cergapa"
    elif value == "CERGRAL":
        sigla = "Cergral"
    elif value == "Ceriluz":
        sigla = "Ceriluz"
    elif value == "CERIM":
        sigla = "Cerim"
    elif value == "CERIPA":
        sigla = "Ceripa"
    elif value == "CERMC":
        sigla = "CERMC"
    elif value == "Cermissões":
        sigla = "Cermissões"
    elif value == "CERMOFUL":
        sigla = "Cermoful"
    elif value == "CERNHE":
        sigla = "Cernhe"
    elif value == "CERON":
        sigla = "ERO"
    elif value == "CERPALO":
        sigla = "Cerpalo"
    elif value == "CERPRO":
        sigla = "Cerpro"
    elif value == "CERRP":
        sigla = "CERRP"
    elif value == "CERSAD":
        sigla = "CERSAD DISTRIBUIDORA"
    elif value == "Cia. Hidroelétrica São Patrício":
        sigla = "Chesp"
    elif value == "COCEL":
        sigla = "Cocel"
    elif value == "COELBA":
        sigla = "COELBA"
    elif value == "COPEL":
        sigla = "COPEL-DIS"
    elif value == "Neoenergia COSERN":
        sigla = "Cosern"
    elif value == "CPFL":
        sigla = "CPFL-PAULISTA"
    elif value == "Enel SP":
        sigla = "ELETROPAULO"
    elif value == "DEMEI Ijuí":
        sigla = "DEMEI"
    elif value == "DME Poços de Caldas":
        sigla = "DMED"
    elif value == "Energisa Paraíba":
        sigla = "EBO"
    elif value == "EDP Escelsa":
        sigla = "EDP ES"
    elif value == "Força e Luz João Cesa":
        sigla = "EFLJC"
    elif value == "EFLUL":
        sigla = "Eflul"
    elif value == "Energisa Acre":
        sigla = "EAC"
    elif value == "Energisa MG":
        sigla = "EMR"
    elif value == "Energisa Mato Grosso do Sul":
        sigla = "EMS"
    elif value == "Energisa Mato Grosso":
        sigla = "EMT"
    elif value == "Energisa Nova Friburgo":
        sigla = "ENF"
    elif value == "Energisa Paraíba":
        sigla = "EPB"
    elif value == "Energisa Sergipe":
        sigla = "ESE"
    elif value == "Energisa Sul Sudeste":
        sigla = "ESS"
    elif value == "Energisa Tocantins":
        sigla = "ETO"
    elif value == "Iguaçu Energia":
        sigla = "Ienergia"
    elif value == "Nova Palma Energia":
        sigla = "Uhenpal"
    elif value == "Enel GO":
        sigla = "EQUATORIAL GO"  # Enel GO foi rebatizada Equatorial Goiás
    elif value == "Light":
        sigla = "LIGHT SESA"
    elif value == "Roraima Energia":
        sigla = "Boa Vista"  # nome anterior da concessionária na base ANEEL
    elif value == "Certel":
        sigla = "CERTEL ENERGIA"
    elif value == "CPFL Piratininga":
        sigla = "CPFL-PIRATINING"
    else:
        sigla = (value)

    return sigla

st.set_page_config(page_title="Otimização de Demanda", page_icon=":zap:", layout="wide", initial_sidebar_state="auto", menu_items=None)

with st.sidebar:
    st.write("Programa desenvolvido com objetivo de otimizar a demanda dos clientes do Grupo A")
    st.write("Em caso de dúvidas ou sugestões para melhoria, estou à disposição para contato")
    st.write("Email: rodolfosixel@gmail.com")
    st.write("---")
    st.write("Desenvolvido por: Rodolfo Almeida Sixel Juliani")
    st.write("---")
    st.write("Versão 1.2")



st.title("""Otimização de Demanda :zap:""")

tab_dados, tab_tarifas, tab_simulacao = st.tabs(
    ["1. Dados de Entrada", "2. Tarifas e Situação Atual", "3. Simulação e Resultados"]
)

with tab_dados:
    st.header("""Entrada de dados""")

    with st.expander(("Padrões para entrada de dados")):
        st.markdown((
            """
            É possível copiar e colar as colunas do Excel diretamente na planilha abaixo.

            É muito importante observar que o programa aceita apenas "." como separador decimal.

            Sugere-se uma conferência dos dados após colagem.
        """
        ))

    sample_data = [
        {
            "mes": months[0],
            "demanda_ponta": 10,
            "demanda_fora_ponta": 20,
            "consumo_ponta": 300,
            "consumo_fora_ponta": 2000,
        },
        {
            "mes": months[1],
            "demanda_ponta": 10,
            "demanda_fora_ponta": 20,
            "consumo_ponta": 300,
            "consumo_fora_ponta": 2000,
        },
        {
            "mes": months[2],
            "demanda_ponta": 10,
            "demanda_fora_ponta": 20,
            "consumo_ponta": 300,
            "consumo_fora_ponta": 2000,
        },
        {
            "mes": months[3],
            "demanda_ponta": 10,
            "demanda_fora_ponta": 20,
            "consumo_ponta": 300,
            "consumo_fora_ponta": 2000,
        },
        {
            "mes": months[4],
            "demanda_ponta": 10.9,
            "demanda_fora_ponta": 20,
            "consumo_ponta": 300,
            "consumo_fora_ponta": 2000,
        },
        {
            "mes": months[5],
            "demanda_ponta": 10,
            "demanda_fora_ponta": 20,
            "consumo_ponta": 300,
            "consumo_fora_ponta": 2000,
        },
        {
            "mes": months[6],
            "demanda_ponta": 10,
            "demanda_fora_ponta": 20,
            "consumo_ponta": 300,
            "consumo_fora_ponta": 2000,
        },
        {
            "mes": months[7],
            "demanda_ponta": 10,
            "demanda_fora_ponta": 20,
            "consumo_ponta": 300,
            "consumo_fora_ponta": 2000,
        },
        {
            "mes": months[8],
            "demanda_ponta": 10,
            "demanda_fora_ponta": 20,
            "consumo_ponta": 300,
            "consumo_fora_ponta": 2000,
        },
        {
            "mes": months[9],
            "demanda_ponta": 10,
            "demanda_fora_ponta": 20,
            "consumo_ponta": 300,
            "consumo_fora_ponta": 2000,
        },
        {
            "mes": months[10],
            "demanda_ponta": 10,
            "demanda_fora_ponta": 20,
            "consumo_ponta": 300,
            "consumo_fora_ponta": 2000,
        },
        {
            "mes": months[11],
            "demanda_ponta": 10,
            "demanda_fora_ponta": 20,
            "consumo_ponta": 300,
            "consumo_fora_ponta": 2000,
        },
    ]

    config = {
        "mes": st.column_config.TextColumn("Mês", required=True),
        "demanda_ponta": st.column_config.NumberColumn("Demanda na ponta"),
        "demanda_fora_ponta": st.column_config.NumberColumn("Demanda fora da ponta"),
        "consumo_ponta": st.column_config.NumberColumn("Consumo na ponta"),
        "consumo_fora_ponta": st.column_config.NumberColumn("Consumo fora da ponta"),
    }

    dados_entrada_planilhas = st.data_editor(sample_data, column_config=config, num_rows="dynamic")

vetor_demanda_ponta = []
vetor_demanda_fp = []
vetor_consumo_ponta= []
vetor_consumo_fp = []

for i in range (0, 12):
    vetor_demanda_ponta.append(dados_entrada_planilhas[i]['demanda_ponta'])
    vetor_demanda_fp.append(dados_entrada_planilhas[i]['demanda_fora_ponta'])
    vetor_consumo_ponta.append(dados_entrada_planilhas[i]['consumo_ponta'])
    vetor_consumo_fp.append(dados_entrada_planilhas[i]['consumo_fora_ponta'])

demanda_maxima = max(vetor_demanda_fp)
limite_demanda = 4*demanda_maxima


def _extrair_tarifas(banco, sigla_norm, grupo, cor, impostos, descricao_fonte):
    """
    Lógica compartilhada de extração de tarifas (TUSD + TE) a partir de um
    DataFrame de tarifas. Reutilizada tanto pela planilha local quanto pela
    API da ANEEL, que alimentam `banco` já no mesmo formato de colunas
    (Sigla_norm, Subgrupo, Modalidade, Detalhe, Base Tarifária, Posto,
    Unidade, TUSD, TE).
    """
    try:
        if cor == 'Verde':
            # filtro para concessionária, grupo e modalidade tarifária
            banco_novo = banco.loc[(banco['Sigla_norm'] == sigla_norm) & (banco['Subgrupo'] == grupo)
                                   & (banco['Modalidade'] == 'Verde') & (banco['Detalhe'] == 'Não se aplica')
                                   & (banco['Base Tarifária'] == 'Tarifa de Aplicação')]

            # filtro para valores de energia fora da ponta
            linha = banco_novo.loc[(banco_novo['Posto'] == 'Fora ponta') & (banco_novo['Unidade'] == 'MWh')]

            # filtros e equação para determinar a tarifa de consumo fora da ponta (kWh) Tarifa = TE + TUSD
            preco_consumo_fp = (1 + impostos) * (float(linha.iloc[0]['TUSD']) + float(linha.iloc[0]['TE'])) / 1000

            # filtro para valores de energia na ponta
            linha = banco_novo.loc[(banco_novo['Posto'] == 'Ponta') & (banco_novo['Unidade'] == 'MWh')]

            # filtros e equação para determinar a tarifa de consumo na ponta (kWh) Tarifa = TE + TUSD
            preco_consumo_ponta = (1 + impostos) * (
                        float(linha.iloc[0]['TUSD']) + float(linha.iloc[0]['TE'])) / 1000

            # filtro para valores de demanda (kW) > 'Não se aplica' é utilizado em Modalidade Verde
            linha = banco_novo.loc[(banco_novo['Posto'] == 'Não se aplica') & (banco_novo['Unidade'] == 'kW')]

            # filtros e equação para determinar a tarifa de demanda
            preco_demanda_fp = (1 + impostos) * float(linha.iloc[0]['TUSD']) + float(linha.iloc[0]['TE'])

            # Multa = 2 * preço
            preco_demanda_ult_fp = 2 * preco_demanda_fp

            # se a Modalidade é verde, então demanda na ponta = 0
            preco_demanda_ponta = 0

            # Multa = 2 * preço
            preco_demanda_ult_ponta = 2 * preco_demanda_ponta

        elif cor == 'Azul':
            # filtro para concessionária, grupo e modalidade tarifária
            banco_novo = banco.loc[(banco['Sigla_norm'] == sigla_norm) & (banco['Subgrupo'] == grupo)
                                   & (banco['Modalidade'] == 'Azul') & (banco['Detalhe'] == 'Não se aplica')
                                   & (banco['Base Tarifária'] == 'Tarifa de Aplicação')]

            # filtro para valores de energia fora da ponta
            linha = banco_novo.loc[(banco_novo['Posto'] == 'Fora ponta') & (banco_novo['Unidade'] == 'MWh')]

            # filtros e equação para determinar a tarifa de consumo fora da ponta (kWh) Tarifa = TE + TUSD
            preco_consumo_fp = round((1 + impostos) * (float(linha.iloc[0]['TUSD']) + float(linha.iloc[0]['TE'])) / 1000, 2)

            # filtro para valores de energia na ponta
            linha = banco_novo.loc[(banco_novo['Posto'] == 'Ponta') & (banco_novo['Unidade'] == 'MWh')]

            # filtros e equação para determinar a tarifa de consumo na ponta (kWh) Tarifa = TE + TUSD
            preco_consumo_ponta = round((1 + impostos) * (
                        float(linha.iloc[0]['TUSD']) + float(linha.iloc[0]['TE'])) / 1000,2)

            # filtro para valores de demanda (kW) fora da ponta
            linha = banco_novo.loc[(banco_novo['Posto'] == 'Fora ponta') & (banco_novo['Unidade'] == 'kW')]

            # filtros e equação para determinar a tarifa de demanda fora da ponta (kW) Tarifa = TE + TUSD
            preco_demanda_fp = round((1 + impostos) * float(linha.iloc[0]['TUSD']) + float(linha.iloc[0]['TE']), 2)

            # Multa = 2 * preço
            preco_demanda_ult_fp = round(2 * preco_demanda_fp, 2)

            # filtro para valores de demanda (kW) na ponta
            linha = banco_novo.loc[(banco_novo['Posto'] == 'Ponta') & (banco_novo['Unidade'] == 'kW')]

            # filtros e equação para determinar a tarifa de demanda na ponta (kW) Tarifa = TE + TUSD
            preco_demanda_ponta = round((1 + impostos) * (float(linha.iloc[0]['TUSD']) + float(linha.iloc[0]['TE'])), 2)

            # Multa = 2 * preço
            preco_demanda_ult_ponta = round(2 * preco_demanda_ponta, 2)

        else:  # definir os valores como zero caso a modalidade tarifária não tenha sido selecionada
            preco_demanda_fp = 0
            preco_demanda_ult_fp = 0
            preco_demanda_ponta = 0
            preco_demanda_ult_ponta = 0
            preco_consumo_fp = 0
            preco_consumo_ponta = 0

    except (IndexError, KeyError):
        st.error(
            f"Não encontrei tarifas (modalidade {cor}, Subgrupo {grupo}) em {descricao_fonte}. Isso "
            f"normalmente significa que esta concessionária não tem tarifa deste subgrupo publicada "
            f"nesta fonte. Tente outro subgrupo, outra concessionária, outra fonte de tarifas ou "
            f"digite os valores manualmente."
        )
        st.stop()

    tarifas = [preco_demanda_fp, preco_demanda_ult_fp, preco_demanda_ponta, preco_demanda_ult_ponta,
               preco_consumo_fp, preco_consumo_ponta]

    return tarifas


def obter_tarifas(cor, fonte='planilha'):
    """
    Obtém as tarifas (TUSD + TE) para a modalidade `cor`, a partir da fonte
    escolhida: 'planilha' (arquivo Excel local) ou 'api' (API de dados
    abertos da ANEEL, consultada em tempo real).
    """
    pis_cofins = 0.08  # valor padronizado de PIS e COFINS
    estado = estado_selecionado  # estado selecionado pelo usuario
    icms = valor_ICMS(estado)  # valor do ICMS para o estado selecionado

    impostos = round(pis_cofins + icms, 2)  # valor de impostos total

    impostos = 0

    conc = concessionaria_selecionada
    sigla = sigla_conc  # Sigla da concessionária
    sigla_norm = normaliza_sigla(sigla)
    grupo = grupo_tarifario  # subgrupo tarifário (A1, A2, A3 ou A4) selecionado pelo usuário

    if fonte == 'api':
        try:
            banco = consultar_tarifas_api_aneel(sigla, grupo)
        except RuntimeError as exc:
            st.error(
                f"Erro ao consultar a API de dados abertos da ANEEL para **{conc}** (Subgrupo {grupo}): "
                f"{exc}. Tente novamente, use a planilha local ou digite as tarifas manualmente."
            )
            st.stop()
        descricao_fonte = "API de dados abertos da ANEEL"
    else:
        excel_file = 'Tarifas_Teste_2025.xlsx'
        banco = carregar_planilha_tarifas(excel_file)
        descricao_fonte = f"planilha local (`{excel_file}`)"

    return _extrair_tarifas(banco, sigla_norm, grupo, cor, impostos, descricao_fonte)


def obter_tarifas_efetivas(cor):
    # Ponto único usado por todos os cálculos (custo atual, varreduras e simulações).
    # Decide se as tarifas vêm da API da ANEEL, da planilha local ou dos campos
    # digitados manualmente pelo usuário na aba "Tarifas e Situação Atual",
    # conforme a opção "Origem das tarifas" selecionada ali.
    origem = st.session_state.get("origem_tarifas")

    if origem == "Inserir manualmente":
        demanda_fp = st.session_state.get("man_demanda_fp", 0.0) or 0.0
        demanda_ponta_azul = st.session_state.get("man_demanda_ponta_azul", 0.0) or 0.0
        consumo_fp = st.session_state.get("man_consumo_fp", 0.0) or 0.0
        consumo_ponta_verde = st.session_state.get("man_consumo_ponta_verde", 0.0) or 0.0
        consumo_ponta_azul = st.session_state.get("man_consumo_ponta_azul", 0.0) or 0.0

        if cor == "Verde":
            return [demanda_fp, 2 * demanda_fp, 0, 0, consumo_fp, consumo_ponta_verde]
        elif cor == "Azul":
            return [demanda_fp, 2 * demanda_fp, demanda_ponta_azul, 2 * demanda_ponta_azul,
                    consumo_fp, consumo_ponta_azul]
        else:
            return [0, 0, 0, 0, 0, 0]

    if origem == "API da ANEEL (tempo real)":
        return obter_tarifas(cor, fonte='api')

    return obter_tarifas(cor, fonte='planilha')


def custo_atual():
    conc = concessionaria_selecionada
    custo_total = 0
    demanda_verde = float(demanda_contratada_verde)
    demanda_azul = float(demanda_contratada_azul)

    demanda_teste, demanda_contratada_teste = vetor_demanda_fp, demanda_verde  # receber valores de demanda
    # vetor_consumo_fp, vetor_consumo_ponta = vetoreceber_valores_consumo()
    demanda_ponta_teste, demanda_contratada_ponta_teste = vetor_demanda_ponta, demanda_azul

    if modalidade == "Verde":
        tarifa_vec = obter_tarifas_efetivas(cor='Verde')  # definição de tarifas
        valor_fp = objetivo_fp(tarifa_vec, demanda_teste, demanda_verde)
        consumo, gasto_consumo_fp_verde, gasto_consumo_ponta_verde = gastos_consumo(tarifa_vec, vetor_consumo_fp,
                                                                                    vetor_consumo_ponta)
        custo_total = valor_fp + consumo
        custo_demanda = valor_fp

    elif modalidade == "Azul":
        tarifa_vec = obter_tarifas_efetivas(cor='Azul')
        valor_fp = objetivo_fp(tarifa_vec, demanda_teste, demanda_verde)
        valor_ponta = objetivo_ponta(tarifa_vec, demanda_ponta_teste, demanda_azul)
        consumo, gasto_consumo_fp_azul, gasto_consumo_ponta_azul = gastos_consumo(tarifa_vec, vetor_consumo_fp,
                                                                                  vetor_consumo_ponta)
        custo_total = round(valor_fp + valor_ponta + consumo, 2)
        custo_demanda = round(valor_fp + valor_ponta, 2)

    return custo_total, custo_demanda


gasto_anual = 30


def objetivo_ponta(tarifas, vetor_demanda, x):
    tarifa_ponta = tarifas[2]
    multa_ponta = tarifas[3]
    f_obj = 0
    custo_demanda = 0
    custo_multa = 0
    vetor_demanda_ult_novo = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
    obj = []

    # preenchendo valores de demanda ultrapassada que geram multa
    for i in range(0, 12):
        if vetor_demanda[i] > 1.05 * x:
            vetor_demanda_ult_novo[i] = vetor_demanda[i] - x
        else:
            vetor_demanda_ult_novo[i] = 0

    for i in range(0, 12):
        # 1 - demanda < contratada ⇨ sem multa e valor faturado é o contratado
        if vetor_demanda[i] <= x:
            f_obj += tarifa_ponta * x
            obj.append(tarifa_ponta * x)
            custo_demanda += tarifa_ponta * x

        # 2 - demanda > 1.05 * contratada ⇨ com multa
        elif vetor_demanda[i] > 1.05 * x:
            f_obj += tarifa_ponta * vetor_demanda[i] + multa_ponta * vetor_demanda_ult_novo[i]
            obj.append(tarifa_ponta * vetor_demanda[i] + multa_ponta * vetor_demanda_ult_novo[i])
            custo_demanda += tarifa_ponta * vetor_demanda[i]
            custo_multa += multa_ponta * vetor_demanda_ult_novo[i]

        # 3 - demanda > contratada, mas não supera os 5% ⇨ sem multa e valor faturado é o medido
        elif x < vetor_demanda[i] < 1.05 * x:
            f_obj += tarifa_ponta * vetor_demanda[i]
            obj.append(tarifa_ponta * vetor_demanda[i])
            custo_demanda += tarifa_ponta * vetor_demanda[i]

    return f_obj


def objetivo_fp(tarifas, vetor_demanda, x):
    tarifa_fp = tarifas[0]
    multa_fp = tarifas[1]
    f_obj = 0  # valor de custo anual
    custo_demanda = 0
    custo_multa = 0
    vetor_demanda_ult_novo = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
    obj = []

    # criando vetor de demanda ultrapassada (caso seja 5% acima do valor contratado)
    for i in range(0, 12):
        if vetor_demanda[i] > 1.05 * x:
            vetor_demanda_ult_novo[i] = vetor_demanda[i] - x
        else:
            vetor_demanda_ult_novo[i] = 0

    for i in range(0, 12):  # adicionar os valores à função objetivo conforme a situação
        # 1 - demanda < contratada ⇨ sem multa e valor faturado é o contratado
        if vetor_demanda[i] <= x:
            f_obj += tarifa_fp * x
            obj.append(tarifa_fp * x)
            custo_demanda += tarifa_fp * x

        # 2 - demanda > 1.05 * contratada ⇨ com multa
        elif vetor_demanda[i] > 1.05 * x:
            f_obj += tarifa_fp * vetor_demanda[i] + multa_fp * vetor_demanda_ult_novo[i]
            obj.append(tarifa_fp * vetor_demanda[i] + multa_fp * vetor_demanda_ult_novo[i])
            custo_demanda += tarifa_fp * x
            # custo_multa = multa_fp * vetor_demanda_ult_novo[i]

        # 3 - demanda > contratada, mas não supera os 5% ⇨ sem multa e valor faturado é o medido
        elif x < vetor_demanda[i] < 1.05 * x:
            f_obj += tarifa_fp * vetor_demanda[i]
            obj.append(tarifa_fp * vetor_demanda[i])
            custo_demanda += tarifa_fp * x
    return f_obj


def gastos_consumo(tarifas, consumo_fp, consumo_ponta):
    gasto_consumo_fp = 0
    gasto_consumo_ponta = 0

    for i in range(0, 12):
        gasto_consumo_fp += tarifas[4] * consumo_fp[i]
        gasto_consumo_ponta += tarifas[5] * consumo_ponta[i]

    total = gasto_consumo_fp + gasto_consumo_ponta

    return total, gasto_consumo_fp, gasto_consumo_ponta

vec_otimo = []


def varredura(a, b, demanda_contratada):
    # Função para cálculo da melhor demanda utilizando busca extensiva por varredura
    # A função roda por todos os valores definidos dentro dos limites (a,b) e checa o custo total para cada demanda
    tarifas = obter_tarifas_efetivas("Verde")

    otimo_varredura = objetivo_fp(tarifas, vetor_demanda_fp, float(demanda_contratada))
    demanda_otima = demanda_contratada
    for x in range(a, b):
        teste = objetivo_fp(tarifas, vetor_demanda_fp, x)
        vec_otimo.append(teste)

        if teste < otimo_varredura:
            otimo_varredura = objetivo_fp(tarifas, vetor_demanda_fp, x)
            demanda_otima = x

    # plot_otimo_verde(vec_otimo,demanda_otima)
    return otimo_varredura, demanda_otima


def varredura_azul(a, b, demanda_contratada, demanda_contratada_azul):
    # Função para cálculo da melhor demanda utilizando busca extensiva por varredura
    # A função roda por todos os valores definidos dentro dos limites (a,b) e checa o custo total para cada demanda
    tarifas = obter_tarifas_efetivas("Verde")

    otimo_varredura_fp = objetivo_fp(tarifas, vetor_demanda_fp, float(demanda_contratada))
    demanda_otima_fp = demanda_contratada
    for x in range(a, b):
        teste = objetivo_fp(tarifas, vetor_demanda_fp, x)
        if teste < otimo_varredura_fp:
            otimo_varredura_fp = objetivo_fp(tarifas, vetor_demanda_fp, x)
            demanda_otima_fp = x

    tarifas = obter_tarifas_efetivas("Azul")
    otimo_varredura_ponta = objetivo_ponta(tarifas, vetor_demanda_ponta, float(demanda_contratada_azul))
    demanda_otima_ponta = demanda_contratada_azul
    for x in range(a, b):
        teste = objetivo_ponta(tarifas, vetor_demanda_ponta, x)
        if teste < otimo_varredura_ponta:
            otimo_varredura_ponta = objetivo_ponta(tarifas, vetor_demanda_ponta, x)
            demanda_otima_ponta = x

    return otimo_varredura_fp, demanda_otima_fp, otimo_varredura_ponta, demanda_otima_ponta

# Paleta categórica (skill de dataviz) validada com scripts/validate_palette.js
# — ordem fixa por série, nunca redistribuída conforme os dados mudam.
COR_SERIE_1 = "#2a78d6"  # azul
COR_SERIE_2 = "#eb6834"  # laranja
COR_SERIE_3 = "#1baf7a"  # verde-água
COR_SERIE_4 = "#eda100"  # amarelo


# As três funções "construir_grafico_*" abaixo só MONTAM a figura Plotly e a devolvem —
# não desenham nada na tela. São usadas tanto pela versão interativa (plotar_*, que
# chama st.plotly_chart) quanto pelo relatório em PDF (que exporta a mesma figura como
# imagem estática via kaleido). Isso garante que o gráfico Verde e o gráfico Azul sigam
# sempre o mesmo padrão visual, tanto na tela quanto no PDF, porque nascem do mesmo código.

def construir_grafico_verde(demanda_contratada, demanda_otima, demanda_fp):
    # Gráfico que compara demanda contratada atual, demanda ótima sugerida e demanda medida (Modalidade Verde)
    demanda_contratada = float(demanda_contratada)
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=months, y=[demanda_contratada] * 12, mode='lines', name='Demanda Atual',
                              line=dict(color=COR_SERIE_1, dash='dash', width=3)))
    fig.add_trace(go.Scatter(x=months, y=[demanda_otima] * 12, mode='lines', name='Demanda Sugerida',
                              line=dict(color=COR_SERIE_2, width=3)))
    fig.add_trace(go.Scatter(x=months, y=demanda_fp, mode='lines+markers', name='Demanda Medida',
                              line=dict(color=COR_SERIE_3, dash='dashdot', width=3), marker=dict(size=8)))
    fig.update_layout(
        title='Simulação: Modalidade Tarifária Verde',
        xaxis_title='Meses',
        yaxis_title='Demanda de Potência (kW)',
        hovermode='x unified',
        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
    )
    return fig


def construir_grafico_azul(demanda_contratada_ponta, demanda_otima_ponta, demanda_fp):
    # Mesmo padrão visual do gráfico Verde (mesmas cores/traços/legenda), só troca o eixo
    demanda_contratada_ponta = float(demanda_contratada_ponta)
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=months, y=[demanda_contratada_ponta] * 12, mode='lines', name='Demanda na Ponta Atual',
                              line=dict(color=COR_SERIE_1, dash='dash', width=3)))
    fig.add_trace(go.Scatter(x=months, y=[demanda_otima_ponta] * 12, mode='lines', name='Demanda na Ponta Sugerida',
                              line=dict(color=COR_SERIE_2, width=3)))
    fig.add_trace(go.Scatter(x=months, y=demanda_fp, mode='lines+markers', name='Demanda Medida',
                              line=dict(color=COR_SERIE_3, dash='dashdot', width=3), marker=dict(size=8)))
    fig.update_layout(
        title='Simulação: Modalidade Tarifária Azul',
        xaxis_title='Meses',
        yaxis_title='Demanda de Potência na Ponta (kW)',
        hovermode='x unified',
        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
    )
    return fig


def construir_grafico_completo(demanda_total_verde, demanda_total_azul, gasto_consumo_fp_verde,
                                gasto_consumo_ponta_verde, gasto_consumo_fp_azul, gasto_consumo_ponta_azul):
    # Comparação empilhada dos custos totais das modalidades Verde e Azul
    categorias = ['Modalidade Verde', 'Modalidade Azul']
    demanda_fp_vals = [demanda_total_verde, demanda_total_verde]
    energia_fp_vals = [gasto_consumo_fp_verde, gasto_consumo_fp_azul]
    demanda_ponta_vals = [0, demanda_total_azul]
    energia_ponta_vals = [gasto_consumo_ponta_verde, gasto_consumo_ponta_azul]

    fig = go.Figure()
    fig.add_trace(go.Bar(name='Demanda Fora Ponta', x=categorias, y=demanda_fp_vals, marker_color=COR_SERIE_1,
                          text=[f"R$ {v:,.0f}" for v in demanda_fp_vals], textposition='inside'))
    fig.add_trace(go.Bar(name='Energia Fora Ponta', x=categorias, y=energia_fp_vals, marker_color=COR_SERIE_2,
                          text=[f"R$ {v:,.0f}" for v in energia_fp_vals], textposition='inside'))
    fig.add_trace(go.Bar(name='Demanda Ponta', x=categorias, y=demanda_ponta_vals, marker_color=COR_SERIE_3,
                          text=[f"R$ {v:,.0f}" for v in demanda_ponta_vals], textposition='inside'))
    fig.add_trace(go.Bar(name='Energia Ponta', x=categorias, y=energia_ponta_vals, marker_color=COR_SERIE_4,
                          text=[f"R$ {v:,.0f}" for v in energia_ponta_vals], textposition='inside'))
    fig.update_layout(
        barmode='stack',
        title='Comparação de Custos',
        yaxis_title='Valor Total (R$)',
        hovermode='x unified',
        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
    )
    return fig


def plotar_verde(demanda_contratada, demanda_otima, demanda_fp, key="chart_verde"):
    st.plotly_chart(construir_grafico_verde(demanda_contratada, demanda_otima, demanda_fp),
                     width='stretch', key=key)


def plotar_azul(demanda_contratada_ponta, demanda_otima_ponta, demanda_fp, key="chart_azul"):
    st.plotly_chart(construir_grafico_azul(demanda_contratada_ponta, demanda_otima_ponta, demanda_fp),
                     width='stretch', key=key)


def plotar_completo(demanda_total_verde, demanda_total_azul, gasto_consumo_fp_verde,
                     gasto_consumo_ponta_verde, gasto_consumo_fp_azul, gasto_consumo_ponta_azul,
                     key="chart_completo"):
    fig = construir_grafico_completo(demanda_total_verde, demanda_total_azul, gasto_consumo_fp_verde,
                                      gasto_consumo_ponta_verde, gasto_consumo_fp_azul, gasto_consumo_ponta_azul)
    st.plotly_chart(fig, width='stretch', key=key)


def fig_para_png(fig, width=900, height=480, scale=2):
    # Exporta a MESMA figura (mesmo tamanho/escala para Verde e Azul) como PNG para uso no relatório PDF
    return fig.to_image(format="png", width=width, height=height, scale=scale)


ESTILO_TABELA_INFO = TableStyle([
    ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#f0efec")),
    ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#c9c8c2")),
    ("FONTSIZE", (0, 0), (-1, -1), 9),
    ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ("TOPPADDING", (0, 0), (-1, -1), 4),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
])

ESTILO_TABELA_MENSAL = TableStyle([
    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2a78d6")),
    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
    ("FONTSIZE", (0, 0), (-1, -1), 8),
    ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#c9c8c2")),
    ("ALIGN", (1, 0), (-1, -1), "CENTER"),
    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f7f7f5")]),
    ("TOPPADDING", (0, 0), (-1, -1), 3),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
])


def _imagem_grafico(fig, largura_cm=16):
    # Insere o gráfico Plotly no PDF preservando a mesma proporção usada em fig_para_png,
    # para que os gráficos Verde e Azul sempre saiam do mesmo tamanho no relatório.
    png_bytes = fig_para_png(fig)
    largura = largura_cm * cm
    altura = largura * (480 / 900)
    return RLImage(io.BytesIO(png_bytes), width=largura, height=altura)


def gerar_relatorio_pdf():
    # Monta o relatório em PDF com os dados de entrada e os resultados das simulações
    # que já foram calculadas nesta sessão (Verde, Azul e/ou Completa - o que estiver disponível).
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4,
        leftMargin=2 * cm, rightMargin=2 * cm, topMargin=2 * cm, bottomMargin=2 * cm,
    )
    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph("Relatório de Otimização de Demanda", styles["Title"]))
    story.append(Paragraph(f"Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}", styles["Normal"]))
    story.append(Spacer(1, 0.5 * cm))

    # --- Dados de entrada ---
    story.append(Paragraph("Dados de Entrada", styles["Heading2"]))
    origem = st.session_state.get("origem_tarifas", "API da ANEEL (tempo real)")
    tabela_info = Table(
        [
            ["Estado", estado_selecionado],
            ["Concessionária", concessionaria_selecionada],
            ["Sigla (base de tarifas)", sigla_conc],
            ["Subgrupo tarifário (tipo de consumidor)", grupo_tarifario],
            ["Modalidade tarifária selecionada", modalidade],
            ["Origem das tarifas", origem],
        ],
        colWidths=[6 * cm, 9 * cm],
    )
    tabela_info.setStyle(ESTILO_TABELA_INFO)
    story.append(tabela_info)
    story.append(Spacer(1, 0.4 * cm))

    # --- Dados mensais informados ---
    story.append(Paragraph("Dados Mensais Informados", styles["Heading2"]))
    linhas = [["Mês", "Demanda Ponta (kW)", "Demanda Fora Ponta (kW)", "Consumo Ponta (kWh)",
               "Consumo Fora Ponta (kWh)"]]
    for i in range(12):
        linhas.append([
            months[i],
            f"{vetor_demanda_ponta[i]:g}",
            f"{vetor_demanda_fp[i]:g}",
            f"{vetor_consumo_ponta[i]:g}",
            f"{vetor_consumo_fp[i]:g}",
        ])
    tabela_mensal = Table(linhas, colWidths=[2.6 * cm, 3.1 * cm, 3.4 * cm, 3.1 * cm, 3.4 * cm], repeatRows=1)
    tabela_mensal.setStyle(ESTILO_TABELA_MENSAL)
    story.append(tabela_mensal)

    def secao_resultado(titulo, linhas_metricas, figuras):
        story.append(PageBreak())
        story.append(Paragraph(titulo, styles["Heading2"]))
        tabela = Table(linhas_metricas, colWidths=[7 * cm, 8 * cm])
        tabela.setStyle(ESTILO_TABELA_INFO)
        story.append(tabela)
        story.append(Spacer(1, 0.4 * cm))
        for fig in figuras:
            story.append(_imagem_grafico(fig))
            story.append(Spacer(1, 0.3 * cm))

    tem_resultado = False

    if "resultado_verde" in st.session_state:
        tem_resultado = True
        r = st.session_state["resultado_verde"]
        secao_resultado(
            "Resultado: Modalidade Verde",
            [
                ["Valor ótimo", f"R$ {money_format(round(r['valor_otimo'], 2))}"],
                ["Demanda sugerida (fora da ponta)", f"{r['demanda_otima']} kW"],
                ["Economia anual estimada", f"R$ {money_format(round(r['economia'], 2))}"],
            ],
            [construir_grafico_verde(r["demanda_contratada"], r["demanda_otima"], r["demanda_fp"])],
        )

    if "resultado_azul" in st.session_state:
        tem_resultado = True
        r = st.session_state["resultado_azul"]
        secao_resultado(
            "Resultado: Modalidade Azul",
            [
                ["Valor ótimo", f"R$ {money_format(round(r['valor_otimo_azul'], 2))}"],
                ["Demanda sugerida (ponta)", f"{r['demanda_otima_azul']} kW"],
                ["Economia anual estimada", f"R$ {money_format(round(r['economia'], 2))}"],
            ],
            [construir_grafico_azul(r["demanda_contratada_azul"], r["demanda_otima_azul"], r["demanda_ponta"])],
        )

    if "resultado_completo" in st.session_state:
        tem_resultado = True
        r = st.session_state["resultado_completo"]
        fig_verde = construir_grafico_verde(r["demanda_contratada_verde"], r["demanda_otima_verde"], r["demanda_fp"])
        fig_azul = construir_grafico_azul(r["demanda_contratada_azul"], r["demanda_otima_azul"], r["demanda_ponta"])
        fig_comp = construir_grafico_completo(
            r["valor_otimo"], r["valor_otimo_azul"], r["gasto_consumo_fp_verde"],
            r["gasto_consumo_ponta_verde"], r["gasto_consumo_fp_azul"], r["gasto_consumo_ponta_azul"],
        )
        secao_resultado(
            "Resultado: Comparação Completa (Verde x Azul)",
            [
                ["Custo Atual", f"R$ {money_format(round(r['custo_soma'], 2))}"],
                ["Valor Total Verde", f"R$ {money_format(round(r['custo_total_verde'], 2))}"],
                ["Valor Total Azul", f"R$ {money_format(round(r['custo_total_azul'], 2))}"],
                ["Demanda Ótima Fora da Ponta", f"{r['demanda_otima_verde']} kW"],
                ["Demanda Ótima na Ponta", f"{r['demanda_sugerida_ponta']} kW"],
                ["Modalidade Sugerida", r["modalidade_sugerida"]],
                ["Economia anual estimada", f"R$ {money_format(round(r['economia'], 2))}"],
            ],
            [fig_verde, fig_azul, fig_comp],
        )

    if not tem_resultado:
        story.append(Spacer(1, 0.4 * cm))
        story.append(Paragraph(
            "Nenhuma simulação foi executada nesta sessão ainda — rode ao menos uma simulação "
            "na aba \"Simulação e Resultados\" antes de gerar o relatório.",
            styles["Normal"],
        ))

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


with tab_tarifas:
    st.header("Situação Atual e Importação de Tarifas")

    with st.expander(("Passo a passo")):
        st.markdown((
            """
            1. Selecionar modalidade tarifária

            2. Inserir valores de demanda contratada

            3. Selecionar o estado e a concessionária de interesse

            4. Obter as tarifas: escolha a origem ("API da ANEEL", "Planilha de 2025" ou
            "Inserir manualmente") e, se aplicável, clique em "Importar tarifas"

            5. Calcular o gasto anual (botão "Calcular gasto anual") atual da unidade consumidora
        """
        ))

    coluna1, coluna2 = st.columns(2)

    with coluna1:

        modalidade = st.radio("Modalidade Tarifária", ["Azul", "Verde"])

        disable = False

        if modalidade == "Verde":
            disable = True

        demanda_contratada_verde = st.number_input(
            "Demanda Contratada Fora da Ponta (kW):", min_value=0.0, value=0.0, step=1.0,
            format="%.2f", key="dcfp"
        )
        demanda_contratada_azul = st.number_input(
            "Demanda Contratada Ponta (kW):", min_value=0.0, value=0.0, step=1.0,
            format="%.2f", disabled=disable, key="dcp"
        )

        options_selectbox1 = estados
        estado_selecionado = st.selectbox("Selecione o estado", options_selectbox1, index=0)

        options_selectbox2 = selecionar_concessionaria(estado_selecionado)
        concessionaria_selecionada = st.selectbox("Selecione a concessionária", options_selectbox2, index=0)

        grupo_tarifario = st.radio(
            "Tipo de consumidor (subgrupo tarifário)",
            ["A1", "A2", "A3", "A4"],
            index=3,
            horizontal=True,
            key="grupo_tarifario",
            help=(
                "Subgrupo de tensão de fornecimento (A1: ≥230 kV, A2: 88-138 kV, "
                "A3: 69 kV, A4: 2,3-25 kV). Determina qual tarifa é buscada na "
                "planilha/API. A maioria das unidades consumidoras atendidas em "
                "média tensão está no A4."
            ),
        )

        icms = valor_ICMS(estado_selecionado)
        sigla_conc = definir_sigla(concessionaria_selecionada)

        demanda_fp_valor = 0

    with coluna2:
        st.subheader("Tarifas")

        origem_tarifas = st.radio(
            "Origem das tarifas",
            ["API da ANEEL (tempo real)", "Planilha de 2025", "Inserir manualmente"],
            key="origem_tarifas",
            horizontal=True,
        )

        if origem_tarifas in ("API da ANEEL (tempo real)", "Planilha de 2025"):
            fonte = "api" if origem_tarifas == "API da ANEEL (tempo real)" else "planilha"
            rotulo_fonte = "API da ANEEL" if fonte == "api" else "planilha local"

            if fonte == "api":
                st.caption(
                    "Busca as tarifas homologadas diretamente no portal de dados abertos da ANEEL "
                    "(consulta em tempo real, requer internet). Se a concessionária não for "
                    "encontrada, tente a planilha com dados de 2025 ou digite as tarifas manualmente."
                )
            else:
                st.caption(
                    "Usa a planilha de tarifas de 2025 já incluída no aplicativo, sem depender de "
                    "conexão com a internet."
                )

            if st.button("Importar tarifas :heavy_dollar_sign:", key="botao_tarifas"):
                with st.spinner(f"Importando tarifas ({rotulo_fonte})..."):
                    st.session_state["tarifas_verde"] = obter_tarifas("Verde", fonte=fonte)
                    st.session_state["tarifas_azul"] = obter_tarifas("Azul", fonte=fonte)
                    st.session_state["tarifas_fonte_importada"] = origem_tarifas

            if (
                "tarifas_verde" in st.session_state
                and "tarifas_azul" in st.session_state
                and st.session_state.get("tarifas_fonte_importada") == origem_tarifas
            ):
                tarifas_verde = st.session_state["tarifas_verde"]
                tarifas_azul = st.session_state["tarifas_azul"]

                m1, m2 = st.columns(2)
                m1.metric("Demanda Fora da Ponta (R$/kW)", f"{tarifas_verde[0]:.2f}")
                m2.metric("Demanda Ponta Azul (R$/kW)", f"{tarifas_azul[2]:.2f}")

                m3, m4, m5 = st.columns(3)
                m3.metric("Consumo Fora da Ponta (R$/kWh)", f"{tarifas_verde[4]:.2f}")
                m4.metric("Consumo Ponta Verde (R$/kWh)", f"{tarifas_verde[5]:.2f}")
                m5.metric("Consumo Ponta Azul (R$/kWh)", f"{tarifas_azul[5]:.2f}")
        else:
            st.caption(
                'Digite os valores de tarifa ("R$/kW" para demanda, "R$/kWh" para consumo) — use "." como '
                "separador decimal. Útil quando a concessionária não está na API/planilha da ANEEL ou "
                "quando você já tem as tarifas da fatura em mãos."
            )
            mt1, mt2 = st.columns(2)
            mt1.number_input("Demanda Fora da Ponta (R$/kW)", min_value=0.0, value=0.0, step=0.01,
                              format="%.4f", key="man_demanda_fp")
            mt2.number_input("Demanda Ponta Azul (R$/kW)", min_value=0.0, value=0.0, step=0.01,
                              format="%.4f", key="man_demanda_ponta_azul")

            mt3, mt4, mt5 = st.columns(3)
            mt3.number_input("Consumo Fora da Ponta (R$/kWh)", min_value=0.0, value=0.0, step=0.0001,
                              format="%.4f", key="man_consumo_fp")
            mt4.number_input("Consumo Ponta Verde (R$/kWh)", min_value=0.0, value=0.0, step=0.0001,
                              format="%.4f", key="man_consumo_ponta_verde")
            mt5.number_input("Consumo Ponta Azul (R$/kWh)", min_value=0.0, value=0.0, step=0.0001,
                              format="%.4f", key="man_consumo_ponta_azul")

        st.write("---")

        if st.button("Calcular gasto anual :dollar:"):
            with st.spinner("Calculando..."):
                custo_total, custo_demanda = custo_atual()
                st.session_state["gasto_anual_demanda"] = custo_demanda

        if "gasto_anual_demanda" in st.session_state:
            gasto_string = money_format(round(st.session_state["gasto_anual_demanda"], 2))
            st.metric("Gasto anual TUSD demanda", f"R$ {gasto_string}")


with tab_simulacao:
    st.header("Simulação e Resultados 📊 💰")

    with st.expander(("Como realizar as simulações")):
        st.markdown((
            """
            O programa realizará as simulações de demanda contratada de acordo com os dados de entrada, exibindo a demanda ótima a ser contratada, a economia anual obtida e os gráficos de otimização.

            Existem 3 possibilidades de simulação de acordo com o interesse do usuário:

            1. Simular Verde: Otimização apenas da demanda fora da ponta.

            2. Simular Azul: Otimização apenas da demanda na ponta.

            3. Simulação Completa: Otimização da demanda fora da ponta e na ponta. O programa realiza o cálculo do custo total na modalidade azul e na modalidade verde e exibe a melhor opção.
        """
        ))

    col_verde, col_azul, col_completo = st.columns(3)
    simular_verde = col_verde.button("Simular Verde 🟢", width='stretch')
    simular_azul = col_azul.button("Simular Azul 🔵", width='stretch')
    simular_completo = col_completo.button("Simular Completo ✅", type="primary",
                                            width='stretch')

    if simular_verde:
        with st.spinner("Calculando a demanda ótima (Modalidade Verde)..."):
            valor_otimo, demanda_otima_verde = varredura(30, limite_demanda, demanda_contratada_verde)
            custo_soma, custo_demanda = custo_atual()
            economia_verde = custo_demanda - valor_otimo
            st.session_state["resultado_verde"] = {
                "valor_otimo": valor_otimo,
                "demanda_otima": demanda_otima_verde,
                "economia": economia_verde,
                "demanda_fp": list(vetor_demanda_fp),
                "demanda_contratada": demanda_contratada_verde,
            }

    if "resultado_verde" in st.session_state:
        r = st.session_state["resultado_verde"]
        with st.container(border=True):
            st.subheader("Resultado: Modalidade Verde 🟢")
            c1, c2 = st.columns(2)
            c1.metric("Valor ótimo", f"R$ {money_format(round(r['valor_otimo'], 2))}")
            c2.metric("Demanda sugerida (fora da ponta)", f"{r['demanda_otima']} kW")
            st.metric("💰 Economia anual estimada", f"R$ {money_format(round(r['economia'], 2))}",
                      delta=round(r['economia'], 2))
            plotar_verde(r["demanda_contratada"], r["demanda_otima"], r["demanda_fp"])

    if simular_azul:
        with st.spinner("Calculando a demanda ótima (Modalidade Azul)..."):
            valor_otimo, demanda_otima_verde, valor_otimo_azul, demanda_otima_azul = \
                varredura_azul(30, limite_demanda, demanda_contratada_verde, demanda_contratada_azul)

            custo_soma, custo_demanda = custo_atual()
            economia_azul = custo_demanda - valor_otimo
            st.session_state["resultado_azul"] = {
                "valor_otimo_azul": valor_otimo_azul,
                "demanda_otima_azul": demanda_otima_azul,
                "economia": economia_azul,
                "demanda_ponta": list(vetor_demanda_ponta),
                "demanda_contratada_azul": demanda_contratada_azul,
            }

    if "resultado_azul" in st.session_state:
        r = st.session_state["resultado_azul"]
        with st.container(border=True):
            st.subheader("Resultado: Modalidade Azul 🔵")
            c1, c2 = st.columns(2)
            c1.metric("Valor ótimo", f"R$ {money_format(round(r['valor_otimo_azul'], 2))}")
            c2.metric("Demanda sugerida (ponta)", f"{r['demanda_otima_azul']} kW")
            st.metric("💰 Economia anual estimada", f"R$ {money_format(round(r['economia'], 2))}",
                      delta=round(r['economia'], 2))
            plotar_azul(r["demanda_contratada_azul"], r["demanda_otima_azul"], r["demanda_ponta"])

    if simular_completo:
        with st.spinner("Calculando comparação completa (Verde x Azul)..."):
            valor_otimo, demanda_otima_verde, valor_otimo_azul, demanda_otima_azul = \
                varredura_azul(30, limite_demanda, demanda_contratada_verde, demanda_contratada_azul)

            tarifas_verde = obter_tarifas_efetivas("Verde")
            total_verde, gasto_consumo_fp_verde, gasto_consumo_ponta_verde = \
                gastos_consumo(tarifas_verde, vetor_consumo_fp, vetor_consumo_ponta)
            custo_total_verde = valor_otimo + total_verde

            tarifas_azul = obter_tarifas_efetivas("Azul")
            total_azul, gasto_consumo_fp_azul, gasto_consumo_ponta_azul = \
                gastos_consumo(tarifas_azul, vetor_consumo_fp, vetor_consumo_ponta)
            custo_total_azul = valor_otimo_azul + valor_otimo + total_azul

            if custo_total_verde < custo_total_azul:
                modalidade_sugerida = "Verde"
                custo_otimo = custo_total_verde
                demanda_sugerida_ponta = "-"
            else:
                modalidade_sugerida = "Azul"
                custo_otimo = custo_total_azul
                demanda_sugerida_ponta = demanda_otima_azul

            custo_soma, custo_demanda = custo_atual()
            economia = custo_soma - custo_otimo

            st.session_state["resultado_completo"] = {
                "custo_soma": custo_soma,
                "custo_total_verde": custo_total_verde,
                "custo_total_azul": custo_total_azul,
                "demanda_otima_verde": demanda_otima_verde,
                "demanda_sugerida_ponta": demanda_sugerida_ponta,
                "modalidade_sugerida": modalidade_sugerida,
                "economia": economia,
                "demanda_otima_azul": demanda_otima_azul,
                "gasto_consumo_fp_verde": gasto_consumo_fp_verde,
                "gasto_consumo_ponta_verde": gasto_consumo_ponta_verde,
                "gasto_consumo_fp_azul": gasto_consumo_fp_azul,
                "gasto_consumo_ponta_azul": gasto_consumo_ponta_azul,
                "valor_otimo": valor_otimo,
                "valor_otimo_azul": valor_otimo_azul,
                "demanda_fp": list(vetor_demanda_fp),
                "demanda_ponta": list(vetor_demanda_ponta),
                "demanda_contratada_verde": demanda_contratada_verde,
                "demanda_contratada_azul": demanda_contratada_azul,
            }

    if "resultado_completo" in st.session_state:
        r = st.session_state["resultado_completo"]
        with st.container(border=True):
            st.subheader("Resultado: Comparação Completa ✅")

            c1, c2, c3 = st.columns(3)
            c1.metric("Custo Atual", f"R$ {money_format(round(r['custo_soma'], 2))}")
            c2.metric("Valor Total Verde", f"R$ {money_format(round(r['custo_total_verde'], 2))}")
            c3.metric("Valor Total Azul", f"R$ {money_format(round(r['custo_total_azul'], 2))}")

            c4, c5, c6 = st.columns(3)
            c4.metric("Demanda Ótima Fora da Ponta", f"{r['demanda_otima_verde']} kW")
            c5.metric("Demanda Ótima na Ponta", f"{r['demanda_sugerida_ponta']} kW")
            c6.metric("Modalidade Sugerida", r["modalidade_sugerida"])

            st.metric("💰 Economia anual estimada", f"R$ {money_format(round(r['economia'], 2))}",
                       delta=round(r['economia'], 2))

            plotar_verde(r["demanda_contratada_verde"], r["demanda_otima_verde"], r["demanda_fp"],
                         key="chart_completo_verde")
            plotar_azul(r["demanda_contratada_azul"], r["demanda_otima_azul"], r["demanda_ponta"],
                        key="chart_completo_azul")
            plotar_completo(r["valor_otimo"], r["valor_otimo_azul"], r["gasto_consumo_fp_verde"],
                             r["gasto_consumo_ponta_verde"], r["gasto_consumo_fp_azul"],
                             r["gasto_consumo_ponta_azul"], key="chart_completo_comparativo")

    tem_algum_resultado = (
        "resultado_verde" in st.session_state
        or "resultado_azul" in st.session_state
        or "resultado_completo" in st.session_state
    )

    if tem_algum_resultado:
        st.write("---")
        st.subheader("Relatório em PDF")
        st.caption(
            "Gera um PDF com os dados de entrada e os resultados das simulações já calculadas "
            "nesta sessão (Verde, Azul e/ou Completa), incluindo os gráficos."
        )

        if st.button("Gerar Relatório PDF 📄"):
            with st.spinner("Gerando PDF..."):
                st.session_state["relatorio_pdf"] = gerar_relatorio_pdf()

        if "relatorio_pdf" in st.session_state:
            nome_arquivo = f"relatorio_otimizacao_demanda_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
            st.download_button(
                "Baixar Relatório PDF ⬇️",
                data=st.session_state["relatorio_pdf"],
                file_name=nome_arquivo,
                mime="application/pdf",
            )