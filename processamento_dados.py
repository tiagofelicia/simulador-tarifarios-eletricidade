import streamlit as st
import pandas as pd
import datetime
from calendar import monthrange
import requests
import io

# --- Carregar ficheiro Excel do GitHub ---
# --- Para simulador de gás
@st.cache_data(ttl=1800, show_spinner=False) # Cache por 30 minutos (1800 segundos)
def carregar_dados_excel_gas(url):
    xls = pd.ExcelFile(url)
    try:
        tarifas_gas_master = xls.parse("Tarifas_Gas_Master")
    except Exception as e:
        # Se a aba não existir, criamos um DataFrame vazio para não quebrar a app
        st.error(f"Atenção: A aba 'Tarifas_Gas_Master' não foi encontrada no Excel. {e}")
        tarifas_gas_master = pd.DataFrame()
    try:
        tos_municipios = xls.parse("TOS")
        # --- Limpar nomes das colunas da aba TOS ---
        if not tos_municipios.empty:
             tos_municipios.columns = [str(c).strip() for c in tos_municipios.columns]
    except Exception:
        st.error("Aviso: A aba 'TOS' (Taxa Ocupação Subsolo) não foi encontrada no Excel.")
        tos_municipios = pd.DataFrame()
    try:
        mibgas_df = xls.parse("MIBGAS")
    except Exception:
        st.warning("Aviso: A aba 'MIBGAS' não foi encontrada no Excel.")
        mibgas_df = pd.DataFrame()
    try:
        info_tab = xls.parse("Info")
    except Exception:
        st.warning("Aviso: A aba 'Info' não foi encontrada no Excel.")
        info_tab = pd.DataFrame()

    constantes = xls.parse("Constantes")
    return constantes, tarifas_gas_master, tos_municipios, mibgas_df, info_tab

# --- Carregar ficheiro Excel do GitHub ---
# --- Para simulador de eletricidade
@st.cache_data(ttl=1800, show_spinner=False)
def carregar_dados_excel_elec(url):
    """
    Carrega os dados do ficheiro Excel de eletricidade a partir de um URL.
    Usa 'requests' para descarregar e o motor 'calamine' para ler de forma robusta.
    """
    try:
        response = requests.get(url)
        response.raise_for_status()
        excel_file_in_memory = io.BytesIO(response.content)
        xls = pd.ExcelFile(excel_file_in_memory, engine='calamine')
        
    except requests.exceptions.RequestException as e:
        st.error(f"Erro ao descarregar o ficheiro Excel do GitHub: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    tarifarios_fixos = xls.parse("Tarifarios_fixos")
    tarifarios_indexados = xls.parse("Indexados")
    omie_perdas_ciclos = xls.parse("OMIE_PERDAS_CICLOS")

    omie_perdas_ciclos.columns = [str(c).strip() for c in omie_perdas_ciclos.columns]
    
    if 'Data' not in omie_perdas_ciclos.columns and 'DataHora' in omie_perdas_ciclos.columns:
        temp_dt = pd.to_datetime(omie_perdas_ciclos['DataHora'])
        omie_perdas_ciclos['Data'] = temp_dt.dt.strftime('%m/%d/%Y')
        omie_perdas_ciclos['Hora'] = temp_dt.dt.strftime('%H:%M')

    if 'Data' in omie_perdas_ciclos.columns and 'Hora' in omie_perdas_ciclos.columns:
        omie_perdas_ciclos['DataHora'] = pd.to_datetime(
            omie_perdas_ciclos['Data'].astype(str) + ' ' + omie_perdas_ciclos['Hora'].astype(str),
            format='%m/%d/%Y %H:%M',
            errors='coerce'
        ).dt.tz_localize(None)
        
        omie_perdas_ciclos.dropna(subset=['DataHora'], inplace=True)
        omie_perdas_ciclos.drop_duplicates(subset=['DataHora'], keep='first', inplace=True)
    else:
        st.error("Colunas 'Data' e 'Hora' não encontradas na aba OMIE_PERDAS_CICLOS.")

    constantes = xls.parse("Constantes")
    return tarifarios_fixos, tarifarios_indexados, omie_perdas_ciclos, constantes



def processar_ficheiro_consumos(ficheiro_excel):
    """
    Lê um ficheiro Excel da E-Redes, com deteção de cabeçalho e ajuste de tempo preciso
    para alinhar com os timestamps do ficheiro OMIE, aplicando a regra de negócio para 00:00.
    Agora suporta múltiplos nomes para a coluna de consumo e potência.
    """
    try:
        df_temp = pd.read_excel(ficheiro_excel, header=None, nrows=20)
        header_row_index = -1
        coluna_consumo_kw = ""
        
        # --- Procurar por uma lista de colunas na ordem desejada ---
        colunas_procurar_consumo = [
            'Consumo Simulado (kW)',
            "Consumo medido na IC, Ativa (kW)",
            "Consumo registado (kW)",
            "Consumo registado, Ativa (kW)"
        ]

        for i, row in df_temp.iterrows():
            row_values = [str(v).strip() for v in row.values]
            for nome_coluna in colunas_procurar_consumo:
                if nome_coluna in row_values:
                    header_row_index = i
                    coluna_consumo_kw = nome_coluna
                    break  # Sai do loop interno assim que encontra uma correspondência
            if coluna_consumo_kw:
                break # Sai do loop externo se já encontrou a coluna

        if header_row_index == -1:
            return None, "Não foi possível encontrar uma linha de cabeçalho com colunas de consumo conhecidas."

        df = pd.read_excel(ficheiro_excel, header=header_row_index)
        df.columns = [str(c).strip() for c in df.columns]
        
        df['Consumo (kWh)'] = pd.to_numeric(df[coluna_consumo_kw], errors='coerce') / 4.0

        # --- Lógica para definir a Potencia_kW_Para_Analise ---
        if "Consumo registado, Ativa (kW)" in df.columns:
            df['Potencia_kW_Para_Analise'] = pd.to_numeric(df["Consumo registado, Ativa (kW)"], errors='coerce')
        elif "Consumo registado (kW)" in df.columns:
            df['Potencia_kW_Para_Analise'] = pd.to_numeric(df["Consumo registado (kW)"], errors='coerce')
        else:
            # Fallback para a coluna de consumo principal, caso as outras não existam
            df['Potencia_kW_Para_Analise'] = pd.to_numeric(df[coluna_consumo_kw], errors='coerce')

        df.dropna(subset=[coluna_consumo_kw], inplace=True)

        df['DataHora'] = pd.to_datetime(
            df['Data'].astype(str) + ' ' + df['Hora'].astype(str),
            errors='coerce'
        ).dt.tz_localize(None)

        # Ajuste para o timestamp 00:00 (lógica existente mantida)
        df['DataHora'] = df['DataHora'].apply(
            lambda ts: ts - pd.Timedelta(minutes=1) if ts.time() == datetime.time(0, 0) else ts
        )
        
        df.dropna(subset=['DataHora', 'Consumo (kWh)'], inplace=True)

        return df[['DataHora', 'Consumo (kWh)', 'Potencia_kW_Para_Analise']], None
    except Exception as e:
        return None, f"Erro ao processar ficheiro: {e}"

def validar_e_juntar_ficheiros(lista_de_ficheiros):
    """
    Processa uma lista de ficheiros da E-Redes, junta os dados, e filtra para incluir
    apenas registos a partir de 01/01/2025, alertando o utilizador se dados mais
    antigos foram ignorados (Lógica Robusta).
    """
    if not lista_de_ficheiros:
        return None, "Nenhum ficheiro carregado."

    dataframes_processados = []
    intervalos_de_datas = []
    
    data_limite_dt = pd.to_datetime('2025-01-01')
    dados_antigos_encontrados = False # Flag para o aviso

    for ficheiro in lista_de_ficheiros:
        df_individual, erro = processar_ficheiro_consumos(ficheiro)
        if erro:
            return None, f"Erro ao processar o ficheiro '{ficheiro.name}': {erro}"
        
        if df_individual.empty:
            continue

        # --- ALTERAÇÃO PRINCIPAL: Lógica de deteção por contagem de linhas ---
        
        # 1. Contar linhas ANTES de filtrar
        linhas_antes = len(df_individual)
        
        # 2. Aplicar o filtro de data
        df_filtrado = df_individual[df_individual['DataHora'] >= data_limite_dt].copy()
        
        # 3. Contar linhas DEPOIS de filtrar
        linhas_depois = len(df_filtrado)
        
        # 4. Se o número de linhas diminuiu, sabemos que dados antigos foram ignorados.
        if linhas_antes > linhas_depois:
            dados_antigos_encontrados = True

        # Se o ficheiro ficar vazio após a filtragem, simplesmente ignoramo-lo.
        if df_filtrado.empty:
            continue

        # A partir daqui, trabalhamos apenas com o df_filtrado
        dataframes_processados.append(df_filtrado)
        min_data = df_filtrado['DataHora'].min()
        max_data = df_filtrado['DataHora'].max()
        intervalos_de_datas.append((min_data, max_data))

    if not dataframes_processados:
        return None, "Nenhum dos ficheiros continha dados válidos a partir de 01/01/2025."

    # Lógica de verificação de sobreposição (mantém-se igual)
    if len(intervalos_de_datas) > 1:
        intervalos_ordenados = sorted(intervalos_de_datas, key=lambda x: x[0])
        for i in range(1, len(intervalos_ordenados)):
            if intervalos_ordenados[i][0] < intervalos_ordenados[i-1][1]:
                return None, "Erro: Sobreposição de datas detetada entre os ficheiros."

    df_final_combinado = pd.concat(dataframes_processados, ignore_index=True)
    df_final_combinado = df_final_combinado.sort_values(by='DataHora').reset_index(drop=True)
    df_final_combinado = df_final_combinado.drop_duplicates(subset=['DataHora'], keep='first')

    # Lógica de retorno da mensagem (mantém-se igual)
    mensagem_retorno = None
    if dados_antigos_encontrados:
        mensagem_retorno = "Aviso: Foram encontrados e ignorados dados anteriores a 01/01/2025."

    return df_final_combinado, mensagem_retorno

def agregar_consumos_por_periodo(df_consumos, df_omie_ciclos):
    if df_consumos is None or df_consumos.empty: return {}

    df_merged = pd.merge(df_consumos, df_omie_ciclos, on='DataHora', how='left')

    consumos_agregados = {'Simples': df_merged['Consumo (kWh)'].sum()}
    
    for ciclo in ['BD', 'BS', 'TD', 'TS']:
        if ciclo in df_merged.columns:
            df_merged[ciclo] = df_merged[ciclo].fillna('Desconhecido')
            soma_por_periodo = df_merged.groupby(ciclo)['Consumo (kWh)'].sum().to_dict()
            consumos_agregados[ciclo] = soma_por_periodo
            
    return consumos_agregados

def calcular_medias_omie_para_todos_ciclos(df_consumos_periodo, df_omie_completo):
    """
    Calcula as médias OMIE para todos os ciclos, com base no intervalo de datas
    do dataframe de consumos fornecido.
    """
    if df_consumos_periodo.empty:
        return {}
    
    min_date = df_consumos_periodo['DataHora'].min()
    max_date = df_consumos_periodo['DataHora'].max()
    
    df_omie_filtrado = df_omie_completo[
        (df_omie_completo['DataHora'] >= min_date) & 
        (df_omie_completo['DataHora'] <= max_date)
    ].copy()

    if df_omie_filtrado.empty:
        return {}

    omie_medios = {'S': df_omie_filtrado['OMIE'].mean()}
    for ciclo in ['BD', 'BS', 'TD', 'TS']:
        if ciclo in df_omie_filtrado.columns:
            agrupado = df_omie_filtrado.groupby(ciclo)['OMIE'].mean()
            for periodo, media in agrupado.items():
                omie_medios[f"{ciclo}_{periodo}"] = media
    return omie_medios

def normalizar_para_ordenacao(texto):
    """
    Remove acentos de um texto e converte para minúsculas para criar uma
    chave de ordenação alfabética que funciona de forma consistente.
    Ex: 'Évora' -> 'evora'
    """
    if not isinstance(texto, str):
        return texto

    # Dicionário de substituição de caracteres
    substituicoes = {
        'á': 'a', 'à': 'a', 'ã': 'a', 'â': 'a',
        'é': 'e', 'ê': 'e',
        'í': 'i',
        'ó': 'o', 'ô': 'o', 'õ': 'o',
        'ú': 'u', 'ü': 'u',
        'ç': 'c',
    }
    
    texto_lower = texto.lower()
    texto_normalizado = ""
    for char in texto_lower:
        texto_normalizado += substituicoes.get(char, char)
        
    return texto_normalizado