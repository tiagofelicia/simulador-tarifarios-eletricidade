import streamlit as st
import pandas as pd
import datetime
import io
import json
import time
import re

from st_aggrid import AgGrid, GridOptionsBuilder
from st_aggrid.shared import GridUpdateMode, JsCode
from bs4 import BeautifulSoup # Para processar o resumo HTML
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill # Para formatação Excel
from openpyxl.utils import get_column_letter # Para nomes de colunas Excel
from calendar import monthrange

st.set_page_config(page_title="Simulador de Tarifários de Eletricidade", layout="wide")

# --- Carregar ficheiro Excel do GitHub ---
@st.cache_data(ttl=1800, show_spinner=False) # Cache por 30 minutos (1800 segundos)
def carregar_dados_excel(url):
    xls = pd.ExcelFile(url)
    tarifarios_fixos = xls.parse("Tarifarios_fixos")
    tarifarios_indexados = xls.parse("Indexados")
    omie_perdas_ciclos = xls.parse("OMIE_PERDAS_CICLOS")
    # Limpar nomes das colunas em OMIE_PERDAS_CICLOS
    omie_perdas_ciclos.columns = [str(c).strip() for c in omie_perdas_ciclos.columns]
    if 'DataHora' not in omie_perdas_ciclos.columns and 'Data' in omie_perdas_ciclos.columns:
        omie_perdas_ciclos.rename(columns={'Data': 'DataHora'}, inplace=True)
    if 'DataHora' in omie_perdas_ciclos.columns:
        omie_perdas_ciclos['DataHora'] = pd.to_datetime(omie_perdas_ciclos['DataHora'])
    else:
        st.error("Coluna 'DataHora' ou 'Data' não encontrada na aba OMIE_PERDAS_CICLOS.")
        omie_perdas_ciclos['DataHora'] = pd.Series(dtype='datetime64[ns]')

    constantes = xls.parse("Constantes")
    return tarifarios_fixos, tarifarios_indexados, omie_perdas_ciclos, constantes

url_excel = "https://github.com/tiago1978/simulador-tarifarios-eletricidade/raw/refs/heads/main/TiagoFelicia_Simulador_Eletricidade.xlsx"
tarifarios_fixos, tarifarios_indexados, OMIE_PERDAS_CICLOS, CONSTANTES = carregar_dados_excel(url_excel)


# --- BLOCO 2 ---

# --- Função para obter valores da aba Constantes ---
def obter_constante(nome_constante, constantes_df):
    constante_row = constantes_df[constantes_df['constante'] == nome_constante]
    if not constante_row.empty:
        valor = constante_row['valor_unitário'].iloc[0]
        try:
            return float(valor)
        except (ValueError, TypeError):
            # st.warning(f"Valor não numérico para constante '{nome_constante}': {valor}")
            return 0.0
    else:
        # st.warning(f"Constante '{nome_constante}' não encontrada.")
        return 0.0

# --- Função para obter valor da TAR energia por período ---
def obter_tar_energia_periodo(opcao_horaria_str, periodo_str, potencia_kva, constantes_df):
    nome_constante = ""
    opcao_lower = str(opcao_horaria_str).lower()
    periodo_upper = str(periodo_str).upper()

    if opcao_lower == "simples": nome_constante = "TAR_Energia_Simples"
    elif opcao_lower.startswith("bi"):
        if periodo_upper == 'V': nome_constante = "TAR_Energia_Bi_Vazio"
        elif periodo_upper == 'F': nome_constante = "TAR_Energia_Bi_ForaVazio"
    elif opcao_lower.startswith("tri"):
        if potencia_kva <= 20.7:
            if periodo_upper == 'V': nome_constante = "TAR_Energia_Tri_Vazio"
            elif periodo_upper == 'C': nome_constante = "TAR_Energia_Tri_Cheias"
            elif periodo_upper == 'P': nome_constante = "TAR_Energia_Tri_Ponta"
        else: # > 20.7 kVA
            if periodo_upper == 'V': nome_constante = "TAR_Energia_Tri_27.6_Vazio"
            elif periodo_upper == 'C': nome_constante = "TAR_Energia_Tri_27.6_Cheias"
            elif periodo_upper == 'P': nome_constante = "TAR_Energia_Tri_27.6_Ponta"

    if nome_constante:
        return obter_constante(nome_constante, constantes_df)
    return 0.0

# --- Função para REINICIAR o simulador para os valores padrão ---
def reiniciar_simulador():
    """
    Repõe explicitamente todos os valores de session_state para os seus defaults.
    Este método é mais robusto do que um simples .clear().
    """
    
    # 1. Repor os inputs principais
    potencias_validas = [1.15, 2.3, 3.45, 4.6, 5.75, 6.9, 10.35, 13.8, 17.25, 20.7, 27.6, 34.5, 41.4]
    st.session_state.sel_potencia = 3.45  # Default da potência
    st.session_state.sel_opcao_horaria = "Simples" # Default da opção horária
    
    meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    mes_atual_idx = datetime.datetime.now().month - 1
    st.session_state.sel_mes = meses[mes_atual_idx] # Default do mês
    
    # Apagar as chaves das datas para que elas se recalculem com base no mês
    for key in ['data_inicio_val', 'data_fim_val', 'dias_manual_val', 'session_initialized_dates', 'previous_mes_for_dates']:
        if key in st.session_state:
            del st.session_state[key]

    # 2. Repor os inputs de consumo para os valores padrão
    st.session_state.exp_consumo_s = "158"
    st.session_state.exp_consumo_v = "63"
    st.session_state.exp_consumo_f = "95"
    st.session_state.exp_consumo_c = "68"
    st.session_state.exp_consumo_p = "27"
    
    # 3. Repor os inputs OMIE (apagar para que usem os defaults calculados)
    for key in ['omie_s_input_field', 'omie_v_input_field', 'omie_f_input_field', 'omie_c_input_field', 'omie_p_input_field', 'omie_foi_editado_manualmente']:
        if key in st.session_state:
            del st.session_state[key]

    # 4. Repor as checkboxes de opções adicionais
    st.session_state.chk_tarifa_social = False
    st.session_state.chk_tarifa_social_val = False
    st.session_state.chk_familia_numerosa = False
    st.session_state.chk_familia_numerosa_val = False
    st.session_state.chk_acp = True
    st.session_state.incluir_quota_acp_val = True
    st.session_state.chk_continente = True
    st.session_state.desconto_continente_val = True
    st.session_state.chk_modo_comparacao_opcoes = False

    # 5. Repor o Meu Tarifário
    st.session_state.chk_meu_tarifario_ativo = False
    if 'meu_tarifario_calculado' in st.session_state:
        del st.session_state['meu_tarifario_calculado']
    
    chaves_meu_tarifario_reset = [
        "energia_meu_s_input_val", "potencia_meu_input_val", "energia_meu_v_input_val",
        "energia_meu_f_input_val", "energia_meu_c_input_val", "energia_meu_p_input_val",
        "meu_tar_energia_val", "meu_tar_potencia_val", "meu_fin_tse_incluido_val",
        "meu_desconto_energia_val", "meu_desconto_potencia_val", "meu_desconto_fatura_val",
        "meu_acrescimo_fatura_val"
    ]
    for key in chaves_meu_tarifario_reset:
        if key in st.session_state:
            del st.session_state[key]
            
    # 6. Repor os filtros da tabela
    st.session_state.filter_segmento_selectbox = "Residencial"
    st.session_state.filter_tipos_multi = []
    st.session_state.filter_faturacao_selectbox = "Todas"
    st.session_state.filter_pagamento_selectbox = "Todos"

# --- Função: Obter valor da TAR potência para a potência contratada ---
def obter_tar_dia(potencia_kva, constantes_df):
    potencia_str = str(float(potencia_kva)) # Formato consistente
    constante_potencia = f'TAR_Potencia {potencia_str}'
    return obter_constante(constante_potencia, constantes_df)

# --- Obter valor constante do Financiamento TSE ---
FINANCIAMENTO_TSE_VAL = obter_constante("Financiamento_TSE", CONSTANTES)

# --- Obter valor constante da Quota ACP ---
VALOR_QUOTA_ACP_MENSAL = obter_constante("Quota_ACP", CONSTANTES)

# --- Função: Determinar o perfil BTN ---
def obter_perfil(consumo_total_kwh, dias, potencia_kva):
    consumo_anual_estimado = consumo_total_kwh * 365 / dias if dias > 0 else consumo_total_kwh
    if potencia_kva > 13.8: return 'perfil_A'
    elif consumo_anual_estimado > 7140: return 'perfil_B'
    else: return 'perfil_C'

# --- Função: Calcular custo de energia com IVA (limite 200 ou 300 kWh/30 dias apenas <= 6.9 kVA), para diferentes opções horárias
def calcular_custo_energia_com_iva(
    consumo_kwh_total_periodo, preco_energia_final_sem_iva_simples,
    precos_energia_final_sem_iva_horario, dias_calculo, potencia_kva,
    opcao_horaria_str, consumos_horarios, familia_numerosa_bool
):
    if not isinstance(opcao_horaria_str, str):
        return {'custo_com_iva': 0.0, 'custo_sem_iva': 0.0, 'valor_iva_6': 0.0, 'valor_iva_23': 0.0}

    opcao_horaria_lower = opcao_horaria_str.lower()
    iva_normal_perc = 0.23
    iva_reduzido_perc = 0.06
    
    custo_total_com_iva = 0.0
    custo_total_sem_iva = 0.0
    total_iva_6_energia = 0.0
    total_iva_23_energia = 0.0

    precos_horarios = precos_energia_final_sem_iva_horario if isinstance(precos_energia_final_sem_iva_horario, dict) else {}
    consumos_periodos = consumos_horarios if isinstance(consumos_horarios, dict) else {}

    # Calcular custo total sem IVA primeiro
    if opcao_horaria_lower == "simples":
        consumo_s = float(consumos_periodos.get('S', 0.0) or 0.0)
        preco_s = float(preco_energia_final_sem_iva_simples or 0.0)
        custo_total_sem_iva = consumo_s * preco_s
    else: # Bi ou Tri
        for periodo, consumo_p in consumos_periodos.items():
            consumo_p_float = float(consumo_p or 0.0)
            preco_h = float(precos_horarios.get(periodo, 0.0) or 0.0)
            custo_total_sem_iva += consumo_p_float * preco_h
            
    # Determinar limite para IVA reduzido
    limite_kwh_periodo_global = 0.0
    if potencia_kva <= 6.9:
        limite_kwh_mensal = 300 if familia_numerosa_bool else 200
        limite_kwh_periodo_global = (limite_kwh_mensal * dias_calculo / 30.0) if dias_calculo > 0 else 0.0

    if limite_kwh_periodo_global == 0.0: # Sem IVA reduzido, tudo a 23%
        total_iva_23_energia = custo_total_sem_iva * iva_normal_perc
        custo_total_com_iva = custo_total_sem_iva + total_iva_23_energia
    else: # Com IVA reduzido/Normal
        if opcao_horaria_lower == "simples":
            consumo_s = float(consumos_periodos.get('S', 0.0) or 0.0)
            preco_s = float(preco_energia_final_sem_iva_simples or 0.0)
            
            consumo_para_iva_reduzido = min(consumo_s, limite_kwh_periodo_global)
            consumo_para_iva_normal = max(0.0, consumo_s - limite_kwh_periodo_global)
            
            base_iva_6 = consumo_para_iva_reduzido * preco_s
            base_iva_23 = consumo_para_iva_normal * preco_s
            
            total_iva_6_energia = base_iva_6 * iva_reduzido_perc
            total_iva_23_energia = base_iva_23 * iva_normal_perc
            custo_total_com_iva = base_iva_6 + total_iva_6_energia + base_iva_23 + total_iva_23_energia
        else: # Bi ou Tri rateado
            consumo_total_real_periodos = sum(float(v or 0.0) for v in consumos_periodos.values())
            if consumo_total_real_periodos > 0:
                for periodo, consumo_periodo in consumos_periodos.items():
                    consumo_periodo_float = float(consumo_periodo or 0.0)
                    preco_periodo = float(precos_horarios.get(periodo, 0.0) or 0.0)
                    
                    fracao_consumo_periodo = consumo_periodo_float / consumo_total_real_periodos
                    limite_para_este_periodo_rateado = limite_kwh_periodo_global * fracao_consumo_periodo
                    
                    consumo_periodo_iva_reduzido = min(consumo_periodo_float, limite_para_este_periodo_rateado)
                    consumo_periodo_iva_normal = max(0.0, consumo_periodo_float - limite_para_este_periodo_rateado)
                    
                    base_periodo_iva_6 = consumo_periodo_iva_reduzido * preco_periodo
                    base_periodo_iva_23 = consumo_periodo_iva_normal * preco_periodo
                    
                    iva_6_este_periodo = base_periodo_iva_6 * iva_reduzido_perc
                    iva_23_este_periodo = base_periodo_iva_23 * iva_normal_perc
                    
                    total_iva_6_energia += iva_6_este_periodo
                    total_iva_23_energia += iva_23_este_periodo
                    custo_total_com_iva += base_periodo_iva_6 + iva_6_este_periodo + base_periodo_iva_23 + iva_23_este_periodo
            else: # Se consumo_total_real_periodos for 0, tudo é zero
                 custo_total_com_iva = 0.0
                 # total_iva_6_energia e total_iva_23_energia permanecem 0.0

    return {
        'custo_com_iva': round(custo_total_com_iva, 4),
        'custo_sem_iva': round(custo_total_sem_iva, 4),
        'valor_iva_6': round(total_iva_6_energia, 4),
        'valor_iva_23': round(total_iva_23_energia, 4)
    }

# --- Função: Calcular custo da potência com IVA ---
def calcular_custo_potencia_com_iva_final(preco_comercializador_dia_sem_iva, tar_potencia_final_dia_sem_iva, dias, potencia_kva):
    iva_normal_perc = 0.23
    iva_reduzido_perc = 0.06
    
    preco_comercializador_dia_sem_iva = float(preco_comercializador_dia_sem_iva or 0.0)
    tar_potencia_final_dia_sem_iva = float(tar_potencia_final_dia_sem_iva or 0.0) # Esta TAR já tem TS, se aplicável
    dias = int(dias or 0)

    if dias <= 0:
        return {'custo_com_iva': 0.0, 'custo_sem_iva': 0.0, 'valor_iva_6': 0.0, 'valor_iva_23': 0.0}

    custo_comerc_siva_periodo = preco_comercializador_dia_sem_iva * dias
    custo_tar_siva_periodo = tar_potencia_final_dia_sem_iva * dias
    custo_total_potencia_siva = custo_comerc_siva_periodo + custo_tar_siva_periodo

    iva_6_pot = 0.0
    iva_23_pot = 0.0
    custo_total_com_iva = 0.0

    # Aplicar IVA separado: 23% no comercializador, 6% na TAR final
    if potencia_kva <= 3.45:
        iva_23_pot = custo_comerc_siva_periodo * iva_normal_perc
        iva_6_pot = custo_tar_siva_periodo * iva_reduzido_perc
        custo_total_com_iva = (custo_comerc_siva_periodo + iva_23_pot) + (custo_tar_siva_periodo + iva_6_pot)
    else: # potencia_kva > 3.45
    # Aplicar IVA normal (23%) à soma das componentes finais
        iva_23_pot = custo_total_potencia_siva * iva_normal_perc
        custo_total_com_iva = custo_total_potencia_siva + iva_23_pot
        # iva_6_pot permanece 0.0

    return {
        'custo_com_iva': round(custo_total_com_iva, 4),
        'custo_sem_iva': round(custo_total_potencia_siva, 4),
        'valor_iva_6': round(iva_6_pot, 4),
        'valor_iva_23': round(iva_23_pot, 4)
    }
    return round(custo_total_com_iva, 4)

IDENTIFICADORES_COMERCIALIZADORES_CAV_FIXA = [
    "CUR",
    "EDP",
    "Galp",
    "Goldenergy",
    "Ibelectra",
    "Iberdrola",
    "Luzigás",
    "Plenitude",
    "YesEnergy"     
    # Adicionar outros identificadores conforme necessário
]

def inicializar_estado():
    """
    Verifica e inicializa os valores no st.session_state se ainda não existirem.
    Isto garante que os valores padrão são definidos apenas uma vez por sessão.
    """
    # Valores padrão para os seletores principais
    if "sel_potencia" not in st.session_state:
        st.session_state.sel_potencia = 3.45

    if "sel_opcao_horaria" not in st.session_state:
        st.session_state.sel_opcao_horaria = "Simples"

    if "sel_mes" not in st.session_state:
        meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
        mes_atual_idx = datetime.datetime.now().month - 1
        st.session_state.sel_mes = meses[mes_atual_idx]
    
    # Pode adicionar outras inicializações aqui se necessário
    if "chk_tarifa_social_val" not in st.session_state:
        st.session_state.chk_tarifa_social_val = False
    
    if "chk_familia_numerosa_val" not in st.session_state:
        st.session_state.chk_familia_numerosa_val = False

inicializar_estado()

# --- Função: Calcular taxas adicionais ---
def calcular_taxas_adicionais(
    consumo_kwh,
    dias_simulacao,           # O número de dias para o cálculo da simulação
    tarifa_social_bool,
    valor_dgeg_mensal,
    valor_cav_mensal,   # O valor da CAV mensal (ex: 2.85€)
    nome_comercializador_atual, # Nome do comercializador do tarifário em análise
    mes_selecionado_simulacao,  # String do mês da simulação (ex: "Janeiro")
    ano_simulacao_atual,      # Inteiro do ano da simulação (ex: 2025)
    valor_iec=0.001           # Valor do IEC por kWh
):
    """
    Calcula taxas adicionais (IEC, DGEG, CAV) com lógica ajustada para CAV mensal fixa
    para comercializadores específicos em meses civis completos.
    """
    iva_normal_perc = 0.23
    iva_reduzido_perc = 0.06 # IVA da CAV é 6%

    consumo_kwh_float = float(consumo_kwh or 0.0)
    dias_simulacao_int = int(dias_simulacao or 0)
    valor_dgeg_mensal_float = float(valor_dgeg_mensal or 0.0)
    valor_cav_mensal = float(valor_cav_mensal or 0.0)
    valor_iec_float = float(valor_iec or 0.0)

    if dias_simulacao_int <= 0:
        return {
            'custo_com_iva': 0.0, 'custo_sem_iva': 0.0,
            'iec_sem_iva': 0.0, 'dgeg_sem_iva': 0.0, 'cav_sem_iva': 0.0,
            'valor_iva_6': 0.0, 'valor_iva_23': 0.0
        }

    # Custos Sem IVA
    # IEC (Imposto Especial de Consumo)
    iec_siva = 0.0 if tarifa_social_bool else (consumo_kwh_float * valor_iec_float)

    # DGEG (Taxa de Exploração da Direção-Geral de Energia e Geologia) - sempre proporcional
    dgeg_siva = (valor_dgeg_mensal_float * 12 / 365.25 * dias_simulacao_int)
    cav_siva = 0.0
    
    # Obter o número de dias no mês civil selecionado para a simulação
    # Utiliza o dicionário 'dias_mes' que já existe no seu script.
    # (Assumindo que 'dias_mes' está acessível aqui ou é passado como argumento/global)
    dias_no_mes_civil_selecionado = dias_mes.get(mes_selecionado_simulacao, 30) # Default 30 se mês não encontrado
    if mes_selecionado_simulacao == "Fevereiro" and \
       ((ano_simulacao_atual % 4 == 0 and ano_simulacao_atual % 100 != 0) or \
        (ano_simulacao_atual % 400 == 0)):
        dias_no_mes_civil_selecionado = 29

    aplica_cav_fixa_mensal = False
    if nome_comercializador_atual and isinstance(nome_comercializador_atual, str):
        nome_comerc_lower_para_verificacao = nome_comercializador_atual.lower()
        # Verifica se algum dos identificadores está contido no nome do comercializador
        if any(identificador.lower() in nome_comerc_lower_para_verificacao for identificador in IDENTIFICADORES_COMERCIALIZADORES_CAV_FIXA):
            # Se for um comercializador da lista, verifica se os dias da simulação correspondem ao mês completo
            if dias_simulacao_int == dias_no_mes_civil_selecionado:
                aplica_cav_fixa_mensal = True
    
    if aplica_cav_fixa_mensal:
        cav_siva = valor_cav_mensal  # Aplica o valor mensal total da CAV
    else:
        # Cálculo proporcional padrão para os outros casos
        cav_siva = (valor_cav_mensal * 12 / 365.25 * dias_simulacao_int)

    # Valores de IVA
    iva_iec = 0.0 if tarifa_social_bool else (iec_siva * iva_normal_perc)
    iva_dgeg = dgeg_siva * iva_normal_perc
    iva_cav = cav_siva * iva_reduzido_perc

    custo_total_siva = iec_siva + dgeg_siva + cav_siva
    custo_total_com_iva = (iec_siva + iva_iec) + (dgeg_siva + iva_dgeg) + (cav_siva + iva_cav)
    
    total_iva_6_calculado = iva_cav
    total_iva_23_calculado = iva_iec + iva_dgeg

    return {
        'custo_com_iva': round(custo_total_com_iva, 4), # Custo total das taxas com IVA
        'custo_sem_iva': round(custo_total_siva, 4),    # Custo total das taxas sem IVA
        'iec_sem_iva': round(iec_siva, 4),
        'dgeg_sem_iva': round(dgeg_siva, 4),
        'cav_sem_iva': round(cav_siva, 4),
        'valor_iva_6': round(total_iva_6_calculado, 4),
        'valor_iva_23': round(total_iva_23_calculado, 4)
    }

# --- Inicializar lista de resultados ---
resultados_list = []

# --- Título e Botão de Limpeza Geral (Layout Revisto) ---

# Linha 1: Logo e Título
col_logo, col_titulo = st.columns([1, 5])

with col_logo:
    st.image("https://raw.githubusercontent.com/tiagofelicia/simulador-tarifarios-eletricidade/refs/heads/main/Logo_Tiago_Felicia.png", width=180)

with col_titulo:
    st.title("🔌 Tiago Felícia - Simulador de Tarifários de Eletricidade")

# Linha 2: Botão de limpeza a ocupar a largura total
st.button(
    "🧹 Limpar e Reiniciar Simulador", 
    on_click=reiniciar_simulador,
    help="Repõe todos os campos do simulador para os valores iniciais.", 
    use_container_width=True
)

# --- Inputs principais ---
# ... (Potência, Opção Horária, Mês) ...
potencias_validas = [1.15, 2.3, 3.45, 4.6, 5.75, 6.9, 10.35, 13.8, 17.25, 20.7, 27.6, 34.5, 41.4]
opcoes_horarias_existentes = list(tarifarios_fixos['opcao_horaria_e_ciclo'].dropna().unique())
if "Simples" not in opcoes_horarias_existentes:
    opcoes_horarias = ["Simples"] + sorted(opcoes_horarias_existentes)
else:
    opcoes_horarias = sorted(opcoes_horarias_existentes)

col1, col2, col3 = st.columns(3)
with col1:
    potencia = st.selectbox("Potência Contratada (kVA)", potencias_validas, key="sel_potencia", help="Potências BTN (1.15 kVA a 41.4 kVA)")

if potencia in [27.6, 34.5, 41.4]:
    opcoes_validas = [o for o in opcoes_horarias if o.startswith("Tri-horário > 20.7 kVA")]
    if not opcoes_validas: opcoes_validas = [o for o in opcoes_horarias if "Tri-horário" in o]
    if not opcoes_validas and "Simples" in opcoes_horarias: opcoes_validas = ["Simples"]
    elif not opcoes_validas: opcoes_validas = opcoes_horarias[:1] if opcoes_horarias else ["Simples"]
else:
    opcoes_validas = [o for o in opcoes_horarias if not o.startswith("Tri-horário > 20.7 kVA")]
    if not opcoes_validas and "Simples" in opcoes_horarias : opcoes_validas = ["Simples"]
    elif not opcoes_validas : opcoes_validas = opcoes_horarias[:1] if opcoes_horarias else ["Simples"]

with col2:
    default_opcao_idx = 0
    if "Simples" in opcoes_validas: default_opcao_idx = opcoes_validas.index("Simples")
    elif any("Bi-horário" in o for o in opcoes_validas): default_opcao_idx = [i for i,o in enumerate(opcoes_validas) if "Bi-horário" in o][0]
    opcao_horaria = st.selectbox("Opção Horária e Ciclo", opcoes_validas, key="sel_opcao_horaria", help="Simples, Bi-horário ou Tri-horário")

with col3:
    mes_atual_idx = datetime.datetime.now().month - 1
    mes = st.selectbox("Mês", ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"], key="sel_mes", help="Se o mês escolhido já tiver terminado, o valor do OMIE é final, se ainda estiver em curso será com Futuros, que pode consultar no site www.tiagofelicia.pt")

# --- Datas e dias ---
dias_mes = {"Janeiro": 31, "Fevereiro": 28, "Março": 31, "Abril": 30, "Maio": 31, "Junho": 30, "Julho": 31, "Agosto": 31, "Setembro": 30, "Outubro": 31, "Novembro": 30, "Dezembro": 31}
ano_atual = datetime.datetime.now().year

if mes == "Fevereiro" and ((ano_atual % 4 == 0 and ano_atual % 100 != 0) or (ano_atual % 400 == 0)):
    dias_mes["Fevereiro"] = 29
mes_num = list(dias_mes.keys()).index(mes) + 1


# --- LÓGICA DE GESTÃO DE ESTADO PARA DATAS ---

# 1. INICIALIZAÇÃO NA PRIMEIRA EXECUÇÃO DA SESSÃO
if 'session_initialized_dates' not in st.session_state:
    hoje = datetime.date.today()
    data_inicial_default_calc = hoje + datetime.timedelta(days=1)
    
    ano_final_calc = data_inicial_default_calc.year
    mes_final_calc = data_inicial_default_calc.month + 1
    if mes_final_calc > 12:
        mes_final_calc = 1
        ano_final_calc += 1
    
    dias_no_mes_final = monthrange(ano_final_calc, mes_final_calc)[1]
    dia_final_calc = min(data_inicial_default_calc.day, dias_no_mes_final)
    data_final_bruta = datetime.date(ano_final_calc, mes_final_calc, dia_final_calc)
    data_final_default_calc = data_final_bruta - datetime.timedelta(days=1)
    
    st.session_state.data_inicio_val = data_inicial_default_calc
    st.session_state.data_fim_val = data_final_default_calc
    st.session_state.previous_mes_for_dates = mes
    st.session_state.session_initialized_dates = True


# 2. VERIFICAR SE O UTILIZADOR MUDOU O MÊS
mes_changed = st.session_state.get('previous_mes_for_dates') != mes
if mes_changed:
    st.session_state.previous_mes_for_dates = mes
    
    primeiro_dia_mes_selecionado = datetime.date(ano_atual, mes_num, 1)
    ultimo_dia_mes_selecionado = datetime.date(ano_atual, mes_num, dias_mes[mes])
    
    st.session_state.data_inicio_val = primeiro_dia_mes_selecionado
    st.session_state.data_fim_val = ultimo_dia_mes_selecionado
    
    if 'dias_manual_val' in st.session_state:
        del st.session_state['dias_manual_val']

# 3. CRIAR OS WIDGETS DE DATA E DIAS
col4, col5, col6 = st.columns(3)

# Guardar os valores antes de os widgets poderem ser alterados, para detetar mudanças manuais
data_inicio_anterior = st.session_state.get('data_inicio_val')
data_fim_anterior = st.session_state.get('data_fim_val')

with col4:
    data_inicio = st.date_input("Data Inicial", 
                                value=st.session_state.data_inicio_val, 
                                format="DD/MM/YYYY",
                                key="data_inicio_key_input", 
                                help="A partir de 01/01/2025. Se não modificar as datas ou o mês, será calculado a partir do dia seguinte ao atual.")

with col5:
    data_fim = st.date_input("Data Final", 
                             value=st.session_state.data_fim_val, 
                             format="DD/MM/YYYY",
                             key="data_fim_key_input", 
                             help="De Data Inicial a 31/12/2025. Se não modificar as datas ou o mês, será calculado até um mês após a data inicial.")

# Lógica de reset para os dias manuais
data_inicio_mudou_manualmente = data_inicio_anterior != data_inicio
data_fim_mudou_manualmente = data_fim_anterior != data_fim

if mes_changed or data_inicio_mudou_manualmente or data_fim_mudou_manualmente:
    if 'dias_manual_val' in st.session_state:
        del st.session_state['dias_manual_val']

# Calcular dias_default com base nas datas ATUAIS
dias_default_calculado = (data_fim - data_inicio).days + 1 if data_fim >= data_inicio else 0

with col6:
    dias_manual_input_val = st.number_input("Número de Dias (manual)", min_value=0,
                                        value=st.session_state.get('dias_manual_val', dias_default_calculado),
                                        step=1, key="dias_manual_input_key", 
                                        help="Pode alterar os dias de forma manual, mas dê preferência às datas ou mês, para ter dados mais fidedignos nos tarifários indexados")
    st.session_state['dias_manual_val'] = dias_manual_input_val

if pd.isna(dias_manual_input_val) or dias_manual_input_val <= 0:
    dias = dias_default_calculado
else:
    dias = int(dias_manual_input_val)

st.write(f"Dias considerados: **{dias} dias**")


# Ler e Processar a Constante Data_Valores_OMIE
data_valores_omie_dt = None
nota_omie = " (Info OMIE Indisp.)" # Default se a data não puder ser processada

constante_row_data_omie = CONSTANTES[CONSTANTES['constante'] == 'Data_Valores_OMIE']
if not constante_row_data_omie.empty:
    valor_raw = constante_row_data_omie['valor_unitário'].iloc[0]
    if pd.notna(valor_raw):
        try:
            # Pandas geralmente lê datas como datetime64[ns] que se tornam Timestamp
            if isinstance(valor_raw, (datetime.datetime, pd.Timestamp)):
                data_valores_omie_dt = valor_raw.date() # Converter para objeto date
            else:
                # Tentar converter para pd.Timestamp primeiro, depois para date
                timestamp_convertido = pd.to_datetime(valor_raw, errors='coerce')
                if pd.notna(timestamp_convertido):
                    data_valores_omie_dt = timestamp_convertido.date()
            
            if pd.isna(data_valores_omie_dt): # Se a conversão resultou em NaT (Not a Time)
                data_valores_omie_dt = None
                st.warning(f"Não foi possível converter 'Data_Valores_OMIE' para uma data válida: {valor_raw}")
        except Exception as e:
            st.warning(f"Erro ao processar 'Data_Valores_OMIE' ('{valor_raw}'): {e}")
            data_valores_omie_dt = None # Garantir que fica None em caso de erro
    else:
        st.warning("Valor para 'Data_Valores_OMIE' está vazio na folha Constantes.")
else:
    st.warning("Constante 'Data_Valores_OMIE' não encontrada na folha Constantes.")

# Determinar a nota para os inputs OMIE
# 'data_fim' já deve ser um objeto datetime.date
if data_valores_omie_dt and isinstance(data_fim, datetime.date):
    if data_fim <= data_valores_omie_dt:
        nota_omie = " (Média Final)"
    else:
        nota_omie = " (Média com Futuros)"
# Se data_valores_omie_dt for None, a nota_omie permanecerá "(Info OMIE Indisp.)"
# FIM - Ler e Processar a Constante Data_Valores_OMIE

# --- LÓGICA DE RESET DOS INPUTS OMIE ---
# Gerar uma chave única para os parâmetros que afetam os defaults OMIE
current_omie_dependency_key = f"{data_inicio}-{data_fim}-{opcao_horaria}"

if st.session_state.get('last_omie_dependency_key_for_inputs') != current_omie_dependency_key:
    st.session_state.last_omie_dependency_key_for_inputs = current_omie_dependency_key
    omie_input_keys_to_reset = ['omie_s_input_field', 'omie_v_input_field', 'omie_f_input_field', 'omie_c_input_field', 'omie_p_input_field']
    for key_in_state in omie_input_keys_to_reset:
        if key_in_state in st.session_state:
            del st.session_state[key_in_state]
    st.session_state.omie_foi_editado_manualmente = {}
    # st.write("Debug: OMIE inputs e flags de edição resetados devido à mudança de parâmetros.")

# --- Calcular valores OMIE médios POR PERÍODO (V, F, C, P) e Global (S) ---
# Estes são os valores CALCULADOS da tabela, antes de qualquer input manual
df_omie_no_periodo_selecionado = pd.DataFrame()
if 'DataHora' in OMIE_PERDAS_CICLOS.columns:
    df_omie_no_periodo_selecionado = OMIE_PERDAS_CICLOS[
        (OMIE_PERDAS_CICLOS['DataHora'] >= pd.to_datetime(data_inicio)) & # Usar data_inicio
        (OMIE_PERDAS_CICLOS['DataHora'] <= pd.to_datetime(data_fim) + pd.Timedelta(hours=23, minutes=59, seconds=59)) # Usar data_fim
    ].copy()
else:
    st.warning("Coluna 'DataHora' não encontrada nos dados OMIE. Não é possível calcular médias OMIE.")

omie_medios_calculados = {'S': 0.0, 'V': 0.0, 'F': 0.0, 'C': 0.0, 'P': 0.0}

if not df_omie_no_periodo_selecionado.empty:
    omie_medios_calculados['S'] = df_omie_no_periodo_selecionado['OMIE'].mean()

    if pd.isna(omie_medios_calculados['S']): omie_medios_calculados['S'] = 0.0
    ciclo_bi_col = 'BD' if "Diário" in opcao_horaria else 'BS'
    ciclo_tri_col = 'TD' if "Diário" in opcao_horaria else 'TS'
    if ciclo_bi_col in df_omie_no_periodo_selecionado.columns:
        omie_bi_calculado = df_omie_no_periodo_selecionado.groupby(ciclo_bi_col)['OMIE'].mean()
        omie_medios_calculados['V'] = omie_bi_calculado.get('V', omie_medios_calculados.get('V', 0.0))
        omie_medios_calculados['F'] = omie_bi_calculado.get('F', omie_medios_calculados.get('F', 0.0))
    if ciclo_tri_col in df_omie_no_periodo_selecionado.columns:
        omie_tri_calculado = df_omie_no_periodo_selecionado.groupby(ciclo_tri_col)['OMIE'].mean()
        omie_medios_calculados['V'] = omie_tri_calculado.get('V', omie_medios_calculados.get('V',0.0))
        omie_medios_calculados['C'] = omie_tri_calculado.get('C', omie_medios_calculados.get('C',0.0))
        omie_medios_calculados['P'] = omie_tri_calculado.get('P', omie_medios_calculados.get('P',0.0))

else:
    st.warning("Não existem dados OMIE para o período selecionado. As médias OMIE serão zero.")


# --- Calcular OMIEs médios para TODOS OS CICLOS POSSÍVEIS ---
omie_medios_calculados_para_todos_ciclos = {'S': 0.0} # Inicializar com Simples

if not df_omie_no_periodo_selecionado.empty and 'OMIE' in df_omie_no_periodo_selecionado.columns:
    omie_medios_calculados_para_todos_ciclos['S'] = df_omie_no_periodo_selecionado['OMIE'].mean()
    if pd.isna(omie_medios_calculados_para_todos_ciclos['S']):
        omie_medios_calculados_para_todos_ciclos['S'] = 0.0

    ciclos_a_processar = {
        'BD': ['V', 'F'], 'BS': ['V', 'F'],
        'TD': ['V', 'C', 'P'], 'TS': ['V', 'C', 'P']
    }
    for ciclo_curto, periodos_ciclo in ciclos_a_processar.items():
        if ciclo_curto in df_omie_no_periodo_selecionado.columns:
            omie_ciclo_calculado = df_omie_no_periodo_selecionado.groupby(ciclo_curto)['OMIE'].mean()
            for p_ciclo in periodos_ciclo:
                chave_completa = f"{ciclo_curto}_{p_ciclo}"
                omie_medios_calculados_para_todos_ciclos[chave_completa] = omie_ciclo_calculado.get(p_ciclo, 0.0)
                if pd.isna(omie_medios_calculados_para_todos_ciclos[chave_completa]):
                     omie_medios_calculados_para_todos_ciclos[chave_completa] = 0.0
        else: # Fallback se a coluna do ciclo não existir
            for p_ciclo in periodos_ciclo:
                chave_completa = f"{ciclo_curto}_{p_ciclo}"
                omie_medios_calculados_para_todos_ciclos[chave_completa] = omie_medios_calculados_para_todos_ciclos['S'] # Usa OMIE Simples como fallback


# Garantir que todas as chaves esperadas existem, mesmo que com valor de OMIE Simples como fallback
chaves_omie_esperadas = ['S']
for ciclo_key_esperado in ['BD', 'BS']:
    for periodo_key_esperado in ['V', 'F']:
        chaves_omie_esperadas.append(f"{ciclo_key_esperado}_{periodo_key_esperado}")
for ciclo_key_esperado in ['TD', 'TS']:
    for periodo_key_esperado in ['V', 'C', 'P']:
        chaves_omie_esperadas.append(f"{ciclo_key_esperado}_{periodo_key_esperado}")

for k_omie_calc in chaves_omie_esperadas:
    if k_omie_calc not in omie_medios_calculados_para_todos_ciclos:
        omie_medios_calculados_para_todos_ciclos[k_omie_calc] = omie_medios_calculados_para_todos_ciclos.get('S', 0.0) # Default para OMIE Simples

# --- Inputs Manuais OMIE pelo utilizador e Deteção de Edição ---
st.session_state.omie_foi_editado_manualmente = st.session_state.get('omie_foi_editado_manualmente', {})
omie_medios_calculados = {'S': 0.0, 'V': 0.0, 'F': 0.0, 'C': 0.0, 'P': 0.0} # Recalcular aqui com base em df_omie_no_periodo_selecionado

if not df_omie_no_periodo_selecionado.empty:
    omie_medios_calculados['S'] = df_omie_no_periodo_selecionado['OMIE'].mean()
    if pd.isna(omie_medios_calculados['S']): omie_medios_calculados['S'] = 0.0
    ciclo_bi_col = 'BD' if "Diário" in opcao_horaria else 'BS'
    ciclo_tri_col = 'TD' if "Diário" in opcao_horaria else 'TS'
    if ciclo_bi_col in df_omie_no_periodo_selecionado.columns:
        omie_bi_calculado = df_omie_no_periodo_selecionado.groupby(ciclo_bi_col)['OMIE'].mean()
        omie_medios_calculados['V'] = omie_bi_calculado.get('V', omie_medios_calculados.get('V', 0.0))
        omie_medios_calculados['F'] = omie_bi_calculado.get('F', omie_medios_calculados.get('F', 0.0))
    if ciclo_tri_col in df_omie_no_periodo_selecionado.columns:
        omie_tri_calculado = df_omie_no_periodo_selecionado.groupby(ciclo_tri_col)['OMIE'].mean()
        omie_medios_calculados['V'] = omie_tri_calculado.get('V', omie_medios_calculados.get('V',0.0))
        omie_medios_calculados['C'] = omie_tri_calculado.get('C', omie_medios_calculados.get('C',0.0))
        omie_medios_calculados['P'] = omie_tri_calculado.get('P', omie_medios_calculados.get('P',0.0))
else:
    st.warning("Não existem dados OMIE para o período selecionado. As médias OMIE serão zero.")

# ... Lógica para Simples para inputs OMIE e flags de edição ...
if opcao_horaria.lower() == "simples":
    default_s = st.session_state.get('omie_s_input_field', round(omie_medios_calculados['S'], 2))
    label_s_completo = f"Valor OMIE (€/MWh) - Simples{nota_omie}"
    
    omie_s_manual = st.number_input(
        label_s_completo, 
        value=default_s, 
        step=1.0, 
        format="%.2f", # <-- Mostra sempre 2 casas decimais
        key="omie_s_input", 
        help="A primeira alteração arredonda para o inteiro mais próximo. As seguintes incrementam de 1 em 1."
    )
    
    # --- LÓGICA PERSONALIZADA DE ARREDONDAMENTO ---
    valor_anterior_s = st.session_state.get('omie_s_input_field', round(omie_medios_calculados['S'], 2))
    
    if omie_s_manual != valor_anterior_s: # Se o utilizador interagiu com o widget
        # E se for a PRIMEIRA vez que é editado
        if not st.session_state.omie_foi_editado_manualmente.get('S', False):
            valor_arredondado = float(round(omie_s_manual))
            st.session_state.omie_s_input_field = valor_arredondado
            st.session_state.omie_foi_editado_manualmente['S'] = True
            st.rerun() # Força a atualização do widget para o valor arredondado
        else: # Se já foi editado antes, apenas atualiza o valor
            st.session_state.omie_s_input_field = omie_s_manual
    else:
        # Se não houve interação, apenas garante que o valor está guardado
        st.session_state.omie_s_input_field = omie_s_manual


elif opcao_horaria.lower().startswith("bi"):
    col_omie1, col_omie2 = st.columns(2)
    with col_omie1:
        default_v = st.session_state.get('omie_v_input_field', round(omie_medios_calculados['V'], 2))
        label_v_completo = f"Valor OMIE (€/MWh) - Vazio{nota_omie}"
        omie_v_manual = st.number_input(label_v_completo, value=default_v, step=1.0, format="%.2f", key="omie_v_input", help="A primeira alteração arredonda para o inteiro mais próximo.")
        
        valor_anterior_v = st.session_state.get('omie_v_input_field', round(omie_medios_calculados['V'], 2))
        if omie_v_manual != valor_anterior_v:
            if not st.session_state.omie_foi_editado_manualmente.get('V', False):
                st.session_state.omie_v_input_field = float(round(omie_v_manual))
                st.session_state.omie_foi_editado_manualmente['V'] = True
                st.rerun()
            else:
                st.session_state.omie_v_input_field = omie_v_manual
        else:
            st.session_state.omie_v_input_field = omie_v_manual

    with col_omie2:
        default_f = st.session_state.get('omie_f_input_field', round(omie_medios_calculados['F'], 2))
        label_f_completo = f"Valor OMIE (€/MWh) - Fora Vazio{nota_omie}"
        omie_f_manual = st.number_input(label_f_completo, value=default_f, step=1.0, format="%.2f", key="omie_f_input", help="A primeira alteração arredonda para o inteiro mais próximo.")

        valor_anterior_f = st.session_state.get('omie_f_input_field', round(omie_medios_calculados['F'], 2))
        if omie_f_manual != valor_anterior_f:
            if not st.session_state.omie_foi_editado_manualmente.get('F', False):
                st.session_state.omie_f_input_field = float(round(omie_f_manual))
                st.session_state.omie_foi_editado_manualmente['F'] = True
                st.rerun()
            else:
                st.session_state.omie_f_input_field = omie_f_manual
        else:
            st.session_state.omie_f_input_field = omie_f_manual


elif opcao_horaria.lower().startswith("tri"):
    col_omie1, col_omie2, col_omie3 = st.columns(3)
    with col_omie1:
        default_v = st.session_state.get('omie_v_input_field', round(omie_medios_calculados['V'], 2))
        label_v_completo = f"Valor OMIE (€/MWh) - Vazio{nota_omie}"
        omie_v_manual = st.number_input(label_v_completo, value=default_v, step=1.0, format="%.2f", key="omie_v_input", help="A primeira alteração arredonda para o inteiro mais próximo.")
        
        valor_anterior_v = st.session_state.get('omie_v_input_field', round(omie_medios_calculados['V'], 2))
        if omie_v_manual != valor_anterior_v:
            if not st.session_state.omie_foi_editado_manualmente.get('V', False):
                st.session_state.omie_v_input_field = float(round(omie_v_manual))
                st.session_state.omie_foi_editado_manualmente['V'] = True
                st.rerun()
            else:
                st.session_state.omie_v_input_field = omie_v_manual
        else:
            st.session_state.omie_v_input_field = omie_v_manual
            
    with col_omie2:
        default_c = st.session_state.get('omie_c_input_field', round(omie_medios_calculados['C'], 2))
        label_c_completo = f"Valor OMIE (€/MWh) - Cheias{nota_omie}"
        omie_c_manual = st.number_input(label_c_completo, value=default_c, step=1.0, format="%.2f", key="omie_c_input", help="A primeira alteração arredonda para o inteiro mais próximo.")
        
        valor_anterior_c = st.session_state.get('omie_c_input_field', round(omie_medios_calculados['C'], 2))
        if omie_c_manual != valor_anterior_c:
            if not st.session_state.omie_foi_editado_manualmente.get('C', False):
                st.session_state.omie_c_input_field = float(round(omie_c_manual))
                st.session_state.omie_foi_editado_manualmente['C'] = True
                st.rerun()
            else:
                st.session_state.omie_c_input_field = omie_c_manual
        else:
            st.session_state.omie_c_input_field = omie_c_manual

    with col_omie3:
        default_p = st.session_state.get('omie_p_input_field', round(omie_medios_calculados['P'], 2))
        label_p_completo = f"Valor OMIE (€/MWh) - Ponta{nota_omie}"
        omie_p_manual = st.number_input(label_p_completo, value=default_p, step=1.0, format="%.2f", key="omie_p_input", help="A primeira alteração arredonda para o inteiro mais próximo.")
        
        valor_anterior_p = st.session_state.get('omie_p_input_field', round(omie_medios_calculados['P'], 2))
        if omie_p_manual != valor_anterior_p:
            if not st.session_state.omie_foi_editado_manualmente.get('P', False):
                st.session_state.omie_p_input_field = float(round(omie_p_manual))
                st.session_state.omie_foi_editado_manualmente['P'] = True
                st.rerun()
            else:
                st.session_state.omie_p_input_field = omie_p_manual
        else:
            st.session_state.omie_p_input_field = omie_p_manual

# --- Alerta para uso de OMIE Manual ---
alertas_omie_manual = []
if st.session_state.omie_foi_editado_manualmente.get('S'): alertas_omie_manual.append("Simples")
if st.session_state.omie_foi_editado_manualmente.get('V'): alertas_omie_manual.append("Vazio")
if st.session_state.omie_foi_editado_manualmente.get('F'): alertas_omie_manual.append("Fora Vazio")
if st.session_state.omie_foi_editado_manualmente.get('C'): alertas_omie_manual.append("Cheias")
if st.session_state.omie_foi_editado_manualmente.get('P'): alertas_omie_manual.append("Ponta")

if alertas_omie_manual:
    st.info(f"ℹ️ Atenção: Os cálculos estão a utilizar valores OMIE manuais (editados) para o(s) período(s): {', '.join(alertas_omie_manual)}. "
            "Para os tarifários quarto-horários, isto significa que este valor OMIE manual (e não os OMIEs horários) será aplicado a todas as horas desse(s) período(s). "
            "Outros períodos não editados usarão os OMIEs horários (para quarto-horários) ou as médias calculadas (para tarifários de média).")

# --- Preparar df_omie_ajustado ---
# Começa com os OMIEs HORÁRIOS ORIGINAIS da tabela para o período selecionado
df_omie_ajustado = df_omie_no_periodo_selecionado.copy()
if not df_omie_ajustado.empty:
    if opcao_horaria.lower() == "simples":
        if st.session_state.omie_foi_editado_manualmente.get('S'):
            df_omie_ajustado['OMIE'] = st.session_state.omie_s_input_field

    elif opcao_horaria.lower().startswith("bi"):
        ciclo_col = 'BD' if "Diário" in opcao_horaria else 'BS'
        if ciclo_col in df_omie_ajustado.columns:
            def substituir_bi_se_manual(row):
                omie_original_hora = row['OMIE']
                if row[ciclo_col] == 'V' and st.session_state.omie_foi_editado_manualmente.get('V'):
                    return st.session_state.omie_v_input_field
                elif row[ciclo_col] == 'F' and st.session_state.omie_foi_editado_manualmente.get('F'):
                    return st.session_state.omie_f_input_field
                return omie_original_hora
            if st.session_state.omie_foi_editado_manualmente.get('V') or st.session_state.omie_foi_editado_manualmente.get('F'):
                 df_omie_ajustado['OMIE'] = df_omie_ajustado.apply(substituir_bi_se_manual, axis=1)

    elif opcao_horaria.lower().startswith("tri"):
        ciclo_col = 'TD' if "Diário" in opcao_horaria else 'TS'
        if ciclo_col in df_omie_ajustado.columns:
            def substituir_tri_se_manual(row):
                omie_original_hora = row['OMIE']
                if row[ciclo_col] == 'V' and st.session_state.omie_foi_editado_manualmente.get('V'):
                    return st.session_state.omie_v_input_field
                elif row[ciclo_col] == 'C' and st.session_state.omie_foi_editado_manualmente.get('C'):
                    return st.session_state.omie_c_input_field
                elif row[ciclo_col] == 'P' and st.session_state.omie_foi_editado_manualmente.get('P'):
                    return st.session_state.omie_p_input_field
                return omie_original_hora
            if st.session_state.omie_foi_editado_manualmente.get('V') or \
               st.session_state.omie_foi_editado_manualmente.get('C') or \
               st.session_state.omie_foi_editado_manualmente.get('P'):
                df_omie_ajustado['OMIE'] = df_omie_ajustado.apply(substituir_tri_se_manual, axis=1)
        else:
            st.warning(f"Coluna de ciclo '{ciclo_col}' não encontrada nos dados OMIE. Não é possível aplicar OMIE manual por período horário.")
else: # df_omie_ajustado está vazio porque df_omie_no_periodo_selecionado estava vazio
    st.warning("DataFrame OMIE ajustado está vazio pois não há dados OMIE para o período.")

# --- Recalcular omie_medio_simples_real_kwh com base nos OMIE ajustados ---
# Este valor é usado por alguns tarifários de MÉDIA (ex: LuziGás)
if not df_omie_ajustado.empty and 'OMIE' in df_omie_ajustado.columns:
    omie_medio_simples_real_kwh = df_omie_ajustado['OMIE'].mean() / 1000.0
    if pd.isna(omie_medio_simples_real_kwh): omie_medio_simples_real_kwh = 0.0
else:
    omie_medio_simples_real_kwh = 0.0

# --- Guardar os valores OMIE médios (para tarifários de MÉDIA) ---
# Estes virão dos inputs, que por sua vez são inicializados com as médias calculadas ou com os valores manuais do utilizador.
omie_para_tarifarios_media = {}
if opcao_horaria.lower() == "simples":
    omie_para_tarifarios_media['S'] = st.session_state.get('omie_s_input_field', omie_medios_calculados['S'])
elif opcao_horaria.lower().startswith("bi"):
    omie_para_tarifarios_media['V'] = st.session_state.get('omie_v_input_field', omie_medios_calculados['V'])
    omie_para_tarifarios_media['F'] = st.session_state.get('omie_f_input_field', omie_medios_calculados['F'])
elif opcao_horaria.lower().startswith("tri"):
    omie_para_tarifarios_media['V'] = st.session_state.get('omie_v_input_field', omie_medios_calculados['V'])
    omie_para_tarifarios_media['C'] = st.session_state.get('omie_c_input_field', omie_medios_calculados['C'])
    omie_para_tarifarios_media['P'] = st.session_state.get('omie_p_input_field', omie_medios_calculados['P'])

# --- Calcular Perdas médias para TODOS os ciclos e períodos (GLOBALMENTE) ---
perdas_medias = {}
if not df_omie_no_periodo_selecionado.empty and 'Perdas' in df_omie_no_periodo_selecionado.columns:
    # Médias para o período selecionado
    perdas_medias['Perdas_M_S'] = df_omie_no_periodo_selecionado['Perdas'].mean()
    
    for ciclo_base_curto in ['BD', 'BS']: # Bi-Diário, Bi-Semanal
        if ciclo_base_curto in df_omie_no_periodo_selecionado.columns:
            perdas_ciclo_periodo = df_omie_no_periodo_selecionado.groupby(ciclo_base_curto)['Perdas'].mean()
            perdas_medias[f'Perdas_M_{ciclo_base_curto}_V'] = perdas_ciclo_periodo.get('V', 1.0)
            perdas_medias[f'Perdas_M_{ciclo_base_curto}_F'] = perdas_ciclo_periodo.get('F', 1.0)
        else: # Fallback se coluna de ciclo não existir para o período selecionado
            perdas_medias[f'Perdas_M_{ciclo_base_curto}_V'] = perdas_medias['Perdas_M_S'] # Usa média simples como fallback
            perdas_medias[f'Perdas_M_{ciclo_base_curto}_F'] = perdas_medias['Perdas_M_S']

    for ciclo_base_curto in ['TD', 'TS']: # Tri-Diário, Tri-Semanal
        if ciclo_base_curto in df_omie_no_periodo_selecionado.columns:
            perdas_ciclo_periodo = df_omie_no_periodo_selecionado.groupby(ciclo_base_curto)['Perdas'].mean()
            perdas_medias[f'Perdas_M_{ciclo_base_curto}_V'] = perdas_ciclo_periodo.get('V', 1.0)
            perdas_medias[f'Perdas_M_{ciclo_base_curto}_C'] = perdas_ciclo_periodo.get('C', 1.0)
            perdas_medias[f'Perdas_M_{ciclo_base_curto}_P'] = perdas_ciclo_periodo.get('P', 1.0)
        else: # Fallback
            perdas_medias[f'Perdas_M_{ciclo_base_curto}_V'] = perdas_medias['Perdas_M_S']
            perdas_medias[f'Perdas_M_{ciclo_base_curto}_C'] = perdas_medias['Perdas_M_S']
            perdas_medias[f'Perdas_M_{ciclo_base_curto}_P'] = perdas_medias['Perdas_M_S']

    # Médias para o ano completo
    df_omie_ano_completo_pm = OMIE_PERDAS_CICLOS[OMIE_PERDAS_CICLOS['DataHora'].dt.year == ano_atual].copy() # Renomeado df_omie_ano_completo
    if not df_omie_ano_completo_pm.empty and 'Perdas' in df_omie_ano_completo_pm.columns:
        perdas_medias['Perdas_Anual_S'] = df_omie_ano_completo_pm['Perdas'].mean()

        for ciclo_base_curto_anual in ['BD', 'BS']:
            if ciclo_base_curto_anual in df_omie_ano_completo_pm.columns:
                perdas_ciclo_anual = df_omie_ano_completo_pm.groupby(ciclo_base_curto_anual)['Perdas'].mean()
                perdas_medias[f'Perdas_Anual_{ciclo_base_curto_anual}_V'] = perdas_ciclo_anual.get('V', 1.0)
                perdas_medias[f'Perdas_Anual_{ciclo_base_curto_anual}_F'] = perdas_ciclo_anual.get('F', 1.0)
            else: # Fallback
                perdas_medias[f'Perdas_Anual_{ciclo_base_curto_anual}_V'] = perdas_medias.get('Perdas_Anual_S', 1.0)
                perdas_medias[f'Perdas_Anual_{ciclo_base_curto_anual}_F'] = perdas_medias.get('Perdas_Anual_S', 1.0)

        for ciclo_base_curto_anual in ['TD', 'TS']:
            if ciclo_base_curto_anual in df_omie_ano_completo_pm.columns:
                perdas_ciclo_anual = df_omie_ano_completo_pm.groupby(ciclo_base_curto_anual)['Perdas'].mean()
                perdas_medias[f'Perdas_Anual_{ciclo_base_curto_anual}_V'] = perdas_ciclo_anual.get('V', 1.0)
                perdas_medias[f'Perdas_Anual_{ciclo_base_curto_anual}_C'] = perdas_ciclo_anual.get('C', 1.0)
                perdas_medias[f'Perdas_Anual_{ciclo_base_curto_anual}_P'] = perdas_ciclo_anual.get('P', 1.0)
            else: # Fallback
                perdas_medias[f'Perdas_Anual_{ciclo_base_curto_anual}_V'] = perdas_medias.get('Perdas_Anual_S', 1.0)
                perdas_medias[f'Perdas_Anual_{ciclo_base_curto_anual}_C'] = perdas_medias.get('Perdas_Anual_S', 1.0)
                perdas_medias[f'Perdas_Anual_{ciclo_base_curto_anual}_P'] = perdas_medias.get('Perdas_Anual_S', 1.0)
    else:
        st.warning("Não existem dados OMIE para o ano completo. Algumas médias de perdas anuais podem não ser calculadas.")
else:
    st.warning("Não existem dados OMIE ou coluna 'Perdas' para o período selecionado. As médias de perdas podem não ser calculadas corretamente.")
# Garantir que todas as chaves esperadas existem em perdas_medias, mesmo que com default 1.0
chaves_perdas_esperadas = [
    'Perdas_M_S', 'Perdas_Anual_S',
    'Perdas_M_BD_V', 'Perdas_Anual_BD_V', 'Perdas_M_BD_F', 'Perdas_Anual_BD_F',
    'Perdas_M_BS_V', 'Perdas_Anual_BS_V', 'Perdas_M_BS_F', 'Perdas_Anual_BS_F',
    'Perdas_M_TD_V', 'Perdas_Anual_TD_V', 'Perdas_M_TD_C', 'Perdas_Anual_TD_C', 'Perdas_M_TD_P', 'Perdas_Anual_TD_P',
    'Perdas_M_TS_V', 'Perdas_Anual_TS_V', 'Perdas_M_TS_C', 'Perdas_Anual_TS_C', 'Perdas_M_TS_P', 'Perdas_Anual_TS_P',
]
for k_perda in chaves_perdas_esperadas:
    if k_perda not in perdas_medias:
        perdas_medias[k_perda] = 1.0 # Default para perdas

# --- FUNÇÃO AUXILIAR PARA OPÇÕES DE FILTRO ---
def get_filter_options_for_multiselect(df_fixos, df_indexados, column_name):
    options = []
    if df_fixos is not None and column_name in df_fixos.columns:
        options.extend(df_fixos[column_name].astype(str).str.strip().dropna().unique())
    if df_indexados is not None and column_name in df_indexados.columns:
        options.extend(df_indexados[column_name].astype(str).str.strip().dropna().unique())
    
    if column_name == 'Tipo':
        options.append("Pessoal")
    if column_name == 'Segmento': # Se o seu "Meu Tarifário" tiver um segmento específico
        options.append("Pessoal") # Adapte conforme necessário ou remova se não aplicável
        
    unique_options_intermediate = set(opt for opt in options if opt and opt.lower() != 'nan')
    
    return sorted(list(unique_options_intermediate))
# --- FIM FUNÇÃO AUXILIAR PARA OPÇÕES DE FILTRO ---

# --- Consumos ---
st.subheader("⚡ Consumos")
st.markdown(
    "Introduza o consumo para cada período (kWh). Pode usar somas ou subtrações simples (ex: `100+50+30` ou `100+50-30`). Os valores serão arredondados para o inteiro mais próximo."
)

# Função para calcular a expressão de consumo (apenas para somas, resultado inteiro)
def calcular_expressao_matematica_simples(expressao_str, periodo_label=""):
    """
    Calcula uma expressão matemática simples de adição e subtração, 
    arredondando o resultado para o inteiro mais próximo.
    Ex: '10+20-5', '10.5 - 2.5 + 0.5'
    """
    if not expressao_str or not isinstance(expressao_str, str):
        return 0, f"Nenhum valor introduzido para {periodo_label}." if periodo_label else "Nenhum valor introduzido."

    # 1. Validação de caracteres permitidos
    # Permite dígitos, ponto decimal, operadores + e -, e espaços.
    valid_chars = set('0123456789.+- ')
    if not all(char in valid_chars for char in expressao_str):
        return 0, f"Expressão inválida para {periodo_label}: '{expressao_str}'. Use apenas números, '.', '+', '-'. O resultado será arredondado."

    expressao_limpa = expressao_str.replace(" ", "") # Remove todos os espaços
    if not expressao_limpa: # Se após remover espaços a string estiver vazia
        return 0, f"Expressão vazia para {periodo_label}."

    # 2. Normalizar operadores duplos (ex: -- para +, +- para -)
    # Este loop garante que sequências como "---" ou "-+-" são corretamente simplificadas.
    temp_expr = expressao_limpa
    while True:
        prev_expr = temp_expr
        temp_expr = temp_expr.replace("--", "+")
        temp_expr = temp_expr.replace("+-", "-")
        temp_expr = temp_expr.replace("-+", "-")
        temp_expr = temp_expr.replace("++", "+")
        if temp_expr == prev_expr: # Termina quando não há mais alterações
            break
    expressao_limpa = temp_expr

    # 3. Verificar se a expressão é apenas um operador ou termina/começa invalidamente com um
    if expressao_limpa in ["+", "-"] or \
       expressao_limpa.endswith(("+", "-")) or \
       (expressao_limpa.startswith(("+", "-")) and len(expressao_limpa) > 1 and expressao_limpa[1] in "+-"): # Ex: "++5", "-+5" já normalizado, mas evita "+5", "-5" aqui
        if not ( (expressao_limpa.startswith(("+", "-")) and len(expressao_limpa) > 1 and expressao_limpa[1].isdigit()) or \
                 (expressao_limpa.startswith(("+", "-")) and len(expressao_limpa) > 2 and expressao_limpa[1] == '.' and expressao_limpa[2].isdigit() ) ): # Permite "+5", "-5", "+.5", "-.5"
            return 0, f"Expressão inválida para {periodo_label}: '{expressao_str}'. Formato de operador inválido."


    # 4. Adicionar um '+' no início se a expressão começar com um número ou ponto decimal, para facilitar a divisão.
    #    Ex: "10-5" -> "+10-5"; ".5+2" -> "+.5+2"
    if expressao_limpa and (expressao_limpa[0].isdigit() or \
        (expressao_limpa.startswith('.') and len(expressao_limpa) > 1 and expressao_limpa[1].isdigit())):
        expressao_limpa = "+" + expressao_limpa
    elif expressao_limpa.startswith('.') and not (len(expressao_limpa) > 1 and (expressao_limpa[1].isdigit() or expressao_limpa[1] in "+-")): # Casos como "." ou ".+"
         return 0, f"Expressão inválida para {periodo_label}: '{expressao_str}'. Ponto decimal mal formatado."

    # 5. Dividir a expressão em operadores ([+\-]) e os operandos que se seguem.
    #    Ex: "+10.5-5" -> ['', '+', '10.5', '-', '5'] (o primeiro '' é por causa do split no início)
    #    Ex: "-5+3" -> ['', '-', '5', '+', '3']
    partes = re.split(r'([+\-])', expressao_limpa)
    
    # Filtrar strings vazias resultantes do split (principalmente a primeira se existir)
    partes_filtradas = [p for p in partes if p]

    if not partes_filtradas:
        return 0, f"Expressão inválida para {periodo_label}: '{expressao_str}'. Não resultou em operandos válidos."

    # A estrutura deve ser [operador, operando, operador, operando, ...]
    # Portanto, o comprimento da lista filtrada deve ser par e pelo menos 2 (ex: ['+', '10'])
    if len(partes_filtradas) % 2 != 0 or len(partes_filtradas) == 0:
        return 0, f"Expressão mal formada para {periodo_label}: '{expressao_str}'. Estrutura de operadores/operandos inválida."

    total = 0.0
    try:
        for i in range(0, len(partes_filtradas), 2):
            operador = partes_filtradas[i]
            operando_str = partes_filtradas[i+1]

            if not operando_str : # Operando em falta
                return 0, f"Expressão mal formada para {periodo_label}: '{expressao_str}'. Operando em falta após operador '{operador}'."

            # Validação robusta do operando antes de converter para float
            # Deve ser um número, pode conter um ponto decimal. Não pode ser apenas "."
            if operando_str == '.' or not operando_str.replace('.', '', 1).isdigit():
                 return 0, f"Operando inválido '{operando_str}' na expressão para {periodo_label}."
            
            valor_operando = float(operando_str)

            if operador == '+':
                total += valor_operando
            elif operador == '-':
                total -= valor_operando
            else: 
                # Esta condição não deve ser atingida devido ao re.split('([+\-])')
                return 0, f"Operador desconhecido '{operador}' na expressão para {periodo_label}."

    except ValueError: # Erro ao converter operando_str para float
        return 0, f"Expressão inválida para {periodo_label}: '{expressao_str}'. Contém valor não numérico ou mal formatado."
    except IndexError: # Falha ao aceder partes_filtradas[i+1], indica erro de parsing não apanhado antes.
        return 0, f"Expressão mal formada para {periodo_label}: '{expressao_str}'. Estrutura inesperada."
    except Exception as e: # Captura outras exceções inesperadas
        return 0, f"Erro ao calcular expressão para {periodo_label} ('{expressao_str}'): {e}"

    # Arredondar para o inteiro mais próximo
    total_arredondado = int(round(total))

    # Manter a lógica original de não permitir consumo negativo
    if total_arredondado < 0:
        return 0, f"Consumo calculado para {periodo_label} não pode ser negativo ({total_arredondado} kWh, calculado de {total:.4f} kWh)."
    
    return total_arredondado, None # Retorna o valor inteiro arredondado e nenhum erro

# --- Lógica de Reset dos Inputs de Consumo ao Mudar Opção Horária ---
# (Manter a lógica de reset que já tinha para as chaves 'exp_consumo_...')
chave_opcao_horaria_consumo = f"consumo_inputs_para_{opcao_horaria}" # Esta chave pode ser simplificada
if st.session_state.get('last_opcao_horaria_for_consumo_calc') != opcao_horaria:
    st.session_state.last_opcao_horaria_for_consumo_calc = opcao_horaria
    keys_exp_a_resetar = ['exp_consumo_s', 'exp_consumo_v', 'exp_consumo_f', 'exp_consumo_c', 'exp_consumo_p',
                          'exp_consumo_s_anterior_valido', 'exp_consumo_v_anterior_valido', 
                          'exp_consumo_f_anterior_valido', 'exp_consumo_tc_anterior_valido', 
                          'exp_consumo_tp_anterior_valido', 'exp_consumo_tv_anterior_valido'] # Adicionar também as de _anterior_valido
    for k_exp in keys_exp_a_resetar:
        if k_exp in st.session_state:
            del st.session_state[k_exp]
    # Resetar os valores no session_state para os defaults da nova opcao_horaria
    if opcao_horaria.lower() == "simples":
        st.session_state.exp_consumo_s = "158"
    elif opcao_horaria.lower().startswith("bi"):
        st.session_state.exp_consumo_v = "63"
        st.session_state.exp_consumo_f = "95"
    elif opcao_horaria.lower().startswith("tri"):
        st.session_state.exp_consumo_v = "63" # Reutilizando a chave para Vazio de bi
        st.session_state.exp_consumo_c = "68"
        st.session_state.exp_consumo_p = "27"


# Inicializar as variáveis de consumo que serão usadas nos cálculos
consumo_simples = 0
consumo_vazio = 0
consumo_fora_vazio = 0
consumo_cheias = 0
consumo_ponta = 0
consumo = 0 # Total

if opcao_horaria.lower() == "simples":
    st.markdown("###### Consumo Simples (kWh)")
    expressao_s = st.text_input(
        "Introduza o consumo total ou um cálculo (ex: 80+20+58 ou 200+58-100)",
        value=st.session_state.get('exp_consumo_s', "158"), # Default como string inteira
        key="exp_consumo_s_widget"
    )
    st.session_state.exp_consumo_s = expressao_s
    
    resultado_s_calc, erro_s = calcular_expressao_matematica_simples(expressao_s, "Simples")
    
    if erro_s:
        st.error(erro_s)
        consumo_simples = int(st.session_state.get('exp_consumo_s_anterior_valido', 0)) # Usar último válido ou 0
    else:
        consumo_simples = resultado_s_calc
        st.session_state.exp_consumo_s_anterior_valido = consumo_simples # Guardar último valor válido
        st.info(f"Consumo Simples considerado: **{consumo_simples:.0f} kWh**") # Exibe como inteiro
    
    consumo = consumo_simples

elif opcao_horaria.lower().startswith("bi"):
    st.markdown("###### Consumos Bi-Horário (kWh)")
    col_bi1, col_bi2 = st.columns(2)
    with col_bi1:
        expressao_v = st.text_input("Vazio (ex: 60+3 ou 70+3-10)", value=st.session_state.get('exp_consumo_v', "63"), key="exp_consumo_v_widget")
        st.session_state.exp_consumo_v = expressao_v
        consumo_vazio, erro_v = calcular_expressao_matematica_simples(expressao_v, "Vazio")
        if erro_v: 
            st.error(erro_v)
            consumo_vazio = int(st.session_state.get('exp_consumo_v_anterior_valido', 0))
        else:
            st.session_state.exp_consumo_v_anterior_valido = consumo_vazio
            st.info(f"Consumo Vazio considerado: **{consumo_vazio:.0f} kWh**")
    with col_bi2:
        expressao_f = st.text_input("Fora Vazio (ex: 90+5 ou 100+5-10)", value=st.session_state.get('exp_consumo_f', "95"), key="exp_consumo_f_widget")
        st.session_state.exp_consumo_f = expressao_f
        consumo_fora_vazio, erro_f = calcular_expressao_matematica_simples(expressao_f, "Fora Vazio")
        if erro_f: 
            st.error(erro_f)
            consumo_fora_vazio = int(st.session_state.get('exp_consumo_f_anterior_valido', 0))
        else:
            st.session_state.exp_consumo_f_anterior_valido = consumo_fora_vazio
            st.info(f"Consumo Fora Vazio considerado: **{consumo_fora_vazio:.0f} kWh**")
        
    consumo = consumo_vazio + consumo_fora_vazio

elif opcao_horaria.lower().startswith("tri"):
    st.markdown("###### Consumos Tri-Horário (kWh)")
    col_tri1, col_tri2, col_tri3 = st.columns(3)
    with col_tri1:
        # Para Vazio em Tri, vamos usar uma chave de session_state diferente para o default se necessário, ou manter a de Bi
        expressao_tv = st.text_input("Vazio (ex: 60+3 ou 70+3-10)", value=st.session_state.get('exp_consumo_v', "63"), key="exp_consumo_tv_widget") # Pode partilhar o default de 'exp_consumo_v' ou ter um 'exp_consumo_tv'
        st.session_state.exp_consumo_v = expressao_tv # Atualiza a mesma chave 'exp_consumo_v' se for partilhado
        consumo_vazio, erro_tv = calcular_expressao_matematica_simples(expressao_tv, "Vazio (Tri)")
        if erro_tv: 
            st.error(erro_tv)
            consumo_vazio = int(st.session_state.get('exp_consumo_tv_anterior_valido', 0))
        else:
            st.session_state.exp_consumo_tv_anterior_valido = consumo_vazio
            st.info(f"Consumo Vazio considerado: **{consumo_vazio:.0f} kWh**")
    with col_tri2:
        expressao_tc = st.text_input("Cheias (ex: 60+8 ou 70+8-10)", value=st.session_state.get('exp_consumo_c', "68"), key="exp_consumo_tc_widget")
        st.session_state.exp_consumo_c = expressao_tc
        consumo_cheias, erro_tc = calcular_expressao_matematica_simples(expressao_tc, "Cheias (Tri)")
        if erro_tc: 
            st.error(erro_tc)
            consumo_cheias = int(st.session_state.get('exp_consumo_tc_anterior_valido', 0))
        else:
            st.session_state.exp_consumo_tc_anterior_valido = consumo_cheias
            st.info(f"Consumo Cheias considerado: **{consumo_cheias:.0f} kWh**")
    with col_tri3:
        expressao_tp = st.text_input("Ponta (ex: 20+7 ou 30+7-10)", value=st.session_state.get('exp_consumo_p', "27"), key="exp_consumo_tp_widget")
        st.session_state.exp_consumo_p = expressao_tp
        consumo_ponta, erro_tp = calcular_expressao_matematica_simples(expressao_tp, "Ponta (Tri)")
        if erro_tp: 
            st.error(erro_tp)
            consumo_ponta = int(st.session_state.get('exp_consumo_tp_anterior_valido', 0))
        else:
            st.session_state.exp_consumo_tp_anterior_valido = consumo_ponta
            st.info(f"Consumo Ponta considerado: **{consumo_ponta:.0f} kWh**")
        
    consumo = consumo_vazio + consumo_cheias + consumo_ponta

st.write(f"Total Consumo a considerar nos cálculos: **{consumo:.0f} kWh**") # Exibir total como inteiro


# ... (Restantes inputs: Taxas DGEG/CAV, Consumos, Opções Adicionais, Meu Tarifário) ...
# Expander para as opções que são menos alteradas ou mais específicas
with st.expander("Opções Adicionais de Simulação (Tarifa Social e condicionais)"):
    st.markdown("##### Definição de Taxas Mensais")
    col_taxa1, col_taxa2 = st.columns(2)
    with col_taxa1:
        valor_dgeg_user = st.number_input(
            "Valor DGEG (€/mês)",
            min_value=0.0, step=0.01, value=0.07, # Mantém os teus defaults
            help="Taxa de Exploração da Direção-Geral de Energia e Geologia - Verifique qual o valor cobrado na sua fatura. Em condições normais, para contratos domésticos o valor é de 0,07 €/mês e os não domésticos têm o valor de 0,35 €/mês."
        )
    with col_taxa2:
        valor_cav_user = st.number_input(
            "Valor Contribuição Audiovisual (€/mês)",
            min_value=0.0, step=0.01, value=2.85, # Mantém os teus defaults
            help="Contribuição Audiovisual (CAV) - Verifique qual o valor cobrado na sua fatura. O valor normal é de 2,85 €/mês. Será 1 €/mês, para alguns casos de Tarifa Social (1º escalão de abono...) Será 0 €/mês, para consumo inferior a 400 kWh/ano."
        )
    
    # --- MOSTRAR BENEFÍCIOS APENAS SE POTÊNCIA PERMITE ---
    if potencia <= 6.9:  # <--- CONDIÇÃO ADICIONADA AQUI
        st.markdown(r"##### Benefícios e Condições Especiais (para potências $\leq 6.9$ kVA)") # Título condicional
        colx1, colx2 = st.columns(2)
        with colx1:
            # A condição interna "if potencia <= 6.9" na checkbox já não é estritamente necessária
            # pois já estamos dentro de um bloco if potencia <= 6.9, mas não faz mal mantê-la por clareza.
            tarifa_social = st.checkbox("Tarifa Social", value=st.session_state.get("chk_tarifa_social_val", False), help="Só pode ter Tarifa Social para potências até 6.9 kVA", key="chk_tarifa_social")
            st.session_state["chk_tarifa_social_val"] = tarifa_social # Guardar estado
        with colx2:
            familia_numerosa = st.checkbox("Família Numerosa", value=st.session_state.get("chk_familia_numerosa_val", False), help="300 kWh com IVA a 6% (em vez dos normais 200 kWh) para potências até 6.9 kVA", key="chk_familia_numerosa")
            st.session_state["chk_familia_numerosa_val"] = familia_numerosa # Guardar estado
    else:
        # Se potencia > 6.9, garantir que as variáveis têm um valor (False)
        # e opcionalmente limpar o estado das checkboxes se estavam ativas
        tarifa_social = False
        familia_numerosa = False
        if st.session_state.get("chk_tarifa_social_val", False): # Se estava True e agora a secção não aparece
            st.session_state.chk_tarifa_social_val = False
            # st.info("Tarifa Social desativada devido à potência selecionada (> 6.9 kVA).") # Mensagem pode ser mostrada fora do if, se desejado.
        if st.session_state.get("chk_familia_numerosa_val", False):
            st.session_state.chk_familia_numerosa_val = False
            # st.info("Benefício de Família Numerosa desativado devido à potência selecionada (> 6.9 kVA).")

    # --- CONDIÇÕES PARA ACP E CONTINENTE (dentro do expander) ---
    mostrar_widgets_acp_continente = False
    if potencia <= 20.7:
        if opcao_horaria.lower() == "simples" or opcao_horaria.lower().startswith("bi"):
            mostrar_widgets_acp_continente = True

    # Inicializar as variáveis para garantir que existem mesmo se a condição acima for falsa
    # e os widgets não forem mostrados. Os valores virão do session_state ou do default da checkbox.
    incluir_quota_acp = st.session_state.get('incluir_quota_acp_val', True)
    desconto_continente = st.session_state.get('desconto_continente_val', True)

    if mostrar_widgets_acp_continente:
        st.markdown("##### Parcerias e Descontos Específicos")
        colx3, colx4 = st.columns(2)
        with colx3:
            incluir_quota_acp = st.checkbox("Incluir Quota ACP", value=st.session_state.get('incluir_quota_acp_val', True), key='chk_acp', help="Inclui o valor da quota do ACP (4,80 €/mês) no valor do tarifário da parceria GE/ACP")
            st.session_state['incluir_quota_acp_val'] = incluir_quota_acp
        with colx4:
            desconto_continente = st.checkbox("Desconto Continente", value=st.session_state.get('desconto_continente_val', True), key='chk_continente', help="Comparar o custo total incluindo o desconto do valor do cupão Continente no tarifário Galp&Continente")
            st.session_state['desconto_continente_val'] = desconto_continente
    # else: as variáveis incluir_quota_acp e desconto_continente já foram inicializadas com os valores do session_state ou default

    # O título mantém-se para organização visual
    #st.markdown("##### Comparação Tarifários Indexados")
    # A variável é agora fixada como True, removendo a checkbox da interface
    comparar_indexados = True

# --- Fim do st.expander "Opções Adicionais de Simulação" ---

# Checkbox para ativar "O Meu Tarifário"
help_O_Meu_Tarifario = """
Para preencher os valores de acordo com o seu tarifário, ou com outro qualquer que queira comparar.

**Atenção às notas sobre as TAR e TSE.**
    """
meu_tarifario_ativo = st.checkbox(
    "**Comparar com O Meu Tarifário?**",
    key="chk_meu_tarifario_ativo",
    help=help_O_Meu_Tarifario
)


# Criação de todos_omie_inputs_utilizador_comp
todos_omie_inputs_utilizador_comp = {
    'S': st.session_state.get('omie_s_input_field', round(omie_medios_calculados.get('S',0), 2)),
    'V': st.session_state.get('omie_v_input_field', round(omie_medios_calculados.get('V',0), 2)),
    'F': st.session_state.get('omie_f_input_field', round(omie_medios_calculados.get('F',0), 2)),
    'C': st.session_state.get('omie_c_input_field', round(omie_medios_calculados.get('C',0), 2)),
    'P': st.session_state.get('omie_p_input_field', round(omie_medios_calculados.get('P',0), 2))
}
# omie_medio_simples_real_kwh já é calculado globalmente (linha 619) e pode ser passado como está
# perdas_medias já é calculado globalmente (linhas 637-678) e pode ser passado como está



#DETERMINAÇÃO DE OPÇÕES HORÁRIAS
def determinar_opcoes_horarias_destino_e_ordenacao(
    opcao_horaria_principal_str,
    potencia_kva_num,
    consumos_input_atuais_dict, # NOVO PARÂMETRO
    opcoes_horarias_existentes_lista # NOVO PARÂMETRO (passar a lista global)
):
    oh_principal_lower = opcao_horaria_principal_str.lower()
    destino_cols_nomes_unicos = []
    coluna_ordenacao_inicial_aggrid = None

    # Nomes EXATOS como na sua BD (Excel)
    SIMPLES_DB = "Simples"
    BI_DIARIO_DB = "Bi-horário - Ciclo Diário"
    BI_SEMANAL_DB = "Bi-horário - Ciclo Semanal"
    TRI_DIARIO_DB = "Tri-horário - Ciclo Diário"
    TRI_SEMANAL_DB = "Tri-horário - Ciclo Semanal"
    TRI_DIARIO_ALTA_DB = "Tri-horário > 20.7 kVA - Ciclo Diário"
    TRI_SEMANAL_ALTA_DB = "Tri-horário > 20.7 kVA - Ciclo Semanal"

    tem_consumo_simples_valido = float(consumos_input_atuais_dict.get('S', 0)) > 0
    tem_consumo_bi_valido = float(consumos_input_atuais_dict.get('V', 0)) > 0 and \
                              float(consumos_input_atuais_dict.get('F', 0)) > 0
    tem_consumo_tri_valido = float(consumos_input_atuais_dict.get('V', 0)) > 0 and \
                               float(consumos_input_atuais_dict.get('C', 0)) > 0 and \
                               float(consumos_input_atuais_dict.get('P', 0)) > 0

    # Lógica baseada na potência primeiro
    if potencia_kva_num > 20.7:
        # Para potências altas, esperamos apenas Tri-Horário se houver consumos Tri válidos
        if tem_consumo_tri_valido:
            if TRI_DIARIO_ALTA_DB in opcoes_horarias_existentes_lista:
                destino_cols_nomes_unicos.append(TRI_DIARIO_ALTA_DB)
            if TRI_SEMANAL_ALTA_DB in opcoes_horarias_existentes_lista:
                destino_cols_nomes_unicos.append(TRI_SEMANAL_ALTA_DB)
        # Se não há consumo Tri válido, ou nenhuma opção Tri de alta potência existe,
        # a lista ficará vazia, o que é correto, pois não há o que comparar.

    else: # Potência <= 20.7 kVA
        # Adicionar Simples se for a opção principal com consumo, ou se puder ser derivada de Bi/Tri
        if (oh_principal_lower == "simples" and tem_consumo_simples_valido) or \
           (oh_principal_lower.startswith("bi") and tem_consumo_bi_valido) or \
           (oh_principal_lower.startswith("tri") and tem_consumo_tri_valido):
            if SIMPLES_DB in opcoes_horarias_existentes_lista:
                destino_cols_nomes_unicos.append(SIMPLES_DB)

        # Adicionar Bi-Horário se a principal for Bi com consumo, ou Tri com consumo
        if (oh_principal_lower.startswith("bi") and tem_consumo_bi_valido) or \
           (oh_principal_lower.startswith("tri") and tem_consumo_tri_valido):
            if BI_DIARIO_DB in opcoes_horarias_existentes_lista:
                destino_cols_nomes_unicos.append(BI_DIARIO_DB)
            if BI_SEMANAL_DB in opcoes_horarias_existentes_lista:
                destino_cols_nomes_unicos.append(BI_SEMANAL_DB)

        # Adicionar Tri-Horário apenas se a principal for Tri com consumo
        if oh_principal_lower.startswith("tri") and tem_consumo_tri_valido and not oh_principal_lower.startswith("tri-horário > 20.7 kva"):
            if TRI_DIARIO_DB in opcoes_horarias_existentes_lista:
                destino_cols_nomes_unicos.append(TRI_DIARIO_DB)
            if TRI_SEMANAL_DB in opcoes_horarias_existentes_lista:
                destino_cols_nomes_unicos.append(TRI_SEMANAL_DB)

    # Ordenar as opções de destino encontradas
    ordem_preferencial = {
        SIMPLES_DB: 0, BI_DIARIO_DB: 1, BI_SEMANAL_DB: 2,
        TRI_DIARIO_DB: 3, TRI_SEMANAL_DB: 4,
        TRI_DIARIO_ALTA_DB: 5, TRI_SEMANAL_ALTA_DB: 6
    }
    destino_cols_nomes_unicos = sorted(
        list(set(destino_cols_nomes_unicos)), # Remove duplicados
        key=lambda x: ordem_preferencial.get(x, 99)
    )

    # Definir coluna de ordenação inicial
    # Tenta usar a opção horária principal do utilizador, se ela for um dos destinos válidos.
    # Caso contrário, usa a primeira opção da lista de destinos (já ordenada).
    if opcao_horaria_principal_str in destino_cols_nomes_unicos:
        coluna_ordenacao_inicial_aggrid = f"Total {opcao_horaria_principal_str} (€)"
    elif destino_cols_nomes_unicos:
        coluna_ordenacao_inicial_aggrid = f"Total {destino_cols_nomes_unicos[0]} (€)"

    colunas_aggrid_destino_formatadas = [f"Total {op_h} (€)" for op_h in destino_cols_nomes_unicos]
    return destino_cols_nomes_unicos, colunas_aggrid_destino_formatadas, coluna_ordenacao_inicial_aggrid

#Função Tarifário Fixo para comparação
def calcular_detalhes_custo_tarifario_fixo(
    dados_tarifario_linha,        # Uma Series do Pandas representando a linha do tarifário
    opcao_horaria_para_calculo, 
    consumos_repartidos_dict,   
    potencia_contratada_kva,
    dias_calculo,
    tarifa_social_ativa,
    familia_numerosa_ativa,
    valor_dgeg_user_input,
    valor_cav_user_input,
    incluir_quota_acp_input,
    desconto_continente_input,
    CONSTANTES_df,
    dias_no_mes_selecionado_dict, 
    mes_selecionado_pelo_user_str, # Já está a receber o mês
    ano_atual_calculo,             # Já está a receber o ano
    data_inicio_periodo_obj, 
    data_fim_periodo_obj
):
    """
    Calcula o custo total e os componentes de tooltip para um DADO TARIFÁRIO FIXO,
    para uma DADA OPÇÃO HORÁRIA DE DESTINO e DADOS CONSUMOS REPARTIDOS.
    Retorna um dicionário com 'Total (€)', 'NomeParaExibirAjustado' e os dicts de tooltip,
    ou None se o cálculo não for aplicável (ex: tarifário não tem preços para a opção).
    """
    try:
        # Extrair nome do comercializador de dados_tarifario_linha
        nome_comercializador_para_taxas = str(dados_tarifario_linha.get('comercializador', 'Desconhecido'))
        nome_tarifario_original = str(dados_tarifario_linha['nome'])
        nome_a_exibir_final = nome_tarifario_original # Começa com o nome original

        # --- Obter Preços e Flags do Tarifário para a OPÇÃO HORÁRIA DE CÁLCULO ---
        # Esta parte é crucial: os preços devem ser os corretos para a 'opcao_horaria_para_calculo'
        preco_energia_input_tf = {}
        oh_calc_lower = opcao_horaria_para_calculo.lower()

        if oh_calc_lower == "simples":
            preco_s = dados_tarifario_linha.get('preco_energia_simples')
            if pd.notna(preco_s):
                preco_energia_input_tf['S'] = float(preco_s)
            else:
                print(f"DEBUG CALC: {nome_tarifario_original} ({opcao_horaria_para_calculo}) - Preço Simples em falta.") # DEBUG
                return None 
        elif oh_calc_lower.startswith("bi-horário"):
            preco_v_bi = dados_tarifario_linha.get('preco_energia_vazio_bi')
            preco_f_bi = dados_tarifario_linha.get('preco_energia_fora_vazio')
            if pd.notna(preco_v_bi) and pd.notna(preco_f_bi):
                preco_energia_input_tf['V'] = float(preco_v_bi)
                preco_energia_input_tf['F'] = float(preco_f_bi)
            else:
                print(f"DEBUG CALC: {nome_tarifario_original} ({opcao_horaria_para_calculo}) - Preços Bi-Horário em falta/inválidos (Vazio: {preco_v_bi}, ForaVazio: {preco_f_bi}).") # DEBUG
                return None 
        elif oh_calc_lower.startswith("tri-horário"): # Cobre variantes normais e >20.7kVA
            if pd.notna(dados_tarifario_linha.get('preco_energia_vazio_tri')) and \
               pd.notna(dados_tarifario_linha.get('preco_energia_cheias')) and \
               pd.notna(dados_tarifario_linha.get('preco_energia_ponta')):
                preco_energia_input_tf['V'] = float(dados_tarifario_linha.get('preco_energia_vazio_tri', 0.0))
                preco_energia_input_tf['C'] = float(dados_tarifario_linha.get('preco_energia_cheias', 0.0))
                preco_energia_input_tf['P'] = float(dados_tarifario_linha.get('preco_energia_ponta', 0.0))
            else: # Tarifário não tem oferta Tri-Horário completa
                return None
        else:
            return None # Opção horária de cálculo desconhecida

        preco_potencia_input_tf = float(dados_tarifario_linha.get('preco_potencia_dia', 0.0))
        tar_incluida_energia_tf = dados_tarifario_linha.get('tar_incluida_energia', True)
        tar_incluida_potencia_tf = dados_tarifario_linha.get('tar_incluida_potencia', True)
        financiamento_tse_incluido_tf = dados_tarifario_linha.get('financiamento_tse_incluido', True)

        # --- Passo 1: Identificar Componentes Base (Sem IVA, Sem TS) ---
        tar_energia_regulada_tf = {}
        for periodo_consumo_key in consumos_repartidos_dict.keys(): # S, V, F, C, P
            tar_energia_regulada_tf[periodo_consumo_key] = obter_tar_energia_periodo(
                opcao_horaria_para_calculo, periodo_consumo_key, potencia_contratada_kva, CONSTANTES_df
            )

        tar_potencia_regulada_tf = obter_tar_dia(potencia_contratada_kva, CONSTANTES_df)

        preco_comercializador_energia_tf = {}
        for periodo_preco_key, preco_val_tf in preco_energia_input_tf.items():
            if periodo_preco_key not in consumos_repartidos_dict: continue # Só se houver consumo nesse período
            if tar_incluida_energia_tf:
                preco_comercializador_energia_tf[periodo_preco_key] = preco_val_tf - tar_energia_regulada_tf.get(periodo_preco_key, 0.0)
            else:
                preco_comercializador_energia_tf[periodo_preco_key] = preco_val_tf
        
        if tar_incluida_potencia_tf:
            preco_comercializador_potencia_tf = preco_potencia_input_tf - tar_potencia_regulada_tf
        else:
            preco_comercializador_potencia_tf = preco_potencia_input_tf
        preco_comercializador_potencia_tf = max(0.0, preco_comercializador_potencia_tf)

        financiamento_tse_a_adicionar_tf = FINANCIAMENTO_TSE_VAL if not financiamento_tse_incluido_tf else 0.0

        # --- Passo 2: Calcular Componentes TAR Finais (Com Desconto TS, Sem IVA) ---
        tar_energia_final_tf = {}
        tar_potencia_final_dia_tf = tar_potencia_regulada_tf
        desconto_ts_energia_aplicado_val = 0.0
        desconto_ts_potencia_aplicado_val = 0.0

        if tarifa_social_ativa:
            desconto_ts_energia_bruto = obter_constante('Desconto TS Energia', CONSTANTES_df)
            desconto_ts_potencia_dia_bruto = obter_constante(f'Desconto TS Potencia {potencia_contratada_kva}', CONSTANTES_df)
            for periodo_calc, tar_reg_val in tar_energia_regulada_tf.items():
                tar_energia_final_tf[periodo_calc] = tar_reg_val - desconto_ts_energia_bruto
            desconto_ts_energia_aplicado_val = desconto_ts_energia_bruto # Para tooltip
            
            tar_potencia_final_dia_tf = max(0.0, tar_potencia_regulada_tf - desconto_ts_potencia_dia_bruto)
            desconto_ts_potencia_aplicado_val = min(tar_potencia_regulada_tf, desconto_ts_potencia_dia_bruto) # Para tooltip
        else:
            tar_energia_final_tf = tar_energia_regulada_tf.copy()

        # --- Passo 3: Calcular Preço Final Energia (€/kWh, Sem IVA) ---
        preco_energia_final_sem_iva_tf_dict = {}
        for periodo_calc in consumos_repartidos_dict.keys(): # Iterar sobre os períodos COM CONSUMO
            if periodo_calc in preco_comercializador_energia_tf: # Verificar se há preço definido para este período
                preco_energia_final_sem_iva_tf_dict[periodo_calc] = (
                    preco_comercializador_energia_tf.get(periodo_calc, 0.0) +
                    tar_energia_final_tf.get(periodo_calc, 0.0) +
                    financiamento_tse_a_adicionar_tf
                )

        # --- Passo 4: Calcular Componentes Finais Potência (€/dia, Sem IVA) ---
        preco_comercializador_potencia_final_sem_iva_tf = preco_comercializador_potencia_tf

        # --- Passo 5 & 6: Calcular Custo Total Energia e Potência (Com IVA) ---
        consumo_total_neste_oh = sum(float(v or 0) for v in consumos_repartidos_dict.values())

        decomposicao_custo_energia_tf = calcular_custo_energia_com_iva(
            consumo_total_neste_oh,
            preco_energia_final_sem_iva_tf_dict.get('S') if opcao_horaria_para_calculo.lower() == "simples" else None,
            {p: v for p, v in preco_energia_final_sem_iva_tf_dict.items() if p != 'S'},
            dias_calculo, potencia_contratada_kva, opcao_horaria_para_calculo,
            consumos_repartidos_dict, # Usar os consumos repartidos para esta opção horária
            familia_numerosa_ativa
        )
        
        decomposicao_custo_potencia_tf_calc = calcular_custo_potencia_com_iva_final(
            preco_comercializador_potencia_final_sem_iva_tf,
            tar_potencia_final_dia_tf,  # <--- USE DIRETAMENTE A VARIÁVEL CORRETA
            dias_calculo, potencia_contratada_kva
        )

        # --- Passo 7: Calcular Taxas Adicionais ---
        decomposicao_taxas_tf = calcular_taxas_adicionais(
            consumo_total_neste_oh, dias_calculo, tarifa_social_ativa,
            valor_dgeg_user_input, valor_cav_user_input,
            nome_comercializador_para_taxas,     # <--- PASSAR O NOME DO COMERCIALIZADOR
            mes_selecionado_pelo_user_str,       # <--- PASSAR O MÊS
            ano_atual_calculo,                   # <--- PASSAR O ANO
            valor_iec=0.001 # O default já está na função
        )


        # --- Passo 8: Calcular Custo Total Final e aplicar descontos específicos ---
        custo_total_antes_desc_fatura_tf = (
            decomposicao_custo_energia_tf['custo_com_iva'] +
            decomposicao_custo_potencia_tf_calc['custo_com_iva'] +
            decomposicao_taxas_tf['custo_com_iva']
        )

        # Lógica de mês completo para descontos mensais
        e_mes_completo_selecionado_calc = False
        if dias_calculo == dias_no_mes_selecionado_dict.get(mes_selecionado_pelo_user_str, 0):
            e_mes_completo_selecionado_calc = True
        
        # Desconto de fatura do Excel
        desconto_fatura_mensal_excel = float(dados_tarifario_linha.get('desconto_fatura_mes', 0.0) or 0.0)
        desconto_fatura_periodo_aplicado = 0.0
        if desconto_fatura_mensal_excel > 0:
            nome_a_exibir_final += f" (+desc. fat. {desconto_fatura_mensal_excel:.2f}€/mês)"
            desconto_fatura_periodo_aplicado = (desconto_fatura_mensal_excel / 30.0) * dias_calculo if not e_mes_completo_selecionado_calc else desconto_fatura_mensal_excel
        
        custo_apos_desc_fatura_excel = custo_total_antes_desc_fatura_tf - desconto_fatura_periodo_aplicado
        
        # Quota ACP
        custo_apos_acp = custo_apos_desc_fatura_excel
        quota_acp_periodo_aplicada = 0.0
        if incluir_quota_acp_input and nome_tarifario_original.startswith("Goldenergy - ACP"):
            quota_acp_a_aplicar = (VALOR_QUOTA_ACP_MENSAL / 30.0) * dias_calculo if not e_mes_completo_selecionado_calc else VALOR_QUOTA_ACP_MENSAL
            custo_apos_acp += quota_acp_a_aplicar
            nome_a_exibir_final += f" (INCLUI Quota ACP - {VALOR_QUOTA_ACP_MENSAL:.2f} €/mês)"
            quota_acp_periodo_aplicada = quota_acp_a_aplicar

        # Desconto MEO
        custo_antes_desconto_meo = custo_apos_acp
        desconto_meo_periodo_aplicado = 0.0
        if "meo energia - tarifa fixa - clientes meo" in nome_tarifario_original.lower() and \
           (consumo_total_neste_oh / dias_calculo * 30.0 if dias_calculo > 0 else 0) >= 216:
            desconto_meo_mensal_base = 0.0
            if opcao_horaria_para_calculo.lower() == "simples": desconto_meo_mensal_base = 2.95
            elif opcao_horaria_para_calculo.lower().startswith("bi-horário"): desconto_meo_mensal_base = 3.50
            elif opcao_horaria_para_calculo.lower().startswith("tri-horário"): desconto_meo_mensal_base = 6.27
            if desconto_meo_mensal_base > 0 and dias_calculo > 0:
                desconto_meo_periodo_aplicado = (desconto_meo_mensal_base / 30.0) * dias_calculo
                custo_antes_desconto_meo -= desconto_meo_periodo_aplicado
                nome_a_exibir_final += f" (Desc. MEO {desconto_meo_periodo_aplicado:.2f}€ incl.)"
        
        # Desconto Continente
        custo_base_para_continente = custo_antes_desconto_meo
        custo_total_final = custo_base_para_continente 
        valor_X_desconto_continente_aplicado = 0.0

        if desconto_continente_input and nome_tarifario_original.startswith("Galp & Continente"):
    
            # PASSO ADICIONAL: CALCULAR O CUSTO BRUTO (SEM TARIFA SOCIAL) APENAS PARA ESTE DESCONTO
    
            # 1. Preço unitário bruto da energia (sem IVA e sem desconto TS)
            preco_energia_bruto_sem_iva = {}
            for p in consumos_repartidos_dict.keys():
                if p in preco_comercializador_energia_tf:
                    preco_energia_bruto_sem_iva[p] = (
                        preco_comercializador_energia_tf[p] + 
                        tar_energia_regulada_tf.get(p, 0.0) + # <--- USA A TAR BRUTA, sem desconto TS
                        financiamento_tse_a_adicionar_tf
                    )

            # 2. Preço unitário bruto da potência (sem IVA e sem desconto TS)
            preco_comerc_pot_bruto = preco_comercializador_potencia_tf
            tar_potencia_bruta = tar_potencia_regulada_tf # <--- USA A TAR BRUTA, sem desconto TS

            # 3. Calcular o custo bruto COM IVA para a energia e potência
            custo_energia_bruto_cIVA = calcular_custo_energia_com_iva(
                consumo_total_neste_oh,
                preco_energia_bruto_sem_iva.get('S'),
                {k: v for k, v in preco_energia_bruto_sem_iva.items() if k != 'S'},
                dias_calculo, potencia_contratada_kva, opcao_horaria_para_calculo,
                consumos_repartidos_dict, familia_numerosa_ativa
            )
            custo_potencia_bruto_cIVA = calcular_custo_potencia_com_iva_final(
                preco_comerc_pot_bruto,
                tar_potencia_bruta,
                dias_calculo, potencia_contratada_kva
            )
    
            # 4. Calcular o valor do cupão sobre os custos brutos
            valor_X_desconto_continente_aplicado = (custo_energia_bruto_cIVA['custo_com_iva'] + custo_potencia_bruto_cIVA['custo_com_iva']) * 0.10
    
            # 5. Aplicar o valor do cupão ao custo final (que já tem o desconto TS, se aplicável)
            custo_total_final = custo_base_para_continente - valor_X_desconto_continente_aplicado
            nome_a_exibir_final += f" (INCLUI desc. Cont. de {valor_X_desconto_continente_aplicado:.2f}€, s/ desc. Cont.={custo_base_para_continente:.2f}€)"        
        # --- Construir Dicionários de Tooltip ---
        # Tooltip Energia
        componentes_tooltip_energia_dict = {}
        for p_key_tt_energia in preco_energia_final_sem_iva_tf_dict.keys():
            componentes_tooltip_energia_dict[f'tooltip_energia_{p_key_tt_energia}_comerc_sem_tar'] = preco_comercializador_energia_tf.get(p_key_tt_energia, 0.0)
            componentes_tooltip_energia_dict[f'tooltip_energia_{p_key_tt_energia}_tar_bruta'] = tar_energia_regulada_tf.get(p_key_tt_energia, 0.0)
            componentes_tooltip_energia_dict[f'tooltip_energia_{p_key_tt_energia}_tse_declarado_incluido'] = financiamento_tse_incluido_tf
            componentes_tooltip_energia_dict[f'tooltip_energia_{p_key_tt_energia}_tse_valor_nominal'] = FINANCIAMENTO_TSE_VAL
            componentes_tooltip_energia_dict[f'tooltip_energia_{p_key_tt_energia}_ts_aplicada_flag'] = tarifa_social_ativa
            componentes_tooltip_energia_dict[f'tooltip_energia_{p_key_tt_energia}_ts_desconto_valor'] = obter_constante('Desconto TS Energia', CONSTANTES_df) if tarifa_social_ativa else 0.0
        
        # Tooltip Potência
        componentes_tooltip_potencia_dict = {
            'tooltip_pot_comerc_sem_tar': preco_comercializador_potencia_tf, # Já após desconto %, mas antes de TS
            'tooltip_pot_tar_bruta': tar_potencia_regulada_tf,
            'tooltip_pot_ts_aplicada': tarifa_social_ativa,
            'tooltip_pot_desconto_ts_valor': desconto_ts_potencia_aplicado_val # Valor efetivo do desconto TS na TAR
        }

        # Preço unitário da potência s/IVA (comercializador + TAR final)
        preco_unit_potencia_siva_tf = preco_comercializador_potencia_final_sem_iva_tf + tar_potencia_final_dia_tf # Esta é a soma correta

        # Tooltip Custo Total
        componentes_tooltip_total_dict = {
            'tt_cte_energia_siva': decomposicao_custo_energia_tf['custo_sem_iva'],
            'tt_cte_potencia_siva': decomposicao_custo_potencia_tf_calc['custo_sem_iva'],
            'tt_cte_iec_siva': decomposicao_taxas_tf['iec_sem_iva'],
            'tt_cte_dgeg_siva': decomposicao_taxas_tf['dgeg_sem_iva'],
            'tt_cte_cav_siva': decomposicao_taxas_tf['cav_sem_iva'],
            'tt_cte_total_siva': decomposicao_custo_energia_tf['custo_sem_iva'] + decomposicao_custo_potencia_tf_calc['custo_sem_iva'] + decomposicao_taxas_tf['custo_sem_iva'],
            'tt_cte_valor_iva_6_total': decomposicao_custo_energia_tf['valor_iva_6'] + decomposicao_custo_potencia_tf_calc['valor_iva_6'] + decomposicao_taxas_tf['valor_iva_6'],
            'tt_cte_valor_iva_23_total': decomposicao_custo_energia_tf['valor_iva_23'] + decomposicao_custo_potencia_tf_calc['valor_iva_23'] + decomposicao_taxas_tf['valor_iva_23'],
            'tt_cte_subtotal_civa': custo_total_antes_desc_fatura_tf,
            'tt_cte_desc_finais_valor': desconto_fatura_periodo_aplicado + desconto_meo_periodo_aplicado + valor_X_desconto_continente_aplicado,
            'tt_cte_acres_finais_valor': quota_acp_periodo_aplicada,
            **{f"tt_preco_unit_energia_{p}_siva": v for p, v in preco_energia_final_sem_iva_tf_dict.items()},
            'tt_preco_unit_potencia_siva': preco_unit_potencia_siva_tf
        }
        
        return {
            'Total (€)': custo_total_final,
            'NomeParaExibirAjustado': nome_a_exibir_final, 
            'componentes_tooltip_custo_total_dict': componentes_tooltip_total_dict,
            'componentes_tooltip_energia_dict': componentes_tooltip_energia_dict,
            'componentes_tooltip_potencia_dict': componentes_tooltip_potencia_dict
        }
    
    except Exception as e:
        st.error(f"Erro ao calcular custo para {dados_tarifario_linha.get('nome', 'Tarifário Desconhecido')} na opção {opcao_horaria_para_calculo}: {e}")
        return None


#Função Tarifário Indexado para comparação
def calcular_detalhes_custo_tarifario_indexado(
    dados_tarifario_indexado_linha,
    opcao_horaria_para_calculo, # A opção de destino (Simples, Bi, Tri) - ex: "Bi-horário - Ciclo Diário"
    opcao_horaria_principal_global, # NOVO PARÂMETRO: A opção principal selecionada pelo user (st.selectbox)
    consumos_repartidos_dict,
    potencia_contratada_kva,
    dias_calculo,
    tarifa_social_ativa,
    familia_numerosa_ativa,
    valor_dgeg_user_input,
    valor_cav_user_input,
    CONSTANTES_df,
    df_omie_ajustado_para_calculo,
    perdas_medias_dict_global,
    todos_omie_inputs_user_global, # Valores dos st.number_input OMIE (S,V,F,C,P)
    omie_medios_calculados_para_todos_ciclos_global, # NOVO: OMIEs médios CALCULADOS para todos os ciclos
    omie_medio_simples_real_kwh_para_luzigas_idx,
    dias_no_mes_selecionado_dict,
    mes_selecionado_pelo_user_str,
    ano_atual_calculo,
    data_inicio_periodo_obj,
    data_fim_periodo_obj
):
    try:
        nome_tarifario_original = str(dados_tarifario_indexado_linha['nome'])
        tipo_tarifario_original = str(dados_tarifario_indexado_linha['tipo'])
        formula_energia_str = str(dados_tarifario_indexado_linha.get('formula_calculo', ''))
        nome_a_exibir_final = nome_tarifario_original

        precos_energia_base_kwh_nesta_oh = {} # Preços base calculados para a opcao_horaria_para_calculo
        oh_calc_lower = opcao_horaria_para_calculo.lower() # ex: "simples", "bi-horário - ciclo diário"
        constantes_dict_local = dict(zip(CONSTANTES_df["constante"], CONSTANTES_df["valor_unitário"]))
        
        # Valores para os preços de energia indexada (resultados do cálculo abaixo)
        preco_idx_s, preco_idx_v, preco_idx_f, preco_idx_c, preco_idx_p = None, None, None, None, None

        # --- BLOCO 1: Cálculo para Indexados Quarto-Horários (BTN ou Luzboa "BTN SPOTDEF") ---
        if 'BTN' in formula_energia_str or nome_tarifario_original == "Luzboa - BTN SPOTDEF":
            soma_calculo_periodo = {p_key: 0.0 for p_key in ['S', 'V', 'F', 'C', 'P']} # Acumuladores para todos os períodos possíveis
            soma_perfil_periodo = {p_key: 0.0 for p_key in ['S', 'V', 'F', 'C', 'P']}

            # Determinar coluna de ciclo e perfil com base na opcao_horaria_para_calculo
            # Nota: opcao_horaria_para_calculo é o nome DB, ex: "Bi-horário - Ciclo Diário"
            coluna_ciclo_qh = None
            if oh_calc_lower.startswith("bi-horário"):
                coluna_ciclo_qh = 'BD' if "diário" in oh_calc_lower else 'BS'
            elif oh_calc_lower.startswith("tri-horário") and not oh_calc_lower.startswith("tri-horário > 20.7 kva"):
                coluna_ciclo_qh = 'TD' if "diário" in oh_calc_lower else 'TS'
            elif oh_calc_lower.startswith("tri-horário > 20.7 kva"):
                 coluna_ciclo_qh = 'TD' if "diário" in oh_calc_lower else 'TS' # Mesma lógica de ciclo

            consumo_total_para_perfil_nesta_oh = sum(v for v in consumos_repartidos_dict.values() if v is not None)
            perfil_nome_str = obter_perfil(consumo_total_para_perfil_nesta_oh, dias_calculo, potencia_contratada_kva) # perfil_A, perfil_B, perfil_C
            perfil_coluna_qh = f"BTN_{perfil_nome_str.split('_')[1].upper()}" # BTN_A, BTN_B, BTN_C

            if perfil_coluna_qh not in df_omie_ajustado_para_calculo.columns:
                # st.warning(f"DEBUG COMP: Coluna de perfil '{perfil_coluna_qh}' não encontrada para '{nome_tarifario_original}' em '{opcao_horaria_para_calculo}'. Energia será zero.")
                # Definir preços como zero se o perfil não existir no DF OMIE
                for p_key_cons in consumos_repartidos_dict.keys(): precos_energia_base_kwh_nesta_oh[p_key_cons] = 0.0
            
            elif nome_tarifario_original == "Luzboa - BTN SPOTDEF":
                # Lógica específica Luzboa (usa médias horárias simples, não ponderadas por perfil BTN)
                soma_luzboa_p = {k: 0.0 for k in ['S', 'V', 'F', 'C', 'P']}
                count_luzboa_p = {k: 0 for k in ['S', 'V', 'F', 'C', 'P']}

                for _, row_omie in df_omie_ajustado_para_calculo.iterrows():
                    if not all(k_luzboa in row_omie and pd.notna(row_omie[k_luzboa]) for k_luzboa in ['OMIE', 'Perdas']): continue
                    omie_val_l = row_omie['OMIE'] / 1000.0
                    perdas_val_l = row_omie['Perdas']
                    cgs_luzboa = constantes_dict_local.get('Luzboa_CGS', 0.0)
                    fa_luzboa = constantes_dict_local.get('Luzboa_FA', 1.0)
                    kp_luzboa = constantes_dict_local.get('Luzboa_Kp', 0.0)
                    valor_hora_luzboa = (omie_val_l + cgs_luzboa) * perdas_val_l * fa_luzboa + kp_luzboa

                    soma_luzboa_p['S'] += valor_hora_luzboa; count_luzboa_p['S'] += 1
                    
                    if coluna_ciclo_qh and coluna_ciclo_qh in row_omie and pd.notna(row_omie[coluna_ciclo_qh]):
                        ciclo_hora_l = row_omie[coluna_ciclo_qh] # V, F, C, P
                        if ciclo_hora_l in soma_luzboa_p: # Para V,F,C,P
                             soma_luzboa_p[ciclo_hora_l] += valor_hora_luzboa
                             count_luzboa_p[ciclo_hora_l] += 1
                
                prec_luzboa = 4
                if oh_calc_lower == "simples":
                    preco_idx_s = round(soma_luzboa_p['S'] / count_luzboa_p['S'], prec_luzboa) if count_luzboa_p['S'] > 0 else 0.0
                elif oh_calc_lower.startswith("bi-horário"):
                    preco_idx_v = round(soma_luzboa_p['V'] / count_luzboa_p['V'], prec_luzboa) if count_luzboa_p['V'] > 0 else 0.0
                    preco_idx_f = round(soma_luzboa_p['F'] / count_luzboa_p['F'], prec_luzboa) if count_luzboa_p['F'] > 0 else 0.0
                elif oh_calc_lower.startswith("tri-horário"):
                    preco_idx_v = round(soma_luzboa_p['V'] / count_luzboa_p['V'], prec_luzboa) if count_luzboa_p['V'] > 0 else 0.0
                    preco_idx_c = round(soma_luzboa_p['C'] / count_luzboa_p['C'], prec_luzboa) if count_luzboa_p['C'] > 0 else 0.0
                    preco_idx_p = round(soma_luzboa_p['P'] / count_luzboa_p['P'], prec_luzboa) if count_luzboa_p['P'] > 0 else 0.0

            else: # Outros Tarifários Quarto-Horários (Coopernico, Repsol, Galp, etc.)
                # Precisam da coluna de ciclo para V,F,C,P
                cycle_column_ok_qh = True
                if oh_calc_lower != "simples":
                    if not coluna_ciclo_qh or coluna_ciclo_qh not in df_omie_ajustado_para_calculo.columns:
                        # st.warning(f"DEBUG COMP: Coluna ciclo '{coluna_ciclo_qh}' em falta para '{nome_tarifario_original}' em '{opcao_horaria_para_calculo}'. Preços V/F/C/P serão 0.")
                        cycle_column_ok_qh = False
                        # Se a coluna de ciclo não existe, os preços para V,F,C,P serão zero, Simples ainda pode ser calculado.
                        preco_idx_v, preco_idx_f, preco_idx_c, preco_idx_p = 0.0, 0.0, 0.0, 0.0

                for _, row_omie in df_omie_ajustado_para_calculo.iterrows():
                    required_cols_qh = ['OMIE', 'Perdas', perfil_coluna_qh]
                    if not all(k_qh in row_omie and pd.notna(row_omie[k_qh]) for k_qh in required_cols_qh): continue
                    
                    omie_val_qh = row_omie['OMIE'] / 1000.0
                    perdas_val_qh = row_omie['Perdas']
                    perfil_val_qh = row_omie[perfil_coluna_qh]
                    if perfil_val_qh <= 0: continue

                    calculo_instantaneo_sem_perfil_qh = 0.0
                    # --- Fórmulas específicas BTN (Quarto-Horário) ---
                    if nome_tarifario_original == "Coopérnico Base 2.0": calculo_instantaneo_sem_perfil_qh = (omie_val_qh + constantes_dict_local.get('Coop_CS_CR', 0.0) + constantes_dict_local.get('Coop_K', 0.0)) * perdas_val_qh
                    elif nome_tarifario_original == "Repsol - Leve Sem Mais": calculo_instantaneo_sem_perfil_qh = (omie_val_qh * perdas_val_qh * constantes_dict_local.get('Repsol_FA', 0.0) + constantes_dict_local.get('Repsol_Q_Tarifa', 0.0))
                    elif nome_tarifario_original == "Repsol - Leve PRO Sem Mais": calculo_instantaneo_sem_perfil_qh = (omie_val_qh * perdas_val_qh * constantes_dict_local.get('Repsol_FA', 0.0) + constantes_dict_local.get('Repsol_Q_Tarifa_Pro', 0.0))
                    elif nome_tarifario_original == "Galp - Plano Flexível / Dinâmico": calculo_instantaneo_sem_perfil_qh = (omie_val_qh + constantes_dict_local.get('Galp_Ci', 0.0)) * perdas_val_qh
                    elif nome_tarifario_original == "Alfa Energia - ALFA POWER INDEX BTN": calculo_instantaneo_sem_perfil_qh = ((omie_val_qh + constantes_dict_local.get('Alfa_CGS', 0.0)) * perdas_val_qh + constantes_dict_local.get('Alfa_K', 0.0))
                    elif nome_tarifario_original == "Plenitude - Tendência": calculo_instantaneo_sem_perfil_qh = ((omie_val_qh + constantes_dict_local.get('Plenitude_CGS', 0.0) + constantes_dict_local.get('Plenitude_GDOs', 0.0)) * perdas_val_qh + constantes_dict_local.get('Plenitude_Fee', 0.0))
                    elif nome_tarifario_original == "Meo Energia - Tarifa Variável": calculo_instantaneo_sem_perfil_qh = (omie_val_qh + constantes_dict_local.get('Meo_K', 0.0)) * perdas_val_qh
                    elif nome_tarifario_original == "EDP - Eletricidade Indexada Horária": calculo_instantaneo_sem_perfil_qh = (omie_val_qh * perdas_val_qh * constantes_dict_local.get('EDP_H_K1', 1.0) + constantes_dict_local.get('EDP_H_K2', 0.0))
                    elif nome_tarifario_original == "EZU - Coletiva": calculo_instantaneo_sem_perfil_qh = (omie_val_qh + constantes_dict_local.get('EZU_K', 0.0) + constantes_dict_local.get('EZU_CGS', 0.0)) * perdas_val_qh
                    elif nome_tarifario_original == "G9 - Smart Dynamic": calculo_instantaneo_sem_perfil_qh = (omie_val_qh * constantes_dict_local.get('G9_FA', 0.0) * perdas_val_qh + constantes_dict_local.get('G9_CGS', 0.0) + constantes_dict_local.get('G9_AC', 0.0))
                    elif nome_tarifario_original == "Iberdrola - Simples Indexado Dinâmico": calculo_instantaneo_sem_perfil_qh = (omie_val_qh * perdas_val_qh + constantes_dict_local.get("Iberdrola_Q", 0.0) + constantes_dict_local.get('Iberdrola_mFRR', 0.0))
                    else: calculo_instantaneo_sem_perfil_qh = omie_val_qh * perdas_val_qh # Fallback

                    soma_calculo_periodo['S'] += calculo_instantaneo_sem_perfil_qh * perfil_val_qh
                    soma_perfil_periodo['S'] += perfil_val_qh
                    
                    if cycle_column_ok_qh and coluna_ciclo_qh and coluna_ciclo_qh in row_omie and pd.notna(row_omie[coluna_ciclo_qh]):
                        ciclo_hora_qh = row_omie[coluna_ciclo_qh] # V, F, C, P
                        if ciclo_hora_qh in soma_calculo_periodo: # Para V,F,C,P
                             soma_calculo_periodo[ciclo_hora_qh] += calculo_instantaneo_sem_perfil_qh * perfil_val_qh
                             soma_perfil_periodo[ciclo_hora_qh] += perfil_val_qh
                
                prec_qh = 4 # Aumentar precisão interna para cálculos
                # Calcular preços médios ponderados para cada período da opcao_horaria_para_calculo
                if nome_tarifario_original in ["Repsol - Leve Sem Mais", "Repsol - Leve PRO Sem Mais"]:
                    # Repsol usa sempre o preço calculado como se fosse Simples para todos os períodos
                    preco_simples_calc_repsol = round(soma_calculo_periodo['S'] / soma_perfil_periodo['S'], prec_qh) if soma_perfil_periodo['S'] > 0 else 0.0
                    preco_idx_s = preco_simples_calc_repsol
                    preco_idx_v = preco_simples_calc_repsol
                    preco_idx_f = preco_simples_calc_repsol
                    preco_idx_c = preco_simples_calc_repsol
                    preco_idx_p = preco_simples_calc_repsol
                else: # Outros BTN
                    if oh_calc_lower == "simples":
                        preco_idx_s = round(soma_calculo_periodo['S'] / soma_perfil_periodo['S'], prec_qh) if soma_perfil_periodo['S'] > 0 else 0.0
                    elif oh_calc_lower.startswith("bi-horário"):
                        preco_idx_v = round(soma_calculo_periodo['V'] / soma_perfil_periodo['V'], prec_qh) if soma_perfil_periodo['V'] > 0 else 0.0
                        preco_idx_f = round(soma_calculo_periodo['F'] / soma_perfil_periodo['F'], prec_qh) if soma_perfil_periodo['F'] > 0 else 0.0
                    elif oh_calc_lower.startswith("tri-horário"):
                        preco_idx_v = round(soma_calculo_periodo['V'] / soma_perfil_periodo['V'], prec_qh) if soma_perfil_periodo['V'] > 0 else 0.0
                        preco_idx_c = round(soma_calculo_periodo['C'] / soma_perfil_periodo['C'], prec_qh) if soma_perfil_periodo['C'] > 0 else 0.0
                        preco_idx_p = round(soma_calculo_periodo['P'] / soma_perfil_periodo['P'], prec_qh) if soma_perfil_periodo['P'] > 0 else 0.0

# --- BLOCO 2: Cálculo para Indexados Média ---
        else: # Tarifários de Média
            prec_media = 4 # Conforme o seu ficheiro .py

            # Iterar sobre os períodos RELEVANTES para a opcao_horaria_para_calculo (destino)
            periodos_relevantes_para_destino = []
            if oh_calc_lower == "simples":
                periodos_relevantes_para_destino.append('S')
            elif oh_calc_lower.startswith("bi-horário"):
                periodos_relevantes_para_destino.extend(['V', 'F'])
            elif oh_calc_lower.startswith("tri-horário"):
                periodos_relevantes_para_destino.extend(['V', 'C', 'P'])

            for p_key_destino in periodos_relevantes_para_destino:
                omie_mwh_final_para_formula = 0.0
                
                # Determinar ciclo real da opção de DESTINO (S, BD, BS, TD, TS)
                ciclo_real_oh_destino = ""
                if oh_calc_lower == "simples":
                    ciclo_real_oh_destino = "S"
                elif oh_calc_lower.startswith("bi-horário"):
                    ciclo_real_oh_destino = "BD" if "diário" in oh_calc_lower else "BS"
                elif oh_calc_lower.startswith("tri-horário"): # Cobre >20.7kVA também para ciclo
                    ciclo_real_oh_destino = "TD" if "diário" in oh_calc_lower else "TS"

                # 1. Obter OMIE MWh base (CALCULADO para o ciclo/período de DESTINO)
                chave_omie_calculado_destino = ""
                if ciclo_real_oh_destino == "S":
                    chave_omie_calculado_destino = "S"
                elif ciclo_real_oh_destino: # Para BD, BS, TD, TS
                    chave_omie_calculado_destino = f"{ciclo_real_oh_destino}_{p_key_destino}"
                
                omie_mwh_base_calculado = omie_medios_calculados_para_todos_ciclos_global.get(chave_omie_calculado_destino, 0.0)

                # 2. Verificar se OMIE manual da OPÇÃO PRINCIPAL deve sobrepor-se
                if opcao_horaria_para_calculo == opcao_horaria_principal_global and \
                   st.session_state.omie_foi_editado_manualmente.get(p_key_destino, False):
                    # Usar o valor do input manual (que corresponde à opção principal)
                    omie_mwh_final_para_formula = todos_omie_inputs_user_global.get(p_key_destino, omie_mwh_base_calculado)
                else:
                    # Usar o OMIE calculado específico para o ciclo de destino
                    omie_mwh_final_para_formula = omie_mwh_base_calculado
                
                omie_kwh_a_usar_na_formula = omie_mwh_final_para_formula / 1000.0
                
                # Lógica de PERDAS (deve usar perdas_medias_dict_global e o ciclo_real_oh_destino)
                perdas_a_usar_val_media = 1.0 # Default
                if nome_tarifario_original in ["LUZiGÁS - Energy 8.8", "LUZiGÁS - Dinâmico Poupança +", "Ibelectra - Solução Família"]:
                    # Perdas Anuais por período da opção de destino
                    chave_perda_anual = f'Perdas_Anual_{ciclo_real_oh_destino}_{p_key_destino}' if ciclo_real_oh_destino != "S" else 'Perdas_Anual_S'
                    perdas_a_usar_val_media = perdas_medias_dict_global.get(chave_perda_anual, 1.0)
                elif nome_tarifario_original == "G9 - Smart Index":
                    # Perdas do PERÍODO SELECIONADO por período da opção de destino
                    chave_perda_periodo = f'Perdas_M_{ciclo_real_oh_destino}_{p_key_destino}' if ciclo_real_oh_destino != "S" else 'Perdas_M_S'
                    perdas_a_usar_val_media = perdas_medias_dict_global.get(chave_perda_periodo, 1.0)
                # Para outros, a lógica de perdas está nas fórmulas ou são constantes.

                # --- Fórmulas específicas para Tarifários de Média ---
                temp_preco_calculado = 0.0
                # OMIE para LuziGás é especial (usa OMIE real simples)
                omie_para_luzigas_kwh = omie_medio_simples_real_kwh_para_luzigas_idx
                
                if nome_tarifario_original == "Iberdrola - Simples Indexado":
                    if p_key_destino == 'S': temp_preco_calculado = omie_kwh_a_usar_na_formula * constantes_dict_local.get('Iberdrola_Perdas', 1.0) + constantes_dict_local.get("Iberdrola_Q", 0.0) + constantes_dict_local.get('Iberdrola_mFRR', 0.0)
                elif nome_tarifario_original == "Goldenergy - Tarifário Indexado 100%":
                    if p_key_destino == 'S':
                        mes_num_calculo = list(dias_no_mes_selecionado_dict.keys()).index(mes_selecionado_pelo_user_str) + 1
                        perdas_mensais_ge_map = {1:1.29,2:1.18,3:1.18,4:1.15,5:1.11,6:1.10,7:1.15,8:1.13,9:1.10,10:1.10,11:1.16,12:1.25}
                        perdas_mensais_ge = perdas_mensais_ge_map.get(mes_num_calculo, 1.0)
                        temp_preco_calculado = omie_kwh_a_usar_na_formula * perdas_mensais_ge + constantes_dict_local.get('GE_Q_Tarifa', 0.0) + constantes_dict_local.get('GE_CG', 0.0)
                elif nome_tarifario_original == "Endesa - Tarifa Indexada":
                    if p_key_destino == 'S': temp_preco_calculado = omie_kwh_a_usar_na_formula + constantes_dict_local.get('Endesa_A_S', 0.0)
                    elif p_key_destino == 'V': temp_preco_calculado = omie_kwh_a_usar_na_formula + constantes_dict_local.get('Endesa_A_V', 0.0)
                    elif p_key_destino == 'F': temp_preco_calculado = omie_kwh_a_usar_na_formula + constantes_dict_local.get('Endesa_A_FV', 0.0)
                elif nome_tarifario_original == "LUZiGÁS - Energy 8.8": # Usa OMIE Real Simples
                    calc_base_luzigas = omie_para_luzigas_kwh + constantes_dict_local.get('Luzigas_8_8_K', 0.0) + constantes_dict_local.get('Luzigas_CGS', 0.0)
                    temp_preco_calculado = calc_base_luzigas * perdas_a_usar_val_media
                elif nome_tarifario_original == "LUZiGÁS - Dinâmico Poupança +": # Usa OMIE Real Simples
                    calc_base_luzigas = omie_para_luzigas_kwh + constantes_dict_local.get('Luzigas_D_K', 0.0) + constantes_dict_local.get('Luzigas_CGS', 0.0)
                    temp_preco_calculado = calc_base_luzigas * perdas_a_usar_val_media
                elif nome_tarifario_original == "Ibelectra - Solução Família": # Usa OMIE do input user p/ período destino
                    temp_preco_calculado = (omie_kwh_a_usar_na_formula + constantes_dict_local.get('Ibelectra_CS', 0.0)) * perdas_a_usar_val_media + constantes_dict_local.get('Ibelectra_K', 0.0)
                elif nome_tarifario_original == "G9 - Smart Index": # Usa OMIE do input user p/ período destino
                    temp_preco_calculado = (omie_kwh_a_usar_na_formula * constantes_dict_local.get('G9_FA', 1.02) * perdas_a_usar_val_media) + constantes_dict_local.get('G9_CGS', 0.01) + constantes_dict_local.get('G9_AC', 0.0055)
                elif nome_tarifario_original == "EDP - Eletricidade Indexada Média": # Usa OMIE do input user p/ período destino
                    temp_preco_calculado = omie_kwh_a_usar_na_formula * constantes_dict_local.get('EDP_M_Perdas', 1.0) * constantes_dict_local.get('EDP_M_K1', 1.0) + constantes_dict_local.get('EDP_M_K2', 0.0)
                else:
                    temp_preco_calculado = omie_kwh_a_usar_na_formula # Fallback
                
                # Atribuir ao respetivo preço_idx_X
                if p_key_destino == 'S': preco_idx_s = round(temp_preco_calculado, prec_media)
                elif p_key_destino == 'V': preco_idx_v = round(temp_preco_calculado, prec_media)
                elif p_key_destino == 'F': preco_idx_f = round(temp_preco_calculado, prec_media)
                elif p_key_destino == 'C': preco_idx_c = round(temp_preco_calculado, prec_media)
                elif p_key_destino == 'P': preco_idx_p = round(temp_preco_calculado, prec_media)
        
        # --- FIM DO CÁLCULO BASE DO PREÇO DE ENERGIA INDEXADA ---

        # Apenas para os períodos relevantes para `opcao_horaria_para_calculo` e `consumos_repartidos_dict`
        if oh_calc_lower == "simples":
            if 'S' in consumos_repartidos_dict: precos_energia_base_kwh_nesta_oh['S'] = preco_idx_s if preco_idx_s is not None else 0.0
        elif oh_calc_lower.startswith("bi-horário"):
            if 'V' in consumos_repartidos_dict: precos_energia_base_kwh_nesta_oh['V'] = preco_idx_v if preco_idx_v is not None else 0.0
            if 'F' in consumos_repartidos_dict: precos_energia_base_kwh_nesta_oh['F'] = preco_idx_f if preco_idx_f is not None else 0.0
        elif oh_calc_lower.startswith("tri-horário"):
            if 'V' in consumos_repartidos_dict: precos_energia_base_kwh_nesta_oh['V'] = preco_idx_v if preco_idx_v is not None else 0.0
            if 'C' in consumos_repartidos_dict: precos_energia_base_kwh_nesta_oh['C'] = preco_idx_c if preco_idx_c is not None else 0.0
            if 'P' in consumos_repartidos_dict: precos_energia_base_kwh_nesta_oh['P'] = preco_idx_p if preco_idx_p is not None else 0.0

        preco_potencia_input_idx = float(dados_tarifario_indexado_linha.get('preco_potencia_dia', 0.0))
        tar_incluida_energia_idx = dados_tarifario_indexado_linha.get('tar_incluida_energia', False)
        tar_incluida_potencia_idx = dados_tarifario_indexado_linha.get('tar_incluida_potencia', True)
        financiamento_tse_incluido_idx = dados_tarifario_indexado_linha.get('financiamento_tse_incluido', False)

        # --- Passo 1 Adaptado: Componentes Base ---
        tar_energia_regulada_idx = {}
        for periodo_consumo_key in consumos_repartidos_dict.keys():
            tar_energia_regulada_idx[periodo_consumo_key] = obter_tar_energia_periodo(
                opcao_horaria_para_calculo, periodo_consumo_key, potencia_contratada_kva, CONSTANTES_df
            )
        tar_potencia_regulada_idx = obter_tar_dia(potencia_contratada_kva, CONSTANTES_df)

        # Para indexados, o preco_comercializador_energia é o próprio preço indexado calculado acima
        # A TAR de energia é adicionada separadamente.
        preco_comercializador_energia_idx_dict = {} # Renomeado para evitar confusão
        for periodo_calc_idx, preco_base_idx in precos_energia_base_kwh_nesta_oh.items():
            if tar_incluida_energia_idx: # Geralmente False para indexados puros
                preco_comercializador_energia_idx_dict[periodo_calc_idx] = preco_base_idx - tar_energia_regulada_idx.get(periodo_calc_idx, 0.0)
            else:
                preco_comercializador_energia_idx_dict[periodo_calc_idx] = preco_base_idx
        # Não limitar a max(0,...) aqui para componentes de indexados, pois podem ser negativos antes da TAR

        # --- Passo 2 Adaptado: TARs Finais ---
        if tar_incluida_potencia_idx:
            preco_comercializador_potencia_idx = preco_potencia_input_idx - tar_potencia_regulada_idx
        else:
            preco_comercializador_potencia_idx = preco_potencia_input_idx

        financiamento_tse_a_adicionar_idx = FINANCIAMENTO_TSE_VAL if not financiamento_tse_incluido_idx else 0.0

        tar_energia_final_idx = {}
        tar_potencia_final_dia_idx = tar_potencia_regulada_idx
        # ... (lógica de Tarifa Social para TARs permanece a mesma)
        desconto_ts_energia_aplicado_val = 0.0 
        desconto_ts_potencia_aplicado_val = 0.0

        if tarifa_social_ativa:
            desconto_ts_energia_bruto = obter_constante('Desconto TS Energia', CONSTANTES_df)
            desconto_ts_potencia_dia_bruto = obter_constante(f'Desconto TS Potencia {potencia_contratada_kva}', CONSTANTES_df)
            for periodo_calc, tar_reg_val in tar_energia_regulada_idx.items():
                tar_energia_final_idx[periodo_calc] = tar_reg_val - desconto_ts_energia_bruto 
            desconto_ts_energia_aplicado_val = desconto_ts_energia_bruto

            tar_potencia_final_dia_idx = max(0.0, tar_potencia_regulada_idx - desconto_ts_potencia_dia_bruto)
            desconto_ts_potencia_aplicado_val = min(tar_potencia_regulada_idx, desconto_ts_potencia_dia_bruto)
        else:
            tar_energia_final_idx = tar_energia_regulada_idx.copy()

        # --- Passo 3 Adaptado: Preços Finais Energia s/IVA ---
        preco_energia_final_sem_iva_idx_dict = {}
        for periodo_calc_idx_final in consumos_repartidos_dict.keys(): # Iterar sobre os períodos COM CONSUMO
            if periodo_calc_idx_final in preco_comercializador_energia_idx_dict:
                preco_energia_final_sem_iva_idx_dict[periodo_calc_idx_final] = (
                    preco_comercializador_energia_idx_dict.get(periodo_calc_idx_final, 0.0) +
                    tar_energia_final_idx.get(periodo_calc_idx_final, 0.0) +
                    financiamento_tse_a_adicionar_idx
                )
            # Se um período de consumo não tem preço em preco_comercializador_energia_idx_dict (ex: tarifário só tem S, mas calculamos para V)
            # o preço será apenas TAR + TSE. Isto deve ser tratado pela lógica de montagem de preco_energia_base_kwh_nesta_oh.

        # --- Passo 4 Adaptado: Componentes Finais Potência s/IVA ---
        preco_comercializador_potencia_final_sem_iva_idx = preco_comercializador_potencia_idx
        # tar_potencia_final_dia_sem_iva_idx é tar_potencia_final_dia_idx

        # --- Passo 5, 6, 7 (Cálculos de Custo com IVA e Taxas) - permanecem muito semelhantes ---
        consumo_total_neste_oh_idx = sum(float(v or 0) for v in consumos_repartidos_dict.values())

        decomposicao_custo_energia_idx_calc = calcular_custo_energia_com_iva(
            consumo_total_neste_oh_idx,
            preco_energia_final_sem_iva_idx_dict.get('S') if oh_calc_lower == "simples" else None,
            {p: v for p, v in preco_energia_final_sem_iva_idx_dict.items() if p != 'S'},
            dias_calculo, potencia_contratada_kva, opcao_horaria_para_calculo,
            consumos_repartidos_dict,
            familia_numerosa_ativa
        )

        decomposicao_custo_potencia_idx_calc = calcular_custo_potencia_com_iva_final(
            preco_comercializador_potencia_final_sem_iva_idx,
            tar_potencia_final_dia_idx,
            dias_calculo,
            potencia_contratada_kva
        )

        decomposicao_taxas_idx_calc = calcular_taxas_adicionais(
            consumo_total_neste_oh_idx, dias_calculo, tarifa_social_ativa,
            valor_dgeg_user_input, valor_cav_user_input,
            nome_comercializador_atual=str(dados_tarifario_indexado_linha.get('comercializador')),
            mes_selecionado_simulacao=mes_selecionado_pelo_user_str,
            ano_simulacao_atual=ano_atual_calculo
        )

        # --- Passo 8: Custo Total e Descontos de Fatura (se aplicável a indexados) ---
        custo_total_antes_desc_fatura_idx_calc = (
            decomposicao_custo_energia_idx_calc['custo_com_iva'] +
            decomposicao_custo_potencia_idx_calc['custo_com_iva'] +
            decomposicao_taxas_idx_calc['custo_com_iva']
        )

        desconto_fatura_mensal_idx_excel = float(dados_tarifario_indexado_linha.get('desconto_fatura_mes', 0.0) or 0.0)
        desconto_fatura_periodo_aplicado_idx = 0.0
        if desconto_fatura_mensal_idx_excel > 0:
            e_mes_completo_selecionado_idx = (dias_calculo == dias_no_mes_selecionado_dict.get(mes_selecionado_pelo_user_str, 0))
            nome_a_exibir_final += f" (+desc. fat. {desconto_fatura_mensal_idx_excel:.2f}€/mês)"
            desconto_fatura_periodo_aplicado_idx = (desconto_fatura_mensal_idx_excel / 30.0) * dias_calculo if not e_mes_completo_selecionado_idx else desconto_fatura_mensal_idx_excel

        custo_total_final_calculado_idx = custo_total_antes_desc_fatura_idx_calc - desconto_fatura_periodo_aplicado_idx

        # Preço unitário da potência s/IVA (comercializador + TAR final)
        preco_unit_potencia_siva_idx = preco_comercializador_potencia_final_sem_iva_idx + tar_potencia_final_dia_idx # Esta é a soma correta

        # --- Montar Tooltips ---
        componentes_tooltip_custo_total_idx = {
            'tt_cte_energia_siva': decomposicao_custo_energia_idx_calc['custo_sem_iva'],
            'tt_cte_potencia_siva': decomposicao_custo_potencia_idx_calc['custo_sem_iva'],
            'tt_cte_iec_siva': decomposicao_taxas_idx_calc['iec_sem_iva'],
            'tt_cte_dgeg_siva': decomposicao_taxas_idx_calc['dgeg_sem_iva'],
            'tt_cte_cav_siva': decomposicao_taxas_idx_calc['cav_sem_iva'],
            'tt_cte_total_siva': decomposicao_custo_energia_idx_calc['custo_sem_iva'] + decomposicao_custo_potencia_idx_calc['custo_sem_iva'] + decomposicao_taxas_idx_calc['custo_sem_iva'],
            'tt_cte_valor_iva_6_total': decomposicao_custo_energia_idx_calc['valor_iva_6'] + decomposicao_custo_potencia_idx_calc['valor_iva_6'] + decomposicao_taxas_idx_calc['valor_iva_6'],
            'tt_cte_valor_iva_23_total': decomposicao_custo_energia_idx_calc['valor_iva_23'] + decomposicao_custo_potencia_idx_calc['valor_iva_23'] + decomposicao_taxas_idx_calc['valor_iva_23'],
            'tt_cte_subtotal_civa': custo_total_antes_desc_fatura_idx_calc,
            'tt_cte_desc_finais_valor': desconto_fatura_periodo_aplicado_idx,
            'tt_cte_acres_finais_valor': 0.0,
            # NOVOS CAMPOS PARA PREÇOS UNITÁRIOS NO TOOLTIP:
            **{f"tt_preco_unit_energia_{p}_siva": v for p, v in preco_energia_final_sem_iva_idx_dict.items()},
            'tt_preco_unit_potencia_siva': preco_unit_potencia_siva_idx
        }

        # Tooltips de Energia e Potência para Indexados (Adaptação da lógica dos Fixos)
        componentes_tooltip_energia_dict_idx_final = {}
        for p_key_tt_idx in preco_comercializador_energia_idx_dict.keys(): # Usar os períodos que têm preço de comercializador
            componentes_tooltip_energia_dict_idx_final[f'tooltip_energia_{p_key_tt_idx}_comerc_sem_tar'] = preco_comercializador_energia_idx_dict.get(p_key_tt_idx, 0.0)
            componentes_tooltip_energia_dict_idx_final[f'tooltip_energia_{p_key_tt_idx}_tar_bruta'] = tar_energia_regulada_idx.get(p_key_tt_idx, 0.0) # TAR Bruta
            componentes_tooltip_energia_dict_idx_final[f'tooltip_energia_{p_key_tt_idx}_tse_declarado_incluido'] = financiamento_tse_incluido_idx
            componentes_tooltip_energia_dict_idx_final[f'tooltip_energia_{p_key_tt_idx}_tse_valor_nominal'] = FINANCIAMENTO_TSE_VAL
            componentes_tooltip_energia_dict_idx_final[f'tooltip_energia_{p_key_tt_idx}_ts_aplicada_flag'] = tarifa_social_ativa
            componentes_tooltip_energia_dict_idx_final[f'tooltip_energia_{p_key_tt_idx}_ts_desconto_valor'] = obter_constante('Desconto TS Energia', CONSTANTES_df) if tarifa_social_ativa else 0.0

        componentes_tooltip_potencia_dict_idx_final = {
            'tooltip_pot_comerc_sem_tar': preco_comercializador_potencia_idx, # Componente comercializador (s/TAR, s/TS)
            'tooltip_pot_tar_bruta': tar_potencia_regulada_idx, # TAR Bruta
            'tooltip_pot_ts_aplicada': tarifa_social_ativa,
            'tooltip_pot_desconto_ts_valor': desconto_ts_potencia_aplicado_val # Desconto TS efetivo na TAR
        }

        return {
            'Total (€)': custo_total_final_calculado_idx,
            'NomeParaExibirAjustado': nome_a_exibir_final,
            'componentes_tooltip_custo_total_dict': componentes_tooltip_custo_total_idx,
            'componentes_tooltip_energia_dict': componentes_tooltip_energia_dict_idx_final,
            'componentes_tooltip_potencia_dict': componentes_tooltip_potencia_dict_idx_final
        }

    except KeyError as ke:
        # st.error(f"DEBUG COMP ERRO KEY: Erro de chave '{ke}' ao calcular custo para indexado {nome_tarifario_original} na opção {opcao_horaria_para_calculo}.")
        return None
    except Exception as e:
        # st.error(f"DEBUG COMP ERRO GERAL: Erro ao calcular custo para tarifário indexado {nome_tarifario_original} na opção {opcao_horaria_para_calculo}: {e}")
        import traceback
        # st.error(traceback.format_exc()) # Para depuração mais detalhada
        return None


def preparar_consumos_para_cada_opcao_destino(
    opcao_horaria_principal_str,
    consumos_input_atuais_dict,
    opcoes_destino_db_nomes_list # Lista das opções de destino válidas (resultado da func anterior)
):
    consumos_para_calculo_por_oh_destino = {}
    oh_principal_lower = opcao_horaria_principal_str.lower()

    # Consumos introduzidos pelo utilizador para a sua opção principal
    c_s_in = float(consumos_input_atuais_dict.get('S', 0))
    c_v_in = float(consumos_input_atuais_dict.get('V', 0))
    c_f_in = float(consumos_input_atuais_dict.get('F', 0)) # Para Bi
    c_c_in = float(consumos_input_atuais_dict.get('C', 0)) # Para Tri
    c_p_in = float(consumos_input_atuais_dict.get('P', 0)) # Para Tri

    for oh_destino_str in opcoes_destino_db_nomes_list:
        oh_destino_lower = oh_destino_str.lower()
        consumos_finais_para_este_destino = {}

        if oh_destino_lower == "simples":
            if oh_principal_lower == "simples":
                consumos_finais_para_este_destino['S'] = c_s_in
            elif oh_principal_lower.startswith("bi-horário"):
                consumos_finais_para_este_destino['S'] = c_v_in + c_f_in
            elif oh_principal_lower.startswith("tri-horário"):
                # Verifica se é Tri normal ou de alta potência, mas a agregação é a mesma
                consumos_finais_para_este_destino['S'] = c_v_in + c_c_in + c_p_in

        elif oh_destino_lower.startswith("bi-horário"):
            # Esta opção de destino só é válida se a principal for Bi ou Tri
            if oh_principal_lower.startswith("bi-horário"):
                consumos_finais_para_este_destino['V'] = c_v_in
                consumos_finais_para_este_destino['F'] = c_f_in
            elif oh_principal_lower.startswith("tri-horário"):
                consumos_finais_para_este_destino['V'] = c_v_in
                consumos_finais_para_este_destino['F'] = c_c_in + c_p_in

        elif oh_destino_lower.startswith("tri-horário"): # Inclui variantes >20.7kVA
            # Esta opção de destino só é válida se a principal for Tri
            if oh_principal_lower.startswith("tri-horário"):
                consumos_finais_para_este_destino['V'] = c_v_in
                consumos_finais_para_este_destino['C'] = c_c_in
                consumos_finais_para_este_destino['P'] = c_p_in
        
        # Apenas adiciona ao dicionário se consumos_finais_para_este_destino não estiver vazio
        # e se a soma dos consumos for maior que zero (para evitar calcular com tudo a zero)
        if consumos_finais_para_este_destino and sum(v for v in consumos_finais_para_este_destino.values() if v is not None) > 0:
            consumos_para_calculo_por_oh_destino[oh_destino_str] = consumos_finais_para_este_destino
            
    return consumos_para_calculo_por_oh_destino
#FIM DETERMINAÇÃO DE OPÇÕES HORÁRIAS

ts_global_ativa = tarifa_social # Flag global de TS

# --- "Meu Tarifário" ---
# Exibe o subheader e o conteúdo apenas se a checkbox estiver selecionada
if meu_tarifario_ativo:
    st.subheader("🧾 O Meu Tarifário (para comparação)")

    # Definir chaves para todos os inputs do Meu Tarifário
    # Preços de Energia e Potência
    key_energia_meu_s = "energia_meu_s_input_val"
    key_potencia_meu = "potencia_meu_input_val"
    key_energia_meu_v = "energia_meu_v_input_val"
    key_energia_meu_f = "energia_meu_f_input_val"
    key_energia_meu_c = "energia_meu_c_input_val"
    key_energia_meu_p = "energia_meu_p_input_val"
    # Checkboxes TAR/TSE
    key_meu_tar_energia = "meu_tar_energia_val"
    key_meu_tar_potencia = "meu_tar_potencia_val"
    key_meu_fin_tse_incluido = "meu_fin_tse_incluido_val"
    # Descontos
    key_meu_desconto_energia = "meu_desconto_energia_val"
    key_meu_desconto_potencia = "meu_desconto_potencia_val"
    key_meu_desconto_fatura = "meu_desconto_fatura_val"
    # Acréscimo
    key_meu_acrescimo_fatura = "meu_acrescimo_fatura_val"

    # Lista de todas as chaves do Meu Tarifário para facilitar a limpeza
    chaves_meu_tarifario = [
        key_energia_meu_s, key_potencia_meu, key_energia_meu_v, key_energia_meu_f,
        key_energia_meu_c, key_energia_meu_p, key_meu_tar_energia, key_meu_tar_potencia,
        key_meu_fin_tse_incluido, key_meu_desconto_energia, key_meu_desconto_potencia,
        key_meu_desconto_fatura, key_meu_acrescimo_fatura
    ]

    # Botão para Limpar Dados do Meu Tarifário
    # Colocado antes dos inputs para que a limpeza ocorra antes da renderização dos inputs
    if st.button("🧹 Limpar Dados do Meu Tarifário", key="btn_limpar_meu_tarifario"):
        for k in chaves_meu_tarifario:
            if k in st.session_state:
                del st.session_state[k]
        # Também podemos limpar o resultado calculado, se existir
        if 'meu_tarifario_calculado' in st.session_state:
            del st.session_state['meu_tarifario_calculado']
        st.success("Dados do 'Meu Tarifário' foram repostos.")

    col_user1, col_user2, col_user3, col_user4 = st.columns(4)

    # Usar st.session_state.get(key, default_value_para_o_input)
    # Default None para number_input faz com que apareça vazio (com placeholder se definido no widget)
    # Default True/False para checkboxes
    # Default 0.0 para descontos

    if opcao_horaria.lower() == "simples":
        with col_user1:
            energia_meu = st.number_input("Preço Energia (€/kWh)", min_value=0.0, step=0.001, format="%g",
                                        value=st.session_state.get(key_energia_meu_s, None), key=key_energia_meu_s)
        with col_user2:
            potencia_meu = st.number_input("Preço Potência (€/dia)", min_value=0.0, step=0.001, format="%g",
                                         value=st.session_state.get(key_potencia_meu, None), key=key_potencia_meu)
    elif opcao_horaria.lower().startswith("bi"):
        with col_user1:
            energia_vazio_meu = st.number_input("Preço Vazio (€/kWh)", min_value=0.0, step=0.001, format="%g",
                                              value=st.session_state.get(key_energia_meu_v, None), key=key_energia_meu_v)
        with col_user2:
            energia_fora_vazio_meu = st.number_input("Preço Fora Vazio (€/kWh)", min_value=0.0, step=0.001, format="%g",
                                                   value=st.session_state.get(key_energia_meu_f, None), key=key_energia_meu_f)
        with col_user3:
            potencia_meu = st.number_input("Preço Potência (€/dia)", min_value=0.0, step=0.001, format="%g",
                                         value=st.session_state.get(key_potencia_meu, None), key=key_potencia_meu)
    elif opcao_horaria.lower().startswith("tri"):
        with col_user1:
            energia_vazio_meu = st.number_input("Preço Vazio (€/kWh)", min_value=0.0, step=0.001, format="%g",
                                              value=st.session_state.get(key_energia_meu_v, None), key=key_energia_meu_v)
        with col_user2:
            energia_cheias_meu = st.number_input("Preço Cheias (€/kWh)", min_value=0.0, step=0.001, format="%g",
                                               value=st.session_state.get(key_energia_meu_c, None), key=key_energia_meu_c)
        with col_user3:
            energia_ponta_meu = st.number_input("Preço Ponta (€/kWh)", min_value=0.0, step=0.001, format="%g",
                                              value=st.session_state.get(key_energia_meu_p, None), key=key_energia_meu_p)
        with col_user4:
            potencia_meu = st.number_input("Preço Potência (€/dia)", min_value=0.0, step=0.001, format="%g",
                                         value=st.session_state.get(key_potencia_meu, None), key=key_potencia_meu)

    col_userx1, col_userx2, col_userx3 = st.columns(3)
    with col_userx1:
        tar_incluida_energia_meu = st.checkbox("TAR incluída na Energia?", value=st.session_state.get(key_meu_tar_energia, True), key=key_meu_tar_energia, help="É muito importante saber se os valores têm ou não as TAR (Tarifas de Acesso às Redes). Alguns comercializadores separam na fatura, outros não. Verifique se há alguma referência a Acesso às Redes na fatura em (€/kWh)")
    with col_userx2:
        tar_incluida_potencia_meu = st.checkbox("TAR incluída na Potência?", value=st.session_state.get(key_meu_tar_potencia, True), key=key_meu_tar_potencia, help="É muito importante saber se os valores têm ou não as TAR (Tarifas de Acesso às Redes). Alguns comercializadores separam na fatura, outros não. Verifique se há alguma referência a Acesso às Redes na fatura em (€/dia)")
    with col_userx3:
        # A checkbox "Inclui Financiamento TSE?" guarda True se ESTIVER incluído.
        # A variável 'adicionar_financiamento_tse_meu' é o inverso.
        checkbox_tse_incluido_estado = st.checkbox("Inclui Financiamento TSE?", value=st.session_state.get(key_meu_fin_tse_incluido, True), key=key_meu_fin_tse_incluido, help="É importante saber se os valores têm ou não incluido o financiamento da Tarifa Social de Eletricidade (TSE). Alguns comercializadores separam na fatura, outros não. Verifique se há alguma referência a Financiamento Tarifa Social na fatura em (€/kWh)")
        adicionar_financiamento_tse_meu = not checkbox_tse_incluido_estado


    col_userd1, col_userd2, col_userd3, col_userd4 = st.columns(4)
    with col_userd1:
        desconto_energia = st.number_input("Desconto Energia (%)", min_value=0.0, max_value=100.0, step=0.1,
                                           value=st.session_state.get(key_meu_desconto_energia, 0.0), key=key_meu_desconto_energia, help="O desconto é aplicado a Energia+TAR. Alguns tarifários não aplicam o desconto nas TAR (por exemplo os da Plenitude), pelo que o desconto não pode ser aqui aplicado. Se não tiver, não necessita preencher!")
    with col_userd2:
        desconto_potencia = st.number_input("Desconto Potência (%)", min_value=0.0, max_value=100.0, step=0.1,
                                            value=st.session_state.get(key_meu_desconto_potencia, 0.0), key=key_meu_desconto_potencia, help="O desconto é aplicado a Potência+TAR. Alguns tarifários não aplicam o desconto nas TAR, pelo que se assim for, o desconto não pode ser aqui aplicado. Se não tiver, não necessita preencher!")
    with col_userd3:
        desconto_fatura_input_meu = st.number_input("Desconto Fatura (€)", min_value=0.0, step=0.01, format="%.2f",
                                                 value=st.session_state.get(key_meu_desconto_fatura, 0.0), key=key_meu_desconto_fatura, help="Se não tiver, não necessita preencher!")
    with col_userd4:
        acrescimo_fatura_input_meu = st.number_input("Acréscimo Fatura (€)", min_value=0.0, step=0.01, format="%.2f",
                                                 value=st.session_state.get(key_meu_acrescimo_fatura, 0.0), key=key_meu_acrescimo_fatura, help="Para outros custos fixos na fatura. Se não tiver, não necessita preencher!")
    


    if st.button("Calcular e Adicionar O Meu Tarifário à Comparação", icon="🧮", type="primary", key="btn_meu_tarifario", use_container_width=True):
        # Ler os valores dos inputs. Como eles têm 'key', já estão no st.session_state.
        # As variáveis locais (energia_meu, potencia_meu, etc.) já contêm os valores dos widgets.

        preco_energia_input_meu = {}
        
        # Acesso aos valores via variáveis locais (que o Streamlit preenche a partir dos widgets com keys)
        if opcao_horaria.lower() == "simples":
            preco_energia_input_meu['S'] = float(energia_meu or 0.0) # Usa a variável local 'energia_meu'
            preco_potencia_input_meu = float(potencia_meu or 0.0)
        elif opcao_horaria.lower().startswith("bi"):
            preco_energia_input_meu['V'] = float(energia_vazio_meu or 0.0)
            preco_energia_input_meu['F'] = float(energia_fora_vazio_meu or 0.0)
            preco_potencia_input_meu = float(potencia_meu or 0.0)
        elif opcao_horaria.lower().startswith("tri"):
            preco_energia_input_meu['V'] = float(energia_vazio_meu or 0.0)
            preco_energia_input_meu['C'] = float(energia_cheias_meu or 0.0)
            preco_energia_input_meu['P'] = float(energia_ponta_meu or 0.0)
            preco_potencia_input_meu = float(potencia_meu or 0.0)
        else: # Fallback caso algo corra mal com opcao_horaria
            preco_potencia_input_meu = 0.0

        preco_potencia_input_meu = float(potencia_meu or 0.0) # Input user potencia
        alert_negativo = False
        if preco_potencia_input_meu < 0: alert_negativo = True

        consumos_horarios_para_func = {} # Dicionário consumos para função IVA

        if opcao_horaria.lower() == "simples":
            preco_energia_input_meu['S'] = float(energia_meu or 0.0)
            if preco_energia_input_meu['S'] < 0: alert_negativo = True
            consumos_horarios_para_func = {'S': consumo_simples}
        elif opcao_horaria.lower().startswith("bi"):
            preco_energia_input_meu['V'] = float(energia_vazio_meu or 0.0)
            preco_energia_input_meu['F'] = float(energia_fora_vazio_meu or 0.0)
            if preco_energia_input_meu['V'] < 0 or preco_energia_input_meu['F'] < 0: alert_negativo = True
            consumos_horarios_para_func = {'V': consumo_vazio, 'F': consumo_fora_vazio}
        elif opcao_horaria.lower().startswith("tri"):
            preco_energia_input_meu['V'] = float(energia_vazio_meu or 0.0)
            preco_energia_input_meu['C'] = float(energia_cheias_meu or 0.0)
            preco_energia_input_meu['P'] = float(energia_ponta_meu or 0.0)
            if preco_energia_input_meu['V'] < 0 or preco_energia_input_meu['C'] < 0 or preco_energia_input_meu['P'] < 0: alert_negativo = True
            consumos_horarios_para_func = {'V': consumo_vazio, 'C': consumo_cheias, 'P': consumo_ponta}

        if alert_negativo:
            st.warning("Atenção: Introduziu um ou mais preços negativos para o seu tarifário.")

    # --- 1. OBTER COMPONENTES BASE (SEM DESCONTOS, SEM TS, SEM IVA) ---

    # ENERGIA (por período p)
        tar_energia_regulada_periodo_meu = {} # TAR da energia por período (€/kWh)
        for p_key in preco_energia_input_meu.keys(): # S, V, F, C, P
            tar_energia_regulada_periodo_meu[p_key] = obter_tar_energia_periodo(opcao_horaria, p_key, potencia, CONSTANTES)

        energia_meu_periodo_comercializador_base = {} # Componente do comercializador para energia (€/kWh)
        for p_key, preco_input_val in preco_energia_input_meu.items():
            preco_input_val_float = float(preco_input_val or 0.0)
            if tar_incluida_energia_meu:
                energia_meu_periodo_comercializador_base[p_key] = preco_input_val_float - tar_energia_regulada_periodo_meu.get(p_key, 0.0)
            else:
                energia_meu_periodo_comercializador_base[p_key] = preco_input_val_float

        # Financiamento TSE (€/kWh, valor único, aplicável a todos os períodos de energia)
        # 'adicionar_financiamento_tse_meu' é (not checkbox_tse_incluido_estado)
        financiamento_tse_a_somar_base = FINANCIAMENTO_TSE_VAL if adicionar_financiamento_tse_meu else 0.0

        # POTÊNCIA (€/dia)
        tar_potencia_regulada_meu_base = obter_tar_dia(potencia, CONSTANTES) # TAR da potência
        preco_potencia_input_meu_float = float(preco_potencia_input_meu or 0.0)
        if tar_incluida_potencia_meu:
            potencia_meu_comercializador_base = preco_potencia_input_meu_float - tar_potencia_regulada_meu_base
        else:
            potencia_meu_comercializador_base = preco_potencia_input_meu_float

    # --- 2. CALCULAR PREÇOS UNITÁRIOS FINAIS (PARA EXIBIR NA TABELA, SEM IVA) ---
    # Estes preços já incluem o desconto percentual do comercializador e o desconto da Tarifa Social.

        preco_energia_final_unitario_sem_iva = {} # Dicionário para {período: preço_final_unitario}
        desconto_monetario_ts_energia = 0.0 # Valor do desconto TS para energia em €/kWh
        if tarifa_social: # Flag global de TS
            desconto_monetario_ts_energia = obter_constante('Desconto TS Energia', CONSTANTES)

        for p_key in energia_meu_periodo_comercializador_base.keys():
            # Base para o desconto percentual da energia (Comercializador + TAR + TSE)
            preco_total_energia_antes_desc_perc = (
                energia_meu_periodo_comercializador_base.get(p_key, 0.0) +
                tar_energia_regulada_periodo_meu.get(p_key, 0.0) +
                financiamento_tse_a_somar_base
            )

            # Aplicar desconto percentual do comercializador à energia
            preco_energia_apos_desc_comerc = preco_total_energia_antes_desc_perc * (1 - (desconto_energia or 0.0) / 100.0)

            # Aplicar desconto da Tarifa Social (se ativo)
            if tarifa_social:
                preco_energia_final_unitario_sem_iva[p_key] = preco_energia_apos_desc_comerc - desconto_monetario_ts_energia
            else:
                preco_energia_final_unitario_sem_iva[p_key] = preco_energia_apos_desc_comerc

        # Preço unitário final da Potência (€/dia, sem IVA)
        desconto_monetario_ts_potencia = 0.0 # Valor do desconto TS para potência em €/dia
        if tarifa_social:
            desconto_monetario_ts_potencia = obter_constante(f'Desconto TS Potencia {potencia}', CONSTANTES)

        # Base para o desconto percentual da potência (Comercializador + TAR)
        preco_total_potencia_antes_desc_perc = potencia_meu_comercializador_base + tar_potencia_regulada_meu_base

        # Aplicar desconto percentual do comercializador à potência
        preco_potencia_apos_desc_comerc = preco_total_potencia_antes_desc_perc * (1 - (desconto_potencia or 0.0) / 100.0)

        # Aplicar desconto da Tarifa Social (se ativo)
        if tarifa_social:
            preco_potencia_final_unitario_sem_iva = max(0.0, preco_potencia_apos_desc_comerc - desconto_monetario_ts_potencia)
        else:
            preco_potencia_final_unitario_sem_iva = preco_potencia_apos_desc_comerc

        desconto_ts_potencia_valor_aplicado_meu = 0.0
        if tarifa_social:
             desconto_ts_potencia_dia_bruto_meu = obter_constante(f'Desconto TS Potencia {potencia}', CONSTANTES)
             # O desconto efetivamente aplicado à TAR para o meu tarifário.
             # tar_potencia_regulada_meu_base é a TAR bruta.
             desconto_ts_potencia_valor_aplicado_meu = min(tar_potencia_regulada_meu_base, desconto_ts_potencia_dia_bruto_meu)

    # --- 3. CALCULAR CUSTOS TOTAIS (COM IVA) E DECOMPOSIÇÃO PARA TOOLTIP ---

        # CUSTO ENERGIA COM IVA
        # Preparar inputs para calcular_custo_energia_com_iva, que usa os preços unitários finais (já com descontos)
        preco_energia_simples_para_iva = None
        precos_energia_horarios_para_iva = {}
        if opcao_horaria.lower() == "simples":
            preco_energia_simples_para_iva = preco_energia_final_unitario_sem_iva.get('S')
        else: # Bi ou Tri
            precos_energia_horarios_para_iva = {
                p: val for p, val in preco_energia_final_unitario_sem_iva.items() if p != 'S'
            }

        decomposicao_custo_energia_meu = calcular_custo_energia_com_iva(
            consumo, # Consumo total global
            preco_energia_simples_para_iva,
            precos_energia_horarios_para_iva,
            dias,
            potencia,
            opcao_horaria,
            consumos_horarios_para_func,
            familia_numerosa
        )

        custo_energia_meu_final_com_iva = decomposicao_custo_energia_meu['custo_com_iva']
        tt_cte_energia_siva_meu = decomposicao_custo_energia_meu['custo_sem_iva']
        tt_cte_energia_iva_6_meu = decomposicao_custo_energia_meu['valor_iva_6']
        tt_cte_energia_iva_23_meu = decomposicao_custo_energia_meu['valor_iva_23']


        # CUSTO POTÊNCIA COM IVA
        # Para aplicar IVA corretamente em potências <= 3.45 kVA, precisamos das componentes "comercializador" e "TAR" APÓS os descontos.

        # Componente do comercializador para potência, após o seu desconto percentual
        comp_comerc_pot_para_iva = potencia_meu_comercializador_base * (1 - (desconto_potencia or 0.0) / 100.0)

        # Componente TAR da potência, após o desconto percentual do comercializador e o desconto TS
        tar_pot_bruta_apos_desc_perc = tar_potencia_regulada_meu_base * (1 - (desconto_potencia or 0.0) / 100.0)

        tar_pot_final_para_iva = 0.0
        if tarifa_social:
            tar_pot_final_para_iva = max(0.0, tar_pot_bruta_apos_desc_perc - desconto_monetario_ts_potencia)
        else:
            tar_pot_final_para_iva = tar_pot_bruta_apos_desc_perc

        decomposicao_custo_potencia_meu = calcular_custo_potencia_com_iva_final(
            comp_comerc_pot_para_iva,
            tar_pot_final_para_iva,
            dias,
            potencia
        )
        custo_potencia_meu_final_com_iva = decomposicao_custo_potencia_meu['custo_com_iva']
        tt_cte_potencia_siva_meu = decomposicao_custo_potencia_meu['custo_sem_iva']
        tt_cte_potencia_iva_6_meu = decomposicao_custo_potencia_meu['valor_iva_6']
        tt_cte_potencia_iva_23 = decomposicao_custo_potencia_meu['valor_iva_23']


        # --- 4. TAXAS ADICIONAIS E CUSTO TOTAL FINAL ---

        # Taxas Adicionais (IEC, DGEG, CAV) - chamada à função não muda

        decomposicao_taxas_meu = calcular_taxas_adicionais(
            consumo, dias, tarifa_social,
            valor_dgeg_user, valor_cav_user,
            nome_comercializador_atual="Pessoal",
            mes_selecionado_simulacao=mes,
            ano_simulacao_atual=ano_atual
        )
        taxas_meu_tarifario_com_iva = decomposicao_taxas_meu['custo_com_iva']
        # tt_cte_taxas_siva = decomposicao_taxas_meu['custo_sem_iva'] # Já teremos as taxas s/IVA individuais
        tt_cte_iec_siva = decomposicao_taxas_meu['iec_sem_iva']
        tt_cte_dgeg_siva = decomposicao_taxas_meu['dgeg_sem_iva']
        tt_cte_cav_siva = decomposicao_taxas_meu['cav_sem_iva']
        tt_cte_taxas_iva_6 = decomposicao_taxas_meu['valor_iva_6']
        tt_cte_taxas_iva_23 = decomposicao_taxas_meu['valor_iva_23']


    # Custo Total antes do desconto de fatura em €
        custo_total_antes_desc_fatura = custo_energia_meu_final_com_iva + custo_potencia_meu_final_com_iva + taxas_meu_tarifario_com_iva

        # Aplicar Desconto Fatura (€) - lógica não muda
        # 'desconto_fatura_input_meu' já é o valor do input numérico
        custo_total_meu_tarifario_com_iva = custo_total_antes_desc_fatura - float(desconto_fatura_input_meu or 0.0) + float(acrescimo_fatura_input_meu or 0.0)

        # Calcular totais para o tooltip
        tt_cte_total_siva_meu = tt_cte_energia_siva_meu + tt_cte_potencia_siva_meu + tt_cte_iec_siva + tt_cte_dgeg_siva + tt_cte_cav_siva
        tt_cte_valor_iva_6_total_meu = tt_cte_energia_iva_6_meu + tt_cte_potencia_iva_6_meu + tt_cte_taxas_iva_6
        tt_cte_valor_iva_23_total_meu = tt_cte_energia_iva_23_meu + tt_cte_potencia_iva_23 + tt_cte_taxas_iva_23

        # NOVO: Calcular Subtotal c/IVA (antes do desconto de fatura)
        tt_cte_subtotal_civa_meu = tt_cte_total_siva_meu + tt_cte_valor_iva_6_total_meu + tt_cte_valor_iva_23_total_meu
        
        # Descontos e Acréscimos Finais
        tt_cte_desc_finais_valor_meu = 0.0
        if 'desconto_fatura_input_meu' in locals() and desconto_fatura_input_meu > 0:
            tt_cte_desc_finais_valor_meu = desconto_fatura_input_meu

        tt_cte_acres_finais_valor_meu = float(acrescimo_fatura_input_meu or 0.0)

        # Adicionar ao resultado_meu_tarifario_dict
        componentes_tooltip_custo_total_dict_meu = {
            'tt_cte_energia_siva': tt_cte_energia_siva_meu,
            'tt_cte_potencia_siva': tt_cte_potencia_siva_meu,
            'tt_cte_iec_siva': tt_cte_iec_siva,
            'tt_cte_dgeg_siva': tt_cte_dgeg_siva,
            'tt_cte_cav_siva': tt_cte_cav_siva,
            'tt_cte_total_siva': tt_cte_total_siva_meu,
            'tt_cte_valor_iva_6_total': tt_cte_valor_iva_6_total_meu,
            'tt_cte_valor_iva_23_total': tt_cte_valor_iva_23_total_meu,
            'tt_cte_subtotal_civa': tt_cte_subtotal_civa_meu,
            'tt_cte_desc_finais_valor': tt_cte_desc_finais_valor_meu,
            'tt_cte_acres_finais_valor': tt_cte_acres_finais_valor_meu
        }

        # --- INÍCIO: CAMPOS PARA TOOLTIPS DE ENERGIA (O MEU TARIFÁRIO) ---
        componentes_tooltip_energia_dict_meu = {}

        # Desconto bruto da Tarifa Social para energia (se TS global estiver ativa)
        desconto_ts_energia_bruto = 0.0
        if tarifa_social: # tarifa_social é a flag global do checkbox
            desconto_ts_energia_bruto = obter_constante('Desconto TS Energia', CONSTANTES)


        for p_key_tooltip in preco_energia_input_meu.keys():
            preco_final_celula_periodo = preco_energia_final_unitario_sem_iva.get(p_key_tooltip, 0.0) # Valor que vai para a célula
    
            # Componentes fixas para o tooltip, conforme as novas regras:
            tar_bruta_para_tooltip = tar_energia_regulada_periodo_meu.get(p_key_tooltip, 0.0) # Regra 1
    
            # Se checkbox_tse_incluido_estado é True, o TSE está "embutido" e o tooltip só faz uma nota.
            # Se False, o tooltip mostra "Financiamento TSE: VALOR"
            tse_valor_para_soma_tooltip = FINANCIAMENTO_TSE_VAL if not checkbox_tse_incluido_estado else 0.0
    
            desconto_ts_bruto_para_tooltip = desconto_ts_energia_bruto if tarifa_social else 0.0 # Regra 2

            # Calcular a componente "Comercializador (s/TAR)" para o tooltip:
            # É o valor residual para que (Comercializador_Tooltip + TAR_Tooltip + TSE_Adicional_Tooltip - DescontoTS_Tooltip) = PrecoFinalCelula
            comerc_final_para_tooltip = (
                preco_final_celula_periodo -
                tar_bruta_para_tooltip -
                tse_valor_para_soma_tooltip + # Subtrai o valor que o tooltip adicionará
                desconto_ts_bruto_para_tooltip # Adiciona de volta o que o tooltip subtrairá
            )

            # Valores para as flags e campos nominais do tooltip JS:
            tooltip_tse_declarado_incluido = checkbox_tse_incluido_estado # Vem da checkbox do utilizador
            tooltip_tse_valor_nominal = FINANCIAMENTO_TSE_VAL # O JS usa isto se a flag acima for false
            tooltip_ts_aplicada_flag = tarifa_social
            tooltip_ts_desconto_valor_bruto = desconto_ts_energia_bruto if tarifa_social else 0.0 # O JS usa isto se a flag acima for true

            componentes_tooltip_energia_dict_meu[f'tooltip_energia_{p_key_tooltip}_comerc_sem_tar'] = comerc_final_para_tooltip
            componentes_tooltip_energia_dict_meu[f'tooltip_energia_{p_key_tooltip}_tar_bruta'] = tar_bruta_para_tooltip
            componentes_tooltip_energia_dict_meu[f'tooltip_energia_{p_key_tooltip}_tse_declarado_incluido'] = tooltip_tse_declarado_incluido
            componentes_tooltip_energia_dict_meu[f'tooltip_energia_{p_key_tooltip}_tse_valor_nominal'] = tooltip_tse_valor_nominal
            componentes_tooltip_energia_dict_meu[f'tooltip_energia_{p_key_tooltip}_ts_aplicada_flag'] = tooltip_ts_aplicada_flag
            componentes_tooltip_energia_dict_meu[f'tooltip_energia_{p_key_tooltip}_ts_desconto_valor'] = tooltip_ts_desconto_valor_bruto
        # --- FIM: CAMPOS PARA TOOLTIPS DE ENERGIA (O MEU TARIFÁRIO) ---

            # Para o tooltip do Preço Potência (O MEU TARIFÁRIO):
            potencia_comerc_base_meu = potencia_meu_comercializador_base
            tar_potencia_bruta_meu = tar_potencia_regulada_meu_base
    
            # Base para o desconto percentual da potência (Comercializador Base + TAR Bruta)
            base_para_desconto_potencia = potencia_comerc_base_meu + tar_potencia_bruta_meu
    
            # Valor do desconto monetário total para potência
            desconto_monetario_total_potencia = base_para_desconto_potencia * ((desconto_potencia or 0.0) / 100.0)

            # Aplicar desconto primeiro à componente do comercializador
            if desconto_monetario_total_potencia <= potencia_comerc_base_meu:
                pot_comerc_final_tooltip = potencia_comerc_base_meu - desconto_monetario_total_potencia
                pot_tar_final_tooltip = tar_potencia_bruta_meu
            else:
                pot_comerc_final_tooltip = 0.0
                desconto_restante_pot_para_tar = desconto_monetario_total_potencia - potencia_comerc_base_meu
                pot_tar_final_tooltip = max(0.0, tar_potencia_bruta_meu - desconto_restante_pot_para_tar)

    # ... (cálculo de 'desconto_ts_potencia_valor_aplicado_meu' permanece o mesmo) ...
            componentes_tooltip_potencia_dict_meu = {
                'tooltip_pot_comerc_sem_tar': pot_comerc_final_tooltip, # Componente do comercializador s/TAR e s/TS
                'tooltip_pot_tar_bruta': tar_potencia_regulada_meu_base,              # TAR bruta s/TS
                'tooltip_pot_ts_aplicada': tarifa_social,                       # True se TS ativa globalmente
                'tooltip_pot_desconto_ts_valor': desconto_ts_potencia_valor_aplicado_meu if tarifa_social else 0.0, # Valor do desconto TS efetivamente aplicado à TAR
            }

        # --- 5. PREPARAR RESULTADOS PARA EXIBIÇÃO NA TABELA ---

        nome_para_exibir_meu_tarifario = "O Meu Tarifário"
        sufixo_nome = "" # String para o sufixo do nome

        # Converter para float, tratando valores nulos para evitar erros
        desconto = float(desconto_fatura_input_meu or 0.0)
        acrescimo = float(acrescimo_fatura_input_meu or 0.0)

        if desconto > 0 and acrescimo > 0:
            # Cenário 1: Ambos existem, calcular o valor líquido
            valor_liquido = desconto - acrescimo
            if valor_liquido > 0:
                sufixo_nome = f" (Inclui desconto líquido de {valor_liquido:.2f}€ no período)"
            elif valor_liquido < 0:
                sufixo_nome = f" (Inclui acréscimo líquido de {abs(valor_liquido):.2f}€ no período)"
            # Se valor_liquido for 0, não adicionamos sufixo pois anulam-se

        elif desconto > 0:
            # Cenário 2: Apenas desconto existe
            sufixo_nome = f" (Inclui desconto de {desconto:.2f}€ no período)"

        elif acrescimo > 0:
            # Cenário 3: Apenas acréscimo existe
            sufixo_nome = f" (Inclui acréscimo de {acrescimo:.2f}€ no período)"

        # Adicionar o sufixo calculado ao nome final
        nome_para_exibir_meu_tarifario += sufixo_nome

        valores_energia_meu_exibir_dict = {}
        if isinstance(preco_energia_final_unitario_sem_iva, dict):
            for p, v_energia in preco_energia_final_unitario_sem_iva.items():
                periodo_nome = ""
                if p == 'S': periodo_nome = "Simples"
                elif p == 'V': periodo_nome = "Vazio"
                elif p == 'F': periodo_nome = "Fora Vazio"
                elif p == 'C': periodo_nome = "Cheias"
                elif p == 'P': periodo_nome = "Ponta"
                if periodo_nome:
                    valores_energia_meu_exibir_dict[f'{periodo_nome} (€/kWh)'] = round(v_energia, 4)

        resultado_meu_tarifario_dict = {
            'NomeParaExibir': nome_para_exibir_meu_tarifario,
            'LinkAdesao': "-",            
            'Tipo': "Pessoal",
            'Comercializador': "-",
            'Segmento': "Pessoal",
            'Faturação': "-",
            'Pagamento': "-",          
            **valores_energia_meu_exibir_dict,
            'Potência (€/dia)': round(preco_potencia_final_unitario_sem_iva, 4), # Arredondar para exibição
            'Total (€)': round(custo_total_meu_tarifario_com_iva, 2),
            'opcao_horaria_calculada': opcao_horaria,
            # CAMPOS DO TOOLTIP DA POTÊNCIA MEU
            **componentes_tooltip_potencia_dict_meu,
            # CAMPOS DO TOOLTIP DA ENERGIA MEU
            **componentes_tooltip_energia_dict_meu, 
            # CAMPOS DO TOOLTIP DA CUSTO TOTAL MEU
            **componentes_tooltip_custo_total_dict_meu, 
            }
        
        st.session_state['meu_tarifario_calculado'] = resultado_meu_tarifario_dict
        st.success(f"Cálculo para 'O Meu Tarifário' adicionado/atualizado. Custo: {custo_total_meu_tarifario_com_iva:.2f} €")
# --- Fim do if st.button ---
        

#Seletor de Modo de Visualização - OPÇÃO HORÁRIA OU DETALHADA
help_modo_de_comparacao = """
    Ative para ver uma tabela que compara o custo de cada tarifário em diferentes opções horárias (Simples, Bi-Horário, Tri-Horário), usando os seus consumos inseridos.

    **Estará ordenada pela opção tarifária escolhida, mas posteriormente pode ordenar por qualquer outra coluna.**
    """
modo_de_comparacao_ativo = st.checkbox(
"🔬 **Comparar custos entre diferentes Opções Horárias para os mesmos consumos?**",
    value=False, # Começa desativado por defeito
    key="chk_modo_comparacao_opcoes",
    help=help_modo_de_comparacao 
)

st.markdown("---") # Separador
st.subheader("🔍 Filtros da Tabela de Resultados")

# Botão para limpar filtros primeiro, para que a ação ocorra antes de ler os valores dos selectbox
# Usaremos colunas para o layout do botão e do título de resultados
col_titulo_resultados, col_btn_limpar = st.columns([3,1]) # Ajuste a proporção conforme necessário


with col_btn_limpar:
    if st.button("🧹 Remover Todos os Filtros", key="btn_remover_filtros_gerais_tabela", use_container_width=True):
        keys_filtros_a_resetar = ["filter_segmentos_multi", "filter_tipos_multi", 
                                  "filter_faturacao_multi", "filter_pagamento_multi"] # Nomes das chaves para multiselect
        for k_f in keys_filtros_a_resetar:
            if k_f in st.session_state:
                st.session_state[k_f] = [] # Reset para lista vazia para multiselect
        st.rerun()

filt_col1, filt_col2, filt_col3, filt_col4 = st.columns(4)

# --- Filtro de Segmento (agora st.selectbox) ---
with filt_col1:
    opcoes_filtro_segmento_user = ["Residencial", "Empresarial", "Ambos"]
    # Encontrar o índice da opção default ("Residencial") para o selectbox
    default_index_segmento = opcoes_filtro_segmento_user.index("Residencial")
    
    selected_segmento_user = st.selectbox(
        "Segmento", 
        opcoes_filtro_segmento_user, 
        index=st.session_state.get("filter_segmento_selectbox_index", default_index_segmento), # Guarda o índice para manter o estado
        key="filter_segmento_selectbox",
        help="Escolha o segmento para a simulação."
    )
    # Guarda o índice selecionado no session_state
    st.session_state.filter_segmento_selectbox_index = opcoes_filtro_segmento_user.index(selected_segmento_user)


with filt_col2:
    tipos_options_ms = get_filter_options_for_multiselect(tarifarios_fixos, tarifarios_indexados, 'tipo')
    
    # Criar o texto de ajuda formatado com Markdown
    help_text_formatado = """
    Deixe em branco para mostrar todos os tipos ou selecione um ou mais para filtrar.

    Os tipos de tarifário são:
    * **Fixo**: O preço da energia (€/kWh) é constante durante o contrato.
    * **Indexado (Média)**: Preço da energia baseado na média do OMIE para os períodos horários.
    * **Indexado (Quarto-Horário)**: Preço da energia baseado nos valores OMIE horários/quarto-horários e no perfil de consumo. Também conhecidos como 'Dinâmicos'.
    """
    
    selected_tipos = st.multiselect("Tipo(s) de Tarifário", tipos_options_ms,
                                 default=st.session_state.get("filter_tipos_multi", []),
                                 key="filter_tipos_multi",
                                 help=help_text_formatado) # <--- Usar a nova variável aqui

# --- Filtro de Faturação (agora st.selectbox) ---
with filt_col3:
    opcoes_faturacao_user = ["Todas", "Fatura eletrónica", "Fatura em papel"]
    selected_faturacao_user = st.selectbox(
        "Faturação",
        opcoes_faturacao_user,
        index=st.session_state.get("filter_faturacao_selectbox_index", 0), # Default "Todas"
        key="filter_faturacao_selectbox",
        help="Escolha o tipo de faturação pretendido."
    )
    st.session_state.filter_faturacao_selectbox_index = opcoes_faturacao_user.index(selected_faturacao_user)


# --- Filtro de Pagamento (agora st.selectbox) ---
with filt_col4:
    opcoes_pagamento_user = ["Todos", "Débito Direto", "Multibanco", "Numerário/Payshop/CTT"]
    selected_pagamento_user = st.selectbox(
        "Pagamento",
        opcoes_pagamento_user,
        index=st.session_state.get("filter_pagamento_selectbox_index", 0), # Default "Todos"
        key="filter_pagamento_selectbox",
        help="Escolha o método de pagamento."
    )
    st.session_state.filter_pagamento_selectbox_index = opcoes_pagamento_user.index(selected_pagamento_user)


# --- Lógica de Aplicação dos Filtros ---

# Começar com os DataFrames completos
tf_processar = tarifarios_fixos.copy()
ti_processar = tarifarios_indexados.copy()

# 1. Lógica para o filtro de Segmento
if selected_segmento_user != "Ambos":
    segmentos_para_filtrar = []
    if selected_segmento_user == "Residencial":
        segmentos_para_filtrar.extend(["Doméstico", "Doméstico e Não Doméstico"])
    elif selected_segmento_user == "Empresarial":
        segmentos_para_filtrar.extend(["Não Doméstico", "Doméstico e Não Doméstico"])
    
    # Aplicar o filtro
    tf_processar = tf_processar[tf_processar['segmento'].astype(str).str.strip().isin(segmentos_para_filtrar)]
    ti_processar = ti_processar[ti_processar['segmento'].astype(str).str.strip().isin(segmentos_para_filtrar)]

# 2. Lógica para o filtro de Tipo
if selected_tipos:
    tf_processar = tf_processar[tf_processar['tipo'].astype(str).str.strip().isin(selected_tipos)]
    ti_processar = ti_processar[ti_processar['tipo'].astype(str).str.strip().isin(selected_tipos)]

# 3. Lógica para o filtro de Faturação
if selected_faturacao_user != "Todas":
    faturacao_para_filtrar = []
    if selected_faturacao_user == "Fatura eletrónica":
        faturacao_para_filtrar.extend(["Fatura eletrónica", "Fatura eletrónica, Fatura em papel"])
    elif selected_faturacao_user == "Fatura em papel":
        faturacao_para_filtrar.extend(["Fatura eletrónica, Fatura em papel"]) 
    
    # Aplicar o filtro
    tf_processar = tf_processar[tf_processar['faturacao'].astype(str).str.strip().isin(faturacao_para_filtrar)]
    ti_processar = ti_processar[ti_processar['faturacao'].astype(str).str.strip().isin(faturacao_para_filtrar)]

# 4. Lógica para o filtro de Pagamento
if selected_pagamento_user != "Todos":
    pagamento_para_filtrar = []
    if selected_pagamento_user == "Débito Direto":
        pagamento_para_filtrar.extend(["Débito Direto", "Débito Direto, Multibanco", "Débito Direto, Multibanco, Numerário/Payshop/CTT"])
    elif selected_pagamento_user == "Multibanco":
        pagamento_para_filtrar.extend(["Débito Direto, Multibanco", "Débito Direto, Multibanco, Numerário/Payshop/CTT"])
    elif selected_pagamento_user == "Numerário/Payshop/CTT":
        pagamento_para_filtrar.extend(["Débito Direto, Multibanco, Numerário/Payshop/CTT"])
    
    # Aplicar o filtro
    tf_processar = tf_processar[tf_processar['pagamento'].astype(str).str.strip().isin(pagamento_para_filtrar)]
    ti_processar = ti_processar[ti_processar['pagamento'].astype(str).str.strip().isin(pagamento_para_filtrar)]


st.markdown("---") # Separador antes da tabela ou do resumo
#FIM Seletor de Modo de Visualização - NORMAL OU OPÇÃO HORÁRIA

# --- Construir Resumo dos Inputs para Exibição ---
# --- Construir Resumo dos Inputs para Exibição ---
cor_texto_resumo = "#333333"  # Um cinza escuro, bom para fundos claros

resumo_html_parts = [
    f"<div style='background-color: #f9f9f9; border: 1px solid #ddd; padding: 15px; border-radius: 6px; margin-bottom: 25px; color: {cor_texto_resumo};'>"
]
resumo_html_parts.append(f"<h5 style='margin-top:0; color: {cor_texto_resumo};'>Resumo da Simulação:</h5>")
resumo_html_parts.append("<ul style='list-style-type: none; padding-left: 0;'>")

# --- NOVO: Adicionar detalhes dos filtros numa única linha ---
# 1. Tratar o caso especial do segmento "Ambos"
if selected_segmento_user == "Ambos":
    segmento_para_resumo = "Residencial e Empresarial"
else:
    segmento_para_resumo = selected_segmento_user

# 2. Construir a string combinada
# &nbsp; é o código HTML para um espaço, para dar um espaçamento agradável
linha_filtros = (
    f"<b>Segmento:</b> {segmento_para_resumo} &nbsp;&nbsp;|&nbsp;&nbsp; "
    f"<b>Faturação:</b> {selected_faturacao_user} &nbsp;&nbsp;|&nbsp;&nbsp; "
    f"<b>Pagamento:</b> {selected_pagamento_user}"
)

# 3. Adicionar a linha única ao resumo
resumo_html_parts.append(f"<li style='margin-bottom: 5px;'>{linha_filtros}</li>")
# --- FIM DO NOVO BLOCO ---

# 1. Potência contratada + Opção Horária e Ciclo
resumo_html_parts.append(f"<li style='margin-bottom: 5px;'><b>{potencia} kVA</b> em <b>{opcao_horaria}</b></li>")

# 2. Consumo dividido por opção
consumo_detalhe_str = ""
if opcao_horaria.lower() == "simples":
    consumo_detalhe_str = f"Simples: {consumo_simples:.0f} kWh"
elif opcao_horaria.lower().startswith("bi"):
    consumo_detalhe_str = f"Vazio: {consumo_vazio:.0f} kWh, Fora Vazio: {consumo_fora_vazio:.0f} kWh"
elif opcao_horaria.lower().startswith("tri"):
    consumo_detalhe_str = f"Vazio: {consumo_vazio:.0f} kWh, Cheias: {consumo_cheias:.0f} kWh, Ponta: {consumo_ponta:.0f} kWh"
resumo_html_parts.append(f"<li style='margin-bottom: 5px;'><b>Consumos ({consumo:.0f} kWh Total):</b> {consumo_detalhe_str}</li>")

# 3. Datas e Dias Faturados
# 'dias_default_calculado' e 'dias' já foram calculados
# 'dias_manual_input_val' é o valor do widget st.number_input para dias manuais
dias_manual_valor_do_input = st.session_state.get('dias_manual_input_key', dias_default_calculado) # Use a key correta do seu input
usou_dias_manuais_efetivamente = False
if pd.notna(dias_manual_valor_do_input) and dias_manual_valor_do_input > 0 and \
   int(dias_manual_valor_do_input) != dias_default_calculado:
    usou_dias_manuais_efetivamente = True

if usou_dias_manuais_efetivamente:
    resumo_html_parts.append(f"<li style='margin-bottom: 5px;'><b>Período:</b> {dias} dias (definido manualmente)</li>")
else:
    resumo_html_parts.append(f"<li style='margin-bottom: 5px;'><b>Período:</b> De {data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')} ({dias} dias)</li>")

# 4. Valores OMIE da opção escolhida, com a referência
omie_valores_str_parts = []
if opcao_horaria.lower() == "simples":
    val_s = st.session_state.get('omie_s_input_field', round(omie_medios_calculados.get('S',0), 2))
    omie_valores_str_parts.append(f"Simples: {val_s:.2f} €/MWh")
elif opcao_horaria.lower().startswith("bi"):
    val_v = st.session_state.get('omie_v_input_field', round(omie_medios_calculados.get('V',0), 2))
    val_f = st.session_state.get('omie_f_input_field', round(omie_medios_calculados.get('F',0), 2))
    omie_valores_str_parts.append(f"Vazio: {val_v:.2f} €/MWh")
    omie_valores_str_parts.append(f"Fora Vazio: {val_f:.2f} €/MWh")
elif opcao_horaria.lower().startswith("tri"):
    val_v = st.session_state.get('omie_v_input_field', round(omie_medios_calculados.get('V',0), 2))
    val_c = st.session_state.get('omie_c_input_field', round(omie_medios_calculados.get('C',0), 2))
    val_p = st.session_state.get('omie_p_input_field', round(omie_medios_calculados.get('P',0), 2))
    omie_valores_str_parts.append(f"Vazio: {val_v:.2f} €/MWh")
    omie_valores_str_parts.append(f"Cheias: {val_c:.2f} €/MWh")
    omie_valores_str_parts.append(f"Ponta: {val_p:.2f} €/MWh")

if omie_valores_str_parts: # Só mostra a secção OMIE se houver valores a exibir
    resumo_html_parts.append(f"<li style='margin-bottom: 5px;'><b>OMIE {nota_omie}:</b> {', '.join(omie_valores_str_parts)}</li>")

# 5. Perfil de consumo utilizado
perfil_consumo_calculado_str = obter_perfil(consumo, dias, potencia) # Chamar a sua função
# Formatar para uma apresentação mais amigável
texto_perfil_apresentacao = perfil_consumo_calculado_str.replace("perfil_", "Perfil ").upper() # Ex: "Perfil A"
resumo_html_parts.append(f"<li style='margin-bottom: 5px;'><b>Perfil de Consumo:</b> {texto_perfil_apresentacao}</li>")

# 6. Tarifa Social (se ativa)
if tarifa_social:
    resumo_html_parts.append(f"<li style='margin-bottom: 5px; color: red;'><b>Benefício Aplicado:</b> Tarifa Social</li>")

# 7. Família Numerosa (se ativa)
if familia_numerosa:
    resumo_html_parts.append(f"<li style='margin-bottom: 5px; color: red;'><b>Benefício Aplicado:</b> Família Numerosa</li>")

resumo_html_parts.append("</ul>")
resumo_html_parts.append("</div>")
html_resumo_final = "".join(resumo_html_parts)

# Exibir o resumo
st.markdown(html_resumo_final, unsafe_allow_html=True)


resultados_list.clear() # Limpar a lista de resultados no início de cada execução

if modo_de_comparacao_ativo:
    # --- INÍCIO DO NOVO BLOCO PARA O MODO DE COMPARAÇÃO DE OPÇÕES HORÁRIAS ---
    st.subheader("📊 Tiago Felícia - Comparação de Custos entre Opções Tarifárias")
    st.markdown("➡️ [**Exportar Tabela de Comparação para Excel**](#exportar-excel-comparacao)")

#Vários JS, alguns repetidos
    #CORES PARA TARIFÁRIOS INDEXADOS:
    cor_fundo_indexado_media_css = "#FFE699"
    cor_texto_indexado_media_css = "black"
    cor_fundo_indexado_dinamico_css = "#4D79BC"  
    cor_texto_indexado_dinamico_css = "white"

    cell_style_nome_tarifario_js = JsCode(f"""
    function(params) {{
        // Estilo base aplicado a todas as células desta coluna
        let styleToApply = {{ 
            textAlign: 'center',
            borderRadius: '6px',  // O teu borderRadius desejado
            padding: '10px 10px'   // O teu padding desejado
            // Podes adicionar um backgroundColor default para células não especiais aqui, se quiseres
            // backgroundColor: '#f0f0f0' // Exemplo para tarifários fixos
        }};                                  

        if (params.data) {{
            const nomeExibir = params.data.NomeParaExibir;
            const tipoTarifario = params.data.Tipo;

            // VERIFICA SE O NOME COMEÇA COM "O Meu Tarifário"
            if (typeof nomeExibir === 'string' && nomeExibir.startsWith('O Meu Tarifário')) {{
                styleToApply.backgroundColor = 'red';
                styleToApply.color = 'white';
                styleToApply.fontWeight = 'bold';
            }} else if (tipoTarifario === 'Indexado Média') {{
                styleToApply.backgroundColor = '{cor_fundo_indexado_media_css}';
                styleToApply.color = '{cor_texto_indexado_media_css}';
            }} else if (tipoTarifario === 'Indexado quarto-horário') {{
                styleToApply.backgroundColor = '{cor_fundo_indexado_dinamico_css}';
                styleToApply.color = '{cor_texto_indexado_dinamico_css}';
            }} else {{
                // Para tarifários fixos ou outros tipos não explicitamente coloridos acima.
                // Eles já terão o textAlign, borderRadius e padding do styleToApply.
                // Se quiseres um fundo específico para eles diferente do default do styleToApply, define aqui.
                // Ex: styleToApply.backgroundColor = '#e9ecef'; // Uma cor neutra para fixos
            }}
            return styleToApply;
        }}
        return styleToApply; 
    }}
    """)

    #tooltip Nome Tarifario
    tooltip_nome_tarifario_getter_js = JsCode("""
    function(params) {
        if (!params.data) { 
            return params.value || ''; 
        }

        const nomeExibir = params.data.NomeParaExibir || '';
        const notas = params.data.info_notas || ''; 

        let tooltipHtmlParts = [];

        if (nomeExibir) {
            tooltipHtmlParts.push("<strong>" + nomeExibir + "</strong>");
        }

        if (notas) {
            const notasHtml = notas.replace(/\\n/g, ' ').replace(/\n/g, ' ');
            // Usando aspas simples para atributos de estilo HTML para simplificar
            tooltipHtmlParts.push("<small style='display: block; margin-top: 5px;'><i>" + notasHtml + "</i></small>");
        }
    
        if (tooltipHtmlParts.length > 0) {
            // Se ambos nomeExibir e notas existem, queremos uma quebra de linha entre eles no tooltip.
            // Se join(''), eles ficam lado a lado. Se join('<br>'), ficam em linhas separadas.
            // Dado que a nota tem display:block, o join('') deve funcionar para colocá-los em "blocos" separados.
            return tooltipHtmlParts.join(''); // Para agora, vamos juntar diretamente.
                                                // Se quiser uma quebra de linha explícita entre o nome e as notas,
                                                 // e ambos existirem, pode usar:
                                                 // return tooltipHtmlParts.join('<br style="margin-bottom:5px">');
        }
    
        return ''; 
    }
            """)
        
    # Custom_Tooltip
    custom_tooltip_component_js = JsCode("""
        class CustomTooltip {
            init(params) {
                // params.value é a string que o seu tooltipValueGetter retorna
                this.eGui = document.createElement('div');
                // Para permitir HTML, definimos o innerHTML
                // É importante que a string de params.value seja HTML seguro se vier de inputs do utilizador,
                // mas no seu caso, está a construí-lo programaticamente.
                this.eGui.innerHTML = params.value; 

                // Aplicar algum estilo básico para o tooltip se desejar
                this.eGui.style.backgroundColor = 'white'; // Ou outra cor de fundo
                this.eGui.style.color = 'black';           // Cor do texto
                this.eGui.style.border = '1px solid #ccc'; // Borda mais suave
                this.eGui.style.padding = '10px';           // Mais padding
                this.eGui.style.borderRadius = '6px';      // Cantos arredondados
                this.eGui.style.boxShadow = '0 2px 5px rgba(0,0,0,0.15)'; // Sombra suave
                this.eGui.style.maxWidth = '400px';        // Largura máxima
                this.eGui.style.fontSize = '1.1em';        // Tamanho da fonte
                this.eGui.style.fontFamily = 'Arial, sans-serif'; // Tipo de fonte                             
                this.eGui.style.whiteSpace = 'normal';     // Para quebra de linha
            }

            getGui() {
                return this.eGui;
            }
        }
    """)


    custom_css = {
        ".ag-header-cell-label": {
            "justify-content": "center !important",
            "text-align": "center !important",
            "font-size": "14px !important",      # <-- aumenta o header
            "font-weight": "bold !important"
        },
        ".ag-center-header": {
            "justify-content": "center !important",
            "text-align": "center !important",
            "font-size": "14px !important"       # <-- reforço para headers
        },
        ".ag-cell": {
            "font-size": "14px !important"       # <-- aumenta valores das células
        },
        ".ag-center-cols-clip": {"justify-content": "center !important", "text-align": "center !important"}
    }

        # --- 1. DEFINIR O JsCode PARA LINK E TOOLTIP ---
    link_tooltip_renderer_js = JsCode("""
    class LinkTooltipRenderer {
        init(params) {
            this.eGui = document.createElement('div');
            let displayText = params.value; // Valor da célula (NomeParaExibir)
            let url = params.data.LinkAdesao; // Acede ao valor da coluna LinkAdesao da mesma linha

            if (url && typeof url === 'string' && url.toLowerCase().startsWith('http')) {
                // HTML para o link clicável
                // O atributo 'title' (tooltip) mostrará "Aderir/Saber mais: [URL]"
                // O texto visível do link será o 'displayText' (NomeParaExibir)
                this.eGui.innerHTML = `<a href="${url}" target="_blank" title="Aderir/Saber mais: ${url}" style="text-decoration: underline; color: inherit;">${displayText}</a>`;
            } else {
                // Se não houver URL válido, apenas mostra o displayText com o próprio displayText como tooltip.
                this.eGui.innerHTML = `<span title="${displayText}">${displayText}</span>`;
            }
        }
        getGui() { return this.eGui; }
    }
    """) # <--- FIM DA DEFINIÇÃO DE link_tooltip_renderer_js


    resultados_comparacao_list = [] # Usar uma lista dedicada para este modo

    # 1. Obter consumos atuais (já tem esta lógica - consumos_input_atuais_dict_comp)
    consumos_input_atuais_dict_comp = {} 
    if opcao_horaria.lower() == "simples":
        consumos_input_atuais_dict_comp['S'] = consumo_simples
    elif opcao_horaria.lower().startswith("bi-horário"):
        consumos_input_atuais_dict_comp['V'] = consumo_vazio
        consumos_input_atuais_dict_comp['F'] = consumo_fora_vazio
    elif opcao_horaria.lower().startswith("tri-horário"):
        consumos_input_atuais_dict_comp['V'] = consumo_vazio
        consumos_input_atuais_dict_comp['C'] = consumo_cheias
        consumos_input_atuais_dict_comp['P'] = consumo_ponta

    # Criação de todos_omie_inputs_utilizador_comp (já tem nas linhas 905-912)
    # Certifique-se que omie_medios_calculados está disponível aqui
    todos_omie_inputs_utilizador_comp_comparacao = {
        'S': st.session_state.get('omie_s_input_field', round(omie_medios_calculados.get('S',0), 2)),
        'V': st.session_state.get('omie_v_input_field', round(omie_medios_calculados.get('V',0), 2)),
        'F': st.session_state.get('omie_f_input_field', round(omie_medios_calculados.get('F',0), 2)),
        'C': st.session_state.get('omie_c_input_field', round(omie_medios_calculados.get('C',0), 2)),
        'P': st.session_state.get('omie_p_input_field', round(omie_medios_calculados.get('P',0), 2))
    }
       
    # 2. Determinar colunas de destino e ordenação (já tem - determinar_opcoes_horarias_destino_e_ordenacao)
    opcoes_destino_db_nomes_comp, colunas_aggrid_custo_comp, coluna_ordenacao_aggrid_comp = \
        determinar_opcoes_horarias_destino_e_ordenacao(
            opcao_horaria, 
            potencia,      
            consumos_input_atuais_dict_comp, 
            opcoes_horarias_existentes 
        )

    if not opcoes_destino_db_nomes_comp:
        st.warning("Não há opções horárias válidas para comparação com os inputs selecionados.")
    else:
        # Preparar consumos repartidos (já tem - preparar_consumos_para_cada_opcao_destino)
        consumos_repartidos_finais_por_oh_comp = preparar_consumos_para_cada_opcao_destino(
            opcao_horaria, 
            consumos_input_atuais_dict_comp,
            opcoes_destino_db_nomes_comp 
        )

        # 3. "Loop Principal Modificado" para o modo de comparação:
        nomes_tarifarios_unicos_para_comparacao = []
        if not tf_processar.empty: # USAR tf_processar
            temp_fixos_unicos = tf_processar[['nome', 'comercializador', 'tipo', 'site_adesao', 'notas']].drop_duplicates(subset=['nome', 'comercializador'])
            for _, row_fixo in temp_fixos_unicos.iterrows():
                nomes_tarifarios_unicos_para_comparacao.append(dict(row_fixo))

        if comparar_indexados and not ti_processar.empty: # USAR ti_processar
            temp_index_unicos = ti_processar[['nome', 'comercializador', 'tipo', 'site_adesao', 'notas', 'formula_calculo']].drop_duplicates(subset=['nome', 'comercializador'])
            for _, row_idx in temp_index_unicos.iterrows():
                if not any(d['nome'] == row_idx['nome'] and d['comercializador'] == row_idx['comercializador'] for d in nomes_tarifarios_unicos_para_comparacao):
                    nomes_tarifarios_unicos_para_comparacao.append(dict(row_idx))

        if meu_tarifario_ativo and 'meu_tarifario_calculado' in st.session_state:
            omt_data = st.session_state['meu_tarifario_calculado']
            # Adicionar informação para identificar "O Meu Tarifário" e a sua opção horária original
            nomes_tarifarios_unicos_para_comparacao.append({
                'nome': omt_data.get('NomeParaExibir', "O Meu Tarifário"),
                'comercializador': omt_data.get('Comercializador', "-"),
                'tipo': omt_data.get('Tipo', "Pessoal"),
                'site_adesao': omt_data.get('LinkAdesao', '-'),
                'notas': omt_data.get('info_notas', ''), # 'info_notas' é o campo esperado pela AgGrid
                'is_meu_tarifario': True, # Flag para identificar "O Meu Tarifário"
                'opcao_horaria_original_meu_tarifario': omt_data.get('opcao_horaria_calculada')
            })
            
        for info_tar_base in nomes_tarifarios_unicos_para_comparacao:
            nome_t_comp = info_tar_base['nome']
            comerc_t_comp = info_tar_base['comercializador']
            tipo_t_comp = str(info_tar_base['tipo'])
            link_t_comp = info_tar_base.get('site_adesao', '-')
            notas_t_comp = info_tar_base.get('notas', '')
            formula_calculo_idx_comp = str(info_tar_base.get('formula_calculo', '')) # Para indexados

            linha_para_aggrid = {
                'NomeParaExibir': nome_t_comp,
                'Comercializador': comerc_t_comp,
                'Tipo': tipo_t_comp,
                'LinkAdesao': link_t_comp, # LinkAdesao é usado pela AgGrid principal, não faz mal estar aqui
                'info_notas': notas_t_comp  # info_notas é usado pela AgGrid principal
            }
            teve_pelo_menos_um_calculo_nesta_linha = False

            for oh_destino_db_nome in opcoes_destino_db_nomes_comp:
                nome_coluna_aggrid_para_este_oh = f"Total {oh_destino_db_nome} (€)"
                linha_para_aggrid[nome_coluna_aggrid_para_este_oh] = None # Inicializa a coluna de custo

                consumos_para_calculo_nesta_oh = consumos_repartidos_finais_por_oh_comp.get(oh_destino_db_nome)
                if not consumos_para_calculo_nesta_oh or sum(v for v in consumos_para_calculo_nesta_oh.values() if v is not None) == 0:
                    continue # Pula se não houver consumos para esta opção de destino
                
                dados_tarifario_especifico_para_calculo = None # DataFrame da linha específica do tarifário
                
                # Lógica para "O Meu Tarifário" na Tabela Comparativa
                if info_tar_base.get('is_meu_tarifario'):
                    if 'meu_tarifario_calculado' in st.session_state:
                        # Recalcular "O Meu Tarifário" para a opção de destino atual (oh_destino_db_nome)
                        # usando os consumos de `consumos_para_calculo_nesta_oh`
                        # Esta parte é complexa porque "O Meu Tarifário" é definido por inputs diretos de preço
                        # e não por uma linha de dados do Excel como os outros.
                        # Precisaremos de uma função adaptada ou lógica aqui para recalcular.
                        # Por simplicidade, vamos mostrar o custo original se for a opção original, senão None.
                        if info_tar_base.get('opcao_horaria_original_meu_tarifario') == oh_destino_db_nome:
                             omt_data_calc = st.session_state['meu_tarifario_calculado']
                             linha_para_aggrid[nome_coluna_aggrid_para_este_oh] = omt_data_calc.get('Total (€)')
                             # linha_para_aggrid[f'Tooltip {nome_coluna_aggrid_para_este_oh}'] = omt_data_calc.get('componentes_tooltip_custo_total_dict_meu')
                             teve_pelo_menos_um_calculo_nesta_linha = True
                        # else: O custo para outras opções de destino para "O Meu Tarifário" não é calculado aqui.
                    continue # Vai para o próximo oh_destino_db_nome

                # Lógica para tarifários do Excel
                if tipo_t_comp == "Fixo":
                    df_match = tarifarios_fixos[
                        (tarifarios_fixos['nome'] == nome_t_comp) &
                        (tarifarios_fixos['comercializador'] == comerc_t_comp) &
                        (tarifarios_fixos['opcao_horaria_e_ciclo'] == oh_destino_db_nome) &
                        (tarifarios_fixos['potencia_kva'] == potencia)
                    ]
                    if not df_match.empty:
                        dados_tarifario_especifico_para_calculo = df_match.iloc[0]
                
                elif tipo_t_comp.startswith("Indexado"):
                    df_match_idx = tarifarios_indexados[
                        (tarifarios_indexados['nome'] == nome_t_comp) &
                        (tarifarios_indexados['comercializador'] == comerc_t_comp) &
                        (tarifarios_indexados['opcao_horaria_e_ciclo'] == oh_destino_db_nome) &
                        (tarifarios_indexados['potencia_kva'] == potencia)
                    ]
                    if not df_match_idx.empty:
                        dados_tarifario_especifico_para_calculo = df_match_idx.iloc[0]

                if dados_tarifario_especifico_para_calculo is not None:
                    resultado_celula = None
                    if tipo_t_comp == "Fixo":
                        resultado_celula = calcular_detalhes_custo_tarifario_fixo(
                            dados_tarifario_especifico_para_calculo, oh_destino_db_nome,
                            consumos_para_calculo_nesta_oh, potencia, dias, tarifa_social, familia_numerosa,
                            valor_dgeg_user, valor_cav_user, incluir_quota_acp, desconto_continente,
                            CONSTANTES, dias_mes, mes, ano_atual, data_inicio, data_fim
                        )
                    elif tipo_t_comp.startswith("Indexado"):
                        resultado_celula = calcular_detalhes_custo_tarifario_indexado(
                            dados_tarifario_especifico_para_calculo, oh_destino_db_nome, opcao_horaria, # Passa opcao_horaria PRINCIPAL
                            consumos_para_calculo_nesta_oh, potencia, dias, tarifa_social, familia_numerosa,
                            valor_dgeg_user, valor_cav_user, CONSTANTES,
                            df_omie_ajustado, # Passar o global
                            perdas_medias,    # Passar o global
                            todos_omie_inputs_utilizador_comp_comparacao, # OMIEs dos inputs PRINCIPAIS
                            omie_medios_calculados_para_todos_ciclos, # OMIEs CALCULADOS para todos os ciclos
                            omie_medio_simples_real_kwh, # OMIE real simples para Luzigas
                            dias_mes, mes, ano_atual, data_inicio, data_fim
                        )
                    
                    if resultado_celula and pd.notna(resultado_celula.get('Total (€)')):
                        linha_para_aggrid[nome_coluna_aggrid_para_este_oh] = resultado_celula['Total (€)']
                        # Guardar tooltips para cada célula da tabela comparativa
                        linha_para_aggrid[f'Tooltip_{nome_coluna_aggrid_para_este_oh}'] = resultado_celula.get('componentes_tooltip_custo_total_dict')
                        teve_pelo_menos_um_calculo_nesta_linha = True
            
            if teve_pelo_menos_um_calculo_nesta_linha:
                 resultados_comparacao_list.append(linha_para_aggrid)

        df_resultados_comparacao_aggrid = pd.DataFrame(resultados_comparacao_list)
        
        if not df_resultados_comparacao_aggrid.empty:
            if 'coluna_ordenacao_aggrid_comp' in locals() and \
               coluna_ordenacao_aggrid_comp and \
               coluna_ordenacao_aggrid_comp in df_resultados_comparacao_aggrid.columns:
                df_resultados_comparacao_aggrid = df_resultados_comparacao_aggrid.sort_values(
                    by=coluna_ordenacao_aggrid_comp,
                    ascending=True,
                    na_position='last'
                ).reset_index(drop=True)
        
            # --- CONFIGURAÇÃO DO AGGRID PARA O MODO DE COMPARAÇÃO ---
            gb_comp = GridOptionsBuilder.from_dataframe(df_resultados_comparacao_aggrid)
            gb_comp.configure_default_column(
                sortable=True, resizable=True, editable=False, wrapText=True, autoHeight=True,
                wrapHeaderText=True, autoHeaderHeight=True
            )
            
            # Coluna NomeParaExibir com Link e Tooltip de Notas
            gb_comp.configure_column(
                field='NomeParaExibir', 
                headerName='Tarifário', 
                minWidth=200, flex=2.5, 
                filter='agTextColumnFilter', 
                cellRenderer=link_tooltip_renderer_js, # <--- APLICAR RENDERER
                cellStyle=cell_style_nome_tarifario_js,
                tooltipValueGetter=tooltip_nome_tarifario_getter_js, 
                tooltipComponent=custom_tooltip_component_js
            )
            
            # Ocultar colunas que agora estão integradas no NomeParaExibir ou não são necessárias
            if 'LinkAdesao' in df_resultados_comparacao_aggrid.columns:
                 gb_comp.configure_column('LinkAdesao', hide=True)
            if 'info_notas' in df_resultados_comparacao_aggrid.columns:
                 gb_comp.configure_column('info_notas', hide=True)
            if 'Comercializador' in df_resultados_comparacao_aggrid.columns:
                gb_comp.configure_column('Comercializador', hide=True)
            if 'Tipo' in df_resultados_comparacao_aggrid.columns:
                gb_comp.configure_column('Tipo', hide=True)

            # Colunas de Comercializador e Tipo (mais simples)
            if 'Comercializador' in df_resultados_comparacao_aggrid.columns:
                gb_comp.configure_column('Comercializador', headerName='Comercializador', minWidth=120, flex=1, cellStyle={'textAlign': 'center'})
            if 'Tipo' in df_resultados_comparacao_aggrid.columns:
                gb_comp.configure_column('Tipo', headerName='Tipo', minWidth=100, flex=1, cellStyle={'textAlign': 'center'})
            
            # Colunas de Custo (ex: "Total Simples (€)")
            if 'colunas_aggrid_custo_comp' in locals(): # Certifica-se de que a variável existe
                # Recriar min_max_data_comparativa_js e cell_style_cores_comparativa_js com base em df_resultados_comparacao_aggrid
                min_max_data_comparativa_js_local = {}
                for col_custo_comp_inner in colunas_aggrid_custo_comp:
                    if col_custo_comp_inner in df_resultados_comparacao_aggrid:
                        series_comp_inner = pd.to_numeric(df_resultados_comparacao_aggrid[col_custo_comp_inner], errors='coerce').dropna()
                        if not series_comp_inner.empty:
                            min_max_data_comparativa_js_local[col_custo_comp_inner] = {'min': series_comp_inner.min(), 'max': series_comp_inner.max()}
                        else:
                            min_max_data_comparativa_js_local[col_custo_comp_inner] = {'min': 0, 'max': 0}
                min_max_data_comparativa_json_string_local = json.dumps(min_max_data_comparativa_js_local)

                cell_style_cores_comparativa_js_local = JsCode(f"""
                function(params) {{
                    const minMaxConfig = {min_max_data_comparativa_json_string_local}; 
                    let style = {{ textAlign: 'center', borderRadius: '6px', padding: '10px 10px' }};
                    if (params.value == null || isNaN(parseFloat(params.value)) || !minMaxConfig[params.colDef.field]) return style;
                    const min_val = minMaxConfig[params.colDef.field].min;
                    const max_val = minMaxConfig[params.colDef.field].max;
                    if (max_val === min_val) {{ style.backgroundColor = 'lightgrey'; return style; }}
                    const normalized_value = Math.max(0, Math.min(1, (parseFloat(params.value) - min_val) / (max_val - min_val)));
                    const cL={{r:99,g:190,b:123}},cM={{r:255,g:255,b:255}},cH={{r:248,g:105,b:107}}; let r,g,b;
                    if(normalized_value < 0.5){{const t=normalized_value/0.5;r=Math.round(cL.r*(1-t)+cM.r*t);g=Math.round(cL.g*(1-t)+cM.g*t);b=Math.round(cL.b*(1-t)+cM.b*t);}}
                    else{{const t=(normalized_value-0.5)/0.5;r=Math.round(cM.r*(1-t)+cH.r*t);g=Math.round(cM.g*(1-t)+cH.g*t);b=Math.round(cM.b*(1-t)+cH.b*t);}}
                    style.backgroundColor=`rgb(${{r}},${{g}},${{b}})`;
                    if((r*0.299+g*0.587+b*0.114)<140)style.color='white';else style.color='black';
                    return style;
                }}
                """)
                
                # Tooltip para colunas de custo na tabela comparativa
                tooltip_custo_total_comparativa_js = JsCode("""
                function(params) {
                    if (!params.data || params.value == null) { return String(params.value); }
                    const colField = params.colDef.field; // Ex: "Total Simples (€)"
                    const tooltipDataKey = "Tooltip_" + colField; 
                    const tooltipData = params.data[tooltipDataKey];

                    const formatCurrency = (num) => (typeof num === 'number' && !isNaN(num)) ? num.toFixed(2) : 'N/A';
                    const formatUnitPrice = (num) => (typeof num === 'number' && !isNaN(num)) ? num.toFixed(4) : 'N/A'; // 4 casas para €/kWh ou €/dia

                    let tooltipParts = [];
                    
                    if (!tooltipData) { 
                        tooltipParts.push("<i>" + (params.data.NomeParaExibir || "Tarifário") + " - " + colField.replace("Total ", "").replace(" (€)", "") + "</i>");
                        tooltipParts.push("<b>Custo Total c/IVA: " + formatCurrency(parseFloat(params.value)) + " €</b>");
                        return tooltipParts.join("<br>");
                    }

                    tooltipParts.push("<i>" + (params.data.NomeParaExibir || "Tarifário") + " - " + colField.replace("Total ", "").replace(" (€)", "") + "</i>");
                    
                    // Preços Unitários s/IVA (NOVOS)
                    tooltipParts.push("<b>Preços Unitários (s/IVA):</b>");
                    let temPrecosUnitarios = false;
                    if (colField.includes("Simples")) {
                        if (tooltipData.tt_preco_unit_energia_S_siva != null) {
                            tooltipParts.push("&nbsp;&nbsp;Energia Simples: " + formatUnitPrice(tooltipData.tt_preco_unit_energia_S_siva) + " €/kWh");
                            temPrecosUnitarios = true;
                        }
                    } else if (colField.includes("Bi-horário")) {
                        if (tooltipData.tt_preco_unit_energia_V_siva != null) {
                            tooltipParts.push("&nbsp;&nbsp;Energia Vazio: " + formatUnitPrice(tooltipData.tt_preco_unit_energia_V_siva) + " €/kWh");
                            temPrecosUnitarios = true;
                        }
                        if (tooltipData.tt_preco_unit_energia_F_siva != null) {
                            tooltipParts.push("&nbsp;&nbsp;Energia Fora Vazio: " + formatUnitPrice(tooltipData.tt_preco_unit_energia_F_siva) + " €/kWh");
                            temPrecosUnitarios = true;
                        }
                    } else if (colField.includes("Tri-horário")) {
                        if (tooltipData.tt_preco_unit_energia_V_siva != null) {
                            tooltipParts.push("&nbsp;&nbsp;Energia Vazio: " + formatUnitPrice(tooltipData.tt_preco_unit_energia_V_siva) + " €/kWh");
                            temPrecosUnitarios = true;
                        }
                        if (tooltipData.tt_preco_unit_energia_C_siva != null) {
                            tooltipParts.push("&nbsp;&nbsp;Energia Cheias: " + formatUnitPrice(tooltipData.tt_preco_unit_energia_C_siva) + " €/kWh");
                            temPrecosUnitarios = true;
                        }
                        if (tooltipData.tt_preco_unit_energia_P_siva != null) {
                            tooltipParts.push("&nbsp;&nbsp;Energia Ponta: " + formatUnitPrice(tooltipData.tt_preco_unit_energia_P_siva) + " €/kWh");
                            temPrecosUnitarios = true;
                        }
                    }
                    if (tooltipData.tt_preco_unit_potencia_siva != null) {
                        tooltipParts.push("&nbsp;&nbsp;Potência: " + formatUnitPrice(tooltipData.tt_preco_unit_potencia_siva) + " €/dia");
                        temPrecosUnitarios = true;
                    }
                    if(temPrecosUnitarios){
                         tooltipParts.push("------------------------------------");
                    }

                    // Decomposição original do tooltip
                    tooltipParts.push("<b>Decomposição Custo Total:</b>");
                    tooltipParts.push("------------------------------------");
                                                
                    tooltipParts.push("Total Energia s/IVA: " + formatCurrency(tooltipData.tt_cte_energia_siva) + " €");
                    tooltipParts.push("Total Potência s/IVA: " + formatCurrency(tooltipData.tt_cte_potencia_siva) + " €");
                    if (tooltipData.tt_cte_iec_siva !== 0) { 
                        tooltipParts.push("IEC s/IVA: " + formatCurrency(tooltipData.tt_cte_iec_siva) + " €");
                    }
                    // ... (resto da decomposição original do tooltip, como estava antes) ...
                    if (tooltipData.tt_cte_dgeg_siva !== 0) { tooltipParts.push("DGEG s/IVA: " + formatCurrency(tooltipData.tt_cte_dgeg_siva) + " €");}
                    if (tooltipData.tt_cte_cav_siva !== 0) { tooltipParts.push("CAV s/IVA: " + formatCurrency(tooltipData.tt_cte_cav_siva) + " €");}
                    tooltipParts.push("<b>Subtotal s/IVA: " + formatCurrency(tooltipData.tt_cte_total_siva) + " €</b>");
                    tooltipParts.push("------------------------------------");
                    if (tooltipData.tt_cte_valor_iva_6_total !== 0) { tooltipParts.push("Valor IVA (6%): " + formatCurrency(tooltipData.tt_cte_valor_iva_6_total) + " €");}
                    if (tooltipData.tt_cte_valor_iva_23_total !== 0) { tooltipParts.push("Valor IVA (23%): " + formatCurrency(tooltipData.tt_cte_valor_iva_23_total) + " €");}
                    tooltipParts.push("<b>Subtotal c/IVA: " + formatCurrency(tooltipData.tt_cte_subtotal_civa) + " €</b>");
                    if (tooltipData.tt_cte_desc_finais_valor !== 0 || tooltipData.tt_cte_acres_finais_valor !== 0) { tooltipParts.push("------------------------------------");}
                    if (tooltipData.tt_cte_desc_finais_valor !== 0) { tooltipParts.push("Outros Descontos: -" + formatCurrency(tooltipData.tt_cte_desc_finais_valor) + " €");}
                    if (tooltipData.tt_cte_acres_finais_valor !== 0) { tooltipParts.push("Outros Acréscimos: +" + formatCurrency(tooltipData.tt_cte_acres_finais_valor) + " €");}
                    if (tooltipData.tt_cte_desc_finais_valor !== 0 || tooltipData.tt_cte_acres_finais_valor !== 0) { tooltipParts.push("------------------------------------");}
                    tooltipParts.push("<b>Custo Total c/IVA: " + formatCurrency(parseFloat(params.value)) + " €</b>");

                    return tooltipParts.filter(part => part !== "").join("<br>");
                }
                """)

                for col_custo_nome_comp in colunas_aggrid_custo_comp:
                    if col_custo_nome_comp in df_resultados_comparacao_aggrid.columns:
                        gb_comp.configure_column(
                            field=col_custo_nome_comp,
                            headerName=col_custo_nome_comp.replace("Total ", "").replace(" (€)", ""),
                            type=["numericColumn"],
                            valueFormatter=JsCode("function(params) { if(params.value == null) return '-'; return Number(params.value).toFixed(2); }"),
                            cellStyle=cell_style_cores_comparativa_js_local,
                            tooltipValueGetter=tooltip_custo_total_comparativa_js,
                            tooltipComponent=custom_tooltip_component_js,
                            minWidth=100, flex=1
                        )
            
            # Ocultar colunas de dados de tooltip que não sejam as de nome/comercializador/tipo ou os totais
            colunas_de_dados_tooltip_para_comparativa = [
                col for col in df_resultados_comparacao_aggrid.columns 
                if col.startswith("Tooltip_Total ")
            ]
            for col_ocultar_tooltip_comp in colunas_de_dados_tooltip_para_comparativa:
                 gb_comp.configure_column(field=col_ocultar_tooltip_comp, hide=True)


            gb_comp.configure_grid_options(
                domLayout='autoHeight',
                suppressContextMenu=True,
                tooltipShowDelay=200, # Atraso para mostrar tooltip em ms
                tooltipMouseTrack=True # Tooltip segue o rato
            )
            gridOptions_comp = gb_comp.build()
            grid_response_comp = AgGrid(
                df_resultados_comparacao_aggrid,
                gridOptions=gridOptions_comp,
                custom_css=custom_css,
                allow_unsafe_jscode=True,
                fit_columns_on_grid_load=True,
                height = min( (len(df_resultados_comparacao_aggrid) + 1) * 40 + 5, 600),
                theme='alpine',
                key="aggrid_comparacao_opcoes_final",
                enable_enterprise_modules=True
            )
        st.markdown("<a id='exportar-excel-comparacao'></a>", unsafe_allow_html=True)

        st.markdown("---") # Separador
        with st.expander("📥 Exportar Tabela de Comparação para Excel"):
            if not df_resultados_comparacao_aggrid.empty:
                # Colunas visíveis na AgGrid comparativa por defeito
                default_cols_export_comp = ['NomeParaExibir']
                if 'colunas_aggrid_custo_comp' in locals():
                    default_cols_export_comp.extend(colunas_aggrid_custo_comp)

                # Todas as colunas disponíveis no DataFrame da tabela comparativa
                all_cols_comp_df = df_resultados_comparacao_aggrid.columns.tolist()
                
                # Outras colunas que podem ser úteis para exportar (originalmente ocultas na AgGrid)
                additional_useful_cols_comp = ['Comercializador', 'Tipo', 'LinkAdesao', 'info_notas']
                tooltip_data_cols_comp = [col for col in all_cols_comp_df if col.startswith("Tooltip_Total ")]
                
                # Construir lista de opções para o multiselect
                export_options_comp = default_cols_export_comp[:] # Começa com os defaults
                for col in additional_useful_cols_comp + tooltip_data_cols_comp:
                    if col in all_cols_comp_df and col not in export_options_comp:
                        export_options_comp.append(col)
                # Adicionar restantes colunas se houver alguma que não foi coberta
                for col in all_cols_comp_df:
                    if col not in export_options_comp:
                        export_options_comp.append(col)

                cols_to_export_comp_selected = st.multiselect(
                    "Selecione as colunas para exportar (Tabela Comparativa):",
                    options=export_options_comp,
                    default=default_cols_export_comp,
                    key="cols_export_excel_comp_selector"
                )

                # --- Função de interpolação de cores ---
                def gerar_estilo_completo_para_valor(valor, minimo, maximo):
                    estilo_css_final = 'text-align: center;' 
                    if pd.isna(valor): return estilo_css_final
                    try: val_float = float(valor)
                    except ValueError: return estilo_css_final
                    if maximo == minimo or minimo is None or maximo is None: return estilo_css_final
        
                    midpoint = (minimo + maximo) / 2
                    r_bg, g_bg, b_bg = 255,255,255 
                    verde_rgb, branco_rgb, vermelho_rgb = (99,190,123), (255,255,255), (248,105,107)

                    if val_float <= midpoint:
                        ratio = (val_float - minimo) / (midpoint - minimo) if midpoint != minimo else 0.0
                        r_bg = int(verde_rgb[0]*(1-ratio) + branco_rgb[0]*ratio)
                        g_bg = int(verde_rgb[1]*(1-ratio) + branco_rgb[1]*ratio)
                        b_bg = int(verde_rgb[2]*(1-ratio) + branco_rgb[2]*ratio)
                    else:
                        ratio = (val_float - midpoint) / (maximo - midpoint) if maximo != midpoint else 0.0
                        r_bg = int(branco_rgb[0]*(1-ratio) + vermelho_rgb[0]*ratio)
                        g_bg = int(branco_rgb[1]*(1-ratio) + vermelho_rgb[1]*ratio)
                        b_bg = int(branco_rgb[2]*(1-ratio) + vermelho_rgb[2]*ratio)
                
                    estilo_css_final += f' background-color: #{r_bg:02X}{g_bg:02X}{b_bg:02X};'
                    luminancia = (0.299 * r_bg + 0.587 * g_bg + 0.114 * b_bg)
                    cor_texto_css = '#000000' if luminancia > 140 else '#FFFFFF'
                    estilo_css_final += f' color: {cor_texto_css};'
                    return estilo_css_final

                                        # --- Função de estilo principal a ser aplicada ao DataFrame ---
                def estilo_geral_dataframe_para_exportar(df_a_aplicar_estilo, tipos_reais_para_estilo_serie, min_max_config_para_cores, nome_coluna_tarifario="Tarifário"):
                    df_com_estilos = pd.DataFrame('', index=df_a_aplicar_estilo.index, columns=df_a_aplicar_estilo.columns)
               
                    # Cores default (pode ajustar)
                    cor_fundo_indexado_media_css_local = "#FFE699"
                    cor_texto_indexado_media_css_local = "black"
                    cor_fundo_indexado_dinamico_css_local = "#4D79BC"  
                    cor_texto_indexado_dinamico_css_local = "white"

                    for nome_coluna_df in df_a_aplicar_estilo.columns:
                        # Estilo para colunas de custo (Total (€) na detalhada, Total [Opção] (€) na comparativa)
                        if nome_coluna_df in min_max_config_para_cores: # min_max_config_para_cores é o min_max_data_X_json_string convertido para dict
                            try:
                                serie_valores_col = pd.to_numeric(df_a_aplicar_estilo[nome_coluna_df], errors='coerce')
                                min_valor_col = min_max_config_para_cores[nome_coluna_df]['min']
                                max_valor_col = min_max_config_para_cores[nome_coluna_df]['max']
                                df_com_estilos[nome_coluna_df] = serie_valores_col.apply(
                                    lambda valor_v: gerar_estilo_completo_para_valor(valor_v, min_valor_col, max_valor_col)
                                )
                            except Exception as e_estilo_custo:
                                print(f"Erro ao aplicar estilo de custo à coluna {nome_coluna_df}: {e_estilo_custo}")
                                df_com_estilos[nome_coluna_df] = 'text-align: center;' 

                        elif nome_coluna_df == nome_coluna_tarifario: # 'Tarifário' ou 'NomeParaExibir'
                            estilos_col_tarif_lista = []
                            for idx_linha_df, valor_nome_col_tarif in df_a_aplicar_estilo[nome_coluna_df].items():
                                tipo_tarif_str = tipos_reais_para_estilo_serie.get(idx_linha_df, '') if tipos_reais_para_estilo_serie is not None else ''
        
                                est_css_tarif = 'text-align: center; padding: 4px;' 
                                bg_cor_val, fonte_cor_val, fonte_peso_val = "#FFFFFF", "#000000", "normal" # Default Fixo/Outro (branco)

                                if isinstance(valor_nome_col_tarif, str) and valor_nome_col_tarif.startswith("O Meu Tarifário"):
                                    bg_cor_val, fonte_cor_val, fonte_peso_val = "#FF0000", "#FFFFFF", "bold"
                                elif tipo_tarif_str == 'Indexado Média':
                                    bg_cor_val, fonte_cor_val = cor_fundo_indexado_media_css_local, cor_texto_indexado_media_css_local
                                elif tipo_tarif_str == 'Indexado quarto-horário':
                                    bg_cor_val, fonte_cor_val = cor_fundo_indexado_dinamico_css_local, cor_texto_indexado_dinamico_css_local
                                elif tipo_tarif_str == 'Fixo': # Para dar um fundo um pouco diferente aos fixos
                                    bg_cor_val = "#F0F0F0" 
                
                                est_css_tarif += f' background-color: {bg_cor_val}; color: {fonte_cor_val}; font-weight: {fonte_peso_val};'
                                estilos_col_tarif_lista.append(est_css_tarif)
                            df_com_estilos[nome_coluna_df] = estilos_col_tarif_lista
                        else: # Outras colunas de texto ou sem estilização de cor baseada em valor
                            df_com_estilos[nome_coluna_df] = 'text-align: center;'
                    return df_com_estilos


                def exportar_excel_completo(df_para_exportar, styler_obj, resumo_html_para_excel, poupanca_texto_para_excel, identificador_cor_cabecalho):
                    output_excel_buffer = io.BytesIO() 
                    with pd.ExcelWriter(output_excel_buffer, engine='openpyxl') as writer_excel:
                        sheet_name_excel = 'Tiago Felicia - Eletricidade'

                        # --- Escrever Resumo ---
                        dados_resumo_formatado = []
                        if resumo_html_para_excel:
                            soup_resumo = BeautifulSoup(resumo_html_para_excel, "html.parser")
                            
                            # --- LÓGICA ALTERADA PARA REARRANJAR O RESUMO ---
                            titulo_resumo = soup_resumo.find('h5')
                            if titulo_resumo:
                                dados_resumo_formatado.append([titulo_resumo.get_text(strip=True), None])

                            itens_lista_resumo = soup_resumo.find_all('li')
                            linha_filtros_texto = ""
                            linha_potencia_texto = ""
                            outras_linhas_resumo = []

                            for item in itens_lista_resumo:
                                # Usar .get_text() para obter o conteúdo limpo do item da lista
                                texto_item = item.get_text(separator=' ', strip=True)
                                
                                if "Segmento:" in texto_item:
                                    linha_filtros_texto = texto_item
                                elif "kVA" in texto_item:
                                    linha_potencia_texto = texto_item
                                else:
                                    # Processar as outras linhas normalmente
                                    parts = texto_item.split(':', 1)
                                    if len(parts) == 2:
                                        outras_linhas_resumo.append([parts[0].strip() + ":", parts[1].strip()])
                                    else:
                                        outras_linhas_resumo.append([texto_item, None])
                            
                            # Adicionar a linha combinada primeiro, na ordem que pediu
                            if linha_filtros_texto or linha_potencia_texto:
                                dados_resumo_formatado.append([linha_filtros_texto, linha_potencia_texto])
                            
                            # Adicionar o resto do resumo
                            dados_resumo_formatado.extend(outras_linhas_resumo)
                            # --- FIM DA LÓGICA ALTERADA ---
        
                        df_resumo_obj = pd.DataFrame(dados_resumo_formatado)

                        # 1. Deixe o Pandas criar/ativar a folha na primeira escrita
                        df_resumo_obj.to_excel(writer_excel, sheet_name=sheet_name_excel, index=False, header=False, startrow=0)

                        # 2. AGORA obtenha a referência à worksheet, que certamente existe
                        worksheet_excel = writer_excel.sheets[sheet_name_excel]


                        # Formatar Resumo (Negrito)
                        bold_font_obj = Font(bold=True) # Font já deve estar importado de openpyxl.styles
                        for i_resumo in range(len(df_resumo_obj)):
                            excel_row_idx_resumo = i_resumo + 1 # Linhas do Excel são 1-based
                            cell_resumo_rotulo = worksheet_excel.cell(row=excel_row_idx_resumo, column=1)
                            cell_resumo_rotulo.font = bold_font_obj
                            if df_resumo_obj.shape[1] > 1 and pd.notna(df_resumo_obj.iloc[i_resumo, 1]):
                                cell_resumo_valor = worksheet_excel.cell(row=excel_row_idx_resumo, column=2)
                                cell_resumo_valor.font = bold_font_obj

                        worksheet_excel.column_dimensions['A'].width = 35
                        worksheet_excel.column_dimensions['B'].width = 65

                        linha_atual_no_excel_escrita = len(df_resumo_obj) + 1

                        # --- Escrever Mensagem de Poupança ---
                        if poupanca_texto_para_excel: # Verifica se há texto para a mensagem de poupança
                            linha_atual_no_excel_escrita += 1 # Adiciona uma linha em branco

                            cor_p = st.session_state.get('poupanca_excel_cor', "000000") # Cor do session_state
                            negrito_p = st.session_state.get('poupanca_excel_negrito', False) # Negrito do session_state
            
                            poupanca_cell_escrita = worksheet_excel.cell(row=linha_atual_no_excel_escrita, column=1, value=poupanca_texto_para_excel)
                            poupanca_font_escrita = Font(bold=negrito_p, color=cor_p)
                            poupanca_cell_escrita.font = poupanca_font_escrita

                            # --- MODIFICAÇÃO PARA JUNTAR CÉLULAS ---
                            worksheet_excel.merge_cells(start_row=linha_atual_no_excel_escrita, start_column=1, end_row=linha_atual_no_excel_escrita, end_column=6)

                            # Aplicar alinhamento à célula mesclada (a célula do canto superior esquerdo, poupanca_cell_escrita)
                            poupanca_cell_escrita.alignment = Alignment(wrap_text=True, horizontal='left', vertical='top')

                            linha_atual_no_excel_escrita += 1 # Avança para a próxima linha após a mensagem de poupança
        
                        linha_inicio_tab_dados_excel = linha_atual_no_excel_escrita + 3

                        # --- Adicionar linha de informação da simulação ---
                        # Adiciona uma linha em branco antes desta nova linha de informação
                        linha_info_simulacao_excel = linha_atual_no_excel_escrita + 1 

                        data_hoje_obj = datetime.date.today() # datetime já deve estar importado
                        data_hoje_formatada_str = data_hoje_obj.strftime('%d/%m/%Y')

                        espacador_info = "                                                                      " #: 70 espaços

                        texto_completo_info = (
                            f"          Simulação em {data_hoje_formatada_str}{espacador_info}"
                            f"https://www.tiagofelicia.pt{espacador_info}"
                            f"Tiago Felícia"
                        )

                        # Escrever o texto completo na primeira célula da área a ser mesclada (Coluna A)
                        info_cell = worksheet_excel.cell(row=linha_info_simulacao_excel, column=1)
                        info_cell.value = texto_completo_info
            
                        # Aplicar negrito à célula
                        # Reutilizar bold_font_obj que já foi definido para o resumo, ou criar um novo se precisar de formatação diferente.
                        # Assumindo que bold_font_obj é Font(bold=True) e está no escopo.
                        # Se não, defina-o: from openpyxl.styles import Font; bold_font_obj = Font(bold=True)
                        if 'bold_font_obj' in locals() or 'bold_font_obj' in globals():
                             info_cell.font = bold_font_obj # Reutiliza o bold_font_obj do resumo
                        else:
                             info_cell.font = Font(bold=True) # Cria um novo se não existir

                        # Mesclar as colunas A, B, C, e D para esta linha
                        worksheet_excel.merge_cells(start_row=linha_info_simulacao_excel, start_column=1, end_row=linha_info_simulacao_excel, end_column=6)

                        # Ajustar alinhamento da célula mesclada (info_cell é a célula do topo-esquerda da área mesclada)
                        # Alinhado à esquerda, centralizado verticalmente, com quebra de linha se necessário.
                        info_cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True) 

                        # A linha de início para a tabela de dados principal virá depois desta linha de informação
                        # Adicionamos +1 para esta linha de informação e +1 para uma linha em branco antes da tabela
                        linha_inicio_tab_dados = linha_info_simulacao_excel + 2 
            
                        # --- Fim da adição da linha de informação ---

                        for row in worksheet_excel.iter_rows(min_row=1, max_row=worksheet_excel.max_row+100, min_col=1, max_col=20):
                            for cell in row:
                                cell.fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

                        # O Styler escreverá na mesma folha 'sheet_name_excel'
                        styler_obj.to_excel(
                            writer_excel,
                            sheet_name=sheet_name_excel,
                            index=False,
                            startrow=linha_inicio_tab_dados_excel - 1, # startrow é 0-indexed
                            columns=df_para_exportar.columns.tolist()
                        )

                        # Determina cor do cabeçalho conforme a opção horária e ciclo
                        opcao_horaria_lower = str(opcao_horaria).lower() if 'opcao_horaria' in locals() else "simples"
                        cor_fundo = "A6A6A6"   # padrão Simples
                        cor_fonte = "000000"    # padrão preto

                        if isinstance(identificador_cor_cabecalho, str):
                            id_lower = identificador_cor_cabecalho.lower()
                            if id_lower == "simples": # Já coberto pelo default
                                pass
                            elif "bi-horário" in id_lower and "diário" in id_lower:
                                cor_fundo = "A9D08E"; cor_fonte = "000000"
                            elif "bi-horário" in id_lower and "semanal" in id_lower:
                                cor_fundo = "8EA9DB"; cor_fonte = "000000"
                            elif "tri-horário" in id_lower and "diário" in id_lower:
                                cor_fundo = "BF8F00"; cor_fonte = "FFFFFF"
                            elif "tri-horário" in id_lower and "semanal" in id_lower:
                                cor_fundo = "C65911"; cor_fonte = "FFFFFF"

                        # linha_inicio_tab_dados_excel já existe na função e corresponde ao header da tabela
                        for col_idx, _ in enumerate(df_para_exportar.columns):
                            celula = worksheet_excel.cell(row=linha_inicio_tab_dados_excel, column=col_idx + 1)
                            celula.fill = PatternFill(start_color=cor_fundo, end_color=cor_fundo, fill_type="solid")
                            celula.font = Font(color=cor_fonte, bold=True)

                                # ---- INÍCIO: ADICIONAR LEGENDA DE CORES APÓS A TABELA ----
                        # Calcular a linha de início para a legenda
                        # df_para_exportar é o DataFrame que foi escrito pelo styler (ex: df_export_final)
                        numero_linhas_dados_tabela_principal = len(df_para_exportar)
                        # linha_inicio_tab_dados_excel é a linha do cabeçalho da tabela principal (1-indexada)
                        ultima_linha_tabela_principal = linha_inicio_tab_dados_excel + numero_linhas_dados_tabela_principal
                    
                        linha_legenda_bloco_inicio = ultima_linha_tabela_principal + 2 # Deixa uma linha em branco após a tabela

                        # Título da Legenda
                        titulo_legenda_cell = worksheet_excel.cell(row=linha_legenda_bloco_inicio, column=1, value="Tipos de Tarifário:")
                        # Reutilizar bold_font_obj se definido ou criar um novo
                        if 'bold_font_obj' in locals() or 'bold_font_obj' in globals():
                            titulo_legenda_cell.font = bold_font_obj
                        else:
                            titulo_legenda_cell.font = Font(bold=True)
                        worksheet_excel.merge_cells(start_row=linha_legenda_bloco_inicio, start_column=1, end_row=linha_legenda_bloco_inicio, end_column=6) # Mesclar para o título

                        linha_legenda_item_atual = linha_legenda_bloco_inicio + 1 # Primeira linha para item da legenda

                        titulo_legenda_cell.alignment = Alignment(horizontal='center', vertical='center')

                        itens_legenda_excel = [
                            {"cf": "FF0000", "ct": "FFFFFF", "b": True, "tA": "O Meu Tarifário", "tB": "Tarifário configurado pelo utilizador."},
                            {"cf": "FFE699", "ct": "000000", "b": False, "tA": "Indexado (Média)", "tB": "Preço de energia baseado na média OMIE do período."},
                            {"cf": "4D79BC", "ct": "FFFFFF", "b": False, "tA": "Indexado (Quarto-Horário)", "tB": "Preço de energia baseado nos valores OMIE horários/quarto-horários e perfil."},
                            {"cf": "F0F0F0", "ct": "333333", "b": False, "tA": "Fixo", "tB": "Preços de energia constantes", "borda_cor": "CCCCCC"}
                        ]
                    
                        # Definir larguras das colunas para a legenda (pode ajustar conforme necessário)
                        worksheet_excel.column_dimensions[get_column_letter(1)].width = 30 # Coluna A para a amostra/nome
                        worksheet_excel.column_dimensions[get_column_letter(2)].width = 200 # Coluna B para a descrição (será junta)

                        for item in itens_legenda_excel:
                            celula_A_legenda = worksheet_excel.cell(row=linha_legenda_item_atual, column=1, value=item["tA"])
                            celula_A_legenda.fill = PatternFill(start_color=item["cf"], end_color=item["cf"], fill_type="solid")
                            celula_A_legenda.font = Font(color=item["ct"], bold=item["b"])
                            celula_A_legenda.alignment = Alignment(horizontal='center', vertical='center', indent=1)

                            if "borda_cor" in item:
                                cor_borda_hex = item["borda_cor"]
                                borda_legenda_obj = Border(
                                    top=Side(border_style="thin", color=cor_borda_hex),
                                    left=Side(border_style="thin", color=cor_borda_hex),
                                    right=Side(border_style="thin", color=cor_borda_hex),
                                    bottom=Side(border_style="thin", color=cor_borda_hex)
                                )
                                celula_A_legenda.border = borda_legenda_obj
                        
                            celula_B_legenda = worksheet_excel.cell(row=linha_legenda_item_atual, column=2, value=item["tB"])
                            celula_B_legenda.alignment = Alignment(vertical='center', wrap_text=True, horizontal='left')
                            # Mesclar colunas B até D (ou ajuste conforme a largura desejada para a descrição)
                            worksheet_excel.merge_cells(start_row=linha_legenda_item_atual, start_column=2,
                                                        end_row=linha_legenda_item_atual, end_column=6) 
                        
                            worksheet_excel.row_dimensions[linha_legenda_item_atual].height = 20 # Ajustar altura da linha da legenda
                            linha_legenda_item_atual += 1
                        # ---- FIM: ADICIONAR LEGENDA DE CORES ----

                        # Ajustar largura das colunas da tabela principal
                        for col_idx_iter, col_nome_iter_width in enumerate(df_para_exportar.columns):
                            col_letra_iter = get_column_letter(col_idx_iter + 1) # get_column_letter já deve estar importado
                            if "Tarifário" in col_nome_iter_width :
                                 worksheet_excel.column_dimensions[col_letra_iter].width = 80    
                            elif "Total Simples (€)" == col_nome_iter_width :
                                worksheet_excel.column_dimensions[col_letra_iter].width = 33
                            elif "Total Bi-horário - Ciclo Diário (€)" in col_nome_iter_width :
                                worksheet_excel.column_dimensions[col_letra_iter].width = 33
                            elif "Total Bi-horário - Ciclo Semanal (€)" in col_nome_iter_width :
                                 worksheet_excel.column_dimensions[col_letra_iter].width = 33    
                            elif "Total Tri-horário - Ciclo Diário (€)" in col_nome_iter_width :
                             worksheet_excel.column_dimensions[col_letra_iter].width = 33    
                            elif "Total Tri-horário - Ciclo Semanal (€)" in col_nome_iter_width :
                                 worksheet_excel.column_dimensions[col_letra_iter].width = 33    
                            else: 
                                worksheet_excel.column_dimensions[col_letra_iter].width = 25


                    output_excel_buffer.seek(0)
                    return output_excel_buffer
                        # --- Fim da definição de exportar_excel_completo ---



                limit_export_comp_selected = st.selectbox(
                    "Número de tarifários a exportar (Tabela Comparativa):",
                    options=["Todos"] + [f"Top {i}" for i in [10, 20, 30, 40, 50]],
                    index=0,
                    key="limit_export_excel_comp"
                )

                if st.button("Preparar Download Excel (Tabela Comparativa)", key="btn_prep_excel_comp_final"):
                    if not cols_to_export_comp_selected:
                        st.warning("Por favor, selecione pelo menos uma coluna para exportar.")
                    else:
                        with st.spinner("A gerar ficheiro Excel da Tabela Comparativa..."):
                            
                            df_data_from_grid_comp = pd.DataFrame()
                            if 'grid_response_comp' in locals() and grid_response_comp and grid_response_comp['data'] is not None:
                                df_data_from_grid_comp = pd.DataFrame(grid_response_comp['data'])
                            else:
                                st.warning("Não foi possível obter os dados da grelha. A exportar com base na tabela original.")
                                df_data_from_grid_comp = df_resultados_comparacao_aggrid.copy()

                            if not df_data_from_grid_comp.empty:
                                
                                tipos_reais_para_estilo_comp = df_data_from_grid_comp['Tipo'] if 'Tipo' in df_data_from_grid_comp else pd.Series(dtype=str)

                                min_max_config_excel = {}
                                colunas_custo_para_cor = [col for col in df_data_from_grid_comp.columns if col.startswith("Total ")]
                                for col_custo in colunas_custo_para_cor:
                                    serie = pd.to_numeric(df_data_from_grid_comp[col_custo], errors='coerce').dropna()
                                    if not serie.empty:
                                        min_max_config_excel[col_custo] = {'min': serie.min(), 'max': serie.max()}
                                    else:
                                        min_max_config_excel[col_custo] = {'min': 0, 'max': 0}

                                valid_export_cols_comp = [col for col in cols_to_export_comp_selected if col in df_data_from_grid_comp.columns]
                                if valid_export_cols_comp:
                                    df_export_comp_final = df_data_from_grid_comp[valid_export_cols_comp].copy()
                                else:
                                    st.warning("Nenhuma das colunas selecionadas para exportação está presente nos dados atuais da tabela comparativa.")
                                    df_export_comp_final = pd.DataFrame()
                                
                                if not df_export_comp_final.empty and limit_export_comp_selected != "Todos":
                                    try:
                                        num_to_export_comp = int(limit_export_comp_selected.split(" ")[1])
                                        df_export_comp_final = df_export_comp_final.head(num_to_export_comp)
                                        tipos_reais_para_estilo_comp = tipos_reais_para_estilo_comp.head(num_to_export_comp)
                                    except Exception as e:
                                        st.warning(f"Não foi possível aplicar o limite de tarifários à tabela comparativa: {e}")
                                
                                if not df_export_comp_final.empty:
                                    if 'NomeParaExibir' in df_export_comp_final.columns:
                                        df_export_comp_final.rename(columns={'NomeParaExibir': 'Tarifário'}, inplace=True)

                                    # --- MUDANÇA PRINCIPAL AQUI: Arredondar os dados diretamente no DataFrame ---
                                    for col in df_export_comp_final.columns:
                                        if "(€)" in col:
                                            # Garante que a coluna é numérica antes de arredondar
                                            df_export_comp_final[col] = pd.to_numeric(df_export_comp_final[col], errors='coerce').round(2)
                                    
                                    styler_comp_excel = df_export_comp_final.style.apply(
                                        lambda df: estilo_geral_dataframe_para_exportar(df, tipos_reais_para_estilo_comp, min_max_config_excel, "Tarifário"),
                                        axis=None
                                    )
                                    
                                    # O .format() já não é estritamente necessário para o arredondamento, mas podemos mantê-lo para formatação (ex: na_rep)
                                    styler_comp_excel = styler_comp_excel.format(formatter="{:.2f}", na_rep="-")
                                    
                                    styler_comp_excel = styler_comp_excel.set_table_styles([
                                        {'selector': 'th', 'props': [('background-color', '#404040'), ('color', 'white'), ('font-weight', 'bold'), ('text-align', 'center'), ('border', '1px solid black'), ('padding', '5px')]},
                                        {'selector': 'td', 'props': [('border', '1px solid #dddddd'), ('padding', '4px')]}
                                    ]).hide(axis="index")

                                    output_excel_comp_bytes = exportar_excel_completo(
                                        df_export_comp_final,
                                        styler_comp_excel,
                                        html_resumo_final,
                                        st.session_state.get('poupanca_excel_texto', ""),
                                        "Comparativa"
                                    )

                                    timestamp_comp_dl = int(time.time())
                                    filename_comp = f"Tiago_Felicia_Eletricidade_Comparacao_{timestamp_comp_dl}.xlsx"
                                    
                                    st.download_button(
                                        label=f"📥 Descarregar Excel ({filename_comp})",
                                        data=output_excel_comp_bytes.getvalue(),
                                        file_name=filename_comp,
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                        key=f"btn_dl_excel_comp_{timestamp_comp_dl}"
                                    )
                                    st.success(f"{filename_comp} pronto para download!")
                                elif df_export_comp_final.empty and not cols_to_export_comp_selected:
                                   pass
                                else:
                                    st.warning("Nenhum dado para exportar com os filtros e colunas selecionados para a tabela comparativa.")
                            else:
                                st.info("Tabela comparativa está vazia, nada para exportar.")
    st.markdown("---")

        # --- FIM DO EXPANDER DE EXPORTAÇÃO DA TABELA COMPARATIVA ---

# --- INÍCIO DO BLOCO PARA TABELA DETALHADA (AgGrid Principal) ---

# --- Comparar Tarifários Fixos ---
tarifarios_filtrados_fixos = tf_processar[
    (tf_processar['opcao_horaria_e_ciclo'] == opcao_horaria) &
    (tf_processar['potencia_kva'] == potencia)
].copy()

if not tarifarios_filtrados_fixos.empty:
    for index, tarifario in tarifarios_filtrados_fixos.iterrows():
        # --- Get tariff specifics ---
        nome_tarifario = tarifario['nome']
        tipo_tarifario = tarifario['tipo']
        comercializador_tarifario = tarifario['comercializador']
        link_adesao_tf = tarifario.get('site_adesao')
        notas_tarifario_tf = tarifario.get('notas', '')
        segmento_tarifario = tarifario.get('segmento', '-')
        faturacao_tarifario = tarifario.get('faturacao', '-')
        pagamento_tarifario = tarifario.get('pagamento', '-')

        # --- Get Inputs and Flags ---
        preco_energia_input_tf = {}
        consumos_horarios_para_func_tf = {}
        if opcao_horaria.lower() == "simples":
            preco_energia_input_tf['S'] = tarifario.get('preco_energia_simples', 0.0)
            consumos_horarios_para_func_tf = {'S': consumo_simples}
        elif opcao_horaria.lower().startswith("bi"):
            preco_energia_input_tf['V'] = tarifario.get('preco_energia_vazio_bi', 0.0)
            preco_energia_input_tf['F'] = tarifario.get('preco_energia_fora_vazio', 0.0)
            consumos_horarios_para_func_tf = {'V': consumo_vazio, 'F': consumo_fora_vazio}
        elif opcao_horaria.lower().startswith("tri"):
            preco_energia_input_tf['V'] = tarifario.get('preco_energia_vazio_tri', 0.0)
            preco_energia_input_tf['C'] = tarifario.get('preco_energia_cheias', 0.0)
            preco_energia_input_tf['P'] = tarifario.get('preco_energia_ponta', 0.0)
            consumos_horarios_para_func_tf = {'V': consumo_vazio, 'C': consumo_cheias, 'P': consumo_ponta}

        preco_potencia_input_tf = tarifario.get('preco_potencia_dia', 0.0)

        # Flags (com defaults sensatos)
        tar_incluida_energia_tf = tarifario.get('tar_incluida_energia', True)
        tar_incluida_potencia_tf = tarifario.get('tar_incluida_potencia', True)
        financiamento_tse_incluido_tf = tarifario.get('financiamento_tse_incluido', True) # Assumindo que fixos geralmente incluem

        # --- Passo 1: Identificar Componentes Base (Sem IVA, Sem TS) ---
        tar_energia_regulada_tf = {}
        for periodo in preco_energia_input_tf.keys():
            tar_energia_regulada_tf[periodo] = obter_tar_energia_periodo(opcao_horaria, periodo, potencia, CONSTANTES)

        tar_potencia_regulada_tf = obter_tar_dia(potencia, CONSTANTES)

        preco_comercializador_energia_tf = {}
        for periodo, preco_in in preco_energia_input_tf.items():
            preco_in_float = float(preco_in or 0.0)
            if tar_incluida_energia_tf:
                preco_comercializador_energia_tf[periodo] = preco_in_float - tar_energia_regulada_tf.get(periodo, 0.0)
            else:
                preco_comercializador_energia_tf[periodo] = preco_in_float

        preco_potencia_input_tf_float = float(preco_potencia_input_tf or 0.0)
        if tar_incluida_potencia_tf:
            preco_comercializador_potencia_tf = preco_potencia_input_tf_float - tar_potencia_regulada_tf
        else:
            preco_comercializador_potencia_tf = preco_potencia_input_tf_float
        preco_comercializador_potencia_tf = max(0.0, preco_comercializador_potencia_tf) # Limitar a 0

        financiamento_tse_a_adicionar_tf = FINANCIAMENTO_TSE_VAL if not financiamento_tse_incluido_tf else 0.0

        # --- Passo 2: Calcular Componentes TAR Finais (Com Desconto TS, Sem IVA) ---
        tar_energia_final_tf = {}
        tar_potencia_final_dia_tf = tar_potencia_regulada_tf

        if tarifa_social: # Flag global
            desconto_ts_energia = obter_constante('Desconto TS Energia', CONSTANTES)
            desconto_ts_potencia_dia = obter_constante(f'Desconto TS Potencia {potencia}', CONSTANTES)
            for periodo, tar_reg in tar_energia_regulada_tf.items():
                tar_energia_final_tf[periodo] = tar_reg - desconto_ts_energia
            tar_potencia_final_dia_tf = max(0.0, tar_potencia_regulada_tf - desconto_ts_potencia_dia)
        else:
            tar_energia_final_tf = tar_energia_regulada_tf.copy()

    # --- INÍCIO: CAMPOS PARA TOOLTIPS FIXOS ---
        # Para o tooltip do Preço Energia:
        componentes_tooltip_energia_dict_tf = {} # Dicionário para os componentes de energia deste tarifário

        # Flag global 'tarifa_social'
        ts_global_ativa = tarifa_social # Flag global de TS

        # Loop pelos períodos de energia (S, V, F, C, P) que existem para este tarifário
        for periodo_key_tf in preco_comercializador_energia_tf.keys():
        
            comp_comerc_energia_base_tf = preco_comercializador_energia_tf.get(periodo_key_tf, 0.0)
            tar_bruta_energia_periodo_tf = tar_energia_regulada_tf.get(periodo_key_tf, 0.0)
        
            # Flag 'financiamento_tse_incluido_tf' lida do Excel para ESTE tarifário fixo
            tse_declarado_incluido_excel_tf = financiamento_tse_incluido_tf 
        
            tse_valor_nominal_const_tf = FINANCIAMENTO_TSE_VAL
        
            ts_aplicada_energia_flag_para_tooltip_tf = ts_global_ativa
            desconto_ts_energia_unitario_para_tooltip_tf = 0.0
            if ts_global_ativa:
                desconto_ts_energia_unitario_para_tooltip_tf = obter_constante('Desconto TS Energia', CONSTANTES)

            # Usar os nomes EXATOS que o JavaScript espera
            componentes_tooltip_energia_dict_tf[f'tooltip_energia_{periodo_key_tf}_comerc_sem_tar'] = comp_comerc_energia_base_tf
            componentes_tooltip_energia_dict_tf[f'tooltip_energia_{periodo_key_tf}_tar_bruta'] = tar_bruta_energia_periodo_tf
            componentes_tooltip_energia_dict_tf[f'tooltip_energia_{periodo_key_tf}_tse_declarado_incluido'] = tse_declarado_incluido_excel_tf
            componentes_tooltip_energia_dict_tf[f'tooltip_energia_{periodo_key_tf}_tse_valor_nominal'] = tse_valor_nominal_const_tf
            componentes_tooltip_energia_dict_tf[f'tooltip_energia_{periodo_key_tf}_ts_aplicada_flag'] = ts_aplicada_energia_flag_para_tooltip_tf
            componentes_tooltip_energia_dict_tf[f'tooltip_energia_{periodo_key_tf}_ts_desconto_valor'] = desconto_ts_energia_unitario_para_tooltip_tf
    
        desconto_ts_potencia_valor_aplicado = 0.0
        if tarifa_social: # Flag global
            desconto_ts_potencia_dia_bruto = obter_constante(f'Desconto TS Potencia {potencia}', CONSTANTES)
            # O desconto efetivamente aplicado é o mínimo entre o desconto e a própria TAR
            desconto_ts_potencia_valor_aplicado = min(tar_potencia_regulada_tf, desconto_ts_potencia_dia_bruto)

        # Para o tooltip do Preço Potência Fixos:
        componentes_tooltip_potencia_dict_tf = {
            'tooltip_pot_comerc_sem_tar': preco_comercializador_potencia_tf,
            'tooltip_pot_tar_bruta': tar_potencia_regulada_tf,
            'tooltip_pot_ts_aplicada': ts_global_ativa,
            'tooltip_pot_desconto_ts_valor': desconto_ts_potencia_valor_aplicado
        }
 
        # --- Passo 3: Calcular Preço Final Energia (€/kWh, Sem IVA) ---
        preco_energia_final_sem_iva_tf = {}
        for periodo in preco_comercializador_energia_tf.keys():
            preco_energia_final_sem_iva_tf[periodo] = (
                preco_comercializador_energia_tf[periodo]
                + tar_energia_final_tf.get(periodo, 0.0)
                + financiamento_tse_a_adicionar_tf
            )

        # --- Passo 4: Calcular Componentes Finais Potência (€/dia, Sem IVA) ---
        preco_comercializador_potencia_final_sem_iva_tf = preco_comercializador_potencia_tf
        tar_potencia_final_dia_sem_iva_tf = tar_potencia_final_dia_tf

        # --- Passo 5: Calcular Custo Total Energia (Com IVA) ---
        custo_energia_tf_com_iva = calcular_custo_energia_com_iva(
            consumo,
            preco_energia_final_sem_iva_tf.get('S') if opcao_horaria.lower() == "simples" else None,
            {p: v for p, v in preco_energia_final_sem_iva_tf.items() if p != 'S'},
            dias, potencia, opcao_horaria,
            consumos_horarios_para_func_tf, # Já definido acima
            familia_numerosa
        )

        # --- Passo 6: Calcular Custo Total Potência (Com IVA) ---
        custo_potencia_tf_com_iva = calcular_custo_potencia_com_iva_final(
            preco_comercializador_potencia_final_sem_iva_tf,
            tar_potencia_final_dia_sem_iva_tf,
            dias,
            potencia
        )

        comercializador_tarifario_tf = tarifario['comercializador'] # Nome do comercializador deste tarifário

        # --- Passo 7: Calcular Taxas Adicionais ---
        taxas_tf = calcular_taxas_adicionais(
            consumo, dias, tarifa_social,
            valor_dgeg_user, valor_cav_user,
            nome_comercializador_atual=comercializador_tarifario_tf,
            mes_selecionado_simulacao=mes, ano_simulacao_atual=ano_atual
        )

        # --- Passo 8: Calcular Custo Total Final ---
        custo_total_antes_desc_fatura_tf = (
        custo_energia_tf_com_iva['custo_com_iva'] +
        custo_potencia_tf_com_iva['custo_com_iva'] +
        taxas_tf['custo_com_iva']
    )

        # Guardar o nome original do tarifário do Excel
        nome_tarifario_excel = str(tarifario['nome'])
        nome_a_exibir = nome_tarifario_excel # Começa com o nome original

        e_mes_completo_selecionado = False
        try:
            dias_no_mes_do_input_widget = dias_mes[mes] # Dias no mês selecionado pelo widget 'mes'
            primeiro_dia_do_mes_widget = datetime.date(ano_atual, mes_num, 1)
            ultimo_dia_do_mes_widget = datetime.date(ano_atual, mes_num, dias_no_mes_do_input_widget)
            if data_inicio == primeiro_dia_do_mes_widget and data_fim == ultimo_dia_do_mes_widget:
                e_mes_completo_selecionado = True
        except Exception: # Lidar com possíveis erros de data, embora improváveis aqui
            e_mes_completo_selecionado = False

        # --- Aplicar desconto_fatura_mes ---
        desconto_fatura_mensal_tf = 0.0
        if 'desconto_fatura_mes' in tarifario.index and pd.notna(tarifario['desconto_fatura_mes']): # Usar .index para Series
            try:
                desconto_fatura_mensal_tf = float(tarifario['desconto_fatura_mes'])
                if desconto_fatura_mensal_tf > 0:
                    # Adiciona nota sobre o desconto de fatura ao nome a ser exibido
                    nome_a_exibir += f" (+desc. fat. {desconto_fatura_mensal_tf:.2f}€/mês)"
            except ValueError:
                desconto_fatura_mensal_tf = 0.0
        
        if e_mes_completo_selecionado:
            desconto_fatura_periodo_tf = desconto_fatura_mensal_tf
        else:
            desconto_fatura_periodo_tf = (desconto_fatura_mensal_tf / 30.0) * dias if dias > 0 else 0

        # Custo após o desconto de fatura do Excel
        custo_apos_desc_fatura_excel_tf = custo_total_antes_desc_fatura_tf - desconto_fatura_periodo_tf
        # --- FIM desconto_fatura_mes ---

        # Adicionar Quota ACP se aplicável
        custo_apos_acp_tf = custo_apos_desc_fatura_excel_tf
        quota_acp_periodo = 0.0
        # A flag incluir_quota_acp vem da checkbox geral
        # VALOR_QUOTA_ACP_MENSAL (constante global)
        if incluir_quota_acp and isinstance(nome_tarifario_excel, str) and nome_tarifario_excel.startswith("Goldenergy - ACP"):
            if e_mes_completo_selecionado:
                quota_acp_periodo = VALOR_QUOTA_ACP_MENSAL
                custo_apos_acp_tf += quota_acp_periodo
                nome_a_exibir += f" (INCLUI Quota ACP - {VALOR_QUOTA_ACP_MENSAL:.2f} €/mês)"
            else:
                quota_acp_periodo = (VALOR_QUOTA_ACP_MENSAL / 30.0) * dias if dias > 0 else 0
                custo_apos_acp_tf += quota_acp_periodo
                nome_a_exibir += f" (INCLUI Quota ACP - {VALOR_QUOTA_ACP_MENSAL:.2f} €/mês)"
            # custo_apos_acp_tf já adiciona quota_acp_periodo

        # Inicializar o custo que será ajustado por este novo desconto MEO
        custo_antes_desconto_meo_tf = custo_apos_acp_tf # Ou custo_apos_desc_fatura_excel_tf se não houver ACP
        desconto_meo_aplicado_periodo = 0.0

        # --- LÓGICA PARA DESCONTO ESPECIAL MEO (NOVO BLOCO) ---
        # Condições: Nome do tarifário e consumo
        nome_original_lower = str(nome_tarifario_excel).lower()
    
        consumo_mensal_equivalente = 0
        if dias > 0:
            consumo_mensal_equivalente = (consumo / dias) * 30.0
    
        # Verifica se o nome contém a frase chave e se o consumo atinge o limite
        if "meo energia - tarifa fixa - clientes meo" in nome_original_lower and consumo_mensal_equivalente >= 216:
            desconto_meo_mensal_base = 0.0
            opcao_horaria_lower = str(opcao_horaria).lower()

            if opcao_horaria_lower == "simples":
                desconto_meo_mensal_base = 0
            elif opcao_horaria_lower.startswith("bi"): # Cobre "bi-horário semanal" e "bi-horário diário"
                desconto_meo_mensal_base = 0
            elif opcao_horaria_lower.startswith("tri"): # Cobre "tri-horário semanal" e "tri-horário diário"
                desconto_meo_mensal_base = 0
        
            if desconto_meo_mensal_base > 0 and dias > 0:
                desconto_meo_aplicado_periodo = (desconto_meo_mensal_base / 30.0) * dias
                custo_antes_desconto_meo_tf -= desconto_meo_aplicado_periodo # Aplicar o desconto
            
                # Adicionar nota ao nome do tarifário
                nome_a_exibir += f" (Desconto MEO Clientes {desconto_meo_aplicado_periodo:.2f}€ incl.)"
        # --- FIM DA LÓGICA DESCONTO ESPECIAL MEO ---

        # --- LÓGICA PARA DESCONTO CONTINENTE (NOVO BLOCO) ---
        # A base para o desconto Continente deve ser o custo APÓS o desconto MEO
        custo_base_para_continente_tf = custo_antes_desconto_meo_tf
        custo_total_estimado_final_tf = custo_base_para_continente_tf
        valor_X_desconto_continente = 0.0

        if desconto_continente and isinstance(nome_tarifario_excel, str) and nome_tarifario_excel.startswith("Galp & Continente"):
    
            # PASSO ADICIONAL: CALCULAR O CUSTO BRUTO (SEM TARIFA SOCIAL) APENAS PARA ESTE DESCONTO
    
            # 1. Preço unitário bruto da energia (sem IVA e sem desconto TS)
            preco_energia_bruto_sem_iva = {}
            for p in preco_comercializador_energia_tf.keys():
                preco_energia_bruto_sem_iva[p] = (
                    preco_comercializador_energia_tf.get(p, 0.0) + 
                    tar_energia_regulada_tf.get(p, 0.0) + # <--- USA A TAR BRUTA, sem desconto TS
                    financiamento_tse_a_adicionar_tf
                )
    
            # 2. Preço unitário bruto da potência (sem IVA e sem desconto TS)
            # Requer as componentes brutas
            preco_comercializador_potencia_bruto = preco_comercializador_potencia_tf 
            tar_potencia_bruta = tar_potencia_regulada_tf # <--- USA A TAR BRUTA, sem desconto TS

            # 3. Calcular o custo bruto COM IVA para a energia e potência
            custo_energia_bruto_cIVA = calcular_custo_energia_com_iva(
                consumo,
                preco_energia_bruto_sem_iva.get('S'),
                {k: v for k, v in preco_energia_bruto_sem_iva.items() if k != 'S'},
                dias, potencia, opcao_horaria, consumos_horarios_para_func_tf, familia_numerosa
            )
            custo_potencia_bruto_cIVA = calcular_custo_potencia_com_iva_final(
                preco_comercializador_potencia_bruto,
                tar_potencia_bruta,
                dias, potencia
            )

            # 4. Calcular o valor do cupão sobre os custos brutos
            valor_X_desconto_continente_energia = custo_energia_bruto_cIVA['custo_com_iva'] * 0.10
            valor_X_desconto_continente_potencia = custo_potencia_bruto_cIVA['custo_com_iva'] * 0.10
            valor_X_desconto_continente = valor_X_desconto_continente_energia + valor_X_desconto_continente_potencia

            # 5. Aplicar o valor do cupão ao custo final (que já tem o desconto TS, se aplicável)
            custo_total_estimado_final_tf = custo_base_para_continente_tf - valor_X_desconto_continente
            nome_a_exibir += f" (INCLUI desc. Cont. de {valor_X_desconto_continente:.2f}€, s/ desc. Cont.={custo_base_para_continente_tf:.2f}€)"        # --- FIM DA LÓGICA DESCONTO CONTINENTE ---

        # --- Passo 9: Preparar Resultados para Exibição ---
        valores_energia_exibir_tf = {} # Recalcular ou usar o já calculado 'preco_energia_final_sem_iva_tf'
        for p, v_energia_sem_iva in preco_energia_final_sem_iva_tf.items(): # Use os preços SEM IVA para exibição na tabela
            periodo_nome = ""
            if p == 'S': periodo_nome = "Simples"
            elif p == 'V': periodo_nome = "Vazio"
            elif p == 'F': periodo_nome = "Fora Vazio"
            elif p == 'C': periodo_nome = "Cheias"
            elif p == 'P': periodo_nome = "Ponta"
            if periodo_nome:
                valores_energia_exibir_tf[f'{periodo_nome} (€/kWh)'] = round(v_energia_sem_iva, 4)

        preco_potencia_total_final_sem_iva_tf = preco_comercializador_potencia_final_sem_iva_tf + tar_potencia_final_dia_sem_iva_tf

        # --- PASSO X: CALCULAR CUSTOS COM IVA E OBTER DECOMPOSIÇÃO PARA TOOLTIP ---

        # ENERGIA (Tarifários Fixos)
        preco_energia_simples_para_iva_tf = None
        precos_energia_horarios_para_iva_tf = {}
        if opcao_horaria.lower() == "simples":
            preco_energia_simples_para_iva_tf = preco_energia_final_sem_iva_tf.get('S')
        else:
            precos_energia_horarios_para_iva_tf = {
                p: val for p, val in preco_energia_final_sem_iva_tf.items() if p != 'S'
            }
            
        decomposicao_custo_energia_tf = calcular_custo_energia_com_iva(
            consumo, # Consumo total global
            preco_energia_simples_para_iva_tf,
            precos_energia_horarios_para_iva_tf,
            dias, potencia, opcao_horaria,
            consumos_horarios_para_func_tf, # Dicionário de consumos por período para este tarifário
            familia_numerosa
        )
        custo_energia_tf_com_iva = decomposicao_custo_energia_tf['custo_com_iva']
        tt_cte_energia_siva_tf = decomposicao_custo_energia_tf['custo_sem_iva']
        tt_cte_energia_iva_6_tf = decomposicao_custo_energia_tf['valor_iva_6']
        tt_cte_energia_iva_23_tf = decomposicao_custo_energia_tf['valor_iva_23']

        # POTÊNCIA (Tarifários Fixos)
        # preco_comercializador_potencia_final_sem_iva_tf e tar_potencia_final_dia_sem_iva_tf já incluem TS (se aplicável)
        decomposicao_custo_potencia_tf = calcular_custo_potencia_com_iva_final(
            preco_comercializador_potencia_final_sem_iva_tf, # Componente comercializador s/IVA, após TS (se TS afetasse isso)
            tar_potencia_final_dia_sem_iva_tf,              # Componente TAR s/IVA, após TS
            dias,
            potencia
        )
        custo_potencia_tf_com_iva = decomposicao_custo_potencia_tf['custo_com_iva']
        tt_cte_potencia_siva_tf = decomposicao_custo_potencia_tf['custo_sem_iva']
        tt_cte_potencia_iva_6_tf = decomposicao_custo_potencia_tf['valor_iva_6']
        tt_cte_potencia_iva_23_tf = decomposicao_custo_potencia_tf['valor_iva_23']
        
        # TAXAS ADICIONAIS (Tarifários Fixos)
        decomposicao_taxas_tf = calcular_taxas_adicionais(
            consumo, dias, tarifa_social,
            valor_dgeg_user, valor_cav_user,
            nome_comercializador_atual=comercializador_tarifario_tf,
            mes_selecionado_simulacao=mes, ano_simulacao_atual=ano_atual
        )
        taxas_tf_com_iva = decomposicao_taxas_tf['custo_com_iva']
        tt_cte_iec_siva_tf = decomposicao_taxas_tf['iec_sem_iva']
        tt_cte_dgeg_siva_tf = decomposicao_taxas_tf['dgeg_sem_iva']
        tt_cte_cav_siva_tf = decomposicao_taxas_tf['cav_sem_iva']
        tt_cte_taxas_iva_6_tf = decomposicao_taxas_tf['valor_iva_6']
        tt_cte_taxas_iva_23_tf = decomposicao_taxas_tf['valor_iva_23']

        # Custo Total antes de outros descontos específicos do tarifário fixo
        custo_total_antes_desc_especificos_tf = custo_energia_tf_com_iva + custo_potencia_tf_com_iva + taxas_tf_com_iva
        
        # Calcular totais para o tooltip do Custo Total Estimado
        tt_cte_total_siva_tf = tt_cte_energia_siva_tf + tt_cte_potencia_siva_tf + tt_cte_iec_siva_tf + tt_cte_dgeg_siva_tf + tt_cte_cav_siva_tf
        tt_cte_valor_iva_6_total_tf = tt_cte_energia_iva_6_tf + tt_cte_potencia_iva_6_tf + tt_cte_taxas_iva_6_tf
        tt_cte_valor_iva_23_total_tf = tt_cte_energia_iva_23_tf + tt_cte_potencia_iva_23_tf + tt_cte_taxas_iva_23_tf

        # NOVO: Calcular Subtotal c/IVA (antes de descontos/acréscimos finais)
        tt_cte_subtotal_civa_tf = tt_cte_total_siva_tf + tt_cte_valor_iva_6_total_tf + tt_cte_valor_iva_23_total_tf
        
        tt_cte_desc_finais_valor_tf = 0.0
        if desconto_fatura_periodo_tf > 0: # Usa o valor proporcionalizado ou fixo já calculado
            tt_cte_desc_finais_valor_tf += desconto_fatura_periodo_tf
        if 'desconto_meo_aplicado_periodo' in locals() and desconto_meo_aplicado_periodo > 0:
            tt_cte_desc_finais_valor_tf += desconto_meo_aplicado_periodo
        if 'valor_X_desconto_continente' in locals() and valor_X_desconto_continente > 0:
            tt_cte_desc_finais_valor_tf += valor_X_desconto_continente
            
        tt_cte_acres_finais_valor_tf = 0.0
        if 'incluir_quota_acp' in locals() and incluir_quota_acp and 'quota_acp_periodo' in locals() and quota_acp_periodo > 0:
            tt_cte_acres_finais_valor_tf += quota_acp_periodo

        # Adicionar os novos campos de tooltip ao resultado_fixo
        componentes_tooltip_custo_total_dict_tf = {
            'tt_cte_energia_siva': tt_cte_energia_siva_tf,
            'tt_cte_potencia_siva': tt_cte_potencia_siva_tf,
            'tt_cte_iec_siva': tt_cte_iec_siva_tf,
            'tt_cte_dgeg_siva': tt_cte_dgeg_siva_tf,
            'tt_cte_cav_siva': tt_cte_cav_siva_tf,
            'tt_cte_total_siva': tt_cte_total_siva_tf,
            'tt_cte_valor_iva_6_total': tt_cte_valor_iva_6_total_tf,
            'tt_cte_valor_iva_23_total': tt_cte_valor_iva_23_total_tf,
            'tt_cte_subtotal_civa': tt_cte_subtotal_civa_tf,
            'tt_cte_desc_finais_valor': tt_cte_desc_finais_valor_tf,
            'tt_cte_acres_finais_valor': tt_cte_acres_finais_valor_tf
        }

        # Preparar o dicionário de resultado
        resultado_fixo = {
            'NomeParaExibir': nome_a_exibir,
            'LinkAdesao': link_adesao_tf,
            'info_notas': notas_tarifario_tf,
            'Tipo': tipo_tarifario,
            'Segmento': segmento_tarifario,
            'Faturação': faturacao_tarifario,
            'Pagamento': pagamento_tarifario,
            'Comercializador': comercializador_tarifario,
            **valores_energia_exibir_tf,
            'Potência (€/dia)': round(preco_potencia_total_final_sem_iva_tf, 4),
            'Total (€)': round(custo_total_estimado_final_tf, 2),
            # CAMPOS DO TOOLTIP DA POTÊNCIA FIXOS
            **componentes_tooltip_potencia_dict_tf,
            # CAMPOS DO TOOLTIP DA ENERGIA FIXOS
            **componentes_tooltip_energia_dict_tf, 
            # CAMPOS DO TOOLTIP DA CUSTO TOTAL FIXOS
            **componentes_tooltip_custo_total_dict_tf, 
            }
        resultados_list.append(resultado_fixo)

# --- Fim do loop for tarifario_fixo ---

# --- Comparar Tarifários Indexados (se a checkbox estiver ativa) ---

if df_omie_ajustado.empty:
    st.warning("Não existem dados OMIE para o período selecionado. Tarifários indexados não podem ser calculados.")
else:
    tarifarios_filtrados_indexados = ti_processar[
        (ti_processar['opcao_horaria_e_ciclo'] == opcao_horaria) &
        (ti_processar['potencia_kva'] == potencia)
    ].copy()

    # Usar a estrutura simplificada do if/else
    if not tarifarios_filtrados_indexados.empty:
        for index, tarifario_indexado in tarifarios_filtrados_indexados.iterrows():
            # --- Get tariff specifics ---
            nome_tarifario = tarifario_indexado['nome']
            tipo_tarifario = tarifario_indexado['tipo']
            comercializador_tarifario = tarifario_indexado['comercializador']
            link_adesao_idx = tarifario_indexado.get('site_adesao')
            notas_tarifario_idx = tarifario_indexado.get('notas', '') 
            segmento_tarifario = tarifario_indexado.get('segmento', '-')
            faturacao_tarifario = tarifario_indexado.get('faturacao', '-')
            pagamento_tarifario = tarifario_indexado.get('pagamento', '-')
            formula_energia = str(tarifario_indexado.get('formula_calculo', '')) # Garantir string
            preco_potencia_dia = tarifario_indexado['preco_potencia_dia']

            constantes = dict(zip(CONSTANTES["constante"], CONSTANTES["valor_unitário"]))

            # Inicializar variáveis de preço
            preco_energia_simples_indexado = None
            preco_energia_vazio_indexado = None
            preco_energia_fora_vazio_indexado = None
            preco_energia_cheias_indexado = None
            preco_energia_ponta_indexado = None

            # --- CALCULAR PREÇO BASE INDEXADO (input energia) ---
                # --- BLOCO 1: Cálculo para Indexados Quarto-Horários (BTN ou Luzboa "BTN SPOTDEF") ---
                # Assume que 'BTN' em formula_energia ou o nome Luzboa identifica corretamente estes tarifários
            if 'BTN' in formula_energia or nome_tarifario == "Luzboa - BTN SPOTDEF":

                # --- Tratamento especial para Luzboa - BTN SPOTDEF ---
                if nome_tarifario == "Luzboa - BTN SPOTDEF":
                    # [LÓGICA LUZBOA - Mantida como estava na versão anterior que funcionava]
                    soma_luzboa_simples, count_luzboa_simples = 0.0, 0
                    soma_luzboa_vazio, count_luzboa_vazio = 0.0, 0
                    soma_luzboa_fv, count_luzboa_fv = 0.0, 0
                    soma_luzboa_cheias, count_luzboa_cheias = 0.0, 0
                    soma_luzboa_ponta, count_luzboa_ponta = 0.0, 0

                    coluna_ciclo_luzboa = None
                    if opcao_horaria.lower().startswith("bi"):
                        coluna_ciclo_luzboa = 'BD' if "Diário" in opcao_horaria else 'BS'
                    elif opcao_horaria.lower().startswith("tri"):
                        coluna_ciclo_luzboa = 'TD' if "Diário" in opcao_horaria else 'TS'

                    if coluna_ciclo_luzboa and coluna_ciclo_luzboa not in df_omie_ajustado.columns and not opcao_horaria.lower() == "simples":
                         st.warning(f"Coluna de ciclo '{coluna_ciclo_luzboa}' não encontrada para Luzboa. Energia será zero.")
                         if opcao_horaria.lower() == "simples": preco_energia_simples_indexado = 0.0
                         else: preco_energia_vazio_indexado, preco_energia_fora_vazio_indexado, preco_energia_cheias_indexado, preco_energia_ponta_indexado = 0.0, 0.0, 0.0, 0.0
                    else: # Calcular apenas se coluna de ciclo existe (ou se for simples)
                        for _, row_omie in df_omie_ajustado.iterrows():
                            if not all(k in row_omie and pd.notna(row_omie[k]) for k in ['OMIE', 'Perdas']): continue
                            omie_val = row_omie['OMIE'] / 1000; perdas_val = row_omie['Perdas']
                            cgs_luzboa = constantes.get('Luzboa_CGS', 0.0); fa_luzboa = constantes.get('Luzboa_FA', 1.0); kp_luzboa = constantes.get('Luzboa_Kp', 0.0)
                            valor_hora_luzboa = (omie_val + cgs_luzboa) * perdas_val * fa_luzboa + kp_luzboa
                            if opcao_horaria.lower() == "simples": soma_luzboa_simples += valor_hora_luzboa; count_luzboa_simples += 1
                            elif coluna_ciclo_luzboa and coluna_ciclo_luzboa in row_omie and pd.notna(row_omie[coluna_ciclo_luzboa]):
                                 ciclo_hora = row_omie[coluna_ciclo_luzboa]
                                 if opcao_horaria.lower().startswith("bi"):
                                     if ciclo_hora == 'V': soma_luzboa_vazio += valor_hora_luzboa; count_luzboa_vazio += 1
                                     elif ciclo_hora == 'F': soma_luzboa_fv += valor_hora_luzboa; count_luzboa_fv += 1
                                 elif opcao_horaria.lower().startswith("tri"):
                                     if ciclo_hora == 'V': soma_luzboa_vazio += valor_hora_luzboa; count_luzboa_vazio += 1
                                     elif ciclo_hora == 'C': soma_luzboa_cheias += valor_hora_luzboa; count_luzboa_cheias += 1
                                     elif ciclo_hora == 'P': soma_luzboa_ponta += valor_hora_luzboa; count_luzboa_ponta += 1
                        prec = 4
                        if opcao_horaria.lower() == "simples": preco_energia_simples_indexado = round(soma_luzboa_simples / count_luzboa_simples, 4) if count_luzboa_simples > 0 else 0.0
                        elif opcao_horaria.lower().startswith("bi"):
                            preco_energia_vazio_indexado = round(soma_luzboa_vazio / count_luzboa_vazio, prec) if count_luzboa_vazio > 0 else 0.0
                            preco_energia_fora_vazio_indexado = round(soma_luzboa_fv / count_luzboa_fv, prec) if count_luzboa_fv > 0 else 0.0
                        elif opcao_horaria.lower().startswith("tri"):
                            preco_energia_vazio_indexado = round(soma_luzboa_vazio / count_luzboa_vazio, prec) if count_luzboa_vazio > 0 else 0.0
                            preco_energia_cheias_indexado = round(soma_luzboa_cheias / count_luzboa_cheias, prec) if count_luzboa_cheias > 0 else 0.0
                            preco_energia_ponta_indexado = round(soma_luzboa_ponta / count_luzboa_ponta, prec) if count_luzboa_ponta > 0 else 0.0
                    # --- FIM LÓGICA LUZBOA ---

                else: # Outros Tarifários Quarto-Horários (Coopernico, Repsol, Galp, etc.)
                    # [LÓGICA PARA OUTROS BTN COM PERFIL - INCLUI AJUSTE REPSOL]
                    perfil_coluna = f"BTN_{obter_perfil(consumo, dias, potencia).split('_')[1].upper()}"
                    # Verifica se coluna de perfil existe
                    if perfil_coluna not in df_omie_ajustado.columns:
                        st.warning(f"Coluna de perfil '{perfil_coluna}' não encontrada para '{nome_tarifario}'. Energia será zero.")
                        if opcao_horaria.lower() == "simples": preco_energia_simples_indexado = 0.0
                        else: preco_energia_vazio_indexado, preco_energia_fora_vazio_indexado, preco_energia_cheias_indexado, preco_energia_ponta_indexado = 0.0, 0.0, 0.0, 0.0
                    else: # Coluna de perfil existe, prosseguir com cálculos
                        soma_calculo_simples, soma_perfil_simples = 0.0, 0.0; soma_calculo_vazio, soma_perfil_vazio = 0.0, 0.0; soma_calculo_fv, soma_perfil_fv = 0.0, 0.0; soma_calculo_cheias, soma_perfil_cheias = 0.0, 0.0; soma_calculo_ponta, soma_perfil_ponta = 0.0, 0.0
                        coluna_ciclo = None
                        cycle_column_ok = True # Assumir que está OK por defeito

                        if not opcao_horaria.lower() == "simples":
                            if opcao_horaria.lower().startswith("bi"): coluna_ciclo = 'BD' if "Diário" in opcao_horaria else 'BS'
                            elif opcao_horaria.lower().startswith("tri"): coluna_ciclo = 'TD' if "Diário" in opcao_horaria else 'TS'
                            
                            if coluna_ciclo and coluna_ciclo not in df_omie_ajustado.columns:
                                st.warning(f"Coluna de ciclo '{coluna_ciclo}' não encontrada para '{nome_tarifario}' com '{opcao_horaria}'. Preços específicos V/F/C/P podem ser zero.")
                                cycle_column_ok = False
                                # Definir preços específicos a zero, mas o simples ainda pode ser calculado
                                preco_energia_vazio_indexado, preco_energia_fora_vazio_indexado, preco_energia_cheias_indexado, preco_energia_ponta_indexado = 0.0, 0.0, 0.0, 0.0

                        # Loop sobre os dados OMIE do período já filtrado (df_omie_ajustado)
                        for _, row_omie in df_omie_ajustado.iterrows():
                            required_cols_check = ['OMIE', 'Perdas', perfil_coluna]
                            if not all(k in row_omie and pd.notna(row_omie[k]) for k in required_cols_check): continue
                            omie = row_omie['OMIE'] / 1000; perdas = row_omie['Perdas']; perfil = row_omie[perfil_coluna]
                            if perfil <= 0: continue

                            calculo_instantaneo_sem_perfil = 0.0
                            # --- Fórmulas específicas BTN ---
                            if nome_tarifario == "Coopérnico Base 2.0": calculo_instantaneo_sem_perfil = (omie + constantes.get('Coop_CS_CR', 0.0) + constantes.get('Coop_K', 0.0)) * perdas
                            elif nome_tarifario == "Repsol - Leve Sem Mais": calculo_instantaneo_sem_perfil = (omie * perdas * constantes.get('Repsol_FA', 0.0) + constantes.get('Repsol_Q_Tarifa', 0.0))
                            elif nome_tarifario == "Repsol - Leve PRO Sem Mais": calculo_instantaneo_sem_perfil = (omie * perdas * constantes.get('Repsol_FA', 0.0) + constantes.get('Repsol_Q_Tarifa_Pro', 0.0))
                            elif nome_tarifario == "Galp - Plano Flexível / Dinâmico": calculo_instantaneo_sem_perfil = (omie + constantes.get('Galp_Ci', 0.0)) * perdas
                            elif nome_tarifario == "Alfa Energia - ALFA POWER INDEX BTN": calculo_instantaneo_sem_perfil = ((omie + constantes.get('Alfa_CGS', 0.0)) * perdas + constantes.get('Alfa_K', 0.0))
                            elif nome_tarifario == "Plenitude - Tendência": calculo_instantaneo_sem_perfil = (((omie + constantes.get('Plenitude_CGS', 0.0) + constantes.get('Plenitude_GDOs', 0.0))) * perdas + constantes.get('Plenitude_Fee', 0.0))
                            elif nome_tarifario == "Meo Energia - Tarifa Variável": calculo_instantaneo_sem_perfil = (omie + constantes.get('Meo_K', 0.0)) * perdas
                            elif nome_tarifario == "EDP - Eletricidade Indexada Horária": calculo_instantaneo_sem_perfil = (omie * perdas * constantes.get('EDP_H_K1', 1.0) + constantes.get('EDP_H_K2', 0.0))
                            elif nome_tarifario == "EZU - Coletiva": calculo_instantaneo_sem_perfil = (omie + constantes.get('EZU_K', 0.0) + constantes.get('EZU_CGS', 0.0)) * perdas
                            elif nome_tarifario == "G9 - Smart Dynamic": calculo_instantaneo_sem_perfil = (omie * constantes.get('G9_FA', 0.0) * perdas + constantes.get('G9_CGS', 0.0) + constantes.get('G9_AC', 0.0))
                            elif nome_tarifario == "Iberdrola - Simples Indexado Dinâmico": calculo_instantaneo_sem_perfil = (omie * perdas + constantes.get("Iberdrola_Q", 0.0) + constantes.get('Iberdrola_mFRR', 0.0))


                            else: calculo_instantaneo_sem_perfil = omie * perdas # Fallback genérico
                            # --- Fim Fórmulas ---

                            # --- Acumular Somas ---
                            # Acumula SEMPRE nas somas simples (gerais ponderadas pelo perfil)
                            soma_calculo_simples += calculo_instantaneo_sem_perfil * perfil
                            soma_perfil_simples += perfil

                            # Acumula nas somas específicas do período SE aplicável e coluna de ciclo OK
                            if cycle_column_ok and coluna_ciclo and coluna_ciclo in row_omie and pd.notna(row_omie[coluna_ciclo]):
                                ciclo_hora = row_omie[coluna_ciclo]
                                if opcao_horaria.lower().startswith("bi"):
                                    if ciclo_hora == 'V': soma_calculo_vazio += calculo_instantaneo_sem_perfil * perfil; soma_perfil_vazio += perfil
                                    elif ciclo_hora == 'F': soma_calculo_fv += calculo_instantaneo_sem_perfil * perfil; soma_perfil_fv += perfil
                                elif opcao_horaria.lower().startswith("tri"):
                                    if ciclo_hora == 'V': soma_calculo_vazio += calculo_instantaneo_sem_perfil * perfil; soma_perfil_vazio += perfil
                                    elif ciclo_hora == 'C': soma_calculo_cheias += calculo_instantaneo_sem_perfil * perfil; soma_perfil_cheias += perfil
                                    elif ciclo_hora == 'P': soma_calculo_ponta += calculo_instantaneo_sem_perfil * perfil; soma_perfil_ponta += perfil
                        # --- Fim loop horas ---
                            
                        prec = 4
                        # --- Cálculo de preços FINAIS para BTN ---
                        if nome_tarifario == "Repsol - Leve Sem Mais":
                            # Repsol usa sempre o preço calculado como se fosse Simples
                            preco_simples_repsol = round(soma_calculo_simples / soma_perfil_simples, prec) if soma_perfil_simples > 0 else 0.0
                            preco_energia_simples_indexado = preco_simples_repsol
                            preco_energia_vazio_indexado = preco_simples_repsol
                            preco_energia_fora_vazio_indexado = preco_simples_repsol
                            preco_energia_cheias_indexado = preco_simples_repsol
                            preco_energia_ponta_indexado = preco_simples_repsol
                        elif nome_tarifario == "Repsol - Leve PRO Sem Mais":
                            # Repsol usa sempre o preço calculado como se fosse Simples
                            preco_simples_repsol_pro = round(soma_calculo_simples / soma_perfil_simples, prec) if soma_perfil_simples > 0 else 0.0
                            preco_energia_simples_indexado = preco_simples_repsol_pro
                            preco_energia_vazio_indexado = preco_simples_repsol_pro
                            preco_energia_fora_vazio_indexado = preco_simples_repsol_pro
                            preco_energia_cheias_indexado = preco_simples_repsol_pro
                            preco_energia_ponta_indexado = preco_simples_repsol_pro                            
                        else:
                            # Cálculo normal para os outros BTN
                            if opcao_horaria.lower() == "simples":
                                preco_energia_simples_indexado = round(soma_calculo_simples / soma_perfil_simples, prec) if soma_perfil_simples > 0 else 0.0
                            elif opcao_horaria.lower().startswith("bi"):
                                preco_energia_vazio_indexado = round(soma_calculo_vazio / soma_perfil_vazio, prec) if soma_perfil_vazio > 0 else 0.0
                                preco_energia_fora_vazio_indexado = round(soma_calculo_fv / soma_perfil_fv, prec) if soma_perfil_fv > 0 else 0.0
                            elif opcao_horaria.lower().startswith("tri"):
                                preco_energia_vazio_indexado = round(soma_calculo_vazio / soma_perfil_vazio, prec) if soma_perfil_vazio > 0 else 0.0
                                preco_energia_cheias_indexado = round(soma_calculo_cheias / soma_perfil_cheias, prec) if soma_perfil_cheias > 0 else 0.0
                                preco_energia_ponta_indexado = round(soma_calculo_ponta / soma_perfil_ponta, prec) if soma_perfil_ponta > 0 else 0.0
                # --- FIM LÓGICA OUTROS BTN ---

            # --- BLOCO 2: Cálculo para Indexados Média ---
            else: # Se não for Quarto-Horário (BTN ou Luzboa)
                # --- INÍCIO LÓGICA MÉDIA CORRIGIDA ---
                omie_medio_simples_input_kwh = None; omie_medio_vazio_kwh = None; omie_medio_fv_kwh = None; omie_medio_cheias_kwh = None; omie_medio_ponta_kwh = None
                if opcao_horaria.lower() == "simples": omie_medio_simples_input_kwh = omie_para_tarifarios_media.get('S', 0.0) / 1000.0
                elif opcao_horaria.lower().startswith("bi"): omie_medio_vazio_kwh = omie_para_tarifarios_media.get('V', 0.0) / 1000.0; omie_medio_fv_kwh = omie_para_tarifarios_media.get('F', 0.0) / 1000.0
                elif opcao_horaria.lower().startswith("tri"): omie_medio_vazio_kwh = omie_para_tarifarios_media.get('V', 0.0) / 1000.0; omie_medio_cheias_kwh = omie_para_tarifarios_media.get('C', 0.0) / 1000.0; omie_medio_ponta_kwh = omie_para_tarifarios_media.get('P', 0.0) / 1000.0
                prec = 4

                if opcao_horaria.lower() == "simples":
                    perdas_a_usar = perdas_medias.get('Perdas_Anual_S', 1.0) # Usa Anual Simples
                    omie_a_usar = omie_medio_simples_input_kwh if omie_medio_simples_input_kwh is not None else 0.0
                    if nome_tarifario == "Iberdrola - Simples Indexado": preco_energia_simples_indexado = round(omie_a_usar * constantes.get('Iberdrola_Perdas', 1.0) + constantes.get("Iberdrola_Q", 0.0) + constantes.get('Iberdrola_mFRR', 0.0), prec)
                    elif nome_tarifario == "Goldenergy - Tarifário Indexado 100%":
                        mes_num_calculo = list(dias_mes.keys()).index(mes) + 1; perdas_mensais_ge_map = {1: 1.29, 2: 1.18, 3: 1.18, 4: 1.15, 5: 1.11, 6: 1.10, 7: 1.15, 8: 1.13, 9: 1.10, 10: 1.10, 11: 1.16, 12: 1.25}; perdas_mensais_ge = perdas_mensais_ge_map.get(mes_num_calculo, 1.0)
                        preco_energia_simples_indexado = round(omie_a_usar * perdas_mensais_ge + constantes.get('GE_Q_Tarifa', 0.0) + constantes.get('GE_CG', 0.0), prec)
                    elif nome_tarifario == "Endesa - Tarifa Indexada": preco_energia_simples_indexado = round(omie_a_usar + constantes.get('Endesa_A_S', 0.0), prec)
                    elif nome_tarifario == "LUZiGÁS - Energy 8.8": preco_energia_simples_indexado = round((omie_a_usar + constantes.get('Luzigas_8_8_K', 0.0) + constantes.get('Luzigas_CGS', 0.0)) * perdas_a_usar, prec)
                    elif nome_tarifario == "LUZiGÁS - Dinâmico Poupança +": preco_energia_simples_indexado = round((omie_a_usar + constantes.get('Luzigas_D_K', 0.0) + constantes.get('Luzigas_CGS', 0.0)) * perdas_a_usar, prec)
                    elif nome_tarifario == "Ibelectra - Solução Família": preco_energia_simples_indexado = round((omie_a_usar + constantes.get('Ibelectra_CS', 0.0)) * perdas_a_usar + constantes.get('Ibelectra_K', 0.0), prec)
                    elif nome_tarifario == "G9 - Smart Index": preco_energia_simples_indexado = round((omie_a_usar * constantes.get('G9_FA', 1.02)) * perdas_medias.get('Perdas_M_S', 1.16) + constantes.get('G9_CGS', 0.01) + constantes.get('G9_AC', 0.0055), prec)
                    elif nome_tarifario == "EDP - Eletricidade Indexada Média": preco_energia_simples_indexado = round(omie_a_usar * constantes.get('EDP_M_Perdas', 1.0) * constantes.get('EDP_M_K1', 1.0) + constantes.get('EDP_M_K2', 0.0), prec)
                    else: st.warning(f"Fórmula não definida para tarifário médio Simples: {nome_tarifario}"); preco_energia_simples_indexado = omie_a_usar
                elif opcao_horaria.lower().startswith("bi"):
                    ciclo_bi = 'BD' if "Diário" in opcao_horaria else 'BS'
                    perdas_v_anual = perdas_medias.get(f'Perdas_Anual_{ciclo_bi}_V', 1.0); perdas_f_anual = perdas_medias.get(f'Perdas_Anual_{ciclo_bi}_F', 1.0)
                    omie_v_a_usar = omie_medio_vazio_kwh if omie_medio_vazio_kwh is not None else 0.0; omie_f_a_usar = omie_medio_fv_kwh if omie_medio_fv_kwh is not None else 0.0
                    if nome_tarifario == "LUZiGÁS - Energy 8.8": k_luzigas = constantes.get('Luzigas_8_8_K', 0.0); cgs_luzigas = constantes.get('Luzigas_CGS', 0.0); calc_base = omie_medio_simples_real_kwh + k_luzigas + cgs_luzigas; preco_energia_vazio_indexado = round(calc_base * perdas_v_anual, prec); preco_energia_fora_vazio_indexado = round(calc_base * perdas_f_anual, prec)
                    elif nome_tarifario == "LUZiGÁS - Dinâmico Poupança +": k_luzigas = constantes.get('Luzigas_D_K', 0.0); cgs_luzigas = constantes.get('Luzigas_CGS', 0.0); calc_base = omie_medio_simples_real_kwh + k_luzigas + cgs_luzigas; preco_energia_vazio_indexado = round(calc_base * perdas_v_anual, prec); preco_energia_fora_vazio_indexado = round(calc_base * perdas_f_anual, prec)
                    elif nome_tarifario == "Endesa - Tarifa Indexada": preco_energia_vazio_indexado = round(omie_v_a_usar + constantes.get('Endesa_A_V', 0.0), prec); preco_energia_fora_vazio_indexado = round(omie_f_a_usar + constantes.get('Endesa_A_FV', 0.0), prec)
                    elif nome_tarifario == "Ibelectra - Solução Família": cs_ib = constantes.get('Ibelectra_CS', 0.0); k_ib = constantes.get('Ibelectra_K', 0.0); preco_energia_vazio_indexado = round((omie_v_a_usar + cs_ib) * perdas_v_anual + k_ib, prec); preco_energia_fora_vazio_indexado = round((omie_f_a_usar + cs_ib) * perdas_f_anual + k_ib, prec)                    
                    elif nome_tarifario == "G9 - Smart Index": preco_energia_vazio_indexado = round((omie_v_a_usar * constantes.get('G9_FA', 1.02) * perdas_medias.get(f'Perdas_M_{ciclo_bi}_V', 1.16)) + constantes.get('G9_CGS', 0.01) + constantes.get('G9_AC', 0.0055), prec); preco_energia_fora_vazio_indexado = round((omie_f_a_usar * constantes.get('G9_FA', 1.02) * perdas_medias.get(f'Perdas_M_{ciclo_bi}_F', 1.16)) + constantes.get('G9_CGS', 0.01) + constantes.get('G9_AC', 0.0055), prec)                    
                    elif nome_tarifario == "EDP - Eletricidade Indexada Média": perdas_const_edp = constantes.get('EDP_M_Perdas', 1.0); k1_edp = constantes.get('EDP_M_K1', 1.0); k2_edp = constantes.get('EDP_M_K2', 0.0); preco_energia_vazio_indexado = round(omie_v_a_usar * perdas_const_edp * k1_edp + k2_edp, prec); preco_energia_fora_vazio_indexado = round(omie_f_a_usar * perdas_const_edp * k1_edp + k2_edp, prec)
                    else: st.warning(f"Fórmula não definida para tarifário médio Bi-horário: {nome_tarifario}"); preco_energia_vazio_indexado = omie_v_a_usar; preco_energia_fora_vazio_indexado = omie_f_a_usar
                elif opcao_horaria.lower().startswith("tri"):
                    ciclo_tri = 'TD' if "Diário" in opcao_horaria else 'TS'; perdas_v_anual = perdas_medias.get(f'Perdas_Anual_{ciclo_tri}_V', 1.0); perdas_c_anual = perdas_medias.get(f'Perdas_Anual_{ciclo_tri}_C', 1.0); perdas_p_anual = perdas_medias.get(f'Perdas_Anual_{ciclo_tri}_P', 1.0)
                    omie_v_a_usar = omie_medio_vazio_kwh if omie_medio_vazio_kwh is not None else 0.0; omie_c_a_usar = omie_medio_cheias_kwh if omie_medio_cheias_kwh is not None else 0.0; omie_p_a_usar = omie_medio_ponta_kwh if omie_medio_ponta_kwh is not None else 0.0
                    if nome_tarifario == "LUZiGÁS - Energy 8.8": k_luzigas = constantes.get('Luzigas_8_8_K', 0.0); cgs_luzigas = constantes.get('Luzigas_CGS', 0.0); calc_base = omie_medio_simples_real_kwh + k_luzigas + cgs_luzigas; preco_energia_vazio_indexado = round(calc_base * perdas_v_anual, prec); preco_energia_cheias_indexado = round(calc_base * perdas_c_anual, prec); preco_energia_ponta_indexado = round(calc_base * perdas_p_anual, prec)
                    elif nome_tarifario == "LUZiGÁS - Dinâmico Poupança +": k_luzigas = constantes.get('Luzigas_D_K', 0.0); cgs_luzigas = constantes.get('Luzigas_CGS', 0.0); calc_base = omie_medio_simples_real_kwh + k_luzigas + cgs_luzigas; preco_energia_vazio_indexado = round(calc_base * perdas_v_anual, prec); preco_energia_cheias_indexado = round(calc_base * perdas_c_anual, prec); preco_energia_ponta_indexado = round(calc_base * perdas_p_anual, prec)
                    elif nome_tarifario == "Ibelectra - Solução Família": cs_ib = constantes.get('Ibelectra_CS', 0.0); k_ib = constantes.get('Ibelectra_K', 0.0); preco_energia_vazio_indexado = round((omie_v_a_usar + cs_ib) * perdas_v_anual + k_ib, prec); preco_energia_cheias_indexado = round((omie_c_a_usar + cs_ib) * perdas_c_anual + k_ib, prec); preco_energia_ponta_indexado = round((omie_p_a_usar + cs_ib) * perdas_p_anual + k_ib, prec)
                    elif nome_tarifario == "G9 - Smart Index": preco_energia_vazio_indexado = round((omie_v_a_usar * constantes.get('G9_FA', 1.02) * perdas_medias.get(f'Perdas_M_{ciclo_tri}_V', 1.16)) + constantes.get('G9_CGS', 0.01) + constantes.get('G9_AC', 0.0055), prec); preco_energia_cheias_indexado = round((omie_c_a_usar * constantes.get('G9_FA', 1.02) * perdas_medias.get(f'Perdas_M_{ciclo_tri}_C', 1.16)) + constantes.get('G9_CGS', 0.01) + constantes.get('G9_AC', 0.0055), prec); preco_energia_ponta_indexado = round((omie_p_a_usar * constantes.get('G9_FA', 1.02) * perdas_medias.get(f'Perdas_M_{ciclo_tri}_P', 1.16)) + constantes.get('G9_CGS', 0.01) + constantes.get('G9_AC', 0.0055), prec) 
                    elif nome_tarifario == "EDP - Eletricidade Indexada Média": perdas_const_edp = constantes.get('EDP_M_Perdas', 1.0); k1_edp = constantes.get('EDP_M_K1', 1.0); k2_edp = constantes.get('EDP_M_K2', 0.0); preco_energia_vazio_indexado = round(omie_v_a_usar * perdas_const_edp * k1_edp + k2_edp, prec); preco_energia_cheias_indexado = round(omie_c_a_usar * perdas_const_edp * k1_edp + k2_edp, prec); preco_energia_ponta_indexado = round(omie_p_a_usar * perdas_const_edp * k1_edp + k2_edp, prec)
                    else: st.warning(f"Fórmula não definida para tarifário médio Tri-horário: {nome_tarifario}"); preco_energia_vazio_indexado = omie_v_a_usar; preco_energia_cheias_indexado = omie_c_a_usar; preco_energia_ponta_indexado = omie_p_a_usar
                # --- FIM LÓGICA MÉDIA ---

            # --- Fim do bloco de cálculo base indexado ---

            # Criar dict de input
            preco_energia_input_idx = {}
            consumos_horarios_para_func_idx = {} # Preencher aqui
            if opcao_horaria.lower() == "simples":
                preco_energia_input_idx['S'] = preco_energia_simples_indexado if 'preco_energia_simples_indexado' in locals() and preco_energia_simples_indexado is not None else 0.0
                consumos_horarios_para_func_idx = {'S': consumo_simples}
            elif opcao_horaria.lower().startswith("bi"):
                preco_energia_input_idx['V'] = preco_energia_vazio_indexado if 'preco_energia_vazio_indexado' in locals() and preco_energia_vazio_indexado is not None else 0.0
                preco_energia_input_idx['F'] = preco_energia_fora_vazio_indexado if 'preco_energia_fora_vazio_indexado' in locals() and preco_energia_fora_vazio_indexado is not None else 0.0
                consumos_horarios_para_func_idx = {'V': consumo_vazio, 'F': consumo_fora_vazio}
            elif opcao_horaria.lower().startswith("tri"):
                preco_energia_input_idx['V'] = preco_energia_vazio_indexado if 'preco_energia_vazio_indexado' in locals() and preco_energia_vazio_indexado is not None else 0.0
                preco_energia_input_idx['C'] = preco_energia_cheias_indexado if 'preco_energia_cheias_indexado' in locals() and preco_energia_cheias_indexado is not None else 0.0
                preco_energia_input_idx['P'] = preco_energia_ponta_indexado if 'preco_energia_ponta_indexado' in locals() and preco_energia_ponta_indexado is not None else 0.0
                consumos_horarios_para_func_idx = {'V': consumo_vazio, 'C': consumo_cheias, 'P': consumo_ponta}


            preco_potencia_input_idx = tarifario_indexado.get('preco_potencia_dia', 0.0)

            # Flags (verificar defaults adequados para indexados)
            tar_incluida_energia_idx = tarifario_indexado.get('tar_incluida_energia', False)
            tar_incluida_potencia_idx = tarifario_indexado.get('tar_incluida_potencia', True)
            financiamento_tse_incluido_idx = tarifario_indexado.get('financiamento_tse_incluido', False)


            # --- Passo 1: Identificar Componentes Base (Sem IVA, Sem TS) ---
            tar_energia_regulada_idx = {}
            for periodo in preco_energia_input_idx.keys():
                tar_energia_regulada_idx[periodo] = obter_tar_energia_periodo(opcao_horaria, periodo, potencia, CONSTANTES)

            tar_potencia_regulada_idx = obter_tar_dia(potencia, CONSTANTES)

            preco_comercializador_energia_idx = {}
            for periodo, preco_in in preco_energia_input_idx.items():
                 preco_in_float = float(preco_in or 0.0) # Ensure float
                 if tar_incluida_energia_idx:
                     preco_comercializador_energia_idx[periodo] = preco_in_float - tar_energia_regulada_idx.get(periodo, 0.0)
                 else:
                     preco_comercializador_energia_idx[periodo] = preco_in_float
                 # Não aplicar max(0,...) aqui, fórmulas de indexados podem dar negativo

            preco_potencia_input_idx_float = float(preco_potencia_input_idx or 0.0) # Ensure float
            if tar_incluida_potencia_idx:
                preco_comercializador_potencia_idx = preco_potencia_input_idx_float - tar_potencia_regulada_idx
            else:
                preco_comercializador_potencia_idx = preco_potencia_input_idx_float
            # Não aplicar max(0,...) aqui

            financiamento_tse_a_adicionar_idx = FINANCIAMENTO_TSE_VAL if not financiamento_tse_incluido_idx else 0.0

            # --- Passo 2: Calcular Componentes TAR Finais (Com Desconto TS, Sem IVA) ---
            tar_energia_final_idx = {}
            tar_potencia_final_dia_idx = tar_potencia_regulada_idx

            if tarifa_social: # Flag global
                desconto_ts_energia = obter_constante('Desconto TS Energia', CONSTANTES)
                desconto_ts_potencia_dia = obter_constante(f'Desconto TS Potencia {potencia}', CONSTANTES)
                for periodo, tar_reg in tar_energia_regulada_idx.items():
                    tar_energia_final_idx[periodo] = tar_reg - desconto_ts_energia
                tar_potencia_final_dia_idx = max(0.0, tar_potencia_regulada_idx - desconto_ts_potencia_dia)
            else:
                tar_energia_final_idx = tar_energia_regulada_idx.copy()

            desconto_ts_potencia_valor_aplicado = 0.0
            if tarifa_social: # Flag global
                desconto_ts_potencia_dia_bruto = obter_constante(f'Desconto TS Potencia {potencia}', CONSTANTES)
                # O desconto efetivamente aplicado é o mínimo entre o desconto e a própria TAR
                desconto_ts_potencia_valor_aplicado = min(tar_potencia_regulada_idx, desconto_ts_potencia_dia_bruto)

            # --- Passo 3: Calcular Preço Final Energia (€/kWh, Sem IVA) ---
            preco_energia_final_sem_iva_idx = {}
            for periodo in preco_comercializador_energia_idx.keys():
                preco_energia_final_sem_iva_idx[periodo] = (
                    preco_comercializador_energia_idx[periodo]
                    + tar_energia_final_idx.get(periodo, 0.0)
                    + financiamento_tse_a_adicionar_idx
                )

            # --- Passo 4: Calcular Componentes Finais Potência (€/dia, Sem IVA) ---
            preco_comercializador_potencia_final_sem_iva_idx = preco_comercializador_potencia_idx
            tar_potencia_final_dia_sem_iva_idx = tar_potencia_final_dia_idx

            # --- Passo 5: Calcular Custo Total Energia (Com IVA) ---
            custo_energia_idx_com_iva = calcular_custo_energia_com_iva(
                consumo,
                preco_energia_final_sem_iva_idx.get('S') if opcao_horaria.lower() == "simples" else None,
                {p: v for p, v in preco_energia_final_sem_iva_idx.items() if p != 'S'},
                dias, potencia, opcao_horaria,
                consumos_horarios_para_func_idx, # Já definido acima
                familia_numerosa
            )

            # --- Passo 6: Calcular Custo Total Potência (Com IVA) ---
            custo_potencia_idx_com_iva = calcular_custo_potencia_com_iva_final(
                preco_comercializador_potencia_final_sem_iva_idx,
                tar_potencia_final_dia_sem_iva_idx,
                dias,
                potencia
            )

            comercializador_tarifario_idx = tarifario_indexado['comercializador'] # Nome do comercializador

            # --- Passo 7: Calcular Taxas Adicionais ---
            taxas_idx = calcular_taxas_adicionais(
                consumo, dias, tarifa_social,
                valor_dgeg_user, valor_cav_user,
                nome_comercializador_atual=comercializador_tarifario_idx, # Passa o comercializador
                mes_selecionado_simulacao=mes,
                ano_simulacao_atual=ano_atual
            )

            # --- Passo 8: Calcular Custo Total Final ---
            custo_total_antes_desc_fatura_idx = custo_energia_idx_com_iva['custo_com_iva'] + custo_potencia_idx_com_iva['custo_com_iva'] + taxas_idx['custo_com_iva']

            # Determinar se o período de simulação é um mês civil completo (mesma lógica acima)
            e_mes_completo_selecionado = False # Recalcular ou passar como argumento se estiver numa função
            try:
                dias_no_mes_do_input_widget = dias_mes[mes]
                primeiro_dia_do_mes_widget = datetime.date(ano_atual, mes_num, 1)
                ultimo_dia_do_mes_widget = datetime.date(ano_atual, mes_num, dias_no_mes_do_input_widget)
                if data_inicio == primeiro_dia_do_mes_widget and data_fim == ultimo_dia_do_mes_widget:
                    e_mes_completo_selecionado = True
            except Exception:
                e_mes_completo_selecionado = False


            # --- Aplicar desconto_fatura_mes ---
            desconto_fatura_mensal_idx = 0.0
            nome_tarifario_original_idx = str(nome_tarifario) # Guardar o nome original

            if 'desconto_fatura_mes' in tarifario_indexado and pd.notna(tarifario_indexado['desconto_fatura_mes']):
                try:
                    desconto_fatura_mensal_idx = float(tarifario_indexado['desconto_fatura_mes'])
                    if desconto_fatura_mensal_idx > 0: # Só adiciona ao nome se o desconto for positivo
                        nome_tarifario += f" (INCLUI desconto {desconto_fatura_mensal_idx:.2f} €/mês)"
                except ValueError:
                    desconto_fatura_mensal_idx = 0.0

            if e_mes_completo_selecionado:
                desconto_fatura_periodo_idx = desconto_fatura_mensal_idx
            else:
                desconto_fatura_periodo_idx = (desconto_fatura_mensal_idx / 30.0) * dias if dias > 0 else 0
            
            custo_total_estimado_idx = custo_total_antes_desc_fatura_idx - desconto_fatura_periodo_idx
            # --- FIM Aplicar desconto_fatura_mes ---

            # --- INÍCIO: CAMPOS PARA TOOLTIPS DE ENERGIA (INDEXADOS) ---
            componentes_tooltip_energia_dict_idx = {}
            ts_global_ativa_idx = tarifa_social # Flag global de TS

            # Loop pelos períodos de energia (S, V, F, C, P) que existem para este tarifário indexado
            # Certifique-se que preco_comercializador_energia_idx.keys() tem os períodos corretos (S ou V,F ou V,C,P)
            for periodo_key_idx in preco_comercializador_energia_idx.keys():
                comp_comerc_energia_base_idx = preco_comercializador_energia_idx.get(periodo_key_idx, 0.0)
                tar_bruta_energia_periodo_idx = tar_energia_regulada_idx.get(periodo_key_idx, 0.0)

                # Flag 'financiamento_tse_incluido_idx' lida do Excel para ESTE tarifário
                tse_declarado_incluido_excel_idx = financiamento_tse_incluido_idx
                tse_valor_nominal_const_idx = FINANCIAMENTO_TSE_VAL

                ts_aplicada_energia_flag_para_tooltip_idx = ts_global_ativa_idx
                desconto_ts_energia_unitario_para_tooltip_idx = 0.0
                if ts_global_ativa_idx:
                    desconto_ts_energia_unitario_para_tooltip_idx = obter_constante('Desconto TS Energia', CONSTANTES)

                componentes_tooltip_energia_dict_idx[f'tooltip_energia_{periodo_key_idx}_comerc_sem_tar'] = comp_comerc_energia_base_idx
                componentes_tooltip_energia_dict_idx[f'tooltip_energia_{periodo_key_idx}_tar_bruta'] = tar_bruta_energia_periodo_idx
                componentes_tooltip_energia_dict_idx[f'tooltip_energia_{periodo_key_idx}_tse_declarado_incluido'] = tse_declarado_incluido_excel_idx
                componentes_tooltip_energia_dict_idx[f'tooltip_energia_{periodo_key_idx}_tse_valor_nominal'] = tse_valor_nominal_const_idx
                componentes_tooltip_energia_dict_idx[f'tooltip_energia_{periodo_key_idx}_ts_aplicada_flag'] = ts_aplicada_energia_flag_para_tooltip_idx
                componentes_tooltip_energia_dict_idx[f'tooltip_energia_{periodo_key_idx}_ts_desconto_valor'] = desconto_ts_energia_unitario_para_tooltip_idx
            # --- FIM: CAMPOS PARA TOOLTIPS DE ENERGIA (INDEXADOS) ---

            desconto_ts_potencia_valor_aplicado_idx = 0.0
            if ts_global_ativa_idx:
                desconto_ts_potencia_dia_bruto_idx = obter_constante(f'Desconto TS Potencia {potencia}', CONSTANTES)
                # tar_potencia_regulada_idx é a TAR bruta para este tarifário indexado
                desconto_ts_potencia_valor_aplicado_idx = min(tar_potencia_regulada_idx, desconto_ts_potencia_dia_bruto_idx)

            # Para o tooltip do Preço Potência Indexados:
            componentes_tooltip_potencia_dict_idx = {
                'tooltip_pot_comerc_sem_tar': preco_comercializador_potencia_idx,
                'tooltip_pot_tar_bruta': tar_potencia_regulada_idx,
                'tooltip_pot_ts_aplicada': ts_global_ativa,
                'tooltip_pot_desconto_ts_valor': desconto_ts_potencia_valor_aplicado
            }

            # --- PASSO X: CALCULAR CUSTOS COM IVA E OBTER DECOMPOSIÇÃO PARA TOOLTIP ---

            # ENERGIA (Tarifários Indexados)
            preco_energia_simples_para_iva_idx = None
            precos_energia_horarios_para_iva_idx = {}
            if opcao_horaria.lower() == "simples":
                preco_energia_simples_para_iva_idx = preco_energia_final_sem_iva_idx.get('S')
            else:
                precos_energia_horarios_para_iva_idx = {
                    p: val for p, val in preco_energia_final_sem_iva_idx.items() if p != 'S'
                }

            decomposicao_custo_energia_idx = calcular_custo_energia_com_iva(
                consumo, # Consumo total global
                preco_energia_simples_para_iva_idx,
                precos_energia_horarios_para_iva_idx,
                dias, potencia, opcao_horaria,
                consumos_horarios_para_func_idx, # Dicionário de consumos por período para este tarifário
                familia_numerosa
            )
            custo_energia_idx_com_iva = decomposicao_custo_energia_idx['custo_com_iva']
            tt_cte_energia_siva_idx = decomposicao_custo_energia_idx['custo_sem_iva']
            tt_cte_energia_iva_6_idx = decomposicao_custo_energia_idx['valor_iva_6']
            tt_cte_energia_iva_23_idx = decomposicao_custo_energia_idx['valor_iva_23']

            # POTÊNCIA (Tarifários Indexados)
            decomposicao_custo_potencia_idx = calcular_custo_potencia_com_iva_final(
                preco_comercializador_potencia_final_sem_iva_idx,
                tar_potencia_final_dia_sem_iva_idx, # Esta já tem TS se aplicável
                dias,
                potencia
            )
            custo_potencia_idx_com_iva = decomposicao_custo_potencia_idx['custo_com_iva']
            tt_cte_potencia_siva_idx = decomposicao_custo_potencia_idx['custo_sem_iva']
            tt_cte_potencia_iva_6_idx = decomposicao_custo_potencia_idx['valor_iva_6']
            tt_cte_potencia_iva_23_idx = decomposicao_custo_potencia_idx['valor_iva_23']
            
            # TAXAS ADICIONAIS (Tarifários Indexados)
            decomposicao_taxas_idx = calcular_taxas_adicionais(
                consumo, dias, tarifa_social,
                valor_dgeg_user, valor_cav_user,
                nome_comercializador_atual=comercializador_tarifario_idx, # Passa o comercializador
                mes_selecionado_simulacao=mes,
                ano_simulacao_atual=ano_atual
            )
            taxas_idx_com_iva = decomposicao_taxas_idx['custo_com_iva']
            tt_cte_iec_siva_idx = decomposicao_taxas_idx['iec_sem_iva']
            tt_cte_dgeg_siva_idx = decomposicao_taxas_idx['dgeg_sem_iva']
            tt_cte_cav_siva_idx = decomposicao_taxas_idx['cav_sem_iva']
            tt_cte_taxas_iva_6_idx = decomposicao_taxas_idx['valor_iva_6']
            tt_cte_taxas_iva_23_idx = decomposicao_taxas_idx['valor_iva_23']

            # Custo Total antes de outros descontos específicos do tarifário indexado
            custo_total_antes_desc_especificos_idx = custo_energia_idx_com_iva + custo_potencia_idx_com_iva + taxas_idx_com_iva

            # Calcular totais para o tooltip do Custo Total Estimado
            tt_cte_total_siva_idx = tt_cte_energia_siva_idx + tt_cte_potencia_siva_idx + tt_cte_iec_siva_idx + tt_cte_dgeg_siva_idx + tt_cte_cav_siva_idx
            tt_cte_valor_iva_6_total_idx = tt_cte_energia_iva_6_idx + tt_cte_potencia_iva_6_idx + tt_cte_taxas_iva_6_idx
            tt_cte_valor_iva_23_total_idx = tt_cte_energia_iva_23_idx + tt_cte_potencia_iva_23_idx + tt_cte_taxas_iva_23_idx

            # NOVO: Calcular Subtotal c/IVA (antes de descontos/acréscimos finais)
            tt_cte_subtotal_civa_idx = tt_cte_total_siva_idx + tt_cte_valor_iva_6_total_idx + tt_cte_valor_iva_23_total_idx
        
            # Consolidar Outros Descontos e Acréscimos Finais
            # Para indexados, geralmente é só o desconto_fatura_periodo_idx
            tt_cte_desc_finais_valor_idx = 0.0
            if 'desconto_fatura_periodo_idx' in locals() and desconto_fatura_periodo_idx > 0:
                 tt_cte_desc_finais_valor_idx = desconto_fatura_periodo_idx
             
            tt_cte_acres_finais_valor_idx = 0.0 # Tipicamente não há para indexados, a menos que adicione

            # Adicionar os novos campos de tooltip ao resultado_indexado
            componentes_tooltip_custo_total_dict_idx = {
                'tt_cte_energia_siva': tt_cte_energia_siva_idx,
                'tt_cte_potencia_siva': tt_cte_potencia_siva_idx,
                'tt_cte_iec_siva': tt_cte_iec_siva_idx,
                'tt_cte_dgeg_siva': tt_cte_dgeg_siva_idx,
                'tt_cte_cav_siva': tt_cte_cav_siva_idx,
                'tt_cte_total_siva': tt_cte_total_siva_idx,
                'tt_cte_valor_iva_6_total': tt_cte_valor_iva_6_total_idx,
                'tt_cte_valor_iva_23_total': tt_cte_valor_iva_23_total_idx,
                'tt_cte_subtotal_civa': tt_cte_subtotal_civa_idx,
                'tt_cte_desc_finais_valor': tt_cte_desc_finais_valor_idx,
                'tt_cte_acres_finais_valor': tt_cte_acres_finais_valor_idx

                }
                # resultados_list.append(resultado_indexado)

            # --- Passo 9: Preparar Resultados para Exibição ---
            valores_energia_exibir_idx = {}
            for p, v in preco_energia_final_sem_iva_idx.items(): # Preços finais s/IVA
                 periodo_nome = ""
                 if p == 'S': periodo_nome = "Simples"
                 elif p == 'V': periodo_nome = "Vazio"
                 elif p == 'F': periodo_nome = "Fora Vazio"
                 elif p == 'C': periodo_nome = "Cheias"
                 elif p == 'P': periodo_nome = "Ponta"
                 if periodo_nome:
                     valores_energia_exibir_idx[f'{periodo_nome} (€/kWh)'] = round(v, 4)

            preco_potencia_total_final_sem_iva_idx = preco_comercializador_potencia_final_sem_iva_idx + tar_potencia_final_dia_sem_iva_idx

            if pd.notna(custo_total_estimado_idx):
                resultado_indexado = {
                    'NomeParaExibir': nome_tarifario,
                    'LinkAdesao': link_adesao_idx,
                    'info_notas': notas_tarifario_idx,
                    'Tipo': tipo_tarifario,
                    'Segmento': segmento_tarifario,
                    'Faturação': faturacao_tarifario,
                    'Pagamento': pagamento_tarifario,
                    'Comercializador': comercializador_tarifario,
                    **valores_energia_exibir_idx,
                    'Potência (€/dia)': round(preco_potencia_total_final_sem_iva_idx, 4), # Preço final s/IVA
                    'Total (€)': round(custo_total_estimado_idx, 2),
                    # CAMPOS DO TOOLTIP DA POTÊNCIA INDEXADOS
                    **componentes_tooltip_potencia_dict_idx,
                    # CAMPOS DO TOOLTIP DA ENERGIA INDEXADOS
                    **componentes_tooltip_energia_dict_idx, 
                    # CAMPOS DO TOOLTIP DA CUSTO TOTAL FIXOS
                    **componentes_tooltip_custo_total_dict_idx, 
                    }
                resultados_list.append(resultado_indexado)
# --- Fim do loop for tarifario_indexado ---

# --- Fim do if comparar_indexados ---

# --- Processamento final e exibição da tabela de resultados ---
st.subheader("💰 Tiago Felícia - Tarifários de Eletricidade - Detalhado")


vista_simplificada = st.checkbox(
    "📱 Ativar vista simplificada (ideal em ecrãs menores)",
    value=True,
    key="chk_vista_simplificada"
    )

st.write("Total com todos os componentes, taxas e impostos e Valores unitários sem IVA")

st.markdown("➡️ [**Exportar Tabela Detalhada para Excel**](#exportar-excel-detalhada)")

# Verifica se "O Meu Tarifário" deve ser incluído
final_results_list = resultados_list.copy() # Começa com os tarifários fixos e/ou indexados
if meu_tarifario_ativo and 'meu_tarifario_calculado' in st.session_state:
    dados_meu_tarifario_guardado = st.session_state['meu_tarifario_calculado']
    # 'opcao_horaria' aqui é o valor atual do selectbox principal "Opção Horária e Ciclo"
    if dados_meu_tarifario_guardado.get('opcao_horaria_calculada') == opcao_horaria:
        final_results_list.append(dados_meu_tarifario_guardado)

df_resultados = pd.DataFrame(final_results_list)

try:
    # Inicializar/resetar variáveis do session_state para a mensagem do Excel
    # Isto garante que não usamos dados de uma execução anterior se as condições atuais não gerarem uma nova mensagem.
    st.session_state.poupanca_excel_texto = ""
    st.session_state.poupanca_excel_cor = "000000"  # Preto por defeito (formato RRGGBB)
    st.session_state.poupanca_excel_negrito = False
    st.session_state.poupanca_excel_disponivel = False # Flag para indicar se há mensagem para o Excel

    if meu_tarifario_ativo and not df_resultados.empty: # df_resultados é o DataFrame da UI
        meu_tarifario_linha = df_resultados[df_resultados['NomeParaExibir'].str.contains("O Meu Tarifário", case=False, na=False)]

        if not meu_tarifario_linha.empty:
            custo_meu_tarifario = meu_tarifario_linha['Total (€)'].iloc[0]
            nome_meu_tarifario_ui = meu_tarifario_linha['NomeParaExibir'].iloc[0] # Usar _ui para clareza

            if pd.notna(custo_meu_tarifario):
                outros_tarifarios_ui_df = df_resultados[
                    ~df_resultados['NomeParaExibir'].str.contains("O Meu Tarifário", case=False, na=False)
                ]
                custos_outros_validos_ui = outros_tarifarios_ui_df['Total (€)'].dropna()
                
                mensagem_poupanca_html_ui = "" # Para a UI

                if not custos_outros_validos_ui.empty:
                    custo_minimo_outros_ui = custos_outros_validos_ui.min()
                    linha_mais_barata_outros_ui = outros_tarifarios_ui_df.loc[custos_outros_validos_ui.idxmin()]
                    nome_tarifario_mais_barato_outros_ui = linha_mais_barata_outros_ui['NomeParaExibir']

                    if custo_meu_tarifario > custo_minimo_outros_ui:
                        poupanca_abs_ui = custo_meu_tarifario - custo_minimo_outros_ui
                        poupanca_rel_ui = (poupanca_abs_ui / custo_meu_tarifario) * 100 if custo_meu_tarifario != 0 else 0
                        
                        mensagem_poupanca_html_ui = (
                            f"<span style='color:red; font-weight:bold;'>Poupança entre '{nome_meu_tarifario_ui}' ({custo_meu_tarifario:.2f} €) e o mais económico da lista, "
                            f"{nome_tarifario_mais_barato_outros_ui} ({custo_minimo_outros_ui:.2f} €): </span>"
                            f"<span style='color:red; font-weight:bold;'>{poupanca_abs_ui:.2f} €</span> "
                            f"<span style='color:red; font-weight:bold;'>({poupanca_rel_ui:.2f} %).</span>"
                        )
                        # Guardar para Excel
                        st.session_state.poupanca_excel_texto = (
                            f"Poupança entre '{nome_meu_tarifario_ui}' ({custo_meu_tarifario:.2f} €) e o mais económico da lista, "
                            f"{nome_tarifario_mais_barato_outros_ui} ({custo_minimo_outros_ui:.2f} €): "
                            f"{poupanca_abs_ui:.2f} € ({poupanca_rel_ui:.2f} %)."
                        )
                        st.session_state.poupanca_excel_cor = "FF0000" # Vermelho
                        st.session_state.poupanca_excel_negrito = True
                        st.session_state.poupanca_excel_disponivel = True
                    
                    elif custo_meu_tarifario <= custo_minimo_outros_ui:
                        mensagem_poupanca_html_ui = f"<span style='color:green; font-weight:bold;'>Parabéns! O seu tarifário ('{nome_meu_tarifario_ui}' - {custo_meu_tarifario:.2f}€) já é o mais económico ou está entre os mais económicos da lista!</span>"
                        st.session_state.poupanca_excel_texto = f"Parabéns! O seu tarifário ('{nome_meu_tarifario_ui}' - {custo_meu_tarifario:.2f}€) já é o mais económico ou está entre os mais económicos da lista!"
                        st.session_state.poupanca_excel_cor = "008000" # Verde
                        st.session_state.poupanca_excel_negrito = True
                        st.session_state.poupanca_excel_disponivel = True
                
                elif len(df_resultados) == 1: # Só "O Meu Tarifário" com custo válido e mais nenhum
                    mensagem_poupanca_html_ui = f"<span style='color:green; font-weight:bold;'>'{nome_meu_tarifario_ui}' ({custo_meu_tarifario:.2f}€) é o único tarifário na lista.</span>"
                    st.session_state.poupanca_excel_texto = f"'{nome_meu_tarifario_ui}' ({custo_meu_tarifario:.2f}€) é o único tarifário na lista."
                    st.session_state.poupanca_excel_cor = "000000" # Preto
                    st.session_state.poupanca_excel_negrito = True # Ou False, conforme preferir
                    st.session_state.poupanca_excel_disponivel = True
                else: # Meu tarifário tem custo, mas não há outros para comparar
                    mensagem_poupanca_html_ui = f"<span style='color:black; font-weight:normal;'>Não há outros tarifários com custos válidos para comparar com '{nome_meu_tarifario_ui}' ({custo_meu_tarifario:.2f}€).</span>"
                    st.session_state.poupanca_excel_texto = f"Não há outros tarifários com custos válidos para comparar com '{nome_meu_tarifario_ui}' ({custo_meu_tarifario:.2f}€)."
                    st.session_state.poupanca_excel_cor = "000000" # Preto
                    st.session_state.poupanca_excel_negrito = False
                    st.session_state.poupanca_excel_disponivel = True
                
                if mensagem_poupanca_html_ui:
                    st.markdown(mensagem_poupanca_html_ui, unsafe_allow_html=True)

            else: # Custo do Meu Tarifário é NaN ou não é válido
                st.info("Custo do 'Meu Tarifário' não pôde ser calculado. Não é possível determinar poupança.")
                st.session_state.poupanca_excel_texto = "Custo do 'Meu Tarifário' não pôde ser calculado. Não é possível determinar poupança."
                st.session_state.poupanca_excel_disponivel = True # Há uma mensagem informativa para o Excel

        # else: Se "Meu Tarifário" não foi encontrado, as variáveis de session_state ficam com os valores de inicialização (mensagem vazia, flag False)
    
    elif meu_tarifario_ativo: # Meu tarifário está ativo mas df_resultados está vazio ou não contém o meu tarifário
        st.info("Ative e calcule 'O Meu Tarifário' ou verifique os resultados para ver a poupança na interface.")
        # Para o Excel, podemos também querer indicar isto
        st.session_state.poupanca_excel_texto = "Informação de poupança não disponível (verifique 'O Meu Tarifário' ou os resultados)."
        st.session_state.poupanca_excel_disponivel = True

    # Se meu_tarifario_ativo for False, não fazemos nada aqui, e poupanca_excel_disponivel permanecerá False

except Exception as e_poupanca: # Renomeado para e_poupanca_ui para evitar conflitos se houver outro try-except
    st.error(f"Erro ao processar a informação de poupança para UI: {e_poupanca}")
    st.session_state.poupanca_excel_texto = "Erro ao calcular a informação de poupança."
    st.session_state.poupanca_excel_disponivel = True # Indica que houve um problema, pode ser útil no Excel
# --- FIM DO BLOCO PARA EXIBIR POUPANÇA ---

#ATENÇÃO, PODE CAUSAR PROBLEMAS
st.empty()
import time
time.sleep(0.2) # Geralmente uma má ideia em apps Streamlit

if not df_resultados.empty:
    if vista_simplificada:
        # Definir a ordem específica para a vista simplificada
        colunas_base_simplificada = ['NomeParaExibir', 'Total (€)']
        nomes_periodos_energia = ["Simples", "Vazio", "Fora Vazio", "Cheias", "Ponta"]
        colunas_energia_existentes = [f'{p_nome} (€/kWh)' for p_nome in nomes_periodos_energia if f'{p_nome} (€/kWh)' in df_resultados.columns]
        coluna_potencia = 'Potência (€/dia)'
        col_order_visivel_aggrid = colunas_base_simplificada + colunas_energia_existentes # colunas_base_simplificada precisa ser definida antes
        if coluna_potencia in df_resultados.columns:
            col_order_visivel_aggrid.append(coluna_potencia)

        colunas_visiveis_presentes = [col for col in col_order_visivel_aggrid if col in df_resultados.columns]

    else:
        # Mapear colunas de exibição desejadas
        nomes_periodos_energia = ["Simples", "Vazio", "Fora Vazio", "Cheias", "Ponta"]
        colunas_energia_esperadas = [f'{p_nome} (€/kWh)' for p_nome in nomes_periodos_energia]
        col_order_visivel_aggrid = ['NomeParaExibir', 'LinkAdesao', 'Total (€)']
        col_order_visivel_aggrid.extend([col for col in colunas_energia_esperadas if col in df_resultados.columns])
        col_order_visivel_aggrid.extend(['Potência (€/dia)'])
        col_order_visivel_aggrid.extend(['Tipo', 'Comercializador', 'Segmento', 'Faturação', 'Pagamento'])
        colunas_visiveis_presentes = [col for col in col_order_visivel_aggrid if col in df_resultados.columns]

    # --- NOVO: Definir colunas necessárias para os dados dos tooltips ---
    colunas_dados_tooltip = [
        'tooltip_pot_comerc_sem_tar', 'tooltip_pot_tar_bruta', 'tooltip_pot_ts_aplicada', 'tooltip_pot_desconto_ts_valor',
        # Energia Simples (S)
        'tooltip_energia_S_comerc_sem_tar', 'tooltip_energia_S_tar_bruta', 
        'tooltip_energia_S_tse_declarado_incluido', 'tooltip_energia_S_tse_valor_nominal',
        'tooltip_energia_S_ts_aplicada_flag', 'tooltip_energia_S_ts_desconto_valor',
    
        # Energia Vazio (V) - Adicione se tiver opção Bi ou Tri
        'tooltip_energia_V_comerc_sem_tar', 'tooltip_energia_V_tar_bruta', 
        'tooltip_energia_V_tse_declarado_incluido', 'tooltip_energia_V_tse_valor_nominal',
        'tooltip_energia_V_ts_aplicada_flag', 'tooltip_energia_V_ts_desconto_valor',

        # Energia Fora Vazio (F) - Adicione se tiver opção Bi
        'tooltip_energia_F_comerc_sem_tar', 'tooltip_energia_F_tar_bruta', 
        'tooltip_energia_F_tse_declarado_incluido', 'tooltip_energia_F_tse_valor_nominal',
        'tooltip_energia_F_ts_aplicada_flag', 'tooltip_energia_F_ts_desconto_valor',

        # Energia Cheias (C) - Adicione se tiver opção Tri
        'tooltip_energia_C_comerc_sem_tar', 'tooltip_energia_C_tar_bruta', 
        'tooltip_energia_C_tse_declarado_incluido', 'tooltip_energia_C_tse_valor_nominal',
        'tooltip_energia_C_ts_aplicada_flag', 'tooltip_energia_C_ts_desconto_valor',

        # Energia Ponta (P) - Adicione se tiver opção Tri
        'tooltip_energia_P_comerc_sem_tar', 'tooltip_energia_P_tar_bruta', 
        'tooltip_energia_P_tse_declarado_incluido', 'tooltip_energia_P_tse_valor_nominal',
        'tooltip_energia_P_ts_aplicada_flag', 'tooltip_energia_P_ts_desconto_valor',

        # Para Custo Total
        'tt_cte_energia_siva', 'tt_cte_potencia_siva', 'tt_cte_iec_siva',
        'tt_cte_dgeg_siva', 'tt_cte_cav_siva', 'tt_cte_total_siva',
        'tt_cte_valor_iva_6_total', 'tt_cte_valor_iva_23_total',
        'tt_cte_subtotal_civa','tt_cte_desc_finais_valor','tt_cte_acres_finais_valor'
    ]

    # Colunas que DEVEM estar presentes nos dados do AgGrid para lógica JS, mesmo que ocultas visualmente
    colunas_essenciais_para_js = ['Tipo', 'NomeParaExibir', 'LinkAdesao', 'info_notas'] # Adicione outras se necessário
    colunas_essenciais_para_js.extend(colunas_dados_tooltip) # as de tooltip já estão aqui

    # Unir colunas visíveis e essenciais para JS, removendo duplicados e mantendo a ordem das visíveis primeiro
    colunas_para_aggrid_final = list(dict.fromkeys(colunas_visiveis_presentes + colunas_essenciais_para_js))
    
    # Filtrar para garantir que todas as colunas em colunas_para_aggrid_final existem em df_resultados
    colunas_para_aggrid_final = [col for col in colunas_para_aggrid_final if col in df_resultados.columns]


    # Verifica se as colunas essenciais 'NomeParaExibir' e 'LinkAdesao' existem
    # Se não existirem, o AgGrid pode não funcionar como esperado para os links.
    if not all(col in df_resultados.columns for col in ['NomeParaExibir', 'LinkAdesao']):
        st.error("Erro: O DataFrame de resultados não contém as colunas 'NomeParaExibir' e/ou 'LinkAdesao' necessárias para o AgGrid. Verifique a construção da lista de resultados.")

    else:
        df_resultados_para_aggrid = df_resultados[colunas_para_aggrid_final].copy()

        if 'Total (€)' in df_resultados_para_aggrid.columns:
            df_resultados_para_aggrid = df_resultados_para_aggrid.sort_values(by='Total (€)')
        df_resultados_para_aggrid = df_resultados_para_aggrid.reset_index(drop=True)

        # ---- INÍCIO DA CONFIGURAÇÃO DO AGGRID ----
        gb = GridOptionsBuilder.from_dataframe(df_resultados_para_aggrid)

        # --- Configurações Padrão para Colunas ---
        gb.configure_default_column(
            sortable=True,
            resizable=True,
            editable=False,
            wrapText=True,
            autoHeight=True,
            wrapHeaderText=True,    # Permite quebra de linha no TEXTO DO CABEÇALHO
            autoHeaderHeight=True   # Ajusta a ALTURA DO CABEÇALHO para o texto quebrado
        )

        # --- 1. DEFINIR O JsCode PARA LINK E TOOLTIP ---
        link_tooltip_renderer_js = JsCode("""
        class LinkTooltipRenderer {
            init(params) {
                this.eGui = document.createElement('div');
                let displayText = params.value; // Valor da célula (NomeParaExibir)
                let url = params.data.LinkAdesao; // Acede ao valor da coluna LinkAdesao da mesma linha

                if (url && typeof url === 'string' && url.toLowerCase().startsWith('http')) {
                    // HTML para o link clicável
                    // O atributo 'title' (tooltip) mostrará "Aderir/Saber mais: [URL]"
                    // O texto visível do link será o 'displayText' (NomeParaExibir)
                    this.eGui.innerHTML = `<a href="${url}" target="_blank" title="Aderir/Saber mais: ${url}" style="text-decoration: underline; color: inherit;">${displayText}</a>`;
                } else {
                    // Se não houver URL válido, apenas mostra o displayText com o próprio displayText como tooltip.
                    this.eGui.innerHTML = `<span title="${displayText}">${displayText}</span>`;
                }
            }
            getGui() { return this.eGui; }
        }
        """) # <--- FIM DA DEFINIÇÃO DE link_tooltip_renderer_js
        
        #CORES PARA TARIFÁRIOS INDEXADOS:
        cor_fundo_indexado_media_css = "#FFE699"
        cor_texto_indexado_media_css = "black"
        cor_fundo_indexado_dinamico_css = "#4D79BC"  
        cor_texto_indexado_dinamico_css = "white"

        cell_style_nome_tarifario_js = JsCode(f"""
        function(params) {{
            // Estilo base aplicado a todas as células desta coluna
            let styleToApply = {{ 
                textAlign: 'center',
                borderRadius: '6px',  // O teu borderRadius desejado
                padding: '10px 10px'   // O teu padding desejado
                // Podes adicionar um backgroundColor default para células não especiais aqui, se quiseres
                // backgroundColor: '#f0f0f0' // Exemplo para tarifários fixos
            }};                                  

            if (params.data) {{
                const nomeExibir = params.data.NomeParaExibir;
                const tipoTarifario = params.data.Tipo;

                // VERIFICA SE O NOME COMEÇA COM "O Meu Tarifário"
                if (typeof nomeExibir === 'string' && nomeExibir.startsWith('O Meu Tarifário')) {{
                    styleToApply.backgroundColor = 'red';
                    styleToApply.color = 'white';
                    styleToApply.fontWeight = 'bold';
                }} else if (tipoTarifario === 'Indexado Média') {{
                    styleToApply.backgroundColor = '{cor_fundo_indexado_media_css}';
                    styleToApply.color = '{cor_texto_indexado_media_css}';
                }} else if (tipoTarifario === 'Indexado quarto-horário') {{
                    styleToApply.backgroundColor = '{cor_fundo_indexado_dinamico_css}';
                    styleToApply.color = '{cor_texto_indexado_dinamico_css}';
                }} else {{
                    // Para tarifários fixos ou outros tipos não explicitamente coloridos acima.
                    // Eles já terão o textAlign, borderRadius e padding do styleToApply.
                    // Se quiseres um fundo específico para eles diferente do default do styleToApply, define aqui.
                    // Ex: styleToApply.backgroundColor = '#e9ecef'; // Uma cor neutra para fixos
                }}
                return styleToApply;
            }}
            return styleToApply; 
        }}
        """)

        #tooltip Nome Tarifario
        tooltip_nome_tarifario_getter_js = JsCode("""
        function(params) {
            if (!params.data) { 
                return params.value || ''; 
            }

            const nomeExibir = params.data.NomeParaExibir || '';
            const notas = params.data.info_notas || ''; 

            let tooltipHtmlParts = [];

            if (nomeExibir) {
                tooltipHtmlParts.push("<strong>" + nomeExibir + "</strong>");
            }

            if (notas) {
                const notasHtml = notas.replace(/\\n/g, ' ').replace(/\n/g, ' ');
                // Usando aspas simples para atributos de estilo HTML para simplificar
                tooltipHtmlParts.push("<small style='display: block; margin-top: 5px;'><i>" + notasHtml + "</i></small>");
            }
    
            if (tooltipHtmlParts.length > 0) {
                // Se ambos nomeExibir e notas existem, queremos uma quebra de linha entre eles no tooltip.
                // Se join(''), eles ficam lado a lado. Se join('<br>'), ficam em linhas separadas.
                // Dado que a nota tem display:block, o join('') deve funcionar para colocá-los em "blocos" separados.
                return tooltipHtmlParts.join(''); // Para agora, vamos juntar diretamente.
                                                 // Se quiser uma quebra de linha explícita entre o nome e as notas,
                                                 // e ambos existirem, pode usar:
                                                 // return tooltipHtmlParts.join('<br style="margin-bottom:5px">');
            }
    
            return ''; 
        }
                """)
        
        # Custom_Tooltip
        custom_tooltip_component_js = JsCode("""
            class CustomTooltip {
                init(params) {
                    // params.value é a string que o seu tooltipValueGetter retorna
                    this.eGui = document.createElement('div');
                    // Para permitir HTML, definimos o innerHTML
                    // É importante que a string de params.value seja HTML seguro se vier de inputs do utilizador,
                    // mas no seu caso, está a construí-lo programaticamente.
                    this.eGui.innerHTML = params.value; 

                    // Aplicar algum estilo básico para o tooltip se desejar
                    this.eGui.style.backgroundColor = 'white'; // Ou outra cor de fundo
                    this.eGui.style.color = 'black';           // Cor do texto
                    this.eGui.style.border = '1px solid #ccc'; // Borda mais suave
                    this.eGui.style.padding = '10px';           // Mais padding
                    this.eGui.style.borderRadius = '6px';      // Cantos arredondados
                    this.eGui.style.boxShadow = '0 2px 5px rgba(0,0,0,0.15)'; // Sombra suave
                    this.eGui.style.maxWidth = '400px';        // Largura máxima
                    this.eGui.style.fontSize = '1.1em';        // Tamanho da fonte
                    this.eGui.style.fontFamily = 'Arial, sans-serif'; // Tipo de fonte                             
                    this.eGui.style.whiteSpace = 'normal';     // Para quebra de linha
                }

                getGui() {
                    return this.eGui;
                }
            }
        """)

        # --- Configuração Coluna Tarifário com Link e Tooltip ---
        gb.configure_column(field='NomeParaExibir', headerName='Tarifário', cellRenderer=link_tooltip_renderer_js, minWidth=100, flex=2, filter='agTextColumnFilter', tooltipValueGetter=tooltip_nome_tarifario_getter_js, tooltipComponent=custom_tooltip_component_js,
    cellStyle=cell_style_nome_tarifario_js)
        if 'LinkAdesao' in df_resultados_para_aggrid.columns:
            gb.configure_column(field='LinkAdesao', hide=True) # Desativar filtro explicitamente
                    
        # --- 2. Formatação Condicional de Cores ---
        cols_para_cor = [
            col for col in df_resultados_para_aggrid.columns
            if '(€/kWh)' in col or '(€/dia)' in col or 'Total (€)' == col
        ]
        min_max_data_for_js = {}

        for col_name in cols_para_cor:
            if col_name in df_resultados_para_aggrid:
                series = pd.to_numeric(df_resultados_para_aggrid[col_name], errors='coerce').dropna()
                if not series.empty:
                    min_max_data_for_js[col_name] = {'min': series.min(), 'max': series.max()}
                else:
                    min_max_data_for_js[col_name] = {'min': 0, 'max': 0}

        min_max_data_json_string = json.dumps(min_max_data_for_js)

    # Função get_color para JavaScript (para cor nas colunas de valor)
        cell_style_cores_js = JsCode(f"""
        function(params) {{
            const colName = params.colDef.field;
            const value = parseFloat(params.value);
            const minMaxConfig = {min_max_data_json_string}; //

            let style = {{ 
                textAlign: 'center',
                borderRadius: '6px',
                padding: '10px 10px'
            }};

            if (isNaN(value) || !minMaxConfig[colName]) {{
                return style; // Sem cor para NaN ou se não houver config min/max
            }}

            const min_val = minMaxConfig[colName].min;
            const max_val = minMaxConfig[colName].max;

            if (max_val === min_val) {{
                style.backgroundColor = 'lightgrey'; // Ou 'transparent'
                return style;
            }}

            const normalized_value = Math.max(0, Math.min(1, (value - min_val) / (max_val - min_val)));
            // Cores alvo do Excel
            const colorLow = {{ r: 99, g: 190, b: 123 }};  // Verde #63BE7B
            const colorMid = {{ r: 255, g: 255, b: 255 }}; // Branco #FFFFFF
            const colorHigh = {{ r: 248, g: 105, b: 107 }}; // Vermelho #F8696B

            let r, g, b;

            if (normalized_value < 0.5) {{
                // Interpolar entre colorLow (Verde) e colorMid (Branco)
                // t vai de 0 (no min) a 1 (no meio)
                const t = normalized_value / 0.5; 
                r = Math.round(colorLow.r * (1 - t) + colorMid.r * t);
                g = Math.round(colorLow.g * (1 - t) + colorMid.g * t);
                b = Math.round(colorLow.b * (1 - t) + colorMid.b * t);
            }} else {{
                // Interpolar entre colorMid (Branco) e colorHigh (Vermelho)
                // t vai de 0 (no meio) a 1 (no max)
                const t = (normalized_value - 0.5) / 0.5;
                r = Math.round(colorMid.r * (1 - t) + colorHigh.r * t);
                g = Math.round(colorMid.g * (1 - t) + colorHigh.g * t);
                b = Math.round(colorMid.b * (1 - t) + colorHigh.b * t);
            }}
            
            style.backgroundColor = `rgb(${'{r}'},${'{g}'},${'{b}'})`;
        
            // Lógica de contraste para a cor do texto (preto/branco)
            // Esta heurística calcula a luminância percebida.
            // Fundos mais escuros (<140-150) recebem texto branco, fundos mais claros recebem texto preto.
            // Pode ajustar o limiar 140 se necessário.
            if ((r * 0.299 + g * 0.587 + b * 0.114) < 140) {{ 
                style.color = 'white';
            }} else {{
                style.color = 'black';
            }}

            return style;
        }}
        """)
         # --- FIM DA DEFINIÇÃO DE cell_style_cores_js

        #Tooltip Preço Energia
        tooltip_preco_energia_js = JsCode("""
        function(params) {
            if (!params.data) { 
                // console.error("Tooltip Energia: params.data está AUSENTE para a célula com valor:", params.value, "e coluna:", params.colDef.field);
                // Decidi retornar apenas o valor da célula se não houver dados, em vez de uma string de erro no tooltip.
                return String(params.value); 
            }
                                                  
            const colField = params.colDef.field;
            let periodoKey = "";
            let nomePeriodoCompletoParaTitulo = "Energia"; // Um default caso algo falhe

            // Determinar a chave do período e o nome completo para o título
            if (colField.includes("Simples")) {
                periodoKey = "S";
                nomePeriodoCompletoParaTitulo = "Simples";
            } else if (colField.includes("Fora Vazio")) { // Importante verificar "Fora Vazio" antes de "Vazio"
                periodoKey = "F";
                nomePeriodoCompletoParaTitulo = "Fora Vazio";
            } else if (colField.includes("Vazio")) {
                periodoKey = "V";
                nomePeriodoCompletoParaTitulo = "Vazio";
            } else if (colField.includes("Cheias")) {
                periodoKey = "C";
                nomePeriodoCompletoParaTitulo = "Cheias";
            } else if (colField.includes("Ponta")) {
                periodoKey = "P";
                nomePeriodoCompletoParaTitulo = "Ponta";
            }

            if (!periodoKey) {
                // console.error("Tooltip Energia: Não foi possível identificar o período a partir de colField:", colField);
                return String(params.value); // Retorna o valor da célula se o período não for identificado
            }
    
            // Nomes exatos dos campos como definidos em Python
            const field_comerc = 'tooltip_energia_' + periodoKey + '_comerc_sem_tar';
            const field_tar_bruta = 'tooltip_energia_' + periodoKey + '_tar_bruta';
            const field_tse_declarado = 'tooltip_energia_' + periodoKey + '_tse_declarado_incluido';
            const field_tse_nominal = 'tooltip_energia_' + periodoKey + '_tse_valor_nominal';
            const field_ts_aplicada = 'tooltip_energia_' + periodoKey + '_ts_aplicada_flag';
            const field_ts_desconto = 'tooltip_energia_' + periodoKey + '_ts_desconto_valor';

            // Verificar se os campos de dados necessários para o tooltip existem
            if (typeof params.data[field_comerc] === 'undefined' || 
                typeof params.data[field_tar_bruta] === 'undefined' ||
                typeof params.data[field_tse_declarado] === 'undefined' ||
                typeof params.data[field_tse_nominal] === 'undefined' ||
                typeof params.data[field_ts_aplicada] === 'undefined' ||
                typeof params.data[field_ts_desconto] === 'undefined') {
        
                // console.warn("Tooltip Energia (" + periodoKey + "): Um ou mais campos de dados para o tooltip estão UNDEFINED. Coluna:", colField);
                return "Info decomposição indisponível."; // Mensagem mais clara se os dados não estiverem lá
            }
                                          
            const comercializador = parseFloat(params.data[field_comerc] || 0);
            const tarBruta = parseFloat(params.data[field_tar_bruta] || 0);
            const tseDeclaradoIncluido = params.data[field_tse_declarado]; // Booleano
            const tseValorNominal = parseFloat(params.data[field_tse_nominal] || 0);
            const tsAplicadaEnergia = params.data[field_ts_aplicada]; // Booleano
            const tsDescontoValorEnergia = parseFloat(params.data[field_ts_desconto] || 0);

            const formatPrice = (num, decimalPlaces) => {
                if (typeof num === 'number' && !isNaN(num)) {
                    return num.toFixed(decimalPlaces);
                }
                // console.warn("formatPrice (Energia): Tentativa de formatar valor não numérico:", num);
                return 'N/A';
            };
            
            // MODIFICADO: Construir o título dinamicamente
            let tituloTooltip = "<b>Decomposição Preço " + nomePeriodoCompletoParaTitulo + " (s/IVA):</b>";
            let tooltipParts = [tituloTooltip];

            tooltipParts.push("Comercializador (s/TAR): " + formatPrice(comercializador, 4) + " €/kWh");
            tooltipParts.push("TAR (Tarifa Acesso Redes): " + formatPrice(tarBruta, 4) + " €/kWh");

            if (tseDeclaradoIncluido === true) {
                tooltipParts.push("<i>(Financiamento TSE incluído no preço)</i>");
            } else if (tseDeclaradoIncluido === false && tseValorNominal > 0) { // Mostrar apenas se houver valor
                tooltipParts.push("Financiamento TSE: " + formatPrice(tseValorNominal, 7) + " €/kWh");
            } else if (tseDeclaradoIncluido !== true && tseDeclaradoIncluido !== false) { // Se não for booleano
                // console.warn("Tooltip Energia ("+periodoKey+"): Flag 'tseDeclaradoIncluido' tem valor inesperado:", tseDeclaradoIncluido);
                tooltipParts.push("<i>(Info Fin. TSE indisponível)</i>");
            }
    
            if (tsAplicadaEnergia === true && tsDescontoValorEnergia > 0) {
                tooltipParts.push("Desconto Tarifa Social: -" + formatPrice(tsDescontoValorEnergia, 4) + " €/kWh");
            }
    
            tooltipParts.push("----------------------------------------------------"); // Separador
            tooltipParts.push("<b>Custo Final : " + formatPrice(parseFloat(params.value), 4) + " €/kWh</b>");
    
            return tooltipParts.join("<br>");
        }
        """)

        # Configuração Coluna 'Preço Energia Simples (€/kWh)'
        col_energia_s_nome = 'Simples (€/kWh)'
        if col_energia_s_nome in df_resultados_para_aggrid.columns:
            casas_decimais_energia = 4
    
            js_value_formatter_energia = JsCode(f"""
                function(params) {{
                    // ... (sua lógica de valueFormatter para Energia)
                    if (params.value == null || typeof params.value === 'undefined' || String(params.value).trim() === '') {{
                        return '';
                    }}
                    let num = Number(params.value);
                    if (isNaN(num)) {{ return ''; }}
                    try {{
                        return num.toFixed({casas_decimais_energia});
                    }} catch (e) {{
                        console.error("Erro valueFormatter Energia:", e, params.value);
                        return String(params.value);
                    }}
                }}
            """)

            gb.configure_column(
                field=col_energia_s_nome,
                headerName=col_energia_s_nome,
                type=["numericColumn"],
                filter=False,
                valueFormatter=js_value_formatter_energia,
                cellStyle=cell_style_cores_js,
                tooltipValueGetter=tooltip_preco_energia_js,
                tooltipComponent=custom_tooltip_component_js,
                minWidth=60, # Ajuste conforme necessário
                flex=1
            )

        # Configuração Coluna 'Preço Energia Vazio (€/kWh)'
        col_energia_v_nome = 'Vazio (€/kWh)'
        if col_energia_v_nome in df_resultados_para_aggrid.columns:
            casas_decimais_energia = 4
    
            js_value_formatter_energia = JsCode(f"""
                function(params) {{
                    // ... (sua lógica de valueFormatter para Energia)
                    if (params.value == null || typeof params.value === 'undefined' || String(params.value).trim() === '') {{
                        return '';
                    }}
                    let num = Number(params.value);
                    if (isNaN(num)) {{ return ''; }}
                    try {{
                        return num.toFixed({casas_decimais_energia});
                    }} catch (e) {{
                        console.error("Erro valueFormatter Energia:", e, params.value);
                        return String(params.value);
                    }}
                }}
            """)

            gb.configure_column(
                field=col_energia_v_nome,
                headerName=col_energia_v_nome,
                type=["numericColumn"],
                filter=False,
                valueFormatter=js_value_formatter_energia,
                cellStyle=cell_style_cores_js,
                tooltipValueGetter=tooltip_preco_energia_js,
                tooltipComponent=custom_tooltip_component_js,
                minWidth=60, # Ajuste conforme necessário
                flex=1
            )

        # Configuração Coluna 'Preço Energia Fora Vazio (€/kWh)'
        col_energia_f_nome = 'Fora Vazio (€/kWh)'
        if col_energia_f_nome in df_resultados_para_aggrid.columns:
            casas_decimais_energia = 4 # Preço energia geralmente tem mais casas decimais
    
            js_value_formatter_energia = JsCode(f"""
                function(params) {{
                    // ... (sua lógica de valueFormatter para Energia)
                    if (params.value == null || typeof params.value === 'undefined' || String(params.value).trim() === '') {{
                        return '';
                    }}
                    let num = Number(params.value);
                    if (isNaN(num)) {{ return ''; }}
                    try {{
                        return num.toFixed({casas_decimais_energia});
                    }} catch (e) {{
                        console.error("Erro valueFormatter Energia:", e, params.value);
                        return String(params.value);
                    }}
                }}
            """)

            gb.configure_column(
                field=col_energia_f_nome,
                headerName=col_energia_f_nome,
                type=["numericColumn"],
                filter=False,
                valueFormatter=js_value_formatter_energia,
                cellStyle=cell_style_cores_js,
                tooltipValueGetter=tooltip_preco_energia_js,
                tooltipComponent=custom_tooltip_component_js,
                minWidth=60, # Ajuste conforme necessário
                flex=1
            )

        # Configuração Coluna 'Preço Energia Cheias (€/kWh)'
        col_energia_c_nome = 'Cheias (€/kWh)'
        if col_energia_c_nome in df_resultados_para_aggrid.columns:
            casas_decimais_energia = 4 # Preço energia geralmente tem mais casas decimais
    
            js_value_formatter_energia = JsCode(f"""
                function(params) {{
                    // ... (sua lógica de valueFormatter para Energia)
                    if (params.value == null || typeof params.value === 'undefined' || String(params.value).trim() === '') {{
                        return '';
                    }}
                    let num = Number(params.value);
                    if (isNaN(num)) {{ return ''; }}
                    try {{
                        return num.toFixed({casas_decimais_energia});
                    }} catch (e) {{
                        console.error("Erro valueFormatter Energia:", e, params.value);
                        return String(params.value);
                    }}
                }}
            """)

            gb.configure_column(
                field=col_energia_c_nome,
                headerName=col_energia_c_nome,
                type=["numericColumn"],
                filter=False,
                valueFormatter=js_value_formatter_energia,
                cellStyle=cell_style_cores_js,
                tooltipValueGetter=tooltip_preco_energia_js,
                tooltipComponent=custom_tooltip_component_js,
                minWidth=60, # Ajuste conforme necessário
                flex=1
            )

        # Configuração Coluna 'Preço Energia Ponta (€/kWh)'
        col_energia_p_nome = 'Ponta (€/kWh)'
        if col_energia_p_nome in df_resultados_para_aggrid.columns:
            casas_decimais_energia = 4
    
            js_value_formatter_energia = JsCode(f"""
                function(params) {{
                    // ... (sua lógica de valueFormatter para Energia)
                    if (params.value == null || typeof params.value === 'undefined' || String(params.value).trim() === '') {{
                        return '';
                    }}
                    let num = Number(params.value);
                    if (isNaN(num)) {{ return ''; }}
                    try {{
                        return num.toFixed({casas_decimais_energia});
                    }} catch (e) {{
                        console.error("Erro valueFormatter Energia:", e, params.value);
                        return String(params.value);
                    }}
                }}
            """)

            gb.configure_column(
                field=col_energia_p_nome,
                headerName=col_energia_p_nome,
                type=["numericColumn"],
                filter=False,
                valueFormatter=js_value_formatter_energia,
                cellStyle=cell_style_cores_js,
                tooltipValueGetter=tooltip_preco_energia_js,
                tooltipComponent=custom_tooltip_component_js,
                minWidth=60, # Ajuste conforme necessário
                flex=1
            )

        #Tooltip Preço Potencia
        tooltip_preco_potencia_js = JsCode("""
        function(params) {
            // params.value é o valor exibido na célula (Potência (€/dia) final sem IVA)
            // params.data contém todos os dados da linha
            if (!params.data) {
                // Se não houver dados da linha, retorna apenas o valor da célula como tooltip
                return String(params.value); 
            }

            // console.log para depurar os valores que chegam:
            console.log("Tooltip Potência Dados:", 
                params.data.tooltip_pot_comerc_sem_tar, 
                params.data.tooltip_pot_tar_bruta, 
                params.data.tooltip_pot_ts_aplicada, 
                params.data.tooltip_pot_desconto_ts_valor,
                params.value // Valor da célula
            );                                   


            // Aceder aos campos que adicionou em Python
            // Use (params.data.NOME_CAMPO || 0) para tratar casos onde o campo pode ser nulo/undefined
            const comercializador = parseFloat(params.data.tooltip_pot_comerc_sem_tar || 0);
            const tarBruta = parseFloat(params.data.tooltip_pot_tar_bruta || 0);
            const tsAplicada = params.data.tooltip_pot_ts_aplicada; 
            const descontoTSValor = parseFloat(params.data.tooltip_pot_desconto_ts_valor || 0);

            // Função helper para formatar números com 4 casas decimais
            const formatPrice = (num) => {
                if (typeof num === 'number' && !isNaN(num)) {
                    return num.toFixed(4);
                }
                return 'N/A'; // Ou algum outro placeholder
            };

            let tooltipParts = [];
            tooltipParts.push("<b>Decomposição Potência (s/IVA):</b>");
            tooltipParts.push("Comercializador (s/TAR): " + formatPrice(comercializador) + " €/dia");
            tooltipParts.push("TAR (Tarifa Acesso Redes): " + formatPrice(tarBruta) + " €/dia");

            if (tsAplicada === true && descontoTSValor > 0) { // Garantir que tsAplicada é explicitamente true
                tooltipParts.push("Desconto Tarifa Social: -" + formatPrice(descontoTSValor) + " €/dia");
            }
    
            tooltipParts.push("----------------------------------------------------");
            tooltipParts.push("<b>Custo Final : " + formatPrice(parseFloat(params.value)) + " €/dia</b>");

            var finalTooltipHtml = tooltipParts.join("<br>");
            // Log para ver o HTML final
            // console.log("Tooltip HTML Final para Potência:", finalTooltipHtml);
            return finalTooltipHtml;
        }
        """)

        # Exemplo para a coluna 'Preço Potência (€/dia)'
        col_potencia_nome = 'Potência (€/dia)'
        if col_potencia_nome in df_resultados_para_aggrid.columns and col_potencia_nome in cols_para_cor:
            casas_decimais_pot = 4
    
            js_value_formatter_potencia = JsCode(f"""
                function(params) {{
                    // ... (sua lógica de valueFormatter para potência)
                    if (params.value == null || typeof params.value === 'undefined' || String(params.value).trim() === '') {{
                        return '';
                    }}
                    let num = Number(params.value);
                    if (isNaN(num)) {{ return ''; }}
                    try {{
                        return num.toFixed({casas_decimais_pot});
                    }} catch (e) {{
                        console.error("Erro valueFormatter Potência:", e, params.value);
                        return String(params.value);
                    }}
                }}
            """)

            #Configuração Coluna Potência
            gb.configure_column(
                field=col_potencia_nome,
                headerName=col_potencia_nome,
                type=["numericColumn"],
                filter=False,
                valueFormatter=js_value_formatter_potencia,
                cellStyle=cell_style_cores_js,
                tooltipValueGetter=tooltip_preco_potencia_js,
                tooltipComponent=custom_tooltip_component_js,
                # tooltipComponentParams={'color': '#aabbcc'}
            )

        #Tooltip Custo Total
        tooltip_custo_total_js = JsCode("""
        function(params) {
            if (!params.data) { return String(params.value); }

            const formatCurrency = (num) => {
                if (typeof num === 'number' && !isNaN(num)) {
                    return num.toFixed(2); // 2 casas decimais para custos
                }
                return 'N/A';
            };

            const nomeTarifario = params.data.NomeParaExibir || "Tarifário"; // Fallback se não existir

            // Linha de título do tooltip atualizada para incluir o nome do tarifário
            let tooltipParts = [
                "<i>" + nomeTarifario + "</i>", // Nome do tarifário em negrito na primeira linha
                "<b>Decomposição Custo Total:</b>" // Título em itálico na segunda linha
                // Pode adicionar uma linha em branco se quiser mais espaçamento: ""
            ];
            tooltipParts.push("------------------------------------");
                                        
            const energia_siva = parseFloat(params.data.tt_cte_energia_siva || 0);
            const potencia_siva = parseFloat(params.data.tt_cte_potencia_siva || 0);
            const iec_siva = parseFloat(params.data.tt_cte_iec_siva || 0);
            const dgeg_siva = parseFloat(params.data.tt_cte_dgeg_siva || 0);
            const cav_siva = parseFloat(params.data.tt_cte_cav_siva || 0);
            const total_siva = parseFloat(params.data.tt_cte_total_siva || 0);
            const iva_6 = parseFloat(params.data.tt_cte_valor_iva_6_total || 0);
            const iva_23 = parseFloat(params.data.tt_cte_valor_iva_23_total || 0);
            const subtotal_civa_antes_desc_acr = parseFloat(params.data.tt_cte_subtotal_civa || 0);
            const desc_finais_valor = parseFloat(params.data.tt_cte_desc_finais_valor || 0);
            const acres_finais_valor = parseFloat(params.data.tt_cte_acres_finais_valor || 0);

            const custo_total_celula = parseFloat(params.value);

            tooltipParts.push("Total Energia s/IVA: " + formatCurrency(energia_siva) + " €");
            tooltipParts.push("Total Potência s/IVA: " + formatCurrency(potencia_siva) + " €");
            // Condição para mostrar IEC: se o valor for maior que zero OU se a tarifa social não estiver ativa.
            // Precisa adicionar 'tarifa_social' (a flag booleana global) aos dados da linha se quiser esta condição precisa.
            // Por agora, vou simplificar para mostrar se iec_siva > 0.
            if (iec_siva !== 0) { 
                tooltipParts.push("IEC s/IVA: " + formatCurrency(iec_siva) + " €");
            }
            if (dgeg_siva !== 0) {
                tooltipParts.push("DGEG s/IVA: " + formatCurrency(dgeg_siva) + " €");
            }
            if (cav_siva !== 0) {
                tooltipParts.push("CAV s/IVA: " + formatCurrency(cav_siva) + " €");
            }
            tooltipParts.push("<b>Subtotal s/IVA: " + formatCurrency(total_siva) + " €</b>");
            tooltipParts.push("------------------------------------");
            if (iva_6 !== 0) {
                tooltipParts.push("Valor IVA (6%): " + formatCurrency(iva_6) + " €");
            }
            if (iva_23 !== 0) {
                tooltipParts.push("Valor IVA (23%): " + formatCurrency(iva_23) + " €");
            }
            tooltipParts.push("<b>Subtotal c/IVA: " + formatCurrency(subtotal_civa_antes_desc_acr) + " €</b>");
            tooltipParts.push("------------------------------------");
    
            // Mostrar descontos e acréscimos apenas se existirem
            if (desc_finais_valor !== 0) {
                tooltipParts.push("Outros Descontos: -" + formatCurrency(desc_finais_valor) + " €");
            }
            if (acres_finais_valor !== 0) {
                tooltipParts.push("Outros Acréscimos: +" + formatCurrency(acres_finais_valor) + " €");
            }
    
            if (desc_finais_valor !== 0 || acres_finais_valor !== 0) {
                tooltipParts.push("------------------------------------");
            }
            tooltipParts.push("<b>Custo Total c/IVA: " + formatCurrency(custo_total_celula) + " €</b>");

            return tooltipParts.join("<br>");
        }
        """)

        # Configuração Coluna 'Custo Total (€)'
        col_custo_total_nome = 'Total (€)'
        if col_custo_total_nome in df_resultados_para_aggrid.columns:
            casas_decimais_total = 2
    
            js_value_formatter_energia = JsCode(f"""
                function(params) {{
                    // ... (sua lógica de valueFormatter para Energia)
                    if (params.value == null || typeof params.value === 'undefined' || String(params.value).trim() === '') {{
                        return '';
                    }}
                    let num = Number(params.value);
                    if (isNaN(num)) {{ return ''; }}
                    try {{
                        return num.toFixed({casas_decimais_total});
                    }} catch (e) {{
                        console.error("Erro valueFormatter Energia:", e, params.value);
                        return String(params.value);
                    }}
                }}
            """)

            gb.configure_column(
                field=col_custo_total_nome,
                headerName=col_custo_total_nome,
                type=["numericColumn"],
                filter=False,
                valueFormatter=js_value_formatter_energia,
                cellStyle=cell_style_cores_js,
                tooltipValueGetter=tooltip_custo_total_js, 
                tooltipComponent=custom_tooltip_component_js, 
                minWidth=80, # Ajuste conforme necessário
                flex=1
            )

        # --- Configuração de Colunas com Set Filter ---
        set_filter_params = {
            'buttons': ['apply', 'reset'],
            'excelMode': 'mac',
            'suppressMiniFilter': False, # Garante que a caixa de pesquisa dentro do Set Filter aparece
        }

        is_visible = col_name in colunas_visiveis_presentes

        text_columns_with_set_filter = ['Tipo', 'Segmento', 'Faturação', 'Pagamento']
        for col_name in text_columns_with_set_filter:
            if col_name in df_resultados_para_aggrid.columns:
                min_w = 100 if col_name == 'Tipo' else 120
                fx = 0.5 if col_name == 'Tipo' else 0.75
                gb.configure_column(
                    col_name, 
                    headerName=col_name,
                    minWidth=min_w,
                    flex=fx,
                    filter='agSetColumnFilter',
                    filterParams=set_filter_params,
                    cellStyle={
                        'textAlign': 'center', 
                        'borderRadius': '6px',
                        'padding': '10px 10px',
                        'backgroundColor': '#f0f0f0' 
                    },
                    hide=(not is_visible) 

                )
        is_visible_comerc = 'Comercializador' in colunas_visiveis_presentes

        # --- Configuração Coluna Comercializador (Text Filter) ---
        if 'Comercializador' in df_resultados_para_aggrid.columns:
            gb.configure_column(
                "Comercializador",
                headerName="Comercializador",
                minWidth=50,
                flex=1,
                filter='agTextColumnFilter',
                cellStyle={
                    'textAlign': 'center', 
                    'borderRadius': '6px', 
                    'padding': '10px 10px', 
                    'backgroundColor': '#f0f0f0'
                },
                hide=(not is_visible_comerc)
            )

        # Configurar outras colunas (Tipo, Comercializador e colunas de dados)
        for col_nome_num in df_resultados_para_aggrid.columns:
            if col_nome_num in cols_para_cor:
                casas_decimais = 4 if '€/kWh' in col_nome_num or '€/dia' in col_nome_num else 2
        
                # Definir o JsCode para valueFormatter DENTRO do loop
                # para que capture o valor correto de 'casas_decimais' para esta coluna específica.
                js_value_formatter_para_coluna = JsCode(f"""
                    function(params) {{
                        // Verificar se o valor é nulo ou não é um número
                        if (params.value == null || typeof params.value === 'undefined' || String(params.value).trim() === '') {{
                            return ''; // Retornar string vazia para valores nulos, indefinidos ou vazios
                        }}
                
                        let num = Number(params.value); // Tentar converter para número
                
                        if (isNaN(num)) {{
                            // Se após a tentativa de conversão ainda for NaN, retornar string vazia
                            // ou pode optar por retornar o valor original se fizer sentido: return params.value;
                            return ''; 
                        }}
                
                        // Se for um número válido, formatá-lo com o número correto de casas decimais
                        try {{
                            return num.toFixed({casas_decimais});
                        }} catch (e) {{
                            console.error("Erro ao formatar o número no valueFormatter:", e, 
                                            "Valor original:", params.value, 
                                            "Casas decimais:", {casas_decimais});
                            return String(params.value); // Fallback para string do valor original em caso de erro inesperado
                        }}
                    }}
                """)

                gb.configure_column(
                    field=col_nome_num,
                    headerName=col_nome_num,
                    type=["numericColumn", "numberColumnFilter"], 
                    valueFormatter=js_value_formatter_para_coluna,
                    cellStyle=cell_style_cores_js,
                    minWidth=50,
                    flex=1
                )

        # --- 3. Formatação para "O Meu Tarifário" ---
        get_row_style_meu_tarifario_js = JsCode("""
        function(params) {
            // Verifica se params.data existe e se NomeParaExibir é uma string que COMEÇA COM "O Meu Tarifário"
            if (params.data && typeof params.data.NomeParaExibir === 'string' && params.data.NomeParaExibir.startsWith('O Meu Tarifário')) {
                return { fontWeight: 'bold' }; // Aplica negrito a toda a linha
            }
            return null; // Sem estilo especial para outras linhas
        }
        """)

        # Ocultar 'Tipo' e 'LinkAdesao' na vista simplificada se não estiverem em colunas_visiveis_presentes
        if vista_simplificada:
            if 'Tipo' not in colunas_visiveis_presentes and 'Tipo' in df_resultados_para_aggrid.columns:
                gb.configure_column(field='Tipo', hide=True)
            if 'LinkAdesao' not in colunas_visiveis_presentes and 'LinkAdesao' in df_resultados_para_aggrid.columns:
                gb.configure_column(field='LinkAdesao', hide=True)
            # Oculte outras colunas que estão nos dados mas não são visíveis na vista simplificada
            colunas_desktop_a_ocultar_na_vista_movel = ['Segmento', 'Faturação', 'Pagamento', 'Comercializador']
            for col_ocultar in colunas_desktop_a_ocultar_na_vista_movel:
                if col_ocultar not in colunas_visiveis_presentes and col_ocultar in df_resultados_para_aggrid.columns:
                     gb.configure_column(field=col_ocultar, hide=True)


        # Ocultar colunas de dados de tooltip
        colunas_de_dados_tooltip_a_ocultar = [
            'info_notas', 
            'tooltip_pot_comerc_sem_tar', 'tooltip_pot_tar_bruta', 'tooltip_pot_ts_aplicada', 'tooltip_pot_desconto_ts_valor',
            # Energia Simples (S)
            'tooltip_energia_S_comerc_sem_tar', 'tooltip_energia_S_tar_bruta', 
            'tooltip_energia_S_tse_declarado_incluido', 'tooltip_energia_S_tse_valor_nominal',
            'tooltip_energia_S_ts_aplicada_flag', 'tooltip_energia_S_ts_desconto_valor',
    
            # Energia Vazio (V) - Adicione se tiver opção Bi ou Tri
            'tooltip_energia_V_comerc_sem_tar', 'tooltip_energia_V_tar_bruta', 
            'tooltip_energia_V_tse_declarado_incluido', 'tooltip_energia_V_tse_valor_nominal',
            'tooltip_energia_V_ts_aplicada_flag', 'tooltip_energia_V_ts_desconto_valor',

            # Energia Fora Vazio (F) - Adicione se tiver opção Bi
            'tooltip_energia_F_comerc_sem_tar', 'tooltip_energia_F_tar_bruta', 
            'tooltip_energia_F_tse_declarado_incluido', 'tooltip_energia_F_tse_valor_nominal',
            'tooltip_energia_F_ts_aplicada_flag', 'tooltip_energia_F_ts_desconto_valor',

            # Energia Cheias (C) - Adicione se tiver opção Tri
            'tooltip_energia_C_comerc_sem_tar', 'tooltip_energia_C_tar_bruta', 
            'tooltip_energia_C_tse_declarado_incluido', 'tooltip_energia_C_tse_valor_nominal',
            'tooltip_energia_C_ts_aplicada_flag', 'tooltip_energia_C_ts_desconto_valor',

            # Energia Ponta (P) - Adicione se tiver opção Tri
            'tooltip_energia_P_comerc_sem_tar', 'tooltip_energia_P_tar_bruta', 
            'tooltip_energia_P_tse_declarado_incluido', 'tooltip_energia_P_tse_valor_nominal',
            'tooltip_energia_P_ts_aplicada_flag', 'tooltip_energia_P_ts_desconto_valor',

            # Para Custo Total
            'tt_cte_energia_siva', 'tt_cte_potencia_siva', 'tt_cte_iec_siva',
            'tt_cte_dgeg_siva', 'tt_cte_cav_siva', 'tt_cte_total_siva',
            'tt_cte_valor_iva_6_total', 'tt_cte_valor_iva_23_total',
            'tt_cte_subtotal_civa','tt_cte_desc_finais_valor','tt_cte_acres_finais_valor'
        ]

        for col_para_ocultar in colunas_de_dados_tooltip_a_ocultar:
            if col_para_ocultar in df_resultados_para_aggrid.columns:
                gb.configure_column(field=col_para_ocultar, hide=True)

        gb.configure_grid_options(
            domLayout='autoHeight', # Para altura automática
            getRowStyle=get_row_style_meu_tarifario_js,
            suppressContextMenu=True  # Adicionado para desativar o menu de contexto
        )

        gb.configure_default_column(headerClass='center-header')
        gridOptions = gb.build()
        custom_css = {
            ".ag-header-cell-label": {
                "justify-content": "center !important",
                "text-align": "center !important",
                "font-size": "14px !important",      # <-- aumenta o header
                "font-weight": "bold !important"
            },
            ".ag-center-header": {
                "justify-content": "center !important",
                "text-align": "center !important",
                "font-size": "14px !important"       # <-- reforço para headers
            },
            ".ag-cell": {
                "font-size": "14px !important"       # <-- aumenta valores das células
            },
            ".ag-center-cols-clip": {"justify-content": "center !important", "text-align": "center !important"}
        }

        # --- Construir a chave dinâmica para a AgGrid ---
        # Inclua todos os inputs cujo estado deve levar a um "reset" da AgGrid
        # (filtros manuais, ordenação manual na grelha são resetados quando a key muda)
        key_parts_para_aggrid = [
            str(potencia), str(opcao_horaria), str(mes),
            str(data_inicio), str(data_fim), str(dias),
            # Consumos
            str(consumo_simples), str(consumo_vazio), str(consumo_fora_vazio),
            str(consumo_cheias), str(consumo_ponta),
            # Opções adicionais (garanta que estas variáveis estão definidas no escopo)
            str(tarifa_social), str(familia_numerosa), str(comparar_indexados),
            str(valor_dgeg_user), str(valor_cav_user),
            str(incluir_quota_acp), str(desconto_continente),
            # Inputs OMIE manuais (usar os valores do session_state que alimentam os number_input)
            str(st.session_state.get('omie_s_input_field')),
           str(st.session_state.get('omie_v_input_field')),
            str(st.session_state.get('omie_f_input_field')),
            str(st.session_state.get('omie_c_input_field')),
            str(st.session_state.get('omie_p_input_field')),
            # Estado e dados do "Meu Tarifário"
            str(meu_tarifario_ativo)
        ]
        if meu_tarifario_ativo and 'meu_tarifario_calculado' in st.session_state:
           # Adicionar uma representação do "Meu Tarifário" à chave,
            # por exemplo, o seu custo total, para que a AgGrid também resete se ele mudar.
            # Isto é útil se a ordenação da grelha depender deste valor.
            key_parts_para_aggrid.append(str(st.session_state['meu_tarifario_calculado'].get('Total (€)', '')))

        # Juntar todas as partes para formar a chave única
        aggrid_dinamica_key = "aggrid_principal_" + "_".join(key_parts_para_aggrid)


        # Exibir a Grelha
        grid_response = AgGrid(
            df_resultados_para_aggrid,
            gridOptions=gridOptions,
            custom_css=custom_css,
            # update_mode diz ao Streamlit para atualizar os dados quando os filtros ou a ordenação mudam na AgGrid
            update_mode=GridUpdateMode.FILTERING_CHANGED | GridUpdateMode.SORTING_CHANGED | GridUpdateMode.SELECTION_CHANGED,
            allow_unsafe_jscode=True,
            fit_columns_on_grid_load=True,
            theme='alpine', 
            key=aggrid_dinamica_key, # Uma key para a instância interativa
            enable_enterprise_modules=True,
            # reload_data=True # Considere usar se os dados de entrada (df_resultados_para_aggrid) puderem mudar dinamicamente por outras interações
        )
        # ---- FIM DA CONFIGURAÇÃO DO AGGRID ----

    st.markdown("<a id='exportar-excel-detalhada'></a>", unsafe_allow_html=True)
    st.markdown("---")
    with st.expander("📥 Exportar Tabela Detalhada para Excel"):
        colunas_dados_tooltip_a_ocultar = [
            'info_notas', 'LinkAdesao',
            'tooltip_pot_comerc_sem_tar', 'tooltip_pot_tar_bruta', # ... e todas as outras ...
            'tt_cte_subtotal_civa','tt_cte_desc_finais_valor','tt_cte_acres_finais_valor'
        ]

        if 'df_resultados_para_aggrid' in locals() and \
           isinstance(df_resultados_para_aggrid, pd.DataFrame) and \
           not df_resultados_para_aggrid.empty and \
           'colunas_visiveis_presentes' in locals() and \
           isinstance(colunas_visiveis_presentes, list) and \
           'colunas_dados_tooltip_a_ocultar' in locals() and \
           isinstance(colunas_dados_tooltip_a_ocultar, list):

        # Colunas disponíveis para seleção:
        # Começamos com todas as colunas que estão no DataFrame que alimenta o AgGrid.
            todas_as_colunas_no_df_aggrid = df_resultados_para_aggrid.columns.tolist()
        
        # Organizar as opções para o multiselect:
        # 1. Colunas visíveis primeiro
        # 2. Depois, colunas de tooltip (que não estão já nas visíveis)
        # 3. Depois, outras colunas (se houver e fizer sentido oferecer)
        
            opcoes_export_excel = []
        # Adicionar colunas visíveis primeiro, mantendo a sua ordem
            for col_vis in colunas_visiveis_presentes:
                if col_vis in todas_as_colunas_no_df_aggrid and col_vis not in opcoes_export_excel:
                    opcoes_export_excel.append(col_vis)
        
        # Adicionar colunas de dados de tooltip que não estão já nas visíveis
        # e que existem no df_resultados_para_aggrid
            for col_tooltip in colunas_dados_tooltip_a_ocultar: # Esta lista contém os nomes das colunas de tooltip
                if col_tooltip in todas_as_colunas_no_df_aggrid and col_tooltip not in opcoes_export_excel:
                    opcoes_export_excel.append(col_tooltip)
        
        # Adicionar quaisquer outras colunas restantes do df_resultados_para_aggrid se desejado
        # (excluindo as que já foram adicionadas)
        # Se 'colunas_para_aggrid_final' foi usado para criar df_resultados_para_aggrid,
        # ele pode já ser uma boa base, mas vamos usar todas_as_colunas_no_df_aggrid para garantir
            for col_restante in todas_as_colunas_no_df_aggrid:
                if col_restante not in opcoes_export_excel:
                    opcoes_export_excel.append(col_restante)

        # Colunas pré-selecionadas: apenas as que estão atualmente visíveis no AgGrid
            default_cols_excel = [col for col in colunas_visiveis_presentes if col in opcoes_export_excel]
            
            if not default_cols_excel and 'NomeParaExibir' in opcoes_export_excel:
                default_cols_excel.append('NomeParaExibir')
            if not default_cols_excel and 'Total (€)' in opcoes_export_excel:
                default_cols_excel.append('Total (€)')


            colunas_para_exportar_excel_selecionadas = st.multiselect(
                "Selecione as colunas para exportar para Excel:",
                options=opcoes_export_excel, 
                default=default_cols_excel,
                key="cols_export_excel_selector_dados_com_tooltips" # Nova key
            )
            
            def exportar_excel_completo(df_para_exportar, styler_obj, resumo_html_para_excel, poupanca_texto_para_excel, identificador_cor_cabecalho):
                output_excel_buffer = io.BytesIO() 
                with pd.ExcelWriter(output_excel_buffer, engine='openpyxl') as writer_excel:
                    sheet_name_excel = 'Tiago Felicia - Eletricidade'

                    # --- Escrever Resumo ---
                    dados_resumo_formatado = []
                    if resumo_html_para_excel:
                        soup_resumo = BeautifulSoup(resumo_html_para_excel, "html.parser")
                        
                        # --- LÓGICA ALTERADA PARA REARRANJAR O RESUMO ---
                        titulo_resumo = soup_resumo.find('h5')
                        if titulo_resumo:
                            dados_resumo_formatado.append([titulo_resumo.get_text(strip=True), None])

                        itens_lista_resumo = soup_resumo.find_all('li')
                        linha_filtros_texto = ""
                        linha_potencia_texto = ""
                        outras_linhas_resumo = []

                        for item in itens_lista_resumo:
                            # Usar .get_text() para obter o conteúdo limpo do item da lista
                            texto_item = item.get_text(separator=' ', strip=True)
                                
                            if "Segmento:" in texto_item:
                                linha_filtros_texto = texto_item
                            elif "kVA" in texto_item:
                                linha_potencia_texto = texto_item
                            else:
                                # Processar as outras linhas normalmente
                                parts = texto_item.split(':', 1)
                                if len(parts) == 2:
                                    outras_linhas_resumo.append([parts[0].strip() + ":", parts[1].strip()])
                                else:
                                    outras_linhas_resumo.append([texto_item, None])
                            
                        # Adicionar a linha combinada primeiro, na ordem que pediu
                        if linha_filtros_texto or linha_potencia_texto:
                            dados_resumo_formatado.append([linha_filtros_texto, linha_potencia_texto])
                        
                        # Adicionar o resto do resumo
                        dados_resumo_formatado.extend(outras_linhas_resumo)
                        # --- FIM DA LÓGICA ALTERADA ---
        
                    df_resumo_obj = pd.DataFrame(dados_resumo_formatado)

                    # 1. Deixe o Pandas criar/ativar a folha na primeira escrita
                    df_resumo_obj.to_excel(writer_excel, sheet_name=sheet_name_excel, index=False, header=False, startrow=0)

                    # 2. AGORA obtenha a referência à worksheet, que certamente existe
                    worksheet_excel = writer_excel.sheets[sheet_name_excel]

                    # Formatar Resumo (Negrito)
                    bold_font_obj = Font(bold=True) # Font já deve estar importado de openpyxl.styles
                    for i_resumo in range(len(df_resumo_obj)):
                        excel_row_idx_resumo = i_resumo + 1 # Linhas do Excel são 1-based
                        cell_resumo_rotulo = worksheet_excel.cell(row=excel_row_idx_resumo, column=1)
                        cell_resumo_rotulo.font = bold_font_obj
                        if df_resumo_obj.shape[1] > 1 and pd.notna(df_resumo_obj.iloc[i_resumo, 1]):
                            cell_resumo_valor = worksheet_excel.cell(row=excel_row_idx_resumo, column=2)
                            cell_resumo_valor.font = bold_font_obj
        
                    worksheet_excel.column_dimensions['A'].width = 35
                    worksheet_excel.column_dimensions['B'].width = 65

                    linha_atual_no_excel_escrita = len(df_resumo_obj) + 1

                    # --- Escrever Mensagem de Poupança ---
                    if poupanca_texto_para_excel: # Verifica se há texto para a mensagem de poupança
                        linha_atual_no_excel_escrita += 1 # Adiciona uma linha em branco
            
                        cor_p = st.session_state.get('poupanca_excel_cor', "000000") # Cor do session_state
                        negrito_p = st.session_state.get('poupanca_excel_negrito', False) # Negrito do session_state
            
                        poupanca_cell_escrita = worksheet_excel.cell(row=linha_atual_no_excel_escrita, column=1, value=poupanca_texto_para_excel)
                        poupanca_font_escrita = Font(bold=negrito_p, color=cor_p)
                        poupanca_cell_escrita.font = poupanca_font_escrita

                        # --- MODIFICAÇÃO PARA JUNTAR CÉLULAS ---
                        worksheet_excel.merge_cells(start_row=linha_atual_no_excel_escrita, start_column=1, end_row=linha_atual_no_excel_escrita, end_column=4)

                        # Aplicar alinhamento à célula mesclada (a célula do canto superior esquerdo, poupanca_cell_escrita)
                        poupanca_cell_escrita.alignment = Alignment(wrap_text=True, horizontal='left', vertical='top')
                
                        linha_atual_no_excel_escrita += 1 # Avança para a próxima linha após a mensagem de poupança
        
                    linha_inicio_tab_dados_excel = linha_atual_no_excel_escrita + 3

                    # --- Adicionar linha de informação da simulação ---
                    # Adiciona uma linha em branco antes desta nova linha de informação
                    linha_info_simulacao_excel = linha_atual_no_excel_escrita + 1 

                    data_hoje_obj = datetime.date.today() # datetime já deve estar importado
                    data_hoje_formatada_str = data_hoje_obj.strftime('%d/%m/%Y')
            
                    espacador_info = "                                                                      " # Exemplo: 70 espaços

                    texto_completo_info = (
                        f"          Simulação em {data_hoje_formatada_str}{espacador_info}"
                        f"https://www.tiagofelicia.pt{espacador_info}"
                        f"Tiago Felícia"
                    )

                    # Escrever o texto completo na primeira célula da área a ser mesclada (Coluna A)
                    info_cell = worksheet_excel.cell(row=linha_info_simulacao_excel, column=1)
                    info_cell.value = texto_completo_info
            
                    # Aplicar negrito à célula
                    # Reutilizar bold_font_obj que já foi definido para o resumo, ou criar um novo se precisar de formatação diferente.
                    # Assumindo que bold_font_obj é Font(bold=True) e está no escopo.
                    # Se não, defina-o: from openpyxl.styles import Font; bold_font_obj = Font(bold=True)
                    if 'bold_font_obj' in locals() or 'bold_font_obj' in globals():
                         info_cell.font = bold_font_obj # Reutiliza o bold_font_obj do resumo
                    else:
                         info_cell.font = Font(bold=True) # Cria um novo se não existir

                    # Mesclar as colunas A, B, C, e D para esta linha
                    worksheet_excel.merge_cells(start_row=linha_info_simulacao_excel, start_column=1, end_row=linha_info_simulacao_excel, end_column=4)
            
                    # Ajustar alinhamento da célula mesclada (info_cell é a célula do topo-esquerda da área mesclada)
                    # Alinhado à esquerda, centralizado verticalmente, com quebra de linha se necessário.
                    info_cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True) 

                    # A linha de início para a tabela de dados principal virá depois desta linha de informação
                    # Adicionamos +1 para esta linha de informação e +1 para uma linha em branco antes da tabela
                    linha_inicio_tab_dados = linha_info_simulacao_excel + 2 
            
                    # --- Fim da adição da linha de informação ---

                    for row in worksheet_excel.iter_rows(min_row=1, max_row=worksheet_excel.max_row+100, min_col=1, max_col=20):
                        for cell in row:
                            cell.fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

                    # O Styler escreverá na mesma folha 'sheet_name_excel'
                    styler_obj.to_excel(
                        writer_excel,
                        sheet_name=sheet_name_excel,
                        index=False,
                        startrow=linha_inicio_tab_dados_excel - 1, # startrow é 0-indexed
                        columns=df_para_exportar.columns.tolist()
                    )

                    # Determina cor do cabeçalho conforme a opção horária e ciclo
                    opcao_horaria_lower = str(opcao_horaria).lower() if 'opcao_horaria' in locals() else "simples"
                    cor_fundo = "A6A6A6"   # padrão Simples
                    cor_fonte = "000000"    # padrão preto

                    if isinstance(identificador_cor_cabecalho, str):
                        id_lower = identificador_cor_cabecalho.lower()
                        if id_lower == "simples": # Já coberto pelo default
                            pass
                        elif "bi-horário" in id_lower and "diário" in id_lower:
                            cor_fundo = "A9D08E"; cor_fonte = "000000"
                        elif "bi-horário" in id_lower and "semanal" in id_lower:
                            cor_fundo = "8EA9DB"; cor_fonte = "000000"
                        elif "tri-horário" in id_lower and "diário" in id_lower:
                            cor_fundo = "BF8F00"; cor_fonte = "FFFFFF"
                        elif "tri-horário" in id_lower and "semanal" in id_lower:
                            cor_fundo = "C65911"; cor_fonte = "FFFFFF"

                    # linha_inicio_tab_dados_excel já existe na função e corresponde ao header da tabela
                    for col_idx, _ in enumerate(df_para_exportar.columns):
                        celula = worksheet_excel.cell(row=linha_inicio_tab_dados_excel, column=col_idx + 1)
                        celula.fill = PatternFill(start_color=cor_fundo, end_color=cor_fundo, fill_type="solid")
                        celula.font = Font(color=cor_fonte, bold=True)

                                # ---- INÍCIO: ADICIONAR LEGENDA DE CORES APÓS A TABELA ----
                    # Calcular a linha de início para a legenda
                    # df_para_exportar é o DataFrame que foi escrito pelo styler (ex: df_export_final)
                    numero_linhas_dados_tabela_principal = len(df_para_exportar)
                    # linha_inicio_tab_dados_excel é a linha do cabeçalho da tabela principal (1-indexada)
                    ultima_linha_tabela_principal = linha_inicio_tab_dados_excel + numero_linhas_dados_tabela_principal
                    
                    linha_legenda_bloco_inicio = ultima_linha_tabela_principal + 2 # Deixa uma linha em branco após a tabela

                    # Título da Legenda
                    titulo_legenda_cell = worksheet_excel.cell(row=linha_legenda_bloco_inicio, column=1, value="Tipos de Tarifário:")
                    # Reutilizar bold_font_obj se definido ou criar um novo
                    if 'bold_font_obj' in locals() or 'bold_font_obj' in globals():
                        titulo_legenda_cell.font = bold_font_obj
                    else:
                        titulo_legenda_cell.font = Font(bold=True)
                    worksheet_excel.merge_cells(start_row=linha_legenda_bloco_inicio, start_column=1, end_row=linha_legenda_bloco_inicio, end_column=4) # Mesclar para o título
                    
                    linha_legenda_item_atual = linha_legenda_bloco_inicio + 1 # Primeira linha para item da legenda

                    titulo_legenda_cell.alignment = Alignment(horizontal='center', vertical='center')

                    itens_legenda_excel = [
                        {"cf": "FF0000", "ct": "FFFFFF", "b": True, "tA": "O Meu Tarifário", "tB": "Tarifário configurado pelo utilizador."},
                        {"cf": "FFE699", "ct": "000000", "b": False, "tA": "Indexado (Média)", "tB": "Preço de energia baseado na média OMIE do período."},
                        {"cf": "4D79BC", "ct": "FFFFFF", "b": False, "tA": "Indexado (Quarto-Horário)", "tB": "Preço de energia baseado nos valores OMIE horários/quarto-horários e perfil."},
                        {"cf": "F0F0F0", "ct": "333333", "b": False, "tA": "Fixo", "tB": "Preços de energia constantes", "borda_cor": "CCCCCC"}
                    ]
                    
                    # Definir larguras das colunas para a legenda (pode ajustar conforme necessário)
                    worksheet_excel.column_dimensions[get_column_letter(1)].width = 30 # Coluna A para a amostra/nome
                    worksheet_excel.column_dimensions[get_column_letter(2)].width = 70 # Coluna B para a descrição (será junta)

                    for item in itens_legenda_excel:
                        celula_A_legenda = worksheet_excel.cell(row=linha_legenda_item_atual, column=1, value=item["tA"])
                        celula_A_legenda.fill = PatternFill(start_color=item["cf"], end_color=item["cf"], fill_type="solid")
                        celula_A_legenda.font = Font(color=item["ct"], bold=item["b"])
                        celula_A_legenda.alignment = Alignment(horizontal='center', vertical='center', indent=1)

                        if "borda_cor" in item:
                            cor_borda_hex = item["borda_cor"]
                            borda_legenda_obj = Border(
                                top=Side(border_style="thin", color=cor_borda_hex),
                                left=Side(border_style="thin", color=cor_borda_hex),
                                right=Side(border_style="thin", color=cor_borda_hex),
                                bottom=Side(border_style="thin", color=cor_borda_hex)
                            )
                            celula_A_legenda.border = borda_legenda_obj
                        
                        celula_B_legenda = worksheet_excel.cell(row=linha_legenda_item_atual, column=2, value=item["tB"])
                        celula_B_legenda.alignment = Alignment(vertical='center', wrap_text=True, horizontal='left')
                        # Mesclar colunas B até D (ou ajuste conforme a largura desejada para a descrição)
                        worksheet_excel.merge_cells(start_row=linha_legenda_item_atual, start_column=2,
                                                    end_row=linha_legenda_item_atual, end_column=4) 
                        
                        worksheet_excel.row_dimensions[linha_legenda_item_atual].height = 20 # Ajustar altura da linha da legenda
                        linha_legenda_item_atual += 1
                    # ---- FIM: ADICIONAR LEGENDA DE CORES ----

                    # Ajustar largura das colunas da tabela principal
                    for col_idx_iter, col_nome_iter_width in enumerate(df_para_exportar.columns):
                        col_letra_iter = get_column_letter(col_idx_iter + 1) # get_column_letter já deve estar importado
                        if "Tarifário" in col_nome_iter_width :
                             worksheet_excel.column_dimensions[col_letra_iter].width = 80    
                        elif "Total (€)" == col_nome_iter_width :
                            worksheet_excel.column_dimensions[col_letra_iter].width = 25
                        elif "(€/kWh)" in col_nome_iter_width or "(€/dia)" in col_nome_iter_width:
                            worksheet_excel.column_dimensions[col_letra_iter].width = 25
                        elif "Comercializador" in col_nome_iter_width :
                             worksheet_excel.column_dimensions[col_letra_iter].width = 30    
                        elif "Faturação" in col_nome_iter_width :
                             worksheet_excel.column_dimensions[col_letra_iter].width = 33    
                        elif "Pagamento" in col_nome_iter_width :
                             worksheet_excel.column_dimensions[col_letra_iter].width = 50    
                        else: 
                            worksheet_excel.column_dimensions[col_letra_iter].width = 25


                output_excel_buffer.seek(0)
                return output_excel_buffer
                    # --- Fim da definição de exportar_excel_completo ---

            # --- Início do Bloco Numero Tarifários exportados ---
            opcoes_limite_export = ["Todos"] + [f"Top {i}" for i in [10, 20, 30, 40, 50]]
            limite_export_selecionado = st.selectbox(
                "Número de tarifários a exportar (ordenados pelo 'Total (€)' atual da tabela):",
                options=opcoes_limite_export,
                index=0, # "Todos" como padrão
                key="limite_tarifarios_export_excel"
            )
            # --- Fim do Bloco Numero Tarifários exportados ---

            # --- Bloco Preparar Download ---
            if st.button("Preparar Download do Ficheiro Excel (Dados Selecionados)", key="btn_prep_excel_download_dados_com_tooltips_corrigido"):
                if not colunas_para_exportar_excel_selecionadas:
                    st.warning("Por favor, selecione pelo menos uma coluna para exportar.")
                else:
                    # O código de geração do Excel e o st.download_button VÊM AQUI DENTRO
                    with st.spinner("A gerar ficheiro Excel..."):
                        #df_export_final = df_resultados_para_aggrid[colunas_para_exportar_excel_selecionadas].copy()
                        if grid_response and grid_response['data'] is not None: # Verifica se grid_response e os dados existem
                        # grid_response['data'] contém os dados filtrados e ordenados da AgGrid como uma lista de dicionários
                            df_dados_filtrados_da_grid = pd.DataFrame(grid_response['data'])


                        if df_dados_filtrados_da_grid.empty and not df_resultados_para_aggrid.empty:
                            # Isto pode acontecer se os filtros resultarem numa tabela vazia
                            st.warning("Os filtros aplicados resultaram numa tabela vazia. A exportar um ficheiro vazio ou com cabeçalhos apenas.")
                            # Decide o que fazer: exportar ficheiro vazio ou parar.
                            # Para exportar ficheiro vazio com cabeçalhos:
                            df_export_final = pd.DataFrame(columns=colunas_para_exportar_excel_selecionadas)

                        elif not df_dados_filtrados_da_grid.empty:
                            # Assegurar que apenas as colunas selecionadas pelo utilizador para exportação são usadas,
                            # e que estas colunas existem no df_dados_filtrados_da_grid.
                            colunas_export_validas_no_filtrado = [
                                col for col in colunas_para_exportar_excel_selecionadas 
                                if col in df_dados_filtrados_da_grid.columns
                            ]
                            if not colunas_export_validas_no_filtrado:
                                st.warning("Nenhuma das colunas selecionadas para exportação está presente nos dados filtrados atuais da tabela.")
                    
                            df_export_final = df_dados_filtrados_da_grid[colunas_export_validas_no_filtrado].copy()
                        else: # Se grid_response['data'] for None ou vazio e df_resultados_para_aggrid também era vazio
                            st.warning("Não há dados na tabela para exportar.")

                        # Aplicar limite de tarifários, se não for "Todos"
                        if limite_export_selecionado != "Todos":
                            try:
                                # Extrai o número do "Top N"
                                num_a_exportar = int(limite_export_selecionado.split(" ")[1])
                                
                                # df_export_final já deve estar na ordem da AgGrid (que é ordenada por 'Total (€)' por defeito)
                                # Se precisar garantir a ordenação por 'Total (€)' ascendentemente aqui:
                                # if 'Total (€)' in df_export_final.columns:
                                #     df_export_final = df_export_final.sort_values(by='Total (€)', ascending=True)
                                
                                if len(df_export_final) > num_a_exportar:
                                    df_export_final = df_export_final.head(num_a_exportar)
                                    st.info(f"A exportar os {num_a_exportar} primeiros tarifários da tabela atual.")
                            except Exception as e_limite_export:
                                st.warning(f"Não foi possível aplicar o limite de tarifários: {e_limite_export}")


                        nome_coluna_tarifario_excel = None
                        if 'NomeParaExibir' in df_export_final.columns:
                            df_export_final.rename(columns={'NomeParaExibir': 'Tarifário'}, inplace=True)
                            nome_coluna_tarifario_excel = 'Tarifário'
                        elif 'Tarifário' in df_export_final.columns:
                            nome_coluna_tarifario_excel = 'Tarifário'

                        # --- Obter a coluna 'Tipo' do DataFrame original para usar na estilização ---
                        # Isto garante que temos os tipos mesmo que a coluna 'Tipo' não seja exportada.
                        # Assumimos que df_export_final mantém o índice de df_resultados_para_aggrid.
                        tipos_reais_para_estilo = None
                        if 'Tipo' in df_dados_filtrados_da_grid.columns: # Usar df_dados_filtrados_da_grid
                            try:
                                # df_export_final agora tem um novo índice (0, 1, 2...).
                                # Precisamos de alinhar com base no índice de df_dados_filtrados_da_grid que corresponde
                                # às linhas de df_export_final. Se df_export_final é apenas uma seleção de colunas
                                # de df_dados_filtrados_da_grid, o índice direto deve funcionar.
                                tipos_reais_para_estilo = df_dados_filtrados_da_grid.loc[df_export_final.index, 'Tipo']
                            except KeyError:
                                tipos_reais_para_estilo = pd.Series(index=df_export_final.index, dtype=str)
                        else:
                            tipos_reais_para_estilo = pd.Series(index=df_export_final.index, dtype=str)

                        # --- Função de interpolação de cores ---
                        def gerar_estilo_completo_para_valor(valor, minimo, maximo):
                            estilo_css_final = 'text-align: center;' 
                            if pd.isna(valor): return estilo_css_final
                            try: val_float = float(valor)
                            except ValueError: return estilo_css_final
                            if maximo == minimo or minimo is None or maximo is None: return estilo_css_final
                
                            midpoint = (minimo + maximo) / 2
                            r_bg, g_bg, b_bg = 255,255,255 
                            verde_rgb, branco_rgb, vermelho_rgb = (99,190,123), (255,255,255), (248,105,107)

                            if val_float <= midpoint:
                                ratio = (val_float - minimo) / (midpoint - minimo) if midpoint != minimo else 0.0
                                r_bg = int(verde_rgb[0]*(1-ratio) + branco_rgb[0]*ratio)
                                g_bg = int(verde_rgb[1]*(1-ratio) + branco_rgb[1]*ratio)
                                b_bg = int(verde_rgb[2]*(1-ratio) + branco_rgb[2]*ratio)
                            else:
                                ratio = (val_float - midpoint) / (maximo - midpoint) if maximo != midpoint else 0.0
                                r_bg = int(branco_rgb[0]*(1-ratio) + vermelho_rgb[0]*ratio)
                                g_bg = int(branco_rgb[1]*(1-ratio) + vermelho_rgb[1]*ratio)
                                b_bg = int(branco_rgb[2]*(1-ratio) + vermelho_rgb[2]*ratio)
                
                            estilo_css_final += f' background-color: #{r_bg:02X}{g_bg:02X}{b_bg:02X};'
                            luminancia = (0.299 * r_bg + 0.587 * g_bg + 0.114 * b_bg)
                            cor_texto_css = '#000000' if luminancia > 140 else '#FFFFFF'
                            estilo_css_final += f' color: {cor_texto_css};'
                            return estilo_css_final


                        # --- Função de estilo principal a ser aplicada ao DataFrame ---
                        def estilo_geral_dataframe_para_exportar(df_a_aplicar_estilo, tipos_reais_para_estilo_serie, min_max_config_para_cores, nome_coluna_tarifario="Tarifário"):
                            df_com_estilos = pd.DataFrame('', index=df_a_aplicar_estilo.index, columns=df_a_aplicar_estilo.columns)
               
                            # Cores default (pode ajustar)
                            cor_fundo_indexado_media_css_local = "#FFE699"
                            cor_texto_indexado_media_css_local = "black"
                            cor_fundo_indexado_dinamico_css_local = "#4D79BC"  
                            cor_texto_indexado_dinamico_css_local = "white"

                            for nome_coluna_df in df_a_aplicar_estilo.columns:
                                # Estilo para colunas de custo (Total (€) na detalhada, Total [Opção] (€) na comparativa)
                                if nome_coluna_df in min_max_config_para_cores: # min_max_config_para_cores é o min_max_data_X_json_string convertido para dict
                                    try:
                                        serie_valores_col = pd.to_numeric(df_a_aplicar_estilo[nome_coluna_df], errors='coerce')
                                        min_valor_col = min_max_config_para_cores[nome_coluna_df]['min']
                                        max_valor_col = min_max_config_para_cores[nome_coluna_df]['max']
                                        df_com_estilos[nome_coluna_df] = serie_valores_col.apply(
                                            lambda valor_v: gerar_estilo_completo_para_valor(valor_v, min_valor_col, max_valor_col)
                                        )
                                    except Exception as e_estilo_custo:
                                        print(f"Erro ao aplicar estilo de custo à coluna {nome_coluna_df}: {e_estilo_custo}")
                                        df_com_estilos[nome_coluna_df] = 'text-align: center;' 
        
                                elif nome_coluna_df == nome_coluna_tarifario: # 'Tarifário' ou 'NomeParaExibir'
                                    estilos_col_tarif_lista = []
                                    for idx_linha_df, valor_nome_col_tarif in df_a_aplicar_estilo[nome_coluna_df].items():
                                        tipo_tarif_str = tipos_reais_para_estilo_serie.get(idx_linha_df, '') if tipos_reais_para_estilo_serie is not None else ''
                
                                        est_css_tarif = 'text-align: center; padding: 4px;' 
                                        bg_cor_val, fonte_cor_val, fonte_peso_val = "#FFFFFF", "#000000", "normal" # Default Fixo/Outro (branco)

                                        if isinstance(valor_nome_col_tarif, str) and valor_nome_col_tarif.startswith("O Meu Tarifário"):
                                            bg_cor_val, fonte_cor_val, fonte_peso_val = "#FF0000", "#FFFFFF", "bold"
                                        elif tipo_tarif_str == 'Indexado Média':
                                            bg_cor_val, fonte_cor_val = cor_fundo_indexado_media_css_local, cor_texto_indexado_media_css_local
                                        elif tipo_tarif_str == 'Indexado quarto-horário':
                                            bg_cor_val, fonte_cor_val = cor_fundo_indexado_dinamico_css_local, cor_texto_indexado_dinamico_css_local
                                        elif tipo_tarif_str == 'Fixo': # Para dar um fundo um pouco diferente aos fixos
                                            bg_cor_val = "#F0F0F0" 
                
                                        est_css_tarif += f' background-color: {bg_cor_val}; color: {fonte_cor_val}; font-weight: {fonte_peso_val};'
                                        estilos_col_tarif_lista.append(est_css_tarif)
                                    df_com_estilos[nome_coluna_df] = estilos_col_tarif_lista
                                else: # Outras colunas de texto ou sem estilização de cor baseada em valor
                                    df_com_estilos[nome_coluna_df] = 'text-align: center;'
                            return df_com_estilos

                        # 1. Aplicar a função de estilo principal que retorna strings CSS
                        # Para a tabela detalhada, min_max_data_for_js é o dicionário com min/max para as colunas de preço e Total.
                        styler_excel = df_export_final.style.apply(
                            lambda df: estilo_geral_dataframe_para_exportar(df, tipos_reais_para_estilo, min_max_data_for_js, "Tarifário"), # Passa min_max_data_for_js
                            axis=None
                        )

                        # 2. Aplicar formatação de número (casas decimais)
                        for coluna_formatar in df_export_final.columns:
                            if '(€/kWh)' in coluna_formatar or '(€/dia)' in coluna_formatar:
                                styler_excel = styler_excel.format(formatter="{:.4f}", subset=[coluna_formatar], na_rep="-")
                            elif 'Total (€)' in coluna_formatar:
                                styler_excel = styler_excel.format(formatter="{:.2f}", subset=[coluna_formatar], na_rep="-")

            
                        # 3. Aplicar estilos de tabela gerais (cabeçalhos, bordas para todas as células td)
                        styler_excel = styler_excel.set_table_styles([
                            {'selector': 'th', 'props': [
                                ('background-color', '#404040'), ('color', 'white'),
                                ('font-weight', 'bold'), ('text-align', 'center'),
                                ('border', '1px solid black'), ('padding', '5px')]},
                            {'selector': 'td', 'props': [ 
                                ('border', '1px solid #dddddd'), ('padding', '4px')
                            ]}
                        ]).hide(axis="index")
            
                        # Obter o resumo_html e a mensagem de poupança
                        # Certifique-se que html_resumo_final está definido e acessível neste escopo
                        # (normalmente é uma variável global no seu script)
                        resumo_html_para_excel_func = html_resumo_final if 'html_resumo_final' in locals() else "Resumo não disponível."
                        poupanca_texto_para_excel_func = st.session_state.get('poupanca_excel_texto', "")

                        output_excel_bytes = exportar_excel_completo( # Sua função exportar_excel_completo
                            df_export_final,
                            styler_excel,
                            resumo_html_para_excel_func, 
                           poupanca_texto_para_excel_func,
                           opcao_horaria    
                        )

                        timestamp_final_dl = int(time.time()) # import time no início do script
                        nome_ficheiro_final_dl = f"Tiago_Felicia_Eletricidade_detalhe_{timestamp_final_dl}.xlsx"
            
                        st.download_button(
                            label=f"📥 Descarregar Excel ({nome_ficheiro_final_dl})",
                            data=output_excel_bytes.getvalue(), # output_excel_bytes é o BytesIO retornado por exportar_excel_completo
                            file_name=nome_ficheiro_final_dl,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"btn_dl_excel_completo_{timestamp_final_dl}" 
                        )
                        st.success(f"{nome_ficheiro_final_dl} pronto para download!")


    # Legenda das Colunas da Tabela Tarifários de Eletricidade
    st.markdown("---")
    st.subheader("📖 Legenda das Colunas da Tabela Tarifários de Eletricidade")
    st.caption("""
    * **Tarifário**: Nome identificativo do tarifário. Pode incluir notas sobre descontos de fatura específicos.
    * **Tipo**: Indica se o tarifário é:
        * `Fixo`: Preços de energia e potência são constantes.
        * `Indexado (Média)`: Preço da energia baseado na média do OMIE para os períodos horários.
        * `Indexado (Quarto-Horário)`: Preço da energia baseado nos valores OMIE horários/quarto-horários e no perfil de consumo. Também conhecidos como "Dinâmicos".
        * `Pessoal`: O seu tarifário, conforme introduzido.
    * **Comercializador**: Empresa que oferece o tarifário.
    * **[...] (€/kWh)**: Custo unitário da energia para o período indicado (Simples, Vazio, Fora Vazio, Cheias, Ponta), **sem IVA**.
        * Para "O Meu Tarifário", este valor já reflete quaisquer descontos percentuais de energia e o desconto da Tarifa Social que tenhas configurado.
        * Para os outros tarifários, é o preço base sem IVA, já considerando o desconto da Tarifa Social se ativa.
    * **Potência (€/dia)**: Custo unitário diário da potência contratada e Termo Fixo **sem IVA**.
        * Para "O Meu Tarifário", este valor já reflete quaisquer descontos percentuais de potência e o desconto da Tarifa Social que tenhas configurado.
        * Para os outros tarifários, é o preço base sem IVA, já considerando o desconto da Tarifa Social se ativa.
    * **Total (€)**: Valor do custo final estimado da fatura para o período simulado. Este custo inclui:
        * Custo da energia consumida (com IVA aplicado conforme as regras).
        * Custo da potência contratada (com IVA aplicado conforme as regras).
        * Taxas adicionais: IEC (Imposto Especial de Consumo, isento com Tarifa Social), DGEG (Taxa de Exploração da Direção-Geral de Energia e Geologia) e CAV (Contribuição Audiovisual).
        * Quaisquer descontos de fatura em euros (para "O Meu Tarifário" ou especificados nos tarifários).
    """)


    st.subheader("🎨 Legenda de Cores por Tipo de Tarifário")

# Cores para "O Meu Tarifário" (replicar do JS)
    cor_fundo_meu_tarifario_legenda = "red"
    cor_texto_meu_tarifario_legenda = "white"

    cor_fundo_fixo_legenda = "#FFFFFF"
    cor_texto_fixo_legenda = "#333333"
    borda_fixo_legenda = "#CCCCCC"     # Borda para o quadrado branco ser visível

# Usar f-strings para construir o HTML da legenda
    legenda_html = f"""
    <div style="font-size: 14px;">
        <div style="display: flex; align-items: center; margin-bottom: 5px;">
            <div style="width: 18px; height: 18px; background-color: {cor_fundo_meu_tarifario_legenda}; border: 1px solid #ccc; border-radius: 4px; margin-right: 8px;"></div>
            <span style="background-color: {cor_fundo_meu_tarifario_legenda}; color: {cor_texto_meu_tarifario_legenda}; padding: 2px 6px; border-radius: 4px; font-weight: bold;">O Meu Tarifário</span>
            <span style="margin-left: 8px;">- Tarifário configurado pelo utilizador.</span>
        </div>
        <div style="display: flex; align-items: center; margin-bottom: 5px;">
            <div style="width: 18px; height: 18px; background-color: {cor_fundo_indexado_media_css}; border: 1px solid #ccc; border-radius: 4px; margin-right: 8px;"></div>
            <span style="background-color: {cor_fundo_indexado_media_css}; color: {cor_texto_indexado_media_css}; padding: 2px 6px; border-radius: 4px;">Indexado (Média)</span>
            <span style="margin-left: 8px;">- Preço de energia baseado na média OMIE do período definido.</span>
        </div>
        <div style="display: flex; align-items: center; margin-bottom: 5px;">
            <div style="width: 18px; height: 18px; background-color: {cor_fundo_indexado_dinamico_css}; border: 1px solid #ccc; border-radius: 4px; margin-right: 8px;"></div>
            <span style="background-color: {cor_fundo_indexado_dinamico_css}; color: {cor_texto_indexado_dinamico_css}; padding: 2px 6px; border-radius: 4px;">Indexado (Quarto-Horário)</span>
            <span style="margin-left: 8px;">- Preço de energia baseado nos valores OMIE horários/quarto-horários e perfil.</span>
        </div>
        <div style="display: flex; align-items: center; margin-bottom: 5px;">
            <div style="width: 18px; height: 18px; background-color: {cor_fundo_fixo_legenda}; border: 1px solid {borda_fixo_legenda}; border-radius: 4px; margin-right: 8px;"></div>
            <span style="background-color: {cor_fundo_fixo_legenda}; color: {cor_texto_fixo_legenda}; padding: 2px 6px; border-radius: 4px;">Tarifário Fixo</span>
            <span style="margin-left: 8px;">- Preços de energia constantes.</span>
        </div>
    </div>
    """
    st.markdown(legenda_html, unsafe_allow_html=True)

else: # df_resultados original estava vazio
    st.info("Não foram encontrados tarifários para a opção selecionada.")

# --- DATAS DE REFERÊNCIA ---
st.markdown("---") # Adiciona um separador visual
st.subheader("📅 Datas de Referência dos Valores de Mercado no simulador")

# 1. Processar e exibir Data_Valores_OMIE
# data_valores_omie_dt já foi processada no início do script.
data_omie_formatada_str = "Não disponível"
if data_valores_omie_dt and isinstance(data_valores_omie_dt, datetime.date):
    try:
        data_omie_formatada_str = data_valores_omie_dt.strftime('%d/%m/%Y')
    except ValueError: # No caso de uma data inválida que passou pelo isinstance
        data_omie_formatada_str = f"Data inválida ({data_valores_omie_dt})"
elif data_valores_omie_dt: # Se existe mas não é um objeto date (pode indicar erro no processamento inicial)
    data_omie_formatada_str = f"Valor não reconhecido como data ({data_valores_omie_dt})"

st.markdown(f"**Valores OMIE (SPOT) até** {data_omie_formatada_str}")

# 2. Processar e exibir Data_Valores_OMIP
data_omip_formatada_str = "Não disponível"
constante_omip_df_row = CONSTANTES[CONSTANTES['constante'] == 'Data_Valores_OMIP'] # Renomeado para evitar conflito

if not constante_omip_df_row.empty:
    valor_bruto_omip = constante_omip_df_row['valor_unitário'].iloc[0]
    if pd.notna(valor_bruto_omip):
        data_omip_dt_temp = None # Variável temporária para a data OMIP
        try:
            # Tenta converter para pd.Timestamp, que é mais flexível, e depois para objeto date
            if isinstance(valor_bruto_omip, (datetime.datetime, pd.Timestamp)):
                data_omip_dt_temp = valor_bruto_omip.date()
            else:
                timestamp_convertido_omip = pd.to_datetime(valor_bruto_omip, errors='coerce')
                if pd.notna(timestamp_convertido_omip):
                    data_omip_dt_temp = timestamp_convertido_omip.date()
            
            if data_omip_dt_temp and isinstance(data_omip_dt_temp, datetime.date):
                data_omip_formatada_str = data_omip_dt_temp.strftime('%d/%m/%Y')
            elif valor_bruto_omip: # Se a conversão falhou mas havia um valor
                data_omip_formatada_str = f"Valor não reconhecido como data ({valor_bruto_omip})"
            # Se valor_bruto_omip for pd.NaT ou a conversão falhar completamente, mantém "Não disponível"
        except Exception: # Captura outros erros de conversão
            if valor_bruto_omip:
                data_omip_formatada_str = f"Erro ao processar valor ({valor_bruto_omip})"
    # Se valor_bruto_omip for NaN, data_omip_formatada_str permanece "Não disponível"
# Se a constante 'Data_Valores_OMIP' não for encontrada, data_omip_formatada_str permanece "Não disponível"

st.markdown(f"**Valores OMIP (Futuros) atualizados em** {data_omip_formatada_str}")
# --- FIM DA NOVA SECÇÃO ---

# --- INÍCIO DA SECÇÃO DE APOIO ---
st.markdown("---") # Adiciona um separador visual antes da secção de apoio
st.subheader("💖 Apoie este Projeto")

st.markdown(
    "Se quiser apoiar a manutenção do site e o desenvolvimento contínuo deste simulador, "
    "pode fazê-lo através de uma das seguintes formas:"
)

# Link para BuyMeACoffee
st.markdown(
    "☕ [**Compre-me um café em BuyMeACoffee**](https://buymeacoffee.com/tiagofelicia)"
)

st.markdown("ou através do botão PayPal:")

# Código HTML para o botão do PayPal
paypal_button_html = """
<div style="text-align: left; margin-top: 10px; margin-bottom: 15px;">
    <form action="https://www.paypal.com/donate" method="post" target="_blank" style="display: inline-block;">
    <input type="hidden" name="hosted_button_id" value="W6KZHVL53VFJC">
    <input type="image" src="https://www.paypalobjects.com/pt_PT/PT/i/btn/btn_donate_SM.gif" border="0" name="submit" title="PayPal - The safer, easier way to pay online!" alt="Faça donativos com o botão PayPal">
    <img alt="" border="0" src="https://www.paypal.com/pt_PT/i/scr/pixel.gif" width="1" height="1">
    </form>
</div>
"""
st.markdown(paypal_button_html, unsafe_allow_html=True)
# --- FIM DA SECÇÃO DE APOIO ---

st.markdown("---")
# Título para as redes sociais
st.subheader("Redes sociais, onde poderão seguir o projeto:")

# URLs das redes sociais
url_x = "https://x.com/tiagofelicia"
url_bluesky = "https://bsky.app/profile/tiagofelicia.bsky.social"
url_youtube = "https://youtube.com/@tiagofelicia"
url_facebook_perfil = "https://www.facebook.com/profile.php?id=61555007360529"


icon_url_x = "https://upload.wikimedia.org/wikipedia/commons/thumb/c/cc/X_icon.svg/120px-X_icon.svg.png?20250519203220"
icon_url_bluesky = "https://upload.wikimedia.org/wikipedia/commons/7/7a/Bluesky_Logo.svg"
icon_url_youtube = "https://upload.wikimedia.org/wikipedia/commons/thumb/f/fd/YouTube_full-color_icon_%282024%29.svg/120px-YouTube_full-color_icon_%282024%29.svg.png"
icon_url_facebook = "https://upload.wikimedia.org/wikipedia/commons/thumb/b/b9/2023_Facebook_icon.svg/120px-2023_Facebook_icon.svg.png"


svg_icon_style_dark_mode_friendly = "filter: invert(0.8) sepia(0) saturate(1) hue-rotate(0deg) brightness(1.5) contrast(0.8);"

col_social1, col_social2, col_social3, col_social4 = st.columns(4)

with col_social1:
    st.markdown(
        f"""
        <a href="{url_x}" target="_blank" style="text-decoration: none; color: inherit; display: flex; flex-direction: column; align-items: center; text-align: center;">
            <img src="{icon_url_x}" width="40" alt="X" style="margin-bottom: 8px; object-fit: contain;">
            X
        </a>
        """,
        unsafe_allow_html=True
    )

with col_social2:
    st.markdown(
        f"""
        <a href="{url_bluesky}" target="_blank" style="text-decoration: none; color: inherit; display: flex; flex-direction: column; align-items: center; text-align: center;">
            <img src="{icon_url_bluesky}" width="40" alt="Bluesky" style="margin-bottom: 8px; object-fit: contain;">
            Bluesky
        </a>
        """,
        unsafe_allow_html=True
    )

with col_social3:
    st.markdown(
        f"""
        <a href="{url_youtube}" target="_blank" style="text-decoration: none; color: inherit; display: flex; flex-direction: column; align-items: center; text-align: center;">
            <img src="{icon_url_youtube}" width="40" alt="YouTube" style="margin-bottom: 8px; object-fit: contain;">
            YouTube
        </a>
        """,
        unsafe_allow_html=True
    )

with col_social4:
    st.markdown(
        f"""
        <a href="{url_facebook_perfil}" target="_blank" style="text-decoration: none; color: inherit; display: flex; flex-direction: column; align-items: center; text-align: center;">
            <img src="{icon_url_facebook}" width="40" alt="Facebook" style="margin-bottom: 8px; object-fit: contain;">
            Facebook
        </a>
        """,
        unsafe_allow_html=True
    )

st.markdown("<br>", unsafe_allow_html=True) # Adiciona um espaço vertical

# Texto de Copyright
ano_copyright = 2025
nome_autor = "Tiago Felícia"
texto_copyright_html = f"© {ano_copyright} Todos os direitos reservados | {nome_autor} | <a href='{url_facebook_perfil}' target='_blank' style='color: inherit;'>Facebook</a>"

st.markdown(
    f"<div style='text-align: center; font-size: 0.9em; color: grey;'>{texto_copyright_html}</div>",
    unsafe_allow_html=True
)
