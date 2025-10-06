import streamlit as st
import pandas as pd
import re
import requests
import numpy as np
from io import StringIO

# Importar as constantes e funções que são necessárias dentro deste módulo

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

# --- Função para obter valores da aba Constantes ---
def obter_constante(nome_constante, constantes_df):
    constante_row = constantes_df[constantes_df['constante'] == nome_constante]
    if not constante_row.empty:
        valor = constante_row['valor_unitário'].iloc[0]
        try:
            return float(valor)
        except (ValueError, TypeError):
            return 0.0
    else:
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

# --- Função: Obter valor da TAR potência para a potência contratada ---
def obter_tar_dia(potencia_kva, constantes_df):
    potencia_str = str(float(potencia_kva)) # Formato consistente
    constante_potencia = f'TAR_Potencia {potencia_str}'
    return obter_constante(constante_potencia, constantes_df)

# --- Função: Determinar o perfil BTN ---
def obter_perfil(consumo_total_kwh, dias, potencia_kva):
    consumo_anual_estimado = consumo_total_kwh * 365 / dias if dias > 0 else consumo_total_kwh
    if potencia_kva > 13.8: return 'perfil_A'
    elif consumo_anual_estimado > 7140: return 'perfil_B'
    else: return 'perfil_C'

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

# --- Função: Calcular taxas adicionais ---
def calcular_taxas_adicionais(
    consumo_kwh,
    dias_simulacao,
    tarifa_social_bool,
    valor_dgeg_mensal,
    valor_cav_mensal,
    nome_comercializador_atual,
    aplica_taxa_fixa_mensal, # NOVO PARÂMETRO
    valor_iec=0.001
):
    """
    Calcula taxas adicionais (IEC, DGEG, CAV) com lógica ajustada para CAV mensal fixa
    para comercializadores específicos.
    A decisão de aplicar a taxa fixa é agora controlada pelo parâmetro 'aplica_taxa_fixa_mensal'.
    """
    iva_normal_perc = 0.23
    iva_reduzido_perc = 0.06 # IVA da CAV é 6%

    consumo_kwh_float = float(consumo_kwh or 0.0)
    dias_simulacao_int = int(dias_simulacao or 0)
    valor_dgeg_mensal_float = float(valor_dgeg_mensal or 0.0)
    valor_cav_mensal_float = float(valor_cav_mensal or 0.0) # Renomeado para consistência
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
    
    # LÓGICA ALTERADA: A decisão vem de fora da função
    aplica_cav_fixa_mensal_final = False
    if nome_comercializador_atual and isinstance(nome_comercializador_atual, str):
        nome_comerc_lower_para_verificacao = nome_comercializador_atual.lower()
        if any(identificador.lower() in nome_comerc_lower_para_verificacao for identificador in IDENTIFICADORES_COMERCIALIZADORES_CAV_FIXA):
            # Se for um comercializador da lista, a decisão depende do parâmetro externo
            if aplica_taxa_fixa_mensal:
                aplica_cav_fixa_mensal_final = True

    if aplica_cav_fixa_mensal_final:
        cav_siva = valor_cav_mensal_float  # Aplica o valor mensal total da CAV
    else:
        # Cálculo proporcional padrão para os outros casos
        cav_siva = (valor_cav_mensal_float * 12 / 365.25 * dias_simulacao_int)

    # Valores de IVA (lógica mantém-se)
    iva_iec = 0.0 if tarifa_social_bool else (iec_siva * iva_normal_perc)
    iva_dgeg = dgeg_siva * iva_normal_perc
    iva_cav = cav_siva * iva_reduzido_perc

    custo_total_siva = iec_siva + dgeg_siva + cav_siva
    custo_total_com_iva = (iec_siva + iva_iec) + (dgeg_siva + iva_dgeg) + (cav_siva + iva_cav)
    
    total_iva_6_calculado = iva_cav
    total_iva_23_calculado = iva_iec + iva_dgeg

    return {
        'custo_com_iva': round(custo_total_com_iva, 4),
        'custo_sem_iva': round(custo_total_siva, 4),
        'iec_sem_iva': round(iec_siva, 4),
        'dgeg_sem_iva': round(dgeg_siva, 4),
        'cav_sem_iva': round(cav_siva, 4),
        'valor_iva_6': round(total_iva_6_calculado, 4),
        'valor_iva_23': round(total_iva_23_calculado, 4)
    }

def calcular_custo_completo_diagrama_carga(tarifario_idx, df_consumos_reais, df_omie_ciclos, constantes_df, dias, potencia, familia_numerosa, tarifa_social, valor_dgeg_user, valor_cav_user, mes, ano_atual, incluir_quota_acp, desconto_continente, FINANCIAMENTO_TSE_VAL,VALOR_QUOTA_ACP_MENSAL):
    """
    Calcula o custo COMPLETO de um tarifário quarto-horário usando os consumos reais,
    incluindo a decomposição detalhada para os tooltips e todos os descontos específicos.
    Devolve um dicionário plano com todos os dados para a tabela detalhada e para os tooltips.
    """
    try:
        # --- Inicializar dicionários para os componentes dos tooltips ---
        componentes_tooltip_energia_dict = {}
        componentes_tooltip_potencia_dict = {}

        # 1. Cruzamento de Dados e Cálculo de Componentes Base
        df_merged = pd.merge(df_consumos_reais, df_omie_ciclos, on='DataHora', how='left')
        df_merged.dropna(subset=['OMIE', 'Perdas'], inplace=True)
        if df_merged.empty: return None

        nome_tarifario = tarifario_idx['nome']
        constantes_dict = dict(zip(constantes_df["constante"], constantes_df["valor_unitário"]))

        def calcular_preco_comercializador_intervalo(row):
            omie_kwh = row['OMIE'] / 1000.0; perdas = row['Perdas']
            if nome_tarifario == "Coopérnico Base 2.0": return (omie_kwh + constantes_dict.get('Coop_CS_CR', 0.0) + constantes_dict.get('Coop_K', 0.0)) * perdas
            elif "Repsol - Leve PRO Sem Mais" in nome_tarifario: return (omie_kwh * perdas * constantes_dict.get('Repsol_FA', 0.0) + constantes_dict.get('Repsol_Q_Tarifa_Pro', 0.0))
            elif "Repsol - Leve Sem Mais" in nome_tarifario: return (omie_kwh * perdas * constantes_dict.get('Repsol_FA', 0.0) + constantes_dict.get('Repsol_Q_Tarifa', 0.0))
            elif "Galp - Plano Flexível / Dinâmico" in nome_tarifario: return (omie_kwh + constantes_dict.get('Galp_Ci', 0.0)) * perdas
            elif "Alfa Energia - ALFA POWER INDEX BTN" in nome_tarifario: return ((omie_kwh + constantes_dict.get('Alfa_CGS', 0.0)) * perdas + constantes_dict.get('Alfa_K', 0.0))
            elif "Plenitude - Tendência" in nome_tarifario: return (((omie_kwh + constantes_dict.get('Plenitude_CGS', 0.0) + constantes_dict.get('Plenitude_GDOs', 0.0))) * perdas + constantes_dict.get('Plenitude_Fee', 0.0))
            elif "Meo Energia - Tarifa Variável" in nome_tarifario: return (omie_kwh + constantes_dict.get('Meo_K', 0.0)) * perdas
            elif "EDP - Eletricidade Indexada Horária" in nome_tarifario: return (omie_kwh * perdas * constantes_dict.get('EDP_H_K1', 1.0) + constantes_dict.get('EDP_H_K2', 0.0))
            elif "EZU - Coletiva" in nome_tarifario: return (omie_kwh + constantes_dict.get('EZU_K', 0.0) + constantes_dict.get('EZU_CGS', 0.0)) * perdas
            elif "G9 - Smart Dynamic" in nome_tarifario: return (omie_kwh * constantes_dict.get('G9_FA', 0.0) * perdas + constantes_dict.get('G9_CGS', 0.0) + constantes_dict.get('G9_AC', 0.0))
            elif "Iberdrola - Simples Indexado Dinâmico" in nome_tarifario: return (omie_kwh * perdas + constantes_dict.get("Iberdrola_Q", 0.0) + constantes_dict.get('Iberdrola_mFRR', 0.0))
            elif "Luzboa - BTN SPOTDEF" in nome_tarifario: return (omie_kwh + constantes_dict.get('Luzboa_CGS', 0.0)) * perdas * constantes_dict.get('Luzboa_FA', 1.0) + constantes_dict.get('Luzboa_Kp', 0.0)
            return omie_kwh * perdas

        df_merged['PrecoComercializadorIntervalo_sIVA'] = df_merged.apply(calcular_preco_comercializador_intervalo, axis=1)
        df_merged['CustoComercializadorIntervalo_sIVA'] = df_merged['PrecoComercializadorIntervalo_sIVA'] * df_merged['Consumo (kWh)']

        # 2. Agregação e Cálculo de Preços Médios Finais
        precos_medios_finais_siva = {}
        opcao_horaria_idx = tarifario_idx['opcao_horaria_e_ciclo']
        consumo_total_real = df_consumos_reais['Consumo (kWh)'].sum()

        financiamento_tse_unitario = obter_constante('Financiamento_TSE', constantes_df) if not tarifario_idx.get('financiamento_tse_incluido', False) else 0.0
        desconto_ts_energia_unitario = obter_constante('Desconto TS Energia', constantes_df) if tarifa_social else 0.0

        ciclo_col_idx = None
        opcao_lower_str = str(opcao_horaria_idx).lower()
        if opcao_lower_str.startswith("bi-horário"):
            ciclo_col_idx = 'BD' if "diário" in opcao_lower_str else 'BS'
        elif opcao_lower_str.startswith("tri-horário"):
            ciclo_col_idx = 'TD' if "diário" in opcao_lower_str else 'TS'

        consumos_repartidos_reais = {'S': consumo_total_real}
        if ciclo_col_idx and ciclo_col_idx in df_merged.columns:
            consumos_repartidos_reais = df_merged.groupby(ciclo_col_idx)['Consumo (kWh)'].sum().to_dict()
            for periodo, group in df_merged.groupby(ciclo_col_idx):
                consumo_p = group['Consumo (kWh)'].sum()
                if consumo_p > 0:
                    comerc_preco_medio = group['CustoComercializadorIntervalo_sIVA'].sum() / consumo_p
                    tar_unitaria = obter_tar_energia_periodo(opcao_horaria_idx, periodo, potencia, constantes_df)
                    precos_medios_finais_siva[periodo] = comerc_preco_medio + tar_unitaria + financiamento_tse_unitario - desconto_ts_energia_unitario
                    componentes_tooltip_energia_dict[f'tooltip_energia_{periodo}_comerc_sem_tar'] = comerc_preco_medio
                    componentes_tooltip_energia_dict[f'tooltip_energia_{periodo}_tar_bruta'] = tar_unitaria
                    componentes_tooltip_energia_dict[f'tooltip_energia_{periodo}_tse_declarado_incluido'] = tarifario_idx.get('financiamento_tse_incluido', False)
                    componentes_tooltip_energia_dict[f'tooltip_energia_{periodo}_tse_valor_nominal'] = FINANCIAMENTO_TSE_VAL
                    componentes_tooltip_energia_dict[f'tooltip_energia_{periodo}_ts_aplicada_flag'] = tarifa_social
                    componentes_tooltip_energia_dict[f'tooltip_energia_{periodo}_ts_desconto_valor'] = desconto_ts_energia_unitario

        comerc_preco_medio_simples = df_merged['CustoComercializadorIntervalo_sIVA'].sum() / consumo_total_real if consumo_total_real > 0 else 0
        tar_media_ponderada = (df_merged.apply(lambda row: obter_tar_energia_periodo(opcao_horaria_idx, row.get(ciclo_col_idx, 'S'), potencia, constantes_df) * row['Consumo (kWh)'], axis=1).sum()) / consumo_total_real if consumo_total_real > 0 else 0
        precos_medios_finais_siva['S'] = comerc_preco_medio_simples + tar_media_ponderada + financiamento_tse_unitario - desconto_ts_energia_unitario
        componentes_tooltip_energia_dict['tooltip_energia_S_comerc_sem_tar'] = comerc_preco_medio_simples
        componentes_tooltip_energia_dict['tooltip_energia_S_tar_bruta'] = tar_media_ponderada
        componentes_tooltip_energia_dict['tooltip_energia_S_tse_declarado_incluido'] = tarifario_idx.get('financiamento_tse_incluido', False)
        componentes_tooltip_energia_dict['tooltip_energia_S_tse_valor_nominal'] = FINANCIAMENTO_TSE_VAL
        componentes_tooltip_energia_dict['tooltip_energia_S_ts_aplicada_flag'] = tarifa_social
        componentes_tooltip_energia_dict['tooltip_energia_S_ts_desconto_valor'] = desconto_ts_energia_unitario

        # 3. Decomposição e Cálculo do Custo Total da Fatura (Componentes Base)
        decomposicao_custo_energia = calcular_custo_energia_com_iva(consumo_total_real, precos_medios_finais_siva.get('S'), {p:v for p,v in precos_medios_finais_siva.items() if p != 'S'}, dias, potencia, opcao_horaria_idx, consumos_repartidos_reais, familia_numerosa)

        tar_potencia_regulada = obter_constante(f'TAR_Potencia {str(float(potencia))}', constantes_df)
        preco_pot_comercializador = float(tarifario_idx.get('preco_potencia_dia', 0.0))
        if tarifario_idx.get('tar_incluida_potencia', True): preco_pot_comercializador -= tar_potencia_regulada
        desconto_ts_pot_bruto = obter_constante(f'Desconto TS Potencia {potencia}', constantes_df) if tarifa_social else 0
        desconto_ts_pot_aplicado = min(tar_potencia_regulada, desconto_ts_pot_bruto)
        tar_potencia_final_siva = tar_potencia_regulada - desconto_ts_pot_aplicado
        decomposicao_custo_potencia = calcular_custo_potencia_com_iva_final(preco_pot_comercializador, tar_potencia_final_siva, dias, potencia)
        
        preco_unit_potencia_siva_final = preco_pot_comercializador + tar_potencia_final_siva

        componentes_tooltip_potencia_dict = {'tooltip_pot_comerc_sem_tar': preco_pot_comercializador, 'tooltip_pot_tar_bruta': tar_potencia_regulada, 'tooltip_pot_ts_aplicada': tarifa_social, 'tooltip_pot_desconto_ts_valor': desconto_ts_pot_aplicado}

        is_billing_month = 28 <= dias <= 31
        decomposicao_taxas = calcular_taxas_adicionais(consumo_total_real, dias, tarifa_social, valor_dgeg_user, valor_cav_user, tarifario_idx.get('comercializador'), aplica_taxa_fixa_mensal=is_billing_month)

        custo_total_antes_desc_especificos = decomposicao_custo_energia['custo_com_iva'] + decomposicao_custo_potencia['custo_com_iva'] + decomposicao_taxas['custo_com_iva']

        # 4. Lógica de Descontos Específicos
        nome_a_exibir = nome_tarifario # ALTERAÇÃO: Começamos com o nome base
        custo_final_com_descontos = custo_total_antes_desc_especificos
        desconto_total_final = 0.0
        acrescimo_total_final = 0.0

        desconto_fatura_mensal = float(tarifario_idx.get('desconto_fatura_mes', 0.0) or 0.0)
        if desconto_fatura_mensal > 0:
            desconto_aplicado = desconto_fatura_mensal if is_billing_month else (desconto_fatura_mensal / 30.0) * dias
            custo_final_com_descontos -= desconto_aplicado
            desconto_total_final += desconto_aplicado
            nome_a_exibir += f" (INCLUI desconto {desconto_fatura_mensal:.2f}€/mês)"
        if incluir_quota_acp and nome_tarifario.startswith("Goldenergy - ACP"):
            quota_aplicada = VALOR_QUOTA_ACP_MENSAL if is_billing_month else (VALOR_QUOTA_ACP_MENSAL / 30.0) * dias
            custo_final_com_descontos += quota_aplicada
            acrescimo_total_final += quota_aplicada
            nome_a_exibir += f" (INCLUI Quota ACP - {VALOR_QUOTA_ACP_MENSAL:.2f} €/mês)"
        consumo_mensal_equivalente = (consumo_total_real / dias) * 30.0 if dias > 0 else 0
        if "meo energia - tarifa fixa - clientes meo" in nome_tarifario.lower() and consumo_mensal_equivalente >= 216:
            desconto_meo_mensal_base = 0.0
            if "simples" in opcao_horaria_idx.lower(): desconto_meo_mensal_base = 2.95
            elif "bi-horário" in opcao_horaria_idx.lower(): desconto_meo_mensal_base = 3.50
            elif "tri-horário" in opcao_horaria_idx.lower(): desconto_meo_mensal_base = 6.27
            if desconto_meo_mensal_base > 0:
                desconto_aplicado = (desconto_meo_mensal_base / 30.0) * dias
                custo_final_com_descontos -= desconto_aplicado
                desconto_total_final += desconto_aplicado
                nome_a_exibir += f" (Desc. MEO Clientes {desconto_aplicado:.2f}€ incl.)"
        if desconto_continente and nome_tarifario.startswith("Galp & Continente"):
            # Lógica base para calcular o custo bruto (comum a ambos os descontos)
            custo_energia_bruto_siva = comerc_preco_medio_simples + tar_media_ponderada + financiamento_tse_unitario
            decomposicao_energia_bruta = calcular_custo_energia_com_iva(consumo_total_real, custo_energia_bruto_siva, {}, dias, potencia, "Simples", {'S': consumo_total_real}, familia_numerosa)
            decomposicao_potencia_bruta = calcular_custo_potencia_com_iva_final(preco_pot_comercializador, tar_potencia_regulada, dias, potencia)
            base_desconto_continente = decomposicao_energia_bruta['custo_com_iva'] + decomposicao_potencia_bruta['custo_com_iva']
            
            # ### DESCONTO DE 10% ###
            if nome_tarifario.startswith("Galp & Continente (-10% DD)"):
                desconto_aplicado = base_desconto_continente * 0.10
                custo_final_com_descontos -= desconto_aplicado
                desconto_total_final += desconto_aplicado
                nome_a_exibir += f" (INCLUI desc. Cont. de {desconto_aplicado:.2f}€, s/ desc. Cont.={(custo_final_com_descontos + desconto_aplicado):.2f}€)"
            
            # ### DESCONTO DE 7% ###
            elif nome_tarifario.startswith("Galp & Continente (-7% s/DD)"):
                desconto_aplicado = base_desconto_continente * 0.07
                custo_final_com_descontos -= desconto_aplicado
                desconto_total_final += desconto_aplicado
                nome_a_exibir += f" (INCLUI desc. Cont. de {desconto_aplicado:.2f}€, s/ desc. Cont.={(custo_final_com_descontos + desconto_aplicado):.2f}€)"

        # Após todos os descontos terem sido adicionados ao nome, acrescentamos o sufixo.
        nome_a_exibir += " - Diagrama"

        # --- INÍCIO DA ALTERAÇÃO ---
        # 5. Montar Dicionário FLAT Final
        componentes_tooltip_custo_total_dict = {
            'tt_cte_energia_siva': decomposicao_custo_energia['custo_sem_iva'],
            'tt_cte_potencia_siva': decomposicao_custo_potencia['custo_sem_iva'],
            'tt_cte_iec_siva': decomposicao_taxas['iec_sem_iva'],
            'tt_cte_dgeg_siva': decomposicao_taxas['dgeg_sem_iva'],
            'tt_cte_cav_siva': decomposicao_taxas['cav_sem_iva'],
            'tt_cte_total_siva': decomposicao_custo_energia['custo_sem_iva'] + decomposicao_custo_potencia['custo_sem_iva'] + decomposicao_taxas['custo_sem_iva'],
            'tt_cte_valor_iva_6_total': decomposicao_custo_energia['valor_iva_6'] + decomposicao_custo_potencia['valor_iva_6'] + decomposicao_taxas['valor_iva_6'],
            'tt_cte_valor_iva_23_total': decomposicao_custo_energia['valor_iva_23'] + decomposicao_custo_potencia['valor_iva_23'] + decomposicao_taxas['valor_iva_23'],
            'tt_cte_subtotal_civa': custo_total_antes_desc_especificos,
            'tt_cte_desc_finais_valor': desconto_total_final,
            'tt_cte_acres_finais_valor': acrescimo_total_final,
            # ADICIONAR PREÇOS UNITÁRIOS AO TOOLTIP
            'tt_preco_unit_energia_S_siva': precos_medios_finais_siva.get('S'),
            'tt_preco_unit_energia_V_siva': precos_medios_finais_siva.get('V'),
            'tt_preco_unit_energia_F_siva': precos_medios_finais_siva.get('F'),
            'tt_preco_unit_energia_C_siva': precos_medios_finais_siva.get('C'),
            'tt_preco_unit_energia_P_siva': precos_medios_finais_siva.get('P'),
            'tt_preco_unit_potencia_siva': preco_unit_potencia_siva_final
        }

        resultado = {
            'NomeParaExibir': nome_a_exibir,
            'Tipo': f"{tarifario_idx.get('tipo')} (Diagrama)",
            'Total (€)': round(custo_final_com_descontos, 2),
            # COLUNAS PARA A TABELA DETALHADA
            'Simples (€/kWh)': round(precos_medios_finais_siva.get('S', 0), 4),
            'Vazio (€/kWh)': round(precos_medios_finais_siva.get('V', 0), 4),
            'Fora Vazio (€/kWh)': round(precos_medios_finais_siva.get('F', 0), 4),
            'Cheias (€/kWh)': round(precos_medios_finais_siva.get('C', 0), 4),
            'Ponta (€/kWh)': round(precos_medios_finais_siva.get('P', 0), 4),
            'Potência (€/dia)': round(preco_unit_potencia_siva_final, 4),
            # OUTRAS COLUNAS DE INFO
            'LinkAdesao': tarifario_idx.get('site_adesao'), 'info_notas': tarifario_idx.get('notas', ''),
            'Comercializador': tarifario_idx.get('comercializador'), 'Segmento': tarifario_idx.get('segmento'),
            'Faturação': tarifario_idx.get('faturacao'), 'Pagamento': tarifario_idx.get('pagamento'),
            # DESEMPACOTAR TODOS OS DADOS DE TOOLTIP NO DICIONÁRIO FINAL
            **componentes_tooltip_energia_dict,
            **componentes_tooltip_potencia_dict,
            **componentes_tooltip_custo_total_dict
        }
        return resultado
        # --- FIM DA ALTERAÇÃO ---

    except Exception as e:
        st.error(f"Erro em `calcular_custo_completo_diagrama_carga` para {tarifario_idx.get('nome', 'desconhecido')}: {e}")
        return None
    
### NOVO: Função de cálculo dedicada para o Tarifário Personalizado ###
def calcular_custo_personalizado(precos_energia_pers, preco_potencia_pers, consumos_para_calculo, flags_pers, CONSTANTES, FINANCIAMENTO_TSE_VAL,**kwargs):
    """
    Função reutilizável para calcular o custo de uma estrutura tarifária personalizada.
    Agora também retorna os dicionários completos para os tooltips.
    """
    # Extrair parâmetros globais
    dias = kwargs.get('dias')
    potencia = kwargs.get('potencia')
    tarifa_social = kwargs.get('tarifa_social')
    familia_numerosa = kwargs.get('familia_numerosa')
    valor_dgeg_user = kwargs.get('valor_dgeg_user')
    valor_cav_user = kwargs.get('valor_cav_user')
    opcao_horaria_ref = kwargs.get('opcao_horaria_ref')

    # 1. Obter componentes base (sem IVA)
    tar_energia_reg = {p: obter_tar_energia_periodo(opcao_horaria_ref, p, potencia, CONSTANTES) for p in consumos_para_calculo.keys()}
    tar_potencia_reg = obter_tar_dia(potencia, CONSTANTES)
    
    comerc_energia = {p: (preco - tar_energia_reg.get(p, 0)) if flags_pers['tar_energia'] else preco for p, preco in precos_energia_pers.items()}
    comerc_potencia = (preco_potencia_pers - tar_potencia_reg) if flags_pers['tar_potencia'] else preco_potencia_pers
    
    financiamento_tse_a_somar = FINANCIAMENTO_TSE_VAL if not flags_pers['tse_incluido'] else 0.0

    # 2. Preços finais sem IVA (já com descontos TS)
    preco_final_siva = {}
    desconto_ts_energia = obter_constante('Desconto TS Energia', CONSTANTES) if tarifa_social else 0
    for p, comp_comerc in comerc_energia.items():
        preco_final_siva[p] = comp_comerc + tar_energia_reg.get(p, 0) - desconto_ts_energia + financiamento_tse_a_somar

    desconto_ts_potencia_bruto = obter_constante(f'Desconto TS Potencia {potencia}', CONSTANTES) if tarifa_social else 0
    desconto_ts_potencia_aplicado = min(tar_potencia_reg, desconto_ts_potencia_bruto) if tarifa_social else 0.0
    preco_potencia_final_siva = comerc_potencia + tar_potencia_reg - desconto_ts_potencia_aplicado

    # 3. Calcular custos totais com IVA
    consumo_total = sum(consumos_para_calculo.values())
    decomposicao_energia = calcular_custo_energia_com_iva(consumo_total, preco_final_siva.get('S'), {k:v for k,v in preco_final_siva.items() if k!='S'}, dias, potencia, opcao_horaria_ref, consumos_para_calculo, familia_numerosa)
    
    tar_potencia_final_com_ts = tar_potencia_reg - desconto_ts_potencia_aplicado
    decomposicao_potencia = calcular_custo_potencia_com_iva_final(comerc_potencia, tar_potencia_final_com_ts, dias, potencia)
    
    is_billing_month = 28 <= dias <= 31
    decomposicao_taxas = calcular_taxas_adicionais(consumo_total, dias, tarifa_social, valor_dgeg_user, valor_cav_user, "Personalizado", is_billing_month)
    
    custo_total = decomposicao_energia['custo_com_iva'] + decomposicao_potencia['custo_com_iva'] + decomposicao_taxas['custo_com_iva']
    
    # 4. Construir dicionários para os tooltips
    componentes_tooltip_energia = {}
    for p_key in preco_final_siva.keys():
        componentes_tooltip_energia[f'tooltip_energia_{p_key}_comerc_sem_tar'] = comerc_energia.get(p_key, 0.0)
        componentes_tooltip_energia[f'tooltip_energia_{p_key}_tar_bruta'] = tar_energia_reg.get(p_key, 0.0)
        componentes_tooltip_energia[f'tooltip_energia_{p_key}_tse_declarado_incluido'] = flags_pers['tse_incluido']
        componentes_tooltip_energia[f'tooltip_energia_{p_key}_tse_valor_nominal'] = FINANCIAMENTO_TSE_VAL
        componentes_tooltip_energia[f'tooltip_energia_{p_key}_ts_aplicada_flag'] = tarifa_social
        componentes_tooltip_energia[f'tooltip_energia_{p_key}_ts_desconto_valor'] = desconto_ts_energia

    componentes_tooltip_potencia = {
        'tooltip_pot_comerc_sem_tar': comerc_potencia,
        'tooltip_pot_tar_bruta': tar_potencia_reg,
        'tooltip_pot_ts_aplicada': tarifa_social,
        'tooltip_pot_desconto_ts_valor': desconto_ts_potencia_aplicado
    }
    
    componentes_tooltip_total = {
        'tt_cte_energia_siva': decomposicao_energia['custo_sem_iva'],
        'tt_cte_potencia_siva': decomposicao_potencia['custo_sem_iva'],
        'tt_cte_iec_siva': decomposicao_taxas['iec_sem_iva'],
        'tt_cte_dgeg_siva': decomposicao_taxas['dgeg_sem_iva'],
        'tt_cte_cav_siva': decomposicao_taxas['cav_sem_iva'],
        'tt_cte_total_siva': decomposicao_energia['custo_sem_iva'] + decomposicao_potencia['custo_sem_iva'] + decomposicao_taxas['custo_sem_iva'],
        'tt_cte_valor_iva_6_total': decomposicao_energia['valor_iva_6'] + decomposicao_potencia['valor_iva_6'] + decomposicao_taxas['valor_iva_6'],
        'tt_cte_valor_iva_23_total': decomposicao_energia['valor_iva_23'] + decomposicao_potencia['valor_iva_23'] + decomposicao_taxas['valor_iva_23'],
        'tt_cte_subtotal_civa': decomposicao_energia['custo_com_iva'] + decomposicao_potencia['custo_com_iva'] + decomposicao_taxas['custo_com_iva'],
        'tt_cte_desc_finais_valor': 0.0,
        'tt_cte_acres_finais_valor': 0.0,
        **{f"tt_preco_unit_energia_{p}_siva": v for p, v in preco_final_siva.items()},
        'tt_preco_unit_potencia_siva': preco_potencia_final_siva
    }

    return {
        'Total (€)': custo_total,
        'PrecosFinaisSemIVA': preco_final_siva,
        'PrecoPotenciaFinalSemIVA': preco_potencia_final_siva,
        'componentes_tooltip_custo_total_dict': componentes_tooltip_total,
        'componentes_tooltip_energia_dict': componentes_tooltip_energia,
        'componentes_tooltip_potencia_dict': componentes_tooltip_potencia
    }

#Função Tarifário Fixo para comparação
def calcular_detalhes_custo_tarifario_fixo(
    dados_tarifario_linha,
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
    mes_selecionado_pelo_user_str,
    ano_atual_calculo,
    data_inicio_periodo_obj,
    data_fim_periodo_obj,
    FINANCIAMENTO_TSE_VAL,
    VALOR_QUOTA_ACP_MENSAL
):
    """
    Calcula o custo total e os componentes de tooltip para um DADO TARIFÁRIO FIXO.
    """
    try:
        nome_comercializador_para_taxas = str(dados_tarifario_linha.get('comercializador', 'Desconhecido'))
        nome_tarifario_original = str(dados_tarifario_linha['nome'])
        nome_a_exibir_final = nome_tarifario_original

        # --- Obter Preços e Flags do Tarifário para a OPÇÃO HORÁRIA DE CÁLCULO ---
        # Esta parte é crucial: os preços devem ser os corretos para a 'opcao_horaria_para_calculo'
        preco_energia_input_tf = {}
        oh_calc_lower = opcao_horaria_para_calculo.lower()

        if oh_calc_lower == "simples":
            preco_s = dados_tarifario_linha.get('preco_energia_simples')
            if pd.notna(preco_s): preco_energia_input_tf['S'] = float(preco_s)
            else: return None
        elif oh_calc_lower.startswith("bi-horário"):
            preco_v_bi = dados_tarifario_linha.get('preco_energia_vazio_bi')
            preco_f_bi = dados_tarifario_linha.get('preco_energia_fora_vazio')
            if pd.notna(preco_v_bi) and pd.notna(preco_f_bi):
                preco_energia_input_tf['V'] = float(preco_v_bi)
                preco_energia_input_tf['F'] = float(preco_f_bi)
            else: return None
        elif oh_calc_lower.startswith("tri-horário"):
            if pd.notna(dados_tarifario_linha.get('preco_energia_vazio_tri')) and pd.notna(dados_tarifario_linha.get('preco_energia_cheias')) and pd.notna(dados_tarifario_linha.get('preco_energia_ponta')):
                preco_energia_input_tf['V'] = float(dados_tarifario_linha.get('preco_energia_vazio_tri', 0.0))
                preco_energia_input_tf['C'] = float(dados_tarifario_linha.get('preco_energia_cheias', 0.0))
                preco_energia_input_tf['P'] = float(dados_tarifario_linha.get('preco_energia_ponta', 0.0))
            else: return None
        else: return None
        
        # --- NOVO: Define se é um mês de faturação completo DENTRO da função ---
        is_billing_month = 28 <= dias_calculo <= 31

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
            consumo_total_neste_oh,
            dias_calculo,
            tarifa_social_ativa,
            valor_dgeg_user_input,
            valor_cav_user_input,
            nome_comercializador_para_taxas,
            aplica_taxa_fixa_mensal=is_billing_month # Usa a variável que definimos acima
        )

        e_mes_completo_selecionado_calc = is_billing_month # Substituição direta

        # --- Passo 8: Calcular Custo Total Final e aplicar descontos específicos ---
        custo_total_antes_desc_fatura_tf = (
            decomposicao_custo_energia_tf['custo_com_iva'] +
            decomposicao_custo_potencia_tf_calc['custo_com_iva'] +
            decomposicao_taxas_tf['custo_com_iva']
        )

        # Desconto de fatura do Excel
        desconto_fatura_mensal_excel = float(dados_tarifario_linha.get('desconto_fatura_mes', 0.0) or 0.0)
        desconto_fatura_periodo_aplicado = 0.0
        if desconto_fatura_mensal_excel > 0:
            nome_a_exibir_final += f" (INCLUI desc. fat. {desconto_fatura_mensal_excel:.2f}€/mês)"
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
    
            # CALCULAR O CUSTO BRUTO (SEM TARIFA SOCIAL) APENAS PARA ESTE DESCONTO
    
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

            # ### DESCONTO DE 10% ###
            if nome_tarifario_original.startswith("Galp & Continente (-10% DD)"):
                valor_X_desconto_continente_aplicado = (custo_energia_bruto_cIVA['custo_com_iva'] + custo_potencia_bruto_cIVA['custo_com_iva']) * 0.10
                custo_total_final = custo_base_para_continente - valor_X_desconto_continente_aplicado
                nome_a_exibir_final += f" (INCLUI desc. Cont. de {valor_X_desconto_continente_aplicado:.2f}€, s/ desc. Cont.={custo_base_para_continente:.2f}€)"
            
            # ### DESCONTO DE 7% ###
            elif nome_tarifario_original.startswith("Galp & Continente (-7% s/DD)"):
                valor_X_desconto_continente_aplicado = (custo_energia_bruto_cIVA['custo_com_iva'] + custo_potencia_bruto_cIVA['custo_com_iva']) * 0.07
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
        
        # Para ter a certeza que retornamos algo, vamos simplificar o retorno por agora
        return {
            'Total (€)': custo_total_final,
            'NomeParaExibirAjustado': nome_a_exibir_final,
            'componentes_tooltip_custo_total_dict': componentes_tooltip_total_dict,
            'componentes_tooltip_energia_dict': componentes_tooltip_energia_dict,
            'componentes_tooltip_potencia_dict': componentes_tooltip_potencia_dict
        }
    
    except Exception as e:
        st.error(f"!!! ERRO DENTRO de `calcular_detalhes_custo_tarifario_fixo` para '{dados_tarifario_linha.get('nome', 'Desconhecido')}' na opção '{opcao_horaria_para_calculo}':")
        st.exception(e) # Isto vai imprimir o traceback completo do erro
        return None
    
#Função Tarifário Indexado para comparação
def calcular_detalhes_custo_tarifario_indexado(
    dados_tarifario_indexado_linha,
    opcao_horaria_para_calculo, 
    opcao_horaria_principal_global,
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
    todos_omie_inputs_user_global, 
    omie_medios_calculados_para_todos_ciclos_global,
    omie_medio_simples_real_kwh_para_luzigas_idx,
    dias_no_mes_selecionado_dict,
    mes_selecionado_pelo_user_str,
    ano_atual_calculo,
    data_inicio_periodo_obj,
    data_fim_periodo_obj,
    FINANCIAMENTO_TSE_VAL
):
    try:
        nome_tarifario_original = str(dados_tarifario_indexado_linha['nome'])
        tipo_tarifario_original = str(dados_tarifario_indexado_linha['tipo'])
        formula_energia_str = str(dados_tarifario_indexado_linha.get('formula_calculo', ''))
        nome_a_exibir_final = nome_tarifario_original

        precos_energia_base_kwh_nesta_oh = {} # Preços base calculados para a opcao_horaria_para_calculo
        oh_calc_lower = opcao_horaria_para_calculo.lower() # ex: "simples", "bi-horário - ciclo diário"
        constantes_dict_local = dict(zip(CONSTANTES_df["constante"], CONSTANTES_df["valor_unitário"]))
        
        # --- Define se é um mês de faturação completo DENTRO da função ---
        is_billing_month = 28 <= dias_calculo <= 31

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
            consumo_total_neste_oh_idx, 
            dias_calculo, 
            tarifa_social_ativa,
            valor_dgeg_user_input, 
            valor_cav_user_input,
            nome_comercializador_atual=str(dados_tarifario_indexado_linha.get('comercializador')),
            aplica_taxa_fixa_mensal=is_billing_month
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
            nome_a_exibir_final += f" (INCLUI desc. fat. {desconto_fatura_mensal_idx_excel:.2f}€/mês)"
            if is_billing_month:
                desconto_fatura_periodo_aplicado_idx = desconto_fatura_mensal_idx_excel
            else:
                desconto_fatura_periodo_aplicado_idx = (desconto_fatura_mensal_idx_excel / 30.0) * dias_calculo

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
    opcoes_destino_db_nomes_list
):
    """
    Prepara os dicionários de consumo para cada opção horária de destino,
    convertendo os consumos manuais da opção principal para as estruturas de destino.
    """
    consumos_para_calculo_por_oh_destino = {}
    oh_principal_lower = opcao_horaria_principal_str.lower()

    # Ler os consumos da opção principal selecionada pelo utilizador
    c_s_in = float(consumos_input_atuais_dict.get('S', 0))
    c_v_in = float(consumos_input_atuais_dict.get('V', 0))
    c_f_in = float(consumos_input_atuais_dict.get('F', 0)) # Para Bi
    c_c_in = float(consumos_input_atuais_dict.get('C', 0)) # Para Tri
    c_p_in = float(consumos_input_atuais_dict.get('P', 0)) # Para Tri

    for oh_destino_str in opcoes_destino_db_nomes_list:
        oh_destino_lower = oh_destino_str.lower()
        consumos_finais_para_este_destino = {}

        # Calcular os consumos para a opção de destino com base na opção principal
        if oh_destino_lower == "simples":
            if oh_principal_lower.startswith("simples"):
                consumos_finais_para_este_destino['S'] = c_s_in
            elif oh_principal_lower.startswith("bi-horário"):
                consumos_finais_para_este_destino['S'] = c_v_in + c_f_in
            elif oh_principal_lower.startswith("tri-horário"):
                consumos_finais_para_este_destino['S'] = c_v_in + c_c_in + c_p_in

        elif oh_destino_lower.startswith("bi-horário"):
            # A conversão para Bi-horário só é necessária se a origem for Tri-horário ou Bi-horário
            if oh_principal_lower.startswith("tri-horário"):
                consumos_finais_para_este_destino['V'] = c_v_in
                consumos_finais_para_este_destino['F'] = c_c_in + c_p_in
            elif oh_principal_lower.startswith("bi-horário"):
                consumos_finais_para_este_destino['V'] = c_v_in
                consumos_finais_para_este_destino['F'] = c_f_in
        
        elif oh_destino_lower.startswith("tri-horário"):
             # A conversão para Tri-horário só é possível se a origem também for Tri-horário
             if oh_principal_lower.startswith("tri-horário"):
                consumos_finais_para_este_destino['V'] = c_v_in
                consumos_finais_para_este_destino['C'] = c_c_in
                consumos_finais_para_este_destino['P'] = c_p_in
        
        # Adicionar ao dicionário final apenas se houver consumos calculados
        if consumos_finais_para_este_destino and sum(v for v in consumos_finais_para_este_destino.values() if v is not None) > 0:
            consumos_para_calculo_por_oh_destino[oh_destino_str] = consumos_finais_para_este_destino
            
    return consumos_para_calculo_por_oh_destino
#FIM DETERMINAÇÃO DE OPÇÕES HORÁRIAS

#DETERMINAÇÃO DE OPÇÕES HORÁRIAS
def determinar_opcoes_horarias_destino_e_ordenacao(
    opcao_horaria_principal_str,
    potencia_kva_num,
    opcoes_horarias_existentes_lista,
    is_file_loaded: bool
):
    """
    Determina as opções horárias de destino para a tabela de comparação,
    seguindo as regras específicas para o modo com e sem ficheiro da E-Redes.
    """
    oh_principal_lower = opcao_horaria_principal_str.lower()
    destino_cols_nomes_unicos = []
    coluna_ordenacao_inicial_aggrid = None

    # Nomes EXATOS da base de dados (Excel) para consistência
    SIMPLES_DB = "Simples"
    BI_DIARIO_DB = "Bi-horário - Ciclo Diário"
    BI_SEMANAL_DB = "Bi-horário - Ciclo Semanal"
    TRI_DIARIO_DB = "Tri-horário - Ciclo Diário"
    TRI_SEMANAL_DB = "Tri-horário - Ciclo Semanal"
    TRI_DIARIO_ALTA_DB = "Tri-horário > 20.7 kVA - Ciclo Diário"
    TRI_SEMANAL_ALTA_DB = "Tri-horário > 20.7 kVA - Ciclo Semanal"
    
    opcoes_bi_horario = [BI_DIARIO_DB, BI_SEMANAL_DB]
    opcoes_tri_horario_normal = [TRI_DIARIO_DB, TRI_SEMANAL_DB]
    opcoes_tri_horario_alta = [TRI_DIARIO_ALTA_DB, TRI_SEMANAL_ALTA_DB]

    if is_file_loaded:
        # Lógica para QUANDO HÁ ficheiro (esta parte MANTÉM-SE IGUAL)
        if potencia_kva_num > 20.7:
            destino_cols_nomes_unicos.extend(opcoes_tri_horario_alta)
        else: # <= 20.7 kVA
            destino_cols_nomes_unicos.append(SIMPLES_DB)
            destino_cols_nomes_unicos.extend(opcoes_bi_horario)
            destino_cols_nomes_unicos.extend(opcoes_tri_horario_normal)
    else:
        # LÓGICA NOVA (SEM FICHEIRO) - Implementa as suas 4 regras
        if oh_principal_lower.startswith("tri-horário > 20.7 kva"):
            # Regra 4: se Opção Horária e Ciclo = Tri-horário > 20.7 kVA -> apenas Tri-horário.
            destino_cols_nomes_unicos.extend(opcoes_tri_horario_alta)
        
        elif oh_principal_lower.startswith("tri-horário"):
            # Regra 3: se Opção Horária e Ciclo = Tri-horário <= 20.7 kVA -> Simples, Bi-horário e Tri-horário.
            destino_cols_nomes_unicos.append(SIMPLES_DB)
            destino_cols_nomes_unicos.extend(opcoes_bi_horario)
            destino_cols_nomes_unicos.extend(opcoes_tri_horario_normal)
            
        elif oh_principal_lower.startswith("bi-horário"):
            # Regra 2: se Opção Horária e Ciclo = Bi-horário -> Simples e Bi-horário.
            destino_cols_nomes_unicos.append(SIMPLES_DB)
            destino_cols_nomes_unicos.extend(opcoes_bi_horario)
            
        elif oh_principal_lower == "simples":
            # Regra 1: se Opção Horária e Ciclo = Simples -> apenas Simples.
            destino_cols_nomes_unicos.append(SIMPLES_DB)

    # Filtrar para garantir que apenas opções que realmente existem no ficheiro Excel são incluídas
    destino_cols_nomes_unicos = [
        opt for opt in destino_cols_nomes_unicos if opt in opcoes_horarias_existentes_lista
    ]

    # Ordenar as opções de destino encontradas de forma consistente
    ordem_preferencial = {
        SIMPLES_DB: 0, BI_DIARIO_DB: 1, BI_SEMANAL_DB: 2,
        TRI_DIARIO_DB: 3, TRI_SEMANAL_DB: 4,
        TRI_DIARIO_ALTA_DB: 5, TRI_SEMANAL_ALTA_DB: 6
    }
    destino_cols_nomes_unicos = sorted(
        list(set(destino_cols_nomes_unicos)), # Garante unicidade e ordena
        key=lambda x: ordem_preferencial.get(x, 99)
    )

    # Definir a coluna de ordenação inicial da tabela
    if opcao_horaria_principal_str in destino_cols_nomes_unicos:
        coluna_ordenacao_inicial_aggrid = f"Total {opcao_horaria_principal_str} (€)"
    elif destino_cols_nomes_unicos: # Se a lista não estiver vazia, ordena pela primeira coluna disponível
        coluna_ordenacao_inicial_aggrid = f"Total {destino_cols_nomes_unicos[0]} (€)"
    
    colunas_formatadas = [f"Total {op} (€)" for op in destino_cols_nomes_unicos]

    return destino_cols_nomes_unicos, colunas_formatadas, coluna_ordenacao_inicial_aggrid
    
###############################################################
######################### AUTOCONSUMO #########################
###############################################################
def interpolar_perfis_para_quarto_horario(perfis_horarios):
    """
    Converte perfis horários em quarto-horários usando interpolação linear
    que preserva a energia total e cria uma curva de produção suave.
    """
    perfis_quarto_horarios = {}
    for distrito, perfis_mensais in perfis_horarios.items():
        perfis_quarto_horarios[distrito] = {}
        for mes, perfil_hora in perfis_mensais.items():
            novo_perfil_mes = {}
            horas_ordenadas = sorted(perfil_hora.keys())

            for i, hora_atual in enumerate(horas_ordenadas):
                # Obter os valores da hora anterior, atual e seguinte para calcular a tendência
                valor_anterior = perfil_hora.get(hora_atual - 1, 0)
                valor_atual = perfil_hora[hora_atual]
                valor_seguinte = perfil_hora.get(hora_atual + 1, 0)

                # Se for a primeira ou última hora de produção, a tendência é mais simples
                if i == 0: # Nascer do sol
                    valor_anterior = 0
                if i == len(horas_ordenadas) - 1: # Pôr do sol
                    valor_seguinte = 0

                # Calcular a "taxa de produção" no início e no fim da hora atual
                # A taxa no início da hora é a média entre a hora anterior e a atual
                taxa_inicio_hora = (valor_anterior + valor_atual) / 2.0
                # A taxa no fim da hora é a média entre a hora atual e a seguinte
                taxa_fim_hora = (valor_atual + valor_seguinte) / 2.0
                
                # Com base nestas taxas, calculamos a produção para cada intervalo de 15 minutos
                # usando a fórmula da área de um trapézio, o que garante uma interpolação linear.
                # O fator 0.25 representa o intervalo de 15 minutos (1/4 de hora).
                p00 = (taxa_inicio_hora + (taxa_inicio_hora * 0.75 + taxa_fim_hora * 0.25)) / 2.0 * 0.25
                p15 = ((taxa_inicio_hora * 0.75 + taxa_fim_hora * 0.25) + (taxa_inicio_hora * 0.5 + taxa_fim_hora * 0.5)) / 2.0 * 0.25
                p30 = ((taxa_inicio_hora * 0.5 + taxa_fim_hora * 0.5) + (taxa_inicio_hora * 0.25 + taxa_fim_hora * 0.75)) / 2.0 * 0.25
                p45 = ((taxa_inicio_hora * 0.25 + taxa_fim_hora * 0.75) + taxa_fim_hora) / 2.0 * 0.25
                
                # A soma de p00, p15, p30, p45 será muito próxima do valor_atual original.
                # Corrigimos para garantir que a energia é exatamente conservada.
                soma_calculada = p00 + p15 + p30 + p45
                if soma_calculada > 0:
                    fator_correcao = valor_atual / soma_calculada
                    novo_perfil_mes[(hora_atual, 0)] = p00 * fator_correcao
                    novo_perfil_mes[(hora_atual, 15)] = p15 * fator_correcao
                    novo_perfil_mes[(hora_atual, 30)] = p30 * fator_correcao
                    novo_perfil_mes[(hora_atual, 45)] = p45 * fator_correcao

            perfis_quarto_horarios[distrito][mes] = novo_perfil_mes
            
    return perfis_quarto_horarios


def simular_autoconsumo_completo(df_consumos, potencia_kwp, distrito, inclinacao, orientacao_str):
    """
    Função completa e rigorosa para simular a produção solar, usando:
    1. Dados de produção diária média por mês do PVGIS para cada distrito.
    2. Perfis de distribuição horária distintos para cada mês e para cada distrito.
    """

    # Fonte: PVGIS PVdata. Produção diária média (kWh) para um sistema de 1 kWp otimizado.
    DADOS_PVGIS_DISTRITO = {
        'Aveiro': {1:3.55, 2:4.43, 3:4.79, 4:5.34, 5:5.66, 6:5.66, 7:5.97, 8:5.92, 9:5.44, 10:4.31, 11:3.53, 12:3.34},
        'Beja': {1: 4.09, 2: 4.76, 3: 5.18, 4: 5.46, 5: 5.74, 6: 5.89, 7: 6.29, 8: 6.2, 9: 5.69, 10: 4.8, 11: 4.13, 12: 3.78},
        'Braga': {1: 3.18, 2: 4.01, 3: 4.45, 4: 4.91, 5: 5.35, 6: 5.43, 7: 5.88, 8: 5.84, 9: 5.25, 10: 3.98, 11: 3.18, 12: 2.95},
        'Bragança': {1: 3.22, 2: 4.41, 3: 4.89, 4: 5.25, 5: 5.73, 6: 5.92, 7: 6.41, 8: 6.26, 9: 5.62, 10: 4.28, 11: 3.37, 12: 2.88},
        'Castelo Branco': {1: 3.8, 2: 4.68, 3: 5.06, 4: 5.45, 5: 5.78, 6: 5.94, 7: 6.32, 8: 6.23, 9: 5.64, 10: 4.5, 11: 3.73, 12: 3.48},
        'Coimbra': {1: 3.27, 2: 4.12, 3: 4.5, 4: 4.92, 5: 5.37, 6: 5.37, 7: 5.9, 8: 5.86, 9: 5.31, 10: 4.13, 11: 3.29, 12: 3.05},
        'Évora': {1: 4.03, 2: 4.73, 3: 5.13, 4: 5.37, 5: 5.75, 6: 5.88, 7: 6.3, 8: 6.23, 9: 5.66, 10: 4.66, 11: 4.01, 12: 3.73},
        'Faro': {1: 4.56, 2: 5.11, 3: 5.62, 4: 6.02, 5: 6.26, 6: 6.31, 7: 6.52, 8: 6.42, 9: 6.07, 10: 5.14, 11: 4.51, 12: 4.19},
        'Guarda': {1: 3.5, 2: 4.4, 3: 4.89, 4: 5.23, 5: 5.68, 6: 5.91, 7: 6.45, 8: 6.33, 9: 5.59, 10: 4.37, 11: 3.45, 12: 3.17},
        'Leiria': {1: 3.53, 2: 4.32, 3: 4.72, 4: 5.1, 5: 5.5, 6: 5.52, 7: 5.98, 8: 6.0, 9: 5.47, 10: 4.32, 11: 3.53, 12: 3.32},
        'Lisboa': {1: 3.47, 2: 4.33, 3: 4.96, 4: 5.43, 5: 5.83, 6: 5.93, 7: 6.32, 8: 6.33, 9: 5.77, 10: 4.46, 11: 3.52, 12: 3.23},
        'Portalegre': {1: 3.83, 2: 4.61, 3: 5.0, 4: 5.31, 5: 5.73, 6: 5.93, 7: 6.39, 8: 6.3, 9: 5.6, 10: 4.54, 11: 3.76, 12: 3.54},
        'Porto': {1: 3.37, 2: 4.31, 3: 4.71, 4: 5.34, 5: 5.74, 6: 5.77, 7: 6.08, 8: 5.96, 9: 5.51, 10: 4.24, 11: 3.4, 12: 3.18},
        'Santarém': {1: 3.79, 2: 4.58, 3: 5.09, 4: 5.42, 5: 5.76, 6: 5.87, 7: 6.26, 8: 6.27, 9: 5.71, 10: 4.55, 11: 3.73, 12: 3.54},
        'Setúbal': {1: 3.9, 2: 4.61, 3: 5.19, 4: 5.58, 5: 5.91, 6: 5.99, 7: 6.31, 8: 6.34, 9: 5.88, 10: 4.7, 11: 3.91, 12: 3.66},
        'Viana do Castelo': {1: 3.17, 2: 4.15, 3: 4.62, 4: 5.32, 5: 5.71, 6: 5.81, 7: 6.11, 8: 5.96, 9: 5.47, 10: 4.1, 11: 3.28, 12: 2.95},
        'Vila Real': {1: 3.07, 2: 4.13, 3: 4.66, 4: 5.03, 5: 5.59, 6: 5.76, 7: 6.32, 8: 6.21, 9: 5.57, 10: 4.11, 11: 3.1, 12: 2.76},
        'Viseu': {1: 3.52, 2: 4.31, 3: 4.74, 4: 5.06, 5: 5.51, 6: 5.61, 7: 6.2, 8: 6.13, 9: 5.46, 10: 4.29, 11: 3.43, 12: 3.35},
        'Açores (Ponta Delgada)': {1: 2.9, 2: 3.59, 3: 4.23, 4: 4.63, 5: 4.98, 6: 5.03, 7: 5.11, 8: 5.29, 9: 4.81, 10: 3.8, 11: 3.04, 12: 2.52},
        'Madeira (Funchal)': {1: 3.74, 2: 4.08, 3: 4.8, 4: 4.79, 5: 4.87, 6: 4.64, 7: 5.2, 8: 5.32, 9: 4.48, 10: 4.18, 11: 3.66, 12: 3.49}
    }
    
    COORDENADAS_DISTRITOS = {
        'Guarda': (40.537, -7.268),
        'Aveiro': (40.641, -8.654),
        'Beja': (38.015, -7.863),
        'Braga': (41.545, -8.427),
        'Bragança': (41.806, -6.757),
        'Castelo Branco': (39.822, -7.492),
        'Coimbra': (40.212, -8.429),
        'Évora': (38.567, -7.900),
        'Faro': (37.019, -7.930),
        'Leiria': (39.744, -8.807),
        'Lisboa': (38.717, -9.140),
        'Portalegre': (39.292, -7.429),
        'Porto': (41.150, -8.611),
        'Santarém': (39.236, -8.685),
        'Setúbal': (38.526, -8.890),
        'Viana do Castelo': (41.692, -8.835),
        'Vila Real': (41.301, -7.749),
        'Viseu': (40.657, -7.912),
        'Açores': (37.741, -25.676),
        'Madeira': (32.667, -16.924)
    }

    # Perfis horários por distrito PVGIS PVdata
    PERFIS_HORARIOS_MENSAIS_POR_DISTRITO = {
        'Aveiro': {
            1: {8: 0.0032, 9: 0.0677, 10: 0.1121, 11: 0.1446, 12: 0.1616, 13: 0.1589, 14: 0.1434, 15: 0.1182, 16: 0.0777, 17: 0.0126},
            2: {8: 0.0206, 9: 0.0703, 10: 0.1083, 11: 0.1357, 12: 0.1479, 13: 0.1497, 14: 0.1344, 15: 0.1152, 16: 0.0814, 17: 0.0364},
            3: {7: 0.0035, 8: 0.0374, 9: 0.0761, 10: 0.1056, 11: 0.1263, 12: 0.1359, 13: 0.1418, 14: 0.1314, 15: 0.1118, 16: 0.0815, 17: 0.0432, 18: 0.0055},
            4: {6: 0.0006, 7: 0.0150, 8: 0.0484, 9: 0.0807, 10: 0.1060, 11: 0.1221, 12: 0.1321, 13: 0.1337, 14: 0.1240, 15: 0.1064, 16: 0.0770, 17: 0.0433, 18: 0.0106, 19: 0.0001},
            5: {6: 0.0044, 7: 0.0207, 8: 0.0512, 9: 0.0794, 10: 0.1016, 11: 0.1190, 12: 0.1302, 13: 0.1294, 14: 0.1209, 15: 0.1050, 16: 0.0782, 17: 0.0448, 18: 0.0133, 19: 0.0019},
            6: {6: 0.0061, 7: 0.0207, 8: 0.0472, 9: 0.0743, 10: 0.0962, 11: 0.1156, 12: 0.1287, 13: 0.1304, 14: 0.1234, 15: 0.1072, 16: 0.0809, 17: 0.0477, 18: 0.0169, 19: 0.0045},
            7: {6: 0.0041, 7: 0.0172, 8: 0.0429, 9: 0.0710, 10: 0.0939, 11: 0.1160, 12: 0.1296, 13: 0.1323, 14: 0.1261, 15: 0.1101, 16: 0.0842, 17: 0.0506, 18: 0.0178, 19: 0.0042},
            8: {6: 0.0010, 7: 0.0142, 8: 0.0420, 9: 0.0714, 10: 0.0983, 11: 0.1202, 12: 0.1334, 13: 0.1365, 14: 0.1278, 15: 0.1096, 16: 0.0832, 17: 0.0476, 18: 0.0137, 19: 0.0011},
            9: {7: 0.0121, 8: 0.0456, 9: 0.0770, 10: 0.1067, 11: 0.1276, 12: 0.1376, 13: 0.1402, 14: 0.1283, 15: 0.1080, 16: 0.0754, 17: 0.0372, 18: 0.0043},
            10: {7: 0.0049, 8: 0.0468, 9: 0.0849, 10: 0.1193, 11: 0.1379, 12: 0.1458, 13: 0.1435, 14: 0.1262, 15: 0.1037, 16: 0.0669, 17: 0.0203},
            11: {8: 0.0352, 9: 0.0859, 10: 0.1252, 11: 0.1508, 12: 0.1586, 13: 0.1541, 14: 0.1291, 15: 0.1032, 16: 0.0576, 17: 0.0005},
            12: {8: 0.0065, 9: 0.0765, 10: 0.1232, 11: 0.1540, 12: 0.1677, 13: 0.1600, 14: 0.1394, 15: 0.1109, 16: 0.0617}
        },
        'Beja': {
            1: {8: 0.0165, 9: 0.0739, 10: 0.1156, 11: 0.1404, 12: 0.1559, 13: 0.1518, 14: 0.1378, 15: 0.1136, 16: 0.0775, 17: 0.0172},
            2: {8: 0.0281, 9: 0.0751, 10: 0.1121, 11: 0.1349, 12: 0.1415, 13: 0.1468, 14: 0.1341, 15: 0.1118, 16: 0.0789, 17: 0.0367},
            3: {7: 0.0045, 8: 0.0407, 9: 0.0794, 10: 0.1091, 11: 0.1296, 12: 0.1357, 13: 0.1391, 14: 0.1263, 15: 0.1098, 16: 0.0797, 17: 0.0411, 18: 0.0050},
            4: {6: 0.0006, 7: 0.0170, 8: 0.0534, 9: 0.0874, 10: 0.1096, 11: 0.1244, 12: 0.1319, 13: 0.1290, 14: 0.1196, 15: 0.1018, 16: 0.0748, 17: 0.0415, 18: 0.0090},
            5: {6: 0.0044, 7: 0.0224, 8: 0.0570, 9: 0.0886, 10: 0.1073, 11: 0.1228, 12: 0.1298, 13: 0.1276, 14: 0.1149, 15: 0.0981, 16: 0.0730, 17: 0.0417, 18: 0.0113, 19: 0.0011},
            6: {6: 0.0057, 7: 0.0214, 8: 0.0544, 9: 0.0837, 10: 0.1054, 11: 0.1193, 12: 0.1273, 13: 0.1262, 14: 0.1176, 15: 0.1013, 16: 0.0760, 17: 0.0448, 18: 0.0134, 19: 0.0035},
            7: {6: 0.0038, 7: 0.0170, 8: 0.0493, 9: 0.0804, 10: 0.1047, 11: 0.1213, 12: 0.1299, 13: 0.1289, 14: 0.1202, 15: 0.1033, 16: 0.0781, 17: 0.0462, 18: 0.0137, 19: 0.0031},
            8: {6: 0.0010, 7: 0.0152, 8: 0.0491, 9: 0.0816, 10: 0.1074, 11: 0.1238, 12: 0.1326, 13: 0.1311, 14: 0.1214, 15: 0.1035, 16: 0.0775, 17: 0.0441, 18: 0.0111, 19: 0.0006},
            9: {7: 0.0147, 8: 0.0525, 9: 0.0858, 10: 0.1144, 11: 0.1303, 12: 0.1360, 13: 0.1353, 14: 0.1205, 15: 0.1016, 16: 0.0714, 17: 0.0343, 18: 0.0033},
            10: {7: 0.0086, 8: 0.0537, 9: 0.0908, 10: 0.1221, 11: 0.1378, 12: 0.1426, 13: 0.1412, 14: 0.1238, 15: 0.0981, 16: 0.0624, 17: 0.0190},
            11: {8: 0.0434, 9: 0.0912, 10: 0.1256, 11: 0.1518, 12: 0.1552, 13: 0.1479, 14: 0.1253, 15: 0.1003, 16: 0.0576, 17: 0.0016},
            12: {8: 0.0219, 9: 0.0803, 10: 0.1199, 11: 0.1500, 12: 0.1581, 13: 0.1562, 14: 0.1364, 15: 0.1122, 16: 0.0650}
        },
        'Braga': {
            1: {8: 0.0002, 9: 0.0705, 10: 0.1186, 11: 0.1497, 12: 0.1644, 13: 0.1509, 14: 0.1422, 15: 0.1171, 16: 0.0762, 17: 0.0102},
            2: {8: 0.0191, 9: 0.0739, 10: 0.1133, 11: 0.1389, 12: 0.1459, 13: 0.1468, 14: 0.1315, 15: 0.1130, 16: 0.0809, 17: 0.0367},
            3: {7: 0.0038, 8: 0.0406, 9: 0.0798, 10: 0.1107, 11: 0.1302, 12: 0.1363, 13: 0.1364, 14: 0.1245, 15: 0.1090, 16: 0.0807, 17: 0.0426, 18: 0.0055},
            4: {6: 0.0008, 7: 0.0173, 8: 0.0528, 9: 0.0859, 10: 0.1079, 11: 0.1229, 12: 0.1315, 13: 0.1298, 14: 0.1171, 15: 0.1032, 16: 0.0767, 17: 0.0431, 18: 0.0109, 19: 0.0001},
            5: {6: 0.0050, 7: 0.0230, 8: 0.0561, 9: 0.0851, 10: 0.1039, 11: 0.1196, 12: 0.1249, 13: 0.1250, 14: 0.1171, 15: 0.1023, 16: 0.0771, 17: 0.0451, 18: 0.0138, 19: 0.0021},
            6: {6: 0.0067, 7: 0.0231, 8: 0.0530, 9: 0.0811, 10: 0.1011, 11: 0.1143, 12: 0.1236, 13: 0.1246, 14: 0.1185, 15: 0.1038, 16: 0.0794, 17: 0.0488, 18: 0.0171, 19: 0.0046},
            7: {6: 0.0046, 7: 0.0187, 8: 0.0482, 9: 0.0764, 10: 0.0977, 11: 0.1156, 12: 0.1254, 13: 0.1265, 14: 0.1220, 15: 0.1078, 16: 0.0833, 17: 0.0513, 18: 0.0181, 19: 0.0043},
            8: {6: 0.0013, 7: 0.0158, 8: 0.0472, 9: 0.0778, 10: 0.1025, 11: 0.1208, 12: 0.1296, 13: 0.1304, 14: 0.1224, 15: 0.1076, 16: 0.0817, 17: 0.0480, 18: 0.0137, 19: 0.0012},
            9: {7: 0.0139, 8: 0.0512, 9: 0.0842, 10: 0.1116, 11: 0.1281, 12: 0.1340, 13: 0.1346, 14: 0.1223, 15: 0.1040, 16: 0.0749, 17: 0.0370, 18: 0.0043},
            10: {7: 0.0044, 8: 0.0514, 9: 0.0897, 10: 0.1242, 11: 0.1385, 12: 0.1434, 13: 0.1418, 14: 0.1240, 15: 0.1006, 16: 0.0634, 17: 0.0188},
            11: {8: 0.0315, 9: 0.0877, 10: 0.1316, 11: 0.1563, 12: 0.1593, 13: 0.1495, 14: 0.1266, 15: 0.1014, 16: 0.0560, 17: 0.0002},
            12: {8: 0.0005, 9: 0.0782, 10: 0.1280, 11: 0.1573, 12: 0.1666, 13: 0.1603, 14: 0.1377, 15: 0.1112, 16: 0.0602}
        },
        'Bragança': {
            1: {8: 0.0094, 9: 0.0730, 10: 0.1162, 11: 0.1449, 12: 0.1579, 13: 0.1598, 14: 0.1474, 15: 0.1185, 16: 0.0725, 17: 0.0005},
            2: {8: 0.0286, 9: 0.0756, 10: 0.1140, 11: 0.1374, 12: 0.1440, 13: 0.1484, 14: 0.1348, 15: 0.1098, 16: 0.0774, 17: 0.0302},
            3: {7: 0.0062, 8: 0.0446, 9: 0.0827, 10: 0.1120, 11: 0.1326, 12: 0.1373, 13: 0.1392, 14: 0.1241, 15: 0.1054, 16: 0.0762, 17: 0.0377, 18: 0.0021},
            4: {6: 0.0013, 7: 0.0210, 8: 0.0565, 9: 0.0888, 10: 0.1115, 11: 0.1257, 12: 0.1302, 13: 0.1297, 14: 0.1198, 15: 0.0969, 16: 0.0720, 17: 0.0388, 18: 0.0078},
            5: {6: 0.0056, 7: 0.0265, 8: 0.0606, 9: 0.0904, 10: 0.1100, 11: 0.1227, 12: 0.1288, 13: 0.1254, 14: 0.1116, 15: 0.0959, 16: 0.0706, 17: 0.0397, 18: 0.0107, 19: 0.0013},
            6: {5: 0.0001, 6: 0.0070, 7: 0.0263, 8: 0.0594, 9: 0.0881, 10: 0.1077, 11: 0.1196, 12: 0.1255, 13: 0.1235, 14: 0.1133, 15: 0.0982, 16: 0.0726, 17: 0.0417, 18: 0.0134, 19: 0.0036},
            7: {6: 0.0050, 7: 0.0216, 8: 0.0553, 9: 0.0850, 10: 0.1060, 11: 0.1212, 12: 0.1283, 13: 0.1262, 14: 0.1164, 15: 0.0999, 16: 0.0743, 17: 0.0439, 18: 0.0138, 19: 0.0032},
            8: {6: 0.0020, 7: 0.0192, 8: 0.0542, 9: 0.0863, 10: 0.1092, 11: 0.1243, 12: 0.1302, 13: 0.1291, 14: 0.1188, 15: 0.1001, 16: 0.0742, 17: 0.0414, 18: 0.0103, 19: 0.0006},
            9: {7: 0.0179, 8: 0.0563, 9: 0.0890, 10: 0.1165, 11: 0.1318, 12: 0.1362, 13: 0.1329, 14: 0.1182, 15: 0.0996, 16: 0.0682, 17: 0.0310, 18: 0.0024},
            10: {7: 0.0100, 8: 0.0565, 9: 0.0922, 10: 0.1240, 11: 0.1377, 12: 0.1456, 13: 0.1420, 14: 0.1231, 15: 0.0984, 16: 0.0583, 17: 0.0121},
            11: {8: 0.0432, 9: 0.0896, 10: 0.1246, 11: 0.1509, 12: 0.1569, 13: 0.1527, 14: 0.1295, 15: 0.1011, 16: 0.0515},
            12: {8: 0.0126, 9: 0.0806, 10: 0.1185, 11: 0.1516, 12: 0.1632, 13: 0.1598, 14: 0.1416, 15: 0.1142, 16: 0.0578}
        },
        'Castelo Branco': {
            1: {8: 0.0119, 9: 0.0741, 10: 0.1182, 11: 0.1431, 12: 0.1568, 13: 0.1523, 14: 0.1409, 15: 0.1164, 16: 0.0755, 17: 0.0107},
            2: {8: 0.0273, 9: 0.0744, 10: 0.1099, 11: 0.1378, 12: 0.1447, 13: 0.1445, 14: 0.1342, 15: 0.1132, 16: 0.0794, 17: 0.0346},
            3: {7: 0.0050, 8: 0.0417, 9: 0.0804, 10: 0.1118, 11: 0.1325, 12: 0.1377, 13: 0.1380, 14: 0.1257, 15: 0.1064, 16: 0.0773, 17: 0.0396, 18: 0.0038},
            4: {6: 0.0008, 7: 0.0184, 8: 0.0550, 9: 0.0890, 10: 0.1121, 11: 0.1263, 12: 0.1304, 13: 0.1277, 14: 0.1172, 15: 0.1012, 16: 0.0729, 17: 0.0405, 18: 0.0086},
            5: {6: 0.0048, 7: 0.0241, 8: 0.0595, 9: 0.0905, 10: 0.1116, 11: 0.1241, 12: 0.1273, 13: 0.1220, 14: 0.1136, 15: 0.0975, 16: 0.0716, 17: 0.0410, 18: 0.0112, 19: 0.0013},
            6: {6: 0.0061, 7: 0.0230, 8: 0.0562, 9: 0.0856, 10: 0.1071, 11: 0.1209, 12: 0.1280, 13: 0.1245, 14: 0.1140, 15: 0.0987, 16: 0.0748, 17: 0.0438, 18: 0.0138, 19: 0.0035},
            7: {6: 0.0043, 7: 0.0188, 8: 0.0525, 9: 0.0827, 10: 0.1058, 11: 0.1208, 12: 0.1280, 13: 0.1270, 14: 0.1184, 15: 0.1016, 16: 0.0770, 17: 0.0459, 18: 0.0141, 19: 0.0033},
            8: {6: 0.0014, 7: 0.0168, 8: 0.0520, 9: 0.0844, 10: 0.1088, 11: 0.1241, 12: 0.1306, 13: 0.1296, 14: 0.1199, 15: 0.1021, 16: 0.0759, 17: 0.0430, 18: 0.0109, 19: 0.0007},
            9: {7: 0.0158, 8: 0.0540, 9: 0.0880, 10: 0.1155, 11: 0.1300, 12: 0.1363, 13: 0.1335, 14: 0.1197, 15: 0.1002, 16: 0.0702, 17: 0.0338, 18: 0.0030},
            10: {7: 0.0089, 8: 0.0547, 9: 0.0904, 10: 0.1220, 11: 0.1387, 12: 0.1427, 13: 0.1390, 14: 0.1245, 15: 0.0999, 16: 0.0620, 17: 0.0171},
            11: {8: 0.0437, 9: 0.0912, 10: 0.1278, 11: 0.1493, 12: 0.1552, 13: 0.1481, 14: 0.1281, 15: 0.1009, 16: 0.0556, 17: 0.0002},
            12: {8: 0.0195, 9: 0.0830, 10: 0.1237, 11: 0.1497, 12: 0.1587, 13: 0.1568, 14: 0.1389, 15: 0.1093, 16: 0.0604}
        },
        'Coimbra': {
            1: {8: 0.0005, 9: 0.0122, 10: 0.1253, 11: 0.1549, 12: 0.1727, 13: 0.1655, 14: 0.1517, 15: 0.1214, 16: 0.0822, 17: 0.0135},
            2: {8: 0.0047, 9: 0.0646, 10: 0.1153, 11: 0.1395, 12: 0.1526, 13: 0.1522, 14: 0.1361, 15: 0.1137, 16: 0.0832, 17: 0.0380},
            3: {7: 0.0021, 8: 0.0322, 9: 0.0783, 10: 0.1100, 11: 0.1336, 12: 0.1370, 13: 0.1395, 14: 0.1279, 15: 0.1093, 16: 0.0811, 17: 0.0433, 18: 0.0057},
            4: {6: 0.0006, 7: 0.0167, 8: 0.0530, 9: 0.0843, 10: 0.1084, 11: 0.1240, 12: 0.1305, 13: 0.1301, 14: 0.1193, 15: 0.1036, 16: 0.0757, 17: 0.0431, 18: 0.0108},
            5: {6: 0.0049, 7: 0.0222, 8: 0.0537, 9: 0.0839, 10: 0.1041, 11: 0.1224, 12: 0.1286, 13: 0.1264, 14: 0.1166, 15: 0.1014, 16: 0.0757, 17: 0.0449, 18: 0.0132, 19: 0.0019},
            6: {6: 0.0065, 7: 0.0222, 8: 0.0512, 9: 0.0792, 10: 0.0996, 11: 0.1164, 12: 0.1256, 13: 0.1261, 14: 0.1191, 15: 0.1031, 16: 0.0805, 17: 0.0489, 18: 0.0170, 19: 0.0045},
            7: {6: 0.0043, 7: 0.0175, 8: 0.0446, 9: 0.0736, 10: 0.0978, 11: 0.1178, 12: 0.1296, 13: 0.1295, 14: 0.1230, 15: 0.1071, 16: 0.0830, 17: 0.0507, 18: 0.0173, 19: 0.0041},
            8: {6: 0.0011, 7: 0.0147, 8: 0.0435, 9: 0.0747, 10: 0.1018, 11: 0.1217, 12: 0.1314, 13: 0.1327, 14: 0.1244, 15: 0.1093, 16: 0.0825, 17: 0.0478, 18: 0.0135, 19: 0.0010},
            9: {7: 0.0102, 8: 0.0487, 9: 0.0815, 10: 0.1117, 11: 0.1295, 12: 0.1390, 13: 0.1361, 14: 0.1226, 15: 0.1044, 16: 0.0748, 17: 0.0373, 18: 0.0042},
            10: {7: 0.0020, 8: 0.0356, 9: 0.0895, 10: 0.1262, 11: 0.1433, 12: 0.1490, 13: 0.1436, 14: 0.1254, 15: 0.1003, 16: 0.0651, 17: 0.0200},
            11: {8: 0.0074, 9: 0.0786, 10: 0.1338, 11: 0.1588, 12: 0.1641, 13: 0.1602, 14: 0.1328, 15: 0.1042, 16: 0.0596, 17: 0.0004},
            12: {8: 0.0010, 9: 0.0135, 10: 0.1365, 11: 0.1665, 12: 0.1773, 13: 0.1719, 14: 0.1472, 15: 0.1193, 16: 0.0668}
        },
        'Évora': {
            1: {8: 0.0145, 9: 0.0729, 10: 0.1157, 11: 0.1426, 12: 0.1553, 13: 0.1506, 14: 0.1391, 15: 0.1173, 16: 0.0766, 17: 0.0153},
            2: {8: 0.0267, 9: 0.0741, 10: 0.1127, 11: 0.1374, 12: 0.1440, 13: 0.1446, 14: 0.1321, 15: 0.1115, 16: 0.0797, 17: 0.0371},
            3: {7: 0.0044, 8: 0.0406, 9: 0.0805, 10: 0.1101, 11: 0.1312, 12: 0.1375, 13: 0.1376, 14: 0.1265, 15: 0.1085, 16: 0.0769, 17: 0.0412, 18: 0.0048},
            4: {6: 0.0006, 7: 0.0174, 8: 0.0544, 9: 0.0879, 10: 0.1112, 11: 0.1246, 12: 0.1306, 13: 0.1284, 14: 0.1177, 15: 0.1008, 16: 0.0751, 17: 0.0420, 18: 0.0093},
            5: {6: 0.0045, 7: 0.0224, 8: 0.0571, 9: 0.0872, 10: 0.1078, 11: 0.1241, 12: 0.1289, 13: 0.1243, 14: 0.1150, 15: 0.1005, 16: 0.0733, 17: 0.0424, 18: 0.0115, 19: 0.0012},
            6: {6: 0.0058, 7: 0.0217, 8: 0.0546, 9: 0.0831, 10: 0.1050, 11: 0.1195, 12: 0.1277, 13: 0.1264, 14: 0.1163, 15: 0.1002, 16: 0.0768, 17: 0.0452, 18: 0.0140, 19: 0.0036},
            7: {6: 0.0039, 7: 0.0172, 8: 0.0491, 9: 0.0803, 10: 0.1043, 11: 0.1205, 12: 0.1290, 13: 0.1288, 14: 0.1206, 15: 0.1031, 16: 0.0789, 17: 0.0468, 18: 0.0142, 19: 0.0032},
            8: {6: 0.0010, 7: 0.0154, 8: 0.0493, 9: 0.0822, 10: 0.1077, 11: 0.1243, 12: 0.1318, 13: 0.1302, 14: 0.1209, 15: 0.1033, 16: 0.0775, 17: 0.0445, 18: 0.0112, 19: 0.0006},
            9: {7: 0.0145, 8: 0.0520, 9: 0.0867, 10: 0.1144, 11: 0.1302, 12: 0.1370, 13: 0.1341, 14: 0.1193, 15: 0.1020, 16: 0.0721, 17: 0.0344, 18: 0.0034},
            10: {7: 0.0083, 8: 0.0529, 9: 0.0910, 10: 0.1226, 11: 0.1390, 12: 0.1432, 13: 0.1401, 14: 0.1221, 15: 0.0990, 16: 0.0628, 17: 0.0189},
            11: {8: 0.0433, 9: 0.0908, 10: 0.1295, 11: 0.1493, 12: 0.1560, 13: 0.1496, 14: 0.1235, 15: 0.1000, 16: 0.0568, 17: 0.0011},
            12: {8: 0.0200, 9: 0.0810, 10: 0.1238, 11: 0.1487, 12: 0.1603, 13: 0.1550, 14: 0.1356, 15: 0.1104, 16: 0.0652}
        },
        'Faro': {
            1: {8: 0.0184, 9: 0.0738, 10: 0.1156, 11: 0.1449, 12: 0.1557, 13: 0.1475, 14: 0.1375, 15: 0.1123, 16: 0.0749, 17: 0.0194},
            2: {8: 0.0271, 9: 0.0735, 10: 0.1100, 11: 0.1360, 12: 0.1442, 13: 0.1457, 14: 0.1350, 15: 0.1134, 16: 0.0786, 17: 0.0365},
            3: {7: 0.0042, 8: 0.0401, 9: 0.0780, 10: 0.1077, 11: 0.1296, 12: 0.1359, 13: 0.1400, 14: 0.1294, 15: 0.1102, 16: 0.0785, 17: 0.0414, 18: 0.0049},
            4: {6: 0.0004, 7: 0.0148, 8: 0.0491, 9: 0.0823, 10: 0.1071, 11: 0.1264, 12: 0.1333, 13: 0.1346, 14: 0.1223, 15: 0.1050, 16: 0.0761, 17: 0.0403, 18: 0.0083},
            5: {6: 0.0037, 7: 0.0198, 8: 0.0531, 9: 0.0842, 10: 0.1059, 11: 0.1236, 12: 0.1315, 13: 0.1299, 14: 0.1210, 15: 0.1017, 16: 0.0743, 17: 0.0404, 18: 0.0102, 19: 0.0009,},
            6: {6: 0.0049, 7: 0.0191, 8: 0.0513, 9: 0.0801, 10: 0.1037, 11: 0.1200, 12: 0.1300, 13: 0.1295, 14: 0.1210, 15: 0.1039, 16: 0.0769, 17: 0.0441, 18: 0.0125, 19: 0.0031},
            7: {6: 0.0033, 7: 0.0154, 8: 0.0483, 9: 0.0790, 10: 0.1032, 11: 0.1199, 12: 0.1307, 13: 0.1305, 14: 0.1228, 15: 0.1056, 16: 0.0793, 17: 0.0460, 18: 0.0133, 19: 0.0028},
            8: {6: 0.0008, 7: 0.0140, 8: 0.0477, 9: 0.0801, 10: 0.1052, 11: 0.1222, 12: 0.1326, 13: 0.1328, 14: 0.1236, 15: 0.1065, 16: 0.0791, 17: 0.0440, 18: 0.0108, 19: 0.0005},
            9: {7: 0.0135, 8: 0.0498, 9: 0.0834, 10: 0.1096, 11: 0.1268, 12: 0.1377, 13: 0.1385, 14: 0.1238, 15: 0.1050, 16: 0.0733, 17: 0.0352, 18: 0.0033},
            10: {7: 0.0083, 8: 0.0513, 9: 0.0892, 10: 0.1214, 11: 0.1359, 12: 0.1440, 13: 0.1414, 14: 0.1247, 15: 0.1001, 16: 0.0638, 17: 0.0198},
            11: {8: 0.0434, 9: 0.0902, 10: 0.1275, 11: 0.1504, 12: 0.1548, 13: 0.1465, 14: 0.1262, 15: 0.1013, 16: 0.0570, 17: 0.0028},
            12: {8: 0.0237, 9: 0.0829, 10: 0.1219, 11: 0.1484, 12: 0.1600, 13: 0.1566, 14: 0.1367, 15: 0.1073, 16: 0.0624, 17: 0.0001}
        },
        'Guarda': {
            1: {8: 0.0100, 9: 0.0755, 10: 0.1165, 11: 0.1434, 12: 0.1616, 13: 0.1549, 14: 0.1424, 15: 0.1164, 16: 0.0744, 17: 0.0047},
            2: {8: 0.0284, 9: 0.0774, 10: 0.1119, 11: 0.1361, 12: 0.1452, 13: 0.1455, 14: 0.1318, 15: 0.1117, 16: 0.0785, 17: 0.0335},
            3: {7: 0.0054, 8: 0.0432, 9: 0.0818, 10: 0.1110, 11: 0.1328, 12: 0.1367, 13: 0.1376, 14: 0.1265, 15: 0.1057, 16: 0.0774, 17: 0.0395, 18: 0.0025},
            4: {6: 0.0010, 7: 0.0192, 8: 0.0558, 9: 0.0879, 10: 0.1112, 11: 0.1276, 12: 0.1320, 13: 0.1276, 14: 0.1148, 15: 0.1008, 16: 0.0734, 17: 0.0403, 18: 0.0085},
            5: {6: 0.0052, 7: 0.0256, 8: 0.0597, 9: 0.0905, 10: 0.1098, 11: 0.1234, 12: 0.1281, 13: 0.1230, 14: 0.1134, 15: 0.0969, 16: 0.0709, 17: 0.0407, 18: 0.0114, 19: 0.0013},
            6: {6: 0.0065, 7: 0.0244, 8: 0.0569, 9: 0.0862, 10: 0.1071, 11: 0.1204, 12: 0.1264, 13: 0.1238, 14: 0.1153, 15: 0.0972, 16: 0.0740, 17: 0.0441, 18: 0.0140, 19: 0.0037},
            7: {6: 0.0045, 7: 0.0197, 8: 0.0532, 9: 0.0835, 10: 0.1066, 11: 0.1218, 12: 0.1298, 13: 0.1258, 14: 0.1162, 15: 0.1010, 16: 0.0756, 17: 0.0449, 18: 0.0141, 19: 0.0033},
            8: {6: 0.0016, 7: 0.0175, 8: 0.0526, 9: 0.0850, 10: 0.1098, 11: 0.1257, 12: 0.1315, 13: 0.1281, 14: 0.1182, 15: 0.1015, 16: 0.0749, 17: 0.0423, 18: 0.0107, 19: 0.0007},
            9: {7: 0.0167, 8: 0.0557, 9: 0.0903, 10: 0.1158, 11: 0.1324, 12: 0.1378, 13: 0.1330, 14: 0.1167, 15: 0.0973, 16: 0.0689, 17: 0.0327, 18: 0.0027},
            10: {7: 0.0093, 8: 0.0551, 9: 0.0907, 10: 0.1242, 11: 0.1394, 12: 0.1452, 13: 0.1408, 14: 0.1209, 15: 0.0990, 16: 0.0612, 17: 0.0142},
            11: {8: 0.0436, 9: 0.0895, 10: 0.1258, 11: 0.1494, 12: 0.1569, 13: 0.1494, 14: 0.1282, 15: 0.1021, 16: 0.0552, 17: 0.0001},
            12: {8: 0.0150, 9: 0.0803, 10: 0.1234, 11: 0.1509, 12: 0.1625, 13: 0.1576, 14: 0.1389, 15: 0.1113, 16: 0.0599}
        },
        'Leiria': {
            1: {8: 0.0004, 9: 0.0714, 10: 0.1179, 11: 0.1492, 12: 0.1617, 13: 0.1529, 14: 0.1410, 15: 0.1180, 16: 0.0798, 17: 0.0076},
            2: {8: 0.0212, 9: 0.0716, 10: 0.1107, 11: 0.1373, 12: 0.1472, 13: 0.1495, 14: 0.1327, 15: 0.1108, 16: 0.0797, 17: 0.0393},
            3: {7: 0.0029, 8: 0.0372, 9: 0.0756, 10: 0.1060, 11: 0.1294, 12: 0.1365, 13: 0.1419, 14: 0.1272, 15: 0.1105, 16: 0.0827, 17: 0.0442, 18: 0.0059},
            4: {6: 0.0005, 7: 0.0154, 8: 0.0497, 9: 0.0810, 10: 0.1046, 11: 0.1230, 12: 0.1302, 13: 0.1324, 14: 0.1230, 15: 0.1053, 16: 0.0777, 17: 0.0456, 18: 0.0114, 19: 0.0001},
            5: {6: 0.0043, 7: 0.0207, 8: 0.0515, 9: 0.0820, 10: 0.1032, 11: 0.1202, 12: 0.1284, 13: 0.1274, 14: 0.1190, 15: 0.1027, 16: 0.0790, 17: 0.0458, 18: 0.0140, 19: 0.0019},
            6: {6: 0.0060, 7: 0.0209, 8: 0.0483, 9: 0.0756, 10: 0.0975, 11: 0.1150, 12: 0.1277, 13: 0.1267, 14: 0.1214, 15: 0.1082, 16: 0.0820, 17: 0.0491, 18: 0.0174, 19: 0.0044},
            7: {6: 0.0039, 7: 0.0165, 8: 0.0424, 9: 0.0704, 10: 0.0950, 11: 0.1160, 12: 0.1289, 13: 0.1320, 14: 0.1257, 15: 0.1105, 16: 0.0854, 17: 0.0515, 18: 0.0177, 19: 0.0041},
            8: {6: 0.0008, 7: 0.0136, 8: 0.0414, 9: 0.0728, 10: 0.0996, 11: 0.1212, 12: 0.1317, 13: 0.1357, 14: 0.1269, 15: 0.1104, 16: 0.0828, 17: 0.0482, 18: 0.0138, 19: 0.0011},
            9: {7: 0.0117, 8: 0.0450, 9: 0.0781, 10: 0.1089, 11: 0.1284, 12: 0.1380, 13: 0.1366, 14: 0.1245, 15: 0.1081, 16: 0.0774, 17: 0.0387, 18: 0.0046},
            10: {7: 0.0044, 8: 0.0483, 9: 0.0858, 10: 0.1218, 11: 0.1383, 12: 0.1441, 13: 0.1414, 14: 0.1235, 15: 0.1040, 16: 0.0664, 17: 0.0218},
            11: {8: 0.0327, 9: 0.0882, 10: 0.1293, 11: 0.1562, 12: 0.1571, 13: 0.1505, 14: 0.1265, 15: 0.0999, 16: 0.0592, 17: 0.0003},
            12: {8: 0.0010, 9: 0.0777, 10: 0.1250, 11: 0.1547, 12: 0.1678, 13: 0.1585, 14: 0.1391, 15: 0.1112, 16: 0.0650}
        },
        'Lisboa': {
            1: {8: 0.0006, 9: 0.0316, 10: 0.1197, 11: 0.1465, 12: 0.1645, 13: 0.1617, 14: 0.1529, 15: 0.1290, 16: 0.0887, 17: 0.0048},
            2: {8: 0.0046, 9: 0.0700, 10: 0.1085, 11: 0.1340, 12: 0.1429, 13: 0.1493, 14: 0.1405, 15: 0.1225, 16: 0.0866, 17: 0.0410, 18: 0.0001},
            3: {7: 0.0016, 8: 0.0327, 9: 0.0752, 10: 0.1055, 11: 0.1284, 12: 0.1370, 13: 0.1415, 14: 0.1315, 15: 0.1127, 16: 0.0838, 17: 0.0453, 18: 0.0047},
            4: {6: 0.0003, 7: 0.0117, 8: 0.0490, 9: 0.0808, 10: 0.1058, 11: 0.1209, 12: 0.1294, 13: 0.1323, 14: 0.1245, 15: 0.1080, 16: 0.0802, 17: 0.0457, 18: 0.0113},
            5: {6: 0.0036, 7: 0.0191, 8: 0.0515, 9: 0.0813, 10: 0.1020, 11: 0.1198, 12: 0.1284, 13: 0.1295, 14: 0.1215, 15: 0.1048, 16: 0.0783, 17: 0.0452, 18: 0.0134, 19: 0.0017},
            6: {6: 0.0050, 7: 0.0189, 8: 0.0497, 9: 0.0768, 10: 0.0998, 11: 0.1167, 12: 0.1272, 13: 0.1278, 14: 0.1218, 15: 0.1064, 16: 0.0808, 17: 0.0488, 18: 0.0162, 19: 0.0042},
            7: {6: 0.0032, 7: 0.0148, 8: 0.0450, 9: 0.0744, 10: 0.0995, 11: 0.1177, 12: 0.1284, 13: 0.1300, 14: 0.1241, 15: 0.1087, 16: 0.0832, 17: 0.0505, 18: 0.0168, 19: 0.0038},
            8: {6: 0.0006, 7: 0.0117, 8: 0.0451, 9: 0.0766, 10: 0.1019, 11: 0.1212, 12: 0.1316, 13: 0.1330, 14: 0.1253, 15: 0.1088, 16: 0.0820, 17: 0.0477, 18: 0.0135, 19: 0.0010},
            9: {7: 0.0051, 8: 0.0472, 9: 0.0826, 10: 0.1093, 11: 0.1272, 12: 0.1365, 13: 0.1389, 14: 0.1268, 15: 0.1075, 16: 0.0760, 17: 0.0385, 18: 0.0044},
            10: {7: 0.0016, 8: 0.0423, 9: 0.0886, 10: 0.1184, 11: 0.1389, 12: 0.1436, 13: 0.1458, 14: 0.1282, 15: 0.1058, 16: 0.0679, 17: 0.0188},
            11: {8: 0.0070, 9: 0.0887, 10: 0.1293, 11: 0.1517, 12: 0.1599, 13: 0.1538, 14: 0.1348, 15: 0.1100, 16: 0.0641, 17: 0.0006},
            12: {8: 0.0012, 9: 0.0362, 10: 0.1257, 11: 0.1570, 12: 0.1718, 13: 0.1681, 14: 0.1489, 15: 0.1200, 16: 0.0708, 17: 0.0001}
        },
        'Portalegre': {
            1: {8: 0.0020, 9: 0.0768, 10: 0.1194, 11: 0.1465, 12: 0.1592, 13: 0.1529, 14: 0.1403, 15: 0.1146, 16: 0.0757, 17: 0.0126},
            2: {8: 0.0286, 9: 0.0752, 10: 0.1119, 11: 0.1356, 12: 0.1450, 13: 0.1449, 14: 0.1328, 15: 0.1116, 16: 0.0797, 17: 0.0348},
            3: {7: 0.0038, 8: 0.0417, 9: 0.0811, 10: 0.1113, 11: 0.1297, 12: 0.1377, 13: 0.1385, 14: 0.1266, 15: 0.1067, 16: 0.0790, 17: 0.0399, 18: 0.0041},
            4: {6: 0.0008, 7: 0.0185, 8: 0.0543, 9: 0.0876, 10: 0.1105, 11: 0.1259, 12: 0.1298, 13: 0.1291, 14: 0.1170, 15: 0.1014, 16: 0.0754, 17: 0.0408, 18: 0.0087},
            5: {6: 0.0048, 7: 0.0240, 8: 0.0587, 9: 0.0891, 10: 0.1090, 11: 0.1223, 12: 0.1292, 13: 0.1253, 14: 0.1136, 15: 0.0985, 16: 0.0727, 17: 0.0409, 18: 0.0108, 19: 0.0012},
            6: {6: 0.0062, 7: 0.0229, 8: 0.0553, 9: 0.0850, 10: 0.1063, 11: 0.1199, 12: 0.1275, 13: 0.1250, 14: 0.1152, 15: 0.1002, 16: 0.0756, 17: 0.0439, 18: 0.0136, 19: 0.0034},
            7: {6: 0.0042, 7: 0.0184, 8: 0.0514, 9: 0.0818, 10: 0.1056, 11: 0.1207, 12: 0.1291, 13: 0.1282, 14: 0.1194, 15: 0.1020, 16: 0.0770, 17: 0.0455, 18: 0.0136, 19: 0.0030},
            8: {6: 0.0013, 7: 0.0165, 8: 0.0510, 9: 0.0838, 10: 0.1086, 11: 0.1246, 12: 0.1315, 13: 0.1303, 14: 0.1198, 15: 0.1023, 16: 0.0760, 17: 0.0431, 18: 0.0106, 19: 0.0006},
            9: {7: 0.0161, 8: 0.0543, 9: 0.0891, 10: 0.1155, 11: 0.1295, 12: 0.1354, 13: 0.1327, 14: 0.1185, 15: 0.1003, 16: 0.0714, 17: 0.0343, 18: 0.0030},
            10: {7: 0.0076, 8: 0.0552, 9: 0.0914, 10: 0.1244, 11: 0.1389, 12: 0.1434, 13: 0.1388, 14: 0.1217, 15: 0.0978, 16: 0.0626, 17: 0.0180},
            11: {8: 0.0454, 9: 0.0915, 10: 0.1282, 11: 0.1522, 12: 0.1557, 13: 0.1464, 14: 0.1246, 15: 0.0995, 16: 0.0562, 17: 0.0003},
            12: {8: 0.0120, 9: 0.0856, 10: 0.1254, 11: 0.1535, 12: 0.1631, 13: 0.1565, 14: 0.1368, 15: 0.1068, 16: 0.0603}
        },
        'Porto': {
            1: {8: 0.0021, 9: 0.0677, 10: 0.1154, 11: 0.1460, 12: 0.1668, 13: 0.1576, 14: 0.1439, 15: 0.1180, 16: 0.0775, 17: 0.0050},
            2: {8: 0.0206, 9: 0.0707, 10: 0.1095, 11: 0.1378, 12: 0.1482, 13: 0.1465, 14: 0.1347, 15: 0.1130, 16: 0.0821, 17: 0.0369},
            3: {7: 0.0035, 8: 0.0380, 9: 0.0769, 10: 0.1066, 11: 0.1309, 12: 0.1366, 13: 0.1407, 14: 0.1285, 15: 0.1105, 16: 0.0808, 17: 0.0423, 18: 0.0046},
            4: {6: 0.0006, 7: 0.0152, 8: 0.0488, 9: 0.0802, 10: 0.1039, 11: 0.1229, 12: 0.1335, 13: 0.1352, 14: 0.1236, 15: 0.1057, 16: 0.0772, 17: 0.0427, 18: 0.0105, 19: 0.0001},
            5: {6: 0.0044, 7: 0.0208, 8: 0.0514, 9: 0.0800, 10: 0.1006, 11: 0.1194, 12: 0.1295, 13: 0.1312, 14: 0.1220, 15: 0.1032, 16: 0.0775, 17: 0.0445, 18: 0.0134, 19: 0.0020},
            6: {6: 0.0061, 7: 0.0214, 8: 0.0493, 9: 0.0767, 10: 0.0973, 11: 0.1160, 12: 0.1282, 13: 0.1307, 14: 0.1217, 15: 0.1055, 16: 0.0788, 17: 0.0473, 18: 0.0165, 19: 0.0044},
            7: {6: 0.0043, 7: 0.0175, 8: 0.0457, 9: 0.0733, 10: 0.0956, 11: 0.1151, 12: 0.1279, 13: 0.1315, 14: 0.1247, 15: 0.1089, 16: 0.0832, 17: 0.0504, 18: 0.0178, 19: 0.0043},
            8: {6: 0.0011, 7: 0.0147, 8: 0.0446, 9: 0.0749, 10: 0.0996, 11: 0.1198, 12: 0.1316, 13: 0.1339, 14: 0.1265, 15: 0.1087, 16: 0.0820, 17: 0.0475, 18: 0.0138, 19: 0.0012},
            9: {7: 0.0124, 8: 0.0479, 9: 0.0813, 10: 0.1093, 11: 0.1276, 12: 0.1369, 13: 0.1378, 14: 0.1248, 15: 0.1062, 16: 0.0747, 17: 0.0370, 18: 0.0041},
            10: {7: 0.0051, 8: 0.0479, 9: 0.0863, 10: 0.1213, 11: 0.1393, 12: 0.1462, 13: 0.1440, 14: 0.1265, 15: 0.1011, 16: 0.0647, 17: 0.0174},
            11: {8: 0.0326, 9: 0.0838, 10: 0.1271, 11: 0.1553, 12: 0.1577, 13: 0.1527, 14: 0.1306, 15: 0.1036, 16: 0.0564, 17: 0.0002},
            12: {8: 0.0053, 9: 0.0760, 10: 0.1244, 11: 0.1557, 12: 0.1683, 13: 0.1633, 14: 0.1374, 15: 0.1102, 16: 0.0594}
        },
        'Santarém': {
            1: {8: 0.0061, 9: 0.0695, 10: 0.1149, 11: 0.1420, 12: 0.1584, 13: 0.1541, 14: 0.1415, 15: 0.1171, 16: 0.0793, 17: 0.0172},
            2: {8: 0.0233, 9: 0.0712, 10: 0.1091, 11: 0.1348, 12: 0.1454, 13: 0.1473, 14: 0.1354, 15: 0.1130, 16: 0.0816, 17: 0.0387, 18: 0.0001},
            3: {7: 0.0034, 8: 0.0378, 9: 0.0771, 10: 0.1081, 11: 0.1287, 12: 0.1340, 13: 0.1408, 14: 0.1307, 15: 0.1085, 16: 0.0813, 17: 0.0436, 18: 0.0061},
            4: {6: 0.0005, 7: 0.0157, 8: 0.0525, 9: 0.0858, 10: 0.1078, 11: 0.1229, 12: 0.1296, 13: 0.1303, 14: 0.1202, 15: 0.1030, 16: 0.0774, 17: 0.0438, 18: 0.0106},
            5: {6: 0.0042, 7: 0.0211, 8: 0.0553, 9: 0.0846, 10: 0.1042, 11: 0.1201, 12: 0.1276, 13: 0.1260, 14: 0.1181, 15: 0.1023, 16: 0.0774, 17: 0.0447, 18: 0.0128, 19: 0.0017},
            6: {6: 0.0056, 7: 0.0208, 8: 0.0525, 9: 0.0813, 10: 0.1013, 11: 0.1169, 12: 0.1260, 13: 0.1267, 14: 0.1189, 15: 0.1034, 16: 0.0791, 17: 0.0475, 18: 0.0158, 19: 0.0042},
            7: {6: 0.0037, 7: 0.0163, 8: 0.0477, 9: 0.0764, 10: 0.0990, 11: 0.1175, 12: 0.1282, 13: 0.1292, 14: 0.1224, 15: 0.1072, 16: 0.0823, 17: 0.0498, 18: 0.0164, 19: 0.0038},
            8: {6: 0.0008, 7: 0.0141, 8: 0.0472, 9: 0.0792, 10: 0.1039, 11: 0.1224, 12: 0.1311, 13: 0.1311, 14: 0.1221, 15: 0.1065, 16: 0.0807, 17: 0.0469, 18: 0.0129, 19: 0.0010},
            9: {7: 0.0126, 8: 0.0495, 9: 0.0833, 10: 0.1118, 11: 0.1280, 12: 0.1371, 13: 0.1364, 14: 0.1209, 15: 0.1041, 16: 0.0747, 17: 0.0373, 18: 0.0043},
            10: {7: 0.0057, 8: 0.0495, 9: 0.0858, 10: 0.1205, 11: 0.1391, 12: 0.1414, 13: 0.1433, 14: 0.1253, 15: 0.1035, 16: 0.0649, 17: 0.0210},
            11: {8: 0.0372, 9: 0.0876, 10: 0.1255, 11: 0.1532, 12: 0.1600, 13: 0.1478, 14: 0.1269, 15: 0.1011, 16: 0.0590, 17: 0.0018},
            12: {8: 0.0114, 9: 0.0790, 10: 0.1204, 11: 0.1504, 12: 0.1613, 13: 0.1603, 14: 0.1408, 15: 0.1116, 16: 0.0648}
        },
        'Setúbal': {
            1: {8: 0.0047, 9: 0.0708, 10: 0.1141, 11: 0.1425, 12: 0.1550, 13: 0.1532, 14: 0.1408, 15: 0.1197, 16: 0.0796, 17: 0.0196},
            2: {8: 0.0231, 9: 0.0712, 10: 0.1080, 11: 0.1336, 12: 0.1444, 13: 0.1484, 14: 0.1349, 15: 0.1149, 16: 0.0827, 17: 0.0388, 18: 0.0001},
            3: {7: 0.0030, 8: 0.0365, 9: 0.0746, 10: 0.1058, 11: 0.1283, 12: 0.1381, 13: 0.1420, 14: 0.1311, 15: 0.1103, 16: 0.0808, 17: 0.0432, 18: 0.0064},
            4: {6: 0.0004, 7: 0.0142, 8: 0.0490, 9: 0.0827, 10: 0.1043, 11: 0.1219, 12: 0.1307, 13: 0.1326, 14: 0.1214, 15: 0.1062, 16: 0.0810, 17: 0.0452, 18: 0.0105},
            5: {6: 0.0036, 7: 0.0197, 8: 0.0530, 9: 0.0817, 10: 0.1011, 11: 0.1203, 12: 0.1302, 13: 0.1292, 14: 0.1199, 15: 0.1045, 16: 0.0780, 17: 0.0445, 18: 0.0127, 19: 0.0016},
            6: {6: 0.0050, 7: 0.0190, 8: 0.0492, 9: 0.0786, 10: 0.1003, 11: 0.1176, 12: 0.1280, 13: 0.1272, 14: 0.1212, 15: 0.1060, 16: 0.0806, 17: 0.0479, 18: 0.0155, 19: 0.0040},
            7: {6: 0.0032, 7: 0.0151, 8: 0.0451, 9: 0.0747, 10: 0.0985, 11: 0.1183, 12: 0.1291, 13: 0.1311, 14: 0.1244, 15: 0.1083, 16: 0.0825, 17: 0.0500, 18: 0.0161, 19: 0.0036},
            8: {6: 0.0007, 7: 0.0130, 8: 0.0450, 9: 0.0766, 10: 0.1021, 11: 0.1216, 12: 0.1321, 13: 0.1335, 14: 0.1249, 15: 0.1078, 16: 0.0816, 17: 0.0473, 18: 0.0130, 19: 0.0009},
            9: {7: 0.0116, 8: 0.0465, 9: 0.0815, 10: 0.1082, 11: 0.1257, 12: 0.1365, 13: 0.1391, 14: 0.1257, 15: 0.1074, 16: 0.0761, 17: 0.0372, 18: 0.0044},
            10: {7: 0.0051, 8: 0.0477, 9: 0.0864, 10: 0.1188, 11: 0.1379, 12: 0.1445, 13: 0.1409, 14: 0.1282, 15: 0.1030, 16: 0.0655, 17: 0.0217},
            11: {8: 0.0387, 9: 0.0859, 10: 0.1258, 11: 0.1507, 12: 0.1552, 13: 0.1502, 14: 0.1301, 15: 0.1018, 16: 0.0590, 17: 0.0026},
            12: {8: 0.0103, 9: 0.0762, 10: 0.1206, 11: 0.1526, 12: 0.1617, 13: 0.1601, 14: 0.1398, 15: 0.1137, 16: 0.0649, 17: 0.0001}
        },
        'Viana do Castelo': {
            1: {8: 0.0004, 9: 0.0672, 10: 0.1149, 11: 0.1444, 12: 0.1603, 13: 0.1568, 14: 0.1458, 15: 0.1193, 16: 0.0797, 17: 0.0110},
            2: {8: 0.0199, 9: 0.0697, 10: 0.1085, 11: 0.1356, 12: 0.1433, 13: 0.1511, 14: 0.1370, 15: 0.1164, 16: 0.0811, 17: 0.0374},
            3: {7: 0.0032, 8: 0.0369, 9: 0.0768, 10: 0.1057, 11: 0.1282, 12: 0.1360, 13: 0.1391, 14: 0.1311, 15: 0.1121, 16: 0.0817, 17: 0.0434, 18: 0.0059},
            4: {6: 0.0006, 7: 0.0151, 8: 0.0485, 9: 0.0789, 10: 0.1021, 11: 0.1234, 12: 0.1334, 13: 0.1346, 14: 0.1231, 15: 0.1068, 16: 0.0789, 17: 0.0437, 18: 0.0109, 19: 0.0001},
            5: {6: 0.0043, 7: 0.0209, 8: 0.0519, 9: 0.0796, 10: 0.1003, 11: 0.1185, 12: 0.1300, 13: 0.1300, 14: 0.1221, 15: 0.1045, 16: 0.0771, 17: 0.0448, 18: 0.0138, 19: 0.0022},
            6: {6: 0.0060, 7: 0.0210, 8: 0.0498, 9: 0.0771, 10: 0.0971, 11: 0.1162, 12: 0.1272, 13: 0.1299, 14: 0.1215, 15: 0.1051, 16: 0.0798, 17: 0.0476, 18: 0.0170, 19: 0.0045},
            7: {6: 0.0042, 7: 0.0171, 8: 0.0454, 9: 0.0727, 10: 0.0947, 11: 0.1154, 12: 0.1283, 13: 0.1301, 14: 0.1249, 15: 0.1092, 16: 0.0841, 17: 0.0513, 18: 0.0183, 19: 0.0043},
            8: {6: 0.0011, 7: 0.0145, 8: 0.0446, 9: 0.0739, 10: 0.0981, 11: 0.1193, 12: 0.1325, 13: 0.1343, 14: 0.1262, 15: 0.1090, 16: 0.0825, 17: 0.0484, 18: 0.0142, 19: 0.0013},
            9: {7: 0.0122, 8: 0.0473, 9: 0.0800, 10: 0.1076, 11: 0.1260, 12: 0.1371, 13: 0.1379, 14: 0.1260, 15: 0.1082, 16: 0.0759, 17: 0.0372, 18: 0.0046},
            10: {7: 0.0047, 8: 0.0472, 9: 0.0842, 10: 0.1195, 11: 0.1354, 12: 0.1484, 13: 0.1435, 14: 0.1265, 15: 0.1047, 16: 0.0665, 17: 0.0193},
            11: {8: 0.0317, 9: 0.0841, 10: 0.1269, 11: 0.1557, 12: 0.1576, 13: 0.1543, 14: 0.1303, 15: 0.1030, 16: 0.0560, 17: 0.0003},
            12: {8: 0.0041, 9: 0.0748, 10: 0.1239, 11: 0.1558, 12: 0.1670, 13: 0.1623, 14: 0.1399, 15: 0.1112, 16: 0.0609}
        },
        'Vila Real': {
            1: {8: 0.0004, 9: 0.0732, 10: 0.1189, 11: 0.1453, 12: 0.1616, 13: 0.1572, 14: 0.1407, 15: 0.1193, 16: 0.0782, 17: 0.0051},
            2: {8: 0.0194, 9: 0.0752, 10: 0.1128, 11: 0.1394, 12: 0.1420, 13: 0.1488, 14: 0.1358, 15: 0.1125, 16: 0.0797, 17: 0.0345},
            3: {7: 0.0038, 8: 0.0419, 9: 0.0816, 10: 0.1119, 11: 0.1319, 12: 0.1345, 13: 0.1389, 14: 0.1247, 15: 0.1083, 16: 0.0780, 17: 0.0406, 18: 0.0039},
            4: {6: 0.0010, 7: 0.0191, 8: 0.0554, 9: 0.0874, 10: 0.1100, 11: 0.1282, 12: 0.1314, 13: 0.1307, 14: 0.1141, 15: 0.0988, 16: 0.0731, 17: 0.0411, 18: 0.0096},
            5: {6: 0.0051, 7: 0.0251, 8: 0.0600, 9: 0.0895, 10: 0.1081, 11: 0.1222, 12: 0.1268, 13: 0.1231, 14: 0.1138, 15: 0.0967, 16: 0.0734, 17: 0.0424, 18: 0.0123, 19: 0.0016},
            6: {6: 0.0066, 7: 0.0245, 8: 0.0565, 9: 0.0849, 10: 0.1047, 11: 0.1198, 12: 0.1267, 13: 0.1238, 14: 0.1143, 15: 0.0996, 16: 0.0746, 17: 0.0447, 18: 0.0153, 19: 0.0040},
            7: {6: 0.0045, 7: 0.0196, 8: 0.0530, 9: 0.0820, 10: 0.1049, 11: 0.1195, 12: 0.1269, 13: 0.1259, 14: 0.1174, 15: 0.1022, 16: 0.0778, 17: 0.0472, 18: 0.0156, 19: 0.0036},
            8: {6: 0.0015, 7: 0.0171, 8: 0.0521, 9: 0.0841, 10: 0.1075, 11: 0.1241, 12: 0.1304, 13: 0.1295, 14: 0.1181, 15: 0.1017, 16: 0.0769, 17: 0.0443, 18: 0.0119, 19: 0.0009},
            9: {7: 0.0157, 8: 0.0535, 9: 0.0875, 10: 0.1154, 11: 0.1314, 12: 0.1383, 13: 0.1339, 14: 0.1187, 15: 0.0992, 16: 0.0689, 17: 0.0341, 18: 0.0032},
            10: {7: 0.0040, 8: 0.0547, 9: 0.0919, 10: 0.1242, 11: 0.1408, 12: 0.1464, 13: 0.1394, 14: 0.1220, 15: 0.0978, 16: 0.0620, 17: 0.0168},
            11: {8: 0.0342, 9: 0.0934, 10: 0.1291, 11: 0.1533, 12: 0.1577, 13: 0.1488, 14: 0.1265, 15: 0.1010, 16: 0.0560, 17: 0.0001},
            12: {8: 0.0009, 9: 0.0785, 10: 0.1228, 11: 0.1526, 12: 0.1643, 13: 0.1610, 14: 0.1410, 15: 0.1158, 16: 0.0632}
        },
        'Viseu': {
            1: {8: 0.0072, 9: 0.0737, 10: 0.1186, 11: 0.1445, 12: 0.1604, 13: 0.1554, 14: 0.1405, 15: 0.1136, 16: 0.0760, 17: 0.0102},
            2: {8: 0.0254, 9: 0.0746, 10: 0.1117, 11: 0.1375, 12: 0.1458, 13: 0.1471, 14: 0.1333, 15: 0.1111, 16: 0.0784, 17: 0.0350},
            3: {7: 0.0045, 8: 0.0408, 9: 0.0804, 10: 0.1101, 11: 0.1302, 12: 0.1369, 13: 0.1377, 14: 0.1247, 15: 0.1100, 16: 0.0791, 17: 0.0410, 18: 0.0047},
            4: {6: 0.0008, 7: 0.0180, 8: 0.0537, 9: 0.0851, 10: 0.1110, 11: 0.1274, 12: 0.1329, 13: 0.1289, 14: 0.1162, 15: 0.1005, 16: 0.0748, 17: 0.0411, 18: 0.0096},
            5: {6: 0.0050, 7: 0.0237, 8: 0.0575, 9: 0.0876, 10: 0.1078, 11: 0.1237, 12: 0.1289, 13: 0.1248, 14: 0.1129, 15: 0.0976, 16: 0.0736, 17: 0.0429, 18: 0.0124, 19: 0.0017},
            6: {6: 0.0064, 7: 0.0230, 8: 0.0535, 9: 0.0821, 10: 0.1027, 11: 0.1184, 12: 0.1271, 13: 0.1253, 14: 0.1167, 15: 0.1016, 16: 0.0770, 17: 0.0464, 18: 0.0157, 19: 0.0041},
            7: {6: 0.0044, 7: 0.0185, 8: 0.0491, 9: 0.0792, 10: 0.1021, 11: 0.1199, 12: 0.1293, 13: 0.1284, 14: 0.1194, 15: 0.1033, 16: 0.0789, 17: 0.0480, 18: 0.0157, 19: 0.0037},
            8: {6: 0.0013, 7: 0.0160, 8: 0.0486, 9: 0.0813, 10: 0.1064, 11: 0.1236, 12: 0.1317, 13: 0.1311, 14: 0.1211, 15: 0.1036, 16: 0.0774, 17: 0.0451, 18: 0.0121, 19: 0.0009},
            9: {7: 0.0148, 8: 0.0528, 9: 0.0871, 10: 0.1138, 11: 0.1310, 12: 0.1378, 13: 0.1343, 14: 0.1199, 15: 0.0993, 16: 0.0712, 17: 0.0346, 18: 0.0035},
            10: {7: 0.0073, 8: 0.0539, 9: 0.0918, 10: 0.1226, 11: 0.1399, 12: 0.1443, 13: 0.1391, 14: 0.1225, 15: 0.0978, 16: 0.0628, 17: 0.0181},
            11: {8: 0.0403, 9: 0.0908, 10: 0.1293, 11: 0.1517, 12: 0.1565, 13: 0.1477, 14: 0.1285, 15: 0.1001, 16: 0.0548, 17: 0.0001},
            12: {8: 0.0111, 9: 0.0819, 10: 0.1238, 11: 0.1541, 12: 0.1645, 13: 0.1603, 14: 0.1361, 15: 0.1086, 16: 0.0595}
        },
        'Açores (Ponta Delgada)': {
            1: {8: 0.0015, 9: 0.0515, 10: 0.0973, 11: 0.1393, 12: 0.1565, 13: 0.1617, 14: 0.1575, 15: 0.1278, 16: 0.0812, 17: 0.0259},
            2: {8: 0.0132, 9: 0.0549, 10: 0.0953, 11: 0.1215, 12: 0.1533, 13: 0.1565, 14: 0.1496, 15: 0.1270, 16: 0.0867, 17: 0.0411, 18: 0.0011},
            3: {7: 0.0016, 8: 0.0275, 9: 0.0645, 10: 0.0994, 11: 0.1229, 12: 0.1431, 13: 0.1443, 14: 0.1420, 15: 0.1178, 16: 0.0844, 17: 0.0441, 18: 0.0086},
            4: {6: 0.0001, 7: 0.0112, 8: 0.0402, 9: 0.0716, 10: 0.1025, 11: 0.1225, 12: 0.1351, 13: 0.1355, 14: 0.1301, 15: 0.1112, 16: 0.0815, 17: 0.0462, 18: 0.0122, 19: 0.0001},
            5: {6: 0.0027, 7: 0.0177, 8: 0.0451, 9: 0.0745, 10: 0.1026, 11: 0.1212, 12: 0.1289, 13: 0.1324, 14: 0.1282, 15: 0.1075, 16: 0.0782, 17: 0.0448, 18: 0.0145, 19: 0.0017},
            6: {6: 0.0043, 7: 0.0179, 8: 0.0454, 9: 0.0722, 10: 0.0977, 11: 0.1182, 12: 0.1290, 13: 0.1308, 14: 0.1272, 15: 0.1072, 16: 0.0810, 17: 0.0476, 18: 0.0177, 19: 0.0038},
            7: {6: 0.0025, 7: 0.0150, 8: 0.0426, 9: 0.0688, 10: 0.0964, 11: 0.1163, 12: 0.1285, 13: 0.1319, 14: 0.1282, 15: 0.1111, 16: 0.0840, 17: 0.0511, 18: 0.0197, 19: 0.0039},
            8: {6: 0.0002, 7: 0.0110, 8: 0.0414, 9: 0.0742, 10: 0.1023, 11: 0.1208, 12: 0.1311, 13: 0.1304, 14: 0.1281, 15: 0.1115, 16: 0.0837, 17: 0.0483, 18: 0.0157, 19: 0.0013},
            9: {7: 0.0087, 8: 0.0405, 9: 0.0774, 10: 0.1087, 11: 0.1277, 12: 0.1378, 13: 0.1314, 14: 0.1324, 15: 0.1105, 16: 0.0789, 17: 0.0394, 18: 0.0066},
            10: {7: 0.0023, 8: 0.0367, 9: 0.0797, 10: 0.1120, 11: 0.1367, 12: 0.1527, 13: 0.1465, 14: 0.1358, 15: 0.1064, 16: 0.0675, 17: 0.0237, 18: 0.0001},
            11: {8: 0.0244, 9: 0.0740, 10: 0.1160, 11: 0.1463, 12: 0.1602, 13: 0.1477, 14: 0.1476, 15: 0.1105, 16: 0.0647, 17: 0.0084},
            12: {8: 0.0047, 9: 0.0647, 10: 0.1079, 11: 0.1471, 12: 0.1691, 13: 0.1579, 14: 0.1563, 15: 0.1197, 16: 0.0691, 17: 0.0034}
        },
        'Madeira (Funchal)': {
            1: {9: 0.0529, 10: 0.1059, 11: 0.1348, 12: 0.1474, 13: 0.1440, 14: 0.1343, 15: 0.1201, 16: 0.0989, 17: 0.0574, 18: 0.0045},
            2: {8: 0.0047, 9: 0.0553, 10: 0.1017, 11: 0.1265, 12: 0.1289, 13: 0.1419, 14: 0.1332, 15: 0.1226, 16: 0.0994, 17: 0.0642, 18: 0.0216},
            3: {8: 0.0198, 9: 0.0638, 10: 0.0997, 11: 0.1212, 12: 0.1263, 13: 0.1356, 14: 0.1296, 15: 0.1179, 16: 0.0945, 17: 0.0645, 18: 0.0270, 19: 0.0001},
            4: {7: 0.0027, 8: 0.0334, 9: 0.0767, 10: 0.1047, 11: 0.1181, 12: 0.1252, 13: 0.1268, 14: 0.1195, 15: 0.1106, 16: 0.0904, 17: 0.0624, 18: 0.0270, 19: 0.0025},
            5: {7: 0.0067, 8: 0.0387, 9: 0.0781, 10: 0.0986, 11: 0.1152, 12: 0.1232, 13: 0.1243, 14: 0.1207, 15: 0.1110, 16: 0.0884, 17: 0.0606, 18: 0.0282, 19: 0.0061},
            6: {7: 0.0092, 8: 0.0381, 9: 0.0789, 10: 0.0981, 11: 0.1131, 12: 0.1233, 13: 0.1201, 14: 0.1162, 15: 0.1091, 16: 0.0901, 17: 0.0628, 18: 0.0315, 19: 0.0094, 20: 0.0001},
            7: {7: 0.0068, 8: 0.0313, 9: 0.0721, 10: 0.0981, 11: 0.1158, 12: 0.1253, 13: 0.1253, 14: 0.1207, 15: 0.1109, 16: 0.0910, 17: 0.0627, 18: 0.0316, 19: 0.0083, 20: 0.0001},
            8: {7: 0.0036, 8: 0.0305, 9: 0.0722, 10: 0.0998, 11: 0.1177, 12: 0.1263, 13: 0.1281, 14: 0.1222, 15: 0.1125, 16: 0.0908, 17: 0.0623, 18: 0.0293, 19: 0.0047},
            9: {7: 0.0009, 8: 0.0368, 9: 0.0820, 10: 0.1120, 11: 0.1245, 12: 0.1278, 13: 0.1263, 14: 0.1177, 15: 0.1075, 16: 0.0861, 17: 0.0568, 18: 0.0213, 19: 0.0003},
            10: {8: 0.0344, 9: 0.0830, 10: 0.1181, 11: 0.1313, 12: 0.1352, 13: 0.1345, 14: 0.1200, 15: 0.1079, 16: 0.0831, 17: 0.0463, 18: 0.0062},
            11: {8: 0.0189, 9: 0.0748, 10: 0.1194, 11: 0.1418, 12: 0.1455, 13: 0.1415, 14: 0.1245, 15: 0.1103, 16: 0.0831, 17: 0.0401},
            12: {8: 0.0001, 9: 0.0616, 10: 0.1121, 11: 0.1408, 12: 0.1477, 13: 0.1508, 14: 0.1310, 15: 0.1210, 16: 0.0899, 17: 0.0450}
        },
    }

    if df_consumos is None or df_consumos.empty:
        return df_consumos.copy()

    # Garante que os dicionários de dados estão disponíveis no escopo desta função
    # Se não estiverem globais, pode passá-los como parâmetros
    dados_producao_distrito = DADOS_PVGIS_DISTRITO.get(distrito)
    perfis_horarios_distrito = PERFIS_HORARIOS_MENSAIS_POR_DISTRITO.get(distrito)

    if not dados_producao_distrito or not perfis_horarios_distrito:
        # st.error(f"Não foram encontrados dados de backup para o distrito '{distrito}'.")
        # Retorna um DF vazio para não quebrar a aplicação
        return pd.DataFrame(columns=['DataHora', 'Consumo (kWh)', 'Producao_Solar_kWh', 'Autoconsumo_kWh', 'Excedente_kWh', 'Consumo_Rede_kWh'])

    df_resultado = df_consumos.copy()
    
    perfis_quarto_horarios = interpolar_perfis_para_quarto_horario({distrito: perfis_horarios_distrito})[distrito]

    # --- FATORES DE AJUSTE (Lógica da versão de referência) ---
    fator_inclinacao = 1.0 - (abs(inclinacao - 35) / 100) * 0.5
    
    # Conversão da string de orientação para uma lógica numérica consistente
    fator_orientacao = 1.0  # Sul (Padrão)
    if "Sudeste / Sudoeste" in orientacao_str:
        fator_orientacao = 0.95
    elif "Este / Oeste" in orientacao_str:
        fator_orientacao = 0.80
    
    # Fator de perdas alinhado com a versão de referência (assumindo 14% como padrão)
    system_loss = 14.0 
    fator_perdas_sistema = system_loss / 100.0

    def calcular_producao_por_linha(row):
        timestamp_inicio = row['DataHora'] - pd.Timedelta(minutes=15)
        mes, hora, minuto = timestamp_inicio.month, timestamp_inicio.hour, timestamp_inicio.minute

        energia_diaria_base = dados_producao_distrito.get(mes, 0)
        
        # Fórmula de cálculo alinhada com a versão de referência
        energia_diaria_total_sistema = (
            energia_diaria_base * potencia_kwp *
            fator_inclinacao * fator_orientacao *
            (1 - fator_perdas_sistema)
        )
        
        perfil_mensal = perfis_quarto_horarios.get(mes, {})
        fator_distribuicao = perfil_mensal.get((hora, minuto), 0)
        
        producao_kwh_intervalo = energia_diaria_total_sistema * fator_distribuicao
        return producao_kwh_intervalo

    df_resultado['Producao_Solar_kWh'] = df_resultado.apply(calcular_producao_por_linha, axis=1)

    soma_original_precisa = df_resultado['Producao_Solar_kWh'].sum()
    df_resultado['Producao_Solar_kWh'] = df_resultado['Producao_Solar_kWh'].rolling(window=4, center=False, min_periods=1).mean()
    soma_apos_suavizar = df_resultado['Producao_Solar_kWh'].sum()
    if soma_apos_suavizar > 0:
        fator_correcao = soma_original_precisa / soma_apos_suavizar
        df_resultado['Producao_Solar_kWh'] *= fator_correcao

    # O cálculo final agora será sobre a produção suavizada e corrigida
    df_resultado['Autoconsumo_kWh'] = np.minimum(df_resultado['Consumo (kWh)'], df_resultado['Producao_Solar_kWh'])
    df_resultado['Excedente_kWh'] = np.maximum(0, df_resultado['Producao_Solar_kWh'] - df_resultado['Consumo (kWh)'])
    df_resultado['Consumo_Rede_kWh'] = np.maximum(0, df_resultado['Consumo (kWh)'] - df_resultado['Autoconsumo_kWh'])

    return df_resultado

def calcular_detalhes_custo_meu_tarifario(
    st_session_state,
    opcao_horaria,
    consumos_para_calculo,
    potencia,
    dias,
    tarifa_social,
    familia_numerosa,
    valor_dgeg_user,
    valor_cav_user,
    CONSTANTES,
    FINANCIAMENTO_TSE_VAL
):
    """
    Calcula o custo completo para a funcionalidade "O Meu Tarifário",
    usando os inputs guardados no st.session_state.
    """
    try:
        # --- PASSO 1: EXTRAIR INPUTS DO st.session_state ---
        # Preços de Energia e Potência
        energia_meu_s = st_session_state.get("energia_meu_s_input_val", 0.0)
        potencia_meu = st_session_state.get("potencia_meu_input_val", 0.0)
        energia_meu_v = st_session_state.get("energia_meu_v_input_val", 0.0)
        energia_meu_f = st_session_state.get("energia_meu_f_input_val", 0.0)
        energia_meu_c = st_session_state.get("energia_meu_c_input_val", 0.0)
        energia_meu_p = st_session_state.get("energia_meu_p_input_val", 0.0)
        # Checkboxes TAR/TSE
        tar_incluida_energia_meu = st_session_state.get("meu_tar_energia_val", True)
        tar_incluida_potencia_meu = st_session_state.get("meu_tar_potencia_val", True)
        checkbox_tse_incluido_estado = st_session_state.get("meu_fin_tse_incluido_val", True)
        adicionar_financiamento_tse_meu = not checkbox_tse_incluido_estado
        # Descontos e Acréscimos
        desconto_energia = st_session_state.get("meu_desconto_energia_val", 0.0)
        desconto_potencia = st_session_state.get("meu_desconto_potencia_val", 0.0)
        desconto_fatura_input_meu = st_session_state.get("meu_desconto_fatura_val", 0.0)
        acrescimo_fatura_input_meu = st_session_state.get("meu_acrescimo_fatura_val", 0.0)

        # --- PASSO 2: PREPARAR DICIONÁRIOS DE PREÇOS E CONSUMOS ---
        is_billing_month = 28 <= dias <= 31
        preco_energia_input_meu = {}
        
        oh_lower = opcao_horaria.lower()
        if oh_lower == "simples":
            preco_energia_input_meu['S'] = float(energia_meu_s or 0.0)
            preco_potencia_input_meu = float(potencia_meu or 0.0)
        elif oh_lower.startswith("bi"):
            preco_energia_input_meu['V'] = float(energia_meu_v or 0.0)
            preco_energia_input_meu['F'] = float(energia_meu_f or 0.0)
            preco_potencia_input_meu = float(potencia_meu or 0.0)
        elif oh_lower.startswith("tri"):
            preco_energia_input_meu['V'] = float(energia_meu_v or 0.0)
            preco_energia_input_meu['C'] = float(energia_meu_c or 0.0)
            preco_energia_input_meu['P'] = float(energia_meu_p or 0.0)
            preco_potencia_input_meu = float(potencia_meu or 0.0)
        
        # --- PASSO 3: CÁLCULO DETALHADO (Lógica que já tinha) ---
        # (Esta parte é uma adaptação direta do seu código original)
        
        tar_energia_regulada_periodo_meu = {p: obter_tar_energia_periodo(opcao_horaria, p, potencia, CONSTANTES) for p in preco_energia_input_meu.keys()}
        tar_potencia_regulada_meu_base = obter_tar_dia(potencia, CONSTANTES)

        energia_meu_periodo_comercializador_base = {}
        for p_key, preco_input_val in preco_energia_input_meu.items():
            preco_input_val_float = float(preco_input_val or 0.0)
            energia_meu_periodo_comercializador_base[p_key] = preco_input_val_float - tar_energia_regulada_periodo_meu.get(p_key, 0.0) if tar_incluida_energia_meu else preco_input_val_float

        potencia_meu_comercializador_base = (float(preco_potencia_input_meu or 0.0) - tar_potencia_regulada_meu_base) if tar_incluida_potencia_meu else float(preco_potencia_input_meu or 0.0)
        
        financiamento_tse_a_somar_base = FINANCIAMENTO_TSE_VAL if adicionar_financiamento_tse_meu else 0.0

        desconto_monetario_ts_energia = obter_constante('Desconto TS Energia', CONSTANTES) if tarifa_social else 0.0
        preco_energia_final_unitario_sem_iva = {}
        for p_key in energia_meu_periodo_comercializador_base.keys():
            base_desc_perc = energia_meu_periodo_comercializador_base.get(p_key, 0.0) + tar_energia_regulada_periodo_meu.get(p_key, 0.0) + financiamento_tse_a_somar_base
            apos_desc_comerc = base_desc_perc * (1 - (desconto_energia or 0.0) / 100.0)
            preco_energia_final_unitario_sem_iva[p_key] = apos_desc_comerc - desconto_monetario_ts_energia if tarifa_social else apos_desc_comerc
            
        desconto_monetario_ts_potencia = obter_constante(f'Desconto TS Potencia {potencia}', CONSTANTES) if tarifa_social else 0.0
        base_desc_pot_perc = potencia_meu_comercializador_base + tar_potencia_regulada_meu_base
        apos_desc_pot_comerc = base_desc_pot_perc * (1 - (desconto_potencia or 0.0) / 100.0)
        preco_potencia_final_unitario_sem_iva = apos_desc_pot_comerc - desconto_monetario_ts_potencia if tarifa_social else apos_desc_pot_comerc

        consumo_total = sum(consumos_para_calculo.values())
        decomposicao_energia = calcular_custo_energia_com_iva(consumo_total, preco_energia_final_unitario_sem_iva.get('S'), {k:v for k,v in preco_energia_final_unitario_sem_iva.items() if k!='S'}, dias, potencia, opcao_horaria, consumos_para_calculo, familia_numerosa)
        
        comerc_pot_para_iva = potencia_meu_comercializador_base * (1 - (desconto_potencia or 0.0) / 100.0)
        tar_pot_bruta_apos_desc = tar_potencia_regulada_meu_base * (1 - (desconto_potencia or 0.0) / 100.0)
        tar_pot_final_para_iva = tar_pot_bruta_apos_desc - desconto_monetario_ts_potencia if tarifa_social else tar_pot_bruta_apos_desc
        decomposicao_potencia = calcular_custo_potencia_com_iva_final(comerc_pot_para_iva, tar_pot_final_para_iva, dias, potencia)
        
        decomposicao_taxas = calcular_taxas_adicionais(consumo_total, dias, tarifa_social, valor_dgeg_user, valor_cav_user, "Pessoal", is_billing_month)
        
        custo_total_antes_desc_fatura = decomposicao_energia['custo_com_iva'] + decomposicao_potencia['custo_com_iva'] + decomposicao_taxas['custo_com_iva']
        custo_final = custo_total_antes_desc_fatura - float(desconto_fatura_input_meu or 0.0) + float(acrescimo_fatura_input_meu or 0.0)

        # --- PASSO 4: MONTAR O DICIONÁRIO DE RETORNO ---
        nome_para_exibir = "O Meu Tarifário"
        sufixo = ""
        desconto = float(desconto_fatura_input_meu or 0.0)
        acrescimo = float(acrescimo_fatura_input_meu or 0.0)
        if desconto > 0 or acrescimo > 0:
            liquido = desconto - acrescimo
            if liquido > 0: sufixo = f" (Inclui desc. líquido de {liquido:.2f}€)"
            elif liquido < 0: sufixo = f" (Inclui acréscimo líquido de {abs(liquido):.2f}€)"
        nome_para_exibir += sufixo

        return { 'Total (€)': custo_final, 'NomeParaExibir': nome_para_exibir }

    except Exception as e:
        st.error(f"Erro ao calcular 'O Meu Tarifário': {e}")
        return None



# PARA GÁS NATURAL
# --- Função: Obter TAR Gás Fixo por Escalão ---
def obter_tar_gas_fixo(escalao_num, constantes_df):
    """Obtém a TAR fixa diária para um determinado escalão de gás."""
    nome_constante = f'TAR_Gas_Fixo_E{escalao_num}'
    return obter_constante(nome_constante, constantes_df)

# --- Função: Obter TAR Gás Energia por Escalão ---
def obter_tar_gas_energia(escalao_num, constantes_df):
    """Obtém a TAR de energia (€/kWh) para um determinado escalão de gás."""
    nome_constante = f'TAR_Gas_Energia_E{escalao_num}'
    return obter_constante(nome_constante, constantes_df)

# --- Função: Obter Desconto TS Gás Fixo por Escalão ---
def obter_desconto_ts_gas_fixo(escalao_num, constantes_df):
    """Obtém o desconto da Tarifa Social sobre o termo fixo de gás."""
    nome_constante = f'Desconto_TS_Gas_Fixo_E{escalao_num}'
    return obter_constante(nome_constante, constantes_df)

# --- Função: Obter Desconto TS Gás Energia por Escalão ---
def obter_desconto_ts_gas_energia(escalao_num, constantes_df):
    """Obtém o desconto da Tarifa Social sobre o termo de energia de gás."""
    nome_constante = f'Desconto_TS_Gas_Energia_E{escalao_num}'
    return obter_constante(nome_constante, constantes_df)

# --- Função Principal: Calcular Custo Total do Gás ---
def calcular_custo_gas_completo(
    dados_tarifa_gas_linha, 
    consumo_kwh_periodo,    
    dias_periodo,           
    escalao_num,
    tarifa_social_ativa,
    constantes_df,
    tos_fixo_dia_val,        
    tos_variavel_kwh_val,   
    mibgas_price_mwh_input,
    isp_gas_valor_manual,
    # --- Argumentos da V14 ---
    acp_gas_flag,
    desconto_continente_gas_flag,
    VALOR_QUOTA_ACP_MENSAL_CONST
):
    """
    (V15) Adiciona as fórmulas de cálculo detalhadas para tarifários indexados de Gás,
    replicando a arquitetura do simulador de eletricidade (Opção 1).
    """
    try:
        IVA_NORMAL_PERC = 0.23
        IVA_REDUZIDO_PERC = 0.06
        
        tipo_tarifa = dados_tarifa_gas_linha.get('tipo', 'Fixo') 
        nome_original_tarifario = dados_tarifa_gas_linha['Nome_Tarifa_G']
        nome_a_exibir_final = nome_original_tarifario 

        # --- 1. Calcular Preço Base de Energia (Comercializador) ---
        
        # Obter o MIBGAS em €/kWh
        mibgas_kwh = mibgas_price_mwh_input / 1000.0
        # Obter Fator de Perdas (constante global de gás, ex: 0.04)
        perdas_dec = obter_constante("PERDAS_GAS_GBL", constantes_df)
        perdas_coef = 1.0 + perdas_dec # (ex: 1.04)

        preco_energia_comerc_input = 0.0 # Inicializar

        if tipo_tarifa == 'Fixo':
            preco_energia_comerc_input = float(dados_tarifa_gas_linha.get('Termo_Energia_eur_kwh', 0.0))
        
        elif tipo_tarifa == 'Indexado':
            # --- INÍCIO DA LÓGICA DE FÓRMULAS INDEXADAS (REQ. V19) ---
            
            if nome_original_tarifario == "Luzigás - Plano Gás":
                # Fórmula: (MIBGAS + K + CGS) (TAR False)
                k = obter_constante("Luzigas_Gas_K", constantes_df)
                cgs = obter_constante("Luzigas_Gas_CGS", constantes_df)
                preco_energia_comerc_input = mibgas_kwh + k + cgs
            
            elif nome_original_tarifario == "EDP - Gás Indexado":
                # Fórmula: (MIBGAS * (1+Perdas) * K1 + K2) (TAR False)
                EDP_Perdas = obter_constante("EDP_Gas_(1+Perdas)", constantes_df)
                k1 = obter_constante("EDP_Gas_K1", constantes_df)
                k2 = obter_constante("EDP_Gas_K2", constantes_df)
                preco_energia_comerc_input = (mibgas_kwh * EDP_Perdas * k1) + k2

            elif nome_original_tarifario == "Galp Plano Flexível - Gás":
                 # Fórmula: (MIBGAS + C) * (1+L) (TAR False)
                c = obter_constante("Galp_Gas_C", constantes_df)
                Galp_Perdas = obter_constante("Galp_Gas_(1+L)", constantes_df)
                preco_energia_comerc_input = (mibgas_kwh + c) * Galp_Perdas

            elif nome_original_tarifario == "Endesa Gás Tarifa Indexada":
                # Fórmula: (MIBGAS + A[escalão]) (TAR True)
                # Esta fórmula depende do escalão.
                a_val = obter_constante(f"Endesa_Gas_A{escalao_num}", constantes_df)
                preco_energia_comerc_input = mibgas_kwh + a_val
                # Nota: A flag tar_incluida_energia=True será lida abaixo e tratará isto corretamente.

            elif nome_original_tarifario == "Goldenergy Tarifa Index Gas 100% Online":
                # Fórmula: (Pmibgas * (1 + Perdas) + QTarifa + CG) (TAR False)
                GE_Perdas = obter_constante("GE_Gas_(1+Perdas)", constantes_df)
                qtarifa = obter_constante("GE_Gas_QTarifa", constantes_df)
                cg = obter_constante("GE_Gas_CG", constantes_df)
                preco_energia_comerc_input = mibgas_kwh * GE_Perdas + qtarifa + cg

            else:
                # Fallback para tarifários indexados genéricos (se existirem)
                margem_generica = float(dados_tarifa_gas_linha.get('Margem_Index', 0.0))
                preco_energia_comerc_input = mibgas_kwh + margem_generica
                if margem_generica == 0.0:
                    st.warning(f"Aviso: Tarifário indexado '{nome_original_tarifario}' não tem fórmula dedicada nem Margem_Index no Excel. Custo de energia pode ser zero.")
            
            # --- FIM DA LÓGICA DE FÓRMULAS ---

        # --- 2. Obter Preço Fixo e Flags ---
        preco_fixo_comerc_input = float(dados_tarifa_gas_linha.get('Termo_Fixo_eur_dia', 0.0))
        tar_fixo_incluida_flag = dados_tarifa_gas_linha.get('tar_incluida_termo_fixo', True)
        tar_energia_incluida_flag = dados_tarifa_gas_linha.get(
            'tar_incluida_energia', 
            False if tipo_tarifa == 'Indexado' else True
        )

        # 3. Obter TARs Reguladas (Base)
        tar_fixo_regulada_base = obter_tar_gas_fixo(escalao_num, constantes_df)
        tar_energia_regulada_base = obter_tar_gas_energia(escalao_num, constantes_df)
        
        # 4. Obter ISP (do input manual)
        isp_gas_kwh = isp_gas_valor_manual 

        # 5. Determinar componentes do Comercializador (separar TARs)
        comp_fixo_comercializador_dia = (preco_fixo_comerc_input - tar_fixo_regulada_base) if tar_fixo_incluida_flag else preco_fixo_comerc_input
        # Se a flag for True (ex: Endesa), subtrai a TAR. Se for False (outros indexados), o preço é apenas a componente comercial.
        comp_energia_comercializador_kwh = (preco_energia_comerc_input - tar_energia_regulada_base) if tar_energia_incluida_flag else preco_energia_comerc_input


        # 6. Aplicar Tarifa Social (TS) - Desconto aplica-se às TARs
        tar_fixo_final_a_pagar = tar_fixo_regulada_base
        tar_energia_final_a_pagar = tar_energia_regulada_base
        isp_total_s_iva_periodo = consumo_kwh_periodo * isp_gas_kwh 
        
        desconto_ts_fixo_valor_aplicado = 0.0
        desconto_ts_energia_valor_aplicado = 0.0

        if tarifa_social_ativa and escalao_num in [1, 2]: 
            desconto_ts_fixo_bruto = obter_desconto_ts_gas_fixo(escalao_num, constantes_df)
            desconto_ts_energia_bruto = obter_desconto_ts_gas_energia(escalao_num, constantes_df)
            
            tar_fixo_final_a_pagar = max(0.0, tar_fixo_regulada_base - desconto_ts_fixo_bruto)
            tar_energia_final_a_pagar = max(0.0, tar_energia_regulada_base - desconto_ts_energia_bruto)
            
            desconto_ts_fixo_valor_aplicado = tar_fixo_regulada_base - tar_fixo_final_a_pagar
            desconto_ts_energia_valor_aplicado = tar_energia_regulada_base - tar_energia_final_a_pagar
            
            isp_total_s_iva_periodo = 0.0
            
        # 7. Preços Unitários Finais (Sem IVA) - Para exibir na tabela
        preco_fixo_final_s_iva_dia = comp_fixo_comercializador_dia + tar_fixo_final_a_pagar
        # Preço final da energia = Componente Comercial (da fórmula) + TAR Final (pós-TS)
        preco_energia_final_s_iva_kwh = comp_energia_comercializador_kwh + tar_energia_final_a_pagar

        # --- 8. LÓGICA DE CÁLCULO DE CUSTO (SEPARADA POR IVA) ---
        custo_tar_fixo_periodo_s_iva = tar_fixo_final_a_pagar * dias_periodo
        custo_comerc_fixo_periodo_s_iva = comp_fixo_comercializador_dia * dias_periodo
        custo_tar_energia_periodo_s_iva = tar_energia_final_a_pagar * consumo_kwh_periodo
        custo_comerc_energia_periodo_s_iva = comp_energia_comercializador_kwh * consumo_kwh_periodo
        custo_tos_fixo_periodo_s_iva = tos_fixo_dia_val * dias_periodo
        custo_tos_variavel_periodo_s_iva = tos_variavel_kwh_val * consumo_kwh_periodo

        total_base_iva_reduzido = custo_tar_fixo_periodo_s_iva
        total_base_iva_normal = (
            custo_comerc_fixo_periodo_s_iva +
            custo_tar_energia_periodo_s_iva +
            custo_comerc_energia_periodo_s_iva +
            isp_total_s_iva_periodo + 
            custo_tos_fixo_periodo_s_iva +
            custo_tos_variavel_periodo_s_iva 
        )

        iva_total_reduzido = total_base_iva_reduzido * IVA_REDUZIDO_PERC
        iva_total_normal = total_base_iva_normal * IVA_NORMAL_PERC
        iva_total_periodo = iva_total_reduzido + iva_total_normal

        custo_subtotal_c_iva = total_base_iva_reduzido + total_base_iva_normal + iva_total_periodo

        # --- 9. LÓGICA DE DESCONTOS FINAIS (V14) ---
        
        is_billing_month = 28 <= dias_periodo <= 31
        desconto_total_final_eur = 0.0
        acrescimo_total_final_eur = 0.0
        
        desconto_fatura_mensal_excel = float(dados_tarifa_gas_linha.get('desconto_fatura_mes', 0.0) or 0.0)
        if desconto_fatura_mensal_excel > 0:
            desconto_aplicado = desconto_fatura_mensal_excel if is_billing_month else (desconto_fatura_mensal_excel / 30.0) * dias_periodo
            desconto_total_final_eur += desconto_aplicado
            nome_a_exibir_final += f" (INCLUI desc. {desconto_fatura_mensal_excel:.2f}€/mês)" 
        
        if acp_gas_flag and nome_original_tarifario.startswith("Goldenergy - ACP"):
            quota_aplicada = VALOR_QUOTA_ACP_MENSAL_CONST if is_billing_month else (VALOR_QUOTA_ACP_MENSAL_CONST / 30.0) * dias_periodo
            acrescimo_total_final_eur += quota_aplicada
            nome_a_exibir_final += f" (INCLUI Quota ACP)"
        
        if desconto_continente_gas_flag and nome_original_tarifario.startswith("Galp & Continente"):
            # Base de custo ANTES da Tarifa Social
            custo_energia_bruto_s_iva = (comp_energia_comercializador_kwh + tar_energia_regulada_base) * consumo_kwh_periodo
            custo_fixo_bruto_s_iva = (comp_fixo_comercializador_dia + tar_fixo_regulada_base) * dias_periodo
            
            base_iva_reduzido_bruto = tar_fixo_regulada_base * dias_periodo
            base_iva_normal_bruto_comerc = comp_fixo_comercializador_dia * dias_periodo
            
            # Custo Energia c/IVA (Bruto)
            custo_energia_c_iva_bruto = (custo_energia_bruto_s_iva * (1 + IVA_NORMAL_PERC))
            # Custo Fixo c/IVA (Bruto)
            custo_fixo_c_iva_bruto = (base_iva_reduzido_bruto * (1 + IVA_REDUZIDO_PERC)) + (base_iva_normal_bruto_comerc * (1 + IVA_NORMAL_PERC))

            base_desconto_continente_c_iva = custo_energia_c_iva_bruto + custo_fixo_c_iva_bruto
            
            desconto_continente_aplicado = 0.0
            if nome_original_tarifario.startswith("Galp & Continente (-10% DD)"):
                desconto_continente_aplicado = base_desconto_continente_c_iva * 0.10
            elif nome_original_tarifario.startswith("Galp & Continente (-7% s/DD)"):
                desconto_continente_aplicado = base_desconto_continente_c_iva * 0.07

            desconto_total_final_eur += desconto_continente_aplicado
            # Custo ANTES do desconto continente = Custo Subtotal (com TS) + Acréscimos (ACP) - Outros descontos (Fatura)
            custo_antes_continente = custo_subtotal_c_iva - (desconto_total_final_eur - desconto_continente_aplicado) + acrescimo_total_final_eur
            nome_a_exibir_final += f" (INCLUI desc. Cont. de {desconto_continente_aplicado:.2f}€, s/ desc. Cont.={custo_antes_continente:.2f}€)"
            
        custo_final_total_periodo_c_iva = custo_subtotal_c_iva - desconto_total_final_eur + acrescimo_total_final_eur

        # --- 10. Construir Dicionários de Tooltip ---
        componentes_tooltip_termo_fixo_dict = {
            'tooltip_fixo_comerc_sem_tar': comp_fixo_comercializador_dia,
            'tooltip_fixo_tar_bruta': tar_fixo_regulada_base,
            'tooltip_fixo_ts_aplicada_flag': tarifa_social_ativa and escalao_num in [1, 2],
            'tooltip_fixo_ts_desconto_valor': desconto_ts_fixo_valor_aplicado
        }
        componentes_tooltip_termo_energia_dict = {
            'tooltip_energia_comerc_sem_tar': comp_energia_comercializador_kwh,
            'tooltip_energia_tar_bruta': tar_energia_regulada_base,
            'tooltip_energia_ts_aplicada_flag': tarifa_social_ativa and escalao_num in [1, 2],
            'tooltip_energia_ts_desconto_valor': desconto_ts_energia_valor_aplicado
        }
        componentes_tooltip_custo_total_dict = {
            'tt_cte_energia_siva': custo_tar_energia_periodo_s_iva + custo_comerc_energia_periodo_s_iva,
            'tt_cte_fixo_siva': custo_tar_fixo_periodo_s_iva + custo_comerc_fixo_periodo_s_iva,
            'tt_cte_isp_siva': isp_total_s_iva_periodo,
            'tt_cte_tos_fixo_siva': custo_tos_fixo_periodo_s_iva,
            'tt_cte_tos_var_siva': custo_tos_variavel_periodo_s_iva,
            'tt_cte_total_siva': total_base_iva_reduzido + total_base_iva_normal,
            'tt_cte_valor_iva_6_total': iva_total_reduzido,
            'tt_cte_valor_iva_23_total': iva_total_normal,
            'tt_cte_subtotal_civa': custo_subtotal_c_iva, 
            'tt_cte_desc_finais_valor': desconto_total_final_eur, 
            'tt_cte_acres_finais_valor': acrescimo_total_final_eur 
        }

        return {
            'NomeParaExibir': nome_a_exibir_final, 
            'Comercializador': dados_tarifa_gas_linha['Comercializador'],
            'Termo Fixo (€/dia)': round(preco_fixo_final_s_iva_dia, 5), 
            'Termo Energia (€/kWh)': round(preco_energia_final_s_iva_kwh, 5),
            'Total Período (€)': round(custo_final_total_periodo_c_iva, 2), 
            'tipo': tipo_tarifa,
            'Segmento': dados_tarifa_gas_linha.get('segmento', '-'),
            'Faturação': dados_tarifa_gas_linha.get('faturacao', '-'),
            'Pagamento': dados_tarifa_gas_linha.get('pagamento', '-'),
            **componentes_tooltip_termo_fixo_dict,
            **componentes_tooltip_termo_energia_dict,
            **componentes_tooltip_custo_total_dict
        }
        
    except Exception as e:
        st.error(f"Erro ao calcular custo de gás (V15) para {dados_tarifa_gas_linha.get('Nome_Tarifa_G', 'Desconhecido')}: {e}")
        import traceback
        st.text(traceback.format_exc()) # Para debug detalhado
        return None
    
# --- NOVA FUNÇÃO: Calcular "O Meu Tarifário" de Gás ---
def calcular_custo_meu_tarifario_gas(
    st_session_state, 
    consumo_kwh_periodo,
    dias_periodo,
    escalao_num,
    tarifa_social_ativa,
    constantes_df,
    tos_fixo_dia_val,
    tos_variavel_kwh_val,
    isp_gas_valor_manual
):
    """
    Calcula 'O Meu Tarifário', lógica V11 (TS subtrai ao preço final),
    DEVOLVE DICIONÁRIOS DE TOOLTIP completos e colunas de Segmento (Pessoal).
    """
    try:
        IVA_NORMAL_PERC = 0.23
        IVA_REDUZIDO_PERC = 0.06

        # 1. Obter inputs do utilizador
        preco_fixo_input_dia = float(st_session_state.get('meu_termo_fixo_gas', 0.0) or 0.0)
        preco_energia_input_kwh = float(st_session_state.get('meu_termo_energia_gas', 0.0) or 0.0)
        
        desc_fixo_perc = float(st_session_state.get('meu_gas_desconto_fixo_perc', 0.0) or 0.0)
        desc_energia_perc = float(st_session_state.get('meu_gas_desconto_energia_perc', 0.0) or 0.0)
        desc_fatura_eur_periodo = float(st_session_state.get('meu_gas_desconto_fatura_eur', 0.0) or 0.0)
        acresc_fatura_eur_periodo = float(st_session_state.get('meu_gas_acrescimo_fatura_eur', 0.0) or 0.0)
        
        # 2. Obter Constantes (TARs, ISP)
        tar_fixo_regulada_base_dia = obter_tar_gas_fixo(escalao_num, constantes_df)
        tar_energia_regulada_base_kwh = obter_tar_gas_energia(escalao_num, constantes_df)
        isp_gas_kwh = isp_gas_valor_manual 

        # 3. Aplicar Tarifa Social (TS) - Afeta TARs e ISP
        tar_fixo_final_pos_ts_dia = tar_fixo_regulada_base_dia
        tar_energia_final_pos_ts_kwh = tar_energia_regulada_base_kwh
        isp_total_s_iva_periodo = consumo_kwh_periodo * isp_gas_kwh 

        desconto_ts_fixo_valor_aplicado = 0.0
        desconto_ts_energia_valor_aplicado = 0.0

        if tarifa_social_ativa and escalao_num in [1, 2]:
            desconto_ts_fixo_bruto = obter_desconto_ts_gas_fixo(escalao_num, constantes_df)
            desconto_ts_energia_bruto = obter_desconto_ts_gas_energia(escalao_num, constantes_df)
            
            tar_fixo_final_pos_ts_dia = max(0.0, tar_fixo_regulada_base_dia - desconto_ts_fixo_bruto)
            tar_energia_final_pos_ts_kwh = max(0.0, tar_energia_regulada_base_kwh - desconto_ts_energia_bruto)
            
            desconto_ts_fixo_valor_aplicado = tar_fixo_regulada_base_dia - tar_fixo_final_pos_ts_dia
            desconto_ts_energia_valor_aplicado = tar_energia_regulada_base_kwh - tar_energia_final_pos_ts_kwh

            isp_total_s_iva_periodo = 0.0
            
        # --- 4. APLICAR LÓGICA DE DESCONTO COMERCIAL E TS (V11) ---
        preco_fixo_pos_desc_perc = preco_fixo_input_dia * (1 - desc_fixo_perc / 100.0)
        preco_energia_pos_desc_perc = preco_energia_input_kwh * (1 - desc_energia_perc / 100.0)
        preco_fixo_final_cobrar_dia = preco_fixo_pos_desc_perc - desconto_ts_fixo_valor_aplicado
        preco_energia_final_cobrar_kwh = preco_energia_pos_desc_perc - desconto_ts_energia_valor_aplicado
        comp_fixo_comerc_final_dia = preco_fixo_final_cobrar_dia - tar_fixo_final_pos_ts_dia
        comp_energia_comerc_final_kwh = preco_energia_final_cobrar_kwh - tar_energia_final_pos_ts_kwh

        # --- 5. MONTAR OS "BALDES" DE IVA ---
        custo_tar_fixo_periodo_s_iva = tar_fixo_final_pos_ts_dia * dias_periodo
        custo_tar_energia_periodo_s_iva = tar_energia_final_pos_ts_kwh * consumo_kwh_periodo
        custo_comerc_fixo_periodo_s_iva = comp_fixo_comerc_final_dia * dias_periodo
        custo_comerc_energia_periodo_s_iva = comp_energia_comerc_final_kwh * consumo_kwh_periodo
        custo_tos_fixo_periodo_s_iva = tos_fixo_dia_val * dias_periodo
        custo_tos_variavel_periodo_s_iva = tos_variavel_kwh_val * consumo_kwh_periodo

        total_base_iva_reduzido = custo_tar_fixo_periodo_s_iva
        total_base_iva_normal = (
            custo_comerc_fixo_periodo_s_iva +
            custo_tar_energia_periodo_s_iva +
            custo_comerc_energia_periodo_s_iva +
            isp_total_s_iva_periodo +
            custo_tos_fixo_periodo_s_iva +
            custo_tos_variavel_periodo_s_iva
        )

        iva_total_reduzido = total_base_iva_reduzido * IVA_REDUZIDO_PERC
        iva_total_normal = total_base_iva_normal * IVA_NORMAL_PERC
        iva_total_periodo = iva_total_reduzido + iva_total_normal

        # 6. Custo Subtotal
        custo_subtotal_c_iva = total_base_iva_reduzido + total_base_iva_normal + iva_total_periodo
        
        # 7. Descontos/Acréscimos Finais de Fatura
        custo_final_com_tudo = custo_subtotal_c_iva - desc_fatura_eur_periodo + acresc_fatura_eur_periodo
        
        # 8. Preços Unitários Finais (para exibição)
        preco_fixo_final_s_iva_dia = preco_fixo_final_cobrar_dia
        preco_energia_final_s_iva_kwh = preco_energia_final_cobrar_kwh

        # --- 9. Construir Dicionários de Tooltip ---
        componentes_tooltip_termo_fixo_dict = {
            'tooltip_fixo_comerc_sem_tar': comp_fixo_comerc_final_dia, 
            'tooltip_fixo_tar_bruta': tar_fixo_regulada_base_dia,
            'tooltip_fixo_ts_aplicada_flag': tarifa_social_ativa and escalao_num in [1, 2],
            'tooltip_fixo_ts_desconto_valor': desconto_ts_fixo_valor_aplicado
        }
        componentes_tooltip_termo_energia_dict = {
            'tooltip_energia_comerc_sem_tar': comp_energia_comerc_final_kwh, 
            'tooltip_energia_tar_bruta': tar_energia_regulada_base_kwh,
            'tooltip_energia_ts_aplicada_flag': tarifa_social_ativa and escalao_num in [1, 2],
            'tooltip_energia_ts_desconto_valor': desconto_ts_energia_valor_aplicado
        }
        componentes_tooltip_custo_total_dict = {
            'tt_cte_energia_siva': custo_tar_energia_periodo_s_iva + custo_comerc_energia_periodo_s_iva,
            'tt_cte_fixo_siva': custo_tar_fixo_periodo_s_iva + custo_comerc_fixo_periodo_s_iva,
            'tt_cte_isp_siva': isp_total_s_iva_periodo,
            'tt_cte_tos_fixo_siva': custo_tos_fixo_periodo_s_iva,
            'tt_cte_tos_var_siva': custo_tos_variavel_periodo_s_iva,
            'tt_cte_total_siva': total_base_iva_reduzido + total_base_iva_normal,
            'tt_cte_valor_iva_6_total': iva_total_reduzido,
            'tt_cte_valor_iva_23_total': iva_total_normal,
            'tt_cte_subtotal_civa': custo_subtotal_c_iva,
            'tt_cte_desc_finais_valor': desc_fatura_eur_periodo,
            'tt_cte_acres_finais_valor': acresc_fatura_eur_periodo
        }
        
        return {
            # --- Colunas Principais para AgGrid ---
            'NomeParaExibir': "O Meu Tarifário (Gás)",
            'Comercializador': "Pessoal",
            'Termo Fixo (€/dia)': round(preco_fixo_final_s_iva_dia, 5),
            'Termo Energia (€/kWh)': round(preco_energia_final_s_iva_kwh, 5),
            'Total Período (€)': round(custo_final_com_tudo, 2),
            'tipo': "Pessoal",
            
            # --- Colunas de Detalhe (para vista normal) ---
            'Segmento': "Pessoal",
            'Faturação': "-",
            'Pagamento': "-",

            # --- NOVAS CHAVES PARA LINKS E NOTAS ---
            'LinkAdesao': "-",
            'info_notas': "Tarifário pessoal configurado pelo utilizador.",

            # --- Dicionários de Tooltip ---
            **componentes_tooltip_termo_fixo_dict,
            **componentes_tooltip_termo_energia_dict,
            **componentes_tooltip_custo_total_dict
        }

    except Exception as e:
        st.error(f"Erro ao calcular 'O Meu Tarifário Gás': {e}")
        return None
    
def calcular_custo_personalizado_gas(
    st_session_state_inputs, # Dicionário de inputs do st.session_state
    consumo_kwh_periodo,    
    dias_periodo,           
    escalao_num,
    tarifa_social_ativa,
    constantes_df,
    tos_fixo_dia_val,        
    tos_variavel_kwh_val,   
    isp_gas_valor_manual
):
    """
    (V1 - Gás) Calcula o custo para o "Tarifário Personalizado", replicando a lógica
    dos outros cálculos de gás (V13) e devolvendo todos os dicionários de tooltip.
    """
    try:
        IVA_NORMAL_PERC = 0.23
        IVA_REDUZIDO_PERC = 0.06

        # 1. Obter inputs do utilizador (via st.session_state_inputs)
        preco_fixo_input_dia = float(st_session_state_inputs.get('pers_gas_fixo', 0.0) or 0.0)
        preco_energia_input_kwh = float(st_session_state_inputs.get('pers_gas_energia', 0.0) or 0.0)
        
        # Flags
        tar_fixo_incluida_flag = st_session_state_inputs.get('pers_gas_tar_fixo', True)
        tar_energia_incluida_flag = st_session_state_inputs.get('pers_gas_tar_energia', True)

        # 2. Obter TARs Reguladas (Base)
        tar_fixo_regulada_base = obter_tar_gas_fixo(escalao_num, constantes_df)
        tar_energia_regulada_base = obter_tar_gas_energia(escalao_num, constantes_df)
        
        # 3. Obter ISP
        isp_gas_kwh = isp_gas_valor_manual 

        # 4. Determinar componentes do Comercializador
        comp_fixo_comercializador_dia = (preco_fixo_input_dia - tar_fixo_regulada_base) if tar_fixo_incluida_flag else preco_fixo_input_dia
        comp_energia_comercializador_kwh = (preco_energia_input_kwh - tar_energia_regulada_base) if tar_energia_incluida_flag else preco_energia_input_kwh

        # 5. Aplicar Tarifa Social (TS)
        tar_fixo_final_a_pagar = tar_fixo_regulada_base
        tar_energia_final_a_pagar = tar_energia_regulada_base
        isp_total_s_iva_periodo = consumo_kwh_periodo * isp_gas_kwh 
        
        desconto_ts_fixo_valor_aplicado = 0.0
        desconto_ts_energia_valor_aplicado = 0.0

        if tarifa_social_ativa and escalao_num in [1, 2]: 
            desconto_ts_fixo_bruto = obter_desconto_ts_gas_fixo(escalao_num, constantes_df)
            desconto_ts_energia_bruto = obter_desconto_ts_gas_energia(escalao_num, constantes_df)
            
            tar_fixo_final_a_pagar = max(0.0, tar_fixo_regulada_base - desconto_ts_fixo_bruto)
            tar_energia_final_a_pagar = max(0.0, tar_energia_regulada_base - desconto_ts_energia_bruto)
            
            desconto_ts_fixo_valor_aplicado = tar_fixo_regulada_base - tar_fixo_final_a_pagar
            desconto_ts_energia_valor_aplicado = tar_energia_regulada_base - tar_energia_final_a_pagar
            
            isp_total_s_iva_periodo = 0.0
            
        # 6. Preços Unitários Finais (Sem IVA)
        preco_fixo_final_s_iva_dia = comp_fixo_comercializador_dia + tar_fixo_final_a_pagar
        preco_energia_final_s_iva_kwh = comp_energia_comercializador_kwh + tar_energia_final_a_pagar

        # 7. Lógica de Custo (IVA)
        custo_tar_fixo_periodo_s_iva = tar_fixo_final_a_pagar * dias_periodo
        custo_comerc_fixo_periodo_s_iva = comp_fixo_comercializador_dia * dias_periodo
        custo_tar_energia_periodo_s_iva = tar_energia_final_a_pagar * consumo_kwh_periodo
        custo_comerc_energia_periodo_s_iva = comp_energia_comercializador_kwh * consumo_kwh_periodo
        custo_tos_fixo_periodo_s_iva = tos_fixo_dia_val * dias_periodo
        custo_tos_variavel_periodo_s_iva = tos_variavel_kwh_val * consumo_kwh_periodo

        total_base_iva_reduzido = custo_tar_fixo_periodo_s_iva
        total_base_iva_normal = (
            custo_comerc_fixo_periodo_s_iva +
            custo_tar_energia_periodo_s_iva +
            custo_comerc_energia_periodo_s_iva +
            isp_total_s_iva_periodo + 
            custo_tos_fixo_periodo_s_iva +
            custo_tos_variavel_periodo_s_iva 
        )

        iva_total_reduzido = total_base_iva_reduzido * IVA_REDUZIDO_PERC
        iva_total_normal = total_base_iva_normal * IVA_NORMAL_PERC
        iva_total_periodo = iva_total_reduzido + iva_total_normal

        custo_final_total_periodo_c_iva = total_base_iva_reduzido + total_base_iva_normal + iva_total_periodo

        # 8. Construir Dicionários de Tooltip
        componentes_tooltip_termo_fixo_dict = {
            'tooltip_fixo_comerc_sem_tar': comp_fixo_comercializador_dia,
            'tooltip_fixo_tar_bruta': tar_fixo_regulada_base,
            'tooltip_fixo_ts_aplicada_flag': tarifa_social_ativa and escalao_num in [1, 2],
            'tooltip_fixo_ts_desconto_valor': desconto_ts_fixo_valor_aplicado
        }
        
        componentes_tooltip_termo_energia_dict = {
            'tooltip_energia_comerc_sem_tar': comp_energia_comercializador_kwh,
            'tooltip_energia_tar_bruta': tar_energia_regulada_base,
            'tooltip_energia_ts_aplicada_flag': tarifa_social_ativa and escalao_num in [1, 2],
            'tooltip_energia_ts_desconto_valor': desconto_ts_energia_valor_aplicado
        }

        componentes_tooltip_custo_total_dict = {
            'tt_cte_energia_siva': custo_tar_energia_periodo_s_iva + custo_comerc_energia_periodo_s_iva,
            'tt_cte_fixo_siva': custo_tar_fixo_periodo_s_iva + custo_comerc_fixo_periodo_s_iva,
            'tt_cte_isp_siva': isp_total_s_iva_periodo,
            'tt_cte_tos_fixo_siva': custo_tos_fixo_periodo_s_iva,
            'tt_cte_tos_var_siva': custo_tos_variavel_periodo_s_iva,
            'tt_cte_total_siva': total_base_iva_reduzido + total_base_iva_normal,
            'tt_cte_valor_iva_6_total': iva_total_reduzido,
            'tt_cte_valor_iva_23_total': iva_total_normal,
            'tt_cte_subtotal_civa': total_base_iva_reduzido + total_base_iva_normal + iva_total_periodo,
            'tt_cte_desc_finais_valor': 0.0,
            'tt_cte_acres_finais_valor': 0.0
        }

        return {
            # --- Colunas Principais para AgGrid ---
            'NomeParaExibir': "Tarifário Personalizado (Gás)",
            'Comercializador': "Personalizado",
            'Termo Fixo (€/dia)': round(preco_fixo_final_s_iva_dia, 5), 
            'Termo Energia (€/kWh)': round(preco_energia_final_s_iva_kwh, 5),
            'Total Período (€)': round(custo_final_total_periodo_c_iva, 2),
            'tipo': "Pessoal", # Usa o tipo 'Pessoal' para partilhar a cor verde/vermelha
            
            'Segmento': "Pessoal",
            'Faturação': "-",
            'Pagamento': "-",
            'LinkAdesao': "-",
            'info_notas': "Tarifário personalizado configurado pelo utilizador.",
            
            # --- Dicionários de Tooltip ---
            **componentes_tooltip_termo_fixo_dict,
            **componentes_tooltip_termo_energia_dict,
            **componentes_tooltip_custo_total_dict
        }
        
    except Exception as e:
        st.error(f"Erro ao calcular 'Tarifário Personalizado Gás': {e}")
        return None
    
def calcular_media_mibgas_datas(df_gwdes, data_inicio, data_fim):
    """
    Calcula o preço médio do MIBGAS (€/MWh) de um DataFrame GWDES para um período específico.
    VERSÃO ATUALIZADA: Assume que a aba GWDES tem preços DIÁRIOS (coluna 'Data') e não horários ('DataHora').
    
    data_inicio e data_fim SÃO objetos datetime.date (vindos do st.date_input).
    """
    if df_gwdes.empty:
        st.warning("A aba 'GWDES' (MIBGAS) está vazia ou não foi carregada.")
        return 0.0

    # --- DEFINIR NOMES DAS COLUNAS ESPERADAS NO EXCEL (NA ABA GWDES) ---
    coluna_data = 'Data'
    coluna_preco_mibgas = 'Preço' 
    # -----------------------------------------------------------------

    # Verificar se as colunas necessárias existem
    if coluna_data not in df_gwdes.columns:
        st.error(f"Erro Crítico: A sua aba 'GWDES' no Excel não tem uma coluna chamada '{coluna_data}'. Não é possível calcular a média MIBGAS.")
        return 0.0
    
    if coluna_preco_mibgas not in df_gwdes.columns:
        st.error(f"Erro Crítico: A sua aba 'GWDES' no Excel não tem uma coluna chamada '{coluna_preco_mibgas}'.")
        return 0.0

    try:
        # 1. Processar a coluna 'Data'. Vindo do Excel, pode já ser datetime.
        # Convertemos para datetime e depois extraímos APENAS a data (.dt.date) para garantir 
        # que removemos quaisquer componentes de hora (como 00:00:00).
        df_gwdes[coluna_data] = pd.to_datetime(df_gwdes[coluna_data], errors='coerce').dt.date
        
        # 2. Garantir que o preço é numérico
        df_gwdes[coluna_preco_mibgas] = pd.to_numeric(df_gwdes[coluna_preco_mibgas], errors='coerce')

        # 3. Remover linhas onde a conversão falhou
        df_gwdes.dropna(subset=[coluna_data, coluna_preco_mibgas], inplace=True)
        if df_gwdes.empty:
            st.error("Aba GWDES processada está vazia (verifique formato de datas e preços).")
            return 0.0

    except Exception as e:
        st.error(f"Erro ao processar dados da aba GWDES: {e}")
        return 0.0

    # 4. Filtrar o DataFrame. Agora comparamos data com data (ambos são datetime.date).
    # Isto é muito mais simples do que o filtro datetime.
    df_periodo = df_gwdes[
        (df_gwdes[coluna_data] >= data_inicio) &
        (df_gwdes[coluna_data] <= data_fim)
    ].copy()

    if df_periodo.empty:
        st.warning(f"Não foram encontrados dados MIBGAS (na aba GWDES) para o período de {data_inicio.strftime('%Y-%m-%d')} a {data_fim.strftime('%Y-%m-%d')}.")
        return 0.0

    # 5. Calcular a média e devolver
    media_mibgas = df_periodo[coluna_preco_mibgas].mean()
    
    if pd.isna(media_mibgas) or media_mibgas == 0.0:
        st.warning("A média MIBGAS calculada é zero ou inválida para o período.")
        return 0.0  # Retorna 0.0 para acionar o fallback (Default das Constantes) no script principal
        
    return round(media_mibgas, 2)