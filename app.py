import streamlit as st
import pandas as pd
import datetime

st.set_page_config(page_title="Simulador de Eletricidade", layout="wide")
st.title("🔌 Simulador de Tarifários de Eletricidade")

# --- Inputs principais ---
st.sidebar.header("Configurações Gerais")

potencias = [1.15, 2.3, 3.45, 4.6, 5.75, 6.9, 10.35, 13.8, 17.25, 20.7, 27.6, 34.5, 41.4]
opcoes_horarias = {
    "Simples": "simples",
    "Bi-horário - Ciclo Diário": "bi-diario",
    "Bi-horário - Ciclo Semanal": "bi-semanal",
    "Tri-horário - Ciclo Diário": "tri-diario",
    "Tri-horário - Ciclo Semanal": "tri-semanal",
    "Tri-horário > 20.7 kVA - Ciclo Diário": "tri-207-diario",
    "Tri-horário > 20.7 kVA - Ciclo Semanal": "tri-207-semanal",
}

potencia = st.sidebar.selectbox("Potência Contratada (kVA):", potencias, index=2)
opcao_horaria = st.sidebar.selectbox("Opção Horária e Ciclo:", list(opcoes_horarias.keys()))

meses = {
    "Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4, "Maio": 5, "Junho": 6,
    "Julho": 7, "Agosto": 8, "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12
}
mes = st.sidebar.selectbox("Mês:", list(meses.keys()), index=datetime.datetime.now().month - 1)

data_inicial = st.sidebar.date_input("Data Inicial", value=None)
data_final = st.sidebar.date_input("Data Final", value=None)

num_dias = st.sidebar.number_input("Nº de Dias:", min_value=1, value=30)
valor_mibel = st.sidebar.number_input("Valor MIBEL/OMIE (€/MWh):", min_value=0.0, step=0.01)

st.sidebar.markdown("---")

quota_acp = st.sidebar.checkbox("Incluir Quota ACP", value=True)
desconto_continente = st.sidebar.checkbox("Desconto Continente", value=True)
comparar = st.sidebar.checkbox("Comparar 'O Meu Tarifário?'", value=False)
tarifa_social = st.sidebar.checkbox("Tarifa Social")
familia_numerosa = st.sidebar.checkbox("Família Numerosa")

# --- Consumos ---
st.subheader("Consumos")
tipo = opcoes_horarias[opcao_horaria]

if tipo == "simples":
    consumo_simples = st.number_input("Consumo Simples (kWh)", min_value=0.0, value=158.0)
elif tipo.startswith("bi"):
    consumo_vazio = st.number_input("Consumo em Vazio (kWh)", min_value=0.0, value=63.0)
    consumo_fora = st.number_input("Consumo em Fora Vazio (kWh)", min_value=0.0, value=95.0)
elif tipo.startswith("tri"):
    consumo_vazio = st.number_input("Consumo em Vazio (kWh)", min_value=0.0, value=63.0)
    consumo_cheias = st.number_input("Consumo em Cheias (kWh)", min_value=0.0, value=68.0)
    consumo_ponta = st.number_input("Consumo em Ponta (kWh)", min_value=0.0, value=27.0)

# --- Cálculo de dias automático ---
if data_inicial and data_final and data_final > data_inicial:
    num_dias = (data_final - data_inicial).days + 1
    st.info(f"Dias calculados automaticamente: {num_dias} dias")
elif not data_inicial and not data_final:
    dias_no_mes = (datetime.date(2025, meses[mes], 1).replace(day=28) + datetime.timedelta(days=4)).replace(day=1) - datetime.timedelta(days=1)
    num_dias = dias_no_mes.day
    st.info(f"Dias do mês selecionado: {num_dias} dias")

# --- Resumo ---
st.subheader("Resumo da Simulação")
resumo = {
    "Potência Contratada (kVA)": potencia,
    "Opção Horária": opcao_horaria,
    "Nº de Dias": num_dias,
    "Tarifa Social": tarifa_social,
    "Família Numerosa": familia_numerosa,
    "Quota ACP": quota_acp,
    "Comparar Tarifário": comparar
}

if tipo == "simples":
    resumo["Consumo Simples"] = consumo_simples
elif tipo.startswith("bi"):
    resumo["Consumo Vazio"] = consumo_vazio
    resumo["Consumo Fora Vazio"] = consumo_fora
elif tipo.startswith("tri"):
    resumo["Consumo Vazio"] = consumo_vazio
    resumo["Consumo Cheias"] = consumo_cheias
    resumo["Consumo Ponta"] = consumo_ponta

st.json(resumo)

st.warning("🔧 Em breve: cálculo com base em tarifários automáticos e Google Sheets")
