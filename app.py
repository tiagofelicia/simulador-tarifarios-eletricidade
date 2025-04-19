import streamlit as st
import pandas as pd
import datetime

st.set_page_config(page_title="Simulador de Eletricidade", layout="wide")
st.title("🔌 Simulador de Tarifários de Eletricidade")

# --- Inputs principais ---
col1, col2, col3 = st.columns(3)

with col1:
    potencia = st.selectbox("Potência Contratada (kVA)", [
        1.15, 2.3, 3.45, 4.6, 5.75, 6.9, 10.35, 13.8, 17.25, 20.7, 27.6, 34.5, 41.4
    ], index=2)

with col2:
    opcao_horaria = st.selectbox("Opção Horária e Ciclo", [
        "Simples",
        "Bi-horário - Ciclo Diário",
        "Bi-horário - Ciclo Semanal",
        "Tri-horário - Ciclo Diário",
        "Tri-horário - Ciclo Semanal",
        "Tri-horário > 20.7 kVA - Ciclo Diário",
        "Tri-horário > 20.7 kVA - Ciclo Semanal",
    ])

with col3:
    mes = st.selectbox("Mês", [
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
    ], index=datetime.datetime.now().month - 1)

# --- Datas e dias ---
col4, col5, col6 = st.columns(3)

with col4:
    data_inicio = st.date_input("Data Inicial", value=datetime.date(2025, 1, 1))
with col5:
    data_fim = st.date_input("Data Final", value=datetime.date(2025, 1, 31))
with col6:
    dias = (data_fim - data_inicio).days + 1
    st.markdown(f"**Dias calculados:** {dias}")

# --- Consumo conforme tipo tarifário ---
st.subheader("Consumo (kWh)")

consumo = {}
if opcao_horaria == "Simples":
    consumo["simples"] = st.number_input("Consumo Simples", min_value=0.0, value=158.0)

elif opcao_horaria.startswith("Bi"):
    consumo["vazio"] = st.number_input("Consumo em Vazio", min_value=0.0, value=63.0)
    consumo["fora_vazio"] = st.number_input("Consumo em Fora Vazio", min_value=0.0, value=95.0)

elif opcao_horaria.startswith("Tri"):
    consumo["vazio"] = st.number_input("Consumo em Vazio", min_value=0.0, value=63.0)
    consumo["cheias"] = st.number_input("Consumo em Cheias", min_value=0.0, value=68.0)
    consumo["ponta"] = st.number_input("Consumo em Ponta", min_value=0.0, value=27.0)

# --- Checkboxes adicionais ---
st.subheader("Opções")
col7, col8, col9 = st.columns(3)
with col7:
    quota_acp = st.checkbox("Incluir Quota ACP", value=True)
with col8:
    desconto_continente = st.checkbox("Desconto Continente", value=True)
with col9:
    tarifa_social = st.checkbox("Tarifa Social")

familia_numerosa = st.checkbox("Família Numerosa")
comparar = st.checkbox("Comparar 'O Meu Tarifário?'")

# --- Campo para valor MIBEL ---
valor_mibel = st.number_input("Introduzir valor MIBEL/OMIE (€/MWh)", min_value=0.0, step=0.01)

# --- Comparação personalizada ---
if comparar:
    st.markdown("---")
    st.subheader("O Meu Tarifário")

    energia_input = st.number_input("Preço da Energia (€/kWh)", min_value=0.0, step=0.0001, format="%.4f")
    potencia_input = st.number_input("Potência (€/dia)", min_value=0.0, step=0.0001, format="%.4f")
    desconto_energia = st.number_input("Desconto na Energia (%)", min_value=0.0, max_value=100.0)
    desconto_potencia = st.number_input("Desconto na Potência (%)", min_value=0.0, max_value=100.0)
    desconto_fatura = st.number_input("Desconto em fatura (€)", min_value=0.0)

    col10, col11, col12 = st.columns(3)
    with col10:
        tar_in_energia = st.checkbox("TAR incluída na energia", value=True)
    with col11:
        tar_in_potencia = st.checkbox("TAR incluída na potência", value=True)
    with col12:
        tse = st.checkbox("Inclui Financiamento TSE", value=True)

# --- Total de consumo ---
total_consumo = sum(consumo.values())
st.markdown(f"### Total de Consumo: **{total_consumo:.2f} kWh**")
