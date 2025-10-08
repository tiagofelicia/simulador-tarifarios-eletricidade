# --- Carregar as bibliotecas necessárias ---
import pandas as pd
import numpy as np
import requests
import openpyxl
from datetime import datetime
import io
import re

print("✅ Bibliotecas carregadas")

# ===================================================================
# ---- CONFIGURAÇÕES ----
# ===================================================================
DATA_INICIO_ATUALIZACAO = pd.to_datetime("2025-10-01")
FICHEIRO_EXCEL = "Tarifarios_🔌_Eletricidade_Tiago_Felicia.xlsx"
ABA_EXCEL = "OMIE_PERDAS_CICLOS"
COLUNA_PARA_ESCREVER = 11 # Coluna K

print(f"ℹ️ Data de início da atualização definida para: {DATA_INICIO_ATUALIZACAO.date()}")
# ===================================================================

def run_update_process():
    """
    Função principal que encapsula todo o processo de ETL.
    """
    try:
        # ========================================================
        # PASSO 1: Extração de Dados de Futuros (OMIP)
        # ========================================================
        
        print("\n⏳ Passo 1: A extrair dados de futuros do ficheiro OMIPdaily.xlsx...")
        url_omip_excel = "https://www.omip.pt/sites/default/files/dados/eod/omipdaily.xlsx"
        resposta_http = requests.get(url_omip_excel, timeout=20)
        resposta_http.raise_for_status()

        ficheiro_omip_memoria = io.BytesIO(resposta_http.content)
        valor_celula_data = pd.read_excel(ficheiro_omip_memoria, sheet_name="OMIP Daily", header=None, skiprows=4, usecols="E", nrows=1).iloc[0, 0]
        data_relatorio_omip = pd.to_datetime(valor_celula_data, dayfirst=True)
        print(f"   - Data do relatório extraída: {data_relatorio_omip.date()}")

        ficheiro_omip_memoria.seek(0)
        df = pd.read_excel(ficheiro_omip_memoria, sheet_name="OMIP Daily", header=None, skiprows=10, usecols=[1, 10], names=['Nome', 'Preco'])

        df = df.dropna(subset=['Nome'])
        df = df[df['Nome'].str.startswith('FPB')]

        # Classificar tipos de futuros
        conditions = [
            df['Nome'].str.contains(" D "), 
            df['Nome'].str.contains(" Wk"),
            df['Nome'].str.contains(" M "), 
            df['Nome'].str.contains(" Q"),
            df['Nome'].str.contains(" YR-")
        ]
        choices = ["Dia", "Semana", "Mês", "Trimestre", "Ano"]
        df['Classificacao'] = np.select(conditions, choices, default=None)

        df = df.dropna(subset=['Classificacao'])
        df['Preco'] = pd.to_numeric(df['Preco'], errors='coerce')
        df['AnoRaw'] = "20" + df['Nome'].str.extract(r'(\d{2})$')[0]

        # Calcular datas
        datas = []
        for index, row in df.iterrows():
            nome, ano = row['Nome'], row['AnoRaw']
            try:
                if row['Classificacao'] == 'Dia':
                    # Extrair data diretamente do nome (ex: "Th09Oct-25")
                    partes = nome.split(" ")
                    data_str = partes[2] if len(partes) > 2 else partes[1]
                    
                    # Remover sufixo "-25" se existir
                    if '-' in data_str:
                        data_str = data_str.split('-')[0]
                    
                    # Remover prefixo do dia da semana se existir (Th, Fr, We, Sa, Su)
                    if data_str[:2].isalpha():
                        data_str = data_str[2:]
                    
                    # Agora data_str deve ser algo como "09Oct"
                    datas.append(pd.to_datetime(data_str + ano, format='%d%b%Y'))
                    
                elif row['Classificacao'] == 'Semana':
                    week_num = int(nome.split(" Wk")[1].split("-")[0])
                    # Usar a lógica ISO: semana 1 contém 4 de Janeiro
                    # Calcular a segunda-feira dessa semana
                    jan_4 = pd.Timestamp(f'{ano}-01-04')
                    # Encontrar a segunda-feira da semana 1
                    seg_semana_1 = jan_4 - pd.Timedelta(days=jan_4.weekday())
                    # Adicionar semanas
                    data_semana = seg_semana_1 + pd.Timedelta(weeks=week_num - 1)
                    datas.append(data_semana)
                    
                elif row['Classificacao'] == 'Mês':
                    mes_str = nome.split(" ")[2].split("-")[0]
                    datas.append(pd.to_datetime(f'01-{mes_str}-{ano}', format='%d-%b-%Y'))
                    
                elif row['Classificacao'] == 'Trimestre':
                    trimestre = int(nome.split(" Q")[1][0])
                    mes_inicio = (trimestre - 1) * 3 + 1
                    datas.append(pd.to_datetime(f'{ano}-{mes_inicio:02d}-01'))
                    
                elif row['Classificacao'] == 'Ano':
                    datas.append(pd.to_datetime(f'{ano}-01-01'))
                    
                else: 
                    datas.append(pd.NaT)
            except Exception as e:
                print(f"   ⚠️ Erro ao processar '{nome}': {e}")
                datas.append(pd.NaT)

        df['Data'] = datas

        # Finalizar
        dados_web = df.dropna(subset=['Preco', 'Data'])[['Data', 'Preco', 'Classificacao', 'Nome']]
        dados_web = dados_web.drop_duplicates(subset=['Nome'], keep='first').reset_index(drop=True)

        print("✅ Dados de futuros extraídos e processados.")
        print(f"   - Total: {len(dados_web)} futuros")
        print(f"   - Dias: {len(dados_web[dados_web['Classificacao']=='Dia'])}")
        print(f"   - Semanas: {len(dados_web[dados_web['Classificacao']=='Semana'])}")
        print(f"   - Meses: {len(dados_web[dados_web['Classificacao']=='Mês'])}")

        # DEBUG: Mostrar alguns futuros para verificação
        print("\n   📋 Amostra de futuros extraídos:")
        for tipo in ['Dia', 'Semana', 'Mês']:
            amostra = dados_web[dados_web['Classificacao'] == tipo].head(3)
            if not amostra.empty:
                print(f"\n   {tipo}:")
                for _, row in amostra.iterrows():
                    print(f"      {row['Nome']:30s} -> {row['Data'].strftime('%Y-%m-%d (%A)')} = {row['Preco']:.2f} €/MWh")

        # ========================================================
        # PASSO 2: Leitura e Combinação dos Dados OMIE
        # ========================================================

        print("\n⏳ Passo 2: A ler e combinar dados OMIE (todos quarto-horários)...")
        
        fontes_qh = []
        print("   - 2a: A ler 'MIBEL.xlsx'...")
        try:
            dados_base_qh = pd.read_excel("MIBEL.xlsx", usecols=['Data', 'Hora', 'Preço marginal no sistema português (EUR/MWh)'])
            dados_base_qh = dados_base_qh.rename(columns={'Preço marginal no sistema português (EUR/MWh)': 'Preco'})
            dados_base_qh['Data'] = pd.to_datetime(dados_base_qh['Data'])
            fontes_qh.append(dados_base_qh)
        except Exception as e: print(f"   - Aviso: Não foi possível ler 'MIBEL.xlsx'. {e}")
        
        print("   - 2b: A ler dados recentes (ACUM)...")
        try:
            dados_acum_qh = pd.read_csv("https://www.omie.es/sites/default/files/dados/NUEVA_SECCION/INT_PBC_EV_H_ACUM.TXT", sep=';', skiprows=2, header=0, usecols=[0, 1, 3], decimal=',', encoding='windows-1252')
            dados_acum_qh.columns = ['Data', 'Hora', 'Preco']
            dados_acum_qh['Data'] = pd.to_datetime(dados_acum_qh['Data'], format='%d/%m/%Y', errors='coerce')
            fontes_qh.append(dados_acum_qh.dropna())
        except Exception as e: print(f"   - Aviso: Falha ao ler dados (ACUM). {e}")

        print("   - 2c: A ler dados do dia seguinte (INDICADORES)...")
        try:
            r = requests.get("https://www.omie.es/sites/default/files/dados/diario/INDICADORES.DAT", timeout=10)
            linhas = r.content.decode('utf-8').splitlines()
            data_sessao = pd.to_datetime([l for l in linhas if l.startswith("SESION;")][0].split(';')[1], format='%d/%m/%Y')
            linhas_dados = [l for l in linhas if re.match(r'^H\d{2}Q[1-4];', l)]
            if linhas_dados:
                dados_ind_list = [{'Data': data_sessao, 'Hora': (int(l.split(';')[0][1:3])-1)*4+int(l.split(';')[0][4:5]), 'Preco': float(l.split(';')[2].replace(',', '.'))} for l in linhas_dados]
                fontes_qh.append(pd.DataFrame(dados_ind_list))
        except Exception as e: print(f"   - Aviso: Falha ao ler dados (INDICADORES). {e}")
        
        print("   - 2d: A combinar fontes de dados...")
        todos_dados_qh = pd.concat(fontes_qh).drop_duplicates(subset=['Data', 'Hora'], keep='last')
        
        dados_para_manter = todos_dados_qh[todos_dados_qh['Data'] < DATA_INICIO_ATUALIZACAO]
        dados_para_atualizar = todos_dados_qh[todos_dados_qh['Data'] >= DATA_INICIO_ATUALIZACAO]

        dados_combinados_qh = pd.concat([dados_para_manter, dados_para_atualizar]).sort_values(['Data', 'Hora']).reset_index(drop=True)
        print("✅ Todas as fontes de dados OMIE foram combinadas.")

        # =================================================================
        # PASSO 3: Criar calendário e aplicar futuros com a lógica correta
        # =================================================================

        print("\n⏳ Passo 3: A criar calendário e aplicar futuros (lógica do R)...")

        # 3a. Criar calendário base
        calendario_es = pd.DataFrame({
            'Data': pd.date_range(start='2025-01-01', end='2026-12-31', freq='D')
        })
        calendario_es['Ano'] = calendario_es['Data'].dt.year
        calendario_es['Mes'] = calendario_es['Data'].dt.month
        calendario_es['Trimestre'] = calendario_es['Data'].dt.quarter
        calendario_es['Semana'] = calendario_es['Data'].dt.isocalendar().week

        # 3b. Preparar futuros por tipo
        print("   - A preparar futuros diários...")
        dados_web_dia = dados_web[dados_web['Classificacao'] == 'Dia'][['Data', 'Preco']].rename(columns={'Preco': 'Preco_Dia'})
        dados_web_dia = dados_web_dia.drop_duplicates(subset=['Data'], keep='first')

        print("   - A preparar futuros semanais...")
        dados_web_semana = dados_web[dados_web['Classificacao'] == 'Semana'].copy()
        dados_web_semana['Semana'] = dados_web_semana['Data'].dt.isocalendar().week
        dados_web_semana['Ano'] = dados_web_semana['Data'].dt.year
        dados_web_semana = dados_web_semana[['Ano', 'Semana', 'Preco']].rename(columns={'Preco': 'Preco_Semana'})

        print("   - A preparar futuros mensais...")
        dados_web_mes = dados_web[dados_web['Classificacao'] == 'Mês'][['Data', 'Preco']].rename(columns={'Preco': 'Preco_Mes'})

        print("   - A preparar futuros trimestrais...")
        dados_web_trimestre = dados_web[dados_web['Classificacao'] == 'Trimestre'][['Data', 'Preco']].rename(columns={'Preco': 'Preco_Trimestre'})

        # 3c. Juntar futuros ao calendário
        print("   - A fazer merge dos futuros...")

        # Merge semanal (por Ano + Semana)
        calendario_es = pd.merge(calendario_es, dados_web_semana, on=['Ano', 'Semana'], how='left')

        # Merge mensal (por Data - primeiro dia do mês)
        calendario_es = pd.merge(calendario_es, dados_web_mes, on='Data', how='left')

        # Merge trimestral (por Data - primeiro dia do trimestre)
        calendario_es = pd.merge(calendario_es, dados_web_trimestre, on='Data', how='left')

        # 3d. CHAVE: Aplicar fill (propagação) dentro de cada grupo
        print("   - A propagar futuros dentro dos períodos (fill)...")

        # Para semanas: propagar Preco_Semana dentro de cada (Ano, Semana)
        calendario_es['Preco_Semana'] = calendario_es.groupby(['Ano', 'Semana'])['Preco_Semana'].ffill().bfill()

        # Para meses: propagar Preco_Mes dentro de cada (Ano, Mes)
        calendario_es['Preco_Mes'] = calendario_es.groupby(['Ano', 'Mes'])['Preco_Mes'].ffill().bfill()

        # Para trimestres: propagar Preco_Trimestre dentro de cada (Ano, Trimestre)
        calendario_es['Preco_Trimestre'] = calendario_es.groupby(['Ano', 'Trimestre'])['Preco_Trimestre'].ffill().bfill()

        # 3e. Juntar dados históricos reais
        print("   - A juntar dados históricos reais...")
        dados_historicos_diarios = dados_combinados_qh.groupby('Data')['Preco'].mean().rename('Preco_Diario_Real')
        calendario_es = pd.merge(calendario_es, dados_historicos_diarios, left_on='Data', right_index=True, how='left')

        # 3f. Juntar futuros diários (último, porque têm prioridade sobre semanais)
        calendario_es = pd.merge(calendario_es, dados_web_dia, left_on='Data', right_on='Data', how='left')

        # 3g. Aplicar a hierarquia de preços (ordem correta!)
        print("   - A aplicar hierarquia de preços...")
        calendario_es['Preco_Final_Diario'] = (
            calendario_es['Preco_Diario_Real']
            .fillna(calendario_es['Preco_Dia'])
            .fillna(calendario_es['Preco_Semana'])
            .fillna(calendario_es['Preco_Mes'])
            .fillna(calendario_es['Preco_Trimestre'])
        )

        print("✅ Preços diários (reais e projetados) calculados.")

        # 3h. Criar grelha quarto-horária em hora de Espanha
        print("   - A criar grelha quarto-horária...")

        def num_quartos_dia(data):
            """Calcula número de quartos horários considerando DST"""
            tz_es = 'Europe/Madrid'
            dt0 = pd.Timestamp(f"{data} 00:00:00", tz=tz_es)
            dt24 = pd.Timestamp(f"{data} 23:59:59", tz=tz_es)
            horas = (dt24 - dt0).total_seconds() / 3600
            return int(round(horas * 4))

        # Gerar datas futuras a partir da última data histórica
        ultima_data_historica = dados_combinados_qh['Data'].max()
        datas_futuras = pd.date_range(start=ultima_data_historica + pd.Timedelta(days=1), end='2026-01-01', freq='D')

        # Criar tabela de futuros quarto-horários
        futuro_qh = []
        for data in datas_futuras:
            n_quartos = num_quartos_dia(data)
            for hora in range(1, n_quartos + 1):
                futuro_qh.append({'Data': data, 'Hora': hora})

        futuro_qh = pd.DataFrame(futuro_qh)

        # Combinar histórico + futuros
        dados_finais_es = pd.concat([dados_combinados_qh, futuro_qh], ignore_index=True)
        dados_finais_es = dados_finais_es.merge(
            calendario_es[['Data', 'Preco_Final_Diario']], 
            on='Data', 
            how='left'
        )

        # Manter histórico real; preencher apenas futuros
        dados_finais_es['Preco'] = dados_finais_es['Preco'].fillna(dados_finais_es['Preco_Final_Diario'])
        dados_finais_es = dados_finais_es.sort_values(['Data', 'Hora']).reset_index(drop=True)

        print("✅ Estrutura ES criada com número correto de quartos-horários.")

        # DEBUG: Verificar aplicação de futuros
        print("\n   🔍 Verificação da aplicação de futuros:")
        dias_teste = pd.to_datetime(['2025-10-08', '2025-10-09', '2025-10-10', '2025-10-13', '2025-11-01'])
        for dia in dias_teste:
            linha = calendario_es[calendario_es['Data'] == dia]
            if not linha.empty:
                linha = linha.iloc[0]
                print(f"      {dia.strftime('%Y-%m-%d (%A)')}:")
                print(f"         Real: {linha.get('Preco_Diario_Real', 'N/A')}")
                print(f"         Dia:  {linha.get('Preco_Dia', 'N/A')}")
                print(f"         Sem:  {linha.get('Preco_Semana', 'N/A')}")
                print(f"         Mês:  {linha.get('Preco_Mes', 'N/A')}")
                print(f"         ➡️  FINAL: {linha['Preco_Final_Diario']:.2f} €/MWh")
                
        # ============================================================
        # PASSO 4: Conversão para hora de Portugal
        # ============================================================

        print("\n⏳ Passo 4: A converter para hora de Portugal...")

        # Gerar datetime em hora de Espanha
        def gerar_datetime_es(row):
            """Gera timestamp correto considerando DST"""
            data = row['Data']
            hora = row['Hora']
            inicio_dia = pd.Timestamp(f"{data} 00:00:00", tz='Europe/Madrid')
            return inicio_dia + pd.Timedelta(minutes=15 * (hora - 1))

        dados_finais_es['datetime_es'] = dados_finais_es.apply(gerar_datetime_es, axis=1)
        dados_finais_es['datetime_pt'] = dados_finais_es['datetime_es'].dt.tz_convert('Europe/Lisbon')
        dados_finais_es['Data_PT'] = dados_finais_es['datetime_pt'].dt.date

        # Renumerar horas em hora de Portugal
        dados_finais_pt = dados_finais_es.sort_values('datetime_pt').copy()
        dados_finais_pt['Hora_PT'] = dados_finais_pt.groupby('Data_PT').cumcount() + 1

        # Selecionar apenas 2025 e 2026
        dados_finais_pt = dados_finais_pt[dados_finais_pt['datetime_pt'].dt.year.isin([2025, 2026])].copy()
        dados_finais_pt = dados_finais_pt[['Data_PT', 'Hora_PT', 'Preco']].rename(
            columns={'Data_PT': 'Data', 'Hora_PT': 'Hora'}
        )
        dados_finais_pt = dados_finais_pt.dropna(subset=['Preco']).reset_index(drop=True)

        print(f"✅ {len(dados_finais_pt)} registos finais preparados em hora de Portugal.")

        # Validação de quartos
        check_quartos = dados_finais_pt.groupby('Data').size().reset_index(name='n')
        dias_estranhos = check_quartos[~check_quartos['n'].isin([92, 96, 100])]

        if not dias_estranhos.empty:
            print("⚠️ Aviso: Dias com número de quartos inesperado:")
            print(dias_estranhos.to_string(index=False))
        else:
            print("✅ Todos os dias têm número de quartos esperado (92, 96 ou 100).")

        # ============================================================
        # PASSO 5: Atualização do ficheiro Excel
        # ============================================================

        print(f"\n⏳ Passo 5: A atualizar o ficheiro '{FICHEIRO_EXCEL}'...")
        
        wb = openpyxl.load_workbook(FICHEIRO_EXCEL)
        sheet = wb[ABA_EXCEL]
        for i, preco in enumerate(dados_finais_pt['Preco'].tolist()):
            sheet.cell(row=i + 2, column=COLUNA_PARA_ESCREVER, value=preco)
            
        sheet_const = wb["Constantes"]
        ultima_data_omie = dados_combinados_qh['Data'].max()
        sheet_const['B90'] = ultima_data_omie.strftime('%d/%m/%Y')
        sheet_const['B91'] = data_relatorio_omip.strftime('%d/%m/%Y')

        wb.save(FICHEIRO_EXCEL)
        print(f"✅ O ficheiro Excel foi atualizado com sucesso!\n   Data_Valores_OMIE = {ultima_data_omie.date()}\n   Data_Valores_OMIP = {data_relatorio_omip.date()}")

    except Exception as e:
        import traceback
        print(f"❌ Ocorreu um erro inesperado: {e}")
        traceback.print_exc()

if __name__ == "__main__":
    run_update_process()
