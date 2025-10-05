# --- Carregar as bibliotecas necessárias ---
library(dplyr)
library(stringr)
library(lubridate)
library(httr)
library(rvest)
library(tidyr)
library(purrr)
library(readxl)
library(openxlsx)
library(readr)

cat("✅ Bibliotecas carregadas\n")

# ===================================================================
# ---- CONFIGURAÇÕES ----
# ===================================================================
data_inicio_atualizacao <- as.Date("2025-10-01")
ficheiro_excel <- "Tarifarios_🔌_Eletricidade_Tiago_Felicia.xlsx"
aba_excel <- "OMIE_PERDAS_CICLOS"
coluna_para_escrever <- 11 # Coluna K
cat(paste0("ℹ️ Data de início da atualização definida para: ", data_inicio_atualizacao, "\n"))
# ===================================================================

### Passo 1: Extração de Dados de Futuros (OMIP a partir do ficheiro Excel)

cat("⏳ A extrair dados de futuros do ficheiro OMIPdaily.xlsx...\n")

# URL e caminho local mantêm-se
url_omip_excel <- "https://www.omip.pt/sites/default/files/dados/eod/omipdaily.xlsx"
caminho_local_ficheiro <- "omipdaily_temp.xlsx"

cat("   - 1/4: A descarregar o ficheiro do URL...\n")
resposta_http <- tryCatch({
  GET(url_omip_excel, write_disk(caminho_local_ficheiro, overwrite = TRUE), timeout(20))
}, error = function(e) { stop("ERRO de Download: ", e$message) })

if (status_code(resposta_http) != 200) {
  stop(paste0("ERRO de Download: Código de estado ", status_code(resposta_http)))
}
cat("   - 2/4: Ficheiro descarregado. A extrair a data do relatório...\n")
data_relatorio_omip <- tryCatch({
  data_lida <- read_excel(caminho_local_ficheiro, sheet = "OMIP Daily", range = "E5", col_names = FALSE)
  # Usar lubridate::dmy para uma conversão de data mais segura
  lubridate::dmy(pull(data_lida))
}, error = function(e) {
  warning("Aviso: Não foi possível ler a data do relatório OMIP. A usar a data de hoje como alternativa.")
  Sys.Date()
})
cat(paste0("   - Data do relatório extraída: ", format(data_relatorio_omip, "%Y-%m-%d"), "\n"))

cat("   - 3/4: Ficheiro descarregado. A ler os dados...\n")
dados_futuros_raw <- tryCatch({
  read_excel(
    caminho_local_ficheiro,
    sheet = "OMIP Daily", 
    skip = 10,            
    col_names = FALSE     
  )
}, error = function(e) { stop("ERRO de Leitura: ", e$message) })

cat("   - 4/4: A processar, limpar e desdobrar os dados dos futuros...\n")

# Passo A: Processamento inicial e cálculo de data unificado
dados_web_processados <- dados_futuros_raw %>%
  select(Nome = ...2, Preco = ...11) %>%
  filter(!is.na(Nome), str_starts(Nome, "FPB")) %>%
  mutate(
    Preco = as.numeric(Preco),
    Classificacao = case_when(
      str_detect(Nome, " WE ") ~ "FimDeSemana",
      str_detect(Nome, " D ") ~ "Dia",
      str_detect(Nome, " Wk") ~ "Semana",
      str_detect(Nome, " M ") ~ "Mês",
      str_detect(Nome, " Q") ~ "Trimestre",
      str_detect(Nome, " YR-") ~ "Ano"
    ),
    AnoRaw = paste0("20", str_extract(Nome, "\\d{2}$")),
    # LÓGICA DE DATA CORRIGIDA E UNIFICADA
    Data = case_when(
      # Para Dia e Fim de Semana, a data é extraída diretamente
      Classificacao %in% c("Dia", "FimDeSemana") ~ dmy(paste0(str_extract(Nome, "\\d{2}[A-Za-z]{3}"), "-", AnoRaw), quiet = TRUE),
      # Para os outros, calculamos o início do período
      Classificacao == "Semana" ~ floor_date(make_date(year = AnoRaw, month = 1, day = 1) + weeks(as.numeric(str_extract(Nome, "(?<=Wk)\\d+")) - 1), "week", week_start = 1),
      Classificacao == "Mês" ~ floor_date(my(paste(str_extract(Nome, "[A-Z][a-z]{2}"), AnoRaw)), "month"),
      Classificacao == "Trimestre" ~ floor_date(make_date(year = AnoRaw, month = (as.numeric(str_extract(Nome, "(?<= Q)\\d")) - 1) * 3 + 1, day = 1), "quarter"),
      Classificacao == "Ano" ~ floor_date(make_date(year = AnoRaw, month = 1, day = 1), "year")
    )
  ) %>%
  filter(!is.na(Preco), !is.na(Data))

# Passo B: Separar e desdobrar os fins de semana
dados_we <- dados_web_processados %>% filter(Classificacao == "FimDeSemana")
dados_normais <- dados_web_processados %>% filter(Classificacao != "FimDeSemana")

if (nrow(dados_we) > 0) {
  dados_we_desdobrados <- dados_we %>%
    rowwise() %>%
    reframe(
      tibble(
        Nome = c(
          paste0(str_sub(Nome, 1, 6), "Sa", str_sub(Nome, 9)), 
          paste0(str_sub(Nome, 1, 6), "Su", str_sub(Nome, 9))
        ),
        Preco = Preco,
        Classificacao = "Dia", # Ambos são agora "Dia"
        AnoRaw = AnoRaw,
        Data = c(Data, Data + days(1)) # Sábado (data original) e Domingo (dia seguinte)
      )
    ) %>%
    ungroup()
  
  dados_juntos <- bind_rows(dados_normais, dados_we_desdobrados)
} else {
  dados_juntos <- dados_web_processados
}

# Passo C: Finalizar o dataframe 'dados_web' para ser usado nos passos seguintes
dados_web <- dados_juntos %>%
  mutate(
    Ano = year(Data),
    Classificacao = factor(Classificacao, levels = c("Dia", "Semana", "Mês", "Trimestre", "Ano"))
  ) %>%
  select(Nome, Preco, Classificacao, Data, Ano) %>%
  distinct(Nome, .keep_all = TRUE)

cat("✅ Dados de futuros extraídos, processados e com priorização diária restaurada.\n")

### Passo 2: Leitura e Combinação dos Dados OMIE (Tudo em Hora de Espanha)
cat("⏳ 2a: A ler o ficheiro histórico 'MIBEL.xlsx' (em hora de Espanha)...\n")
dados_base_qh <- tryCatch({
  read_excel("MIBEL.xlsx") %>%
    select(Data, Hora, Preco = `Preço marginal no sistema português (EUR/MWh)`) %>%
    mutate(Data = as.Date(Data))
}, error = function(e) { stop("ERRO: 'MIBEL.xlsx'. Detalhes: ", e$message) })

cat("⏳ 2b: A ler dados recentes do ano (INT_PBC_EV_H_ACUM.TXT)...\n")
url_acum <- "https://www.omie.es/sites/default/files/dados/NUEVA_SECCION/INT_PBC_EV_H_ACUM.TXT"
dados_acum_qh <- tryCatch({
  read_delim(url_acum, delim = ";", col_names = FALSE, skip = 2, col_types = cols(.default = col_character()), locale = locale(encoding = "windows-1252")) %>%
    select(DataStr = X1, Hora = X2, PrecoStr = X4) %>%
    mutate(
      Data = dmy(DataStr),
      Hora = as.integer(Hora),
      Preco = as.numeric(str_replace(PrecoStr, ",", "."))
    ) %>%
    filter(!is.na(Data), !is.na(Preco)) %>%
    select(Data, Hora, Preco)
}, error = function(e) { warning("Aviso: Falha ao ler dados recentes."); tibble() })

cat("⏳ 2c: A ler dados do dia seguinte (INDICADORES.DAT)...\n")
url_ind <- "https://www.omie.es/sites/default/files/dados/diario/INDICADORES.DAT"
dados_ind_qh <- tryCatch({
  linhas <- readLines(url_ind, encoding = "UTF-8")
  linha_sesion <- grep("^SESION;", linhas, value = TRUE)
  data_sessao <- dmy(strsplit(linha_sesion, ";")[[1]][2])
  linhas_dados <- grep("^H[0-9]{2}Q[1-4];", linhas, value = TRUE)
  if (length(linhas_dados) > 0) {
    dados <- strsplit(linhas_dados, ";")
    id_str <- sapply(dados, function(x) x[1])
    hora_val <- as.numeric(str_sub(id_str, 2, 3))
    quarto_val <- as.numeric(str_sub(id_str, 5, 5))
    tibble(
      Data = data_sessao,
      Hora = (hora_val - 1) * 4 + quarto_val,
      Preco = as.numeric(gsub(",", ".", sapply(dados, function(x) x[3])))
    )
  } else { tibble() }
}, error = function(e) { warning("Aviso: Falha ao ler dados do dia seguinte."); tibble() })

cat("⏳ 2d: A combinar fontes de dados com proteção de histórico...\n")
dados_internet_qh <- bind_rows(dados_acum_qh, dados_ind_qh) %>% distinct(Data, Hora, .keep_all = TRUE)
dados_para_manter <- dados_base_qh %>% filter(Data < data_inicio_atualizacao)
dados_base_para_atualizar <- dados_base_qh %>% filter(Data >= data_inicio_atualizacao)
dados_internet_para_atualizar <- dados_internet_qh %>% filter(Data >= data_inicio_atualizacao)
dados_atualizados <- bind_rows(dados_base_para_atualizar, dados_internet_para_atualizar) %>% distinct(Data, Hora, .keep_all = TRUE)
dados_combinados_qh <- bind_rows(dados_para_manter, dados_atualizados)

dados_historicos_diarios <- dados_combinados_qh %>%
  group_by(Data) %>%
  summarise(Preco_Diario_Real = mean(Preco, na.rm = TRUE)) %>%
  ungroup()
cat("✅ Todas as fontes de dados OMIE foram combinadas e processadas.\n")

### Passo 3: Criação do Dataframe de Datas e Junção dos Futuros (em Hora de Espanha)
calendario <- tibble(Data = seq(as.Date("2025-01-01"), as.Date("2026-12-31"), by = "day")) %>%
  mutate(Ano = year(Data), Mes = month(Data), Trimestre = quarter(Data), Semana = isoweek(Data))
cat("⏳ A combinar dados históricos e futuros...\n")
dados_web_dia <- dados_web %>% 
  filter(Classificacao == "Dia") %>% 
  select(Data, Preco_Dia = Preco) %>%
  distinct(Data, .keep_all = TRUE) # <-- ADICIONAR ESTA LINHA
dados_web_semana <- dados_web %>% filter(Classificacao == "Semana") %>% mutate(Semana = isoweek(Data)) %>% select(Ano, Semana, Preco_Semana = Preco)
df_futuros <- calendario %>%
  left_join(dados_web_semana, by = c("Ano", "Semana")) %>%
  left_join(dados_web %>% filter(Classificacao == "Mês") %>% select(Data, Preco_Mes = Preco), by = "Data") %>%
  left_join(dados_web %>% filter(Classificacao == "Trimestre") %>% select(Data, Preco_Trimestre = Preco), by = "Data") %>%
  group_by(Ano, Semana) %>% fill(Preco_Semana, .direction = "downup") %>% ungroup() %>%
  group_by(Ano, Mes) %>% fill(Preco_Mes, .direction = "downup") %>% ungroup() %>%
  group_by(Ano, Trimestre) %>% fill(Preco_Trimestre, .direction = "downup") %>% ungroup()

dados_diarios_finais <- calendario %>%
  left_join(dados_historicos_diarios, by = "Data") %>%
  left_join(dados_web_dia, by = "Data") %>%
  left_join(df_futuros %>% select(Data, Preco_Semana, Preco_Mes, Preco_Trimestre), by = "Data") %>%
  mutate(
    Preco_Final_Diario = coalesce(Preco_Diario_Real, Preco_Dia, Preco_Semana, Preco_Mes, Preco_Trimestre)
  )
cat("✅ Preços diários (reais e projetados) calculados.\n")

# ===================================================================
# ---- PASSO 4 a 6: TABELA FINAL QUARTO-HORÁRIOS + CONVERSÃO PT ----
# ===================================================================

cat("⏳ Passo 4: Preparar estrutura completa de quartos-horários (ES)...\n")

# Função para número de quartos por dia (92, 96, 100)
num_quartos_dia <- function(data) {
  tz_es <- "Europe/Madrid"
  dt0 <- as.POSIXct(paste0(data, " 00:00:00"), tz = tz_es)
  dt24 <- as.POSIXct(paste0(data, " 23:59:59"), tz = tz_es)
  # Diferença de horas multiplicada por 4
  as.integer(round(as.numeric(difftime(dt24, dt0, units = "hours")) * 4))
}

# Datas futuras a partir da última histórica
datas_futuras <- seq(from = max(dados_combinados_qh$Data) + 1, to = as.Date("2026-01-01"), by = "day")

# Criar tabela de futuros quarto-horários
futuro_qh <- map_dfr(datas_futuras, function(d) {
  tibble(Data = d, Hora = 1:num_quartos_dia(d))
}) %>%
  # Evitar sobreposição com histórico
  anti_join(dados_combinados_qh, by = c("Data", "Hora"))

# Combinar histórico + futuros
dados_finais_es <- bind_rows(dados_combinados_qh, futuro_qh) %>%
  left_join(dados_diarios_finais %>% select(Data, Preco_Final_Diario), by = "Data") %>%
  mutate(
    # Mantém histórico real; preenche apenas com projeção diária/futura
    Preco = coalesce(Preco, Preco_Final_Diario)
  ) %>%
  arrange(Data, Hora)

cat("✅ Estrutura ES criada com número correto de quartos-horários e projeções aplicadas apenas onde necessário.\n")

cat("⏳ Passo 5: A converter dados finais para a hora de Portugal (respeitando DST e quartos-horários irregulares)...\n")

# --- 1) Combinar histórico + futuros preenchidos ---
dados_completos_es <- bind_rows(
  dados_finais_es,  # histórico + futuros gerados anteriormente
  futuro_qh %>% 
    left_join(dados_diarios_finais %>% select(Data, Preco_Final_Diario), by = "Data") %>%
    mutate(Preco = Preco_Final_Diario)
) %>%
  arrange(Data, Hora) %>%
  distinct(Data, Hora, .keep_all = TRUE)

# --- 2) Gerar datetime ES e converter para PT ---
dados_finais_pt <- dados_completos_es %>%
  filter(!is.na(Preco)) %>%
  group_by(Data) %>%
  arrange(Hora, .by_group = TRUE) %>%
  mutate(
    datetime_es = {
      seqs <- seq(
        as.POSIXct(paste0(unique(Data), " 00:00:00"), tz = "Europe/Madrid"),
        as.POSIXct(paste0(unique(Data + 1), " 00:00:00"), tz = "Europe/Madrid") - minutes(15),
        by = "15 min"
      )
      seqs[Hora]  # mapear cada Hora para o timestamp correto
    },
    datetime_pt = with_tz(datetime_es, "Europe/Lisbon"),
    Data_PT = as.Date(datetime_pt)
  ) %>%
  ungroup() %>%
  arrange(datetime_pt) %>%
  group_by(Data_PT) %>%
  mutate(Hora_PT = row_number()) %>%
  ungroup() %>%
  select(Data = Data_PT, Hora = Hora_PT, Preco) %>%
  filter(year(Data) %in% c(2025, 2026))

cat(paste0("✅ ", nrow(dados_finais_pt), " registos finais preparados em hora de Portugal.\n"))

# --- 3) Função de validação automática de quartos ---
validar_quartos_dia <- function(dias, df_final_pt) {
  dias <- as.Date(dias)
  check <- df_final_pt %>%
    filter(Data %in% dias) %>%
    count(Data) %>%
    mutate(
      tipo = case_when(
        n == 92 ~ "Primavera (92 quartos)",
        n == 96 ~ "Normal (96 quartos)",
        n == 100 ~ "Outono (100 quartos)",
        TRUE ~ paste0("Inesperado: ", n)
      )
    )
  
  cat("⚡ Validação automática de quartos-horários:\n")
  print(check, n = Inf)
  return(invisible(check))
}

# --- 4) Check automático de todos os dias futuros críticos ---
# Ajustar as datas de mudança de hora reais do ano em questão
datas_mudanca_hora <- c("2025-03-30", "2025-10-26", "2026-03-29", "2026-10-25")
validar_quartos_dia(datas_mudanca_hora, dados_finais_pt)

# --- 5) Check geral de todos os dias com número diferente do normal ---
check_quartos <- dados_finais_pt %>%
  count(Data) %>%
  mutate(tipo = case_when(
    n == 92 ~ "Primavera (92 quartos)",
    n == 96 ~ "Normal (96 quartos)",
    n == 100 ~ "Outono (100 quartos)",
    TRUE ~ paste0("Inesperado: ", n)
  ))

dias_estranhos <- check_quartos %>% filter(tipo != "Normal (96 quartos)")
if(nrow(dias_estranhos) > 0){
  cat("⚠️ Aviso: Dias com número de quartos diferente do normal (96):\n")
  print(dias_estranhos, n = Inf)
} else {
  cat("✅ Todos os dias têm número de quartos esperado.\n")
}

# ===================================================================
# ---- PASSO 6: Atualização do ficheiro Excel ----
# ===================================================================
cat(paste0("⏳ A atualizar o ficheiro '", ficheiro_excel, "'...\n"))

tryCatch({
  # Abrir workbook
  wb <- loadWorkbook(ficheiro_excel)
  
  # --- 1) Escrever os valores finais na coluna definida ---
  dados_finais_para_excel <- dados_finais_pt %>% pull(Preco) %>% as.data.frame()
  writeData(
    wb,
    sheet = aba_excel,
    x = dados_finais_para_excel,
    startCol = coluna_para_escrever,
    startRow = 2,
    colNames = FALSE
  )
  
  # --- 2) Atualizar datas de referência na aba 'Constantes' ---
  ultima_data_omie <- max(dados_historicos_diarios$Data, na.rm = TRUE)
  writeData(wb, sheet = "Constantes", x = format(ultima_data_omie, "%m/%d/%Y"), startCol = 2, startRow = 81, colNames = FALSE)
  writeData(wb, sheet = "Constantes", x = format(data_relatorio_omip, "%m/%d/%Y"), startCol = 2, startRow = 82, colNames = FALSE)
  
  # --- 3) Salvar workbook ---
  saveWorkbook(wb, ficheiro_excel, overwrite = TRUE)
  
  cat(paste0("✅ O ficheiro Excel foi atualizado com sucesso!\n",
             "   Data_Valores_OMIE = ", ultima_data_omie, "\n",
             "   Data_Valores_OMIP = ", data_relatorio_omip, "\n"))
  
}, error = function(e) {
  stop("ERRO: Falha ao escrever no ficheiro Excel. Detalhes: ", e$message)
})

cat(paste0("🏁 Atualização concluída em ", round(difftime(Sys.time(), start_time, units = "mins"), 1), " minutos.\n"))
