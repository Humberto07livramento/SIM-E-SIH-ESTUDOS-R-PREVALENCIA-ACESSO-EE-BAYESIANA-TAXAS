# -------------------------------------------------------------------------
# SCRIPT COMPLETO: ANÁLISE EPIDEMIOLÓGICA COM FICHEIRO ÚNICO E LAUDO
# Estudo Comparativo: Região de Serrinha vs Demais Regiões
# Autor: Humberto Livramento
# -------------------------------------------------------------------------

# 1. Configuração do Diretório de Trabalho
setwd("C:/Users/humbertolivramento/Documents/estudo serrinha")
cat("Diretório configurado.\n")

# 2. Instalação e Carregamento de Pacotes (Rigor Estatístico)
if(!require(pacman)) install.packages("pacman")
pacman::p_load(tidyverse, sandwich, lmtest, writexl)
cat("Pacotes carregados com sucesso.\n")

# 3. Leitura do Arquivo Único de Óbitos Maternos
cat("Lendo o arquivo maternas_obitos_2015_2024.csv...\n")
obitos <- read.csv("maternas_obitos_2015_2024.csv", sep = ";", stringsAsFactors = FALSE)

# 4. Definição da Região de Serrinha (Códigos IBGE 6 dígitos)
codigos_serrinha <- c(
  290230, # Araci
  290327, # Barrocas
  290360, # Biritinga
  290680, # Cansanção
  290840, # Conceição do Coité
  291680, # Itiúba
  291880, # Lamarão
  292150, # Monte Santo
  292265, # Nordestina
  292590, # Queimadas
  292740, # Retirolândia
  292800, # Santaluz
  292890, # São Domingos
  292930, # Serrinha
  293150, # Teofilândia
  293190, # Tucano
  293240  # Valente
)

# 5. Limpeza e Recodificação das Variáveis (Dicionário SIM)
cat("Processando e recodificando as variáveis socioeconômicas...\n")
dados_analise <- obitos %>%
  mutate(
    # Exposição: Região de Serrinha (1) vs Demais (0)
    RegiaoSerrinha = ifelse(CODMUNRES %in% codigos_serrinha, 1, 0),
    
    # Idade: Códigos 400 a 499 indicam anos de vida
    idade_anos = ifelse(IDADE >= 400 & IDADE < 500, IDADE - 400, NA),
    Idade_Risco = ifelse(idade_anos < 20 | idade_anos >= 35, 1, 0),
    
    # Raça/Cor: Negras (Pretas e Pardas) vs Brancas
    Raca_Negra = ifelse(RACACOR %in% c(2, 4), 1, ifelse(RACACOR == 1, 0, NA)),
    
    # Escolaridade: Baixa Escolaridade (Até 7 anos) vs Alta (8+ anos)
    Baixa_Escolaridade = ifelse(ESC %in% c(1, 2, 3), 1, ifelse(ESC %in% c(4, 5), 0, NA)),
    
    # Estado Civil: Sem Cônjuge vs Com Cônjuge
    Sem_Conjuge = ifelse(ESTCIV %in% c(1, 3, 4), 1, ifelse(ESTCIV %in% c(2, 5), 0, NA))
  )

# 6. Função Epidemiológica: Regressão de Poisson com Variância Robusta
calcular_rp <- function(df, var_resposta, nome_variavel) {
  df_modelo <- df %>% filter(!is.na(df[[var_resposta]]), !is.na(RegiaoSerrinha))
  
  modelo <- glm(as.formula(paste(var_resposta, "~ RegiaoSerrinha")), 
                family = poisson(link = "log"), data = df_modelo)
  
  cov_robusta <- vcovHC(modelo, type = "HC0")
  coefs <- coeftest(modelo, vcov. = cov_robusta)
  
  est <- exp(coefs["RegiaoSerrinha", "Estimate"])
  ic_inf <- exp(coefs["RegiaoSerrinha", "Estimate"] - 1.96 * coefs["RegiaoSerrinha", "Std. Error"])
  ic_sup <- exp(coefs["RegiaoSerrinha", "Estimate"] + 1.96 * coefs["RegiaoSerrinha", "Std. Error"])
  p_val <- coefs["RegiaoSerrinha", "Pr(>|z|)"]
  
  n_serr <- sum(df_modelo[[var_resposta]][df_modelo$RegiaoSerrinha == 1] == 1)
  t_serr <- sum(df_modelo$RegiaoSerrinha == 1)
  prev_serr <- (n_serr / t_serr) * 100
  
  n_dem <- sum(df_modelo[[var_resposta]][df_modelo$RegiaoSerrinha == 0] == 1)
  t_dem <- sum(df_modelo$RegiaoSerrinha == 0)
  prev_dem <- (n_dem / t_dem) * 100
  
  data.frame(
    Variavel = nome_variavel,
    Prev_Demais = sprintf("%.1f%%", prev_dem),
    Prev_Serrinha = sprintf("%.1f%%", prev_serr),
    Razao_Prevalencia = round(est, 2),
    IC_95 = sprintf("[%.2f; %.2f]", ic_inf, ic_sup),
    P_Valor = ifelse(p_val < 0.001, "< 0.001", sprintf("%.4f", p_val))
  )
}

# 7. Execução das Análises
cat("Calculando a Razão de Prevalência (RP) com Modelos de Poisson...\n")
res_raca   <- calcular_rp(dados_analise, "Raca_Negra", "Raça Negra (Preta/Parda) vs Branca")
res_esc    <- calcular_rp(dados_analise, "Baixa_Escolaridade", "Baixa Escolaridade (<= 7 anos) vs Alta")
res_idade  <- calcular_rp(dados_analise, "Idade_Risco", "Idade de Risco (<20 ou >=35) vs 20-34 anos")
res_estciv <- calcular_rp(dados_analise, "Sem_Conjuge", "Sem Cônjuge vs Com Cônjuge")

# Aba 1: Tabela Consolidada de Resultados
tabela_final <- bind_rows(res_raca, res_esc, res_idade, res_estciv)

cat("\n--- RESULTADOS FINAIS ---\n")
print(tabela_final)

# 8. Criação do Laudo Epidemiológico Formal (Aba 2)
# Estruturando um Data Frame para gerar um texto claro e acadêmico no Excel
tabela_laudo <- data.frame(
  Secao = c(
    "1. OBJETIVO", 
    "2. METODOLOGIA", 
    "3. RESULTADOS OBSERVADOS", 
    "4. INTERPRETAÇÃO EPIDEMIOLÓGICA", 
    "5. CONCLUSÃO E RECOMENDAÇÃO"
  ),
  Descricao = c(
    "Avaliar a associação entre características socioeconômicas/demográficas e a ocorrência de óbito materno na Região de Serrinha em comparação às demais regiões.",
    "Utilizou-se Regressão de Poisson com estimador de variância robusta (Sanduíche HC0) para obtenção de estimativas precisas de Razão de Prevalência (RP), mitigando o viés de superestimação do Odds Ratio (OR). Adotou-se Intervalo de Confiança de 95% e nível de significância de 5% (p < 0,05).",
    "As estimativas demonstraram que: Raça Negra (RP=1,04; p=0,1480); Baixa Escolaridade (RP=0,98; p=0,8333); Idade de Risco (RP=1,13; p=0,1028) e Ausência de Cônjuge (RP=1,05; p=0,5191). Todas as variáveis apresentaram Intervalos de Confiança que interceptam a unidade (1,00).",
    "Na epidemiologia clássica, quando o IC 95% atravessa o valor 1,00 e o p-valor é superior a 0,05, considera-se a hipótese nula. Isso significa que não há associação estatisticamente significativa entre as vulnerabilidades socioeconômicas testadas e o fato do óbito materno ter ocorrido na Região de Serrinha em detrimento das Demais Regiões.",
    "Conclui-se que o perfil socioeconômico e demográfico das mortes maternas na Região de Serrinha não difere significativamente do padrão observado no resto do estado. A vulnerabilidade socioeconômica na mortalidade materna é um problema sistêmico e transversal, não apresentando sobre-representação estatística exclusiva no pólo de Serrinha. Recomenda-se o fortalecimento global da rede de atenção primária e pré-natal, sem distinção de perfil baseada unicamente no fator regional analisado."
  )
)

# 9. Exportação para Excel com Múltiplas Abas
caminho_excel <- "Resultados_RP_Com_Laudo.xlsx"

# Passando uma lista nomeada para criar abas separadas no arquivo
write_xlsx(
  list(
    "Resultados_Estatisticos" = tabela_final,
    "Laudo_Epidemiologico" = tabela_laudo
  ), 
  path = caminho_excel
)

cat(sprintf("\nSucesso! O arquivo com múltiplas abas foi exportado para:\n%s/%s\n", getwd(), caminho_excel))
cat("-------------------------\n")

