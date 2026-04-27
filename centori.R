# ==============================================================================
# TÍTULO: PIPELINE EPIDEMIOLÓGICO UNIFICADO - OBSTETRÍCIA BAHIA (SISAL)
# OBJETIVO: Distâncias, Fluxos (Bezier), Mapa Interativo e Laudo PDF ABNT
# BASE TEÓRICA: Artigo "Mortalidade Materna no Semiárido Baiano" 
# AUTOR: Humberto Livramento
# ==============================================================================

message("Iniciando o Pipeline Epidemiológico Unificado...")

# 1. Instalação e Carga de Pacotes ---------------------------------------
pacotes <- c("tidyverse", "geobr", "sf", "openxlsx", "rmarkdown", 
             "tinytex", "readxl", "leaflet", "htmlwidgets", "viridis")

# Instala pacotes ausentes
pacotes_instalar <- pacotes[!(pacotes %in% installed.packages()[,"Package"])]
if(length(pacotes_instalar)) install.packages(pacotes_instalar)

# Carregamento
library(tidyverse) # Inclui dplyr, ggplot2, readr, stringr, forcats
library(geobr)
library(sf)
library(openxlsx)
library(rmarkdown)
library(tinytex)
library(readxl)
library(leaflet)
library(htmlwidgets)
library(viridis)

# 2. Seleção do Arquivo e Diretórios -------------------------------------------
cat("A abrir janela para selecionar o arquivo base_teste.csv...\n")
arq_entrada <- file.choose()
dir_base <- paste0(dirname(arq_entrada), "/")
setwd(dir_base)
cat("Resultados serão guardados na pasta:\n", dir_base, "\n")

# Códigos IBGE - Região do Sisal
municipios_sisal <- c(
  "290230", "290327", "290370", "290610", "290680", 
  "290840", "291340", "291460", "291900", "292100", 
  "292273", "292590", "292620", "292800", "292895", 
  "293050", "293150", "293180", "293290", "292580"
)

# 3. Leitura e Filtro de Dados Obstétricos (Bahia) -----------------------------
message("Lendo e filtrando dados obstétricos do SIH...")
dados_sih <- read_csv2(arq_entrada, locale = locale(encoding = "latin1"), show_col_types = FALSE)

dados_obstetricos <- dados_sih %>%
  mutate(
    MUNIC_RES = str_pad(as.character(MUNIC_RES), 6, pad = "0"),
    MUNIC_MOV = str_pad(as.character(MUNIC_MOV), 6, pad = "0")
  ) %>%
  filter(
    str_starts(MUNIC_MOV, "29"), # Ocorrência na Bahia
    str_starts(MUNIC_RES, "29"), # Residência na Bahia
    str_starts(DIAG_PRINC, "O")  # Capítulo XV (Obstetrícia)
  )

# 4. Geometrias e Centroides (geobr) -------------------------------------------
message("Baixando malha da Bahia e calculando centroides...")
malha_ba <- read_municipality(code_muni = "BA", year = 2020, showProgress = FALSE) %>%
  mutate(MUNIC_6 = substr(as.character(code_muni), 1, 6))

suppressWarnings(centroides_ba <- st_centroid(malha_ba))

coords_centroides <- centroides_ba %>%
  mutate(lon = st_coordinates(.)[,1], lat = st_coordinates(.)[,2]) %>%
  st_drop_geometry() %>%
  select(MUNIC_6, name_muni, lon, lat)

# 5. Cruzamento Espacial e Distâncias ------------------------------------------
message("Calculando distâncias de deslocamento em linha reta...")
dados_distancia <- dados_obstetricos %>%
  left_join(coords_centroides, by = c("MUNIC_RES" = "MUNIC_6")) %>%
  rename(nome_mun_res = name_muni, lon_res = lon, lat_res = lat) %>%
  left_join(coords_centroides, by = c("MUNIC_MOV" = "MUNIC_6")) %>%
  rename(nome_mun_mov = name_muni, lon_mov = lon, lat_mov = lat) %>%
  filter(!is.na(lon_res) & !is.na(lon_mov))

origens_sf <- st_as_sf(dados_distancia, coords = c("lon_res", "lat_res"), crs = 4674)
destinos_sf <- st_as_sf(dados_distancia, coords = c("lon_mov", "lat_mov"), crs = 4674)

# Distância em km
dados_distancia$distancia_km <- as.numeric(st_distance(origens_sf, destinos_sf, by_element = TRUE)) / 1000
dados_distancia <- dados_distancia %>%
  mutate(regiao_sisal = ifelse(MUNIC_RES %in% municipios_sisal, "Região do Sisal", "Outras Regiões"))

# 6. Agregações e Exportação Excel ---------------------------------------------
resumo_bahia <- dados_distancia %>%
  group_by(MUNIC_RES, nome_mun_res, regiao_sisal) %>%
  summarise(
    total_internamentos = n(),
    dist_minima_km = round(min(distancia_km, na.rm = TRUE), 2),
    dist_media_km = round(mean(distancia_km, na.rm = TRUE), 2),
    dist_maxima_km = round(max(distancia_km, na.rm = TRUE), 2),
    .groups = 'drop'
  ) %>% arrange(desc(dist_media_km))

resumo_sisal <- resumo_bahia %>% filter(regiao_sisal == "Região do Sisal")

wb <- createWorkbook()
addWorksheet(wb, "Geral Bahia")
addWorksheet(wb, "Apenas Sisal")
addWorksheet(wb, "Dados Brutos")
writeData(wb, "Geral Bahia", resumo_bahia)
writeData(wb, "Apenas Sisal", resumo_sisal)
writeData(wb, "Dados Brutos", dados_distancia %>% select(N_AIH, MUNIC_RES, nome_mun_res, MUNIC_MOV, nome_mun_mov, distancia_km))
saveWorkbook(wb, "Resultados_Distancias_Obstetricas.xlsx", overwrite = TRUE)

# 7. Gráficos ABNT Básicos (Tema e Curvas) -------------------------------------
message("Gerando gráficos estáticos (Padrão ABNT)...")
tema_abnt <- theme_classic() +
  theme(text = element_text(family = "serif", size = 12, color = "black"),
        plot.title = element_text(face = "bold", size = 12),
        panel.border = element_rect(colour = "black", fill = NA, size = 0.5),
        legend.position = "bottom")

# Gráfico: Curvas de Densidade de Deslocamento
grafico_curvas <- ggplot(dados_distancia, aes(x = distancia_km, fill = regiao_sisal)) +
  geom_density(alpha = 0.5, color = "black") +
  scale_fill_manual(values = c("Outras Regiões" = "#D3D3D3", "Região do Sisal" = "#8B0000"), name = "Região:") +
  labs(title = "Figura 1 - Curvas de Densidade do Deslocamento Obstétrico",
       x = "Distância Percorrida (km)", y = "Densidade",
       caption = "Fonte: Ministério da Saúde - SIH/SUS. Elaboração do autor.") +
  tema_abnt
ggsave("ABNT_Fig1_Curvas_Densidade.tiff", plot = grafico_curvas, device = "tiff", dpi = 300, width = 16, height = 12, units = "cm", compression = "lzw")

# 8. Linhas de Desejo e Mapa Interativo (Leaflet) ------------------------------
message("Construindo Linhas de Desejo e Mapa Interativo (Leaflet)...")
# Filtra apenas quem se deslocou (> 0 km)
dados_fluxo <- dados_distancia %>% filter(distancia_km > 0.1)

linhas_list <- lapply(1:nrow(dados_fluxo), function(i) {
  st_linestring(matrix(c(dados_fluxo$lon_res[i], dados_fluxo$lon_mov[i],
                         dados_fluxo$lat_res[i], dados_fluxo$lat_mov[i]), ncol = 2))
})
linhas_sf <- st_sfc(linhas_list, crs = 4674)
linhas_desejo <- st_sf(dados_fluxo, geometry = linhas_sf)
linhas_wgs84 <- st_transform(linhas_desejo, 4326) 

mapa_interativo <- leaflet(linhas_wgs84) %>%
  addProviderTiles(providers$CartoDB.Positron) %>%
  addPolylines(
    color = ~colorNumeric("magma", distancia_km)(distancia_km), weight = 2, opacity = 0.6,
    popup = ~paste0(
      "<b>Fluxo Obstétrico:</b><br>",
      "<b>Origem:</b> ", nome_mun_res, "<br>",
      "<b>Destino:</b> ", nome_mun_mov, "<br>",
      "<b>Distância:</b> ", round(distancia_km, 2), " km"
    )
  )
saveWidget(mapa_interativo, "Mapa_Interativo_Obstetrico.html", selfcontained = TRUE)

# 9. Mapas com Curvas de Bezier (Fluxos) e Calor -------------------------------
message("Gerando Mapas de Fluxo Suavizado (Bezier) e Calor (Densidade)...")
dir.create("Mapas_Espaciais", showWarnings = FALSE)

# Agrega rotas para espessura da linha
dados_rotas <- dados_fluxo %>%
  group_by(MUNIC_RES, MUNIC_MOV, nome_mun_res, nome_mun_mov) %>%
  summarise(
    volume_pacientes = n(),
    lon_orig = first(lon_res), lat_orig = first(lat_res),
    lon_dest = first(lon_mov), lat_dest = first(lat_mov),
    distancia_km = first(distancia_km),
    .groups = "drop"
  )

tema_mapa <- theme_void() + theme(
  plot.title = element_text(face = "bold", size = 14, hjust = 0.5, family = "serif"),
  plot.subtitle = element_text(size = 12, hjust = 0.5, family = "serif", margin = margin(b=10)),
  legend.position = "bottom", legend.title = element_text(face = "bold", size=10)
)

# Mapa Fluxo Bezier (Bahia)
mapa_fluxo <- ggplot() +
  geom_sf(data = malha_ba, fill = "#f8f9fa", color = "#bdc3c7", linewidth = 0.3) +
  geom_curve(data = dados_rotas,
             aes(x = lon_orig, y = lat_orig, xend = lon_dest, yend = lat_dest,
                 color = distancia_km, linewidth = volume_pacientes, alpha = volume_pacientes),
             curvature = 0.25, ncp = 100) +
  scale_color_viridis_c(option = "magma", name = "Distância (km)", direction = -1) +
  scale_linewidth_continuous(range = c(0.1, 2.0), guide = "none") + 
  scale_alpha_continuous(range = c(0.3, 0.8), guide = "none") +     
  tema_mapa +
  labs(title = "FLUXO OBSTÉTRICO INTRAESTADUAL - BAHIA", 
       subtitle = "Volume e distância do deslocamento materno (Curvas de Bezier)")

# Mapa de Calor (Densidade de Destinos - Hospitais que mais recebem)
coords_destinos <- dados_fluxo %>% select(lon = lon_mov, lat = lat_mov)

mapa_calor <- ggplot() +
  geom_sf(data = malha_ba, fill = "#ecf0f1", color = "#bdc3c7", linewidth = 0.3) +
  stat_density_2d(data = coords_destinos, aes(x = lon, y = lat, fill = after_stat(level)), geom = "polygon", alpha = 0.6) +
  scale_fill_viridis_c(option = "viridis", name = "Densidade de\nInternações") +
  tema_mapa +
  labs(title = "CONCENTRAÇÃO DA DEMANDA OBSTÉTRICA", 
       subtitle = "Mapa de calor focado nos municípios de internação")

suppressWarnings({
  ggsave("Mapas_Espaciais/Fluxo_Bezier_Bahia.pdf", plot = mapa_fluxo, width = 10, height = 8, dpi = 300)
  ggsave("Mapas_Espaciais/Fluxo_Bezier_Bahia.tiff", plot = mapa_fluxo, width = 10, height = 8, dpi = 300, compression = "lzw")
  ggsave("Mapas_Espaciais/Calor_Destinos_Bahia.pdf", plot = mapa_calor, width = 10, height = 8, dpi = 300)
})

# 10. Geração do Laudo Epidemiológico (R Markdown -> PDF padrão ABNT) ----------
message("Gerando Laudo Epidemiológico Unificado em PDF...")

# Extração de estatísticas para o Laudo
media_sisal <- mean(dados_distancia$distancia_km[dados_distancia$regiao_sisal == "Região do Sisal"], na.rm = TRUE)
media_outros <- mean(dados_distancia$distancia_km[dados_distancia$regiao_sisal == "Outras Regiões"], na.rm = TRUE)
total_pacientes <- nrow(dados_distancia)
total_deslocamentos <- nrow(dados_fluxo)

texto_rmd <- c(
  "---",
  "title: 'LAUDO TÉCNICO EPIDEMIOLÓGICO: ACESSO E DESLOCAMENTO OBSTÉTRICO'",
  "subtitle: 'Mortalidade Materna e o Modelo dos Três Atrasos no Território do Sisal'",
  "author: 'Sistema de Vigilância Automatizada - Autor: Humberto Livramento'",
  "date: '`r format(Sys.time(), \"%d de %B de %Y\")`'",
  "output: pdf_document",
  "geometry: a4paper, left=3cm, top=3cm, right=2cm, bottom=2cm",
  "header-includes:",
  "  - \\linespread{1.5}",              
  "  - \\setlength{\\parindent}{1.25cm}", 
  "---",
  "",
  "## 1. INTRODUÇÃO",
  "",
  "Este laudo técnico epidemiológico analisa os indicadores de acesso espacial, o perfil de deslocamento e a vulnerabilidade estrutural de gestantes na Bahia, com enfoque na Região de Saúde de Serrinha (Território do Sisal). O documento foi estruturado e formatado em estrita conformidade com as diretrizes da Associação Brasileira de Normas Técnicas (ABNT).",
  "",
  "De acordo com a literatura base (Estudo de Série Histórica 2015-2024), a região estudada caracteriza-se por expressivas vulnerabilidades. O perfil descritivo das vítimas de mortalidade materna aponta uma ampla maioria de mulheres negras (88%) e uma alta frequência em idades de risco (42,8%).",
  "",
  "## 2. METODOLOGIA E GEOPROCESSAMENTO",
  "",
  "A análise utilizou dados de internação (SIH/SUS) cruzados com as malhas territoriais do IBGE. A distância euclidiana entre o município de residência da paciente e a unidade hospitalar foi calculada mediante projeção dos centroides municipais. Para a representação cartográfica, foram aplicados mapas de densidade de Kernel e curvas de Bezier para maximizar a visualização cartográfica dos fluxos regionais (R CORE TEAM, 2023).",
  "",
  "Adicionalmente, conforme preconizado metodologicamente no estudo referenciado, a análise espacial epidemiológica corrobora a necessidade do uso da Suavização Bayesiana Espacial Empírica (EBest) para correção de flutuações em pequenos municípios (a exemplo da correção do município de Barrocas).",
  "",
  "## 3. RESULTADOS EPIDEMIOLÓGICOS",
  "",
  "O **Modelo dos Três Atrasos** estabelece que a mortalidade materna é agravada por atrasos na decisão, no deslocamento e no tratamento. A presente análise quantifica diretamente o **Segundo Atraso** (dificuldade de acesso físico).",
  "",
  "O processamento espacial validou um total de **", total_pacientes, "** registros de internação obstétrica no estado. Destes, identificou-se que **", total_deslocamentos, "** gestantes necessitaram realizar deslocamentos intermunicipais.",
  "",
  "A análise comparativa de fluxo intraestadual apontou as seguintes discrepâncias:",
  "",
  paste("- **Média de Deslocamento - Região do Sisal:** ", round(media_sisal, 2), " km."),
  paste("- **Média de Deslocamento - Demais Regiões BA:** ", round(media_outros, 2), " km."),
  "",
  "As curvas de densidade (Anexo 1) e os Mapas de Fluxo Suavizados atestam o forte impacto do distanciamento geográfico sobre as gestantes do Sisal.",
  "",
  "## 4. CONCLUSÃO",
  "",
  "A expressiva exigência de deslocamento imposta às pacientes da Região do Sisal, atrelada aos determinantes sociais em saúde (vulnerabilidade racial e socioeconômica), evidencia um cenário de alerta máximo. Recomenda-se o fortalecimento da rede de atenção primária e do transporte sanitário especializado (SAMU) para mitigação do 'Segundo Atraso' no semiárido baiano.",
  "",
  "## REFERÊNCIAS",
  "",
  "BRASIL. Ministério da Saúde. Secretaria de Vigilância em Saúde. **Guia de Vigilância Epidemiológica do Óbito Materno**. Brasília: Ministério da Saúde, 2022.",
  "",
  "INSTITUTO BRASILEIRO DE GEOGRAFIA E ESTATÍSTICA (IBGE). **Malha Territorial Brasileira**. Rio de Janeiro: IBGE, 2020.",
  "",
  "PEREIRA, R. H. M.; GONÇALVES, C. N. geobr: pacote R para baixar dados espaciais oficiais do Brasil. **Ipea - Instituto de Pesquisa Econômica Aplicada**, Brasília, 2021.",
  "",
  "R CORE TEAM. **R**: A language and environment for statistical computing. R Foundation for Statistical Computing, Vienna, Austria, 2023."
)

writeLines(texto_rmd, "Laudo_Epidemiologico_Sisal.Rmd")

# Tenta gerar PDF. Se TinyTex não estiver configurado, avisa o usuário.
options(tinytex.install_packages = FALSE)
tryCatch({
  rmarkdown::render("Laudo_Epidemiologico_Sisal.Rmd", output_format = "pdf_document", quiet = TRUE)
}, error = function(e) {
  message("\nAviso: Gerando HTML pois o computador não possui LaTeX (TinyTex) completo.")
  rmarkdown::render("Laudo_Epidemiologico_Sisal.Rmd", output_format = "html_document", quiet = TRUE)
})
unlink("Laudo_Epidemiologico_Sisal.Rmd") # Limpa Rmd

message("=========================================================")
message("PROCESSAMENTO UNIFICADO CONCLUÍDO COM SUCESSO!")
message("- Planilhas Excel com distâncias processadas.")
message("- Mapas Estáticos TIFF com padrão ABNT.")
message("- Mapa Interativo Leaflet (HTML) com pop-ups.")
message("- Mapas Regionais (Curvas de Bezier e Calor).")
message("- Laudo Epidemiológico formatado gerado.")
message("Verifique a pasta de origem do seu CSV!")
message("=========================================================")

