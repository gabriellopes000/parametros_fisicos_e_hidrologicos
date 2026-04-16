# =============================================================================
# SCRIPT: Clip de Uso do Solo por Áreas de Drenagem e Compilação de Áreas
# =============================================================================
# DESCRIÇÃO:
#   Este script integra o processamento de uso do solo com múltiplas áreas de
#   drenagem, executando as seguintes etapas:
#
#   1. Seleciona o shapefile de Uso do Solo geral via caixa de diálogo
#      - Verifica a presença obrigatória da coluna "Classe"
#      - Realiza dissolve por classe (mescla polígonos da mesma classe)
#   2. Seleciona múltiplos shapefiles de Áreas de Drenagem via caixa de diálogo
#   3. Seleciona a pasta de saída via caixa de diálogo
#   4. Para cada Área de Drenagem:
#      - Clipa o uso do solo pela extensão da área de drenagem
#      - Recalcula as áreas em m², ha e km²
#      - Salva o shapefile resultante na pasta de saída com sufixo _v2
#   5. Compila uma tabela única com:
#      - Linhas  = Classes de uso do solo
#      - Colunas = Área de Drenagem (com valores em m², ha e km²)
#      - Salva a tabela como CSV na pasta de saída
#
# REQUISITOS:
#   - Pacotes: sf, dplyr, tidyr, tcltk
#   - Shapefile de uso do solo deve conter a coluna "Classe"
#
# AUTOR: [Seu nome]
# DATA:  [Data]
# =============================================================================

library(sf)
library(dplyr)
library(tidyr)
library(tcltk)

# -----------------------------------------------------------------------------
# 1. SELEÇÃO DO SHAPEFILE DE USO DO SOLO GERAL
# -----------------------------------------------------------------------------
cat("=== PASSO 1: Selecione o shapefile de Uso do Solo geral ===\n")

caminho_uso_solo <- tk_choose.files(
  caption = "Selecione o shapefile de Uso do Solo geral",
  filter  = matrix(c("Shapefiles", "*.shp"), ncol = 2),
  multi   = FALSE
)

if (length(caminho_uso_solo) == 0 || caminho_uso_solo == "") {
  stop("Nenhum shapefile de uso do solo selecionado. Encerrando.")
}

uso_solo_raw <- st_read(caminho_uso_solo, quiet = TRUE)
cat("Uso do solo importado:", nrow(uso_solo_raw), "feições\n")

# Verificação obrigatória da coluna "Classe"
if (!"Classe" %in% names(uso_solo_raw)) {
  stop("ERRO: O shapefile de uso do solo não contém a coluna 'Classe'. Verifique o arquivo.")
}
cat("Coluna 'Classe' encontrada.\n")
cat("Classes:", paste(unique(uso_solo_raw$Classe), collapse = ", "), "\n\n")

# -----------------------------------------------------------------------------
# 2. DISSOLVE POR CLASSE NO USO DO SOLO GERAL
# -----------------------------------------------------------------------------
cat("=== PASSO 2: Dissolve por classe ===\n")

classes_duplas <- uso_solo_raw %>%
  st_drop_geometry() %>%
  count(Classe) %>%
  filter(n > 1)

if (nrow(classes_duplas) > 0) {
  cat("Classes com múltiplos polígonos detectadas — realizando dissolve...\n")
  print(classes_duplas)
  
  uso_solo <- uso_solo_raw %>%
    group_by(Classe) %>%
    summarise(geometry = st_union(geometry), .groups = "drop") %>%
    st_zm(drop = TRUE, what = "ZM")
  
  cat("Dissolve concluído:", nrow(uso_solo), "feições resultantes\n\n")
} else {
  cat("Nenhuma classe duplicada — dissolve não necessário.\n\n")
  uso_solo <- uso_solo_raw %>%
    select(Classe, geometry) %>%
    st_zm(drop = TRUE, what = "ZM")
}

# -----------------------------------------------------------------------------
# 3. SELEÇÃO DOS SHAPEFILES DE ÁREAS DE DRENAGEM
# -----------------------------------------------------------------------------
cat("=== PASSO 3: Selecione os shapefiles de Áreas de Drenagem ===\n")

caminhos_ad <- tk_choose.files(
  caption = "Selecione os shapefiles de Áreas de Drenagem (múltipla seleção)",
  filter  = matrix(c("Shapefiles", "*.shp"), ncol = 2),
  multi   = TRUE
)

if (length(caminhos_ad) == 0) {
  stop("Nenhuma Área de Drenagem selecionada. Encerrando.")
}

cat("Áreas de Drenagem selecionadas:", length(caminhos_ad), "\n")
cat(paste("-", basename(caminhos_ad), collapse = "\n"), "\n\n")

# -----------------------------------------------------------------------------
# 4. SELEÇÃO DA PASTA DE SAÍDA
# -----------------------------------------------------------------------------
cat("=== PASSO 4: Selecione a pasta de saída ===\n")

pasta_saida <- tk_choose.dir(caption = "Selecione a pasta de saída")

if (is.na(pasta_saida) || pasta_saida == "") {
  stop("Nenhuma pasta de saída selecionada. Encerrando.")
}

cat("Pasta de saída:", pasta_saida, "\n\n")

# -----------------------------------------------------------------------------
# 5. PROCESSAMENTO: CLIP + CÁLCULO DE ÁREAS + SALVAMENTO
# -----------------------------------------------------------------------------
cat("=== PASSO 5: Processando cada Área de Drenagem ===\n\n")

# Garantir mesmo CRS entre uso do solo e áreas de drenagem
resultados_lista <- list()

for (caminho_ad in caminhos_ad) {
  
  nome_ad <- tools::file_path_sans_ext(basename(caminho_ad))
  cat("Processando:", nome_ad, "...\n")
  
  # Leitura e padronização da AD
  ad <- st_read(caminho_ad, quiet = TRUE) %>%
    st_zm(drop = TRUE, what = "ZM")
  
  # Reprojetar AD para o CRS do uso do solo, se necessário
  if (st_crs(ad) != st_crs(uso_solo)) {
    ad <- st_transform(ad, st_crs(uso_solo))
    cat("  CRS reprojetado para coincidir com o uso do solo.\n")
  }
  
  # Clip do uso do solo pela área de drenagem
  uso_clipado <- tryCatch({
    resultado <- st_intersection(uso_solo, st_union(ad))
    
    # Forçar 2D
    resultado <- st_zm(resultado, drop = TRUE, what = "ZM")
    
    # Manter apenas geometrias do tipo polígono (remover linhas e pontos residuais)
    resultado <- resultado[st_geometry_type(resultado) %in% c("POLYGON", "MULTIPOLYGON"), ]
    
    # Garantir que tudo seja MULTIPOLYGON (tipo uniforme)
    resultado <- st_cast(resultado, "MULTIPOLYGON")
    
    resultado
  }, error = function(e) {
    cat("  AVISO: Erro no clip de", nome_ad, "—", conditionMessage(e), "\n")
    return(NULL)
  })
  
  # Recalcular áreas
  uso_clipado <- uso_clipado %>%
    mutate(
      Id     = 0,
      AD_m2  = as.numeric(st_area(geometry)),
      AD_ha  = AD_m2 / 10000,
      AD_km2 = AD_m2 / 1e6
    ) %>%
    select(Id, Classe, AD_m2, AD_ha, AD_km2, geometry)
  
  # Salvar shapefile clipado
  nome_saida    <- paste0(nome_ad, "_v2.shp")
  caminho_saida <- file.path(pasta_saida, nome_saida)
  st_write(uso_clipado, caminho_saida, delete_layer = TRUE, quiet = TRUE)
  cat("  Salvo:", nome_saida, "\n")
  
  # Guardar resultados para a tabela
  resumo <- uso_clipado %>%
    st_drop_geometry() %>%
    select(Classe, AD_m2, AD_ha, AD_km2) %>%
    mutate(AD = nome_ad)
  
  resultados_lista[[nome_ad]] <- resumo
  cat("  Feições clipadas:", nrow(uso_clipado), "\n\n")
}

# -----------------------------------------------------------------------------
# 6. COMPILAÇÃO DA TABELA FINAL
# -----------------------------------------------------------------------------
cat("=== PASSO 6: Compilando tabela final ===\n")

if (length(resultados_lista) == 0) {
  stop("Nenhum resultado gerado. Verifique os arquivos de entrada.")
}

dados_todos <- bind_rows(resultados_lista)

# Pivotar: uma coluna por AD × unidade
tabela <- dados_todos %>%
  pivot_wider(
    names_from  = AD,
    values_from = c(AD_m2, AD_ha, AD_km2),
    values_fill = 0,
    names_glue  = "{AD}__{.value}"   # ex: AD_Incremental_1__AD_m2
  ) %>%
  arrange(Classe)

# Reordenar colunas: agrupar por AD (m², ha, km² juntos)
nomes_ad    <- tools::file_path_sans_ext(basename(caminhos_ad))
cols_ordem  <- c("Classe", unlist(lapply(nomes_ad, function(ad) {
  c(paste0(ad, "__AD_m2"),
    paste0(ad, "__AD_ha"),
    paste0(ad, "__AD_km2"))
})))
cols_ordem  <- cols_ordem[cols_ordem %in% names(tabela)]
tabela      <- tabela %>% select(all_of(cols_ordem))

cat("Tabela compilada com sucesso!\n\n")
print(tabela)

# -----------------------------------------------------------------------------
# 7. SALVAR TABELA CSV
# -----------------------------------------------------------------------------
caminho_csv <- file.path(pasta_saida, "tabela_uso_solo_por_AD.csv")
write.csv(tabela, caminho_csv, row.names = FALSE, fileEncoding = "UTF-8")

cat("\nTabela salva em:\n", caminho_csv, "\n")
cat("\n=== PROCESSAMENTO CONCLUÍDO ===\n")