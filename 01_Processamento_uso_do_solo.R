# =============================================================================
# SCRIPT: Dissolve e Recálculo de Áreas por Classe de Uso do Solo
# =============================================================================
# DESCRIÇÃO:
#   Este script realiza o processamento de shapefiles de uso do solo com as
#   seguintes etapas:
#     1. Abre uma caixa de diálogo para seleção do shapefile de entrada
#     2. Verifica se existem polígonos com a mesma classe (campo "Classe")
#     3. Caso existam, realiza o dissolve (mesclagem) dos polígonos por classe
#     4. Recalcula as áreas em m² e km² com base nas geometrias resultantes
#     5. Salva o shapefile processado na mesma pasta de origem com sufixo "_v2"
#
# REQUISITOS:
#   - Pacotes: sf, dplyr, tcltk
#   - O shapefile de entrada deve conter as colunas:
#     Id | AD_m2 | AD_km2 | Classe | geometry
#
# AUTOR: [Seu nome]
# DATA:  [Data]
# =============================================================================

library(sf)
library(dplyr)
library(tcltk)

# -----------------------------------------------------------------------------
# 1. SELEÇÃO DO ARQUIVO VIA CAIXA DE DIÁLOGO
# -----------------------------------------------------------------------------
cat("Aguardando seleção do shapefile...\n")

caminho_shp <- tk_choose.files(
  caption = "Selecione o shapefile de Uso do Solo",
  filter  = matrix(c("Shapefiles", "*.shp"), ncol = 2)
)

if (length(caminho_shp) == 0 || caminho_shp == "") {
  stop("Nenhum arquivo selecionado. Encerrando o script.")
}

cat("Arquivo selecionado:\n", caminho_shp, "\n\n")

# -----------------------------------------------------------------------------
# 2. LEITURA DO SHAPEFILE
# -----------------------------------------------------------------------------
dados <- st_read(caminho_shp, quiet = TRUE)

cat("Shapefile importado com sucesso!\n")
cat("Feições originais:", nrow(dados), "\n")
cat("Classes encontradas:", paste(unique(dados$Classe), collapse = ", "), "\n\n")

# -----------------------------------------------------------------------------
# 3. VERIFICAÇÃO E DISSOLVE POR CLASSE
# -----------------------------------------------------------------------------
classes_duplicadas <- dados %>%
  st_drop_geometry() %>%
  count(Classe) %>%
  filter(n > 1)

if (nrow(classes_duplicadas) > 0) {
  cat("Classes com múltiplos polígonos detectadas:\n")
  print(classes_duplicadas)
  cat("\nRealizando dissolve por classe...\n")
  
  dados_proc <- dados %>%
    group_by(Classe) %>%
    summarise(geometry = st_union(geometry), .groups = "drop")
  
  cat("Dissolve concluído. Feições após dissolve:", nrow(dados_proc), "\n\n")
  
} else {
  cat("Nenhuma classe duplicada encontrada. Dissolve não necessário.\n\n")
  dados_proc <- dados %>% select(Classe, geometry)
}

# -----------------------------------------------------------------------------
# 4. RECÁLCULO DAS ÁREAS
# -----------------------------------------------------------------------------
cat("Recalculando áreas...\n")

dados_proc <- dados_proc %>%
  mutate(
    Id     = 0,
    AD_m2  = as.numeric(st_area(geometry)),
    AD_km2 = AD_m2 / 1e6
  ) %>%
  select(Id, AD_m2, AD_km2, Classe, geometry)

cat("Áreas recalculadas com sucesso!\n\n")
print(st_drop_geometry(dados_proc))

# -----------------------------------------------------------------------------
# 5. SALVAMENTO DO SHAPEFILE COM SUFIXO _v2
# -----------------------------------------------------------------------------
pasta_origem  <- dirname(caminho_shp)
nome_original <- tools::file_path_sans_ext(basename(caminho_shp))
nome_saida    <- paste0(nome_original, "_v2.shp")
caminho_saida <- file.path(pasta_origem, nome_saida)

# Forçar geometria 2D antes de salvar
dados_proc <- st_zm(dados_proc, drop = TRUE, what = "ZM")

st_write(dados_proc, caminho_saida, delete_layer = TRUE, quiet = TRUE)

cat("Shapefile salvo com sucesso em:\n", caminho_saida, "\n")