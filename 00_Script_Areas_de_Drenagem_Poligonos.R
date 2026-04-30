# ============================================================
# SCRIPT: Cálculo de Áreas de Shapefiles de Polígonos
# ------------------------------------------------------------
# Este script realiza as seguintes etapas:
#
# 1. Abre uma caixa de diálogo para seleção da pasta de trabalho
# 2. Lista e lê todos os arquivos shapefile (.shp) da pasta
# 3. Carrega os shapefiles no ambiente do R
# 4. Verifica se há feições com múltiplas partes (multipart)
#    e as converte para single-part automaticamente
# 5. Reprojeta o CRS para SIRGAS 2000 Policônica (EPSG:5880)
#    caso o shapefile esteja em coordenadas geográficas (graus),
#    garantindo que as áreas sejam calculadas em metros
# 6. Calcula as áreas de cada polígono em m², km² e ha,
#    adicionando esses valores como novos campos nos shapefiles
# 7. Sobrescreve os shapefiles originais com os novos campos
# 8. Gera uma planilha Excel (areas_poligonos.xlsx) na pasta
#    de origem com o nome de cada shapefile e suas áreas,
#    sobrescrevendo a planilha caso ela já exista
# ============================================================

library(sf)
library(dplyr)
library(openxlsx)
library(tcltk)

# ============================================================
# 1. Selecionar pasta via caixa de diálogo
# ============================================================
pasta <- tk_choose.dir(caption = "Selecione a pasta com os shapefiles")

if (is.na(pasta) || pasta == "") stop("Nenhuma pasta selecionada.")

# ============================================================
# 2. Listar shapefiles na pasta
# ============================================================
arquivos_shp <- list.files(pasta, pattern = "\\.shp$", full.names = TRUE, ignore.case = TRUE)

if (length(arquivos_shp) == 0) stop("Nenhum shapefile encontrado na pasta selecionada.")

cat("Shapefiles encontrados:\n")
cat(paste0(" - ", basename(arquivos_shp), "\n"))

# ============================================================
# 3, 4, 5 e 6. Para cada shapefile:
#   - Ler o arquivo
#   - Verificar e corrigir feições multipart
#   - Reprojetar se necessário
#   - Calcular e adicionar campos de área
# ============================================================
lista_sf    <- list()   # objetos sf no ambiente
tabela_area <- list()   # dados para a planilha

for (arq in arquivos_shp) {
  
  nome <- tools::file_path_sans_ext(basename(arq))
  cat(sprintf("\nProcessando: %s\n", nome))
  
  # Leitura do shapefile
  camada <- st_read(arq, quiet = TRUE)
  
  # ── 5. Reprojetar para sistema métrico se necessário ──────
  # Shapefiles em graus (lat/lon) não permitem cálculo de área
  # em metros. Nesse caso, reprojeta para SIRGAS 2000 Policônica.
  if (st_is_longlat(camada)) {
    camada <- st_transform(camada, crs = 5880)
    cat("  → CRS geográfico detectado. Reprojetado para SIRGAS 2000 Policônica (EPSG:5880).\n")
  }
  
  # ── 4. Verificar e converter feições multipart ────────────
  # Feições do tipo MULTIPOLYGON possuem mais de uma