# =============================================================================
# SCRIPT: Compilação de Áreas por Classe de Uso do Solo - Múltiplos Shapefiles
# =============================================================================
# DESCRIÇÃO:
#   Este script lê todos os shapefiles de uma pasta, extrai as áreas por classe
#   de uso do solo e organiza os resultados em uma tabela comparativa onde:
#     - Linhas  = Classes de uso do solo
#     - Colunas = Nome de cada shapefile
#     - Valores = Área em km² (ou m², configurável)
#
# REQUISITOS:
#   - Pacotes: sf, dplyr, tidyr
#   - Os shapefiles devem conter as colunas: Id | AD_m2 | AD_km2 | Classe
# =============================================================================

library(sf)
library(dplyr)
library(tidyr)

# -----------------------------------------------------------------------------
# 1. DEFINIR PASTA E LISTAR SHAPEFILES
# -----------------------------------------------------------------------------
pasta <- "C:/Users/gabriel.paula/Desktop/Projetos/13_Dique_2/15_Parametros_Fisicos_e_Hidrologicos/02_Uso_do_Solo/Atualizado"

arquivos_shp <- list.files(pasta, pattern = "\\.shp$", full.names = TRUE)

if (length(arquivos_shp) == 0) stop("Nenhum shapefile encontrado na pasta.")

cat("Shapefiles encontrados:", length(arquivos_shp), "\n")
cat(paste("-", basename(arquivos_shp), collapse = "\n"), "\n\n")

# -----------------------------------------------------------------------------
# 2. LER TODOS OS SHAPEFILES E EMPILHAR
# -----------------------------------------------------------------------------
lista <- lapply(arquivos_shp, function(arq) {
  shp <- st_read(arq, quiet = TRUE)
  shp %>%
    st_drop_geometry() %>%
    select(Classe, AD_km2) %>%
    mutate(Shapefile = tools::file_path_sans_ext(basename(arq)))
})

dados_todos <- bind_rows(lista)

# -----------------------------------------------------------------------------
# 3. PIVOTAR: linhas = Classe, colunas = Shapefile
# -----------------------------------------------------------------------------
tabela <- dados_todos %>%
  pivot_wider(
    names_from  = Shapefile,
    values_from = AD_km2,
    values_fill = 0          # classes ausentes em algum shapefile recebem 0
  ) %>%
  arrange(Classe)

# Adicionar coluna de total por classe (soma das linhas)
tabela <- tabela %>%
  mutate(TOTAL = rowSums(select(., -Classe)))

cat("Tabela gerada com sucesso!\n\n")
print(tabela)

# -----------------------------------------------------------------------------
# 4. SALVAR COMO CSV NA MESMA PASTA
# -----------------------------------------------------------------------------
caminho_csv <- file.path(pasta, "tabela_uso_solo_compilada.csv")
write.csv(tabela, caminho_csv, row.names = FALSE)

cat("\nTabela salva em:\n", caminho_csv, "\n")