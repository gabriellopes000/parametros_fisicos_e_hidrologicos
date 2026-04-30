# =============================================================================
# SCRIPT: Clip de Uso do Solo por Áreas de Drenagem e Compilação de Áreas
# =============================================================================

library(sf)
library(dplyr)
library(tidyr)
library(tcltk)
library(openxlsx)

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

if (!"Classe" %in% names(uso_solo_raw)) {
  stop("ERRO: O shapefile de uso do solo não contém a coluna 'Classe'. Verifique o arquivo.")
}
cat("Coluna 'Classe' encontrada.\n")
cat("Classes:", paste(unique(uso_solo_raw$Classe), collapse = ", "), "\n\n")

# -----------------------------------------------------------------------------
# 2. DISSOLVE POR CLASSE
# -----------------------------------------------------------------------------
cat("=== PASSO 2: Dissolve por classe ===\n")

classes_duplas <- uso_solo_raw %>%
  st_drop_geometry() %>%
  count(Classe) %>%
  filter(n > 1)

if (nrow(classes_duplas) > 0) {
  cat("Classes com múltiplos polígonos detectadas — realizando dissolve...\n")
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

if (length(caminhos_ad) == 0) stop("Nenhuma Área de Drenagem selecionada. Encerrando.")

cat("Áreas de Drenagem selecionadas:", length(caminhos_ad), "\n")
cat(paste("-", basename(caminhos_ad), collapse = "\n"), "\n\n")

# -----------------------------------------------------------------------------
# 4. SELEÇÃO DA PASTA DE SAÍDA
# -----------------------------------------------------------------------------
cat("=== PASSO 4: Selecione a pasta de saída ===\n")

pasta_saida <- tk_choose.dir(caption = "Selecione a pasta de saída")
if (is.na(pasta_saida) || pasta_saida == "") stop("Nenhuma pasta de saída selecionada. Encerrando.")
cat("Pasta de saída:", pasta_saida, "\n\n")

# -----------------------------------------------------------------------------
# 5. PROCESSAMENTO: CLIP + CÁLCULO DE ÁREAS + SALVAMENTO
# -----------------------------------------------------------------------------
cat("=== PASSO 5: Processando cada Área de Drenagem ===\n\n")

resultados_lista <- list()

for (caminho_ad in caminhos_ad) {
  
  nome_ad <- tools::file_path_sans_ext(basename(caminho_ad))
  cat("Processando:", nome_ad, "...\n")
  
  ad <- st_read(caminho_ad, quiet = TRUE) %>% st_zm(drop = TRUE, what = "ZM")
  
  if (st_crs(ad) != st_crs(uso_solo)) {
    ad <- st_transform(ad, st_crs(uso_solo))
    cat("  CRS reprojetado.\n")
  }
  
  uso_clipado <- tryCatch({
    res <- st_intersection(uso_solo, st_union(ad)) %>%
      st_zm(drop = TRUE, what = "ZM")
    res <- res[st_geometry_type(res) %in% c("POLYGON", "MULTIPOLYGON"), ]
    st_cast(res, "MULTIPOLYGON")
  }, error = function(e) {
    cat("  AVISO: Erro no clip de", nome_ad, "—", conditionMessage(e), "\n")
    return(NULL)
  })
  
  if (is.null(uso_clipado)) next
  
  uso_clipado <- uso_clipado %>%
    mutate(
      Id     = 0,
      AD_m2  = as.numeric(st_area(geometry)),
      AD_ha  = AD_m2 / 10000,
      AD_km2 = AD_m2 / 1e6
    ) %>%
    select(Id, Classe, AD_m2, AD_ha, AD_km2, geometry)
  
  nome_saida <- paste0(nome_ad, "_v2.shp")
  st_write(uso_clipado, file.path(pasta_saida, nome_saida), delete_layer = TRUE, quiet = TRUE)
  cat("  Salvo:", nome_saida, "\n")
  
  resumo <- uso_clipado %>%
    st_drop_geometry() %>%
    select(Classe, AD_m2, AD_ha, AD_km2) %>%
    mutate(AD = nome_ad)
  
  resultados_lista[[nome_ad]] <- resumo
  cat("  Feições clipadas:", nrow(uso_clipado), "\n\n")
}

# -----------------------------------------------------------------------------
# 6. COMPILAÇÃO DOS DADOS
# -----------------------------------------------------------------------------
cat("=== PASSO 6: Compilando tabela final ===\n\n")

if (length(resultados_lista) == 0) stop("Nenhum resultado gerado.")

nomes_ad    <- tools::file_path_sans_ext(basename(caminhos_ad))
dados_todos <- bind_rows(resultados_lista) %>%
  mutate(AD = factor(AD, levels = nomes_ad))

classes_unicas <- sort(unique(as.character(dados_todos$Classe)))
n_ad           <- length(nomes_ad)

# -----------------------------------------------------------------------------
# 7. EXPORTAÇÃO EXCEL FORMATADO
# -----------------------------------------------------------------------------
cat("=== PASSO 7: Gerando Excel formatado ===\n")

wb <- createWorkbook()
addWorksheet(wb, "Uso do Solo por AD")
ws <- "Uso do Solo por AD"

# --- Estilos ------------------------------------------------------------------

st_titulo <- createStyle(
  fontSize = 13, fontColour = "#FFFFFF", fontName = "Arial",
  fgFill = "#1F4E79", halign = "center", valign = "center",
  textDecoration = "bold"
)
st_subtitulo <- createStyle(
  fontSize = 9, fontColour = "#595959", fontName = "Arial",
  halign = "left", valign = "center", textDecoration = "italic"
)
st_cab_ad <- createStyle(
  fontSize = 10, fontColour = "#FFFFFF", fontName = "Arial",
  fgFill = "#2E75B6", halign = "center", valign = "center",
  textDecoration = "bold", wrapText = TRUE,
  border = "TopBottomLeftRight", borderColour = "#FFFFFF"
)
st_cab_sub <- createStyle(
  fontSize = 9, fontColour = "#FFFFFF", fontName = "Arial",
  fgFill = "#5B9BD5", halign = "center", valign = "center",
  textDecoration = "bold",
  border = "TopBottomLeftRight", borderColour = "#FFFFFF"
)
st_cab_classe <- createStyle(
  fontSize = 10, fontColour = "#FFFFFF", fontName = "Arial",
  fgFill = "#1F4E79", halign = "center", valign = "center",
  textDecoration = "bold",
  border = "TopBottomLeftRight", borderColour = "#FFFFFF"
)
st_dado_par <- createStyle(
  fontSize = 9, fontName = "Arial", halign = "right", valign = "center",
  fgFill = "#DEEAF1",
  border = "TopBottomLeftRight", borderColour = "#BDD7EE",
  numFmt = "#,##0.00"
)
st_dado_impar <- createStyle(
  fontSize = 9, fontName = "Arial", halign = "right", valign = "center",
  fgFill = "#FFFFFF",
  border = "TopBottomLeftRight", borderColour = "#BDD7EE",
  numFmt = "#,##0.00"
)
st_classe_par <- createStyle(
  fontSize = 9, fontName = "Arial", halign = "left", valign = "center",
  fgFill = "#DEEAF1", textDecoration = "bold",
  border = "TopBottomLeftRight", borderColour = "#BDD7EE"
)
st_classe_impar <- createStyle(
  fontSize = 9, fontName = "Arial", halign = "left", valign = "center",
  fgFill = "#FFFFFF", textDecoration = "bold",
  border = "TopBottomLeftRight", borderColour = "#BDD7EE"
)
st_total_label <- createStyle(
  fontSize = 10, fontName = "Arial", halign = "left", valign = "center",
  fgFill = "#1F4E79", fontColour = "#FFFFFF", textDecoration = "bold",
  border = "TopBottomLeftRight", borderColour = "#FFFFFF"
)
st_total_val <- createStyle(
  fontSize = 10, fontName = "Arial", halign = "right", valign = "center",
  fgFill = "#2E75B6", fontColour = "#FFFFFF", textDecoration = "bold",
  border = "TopBottomLeftRight", borderColour = "#FFFFFF",
  numFmt = "#,##0.00"
)
st_dado_m2_par <- createStyle(
  fontSize = 9, fontName = "Arial", halign = "right", valign = "center",
  fgFill = "#DEEAF1",
  border = "TopBottomLeftRight", borderColour = "#BDD7EE",
  numFmt = "#,##0.00"
)
st_dado_m2_impar <- createStyle(
  fontSize = 9, fontName = "Arial", halign = "right", valign = "center",
  fgFill = "#FFFFFF",
  border = "TopBottomLeftRight", borderColour = "#BDD7EE",
  numFmt = "#,##0.00"
)
st_total_m2 <- createStyle(
  fontSize = 10, fontName = "Arial", halign = "right", valign = "center",
  fgFill = "#2E75B6", fontColour = "#FFFFFF", textDecoration = "bold",
  border = "TopBottomLeftRight", borderColour = "#FFFFFF",
  numFmt = "#,##0.00"
)

# --- Layout de linhas e colunas -----------------------------------------------
# Linha 1: Título
# Linha 2: Subtítulo
# Linha 3: Cabeçalho AD (mesclado a cada 3 colunas)
# Linha 4: Sub-cabeçalho (m², ha, km²)
# Linhas 5 ... 5+n_classes-1: Dados
# Última linha: Totais

linha_titulo    <- 1
linha_subtitulo <- 2
linha_cab_ad    <- 3
linha_cab_sub   <- 4
linha_dados_ini <- 5
linha_dados_fim <- linha_dados_ini + length(classes_unicas) - 1
linha_total     <- linha_dados_fim + 1

col_classe  <- 1                          # coluna A = Classe
col_ad_ini  <- function(i) 2 + (i - 1) * 3   # cada AD ocupa 3 colunas: m², ha, km²
col_ad_m2   <- function(i) col_ad_ini(i)
col_ad_ha   <- function(i) col_ad_ini(i) + 1
col_ad_km2  <- function(i) col_ad_ini(i) + 2
col_fim     <- col_ad_km2(n_ad)

# --- Título e subtítulo -------------------------------------------------------
mergeCells(wb, ws, rows = linha_titulo, cols = col_classe:col_fim)
writeData(wb, ws, "TABELA DE USO DO SOLO POR ÁREA DE DRENAGEM",
          startRow = linha_titulo, startCol = col_classe)
addStyle(wb, ws, st_titulo, rows = linha_titulo, cols = col_classe:col_fim, gridExpand = TRUE)
setRowHeights(wb, ws, linha_titulo, 22)

mergeCells(wb, ws, rows = linha_subtitulo, cols = col_classe:col_fim)
writeData(wb, ws, "Áreas calculadas a partir de clip vetorial (sf::st_intersection)",
          startRow = linha_subtitulo, startCol = col_classe)
addStyle(wb, ws, st_subtitulo, rows = linha_subtitulo, cols = col_classe:col_fim, gridExpand = TRUE)
setRowHeights(wb, ws, linha_subtitulo, 16)

# --- Cabeçalho: coluna Classe (linhas 3 e 4 mescladas) ----------------------
mergeCells(wb, ws, rows = c(linha_cab_ad, linha_cab_sub), cols = col_classe)
writeData(wb, ws, "Classe", startRow = linha_cab_ad, startCol = col_classe)
addStyle(wb, ws, st_cab_classe,
         rows = c(linha_cab_ad, linha_cab_sub), cols = col_classe, gridExpand = TRUE)

# --- Cabeçalho: uma seção por AD ---------------------------------------------
for (i in seq_along(nomes_ad)) {
  
  c_ini <- col_ad_ini(i)
  c_fim_bloco <- col_ad_km2(i)
  
  # Nome da AD (mesclado sobre as 3 sub-colunas)
  mergeCells(wb, ws, rows = linha_cab_ad, cols = c_ini:c_fim_bloco)
  writeData(wb, ws, nomes_ad[i], startRow = linha_cab_ad, startCol = c_ini)
  addStyle(wb, ws, st_cab_ad,
           rows = linha_cab_ad, cols = c_ini:c_fim_bloco, gridExpand = TRUE)
  
  # Sub-cabeçalhos
  writeData(wb, ws, "Área (m²)", startRow = linha_cab_sub, startCol = col_ad_m2(i))
  writeData(wb, ws, "Área (ha)", startRow = linha_cab_sub, startCol = col_ad_ha(i))
  writeData(wb, ws, "Área (km²)", startRow = linha_cab_sub, startCol = col_ad_km2(i))
  addStyle(wb, ws, st_cab_sub,
           rows = linha_cab_sub,
           cols = c(col_ad_m2(i), col_ad_ha(i), col_ad_km2(i)),
           gridExpand = TRUE)
}

setRowHeights(wb, ws, c(linha_cab_ad, linha_cab_sub), c(28, 20))

# --- Dados por classe --------------------------------------------------------
for (i in seq_along(classes_unicas)) {
  
  cls      <- classes_unicas[i]
  linha_i  <- linha_dados_ini + i - 1
  eh_par   <- (i %% 2 == 0)
  
  st_cls  <- if (eh_par) st_classe_par  else st_classe_impar
  st_val  <- if (eh_par) st_dado_par    else st_dado_impar
  st_m2   <- if (eh_par) st_dado_m2_par else st_dado_m2_impar
  
  # Nome da classe
  writeData(wb, ws, cls, startRow = linha_i, startCol = col_classe)
  addStyle(wb, ws, st_cls, rows = linha_i, cols = col_classe)
  
  # Valores por AD
  for (j in seq_along(nomes_ad)) {
    sub <- dados_todos %>% filter(Classe == cls, AD == nomes_ad[j])
    
    v_m2  <- if (nrow(sub) == 0) 0 else sum(sub$AD_m2)
    v_ha  <- if (nrow(sub) == 0) 0 else sum(sub$AD_ha)
    v_km2 <- if (nrow(sub) == 0) 0 else sum(sub$AD_km2)
    
    writeData(wb, ws, v_m2,  startRow = linha_i, startCol = col_ad_m2(j))
    writeData(wb, ws, v_ha,  startRow = linha_i, startCol = col_ad_ha(j))
    writeData(wb, ws, v_km2, startRow = linha_i, startCol = col_ad_km2(j))
    
    addStyle(wb, ws, st_m2,  rows = linha_i, cols = col_ad_m2(j))
    addStyle(wb, ws, st_val, rows = linha_i, cols = col_ad_ha(j))
    addStyle(wb, ws, st_val, rows = linha_i, cols = col_ad_km2(j))
  }
  
  setRowHeights(wb, ws, linha_i, 16)
}

# --- Linha de totais ----------------------------------------------------------
writeData(wb, ws, "TOTAL", startRow = linha_total, startCol = col_classe)
addStyle(wb, ws, st_total_label, rows = linha_total, cols = col_classe)

for (j in seq_along(nomes_ad)) {
  
  r_ini <- linha_dados_ini
  r_fim <- linha_dados_fim
  
  col_m2  <- col_ad_m2(j)
  col_ha  <- col_ad_ha(j)
  col_km2 <- col_ad_km2(j)
  
  cel_m2  <- paste0(int2col(col_m2),  r_ini, ":", int2col(col_m2),  r_fim)
  cel_ha  <- paste0(int2col(col_ha),  r_ini, ":", int2col(col_ha),  r_fim)
  cel_km2 <- paste0(int2col(col_km2), r_ini, ":", int2col(col_km2), r_fim)
  
  writeFormula(wb, ws, paste0("=SUM(", cel_m2,  ")"), startRow = linha_total, startCol = col_m2)
  writeFormula(wb, ws, paste0("=SUM(", cel_ha,  ")"), startRow = linha_total, startCol = col_ha)
  writeFormula(wb, ws, paste0("=SUM(", cel_km2, ")"), startRow = linha_total, startCol = col_km2)
  
  addStyle(wb, ws, st_total_m2,  rows = linha_total, cols = col_m2)
  addStyle(wb, ws, st_total_val, rows = linha_total, cols = col_ha)
  addStyle(wb, ws, st_total_val, rows = linha_total, cols = col_km2)
}

setRowHeights(wb, ws, linha_total, 20)

# --- Largura das colunas ------------------------------------------------------
setColWidths(wb, ws, cols = col_classe, widths = 30)
for (j in seq_along(nomes_ad)) {
  setColWidths(wb, ws, cols = col_ad_m2(j),  widths = 18)
  setColWidths(wb, ws, cols = col_ad_ha(j),  widths = 13)
  setColWidths(wb, ws, cols = col_ad_km2(j), widths = 13)
}

# --- Congelar painéis na linha de dados e coluna Classe ----------------------
freezePane(wb, ws, firstActiveRow = linha_dados_ini, firstActiveCol = 2)

# --- Salvar ------------------------------------------------------------------
caminho_xlsx <- file.path(pasta_saida, "tabela_uso_solo_por_AD.xlsx")
saveWorkbook(wb, caminho_xlsx, overwrite = TRUE)

cat("Excel salvo em:\n ", caminho_xlsx, "\n")
cat("\n=== PROCESSAMENTO CONCLUÍDO ===\n")
