# =============================================================================
# SCRIPT: Clip de Uso do Solo por Áreas de Drenagem e Compilação de Áreas
# =============================================================================

# -----------------------------------------------------------------------------
# 0. PARÂMETROS DE CN / COEFICIENTE DE RUNOFF
# -----------------------------------------------------------------------------

# Escolha o método: "CN" ou "C"
metodo_runoff <- "C"

# Valores de CN por tipologia
valores_CN <- c(
  "Área industrial"            = 85,
  "Cava"                       = 90,
  "Enrocamento"                = 75,
  "Solo Exposto"               = 91,
  "Vegetação de Baixo Porte"   = 74,
  "Vegetação de Grande Porte"  = 55,
  "Vegetação de Médio Porte"   = 65
)

# Valores de C (coeficiente de Runoff) por tipologia
valores_C <- c(
  "Área industrial"            = 0.38,
  "Cava"                       = 0.65,
  "Enrocamento"                = 0.40,
  "Solo Exposto"               = 0.53,
  "Vegetação de Baixo Porte"   = 0.40,
  "Vegetação de Grande Porte"  = 0.13,
  "Vegetação de Médio Porte"   = 0.38
)

# Seleciona a lista ativa com base no método escolhido
valores_runoff <- switch(metodo_runoff,
                         "CN" = valores_CN,
                         "C"  = valores_C,
                         stop("ERRO: 'metodo_runoff' deve ser \"CN\" ou \"C\".")
)

# -----------------------------------------------------------------------------
# 1. INSTALAÇÃO DOS PACOTES
# -----------------------------------------------------------------------------

library(sf)
library(dplyr)
library(tidyr)
library(tcltk)
library(openxlsx)

# -----------------------------------------------------------------------------
# 2. SELEÇÃO DO SHAPEFILE DE USO DO SOLO GERAL
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
# 3. DISSOLVE POR CLASSE
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
# 4. SELEÇÃO DOS SHAPEFILES DE ÁREAS DE DRENAGEM
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
# 5. SELEÇÃO DA PASTA DE SAÍDA
# -----------------------------------------------------------------------------
cat("=== PASSO 4: Selecione a pasta de saída ===\n")

pasta_saida <- tk_choose.dir(caption = "Selecione a pasta de saída")
if (is.na(pasta_saida) || pasta_saida == "") stop("Nenhuma pasta de saída selecionada. Encerrando.")
cat("Pasta de saída:", pasta_saida, "\n\n")

# -----------------------------------------------------------------------------
# 6. PROCESSAMENTO: CLIP + CÁLCULO DE ÁREAS + SALVAMENTO
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
    res <- st_collection_extract(res, type = "POLYGON")  # <-- correção
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
# 7. COMPILAÇÃO DOS DADOS
# -----------------------------------------------------------------------------
cat("=== PASSO 6: Compilando tabela final ===\n\n")

if (length(resultados_lista) == 0) stop("Nenhum resultado gerado.")

nomes_ad    <- tools::file_path_sans_ext(basename(caminhos_ad))
dados_todos <- bind_rows(resultados_lista) %>%
  mutate(AD = factor(AD, levels = nomes_ad))

classes_unicas <- sort(unique(as.character(dados_todos$Classe)))
n_ad           <- length(nomes_ad)

lista_por_ad <- setNames(
  lapply(nomes_ad, function(ad) {
    df <- dados_todos %>%
      filter(AD == ad) %>%
      group_by(Classe) %>%
      summarise(
        AD_m2  = sum(AD_m2,  na.rm = TRUE),
        AD_ha  = sum(AD_ha,  na.rm = TRUE),
        AD_km2 = sum(AD_km2, na.rm = TRUE),
        .groups = "drop"
      )
    
    # Garante todas as classes, preenchendo ausentes com zero
    classes_faltantes <- setdiff(classes_unicas, df$Classe)
    
    if (length(classes_faltantes) > 0) {
      df_zeros <- data.frame(
        Classe = classes_faltantes,
        AD_m2  = 0,
        AD_ha  = 0,
        AD_km2 = 0,
        stringsAsFactors = FALSE
      )
      df <- bind_rows(df, df_zeros)
    }
    
    df %>% arrange(match(Classe, classes_unicas))
  }),
  nomes_ad
)

teste <- as.data.frame(lista_por_ad[["AD_TR-01_v2"]])

# -----------------------------------------------------------------------------
# PASSO 8: Ponderação de CN / C por AD
# -----------------------------------------------------------------------------
cat("=== PASSO 8: Ponderando", metodo_runoff, "por Área de Drenagem ===\n\n")

# Verificar se todas as tipologias presentes têm valor definido
classes_sem_valor <- setdiff(classes_unicas, names(valores_runoff))

if (length(classes_sem_valor) > 0) {
  stop(
    "ERRO: As seguintes tipologias não possuem valor de ", metodo_runoff,
    " definido em 'valores_runoff':\n",
    paste(" -", classes_sem_valor, collapse = "\n")
  )
}

# Ponderar por AD
runoff_por_ad <- setNames(
  lapply(nomes_ad, function(ad) {
    
    df <- lista_por_ad[[ad]] %>%
      mutate(
        Valor_runoff   = valores_runoff[Classe],
        Area_total_m2  = sum(AD_m2),
        Perc_area      = ifelse(Area_total_m2 > 0, AD_m2 / Area_total_m2, 0),
        Contribuicao   = Valor_runoff * Perc_area
      )
    
    valor_ponderado <- sum(df$Contribuicao, na.rm = TRUE)
    
    cat(sprintf("  %-30s %s ponderado = %.2f\n", ad, metodo_runoff, valor_ponderado))
    
    list(
      tabela          = df,
      valor_ponderado = valor_ponderado
    )
  }),
  nomes_ad
)
cat("\n")

teste <- as.data.frame(runoff_por_ad[["AD_TR-01_v2"]])

# -----------------------------------------------------------------------------
# 9. EXPORTAÇÃO EXCEL FORMATADO — UMA ABA POR AD
# -----------------------------------------------------------------------------
cat("=== PASSO 7: Gerando Excel formatado ===\n")

wb <- createWorkbook()

# --- Estilos comuns -----------------------------------------------------------
st_titulo <- createStyle(
  fontSize = 13, fontColour = "#FFFFFF", fontName = "Arial",
  fgFill = "#1F4E79", halign = "center", valign = "center",
  textDecoration = "bold"
)
st_subtitulo <- createStyle(
  fontSize = 9, fontColour = "#595959", fontName = "Arial",
  halign = "left", valign = "center", textDecoration = "italic"
)
st_cab <- createStyle(
  fontSize = 10, fontColour = "#FFFFFF", fontName = "Arial",
  fgFill = "#2E75B6", halign = "center", valign = "center",
  textDecoration = "bold", wrapText = TRUE,
  border = "TopBottomLeftRight", borderColour = "#FFFFFF"
)
st_cab_classe <- createStyle(
  fontSize = 10, fontColour = "#FFFFFF", fontName = "Arial",
  fgFill = "#1F4E79", halign = "center", valign = "center",
  textDecoration = "bold",
  border = "TopBottomLeftRight", borderColour = "#FFFFFF"
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
st_num_par <- createStyle(
  fontSize = 9, fontName = "Arial", halign = "right", valign = "center",
  fgFill = "#DEEAF1",
  border = "TopBottomLeftRight", borderColour = "#BDD7EE",
  numFmt = "#,##0.00"
)
st_num_impar <- createStyle(
  fontSize = 9, fontName = "Arial", halign = "right", valign = "center",
  fgFill = "#FFFFFF",
  border = "TopBottomLeftRight", borderColour = "#BDD7EE",
  numFmt = "#,##0.00"
)
st_pct_par <- createStyle(
  fontSize = 9, fontName = "Arial", halign = "right", valign = "center",
  fgFill = "#DEEAF1",
  border = "TopBottomLeftRight", borderColour = "#BDD7EE",
  numFmt = "0.00%"
)
st_pct_impar <- createStyle(
  fontSize = 9, fontName = "Arial", halign = "right", valign = "center",
  fgFill = "#FFFFFF",
  border = "TopBottomLeftRight", borderColour = "#BDD7EE",
  numFmt = "0.00%"
)
st_runoff_par <- createStyle(
  fontSize = 9, fontName = "Arial", halign = "right", valign = "center",
  fgFill = "#FFF2CC", fontColour = "#7F6000",
  border = "TopBottomLeftRight", borderColour = "#BDD7EE",
  numFmt = "#,##0.00"
)
st_runoff_impar <- createStyle(
  fontSize = 9, fontName = "Arial", halign = "right", valign = "center",
  fgFill = "#FFFCE8", fontColour = "#7F6000",
  border = "TopBottomLeftRight", borderColour = "#BDD7EE",
  numFmt = "#,##0.00"
)
st_total_label <- createStyle(
  fontSize = 10, fontName = "Arial", halign = "left", valign = "center",
  fgFill = "#1F4E79", fontColour = "#FFFFFF", textDecoration = "bold",
  border = "TopBottomLeftRight", borderColour = "#FFFFFF"
)
st_total_num <- createStyle(
  fontSize = 10, fontName = "Arial", halign = "right", valign = "center",
  fgFill = "#2E75B6", fontColour = "#FFFFFF", textDecoration = "bold",
  border = "TopBottomLeftRight", borderColour = "#FFFFFF",
  numFmt = "#,##0.00"
)
st_total_pct <- createStyle(
  fontSize = 10, fontName = "Arial", halign = "right", valign = "center",
  fgFill = "#2E75B6", fontColour = "#FFFFFF", textDecoration = "bold",
  border = "TopBottomLeftRight", borderColour = "#FFFFFF",
  numFmt = "0.00%"
)
st_total_runoff <- createStyle(
  fontSize = 10, fontName = "Arial", halign = "right", valign = "center",
  fgFill = "#BF9000", fontColour = "#FFFFFF", textDecoration = "bold",
  border = "TopBottomLeftRight", borderColour = "#FFFFFF",
  numFmt = "#,##0.00"
)
st_rodape <- createStyle(
  fontSize = 9, fontColour = "#595959", fontName = "Arial",
  halign = "left", valign = "center", textDecoration = "italic"
)

# --- Layout fixo de colunas por aba ------------------------------------------
# Col 1: Classe
# Col 2: Área (m²)
# Col 3: Área (ha)
# Col 4: Área (km²)
# Col 5: % da AD
# Col 6: Valor CN ou C
# Col 7: Contribuição ponderada

COL_CLASSE <- 1
COL_M2     <- 2
COL_HA     <- 3
COL_KM2    <- 4
COL_PCT    <- 5
COL_RUNOFF <- 6
COL_CONTRIB <- 7
N_COLS     <- 7

linha_titulo    <- 1
linha_subtitulo <- 2
linha_cab       <- 3
linha_dados_ini <- 4

# --- Uma aba por AD ----------------------------------------------------------
for (ad in nomes_ad) {
  
  cat("  Gerando aba:", ad, "\n")
  
  # Nome da aba truncado para 31 caracteres (limite do Excel)
  nome_aba <- substr(ad, 1, 31)
  addWorksheet(wb, nome_aba)
  
  df      <- runoff_por_ad[[ad]]$tabela
  vp      <- runoff_por_ad[[ad]]$valor_ponderado
  n_cls   <- nrow(df)
  
  linha_dados_fim <- linha_dados_ini + n_cls - 1
  linha_total     <- linha_dados_fim + 1
  linha_rodape    <- linha_total + 2
  
  # Título
  mergeCells(wb, nome_aba, rows = linha_titulo, cols = COL_CLASSE:N_COLS)
  writeData(wb, nome_aba,
            paste0("USO DO SOLO E ", metodo_runoff, " PONDERADO — ", ad),
            startRow = linha_titulo, startCol = COL_CLASSE)
  addStyle(wb, nome_aba, st_titulo,
           rows = linha_titulo, cols = COL_CLASSE:N_COLS, gridExpand = TRUE)
  setRowHeights(wb, nome_aba, linha_titulo, 22)
  
  # Subtítulo
  mergeCells(wb, nome_aba, rows = linha_subtitulo, cols = COL_CLASSE:N_COLS)
  writeData(wb, nome_aba,
            paste0("Método: ", metodo_runoff,
                   " | Área total da AD: ",
                   format(round(sum(df$AD_m2), 2), big.mark = ".", decimal.mark = ","),
                   " m²"),
            startRow = linha_subtitulo, startCol = COL_CLASSE)
  addStyle(wb, nome_aba, st_subtitulo,
           rows = linha_subtitulo, cols = COL_CLASSE:N_COLS, gridExpand = TRUE)
  setRowHeights(wb, nome_aba, linha_subtitulo, 16)
  
  # Cabeçalhos
  cabecalhos <- c(
    "Classe de Uso do Solo",
    "Área (m²)", "Área (ha)", "Área (km²)",
    "% da AD",
    metodo_runoff,
    paste0(metodo_runoff, " × %")
  )
  for (k in seq_along(cabecalhos)) {
    writeData(wb, nome_aba, cabecalhos[k], startRow = linha_cab, startCol = k)
    st_cab_k <- if (k == COL_CLASSE) st_cab_classe else st_cab
    addStyle(wb, nome_aba, st_cab_k, rows = linha_cab, cols = k)
  }
  setRowHeights(wb, nome_aba, linha_cab, 24)
  
  # Dados
  for (i in seq_len(n_cls)) {
    linha_i <- linha_dados_ini + i - 1
    eh_par  <- (i %% 2 == 0)
    
    st_cls  <- if (eh_par) st_classe_par  else st_classe_impar
    st_num  <- if (eh_par) st_num_par     else st_num_impar
    st_pct  <- if (eh_par) st_pct_par     else st_pct_impar
    st_ro   <- if (eh_par) st_runoff_par  else st_runoff_impar
    
    writeData(wb, nome_aba, df$Classe[i],       startRow = linha_i, startCol = COL_CLASSE)
    writeData(wb, nome_aba, df$AD_m2[i],        startRow = linha_i, startCol = COL_M2)
    writeData(wb, nome_aba, df$AD_ha[i],        startRow = linha_i, startCol = COL_HA)
    writeData(wb, nome_aba, df$AD_km2[i],       startRow = linha_i, startCol = COL_KM2)
    writeData(wb, nome_aba, df$Perc_area[i],    startRow = linha_i, startCol = COL_PCT)
    writeData(wb, nome_aba, df$Valor_runoff[i], startRow = linha_i, startCol = COL_RUNOFF)
    writeData(wb, nome_aba, df$Contribuicao[i], startRow = linha_i, startCol = COL_CONTRIB)
    
    addStyle(wb, nome_aba, st_cls, rows = linha_i, cols = COL_CLASSE)
    addStyle(wb, nome_aba, st_num, rows = linha_i, cols = c(COL_M2, COL_HA, COL_KM2, COL_CONTRIB), gridExpand = TRUE)
    addStyle(wb, nome_aba, st_pct, rows = linha_i, cols = COL_PCT)
    addStyle(wb, nome_aba, st_ro,  rows = linha_i, cols = COL_RUNOFF)
    
    setRowHeights(wb, nome_aba, linha_i, 16)
  }
  
  # Linha de totais
  writeData(wb, nome_aba, "TOTAL / PONDERADO", startRow = linha_total, startCol = COL_CLASSE)
  addStyle(wb, nome_aba, st_total_label, rows = linha_total, cols = COL_CLASSE)
  
  # Totais de área via fórmula
  for (col in c(COL_M2, COL_HA, COL_KM2)) {
    cel <- paste0(int2col(col), linha_dados_ini, ":", int2col(col), linha_dados_fim)
    writeFormula(wb, nome_aba, paste0("=SUM(", cel, ")"), startRow = linha_total, startCol = col)
    addStyle(wb, nome_aba, st_total_num, rows = linha_total, cols = col)
  }
  
  # % total (deve ser 100%)
  cel_pct <- paste0(int2col(COL_PCT), linha_dados_ini, ":", int2col(COL_PCT), linha_dados_fim)
  writeFormula(wb, nome_aba, paste0("=SUM(", cel_pct, ")"), startRow = linha_total, startCol = COL_PCT)
  addStyle(wb, nome_aba, st_total_pct, rows = linha_total, cols = COL_PCT)
  
  # CN/C ponderado final
  cel_contrib <- paste0(int2col(COL_CONTRIB), linha_dados_ini, ":", int2col(COL_CONTRIB), linha_dados_fim)
  writeFormula(wb, nome_aba, paste0("=SUM(", cel_contrib, ")"), startRow = linha_total, startCol = COL_RUNOFF)
  addStyle(wb, nome_aba, st_total_runoff, rows = linha_total, cols = COL_RUNOFF)
  
  # Célula CN/C × % não se aplica ao total — deixar em branco com estilo
  addStyle(wb, nome_aba, st_total_num, rows = linha_total, cols = COL_CONTRIB)
  
  setRowHeights(wb, nome_aba, linha_total, 20)
  
  # Rodapé
  mergeCells(wb, nome_aba, rows = linha_rodape, cols = COL_CLASSE:N_COLS)
  writeData(wb, nome_aba,
            paste0(metodo_runoff, " ponderado calculado como: Σ(", metodo_runoff,
                   "_i × A_i) / A_total, onde A_i é a área de cada tipologia e A_total é a soma das áreas clipadas."),
            startRow = linha_rodape, startCol = COL_CLASSE)
  addStyle(wb, nome_aba, st_rodape,
           rows = linha_rodape, cols = COL_CLASSE:N_COLS, gridExpand = TRUE)
  setRowHeights(wb, nome_aba, linha_rodape, 20)
  
  # Larguras de coluna
  setColWidths(wb, nome_aba, cols = COL_CLASSE,  widths = 30)
  setColWidths(wb, nome_aba, cols = COL_M2,      widths = 18)
  setColWidths(wb, nome_aba, cols = COL_HA,      widths = 13)
  setColWidths(wb, nome_aba, cols = COL_KM2,     widths = 13)
  setColWidths(wb, nome_aba, cols = COL_PCT,     widths = 10)
  setColWidths(wb, nome_aba, cols = COL_RUNOFF,  widths = 12)
  setColWidths(wb, nome_aba, cols = COL_CONTRIB, widths = 16)
  
  # Congelar painel
  freezePane(wb, nome_aba, firstActiveRow = linha_dados_ini, firstActiveCol = 2)
}

# --- Salvar com timestamp ----------------------------------------------------
timestamp     <- format(Sys.time(), "%Y%m%d_%H%M%S")
nome_arquivo  <- paste0("tabela_uso_solo_por_AD_", timestamp, ".xlsx")
caminho_xlsx  <- file.path(pasta_saida, nome_arquivo)
saveWorkbook(wb, caminho_xlsx, overwrite = TRUE)

cat("Excel salvo em:\n ", caminho_xlsx, "\n")
cat("\n=== PROCESSAMENTO CONCLUÍDO ===\n")
