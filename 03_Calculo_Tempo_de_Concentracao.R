library(terra)
library(sf)
library(tcltk)

# ── 1. CAIXAS DE DIÁLOGO ────────────────────────────────────────────────────

# MDT
mdt_path <- tk_choose.files(
  caption = "Selecione o MDT (.tif)",
  filter  = matrix(c("GeoTIFF", "*.tif;*.tiff"), 1, 2)
)
if (!length(mdt_path)) stop("Nenhum MDT selecionado.")

# Shapefiles de linha (múltiplos)
shp_paths <- tk_choose.files(
  caption = "Selecione os Shapefiles de linha (pode selecionar vários)",
  filter  = matrix(c("Shapefile", "*.shp"), 1, 2),
  multi   = TRUE
)
if (!length(shp_paths)) stop("Nenhum shapefile selecionado.")

# Pasta de saída
pasta_saida <- tk_choose.dir(caption = "Selecione a pasta de saída")
if (is.na(pasta_saida)) stop("Nenhuma pasta de saída selecionada.")

# ── 2. CARREGA MDT ───────────────────────────────────────────────────────────

mdt <- rast(mdt_path)
cat("\nMDT carregado:", mdt_path, "\n")

# ── 3. FUNÇÃO: declividade de uma feição ─────────────────────────────────────

calc_decliv <- function(linha_sf, mdt) {
  
  # Reprojeção
  linha_sf <- st_transform(linha_sf, crs(mdt))
  
  # Comprimento 2D
  comprimento_total <- as.numeric(st_length(linha_sf))
  
  # Pontos a cada 1m
  pts <- st_line_sample(linha_sf, density = 1)
  pts <- st_cast(pts, "POINT")
  
  # Garante que há pontos suficientes
  if (length(pts) < 2) {
    warning("Feição muito curta para amostrar — ignorada.")
    return(NULL)
  }
  
  # Extrai cotas
  cotas  <- extract(mdt, vect(pts))[, 2]
  coords <- st_coordinates(pts)
  
  n          <- length(cotas)
  L2D_vec    <- numeric(n - 1)
  L3D_vec    <- numeric(n - 1)
  slope_vec  <- numeric(n - 1)
  
  for (i in 1:(n - 1)) {
    dx <- coords[i+1, "X"] - coords[i, "X"]
    dy <- coords[i+1, "Y"] - coords[i, "Y"]
    dz <- cotas[i+1] - cotas[i]
    
    L2D <- sqrt(dx^2 + dy^2)
    L3D <- sqrt(dx^2 + dy^2 + dz^2)
    
    L2D_vec[i]  <- L2D
    L3D_vec[i]  <- L3D
    slope_vec[i] <- abs(dz) / L2D * 100
  }
  
  list(
    comprimento_2D  = comprimento_total,
    n_segmentos     = n - 1,
    decliv_media    = sum(slope_vec * L3D_vec, na.rm = TRUE) / sum(L3D_vec, na.rm = TRUE),
    decliv_min      = min(slope_vec, na.rm = TRUE),
    decliv_max      = max(slope_vec, na.rm = TRUE),
    cota_inicial    = cotas[1],
    cota_final      = cotas[n],
    delta_H         = abs(cotas[n] - cotas[1])
  )
}

# ── 4. LOOP PRINCIPAL ────────────────────────────────────────────────────────

resultados_todos <- list()

for (shp_path in shp_paths) {
  
  nome_arq <- tools::file_path_sans_ext(basename(shp_path))
  cat("\n========================================\n")
  cat("Processando:", nome_arq, "\n")
  
  camada <- st_read(shp_path, quiet = TRUE)
  
  # Garante geometria de linha
  camada <- camada[st_geometry_type(camada) %in% c("LINESTRING", "MULTILINESTRING"), ]
  if (nrow(camada) == 0) {
    warning("Nenhuma feição de linha em: ", nome_arq)
    next
  }
  
  # Explode MULTILINESTRING em LINESTRING individuais
  camada <- st_cast(camada, "LINESTRING")
  n_feicoes <- nrow(camada)
  cat("Feições encontradas:", n_feicoes, "\n")
  
  res_arq <- data.frame()
  
  for (j in seq_len(n_feicoes)) {
    
    cat("  Feição", j, "/", n_feicoes, "... ")
    feicao <- camada[j, ]
    
    res <- tryCatch(
      calc_decliv(feicao, mdt),
      error = function(e) { message("ERRO: ", e$message); NULL }
    )
    
    if (is.null(res)) next
    
    cat("Média:", round(res$decliv_media, 2), "%\n")
    
    res_arq <- rbind(res_arq, data.frame(
      arquivo        = nome_arq,
      feicao_id      = j,
      comprimento_2D = round(res$comprimento_2D, 1),
      n_segmentos    = res$n_segmentos,
      cota_inicial   = round(res$cota_inicial, 2),
      cota_final     = round(res$cota_final, 2),
      delta_H        = round(res$delta_H, 2),
      decliv_media   = round(res$decliv_media, 2),
      decliv_min     = round(res$decliv_min, 2),
      decliv_max     = round(res$decliv_max, 2)
    ))
  }
  
  resultados_todos[[nome_arq]] <- res_arq
  
  # CSV por shapefile
  csv_path <- file.path(pasta_saida, paste0(nome_arq, "_declividade.csv"))
  write.csv(res_arq, csv_path, row.names = FALSE)
  cat("CSV salvo:", csv_path, "\n")
}

# ── 5. CSV CONSOLIDADO ───────────────────────────────────────────────────────

todos <- do.call(rbind, resultados_todos)

# ── 6. DECLIVIDADE MÉDIA EM m/m ─────────────────────────────────────────────

todos$decliv_media_m_m <- round(todos$decliv_media / 100, 6)

# ── 7. TEMPO DE CONCENTRAÇÃO – KIRPICH ──────────────────────────────────────

todos$tc_kirpich_min <- round(
  0.0195 * (todos$comprimento_2D ^ 0.77) / (todos$decliv_media_m_m ^ 0.385),
  2
)

# ── 8. SALVA CSV CONSOLIDADO COM TIMESTAMP ───────────────────────────────────

timestamp <- format(Sys.time(), "%Y%m%d_%H%M%S")
csv_consolidado <- file.path(pasta_saida, paste0("declividade_consolidado_", timestamp, ".csv"))
write.csv(todos, csv_consolidado, row.names = FALSE)

cat("\n========================================\n")
cat("CONCLUÍDO!\n")
cat("CSV consolidado salvo em:\n", csv_consolidado, "\n")
print(todos)