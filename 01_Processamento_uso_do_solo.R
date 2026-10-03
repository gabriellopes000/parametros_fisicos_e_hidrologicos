# =============================================================================
# SCRIPT: Ponderação de CN / C por Área de Drenagem (Shiny local)
# =============================================================================
# Substitui os scripts 00, 01, 02 e 02_v2.
#
# FLUXO:
#   1. Uso do Solo  → lê o shapefile geral, padroniza as classes (acentos,
#                     maiúsculas, variantes) e define os valores de CN e C
#   2. ADs          → lê os shapefiles das Áreas de Drenagem e permite
#                     renomear Dispositivo / AD
#   3. Resultados   → recorte, ponderação, mapa e exportação (Excel + shapefiles)
#
# O recorte é feito uma única vez e fica em cache: mudar CN↔C, os valores ou os
# nomes das ADs não refaz a interseção.
# =============================================================================

# -----------------------------------------------------------------------------
# PACOTES: instala os que faltam e carrega todos
# -----------------------------------------------------------------------------
carregar_pacotes <- function(pacotes) {
  faltando <- pacotes[!vapply(pacotes, requireNamespace, logical(1), quietly = TRUE)]
  
  if (length(faltando) > 0) {
    cat("Instalando pacotes:", paste(faltando, collapse = ", "), "\n")
    install.packages(faltando, dependencies = TRUE)
  }
  
  for (p in pacotes) {
    suppressPackageStartupMessages(library(p, character.only = TRUE))
  }
  cat("Pacotes carregados:", paste(pacotes, collapse = ", "), "\n")
}

carregar_pacotes(c("shiny", "sf", "dplyr", "tidyr", "stringi",
                   "rhandsontable", "leaflet", "openxlsx", "tcltk"))
# -----------------------------------------------------------------------------
# 0. VALORES PADRÃO DE CN / C (usados como dicionário inicial)
# -----------------------------------------------------------------------------
VALORES_PADRAO <- data.frame(
  Classe_padronizada = c("Área industrial", "Cava", "Enrocamento", "Solo Exposto",
                         "Vegetação de Baixo Porte", "Vegetação de Grande Porte",
                         "Vegetação de Médio Porte"),
  CN = c(85, 90, 75, 91, 74, 55, 65),
  C  = c(0.38, 0.65, 0.40, 0.53, 0.40, 0.13, 0.38),
  stringsAsFactors = FALSE
)

ABAS_RESERVADAS <- c("resumo_dispositivos", "padronizacao_classes", "valores")
LINHA_INI       <- 4   # primeira linha de dados nas abas do Excel

# =============================================================================
# 1. FUNÇÕES DE LEITURA E PREPARAÇÃO
# =============================================================================

escolher_arquivos <- function(titulo, filtro, multi = FALSE) {
  x <- tk_choose.files(caption = titulo, filter = filtro, multi = multi)
  x[x != ""]
}

# Chave normalizada: sem acentos, minúsculas, sem espaços extras
chave_classe <- function(x) {
  x <- enc2utf8(as.character(x))
  x <- stri_trans_general(x, "Latin-ASCII")
  x <- gsub("\\s+", " ", trimws(x))
  tolower(x)
}

# Mantém apenas polígonos (st_make_valid / st_intersection podem gerar coleções)
so_poligonos <- function(x) {
  tipos <- as.character(st_geometry_type(x))
  if (any(tipos == "GEOMETRYCOLLECTION")) {
    x <- suppressWarnings(st_collection_extract(x, "POLYGON"))
  }
  x[as.character(st_geometry_type(x)) %in% c("POLYGON", "MULTIPOLYGON"), ]
}

ler_camada <- function(caminho, encoding = "UTF-8") {
  x <- st_read(caminho, options = paste0("ENCODING=", encoding), quiet = TRUE)
  x <- st_zm(x, drop = TRUE, what = "ZM")
  if (is.na(st_crs(x))) stop("Shapefile sem CRS definido: ", basename(caminho))
  
  n_invalidas <- sum(!st_is_valid(x), na.rm = TRUE)
  if (n_invalidas > 0) x <- st_make_valid(x)
  x <- so_poligonos(x[!st_is_empty(x), ])
  
  # Áreas em metros: reprojeta para SIRGAS 2000 Policônica se estiver em graus
  reprojetado <- st_is_longlat(x)
  if (reprojetado) x <- st_transform(x, 5880)
  
  attr(x, "n_invalidas") <- n_invalidas
  attr(x, "reprojetado") <- reprojetado
  x
}

campos_atributos <- function(x) setdiff(names(x), attr(x, "sf_column"))

tem_mojibake <- function(x) any(grepl("\u00C3|\u00C2", x))

# -----------------------------------------------------------------------------
# Tabela 1: padronização das classes (original → padronizada)
# Prioridade da sugestão:
#   1) variante já registrada no dicionário (Classe_original)
#   2) nome padronizado do dicionário com a mesma chave
#   3) variante de maior área entre as que têm a mesma chave
# -----------------------------------------------------------------------------
montar_tabela_classes <- function(uso, campo, dic_classes, dic_valores) {
  tab <- data.frame(
    Classe_original = trimws(as.character(uso[[campo]])),
    area_m2         = as.numeric(st_area(uso)),
    stringsAsFactors = FALSE
  ) %>%
    mutate(Classe_original = ifelse(is.na(Classe_original) | Classe_original == "",
                                    "(sem classe)", Classe_original)) %>%
    group_by(Classe_original) %>%
    summarise(Feicoes = n(), Area_ha = round(sum(area_m2) / 1e4, 2), .groups = "drop") %>%
    mutate(chave = chave_classe(Classe_original)) %>%
    group_by(chave) %>%
    mutate(Classe_padronizada = Classe_original[which.max(Area_ha)],
           Sugestao = ifelse(n() > 1, "Agrupada (acentos/maiúsculas)", "Original")) %>%
    ungroup()
  
  if (!is.null(dic_valores) && nrow(dic_valores) > 0) {
    i <- match(tab$chave, chave_classe(dic_valores$Classe_padronizada))
    tab$Classe_padronizada <- ifelse(is.na(i), tab$Classe_padronizada,
                                     dic_valores$Classe_padronizada[i])
    tab$Sugestao[!is.na(i)] <- "Dicionário"
  }
  if (!is.null(dic_classes) && nrow(dic_classes) > 0) {
    i <- match(tab$chave, chave_classe(dic_classes$Classe_original))
    tab$Classe_padronizada <- ifelse(is.na(i), tab$Classe_padronizada,
                                     dic_classes$Classe_padronizada[i])
    tab$Sugestao[!is.na(i)] <- "Dicionário (variante)"
  }
  
  tab %>%
    arrange(Classe_padronizada, Classe_original) %>%
    select(Classe_original, Feicoes, Area_ha, Sugestao, Classe_padronizada) %>%
    as.data.frame()
}

# -----------------------------------------------------------------------------
# Tabela 2: valores de CN / C por classe padronizada
# Mantém os valores já digitados; completa com o dicionário pela chave
# -----------------------------------------------------------------------------
montar_tabela_valores <- function(padronizadas, valores_atuais, dic_valores) {
  v <- data.frame(Classe_padronizada = sort(unique(trimws(padronizadas))),
                  stringsAsFactors = FALSE)
  fonte <- bind_rows(valores_atuais, dic_valores)
  if (is.null(fonte) || nrow(fonte) == 0) {
    v$CN <- NA_real_; v$C <- NA_real_
    return(v)
  }
  # Para cada classe, o primeiro valor não vazio (tabela atual tem prioridade)
  fonte <- fonte %>%
    mutate(chave = chave_classe(Classe_padronizada),
           CN = as.numeric(CN), C = as.numeric(C)) %>%
    group_by(chave) %>%
    summarise(CN = CN[!is.na(CN)][1], C = C[!is.na(C)][1], .groups = "drop")
  i <- match(chave_classe(v$Classe_padronizada), fonte$chave)
  v$CN <- as.numeric(fonte$CN[i])
  v$C  <- as.numeric(fonte$C[i])
  v
}

# -----------------------------------------------------------------------------
# ADs: uma feição por AD, com id interno fixo (id_ad)
#   modo "arquivo": cada shapefile é uma AD (feições unidas)
#   modo "feicao" : cada valor do campo escolhido é uma AD
# -----------------------------------------------------------------------------
preparar_ads <- function(ads_raw, modo, campo = NULL) {
  lista <- lapply(names(ads_raw), function(caminho) {
    x   <- ads_raw[[caminho]]
    arq <- tools::file_path_sans_ext(basename(caminho))
    if (modo == "arquivo") {
      x <- x %>%
        summarise(geometry = st_union(geometry)) %>%
        mutate(Nome_original = arq)
    } else {
      if (!campo %in% names(x)) stop("Campo '", campo, "' não existe em ", arq)
      nomes <- trimws(as.character(x[[campo]]))
      faltam <- is.na(nomes) | nomes == ""
      nomes[faltam] <- paste0(arq, "_", which(faltam))
      x$Nome_original <- nomes
      x <- x %>%
        group_by(Nome_original) %>%
        summarise(geometry = st_union(geometry), .groups = "drop")
    }
    x %>% mutate(Arquivo = basename(caminho)) %>% select(Arquivo, Nome_original, geometry)
  })
  crs_ref <- st_crs(lista[[1]])
  lista   <- lapply(lista, st_transform, crs_ref)
  ads <- do.call(rbind, lista)
  ads$id_ad <- seq_len(nrow(ads))
  ads
}

montar_tabela_ads <- function(ads, dic_ads) {
  tab <- data.frame(
    id_ad         = ads$id_ad,
    Arquivo       = ads$Arquivo,
    Nome_original = ads$Nome_original,
    Area_ha       = round(as.numeric(st_area(ads)) / 1e4, 2),
    Dispositivo   = ads$Nome_original,
    AD            = ads$Nome_original,
    stringsAsFactors = FALSE
  )
  if (!is.null(dic_ads) && nrow(dic_ads) > 0) {
    i  <- match(tab$Nome_original, dic_ads$Nome_original)
    ok <- !is.na(i)
    tab$Dispositivo[ok] <- dic_ads$Dispositivo[i[ok]]
    tab$AD[ok]          <- dic_ads$AD[i[ok]]
  }
  tab
}

validar_nomes_ad <- function(nomes) {
  erros  <- character()
  vazios <- is.na(nomes) | trimws(nomes) == ""
  if (any(vazios)) erros <- c(erros, "Há ADs sem nome.")
  dup <- unique(nomes[duplicated(tolower(nomes)) & !vazios])   # Excel ignora maiúsculas
  if (length(dup)) erros <- c(erros, paste("Nomes duplicados:", paste(dup, collapse = ", ")))
  inval <- nomes[grepl("[\\[\\]:*?/\\\\<>\"|]", nomes, perl = TRUE)]
  if (length(inval)) erros <- c(erros, paste("Caracteres inválidos ([ ] : * ? / \\ < > \" |):",
                                             paste(inval, collapse = ", ")))
  longos <- nomes[!is.na(nomes) & nchar(nomes) > 31]
  if (length(longos)) erros <- c(erros, paste("Mais de 31 caracteres (limite de aba do Excel):",
                                              paste(longos, collapse = ", ")))
  reserv <- nomes[tolower(nomes) %in% ABAS_RESERVADAS]
  if (length(reserv)) erros <- c(erros, paste("Nome reservado:", paste(reserv, collapse = ", ")))
  erros
}

# =============================================================================
# 2. FUNÇÕES DE CÁLCULO
# =============================================================================

# Interseção uso do solo × ADs com a classe ORIGINAL (fica em cache)
intersectar <- function(uso, campo, ads) {
  uso <- uso %>%
    mutate(Classe_original = trimws(as.character(.data[[campo]])),
           Classe_original = ifelse(is.na(Classe_original) | Classe_original == "",
                                    "(sem classe)", Classe_original)) %>%
    select(Classe_original)
  ads <- st_transform(ads, st_crs(uso)) %>% select(id_ad)
  inter <- suppressWarnings(st_intersection(uso, ads))
  so_poligonos(inter)
}

# Aplica a padronização, une por AD × Classe e pondera
ponderar_runoff <- function(inter, ads, mapa_classes, valores, metodo) {
  geom <- inter %>%
    mutate(Classe = unname(mapa_classes[Classe_original])) %>%
    group_by(id_ad, Classe) %>%
    summarise(.groups = "drop") %>%          # une as geometrias de cada grupo
    so_poligonos() %>%
    st_cast("MULTIPOLYGON", warn = FALSE)
  geom$area_m2 <- as.numeric(st_area(geom))
  geom <- geom[geom$area_m2 > 0, ]
  if (nrow(geom) == 0) stop("Nenhuma AD intersecta o uso do solo. Verifique os CRS e a localização.")
  
  vals <- data.frame(Classe = valores$Classe_padronizada,
                     Valor  = as.numeric(valores[[metodo]]),
                     stringsAsFactors = FALSE)
  geom <- left_join(geom, vals, by = "Classe")
  
  classes_presentes <- sort(unique(geom$Classe))
  tabela <- geom %>%
    st_drop_geometry() %>%
    select(id_ad, Classe, area_m2) %>%
    complete(id_ad = ads$id_ad, Classe = classes_presentes, fill = list(area_m2 = 0)) %>%
    left_join(vals, by = "Classe") %>%
    group_by(id_ad) %>%
    mutate(Perc         = if (sum(area_m2) > 0) area_m2 / sum(area_m2) else NA_real_,
           Contribuicao = Valor * Perc) %>%
    ungroup() %>%
    arrange(id_ad, Classe)
  
  resumo <- tabela %>%
    group_by(id_ad) %>%
    summarise(Area_uso_m2 = sum(area_m2),
              Coeficiente = if (sum(area_m2) > 0) sum(Contribuicao) else NA_real_,
              .groups = "drop") %>%
    left_join(data.frame(id_ad = ads$id_ad, Area_AD_m2 = as.numeric(st_area(ads))),
              by = "id_ad") %>%
    mutate(Cobertura = Area_uso_m2 / Area_AD_m2)
  
  list(geom = geom, tabela = tabela, resumo = resumo, metodo = metodo)
}

# =============================================================================
# 3. FUNÇÕES DE EXPORTAÇÃO
# =============================================================================

est_titulo <- createStyle(fontSize = 13, fontColour = "#FFFFFF", fontName = "Arial",
                          fgFill = "#1F4E79", halign = "center", valign = "center",
                          textDecoration = "bold")
est_subtitulo <- createStyle(fontSize = 9, fontColour = "#595959", fontName = "Arial",
                             halign = "left", valign = "center", textDecoration = "italic")
est_cab <- createStyle(fontSize = 10, fontColour = "#FFFFFF", fontName = "Arial",
                       fgFill = "#2E75B6", halign = "center", valign = "center",
                       textDecoration = "bold", wrapText = TRUE,
                       border = "TopBottomLeftRight", borderColour = "#FFFFFF")

fmt_numero <- function(tipo, fmt_runoff) {
  switch(tipo, texto = "GENERAL", pct = "0.00%", runoff = fmt_runoff, "#,##0.00")
}

est_celula <- function(tipo, par, fmt_runoff) {
  fundo <- if (tipo == "runoff") { if (par) "#FFF2CC" else "#FFFCE8" } else { if (par) "#DEEAF1" else "#FFFFFF" }
  createStyle(fontSize = 9, fontName = "Arial", valign = "center",
              halign = if (tipo == "texto") "left" else "right",
              fgFill = fundo,
              fontColour = if (tipo == "runoff") "#7F6000" else "#000000",
              textDecoration = if (tipo == "texto") "bold" else NULL,
              border = "TopBottomLeftRight", borderColour = "#BDD7EE",
              numFmt = fmt_numero(tipo, fmt_runoff))
}

est_total <- function(tipo, fmt_runoff) {
  createStyle(fontSize = 10, fontName = "Arial", valign = "center",
              halign = if (tipo == "texto") "left" else "right",
              fgFill = switch(tipo, texto = "#1F4E79", runoff = "#BF9000", "#2E75B6"),
              fontColour = "#FFFFFF", textDecoration = "bold",
              border = "TopBottomLeftRight", borderColour = "#FFFFFF",
              numFmt = fmt_numero(tipo, fmt_runoff))
}

soma_col <- function(col, n) {
  paste0("SUM(", int2col(col), LINHA_INI, ":", int2col(col), LINHA_INI + n - 1, ")")
}

# total: lista com um elemento por coluna — NA (vazio), texto, ou fórmula "SUM(...)"
escrever_aba <- function(wb, aba, titulo, subtitulo, df, tipos, larguras,
                         total = NULL, rodape = NULL, fmt_runoff = "0.00") {
  addWorksheet(wb, aba)
  nc <- ncol(df); n <- nrow(df)
  
  mergeCells(wb, aba, cols = 1:nc, rows = 1)
  writeData(wb, aba, titulo, startRow = 1, startCol = 1)
  addStyle(wb, aba, est_titulo, rows = 1, cols = 1:nc, gridExpand = TRUE)
  setRowHeights(wb, aba, 1, 22)
  
  mergeCells(wb, aba, cols = 1:nc, rows = 2)
  writeData(wb, aba, subtitulo, startRow = 2, startCol = 1)
  addStyle(wb, aba, est_subtitulo, rows = 2, cols = 1:nc, gridExpand = TRUE)
  
  writeData(wb, aba, df, startRow = LINHA_INI - 1, startCol = 1, headerStyle = est_cab)
  setRowHeights(wb, aba, LINHA_INI - 1, 30)
  
  for (j in seq_len(nc)) for (par in c(TRUE, FALSE)) {
    linhas <- LINHA_INI - 1 + which((seq_len(n) %% 2 == 0) == par)
    if (length(linhas)) addStyle(wb, aba, est_celula(tipos[j], par, fmt_runoff),
                                 rows = linhas, cols = j)
  }
  
  ultima <- LINHA_INI + n - 1
  if (!is.null(total)) {
    ultima <- ultima + 1
    for (j in seq_len(nc)) {
      t <- total[[j]]
      if (!is.na(t)) {
        if (startsWith(t, "SUM(")) writeFormula(wb, aba, t, startRow = ultima, startCol = j)
        else writeData(wb, aba, t, startRow = ultima, startCol = j)
      }
      tipo_t <- if (j == 1) "texto" else if (tipos[j] == "texto") "num" else tipos[j]
      addStyle(wb, aba, est_total(tipo_t, fmt_runoff), rows = ultima, cols = j)
    }
    setRowHeights(wb, aba, ultima, 20)
  }
  
  if (!is.null(rodape)) {
    lr <- ultima + 2
    mergeCells(wb, aba, cols = 1:nc, rows = lr)
    writeData(wb, aba, rodape, startRow = lr, startCol = 1)
    addStyle(wb, aba, est_subtitulo, rows = lr, cols = 1:nc, gridExpand = TRUE)
  }
  
  setColWidths(wb, aba, cols = 1:nc, widths = larguras)
  freezePane(wb, aba, firstActiveRow = LINHA_INI, firstActiveCol = 2)
}

exportar_excel <- function(res, nomes_ad, tab_classes, tab_valores, caminho) {
  metodo     <- res$metodo
  fmt_runoff <- if (metodo == "C") "0.000" else "0.00"
  wb <- createWorkbook()
  
  resumo <- res$resumo %>% left_join(nomes_ad, by = "id_ad") %>% arrange(id_ad)
  n <- nrow(resumo)
  
  # --- Resumo -----------------------------------------------------------------
  df_res <- data.frame(
    resumo$Dispositivo, resumo$AD,
    resumo$Area_AD_m2, resumo$Area_AD_m2 / 1e4, resumo$Area_AD_m2 / 1e6,
    resumo$Cobertura, resumo$Coeficiente
  )
  names(df_res) <- c("Dispositivo", "AD", "Área (m²)", "Área (ha)", "Área (km²)",
                     "Cobertura do uso do solo", metodo)
  escrever_aba(
    wb, "Resumo_Dispositivos",
    paste0("RESUMO POR DISPOSITIVO — ÁREA E ", metodo, " PONDERADO"),
    paste0("Método: ", metodo, " | Dispositivos: ", n,
           " | Gerado em ", format(Sys.time(), "%d/%m/%Y %H:%M")),
    df_res,
    tipos    = c("texto", "texto", "num", "num", "num", "pct", "runoff"),
    larguras = c(28, 20, 16, 12, 12, 14, 10),
    total    = list("TOTAL", NA, soma_col(3, n), soma_col(4, n), soma_col(5, n), NA, NA),
    rodape   = paste0("Áreas = polígono da AD. ", metodo,
                      " ponderado pela área de uso do solo recortada (cobertura indicada)."),
    fmt_runoff = fmt_runoff
  )
  
  # --- Uma aba por AD ---------------------------------------------------------
  for (k in seq_len(n)) {
    id <- resumo$id_ad[k]
    df <- res$tabela %>% filter(id_ad == id)
    df_ad <- data.frame(df$Classe, df$area_m2, df$area_m2 / 1e4, df$area_m2 / 1e6,
                        df$Perc, df$Valor, df$Contribuicao)
    names(df_ad) <- c("Classe de Uso do Solo", "Área (m²)", "Área (ha)", "Área (km²)",
                      "% da AD", metodo, paste0(metodo, " × %"))
    m <- nrow(df_ad)
    escrever_aba(
      wb, resumo$AD[k],
      paste0("USO DO SOLO E ", metodo, " PONDERADO — ", resumo$AD[k]),
      paste0("Dispositivo: ", resumo$Dispositivo[k],
             " | Área da AD: ", format(round(resumo$Area_AD_m2[k], 2), big.mark = ".", decimal.mark = ","),
             " m² | Cobertura: ", format(round(100 * resumo$Cobertura[k], 2), decimal.mark = ","), "%"),
      df_ad,
      tipos    = c("texto", "num", "num", "num", "pct", "runoff", "num"),
      larguras = c(30, 18, 13, 13, 10, 12, 16),
      total    = list("TOTAL / PONDERADO", soma_col(2, m), soma_col(3, m), soma_col(4, m),
                      soma_col(5, m), soma_col(7, m), NA),
      rodape   = paste0(metodo, " ponderado = Σ(", metodo, "_i × A_i) / A_total, onde A_total ",
                        "é a soma das áreas de uso do solo recortadas na AD."),
      fmt_runoff = fmt_runoff
    )
  }
  
  # --- Rastreabilidade --------------------------------------------------------
  escrever_aba(wb, "Padronizacao_Classes", "PADRONIZAÇÃO DAS CLASSES DE USO DO SOLO",
               "Classe original no shapefile → classe usada no cálculo",
               tab_classes[, c("Classe_original", "Feicoes", "Area_ha", "Classe_padronizada")],
               tipos = c("texto", "num", "num", "texto"), larguras = c(32, 10, 12, 32))
  escrever_aba(wb, "Valores", "VALORES DE CN E C POR CLASSE",
               paste0("Método usado: ", metodo), tab_valores,
               tipos = c("texto", "runoff", "runoff"), larguras = c(32, 10, 10),
               fmt_runoff = "0.000")
  
  saveWorkbook(wb, caminho, overwrite = TRUE)
}

exportar_shapes <- function(res, nomes_ad, pasta) {
  arquivos <- character()
  for (k in seq_len(nrow(nomes_ad))) {
    x <- res$geom %>% filter(id_ad == nomes_ad$id_ad[k])
    if (nrow(x) == 0) next
    x <- x %>%
      transmute(AD = nomes_ad$AD[k], Classe,
                AD_m2 = area_m2, AD_ha = area_m2 / 1e4, AD_km2 = area_m2 / 1e6,
                Valor)
    names(x)[names(x) == "Valor"] <- res$metodo
    arq <- file.path(pasta, paste0(nomes_ad$AD[k], "_uso_solo.shp"))
    st_write(x, arq, delete_layer = TRUE, quiet = TRUE, layer_options = "ENCODING=UTF-8")
    arquivos <- c(arquivos, arq)
  }
  arquivos
}

# Dicionário: um .xlsx com abas Classes, Valores e ADs
salvar_dicionario <- function(caminho, tab_classes, tab_valores, tab_ads, dic_antigo) {
  juntar <- function(novo, antigo, chave) {
    out <- bind_rows(novo, antigo)
    out[!duplicated(chave_classe(out[[chave]])), , drop = FALSE]
  }
  cls <- juntar(tab_classes[, c("Classe_original", "Classe_padronizada")], dic_antigo$classes, "Classe_original")
  val <- juntar(tab_valores, dic_antigo$valores, "Classe_padronizada")
  ads <- if (!is.null(tab_ads)) tab_ads[, c("Nome_original", "Dispositivo", "AD")] else NULL
  ads <- bind_rows(ads, dic_antigo$ads)
  if (!is.null(ads) && nrow(ads)) ads <- ads[!duplicated(ads$Nome_original), ]
  write.xlsx(list(Classes = cls, Valores = val,
                  ADs = if (is.null(ads)) data.frame(Nome_original = character(), Dispositivo = character(), AD = character()) else ads),
             caminho, overwrite = TRUE)
}

ler_dicionario <- function(caminho) {
  abas <- getSheetNames(caminho)
  ler  <- function(aba) if (aba %in% abas) read.xlsx(caminho, sheet = aba) else NULL
  dic  <- list(classes = ler("Classes"), valores = ler("Valores"), ads = ler("ADs"))
  if (!is.null(dic$valores)) {
    dic$valores$CN <- as.numeric(dic$valores$CN)
    dic$valores$C  <- as.numeric(dic$valores$C)
  }
  dic
}

# =============================================================================
# 4. INTERFACE
# =============================================================================

ui <- navbarPage(
  "Ponderação de CN / C",
  
  # --- Aba 1 ------------------------------------------------------------------
  tabPanel("1. Uso do Solo",
           sidebarLayout(
             sidebarPanel(width = 3,
                          actionButton("btn_uso", "Selecionar shapefile de uso do solo", class = "btn-primary", width = "100%"),
                          br(), br(),
                          selectInput("enc_uso", "Codificação do .dbf", c("UTF-8", "CP1252")),
                          selectInput("campo_classe", "Campo com a classe", choices = NULL),
                          hr(),
                          actionButton("btn_carregar_dic", "Carregar dicionário (.xlsx)", width = "100%"),
                          br(), br(),
                          actionButton("btn_salvar_dic", "Salvar dicionário (.xlsx)", width = "100%"),
                          hr(),
                          verbatimTextOutput("info_uso")
             ),
             mainPanel(width = 9,
                       h4("Tabela 1 — Padronização das classes"),
                       helpText("Edite 'Classe_padronizada'. Variantes com o mesmo nome padronizado serão unidas.",
                                "Depois clique em 'Atualizar tabela de valores'."),
                       rHandsontableOutput("tab_classes"),
                       br(),
                       actionButton("btn_atualizar_valores", "Atualizar tabela de valores", class = "btn-success"),
                       hr(),
                       h4("Tabela 2 — Valores de CN e C"),
                       rHandsontableOutput("tab_valores")
             )
           )
  ),
  
  # --- Aba 2 ------------------------------------------------------------------
  tabPanel("2. Áreas de Drenagem",
           sidebarLayout(
             sidebarPanel(width = 3,
                          actionButton("btn_ads", "Selecionar shapefiles das ADs", class = "btn-primary", width = "100%"),
                          br(), br(),
                          selectInput("enc_ads", "Codificação do .dbf", c("UTF-8", "CP1252")),
                          radioButtons("modo_ad", "Cada AD é:",
                                       c("Um shapefile (une as feições)" = "arquivo",
                                         "Uma feição, identificada por um campo" = "feicao")),
                          conditionalPanel("input.modo_ad == 'feicao'",
                                           selectInput("campo_ad", "Campo com o nome da AD", choices = NULL)),
                          actionButton("btn_montar_ads", "Montar tabela de ADs", class = "btn-success", width = "100%"),
                          hr(),
                          verbatimTextOutput("info_ads")
             ),
             mainPanel(width = 9,
                       h4("Nomes dos dispositivos e ADs"),
                       helpText("Edite 'Dispositivo' e 'AD'. O nome da AD vira o nome da aba no Excel e do shapefile exportado",
                                "(máx. 31 caracteres, sem [ ] : * ? / \\ < > \" |)."),
                       rHandsontableOutput("tab_ads")
             )
           )
  ),
  
  # --- Aba 3 ------------------------------------------------------------------
  tabPanel("3. Resultados",
           sidebarLayout(
             sidebarPanel(width = 3,
                          radioButtons("metodo", "Método", c("Coeficiente de runoff (C)" = "C", "Curve Number (CN)" = "CN")),
                          actionButton("btn_calcular", "Calcular", class = "btn-primary", width = "100%"),
                          hr(),
                          actionButton("btn_pasta", "Selecionar pasta de saída", width = "100%"),
                          verbatimTextOutput("info_pasta"),
                          checkboxInput("exp_shp", "Exportar shapefiles recortados por AD", TRUE),
                          actionButton("btn_exportar", "Exportar", class = "btn-success", width = "100%"),
                          hr(),
                          verbatimTextOutput("info_resultado")
             ),
             mainPanel(width = 9,
                       h4("Resumo por dispositivo"),
                       tableOutput("tab_resumo"),
                       hr(),
                       fluidRow(
                         column(5, selectInput("ad_detalhe", "Detalhe da AD", choices = NULL), tableOutput("tab_detalhe")),
                         column(7, leafletOutput("mapa", height = 500))
                       )
             )
           )
  )
)

# =============================================================================
# 5. SERVIDOR
# =============================================================================

server <- function(input, output, session) {
  
  rv <- reactiveValues(
    caminho_uso = NULL, uso = NULL, uso_versao = 0,
    classes_ini = NULL, valores_ini = NULL,
    ads_raw = NULL, ads = NULL, ads_versao = 0, ads_tab_ini = NULL,
    dic = list(classes = NULL, valores = VALORES_PADRAO, ads = NULL),
    cache = NULL, res = NULL, nomes_ad = NULL,
    tab_classes_calc = NULL, tab_valores_calc = NULL,
    pasta_saida = NULL
  )
  
  erro <- function(msg) showNotification(msg, type = "error", duration = 10)
  ok   <- function(msg) showNotification(msg, type = "message")
  
  # Lê a versão editada da tabela. Usa a inicial se ainda não houve edição ou se
  # o navegador ainda guarda uma tabela antiga (chaves diferentes)
  ler_hot <- function(input_hot, inicial, chave) {
    if (is.null(input_hot)) return(inicial)
    x <- tryCatch(hot_to_r(input_hot), error = function(e) NULL)
    if (is.null(x) || nrow(x) != nrow(inicial) || !setequal(x[[chave]], inicial[[chave]])) return(inicial)
    x
  }
  
  # ---------------------------------------------------------------------------
  # ABA 1 — USO DO SOLO
  # ---------------------------------------------------------------------------
  observeEvent(input$btn_uso, {
    caminho <- escolher_arquivos("Selecione o shapefile de Uso do Solo geral",
                                 matrix(c("Shapefiles", "*.shp"), 1, 2))
    if (length(caminho)) rv$caminho_uso <- caminho
  })
  
  observeEvent(list(rv$caminho_uso, input$enc_uso), {
    req(rv$caminho_uso)
    uso <- tryCatch(ler_camada(rv$caminho_uso, input$enc_uso),
                    error = function(e) { erro(conditionMessage(e)); NULL })
    req(uso)
    rv$uso        <- uso
    rv$uso_versao <- rv$uso_versao + 1
    campos <- campos_atributos(uso)
    updateSelectInput(session, "campo_classe", choices = campos,
                      selected = if ("Classe" %in% campos) "Classe" else campos[1])
  })
  
  gerar_tabelas_classes <- function() {
    req(rv$uso, input$campo_classe, input$campo_classe %in% names(rv$uso))
    tab <- montar_tabela_classes(rv$uso, input$campo_classe, rv$dic$classes, rv$dic$valores)
    rv$classes_ini <- tab
    rv$valores_ini <- montar_tabela_valores(tab$Classe_padronizada, NULL, rv$dic$valores)
  }
  
  observeEvent(list(rv$uso_versao, input$campo_classe), gerar_tabelas_classes())
  
  output$info_uso <- renderText({
    req(rv$uso)
    cls <- if (isTruthy(input$campo_classe) && input$campo_classe %in% names(rv$uso))
      as.character(rv$uso[[input$campo_classe]]) else character()
    paste0(
      "Arquivo: ", basename(rv$caminho_uso), "\n",
      "Feições: ", nrow(rv$uso), "\n",
      "CRS: EPSG ", st_crs(rv$uso)$epsg,
      if (isTRUE(attr(rv$uso, "reprojetado"))) " (reprojetado de graus)" else "", "\n",
      if (attr(rv$uso, "n_invalidas") > 0) paste0("Geometrias corrigidas: ", attr(rv$uso, "n_invalidas"), "\n") else "",
      if (tem_mojibake(cls)) "\n⚠ Acentos ilegíveis (ex.: 'Ã§').\nTroque a codificação do .dbf." else ""
    )
  })
  
  output$tab_classes <- renderRHandsontable({
    req(rv$classes_ini)
    opcoes <- sort(unique(c(rv$classes_ini$Classe_padronizada, rv$dic$valores$Classe_padronizada)))
    rhandsontable(rv$classes_ini, rowHeaders = NULL, stretchH = "all", height = 350) %>%
      hot_col(c("Classe_original", "Feicoes", "Area_ha", "Sugestao"), readOnly = TRUE) %>%
      hot_col("Classe_padronizada", type = "autocomplete", source = opcoes, strict = FALSE) %>%
      hot_cols(columnSorting = TRUE)
  })
  
  observeEvent(input$btn_atualizar_valores, {
    req(rv$classes_ini)
    tab_cls  <- ler_hot(input$tab_classes, rv$classes_ini, "Classe_original")
    tab_vals <- ler_hot(input$tab_valores, rv$valores_ini, "Classe_padronizada")
    tab_cls$Classe_padronizada <- trimws(tab_cls$Classe_padronizada)
    if (any(is.na(tab_cls$Classe_padronizada) | tab_cls$Classe_padronizada == "")) {
      return(erro("Há classes sem nome padronizado."))
    }
    rv$classes_ini <- tab_cls
    rv$valores_ini <- montar_tabela_valores(tab_cls$Classe_padronizada, tab_vals, rv$dic$valores)
    ok("Tabela de valores atualizada.")
  })
  
  output$tab_valores <- renderRHandsontable({
    req(rv$valores_ini)
    rhandsontable(rv$valores_ini, rowHeaders = NULL, stretchH = "all") %>%
      hot_col("Classe_padronizada", readOnly = TRUE) %>%
      hot_col("CN", format = "0.00") %>%
      hot_col("C", format = "0.000")
  })
  
  observeEvent(input$btn_carregar_dic, {
    caminho <- escolher_arquivos("Selecione o dicionário", matrix(c("Excel", "*.xlsx"), 1, 2))
    req(length(caminho))
    dic <- tryCatch(ler_dicionario(caminho), error = function(e) { erro(conditionMessage(e)); NULL })
    req(dic)
    if (is.null(dic$valores)) dic$valores <- VALORES_PADRAO
    rv$dic <- dic
    if (!is.null(rv$uso)) gerar_tabelas_classes()
    if (!is.null(rv$ads)) rv$ads_tab_ini <- montar_tabela_ads(rv$ads, rv$dic$ads)
    ok("Dicionário carregado.")
  })
  
  observeEvent(input$btn_salvar_dic, {
    req(rv$classes_ini)
    caminho <- tclvalue(tkgetSaveFile(title = "Salvar dicionário", defaultextension = ".xlsx",
                                      initialfile = "dicionario_uso_solo.xlsx",
                                      filetypes = "{{Excel} {.xlsx}}"))
    req(nzchar(caminho))
    tab_ads <- if (!is.null(rv$ads_tab_ini)) ler_hot(input$tab_ads, rv$ads_tab_ini, "id_ad") else NULL
    tryCatch({
      salvar_dicionario(caminho, ler_hot(input$tab_classes, rv$classes_ini, "Classe_original"),
                        ler_hot(input$tab_valores, rv$valores_ini, "Classe_padronizada"), tab_ads, rv$dic)
      ok(paste("Dicionário salvo:", basename(caminho)))
    }, error = function(e) erro(conditionMessage(e)))
  })
  
  # ---------------------------------------------------------------------------
  # ABA 2 — ÁREAS DE DRENAGEM
  # ---------------------------------------------------------------------------
  observeEvent(input$btn_ads, {
    caminhos <- escolher_arquivos("Selecione os shapefiles das Áreas de Drenagem",
                                  matrix(c("Shapefiles", "*.shp"), 1, 2), multi = TRUE)
    req(length(caminhos))
    raw <- tryCatch(setNames(lapply(caminhos, ler_camada, encoding = input$enc_ads), caminhos),
                    error = function(e) { erro(conditionMessage(e)); NULL })
    req(raw)
    rv$ads_raw <- raw
    campos <- Reduce(intersect, lapply(raw, campos_atributos))
    updateSelectInput(session, "campo_ad", choices = campos)
    ok(paste(length(raw), "shapefile(s) lido(s). Clique em 'Montar tabela de ADs'."))
  })
  
  observeEvent(input$btn_montar_ads, {
    req(rv$ads_raw)
    if (input$modo_ad == "feicao" && !isTruthy(input$campo_ad)) {
      return(erro("Os shapefiles não têm um campo em comum para o nome da AD."))
    }
    ads <- tryCatch(preparar_ads(rv$ads_raw, input$modo_ad, input$campo_ad),
                    error = function(e) { erro(conditionMessage(e)); NULL })
    req(ads)
    rv$ads         <- ads
    rv$ads_versao  <- rv$ads_versao + 1
    rv$ads_tab_ini <- montar_tabela_ads(ads, rv$dic$ads)
  })
  
  output$tab_ads <- renderRHandsontable({
    req(rv$ads_tab_ini)
    rhandsontable(rv$ads_tab_ini, rowHeaders = NULL, stretchH = "all", height = 450) %>%
      hot_col(c("id_ad", "Arquivo", "Nome_original", "Area_ha"), readOnly = TRUE)
  })
  
  output$info_ads <- renderText({
    req(rv$ads)
    paste0("Shapefiles: ", length(rv$ads_raw), "\n",
           "ADs: ", nrow(rv$ads), "\n",
           "Área total: ", format(round(sum(as.numeric(st_area(rv$ads))) / 1e4, 2),
                                  big.mark = ".", decimal.mark = ","), " ha")
  })
  
  # ---------------------------------------------------------------------------
  # ABA 3 — CÁLCULO
  # ---------------------------------------------------------------------------
  observeEvent(input$btn_calcular, {
    if (is.null(rv$uso) || is.null(rv$classes_ini)) return(erro("Carregue o uso do solo (aba 1)."))
    if (is.null(rv$ads)) return(erro("Monte a tabela de ADs (aba 2)."))
    
    tab_cls  <- ler_hot(input$tab_classes, rv$classes_ini, "Classe_original")
    tab_vals <- ler_hot(input$tab_valores, rv$valores_ini, "Classe_padronizada")
    tab_ads  <- ler_hot(input$tab_ads, rv$ads_tab_ini, "id_ad")
    tab_cls$Classe_padronizada <- trimws(tab_cls$Classe_padronizada)
    
    # Validações
    sem_tabela <- setdiff(unique(tab_cls$Classe_padronizada), tab_vals$Classe_padronizada)
    if (length(sem_tabela)) {
      return(erro(paste0("Classes fora da Tabela 2: ", paste(sem_tabela, collapse = ", "),
                         ". Clique em 'Atualizar tabela de valores'.")))
    }
    tab_ads$AD <- trimws(tab_ads$AD)
    erros_ad <- validar_nomes_ad(tab_ads$AD)
    if (length(erros_ad)) return(erro(paste(erros_ad, collapse = "\n")))
    
    res <- withProgress(message = "Calculando...", value = 0, {
      # Interseção em cache: só refaz se o uso do solo, o campo ou as ADs mudaram
      chave_cache <- c(rv$uso_versao, rv$ads_versao, input$campo_classe)
      if (is.null(rv$cache) || !identical(rv$cache$chave, chave_cache)) {
        incProgress(0.2, detail = "Recortando uso do solo pelas ADs")
        inter <- tryCatch(intersectar(rv$uso, input$campo_classe, rv$ads),
                          error = function(e) { erro(paste("Erro no recorte:", conditionMessage(e))); NULL })
        if (!is.null(inter)) rv$cache <- list(chave = chave_cache, inter = inter)
      }
      
      if (is.null(rv$cache) || !identical(rv$cache$chave, chave_cache)) {
        NULL
      } else {
        incProgress(0.5, detail = "Ponderando")
        mapa_classes <- setNames(tab_cls$Classe_padronizada, tab_cls$Classe_original)
        tryCatch(ponderar_runoff(rv$cache$inter, rv$ads, mapa_classes, tab_vals, input$metodo),
                 error = function(e) { erro(conditionMessage(e)); NULL })
      }
    })
    if (is.null(res)) return()
    
    # Classes presentes nas ADs sem valor para o método escolhido
    sem_valor <- unique(res$tabela$Classe[res$tabela$area_m2 > 0 & is.na(res$tabela$Valor)])
    if (length(sem_valor)) {
      return(erro(paste0("Sem valor de ", input$metodo, " para: ", paste(sem_valor, collapse = ", "))))
    }
    
    rv$res      <- res
    rv$nomes_ad <- tab_ads[, c("id_ad", "Dispositivo", "AD")]
    rv$tab_classes_calc <- tab_cls
    rv$tab_valores_calc <- tab_vals
    updateSelectInput(session, "ad_detalhe", choices = setNames(tab_ads$id_ad, tab_ads$AD))
    ok("Cálculo concluído.")
  })
  
  resumo_exibir <- reactive({
    req(rv$res, rv$nomes_ad)
    rv$res$resumo %>%
      left_join(rv$nomes_ad, by = "id_ad") %>%
      arrange(id_ad) %>%
      transmute(Dispositivo, AD,
                `Área (ha)` = Area_AD_m2 / 1e4,
                `Cobertura (%)` = 100 * Cobertura,
                !!rv$res$metodo := Coeficiente)
  })
  
  output$tab_resumo <- renderTable(resumo_exibir(), digits = 3, na = "—")
  
  output$tab_detalhe <- renderTable({
    req(rv$res, input$ad_detalhe)
    rv$res$tabela %>%
      filter(id_ad == as.integer(input$ad_detalhe), area_m2 > 0) %>%
      transmute(Classe, `Área (ha)` = area_m2 / 1e4, `%` = 100 * Perc,
                !!rv$res$metodo := Valor)
  }, digits = 3)
  
  output$info_resultado <- renderText({
    req(rv$res, rv$nomes_ad)
    r <- rv$res$resumo %>% left_join(rv$nomes_ad, by = "id_ad")
    baixa <- r$AD[!is.na(r$Cobertura) & r$Cobertura < 0.99]
    vazia <- r$AD[r$Area_uso_m2 == 0]
    paste0(
      "ADs calculadas: ", nrow(r), "\n",
      if (length(vazia)) paste0("\n⚠ Sem uso do solo:\n  ", paste(vazia, collapse = "\n  "), "\n") else "",
      if (length(baixa)) paste0("\n⚠ Cobertura < 99%:\n  ",
                                paste(sprintf("%s (%.1f%%)", r$AD[r$AD %in% baixa],
                                              100 * r$Cobertura[r$AD %in% baixa]), collapse = "\n  ")) else ""
    )
  })
  
  output$mapa <- renderLeaflet({
    req(rv$res, rv$nomes_ad)
    geom <- st_transform(rv$res$geom, 4326)
    ads  <- st_transform(rv$ads, 4326) %>% left_join(rv$nomes_ad, by = "id_ad")
    pal  <- colorFactor("Set2", domain = geom$Classe)
    leaflet() %>%
      addProviderTiles(providers$Esri.WorldImagery, group = "Satélite") %>%
      addProviderTiles(providers$OpenStreetMap, group = "Mapa") %>%
      addPolygons(data = geom, fillColor = ~pal(Classe), fillOpacity = 0.6,
                  color = "#333333", weight = 0.5, group = "Uso do solo",
                  label = ~sprintf("%s — %.2f ha", Classe, area_m2 / 1e4)) %>%
      addPolygons(data = ads, fill = FALSE, color = "#FF0000", weight = 2, group = "ADs",
                  label = ~AD) %>%
      addLegend(pal = pal, values = geom$Classe, title = "Classe", position = "bottomright") %>%
      addLayersControl(baseGroups = c("Satélite", "Mapa"), overlayGroups = c("Uso do solo", "ADs"))
  })
  
  # ---------------------------------------------------------------------------
  # EXPORTAÇÃO
  # ---------------------------------------------------------------------------
  observeEvent(input$btn_pasta, {
    pasta <- tk_choose.dir(caption = "Selecione a pasta de saída")
    if (!is.na(pasta) && nzchar(pasta)) rv$pasta_saida <- pasta
  })
  
  output$info_pasta <- renderText({
    if (is.null(rv$pasta_saida)) "Nenhuma pasta selecionada" else rv$pasta_saida
  })
  
  observeEvent(input$btn_exportar, {
    if (is.null(rv$res)) return(erro("Calcule antes de exportar."))
    if (is.null(rv$pasta_saida)) return(erro("Selecione a pasta de saída."))
    
    # Usa os nomes atuais da tabela de ADs (renomear não exige recalcular)
    tab_ads <- ler_hot(input$tab_ads, rv$ads_tab_ini, "id_ad")
    tab_ads$AD <- trimws(tab_ads$AD)
    erros_ad <- validar_nomes_ad(tab_ads$AD)
    if (length(erros_ad)) return(erro(paste(erros_ad, collapse = "\n")))
    rv$nomes_ad <- tab_ads[, c("id_ad", "Dispositivo", "AD")]
    
    tryCatch({
      withProgress(message = "Exportando...", {
        timestamp <- format(Sys.time(), "%Y%m%d_%H%M%S")
        xlsx <- file.path(rv$pasta_saida, paste0("ponderacao_", rv$res$metodo, "_por_AD_", timestamp, ".xlsx"))
        exportar_excel(rv$res, rv$nomes_ad, rv$tab_classes_calc, rv$tab_valores_calc, xlsx)
        n_shp <- 0
        if (input$exp_shp) n_shp <- length(exportar_shapes(rv$res, rv$nomes_ad, rv$pasta_saida))
      })
      ok(paste0("Excel salvo: ", basename(xlsx),
                if (n_shp) paste0(" | ", n_shp, " shapefile(s)") else ""))
    }, error = function(e) erro(paste("Erro na exportação:", conditionMessage(e))))
  })
}

# =============================================================================
# 6. EXECUÇÃO
# =============================================================================
runApp(shinyApp(ui, server), launch.browser = TRUE)