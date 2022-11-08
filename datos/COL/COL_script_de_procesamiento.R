# Llamar librerías
library(readxl)
library(openxlsx)
library(reshape2)
library(stringr)
library(plyr)

# Limpiar el área de trabajo
rm(list = ls())
# Se pone la ruta de trabajo en una variable (con "/")
wd <- "C:/Users/amondaini/Desktop/Unidad de Desarrollo Agrícola/Cuenta satelite Bioeconomía/datos/colombia"
# Cambiar la ruta de trabajo con la variable anterior
setwd(wd)
# Lógica recursiva
# ================
# info es lo que resulta de leer el archivo
# info es lo que resulta de leer el archivo
# info es lo que resulta de leer el archivo
# info es lo que resulta de leer el archivo
archivo <- "COU_2014-2019_PRECIOSCORRIENTES_66x61_REFERENCIA.xlsx"
hojas <- excel_sheets(archivo)
# extraemos solo las que nos interesan
#crea objetos con nombre de las hoajs de excel.
hojas <- hojas[-c(1,3,5,7,9,11,13,14,15,16,17,18,19,20,21,22,23,24,25)]
#el inicio es solo de ejemplo, puede ir cualk cosa.
lista <- c("inicio")

for (i in 1:length(hojas)) {
  # Extraemos el año y la unidad de medida
  info <- read_excel(
    archivo,
    range = paste("'", hojas[i], "'!a4:a5", sep = "") ,
    col_names = FALSE,
    col_types = "text",
  )

  # Extraemos el texto de la cadena de caracteres
  anio <- as.numeric((str_extract(info[2, ], "\\d{4}")))
  unidad <- toString(info[1, ]) 
  # unidad de medida
  }
  # precios Corrientes o medidas Encadenados
  if (unidad != "Miles de millones de pesos") {
    precios <- "Encadenados"
  #Dejé los quetzales, porque no me modificaba 
    unidad <- c("Millones de quetzales en medidas encadenadas de volumen con año de referencia 2013")
  }  else {
    precios <- "Corrientes"
  }
  # Cuadro de Oferta
  # ================
  
  oferta <- as.matrix(read_excel(
    archivo,
    range = paste("'" , hojas[i], "'!c11:cb78", sep = ""),
    # Nótese que no incluimos la fila de totales
    col_names = FALSE,
    col_types = "numeric"
  ))
  rownames(oferta) <- c(sprintf("of%03d", seq(1, dim(oferta)[1])))
  colnames(oferta) <- c(sprintf("oc%03d", seq(1, dim(oferta)[2])))
  
  # Columnas a eliminar con subtotales y totales
  # Total oferta a precios comprador (1)
  # Márgenes de comercio (2)
  # Márgenes de transporte (3)
  # Impuestos y derechos a las importaciones (4)
  # IVA no deducible (5)
  # Impuestos a los productos (excepto impuestos a importaciones e IVA no deducible) (6)
  # Subvenciones a los productos (7)
  # Oferta total precios básicos (8)
  # Producción a precios básicos; TOTAL (71)
  # Producción a precios básicos; Para uso final propio (72)
  # Producción a precios básicos; Otra de no mercado (73)
  # Producción a precios básicos; De mercado (74)
  #  (75)
  # Importaciones; Ajustes  CIF/FOB sobre importaciones (76)
  # Importaciones; Bienes (77)
  # Importaciones; Servicios (78)
  oferta <- oferta[, -c(1,2,3,4,5,6,7,8,71,72,73,74,75,76)]
  
  # Desdoblamos
  oferta <- cbind(anio, precios,1, "Oferta", melt(oferta), unidad)
  
  colnames(oferta) <-
    c("Año",
      "Precios",
      "No. Cuadro",
      "Cuadro",
      "Filas",
      "Columnas",
      "Valor",
      "Unidades")
  
  # Cuadro de utilización
  # =====================
  
  utilizacion <- as.matrix(read_excel(
    archivo,
    range = paste("'" , hojas[i], "'!D169:DQ320", sep = ""),
    # Nótese que no incluimos la fila de totales
    col_names = FALSE,
    col_types = "numeric"
  ))
  rownames(utilizacion) <-
    c(sprintf("uf%03d", seq(1, dim(utilizacion)[1])))
  colnames(utilizacion) <-
    c(sprintf("uc%03d", seq(1, dim(utilizacion)[2])))
  
  #   Columnas a eliminar con subtotales y totales
  
  #   uc093	P2 CONSUMO INTERMEDIO (PC)	SUBTOTAL DE MERCADO
  #   uc098	P2 CONSUMO INTERMEDIO (PC)	SUBTOTAL USO FINAL PROPIO
  #   uc108	P2 CONSUMO INTERMEDIO (PC)	SUBTOTAL NO DE MERCADO
  #   uc109	"P2 TOTAL CONSUMO INTERMEDIO
  #   uc112	P6 EXPORTACIONES (FOB)
  #   uc115	P3 GASTO DE CONSUMO FINAL
  #   uc118	TOTAL UTILIZACIÓN
  
  utilizacion <- utilizacion[, -c(93, 98, 108, 109, 112, 115, 118)]
  
  # Desdoblamos
  utilizacion <-
    cbind(anio, 
          precios,
          2,
          "Utilización", 
          melt(utilizacion), 
          unidad)
  
  colnames(utilizacion) <-
    c("Año",
      "Precios",
      "No. Cuadro",
      "Cuadro",
      "Filas",
      "Columnas",
      "Valor",
      "Unidades")
  
  
  # Cuadros de Valor Agregado y empleo solo para precios Corrientes
  # ========================
  
  if (precios == "Corrientes") {
    # Cuadro de Valor Agregado
    # ========================
    
    valorAgregado <- as.matrix(read_excel(
      archivo,
      range = paste("'" , hojas[i], "'!D325:DH330", sep = ""),
      col_names = FALSE,
      col_types = "numeric"
    ))
    rownames(valorAgregado) <-
      c(sprintf("vf%03d", seq(1, dim(valorAgregado)[1])))
    colnames(valorAgregado) <-
      c(sprintf("vc%03d", seq(1, dim(valorAgregado)[2])))
    
    #   Columnas a eliminar con subtotales y totales
    
    #   vc093	SUBTOTAL DE MERCADO
    #   vc098	SUBTOTAL USO FINAL PROPIO
    #   vc108	SUBTOTAL NO DE MERCADO
    #   vc109 TOTAL
    
    valorAgregado <- valorAgregado[, -c(93, 98, 108, 109)]
    
    # Desdoblamos
    valorAgregado <-
      cbind(anio,
            precios,
            3,
            "Valor Agregado",
            melt(valorAgregado),
            unidad)
    
    colnames(valorAgregado) <-
      c("Año",
        "Precios",
        "No. Cuadro",
        "Cuadro",
        "Filas",
        "Columnas",
        "Valor",
        "Unidades")
    
    # Empleo
    # ======
    
    empleo <- as.data.frame(read_excel(
      archivo,
      range = paste("'" , hojas[i], "'!D332:DH332", sep = ""),
      col_names = FALSE,
      col_types = "numeric"
    ))
    rownames(empleo) <- c(sprintf("ef%03d", seq(1, dim(empleo)[1])))
    colnames(empleo) <- c(sprintf("ec%03d", seq(1, dim(empleo)[2])))
    
    #   Columnas a eliminar con subtotales y totales
    
    #   vc093	SUBTOTAL DE MERCADO
    #   vc098	SUBTOTAL USO FINAL PROPIO
    #   vc108	SUBTOTAL NO DE MERCADO
    #   vc109 TOTAL
    
    empleo <- empleo[, -c(93, 98, 108, 109)]
    
    #Desdoblamos
    empleo <- cbind(anio,
                    precios,
                    4,
                    "Empleo",
                    "ef001",
                    melt(empleo),
                    "Puestos de trabajo")
    
    colnames(empleo) <- c("Año",
                          "Precios",
                          "No. Cuadro",
                          "Cuadro",
                          "Filas",
                          "Columnas",
                          "Valor",
                          "Unidades")
    
  }
  
  # Unimos todas las partes
  if (precios == "Corrientes") {
    union <- rbind(oferta, 
                   utilizacion, 
                   valorAgregado, 
                   empleo)
    
    assign(paste("COU_", anio, "_", precios, sep = ""), 
           union)
  }
  else {
    union <- rbind(oferta, utilizacion)
    assign(paste("COU_", anio, "_", precios, sep = ""), 
           union)
  }
  lista <- c(lista, paste("COU_", anio, "_", precios, sep = ""))
}

# Actualizamos nuestra lista de objetos creados
lista <- lapply(lista[-1], as.name)

# Unimos los objetos de todos los años y precios
SCN <- do.call(rbind.data.frame, lista)

# Y borramos los objetos individuales
do.call(rm,lista)



# Le damos significado a las filas y columnas

clasificacionColumnas <- read_xlsx(
  "filas_y_columnas.xlsx",
  sheet = "columnas",
  col_names = TRUE,
)
clasificacionFilas <- read_xlsx(
  "filas_y_columnas.xlsx",
  sheet = "filas",
  col_names = TRUE,
)

SCN <- join(SCN,clasificacionColumnas,by = "Columnas")
SCN <- join(SCN,clasificacionFilas, by = "Filas")
gc()

# Y lo exportamos a Excel
write.xlsx(
  SCN,
  "SCN_BD.xlsx",
  sheetName= "SCNGT_BD",
  rowNames=FALSE,
  colnames=FALSE,
  overwrite = TRUE,
  asTable = FALSE
)

# El formato CSV se exporta muy grande, pero se comprime muy bien a 3mb
# write.csv(
#   SCN,
#   "scn.csv",
#   col.names = TRUE,
#   row.names = FALSE,
#   fileEncoding = "latin1"
# )