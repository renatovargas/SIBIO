# Llamar librerías
library(readxl)
library(openxlsx)
library(reshape2)
library(stringr)
library(plyr)

# Limpiar el área de trabajo
rm(list = ls())
# Se pone la ruta de trabajo en una variable (con "/")

# País
# Ecuador = ECU
iso <- "ECU"

# Lógica recursiva
# ================

# info es lo que resulta de leer el archivo
archivo <- "datos/COU_2014-2019_PRECIOSCORRIENTES_66x61_REFERENCIA.xlsx"
hojas <- excel_sheets(archivo)
# extraemos solo las que nos interesan
#crea objetos con nombre de las hoajs de excel.
hojas <- hojas[-c(1,3,5,7,9,11,13,14,15,16,17,18,19,20,21,22,23,24,25)]

lista <- c("inicio")
i<- 1
for (i in 1:length(hojas)) {
  # Extraemos el año y la unidad de medida
  info <- read_excel(
    archivo,
    range = paste("'", hojas[i], "'!a4:a5", sep = "") ,
    col_names = FALSE,
    col_types = "text",
  )
  
  # Extraemos el texto de la cadena de caracteres
  anio <- as.numeric((str_extract(info[2,], "\\d{4}")))
  unidad <- toString(info[1,])
  # unidad de medida
  
  # precios Corrientes o medidas Encadenados
  if (unidad != "Miles de millones de pesos") {
    precios <- "Encadenados"
    #Dejé los quetzales, porque no me modificaba
    unidad <-
      c("Millones de quetzales en medidas encadenadas de volumen con año de referencia 2013")
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
  oferta[is.na(oferta)] <- 0.0
  
  # Los correlativos deberían tener el código ISO en la forma "ISO3of001"
  # o "ISO3oc001"
  
  rownames(oferta) <- c(sprintf(paste(iso, "of%03d", sep = ""), seq(1, dim(oferta)[1])))
  colnames(oferta) <- c(sprintf(paste(iso, "oc%03d", sep = ""), seq(1, dim(oferta)[2])))
  
  # Columnas a eliminar con subtotales y totales
  # Total oferta a precios comprador (1)
  # Márgenes de comercio y de transporte siempre deben quedar.
  # LOS IMPUESTOS Y SUBVENCIONES TAMBIÉN DEBEN QUEDAR
  # Oferta total precios básicos (8)
  # La 70 la tengo vacía, debo sacarla
  # Producción a precios básicos; TOTAL (71)
  # Producción a precios básicos; Para uso final propio (72)
  # Producción a precios básicos; Otra de no mercado (73)
  # Producción a precios básicos; De mercado (74)
  #  (75)
  # Las importaciones quedan también
  #Bienes y servicios están FOB, LOS AJUSTES CIF/FOB AGREGAN EL COSTO ADICIONAL DE TRAERLO.
  #Nótese que no estamos eliminando filas, pero siempre hay que asegurarse que sea el caso.
  oferta1 <- oferta[,-c(1, 8, 70, 71, 72, 73, 74, 75)]
  #IMPORTANTE, VERIFICAR QUE CUADRE EL CUADRO Y QUE OFERTA SEA IGUAL A UTILIZACIÓN, CASO CONTRARIO, MODIFICAR LA VARIACIÓN DE EXISTENCIAS CON LA DIFERENCIA.
  # Desdoblamos
  oferta <- cbind(anio, precios, 1, "Oferta", melt(oferta1), unidad)
  
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
    range = paste("'" , hojas[i], "'!c97:eh164", sep = ""),
    # Nótese que no incluimos la fila de totales
    col_names = FALSE,
    col_types = "numeric"
  ))
  utilizacion[is.na(utilizacion)] <- 0.0
  rownames(utilizacion) <-
    c(sprintf("uf%03d", seq(1, dim(utilizacion)[1])))
  colnames(utilizacion) <-
    c(sprintf("uc%03d", seq(1, dim(utilizacion)[2])))
  
  #  Columnas a eliminar con subtotales y totales
  #  Total oferta a precios de comprador (1)
  #  Total de consumo intermedio (65)
  #  Vacía (66)
  #  Sacar Consumo Intermedio a precios de comprador (67)
  #  Desde la 68 a 76, tampoco me interesan  ()
  #  Dejamos la 77 y luego quitamos desde la 78 hasta la 84.
  #  Dejamos la 85 y luego quitamos de la 86 a la 92
  #  Dejamos la 93 y la 94 y quitamos de la 95 a la 103.
  #  Dejamos la 104 y quitamos desde la 105 hasta la 111
  #  Dejamos la 112 y quitamos desde 113 hasta la 119
  #  Dejamos la 120 y quitamos desde 121 hasta la 127
  #  Dejamos la 128 y la 129 y quitamos desde 130 hasta la 136
  #  Poner la línea de o y nas as.matrix(rowSums(oferta1))-
  #  as.matrix(rowSums(utilziacion1), esto debiera dar 0 ; 
  # utilziacion 1 debiera dar igual a la primera columna del excel.
  
  utilizacion1 <-
    utilizacion[,-c(1,
                    65:76,
                    78:84,
                    86:92,
                    95:103,
                    105:111,
                    113:119,
                    121:127,
                    130:136)]
  as.matrix(rowSums(utilizacion1))
  #A continuación, va a surgir una diferencia que entiendo,
  #se debe a que en el excel, en la fila (Compras directas en el exterior por residentes),
  #se totaliza a la  suma ente gasto de los hogares a precios básicos y Compras directas en el territorio nacional por no residentes, es así?
  #Importante, verificar que las dimensiones de oferta y utilización sean
  as.matrix(rowSums(oferta1)) - as.matrix(rowSums(utilizacion1))
  # Desdoblamos
  utilizacion <-
    cbind(anio,
          precios,
          2,
          "Utilización",
          melt(utilizacion1),
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
  
  # Cuadro de Valor Agregado
  # ========================
  
  valorAgregado <- as.data.frame(read_excel(
    archivo,
    range = paste("'" , hojas[i], "'!c167:bn167", sep = ""),
    col_names = FALSE,
    col_types = "numeric"
  ))
  rownames(valorAgregado) <-
    c(sprintf("vf%03d", seq(1, dim(valorAgregado)[1])))
  colnames(valorAgregado) <-
    c(sprintf("vc%03d", seq(1, dim(valorAgregado)[2])))
  valorAgregado[is.na(valorAgregado)] <- 0.0
  
  #   Columnas a eliminar con subtotales y totales
  
  #   vc093	SUBTOTAL DE MERCADO
  #   vc098	SUBTOTAL USO FINAL PROPIO
  #   vc108	SUBTOTAL NO DE MERCADO
  #   vc109 TOTAL
  
  valorAgregado <- valorAgregado[,-c(1, 2, 3)]
  
  # Desdoblamos
  valorAgregado <-
    cbind(anio,
          precios,
          3,
          "Valor Agregado",
          "vf001",
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
  #   empleo <- as.data.frame(read_excel(
  #     archivo,
  #     range = paste("'" , hojas[i], "'!D332:DH332", sep = ""),
  #     col_names = FALSE,
  #     col_types = "numeric"
  #   ))
  #   rownames(empleo) <- c(sprintf("ef%03d", seq(1, dim(empleo)[1])))
  #   colnames(empleo) <- c(sprintf("ec%03d", seq(1, dim(empleo)[2])))
  #
  #   #   Columnas a eliminar con subtotales y totales
  #
  #   #   vc093	SUBTOTAL DE MERCADO
  #   #   vc098	SUBTOTAL USO FINAL PROPIO
  #   #   vc108	SUBTOTAL NO DE MERCADO
  #   #   vc109 TOTAL
  #
  #   empleo <- empleo[, -c(93, 98, 108, 109)]
  #
  #   #Desdoblamos
  #   empleo <- cbind(anio,
  #                   precios,
  #                   4,
  #                   "Empleo",
  #                   "ef001",
  #                   melt(empleo),
  #                   "Puestos de trabajo")
  #
  #   colnames(empleo) <- c("Año",
  #                         "Precios",
  #                         "No. Cuadro",
  #                         "Cuadro",
  #                         "Filas",
  #                         "Columnas",
  #                         "Valor",
  #                         "Unidades")
  #
  #
  # # Unimos todas las partes
  # if (precios == "Corrientes") {
  #   union <- rbind(oferta,
  #                  utilizacion,
  #                  valorAgregado,
  #                  empleo)
  #
  union <- rbind(oferta, utilizacion, valorAgregado)
  assign(paste("COU_", anio, "_", precios, sep = ""),
         union)
  lista <- c(lista, paste("COU_", anio, "_", precios, sep = ""))
}

# Actualizamos nuestra lista de objetos creados
lista <- lapply(lista[-1], as.name)

# Unimos los objetos de todos los años y precios
SCN <- do.call(rbind.data.frame, lista)

# Y borramos los objetos individuales
do.call(rm,lista)



# Le damos significado a las filas y columnas

clasificacionColumnas <- read_xlsx("COL_Clasificaciones.xlsx",
                                   sheet = "columnas",
                                   col_names = TRUE,)

clasificacionFilas <- read_xlsx("COL_Clasificaciones.xlsx",
                                sheet = "filas",
                                col_names = TRUE,)

#SCN <- join(SCN,clasificacionColumnas,by = "Columnas")
#SCN <- join(SCN,clasificacionFilas, by = "Filas")
gc()

# Y lo exportamos a Excel
write.xlsx(
  SCN,
  "salidas/COL_SCN_BD.xlsx",
  sheetName= "COL_SCN_BD",
  rowNames=FALSE,
  colnames=FALSE,
  overwrite = TRUE,
  asTable = FALSE
)
