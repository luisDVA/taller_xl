##%######################################################%##
#                                                          #
####              Taller Hojas de Cálculo               ####
#                                                          #
##%######################################################%##


# Cargar paquetes ---------------------------------------------------------
library(readxl)  
library(purrr)   
library(unheadr)
library(dplyr)   
library(tidyxl)  
library(writexl)
library(googlesheets4) 

# Importar con readxl -----------------------------------------------------
read_excel(path = "hojas_calc/rladies.xlsx")

# ejemplo del comportamiento 'discover'
read_excel(path = "hojas_calc/cafeteria.xlsx", sheet = 4)
read_excel(path = "hojas_calc/cafeteria.xlsx", sheet = "Combos")
read_excel(path = "hojas_calc/cafeteria.xlsx", sheet = "Combos",range = "C7:D13")

# Hojas y pestañas --------------------------------------------------------
ruta_cafeteria <- "hojas_calc/cafeteria.xlsx"
ruta_cafeteria %>%
  excel_sheets()

read_excel(ruta_cafeteria,sheet = "Bebidas")  
read_excel(ruta_cafeteria,sheet = "1") # ojo aquí  
read_excel(ruta_cafeteria,sheet =  1)   
read_excel(ruta_cafeteria,sheet =  1:2) #ojo


## muchas pestañas 
ruta_vehiculos <- "hojas_calc/vehiculos.xlsx"

# para obtener una lista de tibbles 
nombres_hojas <- ruta_vehiculos %>%
                    excel_sheets() %>% # obtener los nombres de las hojas
                    set_names()        # asignarle los nombres a cada elemento del vector 
# map(nombres_hojas,read_excel,path = ruta_vehiculos)

lista_fabricantes <- map(nombres_hojas,
    ~read_excel(sheet=.x,path = ruta_vehiculos)) 
lista_fabricantes[[1]]
lista_fabricantes[[2]]

# asignar a tibbles en el entorno global
list2env(lista_fabricantes,envir = .GlobalEnv)

# una sola tibble
vehiculos_todo <- bind_rows(lista_fabricantes)


# Hojas con formato -------------------------------------------------------
ruta_rladies <- "hojas_calc/rladies.xlsx"
ruta_formato <- "hojas_calc/rladies-formato.xlsx"

read_xlsx(ruta_rladies)
read_xlsx(ruta_formato)

# desenmascarar celdas
celdas <-  tidyxl::xlsx_cells(ruta_rladies)
celdasformato <- tidyxl::xlsx_cells(ruta_formato)

# desenmascarar formatos
sinformato <- tidyxl::xlsx_formats(ruta_rladies)
conformato <- tidyxl::xlsx_formats(ruta_formato)

## De la ayuda de tidyxl
# negritas
conformato$local$font$bold[celdasformato$local_format_id]
conformato$local$font$bold[celdasformato$style_format]
negritas <- which(conformato$local$font$bold)
negritas
celdasformato[celdasformato$local_format_id %in% negritas, ] # filtrar

# formato con unheadr -----------------------------------------------------
ruta_formato <- "hojas_calc/rladies-formato.xlsx"
ruta_formato
read_excel(ruta_formato)

annotate_mf(ruta_formato,orig = Grupo, new=Grupo_fmt)


# exportando --------------------------------------------------------------
excel_sheets("hojas_calc/cafeteria.xlsx")
panaderia <- read_excel("hojas_calc/cafeteria.xlsx",sheet = 2)
panaderia2 <- panaderia %>% mutate(precio_usd=Precio*21)

write_xlsx(panaderia2,path = "panaderia.xlsx")
# con nombre de hoja
write_xlsx(list(Panaderia=panaderia2),path = "panaderia.xlsx")

# grupos, una hoja para cada grupo
# para obtener una lista de tibbles 
ruta_vehiculos <- "hojas_calc/vehiculos.xlsx"
nombres_hojas <- ruta_vehiculos %>%
  excel_sheets() %>% # obtener los nombres de las hojas
  set_names()        # asignarle los nombres a cada elemento del vector 
lista_fabricantes <- map(nombres_hojas,
                         ~read_excel(sheet=.x,path = ruta_vehiculos)) 
# cambiamos los nombres para ver que sí funciona
names(lista_fabricantes) <- paste("Hoja",names(lista_fabricantes))
names(lista_fabricantes)
# exportar
write_xlsx(lista_fabricantes,"vehiculos2.xlsx")

# exportar varios archivos
vehic3 <- bind_rows(lista_fabricantes)
vehic3
vehic3 %>% 
   group_by(traccion) %>% 
   group_walk(~write_xlsx(.x,path = file.path(paste0(.x$traccion,".xlsx"))),.keep=TRUE)
   

# google sheets -----------------------------------------------------------
# link https://docs.google.com/spreadsheets/d/1QTwAa5-uod6JNABo_ua_T8y98Enf6Hjjzde6PgUp2Ig/edit?usp=sharing
gs4_deauth()
read_sheet("1QTwAa5-uod6JNABo_ua_T8y98Enf6Hjjzde6PgUp2Ig")
# abrir en el navegador
gs4_browse("1QTwAa5-uod6JNABo_ua_T8y98Enf6Hjjzde6PgUp2Ig")
# mejor forma de identificar nuestro documento
ssmx <- "1QTwAa5-uod6JNABo_ua_T8y98Enf6Hjjzde6PgUp2Ig"
write_sheet(iris,ss = ssmx,sheet = "nueva")
