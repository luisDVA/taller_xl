# Taller 'Importando datos desde hojas de cálculo'

Este repositorio contiene los archivos .xlsx para los ejercicios y demostraciones del taller, junto con un archivo .R con código para los ejemplos que vamos a estar revisando.

El código que vamos a usar contiene rutas de archivo relativas, por lo que recomiendo descargar los archivos de este repositorio y copiarlos dentro de la carpeta de un proyecto de RStudio nuevo, creado para este taller, o bien en la carpeta de algún otro projecto existente. 


Podemos descargar los datos desde R con el paquete `usethis`

``` r
# install.packages("usethis")
usethis::use_course("luisDVA/taller_xl")
```

Este es un ejemplo de cómo sería la estructura de las carpetas y archivos:
(el directorio de trabajo y el archivo .RProj pueden llevar cualquier nombre que nos sirva)  

directorio_de_trabajo/  
├─ hojas_calc/  
│  ├─ cafeteria.xlsx  
│  ├─ rladies.xlsx  
│  ├─ rladies-formato.xlsx  
│  ├─ vehiculos.xlsx  
├─ taller_calc.R  
├─ proyecto_taller.Rproj  

