[enlace]:https://www.nuget.org/packages/EPPlus/ "EPPlus"

# Como Generar un documento en excel
en este documento se explicara como genero mis documentos en excel, en general esto funciona para los reportes que en su mayoria es bantante comun que la pidan en las aplicaciones

## Plug Necesario para Generar el excel

Para generar los excel con poco esfuerso se utiliza el NuGet ** EPPLUS ** 
pues con este plug con nos deja un camino bastante sencillo para generar los archivos.
Pagina del plug visita el siguiente [enlace]


## Pasos para generar el Archivo excel

. preparar el plug y los datos que se deseen exportar.
2. Generar el array de bytes __ bytes[] __; Para esto se puede tomar como ejemplo el siguiente codigo
```C#
            //se genera un memorytStream para crear el documento
            using (MemoryStream ms = new MemoryStream())
            {
                //se agrega el tipo de licencia que va a usar nuestro documento
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                //aqui partimos con el uso de EPPlus y generamos un paquete de excel
                using (ExcelPackage ep = new ExcelPackage())
                {
                    //Generamos una hoja... pueden generarce N hojas
                    ep.Workbook.Worksheets.Add("Titula de la hoja");
                    //se asigna la primera pagina
                    ExcelWorksheet ew = ep.Workbook.Worksheets[0];
                    // generanto stylos basicos
                    ew.Cells.Style.Font.Size = 10;
                    ew.Cells.Style.Font.Name = "Arial";
                    //se empieza a introducir los documentos
                    ew.Cells[1, 1].Value = "Reporte Generado ";
                    ew.Cells[1, 2].Value = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                    ...
                    ...
                    ...
                    // almaceno el memory string en el excel package, asi se genera mi archivo de excel
                    ep.SaveAs(ms);
                    // retorna el array de bytes y ya tengo mi documento para el uso que lo deseo
                    return ms.ToArray();
                }
            }
```
3. ahora transformamos nuestro array de bytes a un archivo, aqui les damos extencion y demas y ya tenemos nuestro archivo listo
```
 //buffer es la variable que recibe nuestro byte[]
 var buffer= ...
 // de este modo enviamos un archivo a las vistas...
 return File(buffer, "application/vnd.ms-excel", "nombreDelDocumento.xlsx");
```

