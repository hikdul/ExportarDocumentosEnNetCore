
[enlace]:https://itextpdf.com/es/products/itext-7/ "IText 7"

# Generar Docmento en PDF

Estos documentos son igual o mas complicados que los de word y rara ves llegan a ser necesario pero en caso de. se puede usar el nugget *IText 7* en el cual se puede obtener mayor detaller de como usar mediante el siguiente [enlace]

## Plug Necesario para Generar el excel

Para generar los excel con poco esfuerso se utiliza el NuGet ** IText 7 ** 
pues con este plug con nos deja un camino bastante sencillo para generar los archivos.
Pagina del plug visita el siguiente [enlace]

##  nota
 recuende ir a la rutar /Explicando todo para obtener los enlaces a los nuggets de este proyecto.

 ### Como Generar un documento en PDF

 primero afirmar que es un poco engorroso. pero explicare los mas breve posible para hacer entender, pero primero recomiendo hacer el llamado a varios namespaces, pue la verdad a veces el mismo Visual Studio o VScode no los consigue.

 ```C#

        using iText.IO.Font.Constants;

        using iText.Kernel.Font;

        using iText.Kernel.Pdf;

        using iText.Layout;

        using iText.Layout.Element;
 ```

 luego generamos nuestro Memory Stream, alli creamos un pdfWriter que es donde sobreponemos nuestro memory stream y por medio de una directiva using creamos el documento pdf

 ```C#
 // creamos el memory strueam
 using (MemoryStream ms = new MemoryStream())
                {
                    // luego el writer
                    PdfWriter writer = new(ms);
                    // y por ultimo nuestro documento
                    using (PdfDocument document = new(writer))
                    {
                        //var font = PdfFontFactory.CreateFont("c:/windows/fonts/arial.ttf", iText.IO.Font.PdfEncodings.IDENTITY_H);
                        var fontSize = 10f;

                       // doc.SetFont(font);
                        doc.SetFontSize(fontSize);
                        ...
                        ...
                        ...
                    
```
ya luego empezamos a introducir los valores dentro del documento, esto se encuentra mas detallado dentro del codigo de este proyecto