using iText.IO.Font.Constants;
using iText.Kernel.Font;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using ExportarDocumentos.Models;

namespace ExportarDocumentos.Helpers
{
    public class pdf : IExp
    {
        /// <summary>
        /// documento
        /// </summary>
        public Byte[] doc { get; }
        /// <summary>
        /// Empty
        /// </summary>
        public pdf()
        {
            this.doc=new byte[0];
        }
        /// <summary>
        /// el que realmente es
        /// </summary>
        /// <param name="datos"></param>
        public pdf(List<Datos> datos)
        {
            try
            {
              using (MemoryStream ms = new MemoryStream())
                {
                    PdfWriter writer = new(ms);
                    using (PdfDocument document = new(writer))
                    {
                        //var font = PdfFontFactory.CreateFont("c:/windows/fonts/arial.ttf", iText.IO.Font.PdfEncodings.IDENTITY_H);
                        var fontSize = 10f;

                        Document doc = new(document);
                       // doc.SetFont(font);
                        doc.SetFontSize(fontSize);

                        // de este modo agregamos un elemento al documento
                        Paragraph titulo = new Paragraph($" REPORTE POR ASISTENCIAS \n Reporte Generado el {DateTime.Now:dd/MM/yyyy}").SetBold().SetFontColor(new iText.Kernel.Colors.DeviceRgb(250, 98, 107));

                        // esto es solo para generar espacios en blanco
                        var white = new Paragraph(" ");


                        titulo.SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER);
                        doc.Add(titulo);
                        doc.Add(white);


                        //Asi se genera un recorrido

                        var tableReportes = new Table(4, true);
                        tableReportes.AddHeaderCell(new Cell().Add(new Paragraph("id").SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).SetBold()));
                        tableReportes.AddHeaderCell(new Cell().Add(new Paragraph("UserId").SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).SetBold()));
                        tableReportes.AddHeaderCell(new Cell().Add(new Paragraph("title").SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).SetBold()));
                        tableReportes.AddHeaderCell(new Cell().Add(new Paragraph("Completed").SetTextAlignment(iText.Layout.Properties.TextAlignment.LEFT).SetBold()));
                        tableReportes.GetHeader().SetBackgroundColor(new iText.Kernel.Colors.DeviceRgb(117, 113, 113)).SetFontColor(iText.Kernel.Colors.ColorConstants.WHITE);

                        if (datos != null && datos.Count > 0)
                            foreach (var item in datos)
                            {
                                tableReportes.AddCell(new Cell().Add(new Paragraph(item.id.ToString()).SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)));
                                tableReportes.AddCell(new Cell().Add(new Paragraph(item.userId.ToString()).SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)));
                                tableReportes.AddCell(new Cell().Add(new Paragraph(item.title).SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)));
                                tableReportes.AddCell(new Cell().Add(new Paragraph(item.complete?"Si":"No").SetTextAlignment(iText.Layout.Properties.TextAlignment.LEFT)));
                            }

                        doc.Add(tableReportes);
                        tableReportes.Flush();
                        tableReportes.Complete();
                        doc.Add(white);
                        doc.Add(white);
                        doc.Add(white);


                        doc.Close();
                        writer.Close();
                    }

                    this.doc = ms.ToArray();
                }
            }
            catch(Exception ee)
            {
                Console.WriteLine(ee.Message);
                this.doc=new byte[0];
            }
        }

    }
}
