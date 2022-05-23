using ExportarDocumentos.Models;
using Syncfusion.DocIO.DLS;

namespace ExportarDocumentos.Helpers
{

    public class Word : IExp
    {
        /// <summary>
        /// documento
        /// </summary>
        public byte[] doc { get; }
        /// <summary>
        /// ctor
        /// </summary>
        public Word()
        {
            this.doc=new byte[0];
        }
        /// <summary>
        ///  constructor para generar el documento
        /// </summary>
        /// <param name="list"></param>
        public Word(List<Datos> list)
        {
            try
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    //declaraciones necesarias para que el reporte termine ordenado
                    WordDocument document = new WordDocument();
                    //formatos del documento
                    var region = (short)Syncfusion.Office.LocaleIDs.es_CL;

                    var formatValue = new WCharacterFormat(document){FontName = "Arial",FontSize = 11f, LocaleIdASCII = region };
                    var formatHeader = new WCharacterFormat(document){FontName = "Arial",FontSize = 11f, LocaleIdASCII = region, Bold = true};
                    var formatHeaderTable = new WCharacterFormat(document) { FontName = "Arial", FontSize = 11f, LocaleIdASCII = region, Bold = true, TextColor = Syncfusion.Drawing.Color.White };

                    // se genera una seccion para ingresar valores al documento
                    WSection section = document.AddSection() as WSection;
                    section.PageSetup.Margins.All = 72;
                    section.PageSetup.PageSize = new Syncfusion.Drawing.SizeF(612, 792);
                    // luego se agrega los datos una a unos... anadiendolos antes a la seccion
                    var titulo = section.AddParagraph();
                    titulo.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                    IWTextRange textRange = titulo.AppendText($"Prueba exportar datos en un Word  \n Documento generado el {DateTime.Now:dd/MM/yyyy}") as WTextRange;
                    textRange.ApplyCharacterFormat(formatHeader);
                    textRange.CharacterFormat.TextColor = Syncfusion.Drawing.Color.FromArgb(1, 250, 98, 107);
                    section.AddParagraph().AppendBreak(BreakType.LineBreak);

                    // Recorrido dentro del documento
                    // se genera la tabla
                    var table = section.AddTable();
                    // si se genera un error o se empieza un recorrido vacio se va a error asi que es mejor validar todo y tomar precacausiones
                    if (list != null && list.Count > 0)
                    {
                        table = section.AddTable();
                        table.ResetCells(1 + list.Count, 4);

                        var cell = table[0, 0].AddParagraph();
                        cell.AppendText("id").ApplyCharacterFormat(formatHeaderTable);
                        cell.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                        table[0, 0].CellFormat.BackColor = Syncfusion.Drawing.Color.FromArgb(117, 113, 113);

                        cell = table[0, 1].AddParagraph();
                        cell.AppendText("userID").ApplyCharacterFormat(formatHeaderTable);
                        cell.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                        table[0, 1].CellFormat.BackColor = Syncfusion.Drawing.Color.FromArgb(117, 113, 113);

                        cell = table[0, 2].AddParagraph();
                        cell.AppendText("Title").ApplyCharacterFormat(formatHeaderTable);
                        cell.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                        table[0, 2].CellFormat.BackColor = Syncfusion.Drawing.Color.FromArgb(117, 113, 113);

                        cell = table[0, 3].AddParagraph();
                        cell.AppendText("Complete").ApplyCharacterFormat(formatHeaderTable);
                        cell.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Justify;
                        table[0, 3].CellFormat.BackColor = Syncfusion.Drawing.Color.FromArgb(117, 113, 113);

                        int i = 1;

                        foreach (var item in list)
                        {
                            cell = table[i, 0].AddParagraph();
                            cell.AppendText(item.id.ToString()).ApplyCharacterFormat(formatValue);
                            cell.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;

                            cell = table[i, 1].AddParagraph();
                            cell.AppendText(item.userId.ToString()).ApplyCharacterFormat(formatValue);
                            cell.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;

                            cell = table[i, 2].AddParagraph();
                            cell.AppendText(item.title.ToString()).ApplyCharacterFormat(formatValue);
                            cell.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;

                            cell = table[i, 3].AddParagraph();
                            cell.AppendText(item.complete?"SI":"NO").ApplyCharacterFormat(formatValue);
                            cell.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Justify;

                            i++;
                        }
                    }

                    // fin del reporte
                    document.Save(ms, Syncfusion.DocIO.FormatType.Docx);
                    this.doc= ms.ToArray();
                }
            }
            catch 
            {
                
                this.doc=new byte[0];
            }
        }
    }
}