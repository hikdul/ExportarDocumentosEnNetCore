using ExportarDocumentos.Models;
using OfficeOpenXml;

namespace ExportarDocumentos.Helpers
{
    /// <summary>
    ///  Clase para generar archivos exportables de excel.
    ///  esta forma es para separar la responsabilidad a un solo elemento
    /// </summary>
    class Excel : IExp
    {
        /// <summary>
        /// este seria el elemento como tal para generar el excel
        /// </summary>
        public byte[] doc { get; }

        /// <summary>
        /// Empty
        /// </summary>
        public Excel()
        {
            this.doc=new byte[0];
        }
        /// <summary>
        /// To Generato this doc whit list to 'Datos' data
        /// </summary>
        /// <param name="list"></param>
        public Excel(List<Datos> list)
        {
            try
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    using (ExcelPackage ep = new ExcelPackage())
                    {
                        ep.Workbook.Worksheets.Add("Reporte Poduccion Entre Periodos");
                        ExcelWorksheet ew = ep.Workbook.Worksheets[0];

                        ew.Cells.Style.Font.Size = 10;
                        ew.Cells.Style.Font.Name = "Arial";

                        ew.Cells[1, 1].Value = "Documento Generado";
                        ew.Cells[1, 2].Value = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                        
                            ew.Cells[2,1].Value="id";
                            ew.Cells[2,2].Value="userId";
                            ew.Cells[2,3].Value="title";
                            ew.Cells[2,4].Value="Completed";
                        for (int i = 0; i < list.Count(); i++)
                        {
                            ew.Cells[i+4,1].Value=list[i].id;
                            ew.Cells[i+4,2].Value=list[i].userId;
                            ew.Cells[i+4,3].Value=list[i].title;
                            ew.Cells[i+4,4].Value=list[i].complete?"Si":"No";
                            
                        }

                        ep.SaveAs(ms);
                        this.doc = ms.ToArray();

                    }}
            }
            catch 
            {
                this.doc=new byte[0];
            }
        }
    }
}