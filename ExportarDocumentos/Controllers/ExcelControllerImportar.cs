
using ExportarDocumentos.Models;
using Microsoft.AspNetCore.Mvc;

namespace ExportarDocumentos.Controllers
{
    /// <summary>
    /// Para generar los ejemplos de importar datos desde un excel
    /// </summary>
    public partial class ExcelController : Controller
    {

        /// <summary>
        /// Solo es una vista vacia... que permite tarto cargar el archivo como obtener uno de ejemplo
        /// </summary>
        /// <returns></returns>
        public ActionResult Importar()
        {
            return View();
        }


        [HttpPost]
        public ActionResult Importar(ImportarDatos ins)
        {
            try
            {
                var ext = Path.GetExtension(ins.doc.FileName);

                if (ext != ".xlsx")
                {
                    ViewBag.Err = "Extencion No Valida";
                    return View();
                }
                //var array = await ins.ConvertToArrayBytes();

                ViewBag.datos = ins.GetListFromExcel();
                
            }
            catch (Exception ee)
            {
                Console.WriteLine(ee.Message);
                
            }
            return View(); 
        }
            

   
      
    }
}