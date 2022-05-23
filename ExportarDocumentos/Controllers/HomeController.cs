using ExportarDocumentos.Helpers;
using ExportarDocumentos.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;

namespace ExportarDocumentos.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }
            /// <summary>
            /// Muestra la lista de elementos que se van a imprir
            /// </summary>
            /// <returns></returns>
        public async Task<IActionResult> Index()
        {

            try
            {
                var model = await Datos.GetList();
                return View(model);
            }
            catch (Exception ee)
            {
                Console.WriteLine(ee.Message);
            }

            return View();
        }
        #region  Excel
        /// <summary>
        ///  aqui obtenemos los datos y los exportamos en su propio elemenos
        /// </summary>
        /// <returns></returns>
        public async Task<FileResult> Excel()
        {
            try
            {
                // Obtenemos nuestros datos
                var datos = await Datos.GetList(); 
                // creamos el excel
                var buffer = new Excel(datos);
                // retornamos los datos en un archivo excel valiod
                return File(buffer.doc, "application/vnd.ms-excel", "ExcelPrueba.xlsx");

            }
            catch (Exception ee)
            {
                Console.Error.WriteLine(ee.Message);
                // en caso de error enviamos un archivo completamente vacio
                return File(new byte[0], "application/vnd.ms-excel", "Empty.xlsx");
            }
            
        }

        #endregion

        #region  Word

        public async Task<FileResult> Word(){

            try
            {
                // Obtenemos nuestros datos
                var datos = await Datos.GetList(); 
                // creamos el excel
                var buffer = new Word(datos);
                // retornamos los datos en un archivo excel valiod
                return File(buffer.doc, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "Ejemplo.Docx");
            }
            catch (Exception ee)
            {
                Console.Error.WriteLine(ee.Message);
                // en caso de error enviamos un archivo completamente vacio
                return File(new byte[0], "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "Empty.Docx");
            }
        }

        #endregion


        #region  pdf

        public async Task<FileResult> pdf(){

            try
            {
                // Obtenemos nuestros datos
                var datos = await Datos.GetList(); 
                // creamos el excel
                var buffer = new pdf(datos);
                // retornamos los datos en un archivo excel valiod
                return File(buffer.doc, "application/pdf", "Ejemplo.PDF");
            }
            catch (Exception ee)
            {
                Console.Error.WriteLine(ee.Message);
                // en caso de error enviamos un archivo completamente vacio
                return File(new byte[0], "application/pdf", "Error.PDF");
            }
        }

        #endregion

        #region  que vienen con el controlador base
        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
        #endregion
    }
}