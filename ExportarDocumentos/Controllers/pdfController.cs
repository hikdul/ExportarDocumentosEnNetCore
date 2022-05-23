
using ExportarDocumentos.Models;
using Microsoft.AspNetCore.Mvc;

namespace ExportarDocumentos.Controllers
{
    public class pdfController : Controller
    {


        public pdfController()
        {
        }


        public ActionResult Index()
        {
            return View();
        }

        
    }
}