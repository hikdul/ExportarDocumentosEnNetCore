
using ExportarDocumentos.Models;
using Microsoft.AspNetCore.Mvc;

namespace ExportarDocumentos.Controllers
{
    public partial class ExcelController : Controller
    {


        public ExcelController()
        {
        }


        public ActionResult Index()
        {
            return View();
        }


        
    }
}