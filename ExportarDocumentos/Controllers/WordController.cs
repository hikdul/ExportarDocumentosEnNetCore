
using ExportarDocumentos.Models;
using Microsoft.AspNetCore.Mvc;

namespace ExportarDocumentos.Controllers
{
    public class WordController : Controller
    {


        public WordController()
        {
        }


        public ActionResult Index()
        {
            return View();
        }

        
    }
}