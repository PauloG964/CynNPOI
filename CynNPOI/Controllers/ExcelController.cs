using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.IO;

namespace ExcelWebApp.Controllers
{
    public class ExcelController : Controller
    {
        private readonly ExcelHelper _excelHelper;

        public ExcelController()
        {
            _excelHelper = new ExcelHelper();
        }
        [HttpGet]
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult Upload(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                ViewBag.Error = "Please select a file.";
                return View("Index");
            }

            var extension = Path.GetExtension(file.FileName);
            using var stream = file.OpenReadStream();
            var data = _excelHelper.ReadExcel(stream, extension);
            ViewBag.Data = data;
            return View("Index");
        }

        [HttpGet]
        public IActionResult Download(string format)
        {
            if (string.IsNullOrWhiteSpace(format))
                format = ".xlsx";

            var fileBytes = _excelHelper.WriteExcel(format);
            string contentType = format == ".xls" ?
                "application/vnd.ms-excel" :
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

            string fileName = format == ".xls" ? "sample.xls" : "sample.xlsx";

            return File(fileBytes, contentType, fileName);
        }
    }
}
