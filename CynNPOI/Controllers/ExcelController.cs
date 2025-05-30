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

            var extension = Path.GetExtension(file.FileName).ToLowerInvariant();

            if (extension != ".xls" && extension != ".xlsx" && extension != ".xlsm")
            {
                ViewBag.Error = "Unsupported file format. Only .xls, .xlsx, and .xlsm are supported.";
                return View("Index");
            }

            using var stream = file.OpenReadStream();
            var data = _excelHelper.ReadExcel(stream, extension);
            ViewBag.Data = data;
            return View("Index");
        }

        [HttpGet]
        public IActionResult Download(string format)
        {
            format = string.IsNullOrWhiteSpace(format) ? ".xlsx" : format.ToLowerInvariant();

            if (format != ".xls" && format != ".xlsx" && format != ".xlsm")
                return BadRequest("Invalid format. Supported formats are .xls, .xlsx, and .xlsm.");

            var fileBytes = _excelHelper.WriteExcel(format);

            string contentType = format switch
            {
                ".xls" => "application/vnd.ms-excel",
                ".xlsx" or ".xlsm" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                _ => "application/octet-stream"
            };

            string fileName = $"sample{format}";

            return File(fileBytes, contentType, fileName);
        }
    }
}
