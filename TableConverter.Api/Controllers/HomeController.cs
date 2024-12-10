using Microsoft.AspNetCore.Mvc;
using TableConverterService;

namespace TableConverter.Controllers
{
    public class HomeController : Controller
    {
        private readonly IPdfToExcelService _pdfToExcelService;

        public HomeController(IPdfToExcelService pdfToExcelService)
        {
            _pdfToExcelService = pdfToExcelService;
        }

        [HttpGet]
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> UploadFile(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                ModelState.AddModelError("", "Please select a valid file.");
                return View("Index");
            }

            var excelFilePath = await _pdfToExcelService.ConvertPdfToExcelAsync(file);

            if (string.IsNullOrEmpty(excelFilePath))
            {
                ModelState.AddModelError("", "Conversion failed. Please try again.");
                return View("Index");
            }

            var fileBytes = System.IO.File.ReadAllBytes(excelFilePath);
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "converted.xlsx");
        }
    }
}