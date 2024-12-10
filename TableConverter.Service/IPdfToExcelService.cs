using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using Microsoft.AspNetCore.Http;
using OfficeOpenXml;

namespace TableConverterService;

public interface IPdfToExcelService
{
    Task<string> ConvertPdfToExcelAsync(IFormFile file);
}

public class PdfToExcelService : IPdfToExcelService
{
    public async Task<string> ConvertPdfToExcelAsync(IFormFile file)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var tempPdfPath = Path.GetTempFileName();
        var tempExcelPath = Path.ChangeExtension(tempPdfPath, ".xlsx");

        // Сохраняем PDF во временный файл
        using (var stream = new FileStream(tempPdfPath, FileMode.Create))
        {
            await file.CopyToAsync(stream);
        }

        // Читаем содержимое PDF
        var pdfReader = new PdfReader(tempPdfPath);
        var pdfDocument = new PdfDocument(pdfReader);

        using (var package = new ExcelPackage())
        {
            var worksheet = package.Workbook.Worksheets.Add("ExtractedData");
            int row = 1;

            for (int i = 1; i <= pdfDocument.GetNumberOfPages(); i++)
            {
                var page = pdfDocument.GetPage(i);
                var text = PdfTextExtractor.GetTextFromPage(page);
                worksheet.Cells[row++, 1].Value = text; // Вставляем текст постранично
            }

            package.SaveAs(new FileInfo(tempExcelPath));
        }

        pdfReader.Close();
        return tempExcelPath;
    }
}