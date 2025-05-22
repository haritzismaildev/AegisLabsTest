using Microsoft.AspNetCore.Mvc;
using System.IO;
using ClosedXML.Excel;
using Rotativa.AspNetCore;

namespace SourceTest1.Controllers
{
    public class ExportController : Controller
    {
        public IActionResult ExportExcel()
        {
            //return View();

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("DataSheet");

                // Tambahkan header
                worksheet.Cell("A1").Value = "ID";
                worksheet.Cell("B1").Value = "Nama";
                worksheet.Cell("C1").Value = "Tanggal";

                // Contoh data: isi baris kedua dan seterusnya
                worksheet.Cell("A2").Value = 1;
                worksheet.Cell("B2").Value = "Haritz";
                worksheet.Cell("C2").Value = System.DateTime.Now.ToString("dd-MM-yyyy");

                // Simpan workbook ke MemoryStream
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    stream.Position = 0;

                    // Kembalikan stream sebagai file Excel dengan MIME type yang sesuai
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "data.xlsx");
                }
            }
        }
        public IActionResult ExportPdf()
        {
            // Jika data dinamis atau menggunakan model, siapkan model-nya di sini.
            // Contoh: var model = _dataService.GetData();

            // Kembalikan view sebagai PDF menggunakan Rotativa/eRotative Output
            return new ViewAsPdf("ExportPdf")
            {
                FileName = "data.pdf",
                PageOrientation = Rotativa.AspNetCore.Options.Orientation.Portrait,
                PageSize = Rotativa.AspNetCore.Options.Size.A4,
                CustomSwitches = "--footer-center \"Page: [page] of [toPage]\" --footer-font-size \"9\""
            };
        }
    }
}
