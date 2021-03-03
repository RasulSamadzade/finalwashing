using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace WpfApp1
{
    class ExcelHandling
    {
        DatabaseHandling databaseHandling;

        public ExcelHandling()
        {
            databaseHandling = new DatabaseHandling();
        }

        public void generateExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage excel = new ExcelPackage())
            {
                excel.Workbook.Worksheets.Add("Worksheet1");
                excel.Workbook.Worksheets.Add("Worksheet2");
                excel.Workbook.Worksheets.Add("Worksheet3");
                var data = databaseHandling.generateListFromDatabase();
                var headerRow = new List<string[]>() { new string[] { "Type", "Name", "IdCode", "Layer", "TopBottom", "IdCode", "Defect", "Input1", "Input2", "Decision", "Date" } };
                string headerRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1";
                string borderRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + (data.Count + 1).ToString();
                var worksheet = excel.Workbook.Worksheets["Worksheet1"];
                setExcelStyle(worksheet, headerRange, borderRange);
                worksheet.Cells[headerRange].LoadFromArrays(headerRow);
                worksheet.Cells[2, 1].LoadFromArrays(data);
                var dir = Directory.GetDirectoryRoot(Directory.GetCurrentDirectory()) + "TechnicalProb";
                FileInfo excelFile = new FileInfo(dir + "\\FinalWashing.xlsx");
                excel.SaveAs(excelFile);
            }
        }

        public void setExcelStyle(ExcelWorksheet worksheet, String headerRange, String borderRange)
        {
            worksheet.Cells[headerRange].Style.Font.Bold = true;
            worksheet.Cells[headerRange].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[headerRange].Style.Fill.BackgroundColor.SetColor(0, 150, 150, 150);
            worksheet.Cells.AutoFitColumns();
            worksheet.Column(2).Width = 20;
            worksheet.Column(3).Width = 20;
            worksheet.Column(4).Width = 20;
            worksheet.Column(5).Width = 20;
            worksheet.Column(6).Width = 20;
            worksheet.Column(7).Width = 20;
            worksheet.Column(8).Width = 20;
            worksheet.Column(9).Width = 20;
            worksheet.Column(10).Width = 20;
            worksheet.Cells[borderRange].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells[borderRange].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells[borderRange].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells[borderRange].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
        }
    }
}
