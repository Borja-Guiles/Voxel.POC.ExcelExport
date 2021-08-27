using System.Collections.Generic;
using ClosedXML.Excel;

namespace Voxel.POC.ExcelExpot
{
    class ClosedXmlExcelExtractor
    {
        public static void CreateExcelUsingClosedXml(List<Invoice> invoices)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Invoices");
                var row = 1;
                worksheet.Column(1).DataType = XLDataType.Text;
                worksheet.Column(2).DataType = XLDataType.Number;
                worksheet.Column(3).DataType = XLDataType.Number;
                worksheet.Column(4).DataType = XLDataType.DateTime;
                worksheet.Row(1).DataType = XLDataType.Text;
                worksheet.Cell(row, 1).SetValue("Reference");
                worksheet.Cell(row, 2).SetValue("SubTotal");
                worksheet.Cell(row, 3).SetValue("Total");
                worksheet.Cell(row, 4).SetValue("InvoiceDate");
                row++;
                foreach (var invoice in invoices)
                {
                    worksheet.Cell(row, 1).SetValue(invoice.Reference);
                    worksheet.Cell(row, 2).SetValue(invoice.SubTotal);
                    worksheet.Cell(row, 3).SetValue(invoice.Total);
                    worksheet.Cell(row, 4).SetValue(invoice.InvoiceDate);
                    row++;
                }

                workbook.SaveAs("V:\\Git\\Voxel.POC.ExcelExport\\Voxel.POC.ExcelExpot\\ClosedXml.xlsx");

            }
        }
    }
}