using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;

namespace Voxel.POC.ExcelExpot
{
    class Program
    {
        private static List<Invoice> _invoices = new List<Invoice>();

        static void Main(string[] args)
        {
            _invoices.Add(new Invoice("0001", 1.5m, 2m, DateTime.Now));
            _invoices.Add(new Invoice("0234", 2113m, 2325m, DateTime.Now.AddHours(+6).AddDays(-1)));
            _invoices.Add(new Invoice("FAST", 100m, 110m, DateTime.Now.AddHours(-2).AddDays(-5)));
            _invoices.Add(new Invoice("F001A", 55m, 60m, DateTime.Now.AddHours(14).AddMonths(-1)));
            ClosedXmlExcelExtractor.CreateExcelUsingClosedXml(_invoices);
            CsvHelperXmlExtractor.CreateExcel(_invoices);
        }
    }
}
