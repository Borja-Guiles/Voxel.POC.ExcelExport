using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using CsvHelper;
using CsvHelper.Excel;

namespace Voxel.POC.ExcelExpot
{
    public static class CsvHelperXmlExtractor
    {
        public static void CreateExcel(List<Invoice> invoices)
        {
            using (var writer = new ExcelWriter("V:\\Git\\Voxel.POC.ExcelExport\\Voxel.POC.ExcelExpot\\CsvHelper.Excel.xlsx"))
            {
                writer.WriteRecords(invoices);
            }
        }
    }
}