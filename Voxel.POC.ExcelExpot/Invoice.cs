using System;

namespace Voxel.POC.ExcelExpot
{
    public class Invoice
    {
        public string Reference { get; set; }
        public decimal SubTotal { get; set; }
        public decimal Total { get; set; }
        public DateTime InvoiceDate { get; set; }

        public Invoice(string reference, decimal subTotal, decimal total, DateTime invoiceDate)
        {
            Reference = reference;
            SubTotal = subTotal;
            Total = total;
            InvoiceDate = invoiceDate;
        }
    }
}