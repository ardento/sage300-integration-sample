using System;
using System.Collections.Generic;

namespace Sage300IntegrationSample.Models
{
    public class InvoiceModel
    {
        public string VendorID { get; set; }
        public string Description { get; set; }
        public string InvoiceNumber { get; set; }
        public DateTime InvoiceDate { get; set; }
        public double DocumentTotal { get; set; }
        
        public List<InvoiceLineModel> Lines { get; set; }
    }
}