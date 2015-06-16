using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Sage300IntegrationSample.Models
{
    public class InvoiceLineModel
    {
        public string Description { get; set; }
        public string AccountCode { get; set; }
        public double AmountEx { get; set; }
    }
}
