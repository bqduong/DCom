using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DigicomDealerReportGenerator.Models
{
    public class ResidualRow
    {
        public string Mrr { get; set; }

        public string AccountId { get; set; }

        public string ActivationDate { get; set; }

        public string CustomerId { get; set; }

        public string MarketId { get; set; }

        public string MarketName { get; set; }

        public string Technology { get; set; }

        public string DealerId { get; set; }

        public string DealerCode { get; set; }

        public string Mac { get; set; }

        public string Agent { get; set; }

        public double ResidualAmount { get; set; }

        public string RevenueClassName { get; set; }
    }
}
