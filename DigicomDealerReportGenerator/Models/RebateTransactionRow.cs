using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DigicomDealerReportGenerator.Models
{
    public class RebateTransactionRow : IRebateRow
    {
        public string DoorCode { get; set; }

        public string DoorName { get; set; }

        public string Address { get; set; }

        public string Location { get; set; }

        public string AccountNo { get; set; }

        public string SubscriberId { get; set; }

        public string Mdn { get; set; }

        public string Esn { get; set; }

        public string Sim { get; set; }

        public string HandsetModel { get; set; }

        public DateTime TransactionDate { get; set; }

        public string ProgramName { get; set; }

        public string RebateType { get; set; }

        public string QualificationStatus { get; set; }

        public decimal RebateAmount { get; set; }

        public DateTime PostedDate { get; set; }
    }
}
