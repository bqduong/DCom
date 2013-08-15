using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DigicomDealerReportGenerator.Models
{
    public class QualifiedTransactionRow : ITransactionRow
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

        public string EsnHistory { get; set; }

        public string SimHistory { get; set; }

        public string HandsetModel { get; set; }

        public DateTime TransactionDate { get; set; }

        public string TransactionType { get; set; }

        public string RatePlan { get; set; }

        public string BoltOn { get; set; }

        public string RatePlanAmount { get; set; }

        public string BoltOnAmount { get; set; }

        public decimal TransactionAmount { get; set; }

        public DateTime PostedDate { get; set; }
    }
}
