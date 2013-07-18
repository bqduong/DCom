using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DigicomDealerReportGenerator.Models
{
    public class CommissionRow
    {
        public string MarketId { get; set; }

        public string MarketName { get; set; }

        public string LoginName { get; set; }

        public string DealerCode { get; set; }

        public string DealerLocation { get; set; }

        public string OicTransactionType { get; set; }

        public DateTime TransactionDate { get; set; }

        public string OfferId { get; set; }

        public string OfferName { get; set; }

        public string ContractType { get; set; }

        public string AccountId { get; set; }

        public string CustomerId { get; set; }

        public DateTime ActivationDate { get; set; }

        public string CustomerFirstName { get; set; }

        public string CustomerLastName { get; set; }

        public string AccountAge { get; set; }

        public string ServiceType { get; set; }

        public string BundleType { get; set; }

        public string EquipmentSerialNumber { get; set; }

        public string Agent { get; set; }

        public string PlanElement { get; set; }

        public string RecurringPrice { get; set; }

        public string SubscriberCount { get; set; }

        public double CommissionAmount { get; set; }
    }
}
