using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DigicomDealerReportGenerator.Models
{
    public interface IRebateRow
    {
        string DoorCode { get; set; }

        string DoorName { get; set; }

        string Address { get; set; }

        string AccountNo { get; set; }

        string SubscriberId { get; set; }

        string Mdn { get; set; }

        string Esn { get; set; }

        string Sim { get; set; }

        string HandsetModel { get; set; }

        DateTime TransactionDate { get; set; }

        string ProgramName { get; set; }

        string RebateType { get; set; }

        string QualificationStatus { get; set; }
        
        string RebateAmount { get; set; }

        DateTime PostedDate { get; set; }
    }
}
