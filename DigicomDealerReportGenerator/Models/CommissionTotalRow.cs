using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DigicomDealerReportGenerator.Models
{
    public class CommissionTotalRow
    {
        public string Agent { get; set; }

        public double Total { get; set; }

        public double CompleteTotal { get; set; }

        public bool IsTerminated { get; set; }

        public bool IsCommission { get; set; }
    }
}
