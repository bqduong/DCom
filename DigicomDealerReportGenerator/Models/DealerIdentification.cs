using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DigicomDealerReportGenerator.Models
{
    class DealerIdentification : IDealerIdentification
    {
        public string DoorCode { get; set; }

        public string DoorName { get; set; }

        public string FullDealerIdentification { get; set; }
    }
}
