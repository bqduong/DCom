using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DigicomDealerReportGenerator.Models
{
    public interface IDealerIdentification
    {
        string DoorCode { get; set; }

        string DoorName { get; set; }

        string FullDealerIdentification { get; set; }
    }
}
