using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DigicomDealerReportGenerator.ViewModels;

using OfficeOpenXml;

namespace DigicomDealerReportGenerator.Models
{
    public class CommissionReportGeneratorModel
    {
        public CommissionReportGeneratorModel(CommissionReportGeneratorViewModel commissionReportGeneratorViewModel)
        {
            this.CommissionReportGeneratorViewModel = commissionReportGeneratorViewModel;
        }

        public CommissionReportGeneratorViewModel CommissionReportGeneratorViewModel { get; set; }

        public void GenerateSingleReport(string dealerCode, ExcelPackage package)
        {
            
        }
    }
}
