using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OfficeOpenXml;

namespace DigicomDealerReportGenerator.ViewModels
{
    public class CallidusReportGenerator
    {
        private DigicomDealerReportGeneratorViewModel viewModel;

        public CallidusReportGenerator(DigicomDealerReportGeneratorViewModel viewModel)
        {
            this.viewModel = viewModel;
        }

        public void GenerateCallidusReport(ExcelPackage package)
        {
            //var reportDataRows = DataHelpers.CreateReportData(doorCode,
            //                                                    this.viewModel.StartDate,
            //                                                    this.viewModel.EndDate,
            //                                                    this.viewModel.IsQualified,
            //                                                    this.viewModel.MasterTransactionList);

            //if (reportDataRows.Any())
            //{
            //    var worksheet = this.AppendReportData(reportDataRows, package, this.viewModel.IsQualified, this.viewModel.StartDate);
            //    this.SaveReportFile(reportDataRows.FirstOrDefault(), worksheet, this.viewModel.IsQualified, this.viewModel.StartDate,
            //                        this.viewModel.EndDate, this.viewModel.DestinationPath);
            //}



            //overwrite or create new source data file (crystalReportViewer file)
        }
    }
}
