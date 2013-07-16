using DigicomDealerReportGenerator.ViewModels;

using OfficeOpenXml;

namespace DigicomDealerReportGenerator.Models
{
    public class CallidusReportGeneratorModel
    {
        private CallidusReportGeneratorViewModel viewModel;

        public CallidusReportGeneratorModel(CallidusReportGeneratorViewModel viewModel)
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
