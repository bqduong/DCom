using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DigicomDealerReportGenerator.ViewModels;

using OfficeOpenXml;

namespace DigicomDealerReportGenerator.Models
{
    public class CommissionResidualReportGeneratorModel
    {
        public CommissionResidualReportGeneratorModel(CommissionResidualReportGeneratorViewModel commissionReportGeneratorViewModel)
        {
            this.CommissionReportGeneratorViewModel = commissionReportGeneratorViewModel;
        }

        public CommissionResidualReportGeneratorViewModel CommissionReportGeneratorViewModel { get; set; }

        public void GenerateSingleCommissionReport(string fullDealerId, ExcelPackage package)
        {
            //var fullDealerSplit = fullDealerId.Split(new char[]{'-'}, 2);
            var reportDataRows =
                this.CommissionReportGeneratorViewModel.MasterCommissionTransactionList.Where(m => /*m.DealerCode.Trim() == fullDealerSplit[0].Trim() &&*/ m.Agent.Trim() == fullDealerId.Trim())
                    .ToList();

            if (reportDataRows.Any())
            {
                var worksheet = this.AppendCommissionReportData(reportDataRows, package);
                this.SaveReportFile(reportDataRows.FirstOrDefault(), worksheet, this.CommissionReportGeneratorViewModel.DestinationPath);
            }
        }


        public void GenerateSingleResidualReport(string fullDealerId, ExcelPackage package)
        {
            //var fullDealerSplit = fullDealerId.Split(new char[] { '-' }, 2);
            var reportDataRows =
                this.CommissionReportGeneratorViewModel.MasterResidualTransactionList.Where(m => /*m.DealerId.Trim() == fullDealerSplit[0].Trim() &&*/ m.Agent.Trim() == fullDealerId.Trim())
                    .ToList();

            if (reportDataRows.Any())
            {
                var worksheet = this.AppendResidualReportData(reportDataRows, package);
                this.SaveResidualReportFile(reportDataRows.FirstOrDefault(), worksheet, this.CommissionReportGeneratorViewModel.DestinationPath);
            }
        }

        protected ExcelWorksheet AppendCommissionReportData(
            IEnumerable<CommissionRow> reportDataRows, ExcelPackage package)
        {
            var worksheet = package.Workbook.Worksheets[1];
            worksheet.Cells.Style.Font.Size = 10;

            var rows = reportDataRows.Select(transactionRow => transactionRow as CommissionRow).ToList();
            var properties = new CommissionRow().GetType().GetProperties();

            var startRow = worksheet.Dimension.End.Row + 1;
            for (int i = 0; i < rows.Count; i++)
            {
                for (int j = 1; j < properties.Length + 1; j++)
                {
                    var value = rows[i].GetType().GetProperty(properties[j - 1].Name).GetValue(rows[i], null);
                    value = DataHelpers.GetDateString(value);
                    value = DataHelpers.ReturnCurrencyString(value);
                    worksheet.SetValue(i + startRow, j, value);
                }
            }

            var sumTotal = rows.Select(r => Math.Round(r.CommissionAmount, 2)).Sum();
            worksheet.SetValue(rows.Count + startRow, properties.Count(), "$" + String.Format("{0:0.00}", sumTotal));

            return worksheet;
        }

        protected ExcelWorksheet AppendResidualReportData(IEnumerable<ResidualRow> reportDataRows, ExcelPackage package)
        {
            var worksheet = package.Workbook.Worksheets[1];
            worksheet.Cells.Style.Font.Size = 10;

            var rows = reportDataRows.Select(transactionRow => transactionRow as ResidualRow).ToList();
            var properties = new ResidualRow().GetType().GetProperties();

            var startRow = worksheet.Dimension.End.Row + 1;
            for (int i = 0; i < rows.Count; i++)
            {
                for (int j = 1; j < properties.Length + 1; j++)
                {
                    var value = rows[i].GetType().GetProperty(properties[j - 1].Name).GetValue(rows[i], null);
                    value = DataHelpers.GetDateString(value);
                    value = DataHelpers.ReturnCurrencyString(value);
                    worksheet.SetValue(i + startRow, j, value);
                }
            }

            var sumTotal = rows.Select(r =>  Math.Round(r.ResidualAmount, 2)).Sum();
            worksheet.SetValue(rows.Count + startRow, properties.Count() - 1, "$" + String.Format("{0:0.00}", sumTotal));

            return worksheet;
        }

        protected void SaveResidualReportFile(
            ResidualRow reportDataRow,
            ExcelWorksheet worksheet,
            string destinationPath)
        {
            var fileName = /*reportDataRow.DealerCode + " - " + */reportDataRow.Agent + " - Residual Report - Week " + this.CommissionReportGeneratorViewModel.WeekInput + ".xlsx";
            fileName = fileName.Replace("/", " ");
            fileName = fileName.Replace(":", " ");

            var filePath = new FileInfo(destinationPath + "\\" + fileName);

            if (File.Exists(destinationPath + "\\" + fileName))
            {
                File.Delete(destinationPath + "\\" + fileName);
            }

            ExcelPackage reportPackage = new ExcelPackage(filePath);
            reportPackage.Workbook.Worksheets.Add("Report", worksheet);
            reportPackage.Save();
            reportPackage.Dispose();
        }

        protected void SaveReportFile(
            CommissionRow reportDataRow,
            ExcelWorksheet worksheet,
            string destinationPath)
        {
            var fileName = /*reportDataRow.DealerCode + " - " + */reportDataRow.Agent + " - Commission Report - Week " + this.CommissionReportGeneratorViewModel.WeekInput + ".xlsx";
            fileName = fileName.Replace("/", " ");
            fileName = fileName.Replace(":", " ");

            var filePath = new FileInfo(destinationPath + "\\" + fileName);

            if (File.Exists(destinationPath + "\\" + fileName))
            {
                File.Delete(destinationPath + "\\" + fileName);
            }

            ExcelPackage reportPackage = new ExcelPackage(filePath);
            reportPackage.Workbook.Worksheets.Add("Report", worksheet);
            reportPackage.Save();
            reportPackage.Dispose();
        }


        protected IEnumerable<CommissionTotalRow> GenerateCommissionTotalRows(IEnumerable<CommissionRow> masterCommissionRows, IEnumerable<ResidualRow> masterResidualRows)
        {
            var agents = masterCommissionRows.GroupBy(m => new { m.Agent }).Select(g => g.Key).ToList();


            var commissionTotalRows = (from agent in agents let sum = masterCommissionRows
                                            .Where(m => m.Agent == agent.Agent)
                                            .Select(c => c.CommissionAmount)
                                            .Sum() select new CommissionTotalRow()
                                                              {
                                                                  Agent = agent.ToString(),
                                                                  CompleteTotal = Math.Round(sum, 2),
                                                                  Total = Math.Round(sum, 2), 
                                                                  IsCommission = true, 
                                                                  IsTerminated = agent.ToString().ToLower().Contains("terminated")
                                                              }).OrderBy(a => a.Agent).ToList();

            agents = masterResidualRows.GroupBy(m => new { m.Agent }).Select(g => g.Key).ToList();

            var groupedResidualRows = (from agent in agents
                                       let sum = masterResidualRows
                      .Where(m => m.Agent == agent.Agent)
                      .Select(c => c.ResidualAmount)
                      .Sum()
                                       select new CommissionTotalRow()
                                       {
                                           Agent = agent.ToString(),
                                           CompleteTotal = Math.Round(sum, 2),
                                           Total = Math.Round(sum, 2),
                                           IsCommission = false,
                                           IsTerminated = agent.ToString().ToLower().Contains("terminated")
                                       }).OrderBy(a => a.Agent).ToList();

            commissionTotalRows.AddRange(groupedResidualRows);
            return commissionTotalRows.OrderBy(c => c.Agent);
        }
    }
}
