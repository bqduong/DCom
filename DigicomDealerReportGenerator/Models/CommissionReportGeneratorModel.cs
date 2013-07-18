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
    public class CommissionReportGeneratorModel
    {
        public CommissionReportGeneratorModel(CommissionReportGeneratorViewModel commissionReportGeneratorViewModel)
        {
            this.CommissionReportGeneratorViewModel = commissionReportGeneratorViewModel;
        }

        public CommissionReportGeneratorViewModel CommissionReportGeneratorViewModel { get; set; }

        public void GenerateSingleReport(string fullDealerId, ExcelPackage package)
        {

            var fullDealerSplit = fullDealerId.Split('-');
            var reportDataRows =
                this.CommissionReportGeneratorViewModel.MasterTransactionList.Where(m => m.DealerCode == fullDealerSplit[0].Trim() && m.Agent == fullDealerSplit[1].Trim())
                    .ToList();

            if (reportDataRows.Any())
            {
                var worksheet = this.AppendReportData(reportDataRows, package);
                this.SaveReportFile(reportDataRows.FirstOrDefault(), worksheet, this.CommissionReportGeneratorViewModel.DestinationPath);
            }
        }


        protected ExcelWorksheet AppendReportData(
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

            var sumTotal = rows.Select(r => r.CommissionAmount).Sum();
            worksheet.SetValue(rows.Count + startRow, properties.Count(), "$" + String.Format("{0:0.00}", sumTotal));

            return worksheet;
        }

        protected void SaveReportFile(
            CommissionRow reportDataRow,
            ExcelWorksheet worksheet,
            string destinationPath)
        {
            var fileName = reportDataRow.DealerCode + " - " + reportDataRow.Agent + " - Commission Report - Week " + this.CommissionReportGeneratorViewModel.WeekInput + ".xlsx";
            fileName = fileName.Replace("/", " ");

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
    }
}
