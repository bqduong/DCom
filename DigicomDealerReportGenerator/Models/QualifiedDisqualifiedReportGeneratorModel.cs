using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using DigicomDealerReportGenerator.FormattingHelper;
using DigicomDealerReportGenerator.ViewModels;

using OfficeOpenXml;

namespace DigicomDealerReportGenerator.Models
{
    public class QualifiedDisqualifiedReportGeneratorModel
    {
        private QualifiedDisqualifiedReportGeneratorViewModel viewModel;

        public QualifiedDisqualifiedReportGeneratorModel(QualifiedDisqualifiedReportGeneratorViewModel viewModel)
        {
            this.viewModel = viewModel;
        }

        public void GenerateSingleReport(string doorCode, ExcelPackage package)
        {
            var reportDataRows = DataHelpers.CreateReportData(doorCode, 
                                                                this.viewModel.StartDate, 
                                                                this.viewModel.EndDate, 
                                                                this.viewModel.IsQualified, 
                                                                this.viewModel.MasterTransactionList);

            if (reportDataRows.Any())
            {
                var worksheet = this.AppendReportData(reportDataRows, package, this.viewModel.IsQualified, this.viewModel.StartDate);
                this.SaveReportFile(reportDataRows.FirstOrDefault(), worksheet, this.viewModel.IsQualified, this.viewModel.StartDate,
                                    this.viewModel.EndDate, this.viewModel.DestinationPath);
            }
        }

        protected ExcelWorksheet AppendReportData(IEnumerable<ITransactionRow> reportDataRows, ExcelPackage package, bool isQualified, DateTime startDate)
        {
            this.viewModel.TemplateWorksheet = package.Workbook.Worksheets[1];
            var worksheet = this.viewModel.TemplateWorksheet;
            worksheet.Cells.Style.Font.Size = 8;

            if (isQualified)
            {
                this.AppendQualifiedWorksheetData(ref worksheet, reportDataRows, startDate);
            }
            else
            {
                this.AppendDisqualifiedWorksheetData(ref worksheet, reportDataRows, startDate);
            }

            return worksheet;
        }

        protected void SaveReportFile(ITransactionRow reportDataRow, ExcelWorksheet worksheet, bool isQualified, DateTime startDate, DateTime endDate, string destinationPath)
        {
            var fileName = DataHelpers.CreateReportFileName(reportDataRow, isQualified, startDate, endDate);
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

        protected void AppendQualifiedWorksheetData(ref ExcelWorksheet worksheet, IEnumerable<ITransactionRow> reportDataRows, DateTime startDate)
        {
            var rows = reportDataRows.Select(transactionRow => transactionRow as QualifiedTransactionRow).ToList();
            var properties = new QualifiedTransactionRow().GetType().GetProperties().ToList();

            properties.RemoveAt(3);

            var startRow = worksheet.Dimension.End.Row + 1;
            for (int i = 0; i < rows.Count; i++)
            {
                for (int j = 1; j < properties.Count + 1; j++)
                {
                    var value = rows[i].GetType().GetProperty(properties[j - 1].Name).GetValue(rows[i], null);
                    value = DataHelpers.GetDateString(value);
                    value = DataHelpers.ReturnCurrencyString(value);
                    worksheet.SetValue(i + startRow, j, value);
                }
            }

            var sumTotal = rows.Select(r => r.TransactionAmount).Sum();
            FormatHelper.FormatQualifiedReport(ref worksheet, startDate, sumTotal, startRow, properties, rows, this.viewModel.IsSoCalReport);
        }

        protected void AppendDisqualifiedWorksheetData(ref ExcelWorksheet worksheet, IEnumerable<ITransactionRow> reportDataRows, DateTime startDate)
        {
            var rows = reportDataRows.Select(transactionRow => transactionRow as DisqualifiedTransactionRow).ToList();
            var properties = new DisqualifiedTransactionRow().GetType().GetProperties();

            FormatHelper.FormatDisqualifedReportLegend(ref worksheet, startDate, this.viewModel.IsSoCalReport);

            var startRow = worksheet.Dimension.End.Row + 1;
            for (int i = 0; i < rows.Count; i++)
            {
                for (int j = 1; j < properties.Length + 1; j++)
                {
                    var value = rows[i].GetType().GetProperty(properties[j - 1].Name).GetValue(rows[i], null);
                    value = DataHelpers.GetDateString(value);
                    value = DataHelpers.ReturnCurrencyString(value);
                    worksheet.SetValue(i + startRow, j, value);
                    worksheet.Cells[i + startRow, j].Style.Font.Size = 8;
                }
            }
        }
    }
}