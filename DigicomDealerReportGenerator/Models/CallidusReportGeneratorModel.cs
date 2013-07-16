using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using DigicomDealerReportGenerator.FormattingHelper;
using DigicomDealerReportGenerator.ViewModels;

using OfficeOpenXml;

namespace DigicomDealerReportGenerator.Models
{
    public class CallidusReportGeneratorModel
    {
        public CallidusReportGeneratorModel(CallidusReportGeneratorViewModel callidusReportGeneratorViewModel)
        {
            this.CallidusReportGeneratorViewModel = callidusReportGeneratorViewModel;
        }

        public CallidusReportGeneratorViewModel CallidusReportGeneratorViewModel { get; set; }

        public IEnumerable<ITransactionRow> MasterBayAreaTransactionList { get; set; }

        public IEnumerable<ITransactionRow> MasterSoCalTransactionList { get; set; }

        public DateTime DateSelect { get; set; }

        public void ProcessQPayReports()
        {
            //overwrite or create new source data file (crystalReportViewer file)
            this.CreateAdjustedCrystalReportsFile(false);

            var date = CallidusReportGeneratorViewModel.DateSelect;
        }

        protected void CreateAdjustedCrystalReportsFile(bool isSoCalReport)
        {
            if (isSoCalReport)
            {
                using (ExcelPackage package = new ExcelPackage(DataHelpers.GetTemplateFile(true, isSoCalReport, this.CallidusReportGeneratorViewModel.ExecutionPath)))
                {
                    var adjustedSoCalData = this.CreateAdjustedReportData(this.MasterSoCalTransactionList, this.DateSelect);
                    var worksheet = this.AppendReportData(adjustedSoCalData, package, this.DateSelect);
                    this.SaveExcelFile(worksheet, new FileInfo("C:\\test.xlsx"));
                }
            }
            else
            {
                using (ExcelPackage package = new ExcelPackage(DataHelpers.GetTemplateFile(true, isSoCalReport, this.CallidusReportGeneratorViewModel.ExecutionPath)))
                {
                    var adjustedBayAreaData = this.CreateAdjustedReportData(this.MasterBayAreaTransactionList, this.DateSelect);
                    var worksheet = this.AppendReportData(adjustedBayAreaData, package, this.DateSelect);
                    this.SaveExcelFile(worksheet, new FileInfo("C:\\test.xlsx"));
                }
            }
        }
        
        public IEnumerable<ITransactionRow> CreateAdjustedReportData(IEnumerable<ITransactionRow> masterList, DateTime dateSelect)
        {
            var adjustedData = DataHelpers.AdjustTransactionDates(masterList, dateSelect);
            return adjustedData;
        }

        protected ExcelWorksheet AppendReportData(IEnumerable<ITransactionRow> reportDataRows, ExcelPackage package, DateTime dateSelect)
        {
            var worksheet = package.Workbook.Worksheets[1];
            worksheet.Cells.Style.Font.Size = 8;

            var rows = reportDataRows.Select(transactionRow => transactionRow as QualifiedTransactionRow).ToList();
            var properties = new QualifiedTransactionRow().GetType().GetProperties();

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

            return worksheet;
        }

        protected void SaveExcelFile(ExcelWorksheet worksheet, FileInfo filePath)
        {
            if (File.Exists(filePath.ToString()))
            {
                File.Delete(filePath.ToString());
            }
            ExcelPackage reportPackage = new ExcelPackage(filePath);
            reportPackage.Workbook.Worksheets.Add("Report", worksheet);
            reportPackage.Save();
            reportPackage.Dispose();
        }
    }
}
