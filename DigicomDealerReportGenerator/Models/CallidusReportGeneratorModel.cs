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

        public string RetailMasterFilePath { get; set; }

        public string RetailOnlineMasterFilePath { get; set; }

        public void ProcessQPayReports()
        {
            //overwrite or create new source data file (crystalReportViewer file)
            this.CreateAdjustedCrystalReportsFile(false);

            //write function that returns number of days in the specified month

            //filter by city per day

            var date = CallidusReportGeneratorViewModel.DateSelect;
        }

        protected void CreateAdjustedCrystalReportsFile(bool isSoCalReport)
        {
            if (isSoCalReport)
            {
                using (ExcelPackage package = new ExcelPackage(DataHelpers.GetTemplateFile(true, isSoCalReport, this.CallidusReportGeneratorViewModel.ExecutionPath)))
                {
                    var adjustedSoCalData = this.CreateAdjustedReportData(this.MasterSoCalTransactionList, this.DateSelect);

                    var testDate = new DateTime(2013, 6, 1);
                    var sum = adjustedSoCalData.Select(a => a as QualifiedTransactionRow)
                                     .ToList()
                                     .Where(q => q.Location == "Fruitvale" && q.TransactionDate.Equals(testDate))
                                     .Select(qa => qa.TransactionAmount)
                                     .Sum();



                    var worksheet = this.AppendReportData(adjustedSoCalData, package, this.DateSelect);
                    this.SaveExcelFile(worksheet, new FileInfo("C:\\test.xlsx"));
                }
            }
            else
            {
                using (ExcelPackage package = new ExcelPackage(DataHelpers.GetTemplateFile(true, isSoCalReport, this.CallidusReportGeneratorViewModel.ExecutionPath)))
                {
                    var adjustedBayAreaData = this.CreateAdjustedReportData(this.MasterBayAreaTransactionList, this.DateSelect);

                    var distinctLocations = this.GetDistinctLocations(adjustedBayAreaData);

                    var dealerListData = new List<dynamic>();
                    var dateList = this.GetAllDatesInMonth(this.DateSelect.Year, this.DateSelect.Month);
                    foreach (var distinctLocation in distinctLocations)
                    {
                        var dailySumList = new List<dynamic>();
                        dailySumList.Add(distinctLocation);
                        foreach (var dateTime in dateList)
                        {
                            dailySumList.Add(this.GetSumPerLocationPerDay(distinctLocation, dateTime, adjustedBayAreaData));
                        }

                        dealerListData.Add(dailySumList);
                    }

                    var dailySumList2 = new List<dynamic>();
                    dailySumList2.Add("Total NorCal");
                    foreach (var dateTime in dateList)
                    {
                        dailySumList2.Add(this.GetTotalSumLocationsPerDay(dateTime, adjustedBayAreaData));
                    }
                    dealerListData.Add(dailySumList2);

                    ExcelPackage master = new ExcelPackage(new FileInfo(this.RetailMasterFilePath));
                    var masterWorksheet = master.Workbook.Worksheets[this.DateSelect.Month];
                    this.SetAllDatesOnWorksheet(ref masterWorksheet, dateList);
                    this.SetAllNorCalReimbursementAmountsOnWorksheet(ref masterWorksheet, dateList, dealerListData);
                    master.Save();
                    master.Dispose();

                    var worksheet = this.AppendReportData(adjustedBayAreaData, package, this.DateSelect);
                    this.SaveExcelFile(worksheet, new FileInfo("C:\\test.xlsx"));
                }
            }
        }

        private void SetAllNorCalReimbursementAmountsOnWorksheet(ref ExcelWorksheet worksheet, List<DateTime> dateList, List<dynamic> sumData)
        {
            var locations = new List<string> { "Fruitvale", "San Jose", "Hayward", "Concord", "Salinas", "Total NorCal" };
            
            foreach (var location in locations)
            {
                var row = this.GetCorrespondingRowForLocation(location);
                var data = this.GetCorrespondingRowDataForLocation(location, sumData);
                foreach (var date in dateList)
                {
                    worksheet.SetValue(row, date.Day + 1, data[date.Day]);
                }
            }
        }

        private void SetAllSoCalReimbursementAmountsOnWorksheet(ref ExcelWorksheet worksheet, List<DateTime> dateList, List<dynamic> sumData)
        {
            var locations = new List<string> { "Rosecrans", "Anaheim", "Imperial" "Santa Maria", "Total SoCal" };

            foreach (var location in locations)
            {
                var row = this.GetCorrespondingRowForLocation(location);
                var data = this.GetCorrespondingRowDataForLocation(location, sumData);
                foreach (var date in dateList)
                {
                    worksheet.SetValue(row, date.Day + 1, data[date.Day]);
                }
            }
        }

        public dynamic GetCorrespondingRowDataForLocation(string location, List<dynamic> sumData)
        {
            var data = new List<dynamic>();
            foreach (var o in sumData)
            {
                if (o[0] == location)
                {
                    data = o;
                }
            }
            return data;
        }

        //horrible hack to get corresponding row - use enums when refactoring
        public int GetCorrespondingRowForLocation(string location)
        {
            switch (location)
            {
                case "Fruitvale":
                    return 5;
                case "San Jose":
                    return 9;
                case "Hayward":
                    return 13;
                case "Concord":
                    return 17;
                case "Salinas":
                    return 21;
                case "Total NorCal":
                    return 25;
                case "Rosecrans":
                    return 31;
                case "Anaheim":
                    return 35;
                case "Imperial":
                    return 39;
                case "Santa Maria":
                    return 41;
                case "Total SoCal":
                    return 47;
                default:
                    return 5;
            }
        }
        
        private void SetAllDatesOnWorksheet(ref ExcelWorksheet worksheet, List<DateTime> dateList)
        {
            var dateRows = new List<int> { 3, 7, 11, 15, 19, 23, 29, 33, 37, 41, 45, 50 };

            foreach (var dateRow in dateRows)
            {
                foreach (var date in dateList)
                {
                    worksheet.SetValue(dateRow, date.Day + 1, date.ToShortDateString());
                }
            }
        }
        
        public List<DateTime> GetAllDatesInMonth(int year, int month)
        {
            var dates = new List<DateTime>();
            int daysInMonth = DateTime.DaysInMonth(year, month);
            for (int i = 0; i < daysInMonth; i++)
            {
                dates.Add(new DateTime(year, month, i + 1));
            }
            return dates;
        }

        public decimal GetSumPerLocationPerDay(string location, DateTime date, IEnumerable<ITransactionRow> adjustedMasterList)
        {
            var matchingLocations = adjustedMasterList.Select(a => a as QualifiedTransactionRow)
                                     .ToList()
                                     .Where(q => q.Location == location && q.TransactionDate.Equals(date)).ToList();

            var sumPerLocationPerDay = matchingLocations.Select(qa => qa.TransactionAmount).Sum();

            return sumPerLocationPerDay;
        }

        public decimal GetTotalSumLocationsPerDay(DateTime date, IEnumerable<ITransactionRow> adjustedMasterList)
        {
            var validLocations = adjustedMasterList.Select(a => a as QualifiedTransactionRow)
                                     .ToList()
                                     .Where(q => q.Location != null && q.TransactionDate.Equals(date)).ToList();

            var totalSumPerLocationsPerDay = validLocations.Select(qa => qa.TransactionAmount).Sum();

            return totalSumPerLocationsPerDay;
        }

        public dynamic GetDistinctLocations(IEnumerable<ITransactionRow> adjustedMasterList)
        {
            var locations = adjustedMasterList.Select(a => a as QualifiedTransactionRow)
                                           .Select(l => l.Location)
                                           .Distinct()
                                           .ToList();
            locations.Remove(null);

            return locations;
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
