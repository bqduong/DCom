﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
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

        public IEnumerable<RebateTransactionRow> MasterBayAreaRebateTransactionList { get; set; }

        public IEnumerable<RebateTransactionRow> MasterSoCalRebateTransactionList { get; set; }
        
        public DateTime DateSelect { get; set; }

        public string RetailMasterFilePath { get; set; }

        public string RetailOnlineMasterFilePath { get; set; }

        public void ProcessQPayReports()
        {
            this.CreateAdjustedCrystalReportsFile(false);
            //this.CreateAdjustedCrystalReportsFile(true);
            MessageBox.Show("Callidus complete.");
        }

        protected void CreateAdjustedCrystalReportsFile(bool isSoCalReport)
        {
            if (isSoCalReport)
            {
                using (ExcelPackage package = new ExcelPackage(DataHelpers.GetTemplateFile(true, isSoCalReport, true, this.CallidusReportGeneratorViewModel.ExecutionPath)))
                {
                    var adjustedSoCalData = this.CreateAdjustedReportData(this.MasterSoCalTransactionList, this.DateSelect);

                    var distinctLocations = this.GetDistinctLocations(adjustedSoCalData);

                    var dealerListData = new List<dynamic>();
                    var dateList = this.GetAllDatesInMonth(this.DateSelect.Year, this.DateSelect.Month);
                    foreach (var distinctLocation in distinctLocations)
                    {
                        var dailySumList = new List<dynamic>();
                        dailySumList.Add(distinctLocation);
                        foreach (var dateTime in dateList)
                        {
                            dailySumList.Add(this.GetSumPerLocationPerDay(distinctLocation, dateTime, adjustedSoCalData));
                        }

                        dealerListData.Add(dailySumList);
                    }

                    ExcelPackage master = new ExcelPackage(new FileInfo(this.RetailMasterFilePath));
                    var masterWorksheet = master.Workbook.Worksheets[this.DateSelect.Month];
                    this.SetAllDatesOnWorksheet(ref masterWorksheet, dateList);
                    this.SetAllSoCalReimbursementAmountsOnWorksheet(ref masterWorksheet, dateList, dealerListData);
                    this.SetAllSoCalRebateAmountsOnWorksheet(ref masterWorksheet, dateList, dealerListData);
                    master.Save();
                    master.Dispose();

                    var worksheet = this.AppendReportData(adjustedSoCalData, package, this.DateSelect);
                    var fileName = "CrystalReportViewer - LA Retail " +
                                   CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(this.DateSelect.Month) + " " +
                                   this.DateSelect.Year + ".xlsx";
                    this.SaveExcelFile(worksheet, new FileInfo(CallidusReportGeneratorViewModel.DestinationPath + "\\" + fileName));
                }
            }
            else
            {
                using (ExcelPackage package = new ExcelPackage(DataHelpers.GetTemplateFile(true, isSoCalReport, true, this.CallidusReportGeneratorViewModel.ExecutionPath)))
                {
                    //for Reimbursement
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


                    //for Rebate
                    //var adjustedRebateBayAreaData = this.CreateAdjustedRebateReportData(this.MasterBayAreaRebateTransactionList, this.DateSelect);

                    var distinctRebateLocations = this.GetDistinctLocations(MasterBayAreaRebateTransactionList);

                    var dealerRebateListData = new List<dynamic>();
                    var dateRebateList = this.GetAllDatesInMonth(this.DateSelect.Year, this.DateSelect.Month);
                    foreach (var distinctLocation in distinctRebateLocations)
                    {
                        var dailySumList = new List<dynamic>();
                        dailySumList.Add(distinctLocation);
                        foreach (var dateTime in dateRebateList)
                        {
                            dailySumList.Add(this.GetSumPerLocationPerDay(distinctLocation, dateTime, MasterBayAreaRebateTransactionList));
                        }

                        dealerRebateListData.Add(dailySumList);
                    }
                    
                    ExcelPackage master = new ExcelPackage(new FileInfo(this.RetailMasterFilePath));
                    var masterWorksheet = master.Workbook.Worksheets[this.DateSelect.Month];
                    this.SetAllDatesOnWorksheet(ref masterWorksheet, dateList);
                    this.SetAllNorCalReimbursementAmountsOnWorksheet(ref masterWorksheet, dateList, dealerListData);
                    this.SetAllNorCalRebateAmountsOnWorksheet(ref masterWorksheet, dateRebateList, dealerRebateListData);
                    master.Save();
                    master.Dispose();

                    var worksheet = this.AppendReportData(adjustedBayAreaData, package, this.DateSelect);
                    var fileName = "CrystalReportViewer - Bay Retail " +
                                   CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(this.DateSelect.Month) + " " +
                                   this.DateSelect.Year + ".xlsx";
                    this.SaveExcelFile(worksheet, new FileInfo(CallidusReportGeneratorViewModel.DestinationPath + "\\" + fileName));
                }
            }
        }

        private void SetAllNorCalReimbursementAmountsOnWorksheet(ref ExcelWorksheet worksheet, List<DateTime> dateList, List<dynamic> sumData)
        {
            var locations = new List<string> { "Fruitvale", "San Jose", "Hayward", "Concord" };
            
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
            var locations = new List<string> { "Rosecrans", "Anaheim", "Imperial" };

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

        private void SetAllNorCalRebateAmountsOnWorksheet(ref ExcelWorksheet worksheet, List<DateTime> dateList, List<dynamic> sumData)
        {
            var locations = new List<string> { "Fruitvale", "San Jose", "Hayward", "Concord" };

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

        private void SetAllSoCalRebateAmountsOnWorksheet(ref ExcelWorksheet worksheet, List<DateTime> dateList, List<dynamic> sumData)
        {
            var locations = new List<string> { "Rosecrans", "Anaheim", "Imperial" };

            foreach (var location in locations)
            {
                var row = this.GetCorrespondingRebateRowForLocation(location);
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
                //2014 Mod
                case "Fruitvale":
                    return 4;
                case "San Jose":
                    return 9;
                case "Hayward":
                    return 14;
                case "Concord":
                    return 19;
                case "Rosecrans":
                    return 31;
                case "Anaheim":
                    return 36;
                case "Imperial":
                    return 41;
                default:
                    return 4;
            }
        }

        public int GetCorrespondingRebateRowForLocation(string location)
        {
            switch (location)
            {
                //2014 Mod
                case "Fruitvale":
                    return 5;
                case "San Jose":
                    return 10;
                case "Hayward":
                    return 15;
                case "Concord":
                    return 20;
                case "Rosecrans":
                    return 32;
                case "Anaheim":
                    return 37;
                case "Imperial":
                    return 42;
                default:
                    return 4;
            }
        }
        
        private void SetAllDatesOnWorksheet(ref ExcelWorksheet worksheet, List<DateTime> dateList)
        {
            //2014
            var dateRows = new List<int> { 2, 7, 12, 17, 22, 29, 34, 39, 45, 51 };
            
            //2013
            //var dateRows = new List<int> { 3, 7, 11, 15, 19, 23, 29, 33, 37, 41, 45, 50 };

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

        public decimal GetSumPerLocationPerDay(string location, DateTime date, IEnumerable<RebateTransactionRow> adjustedMasterList)
        {
            var matchingLocations = adjustedMasterList
                                     .ToList()
                                     .Where(q => q.Location == location && q.TransactionDate.Equals(date)).ToList();

            var sumPerLocationPerDay = matchingLocations.Select(qa => qa.RebateAmount).Sum();

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

        public dynamic GetDistinctLocations(IEnumerable<RebateTransactionRow> adjustedMasterList)
        {
            var locations = adjustedMasterList
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

        //public IEnumerable<ITransactionRow> CreateAdjustedReportData(IEnumerable<RebateTransactionRow> masterList, DateTime dateSelect)
        //{
        //    var adjustedData = DataHelpers.AdjustTransactionDates(masterList, dateSelect);
        //    return adjustedData;
        //}

        protected ExcelWorksheet AppendReportData(IEnumerable<ITransactionRow> reportDataRows, ExcelPackage package, DateTime dateSelect)
        {
            var worksheet = package.Workbook.Worksheets[1];
            worksheet.Cells.Style.Font.Size = 8;

            var rows = reportDataRows.Select(transactionRow => transactionRow as QualifiedTransactionRow).ToList();
            var properties = new QualifiedTransactionRow().GetType().GetProperties().ToList();

            properties.RemoveAt(11);

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
            worksheet.SetValue(4, 17, DataHelpers.GetStartingMonthAndYear(dateSelect));
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
