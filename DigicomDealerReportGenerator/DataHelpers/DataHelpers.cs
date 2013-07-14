﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DigicomDealerReportGenerator.Models;

using LinqToExcel;

namespace DigicomDealerReportGenerator
{
    public static class DataHelpers
    {
        public const string Disqualified = "disqualified";

        public const string Qualified = "qualified";

        public static string GetReportType(string filename)
        {
            return filename.ToLower().Contains(Disqualified) ? Disqualified : Qualified;
        }


        public static bool IsQualified(string filename)
        {
            return !filename.ToLower().Contains(Disqualified);
        }


        public static DateTime GetLatestDate(IEnumerable<ITransactionRow> masterList, bool isQualified)
        {
            var endDate = new DateTime();
            if (isQualified)
            {
                endDate =
                    masterList.Select(transactionRow => transactionRow as QualifiedTransactionRow)
                              .ToList()
                              .Select(l => l.PostedDate)
                              .Where(l => l.Year != 0001)
                              .Max(l => l.Date);
            }
            else
            {
                endDate =
                    masterList.Select(transactionRow => transactionRow as DisqualifiedTransactionRow)
                              .ToList()
                              .Select(l => l.TransactionDate)
                              .Where(l => l.Year != 0001)
                              .Max(l => l.Date);
            }

            return endDate;
        }


        public static DateTime GetEarliestDate(IEnumerable<ITransactionRow> masterList, bool isQualified)
        {
            var startDate = new DateTime();
            if (isQualified)
            {
                startDate =
                    masterList.Select(transactionRow => transactionRow as QualifiedTransactionRow)
                              .ToList()
                              .Select(l => l.PostedDate)
                              .Where(l => l.Year != 0001)
                              .Min(l => l.Date);
            }
            else
            {
                startDate =
                    masterList.Select(transactionRow => transactionRow as DisqualifiedTransactionRow)
                              .ToList()
                              .Select(l => l.TransactionDate)
                              .Where(l => l.Year != 0001)
                              .Min(l => l.Date);
            }

            return startDate;
        }


        public static string GetStartingMonthAndYear(DateTime startDate)
        {
            var month = startDate.ToString("MMMM");
            var year = startDate.Year.ToString();

            return month + " " + year;
        }


        public static dynamic GetDateString(dynamic value)
        {
            if (value is DateTime)
            {
                string dateString = value.ToShortDateString();
                if (value.ToShortDateString().Contains("0001"))
                {
                    dateString = "";
                }
                return dateString;
            }
            else
            {
                return value;
            }
        }


        public static dynamic ReturnCurrencyString(dynamic value)
        {
            return value is double ? "$" + String.Format("{0:0.00}", value) : value;
        }


        public static string CreateReportFileName(ITransactionRow reportDataRow, bool isQualified, DateTime startDate, DateTime endDate)
        {
            var fileString = "";
            if (isQualified)
            {
                fileString = reportDataRow.DoorCode + " - " + reportDataRow.DoorName + " (Qualified - "
                             + startDate.ToString("MM/dd/yy") + " - " + endDate.ToString("MM/dd/yy") + ").xlsx";
            }
            else
            {
                fileString = reportDataRow.DoorCode + " - " + reportDataRow.DoorName + " (Disqualified - " + startDate.ToString("MM/dd/yy")
                   + " - " + endDate.ToString("MM/dd/yy") + ").xlsx";
            }

            return fileString.Replace("/", "-");
        }


        public static IEnumerable<ITransactionRow> CreateReportData(string doorCode, DateTime startDate, DateTime endDate, 
                                                                    bool isQualified, IEnumerable<ITransactionRow> masterTransactionList)
        {
            if (isQualified)
            {
                return
                    masterTransactionList.Select(transactionRow => transactionRow as QualifiedTransactionRow)
                        .Where(t => t.DoorCode == doorCode && t.TransactionDate >= startDate && t.TransactionDate <= endDate)
                        .OrderBy(t => t.TransactionDate)
                        .ToList();
            }
            return
                masterTransactionList.Select(transactionRow => transactionRow as DisqualifiedTransactionRow)
                    .Where(t => t.DoorCode == doorCode && t.TransactionDate >= startDate && t.TransactionDate <= endDate)
                    .OrderBy(t => t.TransactionDate)
                    .ToList();
        }

        //public static IEnumerable<QualifiedTransactionRow> AdjustTransactionDates(IEnumerable<ITransactionRow> masterTransactionList, DateTime startDate)
        //{
        //    foreach (var row in masterTransactionList)
        //    {
        //        var qRow = row as QualifiedTransactionRow;
        //        if (qRow.TransactionDate.Month != startDate.Month)
        //        {
        //            qRow.BoltOn = qRow.TransactionDate.ToShortTimeString();
        //            //qRow.TransactionDate = 
        //        }
        //    }

        //    return masterTransactionList as QualifiedTransactionRow;
        //}

        //public static DateTime GetMatchingTransactionDate(IEnumerable<ITransactionRow> masterTransactionList, DateTime targetPostedDate)
        //{
        //    return null;
        //}


        public static FileInfo GetTemplateFile(bool isQualified, bool isSoCalReport, string executionPath)
        {
            var templatePath = ""; 

            if (isSoCalReport)
            {
                templatePath = isQualified
                                   ? executionPath + "Digicom Templates\\SoCal Qualified Transactions Template.xlsx"
                                   : executionPath + "Digicom Templates\\SoCal Disqualified Transactions Template Final.xlsx";
            }
            else
            {
                templatePath = isQualified
                                   ? executionPath + "Digicom Templates\\Qualified Transactions Template.xlsx"
                                   : executionPath + "Digicom Templates\\Disqualified Transactions Template Final.xlsx";
            }

            return new FileInfo(templatePath);
        }


        public static List<IDealerIdentification> GenerateDoorNameListWithDoorCode(IEnumerable<ITransactionRow> masterList)
        {
            var distinctItems = masterList.GroupBy(i => new { i.DoorCode, i.DoorName }).Select(g => g.Key).ToList();
            var distinctDealers = new List<IDealerIdentification>();

            distinctDealers.Add(new DealerIdentification()
            {
                DoorCode = "All",
                DoorName = "All",
                FullDealerIdentification = "[All Dealers]"
            });

            distinctDealers.AddRange(distinctItems
                .Select(distinctItem => new DealerIdentification()
                {
                    DoorCode = distinctItem.DoorCode,
                    DoorName = distinctItem.DoorName,
                    FullDealerIdentification = distinctItem.DoorCode + " - " + distinctItem.DoorName
                })
                .Cast<IDealerIdentification>().ToList());

            return distinctDealers;
        }


        public static IEnumerable<ITransactionRow> GetMasterListOfTransactionRows(bool isQualified, ExcelQueryFactory excel)
        {
            if (isQualified)
            {
                var validRows = (from x in excel.Worksheet<QualifiedTransactionRow>() select x).ToList();
                validRows.RemoveAll(v => v.DoorCode == null);
                return validRows;
            }
            else
            {
                var validRows = (from x in excel.Worksheet<DisqualifiedTransactionRow>() select x).ToList();
                validRows.RemoveAll(v => v.DoorCode == null);
                return validRows;
            }
        }
    }
}
