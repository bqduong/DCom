using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DigicomDealerReportGenerator.Models;

using LinqToExcel;

namespace DigicomDealerReportGenerator.MappingHelper
{
    public static class LinqToExcelMappingHelpers
    {
        public const string Disqualified = "disqualified";
        public const string Qualified = "qualified";

        public static void MapToLinq(ref ExcelQueryFactory excel, Func<string, string> getReportType, string filename)
        {
            LinqToExcelMappingHelpers.ModifyCommonTransactionRowMappings(ref excel);

            if (getReportType(filename) == Disqualified)
            {
                LinqToExcelMappingHelpers.ModifyDisqualilfiedTransactionRowMappings(ref excel);
            }
            else
            {
                LinqToExcelMappingHelpers.ModifyQualilfiedTransactionRowMappings(ref excel);
            }
        }

        public static void ModifyDisqualilfiedTransactionRowMappings(ref ExcelQueryFactory excel)
        {
            excel.AddMapping<DisqualifiedTransactionRow>(q => q.SubscriberStatus, "Subscriber Status");
            excel.AddMapping<DisqualifiedTransactionRow>(q => q.AccountBalance, "Account Balance");
            excel.AddMapping<DisqualifiedTransactionRow>(q => q.BusinesRuleReasonCode, "Business Rule Reason Code");
        }

        public static void ModifyQualilfiedTransactionRowMappings(ref ExcelQueryFactory excel)
        {
            excel.AddMapping<QualifiedTransactionRow>(q => q.RatePlanAmount, "Rate Plan Amount");
            excel.AddMapping<QualifiedTransactionRow>(q => q.BoltOnAmount, "Bolt On Amount");
            excel.AddMapping<QualifiedTransactionRow>(q => q.TransactionAmount, "Transaction Amount");
            excel.AddMapping<QualifiedTransactionRow>(q => q.PostedDate, "Posted Date");
        }

        public static void ModifyCommonTransactionRowMappings(ref ExcelQueryFactory excel)
        {
            excel.AddMapping<ITransactionRow>(q => q.DoorCode, "Door Code");
            excel.AddMapping<ITransactionRow>(q => q.DoorName, "Door Name");
            excel.AddMapping<ITransactionRow>(q => q.AccountNo, "Account Number");
            excel.AddMapping<ITransactionRow>(q => q.SubscriberId, "Subscriber  ID");
            excel.AddMapping<ITransactionRow>(q => q.Mdn, "MDN");
            excel.AddMapping<ITransactionRow>(q => q.Esn, "ESN");
            excel.AddMapping<ITransactionRow>(q => q.Sim, "SIM");
            excel.AddMapping<ITransactionRow>(q => q.EsnHistory, "ESN History");
            excel.AddMapping<ITransactionRow>(q => q.SimHistory, "SIM History");
            excel.AddMapping<ITransactionRow>(q => q.HandsetModel, "Handset Model");
            excel.AddMapping<ITransactionRow>(q => q.TransactionDate, "Transaction Date");
            excel.AddMapping<ITransactionRow>(q => q.TransactionType, "Transaction Type");
            excel.AddMapping<ITransactionRow>(q => q.RatePlan, "Rate Plan");
            excel.AddMapping<ITransactionRow>(q => q.BoltOn, "Bolt On");
        }

    }
}
