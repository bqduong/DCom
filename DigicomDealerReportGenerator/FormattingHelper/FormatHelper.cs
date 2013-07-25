using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

using DigicomDealerReportGenerator.Models;

using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace DigicomDealerReportGenerator.FormattingHelper
{
    public static class FormatHelper
    {
        public static void FormatDisqualifedReportLegend(ref ExcelWorksheet worksheet, DateTime startDate, bool isSoCalReport)
        {
            int soCalOffset = 0;

            worksheet.InsertRow(1, 16);
            worksheet.Cells[1, 4].Style.Font.Size = 14;
            if (isSoCalReport)
            {
                worksheet.SetValue(1, 1, "MetroPCS Wireless, Inc. Sales Incentive and Compensation Division");
                worksheet.SetValue(2, 1, "Disqualified Transaction Report");
                soCalOffset = 2;
            }
            else
            {
                worksheet.SetValue(1, 4, "Disqualified Transaction Report");
            }

            worksheet.Cells[2 + soCalOffset, 2].Style.Font.Size = 9;
            worksheet.Cells[3 + soCalOffset, 1].Style.Font.Size = 9;
            worksheet.Cells[3 + soCalOffset, 2].Style.Font.Size = 9;
            worksheet.Cells[2 + soCalOffset, 2].Style.Font.Bold = true;
            worksheet.Cells[3 + soCalOffset, 1].Style.Font.Bold = true;
            worksheet.Cells[3 + soCalOffset, 2].Style.Font.Bold = true;
            worksheet.Cells[15 + soCalOffset, 1].Style.Font.Bold = true;

            worksheet.SetValue(4 + soCalOffset, 1, "01");
            worksheet.SetValue(5 + soCalOffset, 1, "02");
            worksheet.SetValue(6 + soCalOffset, 1, "03");
            worksheet.SetValue(7 + soCalOffset, 1, "06");
            worksheet.SetValue(8 + soCalOffset, 1, "07");
            worksheet.SetValue(9 + soCalOffset, 1, "08");
            worksheet.SetValue(10 + soCalOffset, 1, "11");
            worksheet.SetValue(11 + soCalOffset, 1, "13");
            worksheet.SetValue(12 + soCalOffset, 1, "14");
            worksheet.SetValue(13 + soCalOffset, 1, "15");

            worksheet.Cells[4 + soCalOffset, 1].Style.Font.Size = 9;
            worksheet.Cells[4 + soCalOffset, 2].Style.Font.Size = 9;
            worksheet.Cells[5 + soCalOffset, 1].Style.Font.Size = 9;
            worksheet.Cells[5 + soCalOffset, 2].Style.Font.Size = 9;
            worksheet.Cells[6 + soCalOffset, 1].Style.Font.Size = 9;
            worksheet.Cells[6 + soCalOffset, 2].Style.Font.Size = 9;
            worksheet.Cells[7 + soCalOffset, 1].Style.Font.Size = 9;
            worksheet.Cells[7 + soCalOffset, 2].Style.Font.Size = 9;
            worksheet.Cells[8 + soCalOffset, 1].Style.Font.Size = 9;
            worksheet.Cells[8 + soCalOffset, 2].Style.Font.Size = 9;
            worksheet.Cells[9 + soCalOffset, 1].Style.Font.Size = 9;
            worksheet.Cells[9 + soCalOffset, 2].Style.Font.Size = 9;
            worksheet.Cells[10 + soCalOffset, 1].Style.Font.Size = 9;
            worksheet.Cells[10 + soCalOffset, 2].Style.Font.Size = 9;
            worksheet.Cells[11 + soCalOffset, 1].Style.Font.Size = 9;
            worksheet.Cells[11 + soCalOffset, 2].Style.Font.Size = 9;
            worksheet.Cells[12 + soCalOffset, 1].Style.Font.Size = 9;
            worksheet.Cells[12 + soCalOffset, 2].Style.Font.Size = 9;
            worksheet.Cells[13 + soCalOffset, 1].Style.Font.Size = 9;
            worksheet.Cells[13 + soCalOffset, 2].Style.Font.Size = 9;
            worksheet.Cells[15 + soCalOffset, 1].Style.Font.Size = 9;

            worksheet.SetValue(4 + soCalOffset, 2, "Active Sub rule: Subscriber is not active at the end of the day.  Applies to all compensation elements.");
            worksheet.SetValue(5 + soCalOffset, 2, "Prior History rule: Handset upgrade disqualified for prior usage history.");
            worksheet.SetValue(6 + soCalOffset, 2, "Handset rule:  Handset does not have a BrightPoint shipment record and is not a Houdini or BYOD handset.  Applies to New Acts and Handset Upgrades (upgrades not eligible if Houdini).");
            worksheet.SetValue(7 + soCalOffset, 2, "Account balance rule:  Customer account is not current.  Applies to all compensation elements.");
            worksheet.SetValue(8 + soCalOffset, 2, "Same day upgrade rule:  Handset upgrade is disqualified for same day as new activation or another handset upgrade for same subscriber");
            worksheet.SetValue(9 + soCalOffset, 2, "Multi-Upgrade rule: All handset upgrades must not occur within 30 days of a new activation, and BYOD handset upgrades must not occur within 90 days of a previous BYOD handset upgrade for the same subscriber.");
            worksheet.SetValue(10 + soCalOffset, 2, "3-day React Rule: Esn was not disconnected at least 3 full calendar days prior to reactivation");
            worksheet.SetValue(11 + soCalOffset, 2, "SOC eligibility rule: Rate Plan or Feature type is not eligible for compensation.");
            worksheet.SetValue(12 + soCalOffset, 2, "Same Day React Rule: Reacts are disqualified if customer upgrades on same date.");
            worksheet.SetValue(13 + soCalOffset, 2, "Termination rule: Terminated dealers and doors do not qualify for compensation");

            worksheet.SetValue(3 + soCalOffset, 1, "Reason Code");
            worksheet.SetValue(3 + soCalOffset, 2, "Description");

            worksheet.SetValue(2 + soCalOffset, 2, "Business Rule Reason Code Legend");
            worksheet.SetValue(15 + soCalOffset, 1, "The transactions listed below do not qualify for payment per the Dealer Compensation Business Rules");


            //was here before
            //for (int i = 1; i < 19; i++)
            //{
            //    worksheet.Cells[3 + soCalOffset, i].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //    worksheet.Cells[3 + soCalOffset, i].Style.Fill.BackgroundColor.SetColor(
            //        isSoCalReport
            //            ? System.Drawing.Color.FromArgb(217, 217, 217)
            //            : System.Drawing.Color.FromArgb(177, 160, 199));
            //}

            worksheet.SetValue(17, 17, DataHelpers.GetStartingMonthAndYear(startDate));
        }

        public static void FormatQualifiedReport(ref ExcelWorksheet worksheet, DateTime startDate, 
                                                 double sumTotal, int startRow, PropertyInfo[] properties, 
                                                 List<QualifiedTransactionRow> rows, bool isSoCalReport)
        {
            worksheet.SetValue(isSoCalReport ? 4 : 2, 17, DataHelpers.GetStartingMonthAndYear(startDate));
            worksheet.SetValue(rows.Count + startRow, properties.Count() - 1, "$" + String.Format("{0:0.00}", sumTotal));
            worksheet.Cells[rows.Count + startRow, properties.Count() - 1].Style.Font.Bold = true;
            worksheet.Cells[rows.Count + startRow, properties.Count() - 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[rows.Count + startRow, properties.Count() - 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(177, 160, 199));
        }
    }
}
