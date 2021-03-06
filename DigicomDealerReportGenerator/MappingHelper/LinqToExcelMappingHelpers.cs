﻿using System;
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
        public const string Rebate = "rebate";

        public static void MapToLinq(ref ExcelQueryFactory excel, Func<string, string> getReportType, string filename)
        {
            LinqToExcelMappingHelpers.ModifyCommonTransactionRowMappings(ref excel);

            switch (getReportType(filename))
            {
                case Disqualified:
                    LinqToExcelMappingHelpers.ModifyDisqualilfiedTransactionRowMappings(ref excel);
                    break;
                case Rebate:
                    LinqToExcelMappingHelpers.ModifyRebateTransactionRowMappings(ref excel);
                    break;
                default:
                    LinqToExcelMappingHelpers.ModifyQualilfiedTransactionRowMappings(ref excel);
                    break;
            }
        }

        public static void ModifyDisqualilfiedTransactionRowMappings(ref ExcelQueryFactory excel)
        {
            excel.AddMapping<DisqualifiedTransactionRow>(q => q.EsnHistory, "ESN History");
            excel.AddMapping<DisqualifiedTransactionRow>(q => q.SimHistory, "SIM History");
            excel.AddMapping<DisqualifiedTransactionRow>(q => q.TmobileLastNetworkHistory, "T-mobile Last Network History");
            excel.AddMapping<DisqualifiedTransactionRow>(q => q.SubscriberStatus, "Subscriber Status");
            excel.AddMapping<DisqualifiedTransactionRow>(q => q.AccountBalance, "Account Balance");
            excel.AddMapping<DisqualifiedTransactionRow>(q => q.BusinesRuleReasonCode, "Business Rule Reason Code");
            excel.AddMapping<DisqualifiedTransactionRow>(q => q.TransactionType, "Transaction Type");
            excel.AddMapping<DisqualifiedTransactionRow>(q => q.RatePlan, "Rate Plan");
            excel.AddMapping<DisqualifiedTransactionRow>(q => q.BoltOn, "Bolt On");
        }

        public static void ModifyQualilfiedTransactionRowMappings(ref ExcelQueryFactory excel)
        {
            excel.AddMapping<QualifiedTransactionRow>(q => q.EsnHistory, "ESN History");
            excel.AddMapping<QualifiedTransactionRow>(q => q.SimHistory, "SIM History");
            excel.AddMapping<QualifiedTransactionRow>(q => q.TmobileLastNetworkHistory, "T-mobile Last Network History");
            excel.AddMapping<QualifiedTransactionRow>(q => q.Location, "Location");
            excel.AddMapping<QualifiedTransactionRow>(q => q.RatePlanAmount, "Rate Plan Amount");
            excel.AddMapping<QualifiedTransactionRow>(q => q.BoltOnAmount, "Bolt On Amount");
            excel.AddMapping<QualifiedTransactionRow>(q => q.TransactionAmount, "Transaction Amount");
            excel.AddMapping<QualifiedTransactionRow>(q => q.PostedDate, "Posted Date");
            excel.AddMapping<QualifiedTransactionRow>(q => q.TransactionType, "Transaction Type");
            excel.AddMapping<QualifiedTransactionRow>(q => q.RatePlan, "Rate Plan");
            excel.AddMapping<QualifiedTransactionRow>(q => q.BoltOn, "Bolt On");
        }

        public static void ModifyRebateTransactionRowMappings(ref ExcelQueryFactory excel)
        {
            excel.AddMapping<IRebateRow>(q => q.ProgramName, "Program Name");
            excel.AddMapping<IRebateRow>(q => q.RebateType, "Rebate Type");
            excel.AddMapping<RebateTransactionRow>(q => q.Location, "Location");
            excel.AddMapping<IRebateRow>(q => q.QualificationStatus, "Qualification Status");
            excel.AddMapping<RebateTransactionRow>(q => q.RebateAmount, "Rebate Amount");
            excel.AddMapping<IRebateRow>(q => q.SubscriberId, "Subscriber ID");
            excel.AddMapping<IRebateRow>(q => q.PostedDate, "Posted Date");
        }

        public static void ModifyCommonTransactionRowMappings(ref ExcelQueryFactory excel)
        {
            excel.AddMapping<ITransactionRow>(q => q.DoorCode, "Door Code");
            excel.AddMapping<ITransactionRow>(q => q.DoorName, "Door Name");
            excel.AddMapping<ITransactionRow>(q => q.Address, "Address");
            excel.AddMapping<ITransactionRow>(q => q.AccountNo, "Account Number");
            excel.AddMapping<ITransactionRow>(q => q.SubscriberId, "Subscriber  ID");
            excel.AddMapping<ITransactionRow>(q => q.Mdn, "MDN");
            excel.AddMapping<ITransactionRow>(q => q.Esn, "ESN");
            excel.AddMapping<ITransactionRow>(q => q.Sim, "SIM");
            excel.AddMapping<ITransactionRow>(q => q.HandsetModel, "Handset Model");
            excel.AddMapping<ITransactionRow>(q => q.TransactionDate, "Transaction Date");
        }


        public static void MapResidualRowToLinq(ref ExcelQueryFactory excel)
        {
            excel.AddMapping<ResidualRow>(r => r.Mrr, "MRR");  
            excel.AddMapping<ResidualRow>(r => r.AccountId, "ACCOUNT_ID");
            excel.AddMapping<ResidualRow>(r => r.ActivationDate, "ACTIVATION_DATE");
            excel.AddMapping<ResidualRow>(r => r.CustomerId, "CUSTOMER_ID");
            excel.AddMapping<ResidualRow>(r => r.MarketId, "MARKET_ID");  
            excel.AddMapping<ResidualRow>(r => r.MarketName, "MARKET_NAME");  
            excel.AddMapping<ResidualRow>(r => r.Technology, "TECHNOLOGY");  
            excel.AddMapping<ResidualRow>(r => r.DealerId, "DEALER_ID");  
            excel.AddMapping<ResidualRow>(r => r.DealerCode, "DEALER_CODE");  
            excel.AddMapping<ResidualRow>(r => r.Mac, "MAC");
            excel.AddMapping<ResidualRow>(r => r.Agent, "Agent");
            excel.AddMapping<ResidualRow>(r => r.ResidualAmount, "RESIDUAL AMOUNT");
            excel.AddMapping<ResidualRow>(r => r.RevenueClassName, "REVENUE_CLASS_NAME");  
        }


        public static void MapCommissionRowToLinq(ref ExcelQueryFactory excel)
        {
            excel.AddMapping<CommissionRow>(r => r.MarketId, "MARKET_ID");
            excel.AddMapping<CommissionRow>(r => r.MarketName, "MARKET_NAME");
            excel.AddMapping<CommissionRow>(r => r.LoginName, "LOGIN_NAME");
            excel.AddMapping<CommissionRow>(r => r.DealerCode, "DEALER_CODE");
            excel.AddMapping<CommissionRow>(r => r.DealerLocation, "DEALER_LOCATION");
            excel.AddMapping<CommissionRow>(r => r.OicTransactionType, "OIC_TRANSACTION_TYPE");
            excel.AddMapping<CommissionRow>(r => r.TransactionDate, "TRX_DATE");
            excel.AddMapping<CommissionRow>(r => r.OfferId, "OFFER_ID");
            excel.AddMapping<CommissionRow>(r => r.OfferName, "OFFER_NAME");
            excel.AddMapping<CommissionRow>(r => r.ContractType, "CONTRACT_TYPE");
            excel.AddMapping<CommissionRow>(r => r.AccountId, "ACCOUNT_ID");
            excel.AddMapping<CommissionRow>(r => r.CustomerId, "CUSTOMER_ID");
            excel.AddMapping<CommissionRow>(r => r.ActivationDate, "ACTIVATION_DATE");
            excel.AddMapping<CommissionRow>(r => r.CustomerFirstName, "CUSTOMER_FIRST_NAME");
            excel.AddMapping<CommissionRow>(r => r.CustomerLastName, "CUSTOMER_LAST_NAME");
            excel.AddMapping<CommissionRow>(r => r.AccountAge, "ACCOUNT_AGE");
            excel.AddMapping<CommissionRow>(r => r.ServiceType, "SERVICE_TYPE");
            excel.AddMapping<CommissionRow>(r => r.BundleType, "BUNDLE_TYPE");
            excel.AddMapping<CommissionRow>(r => r.EquipmentSerialNumber, "EQUIPMENT_SERIAL_NUMBER");
            excel.AddMapping<CommissionRow>(r => r.Agent, "Agent");
            excel.AddMapping<CommissionRow>(r => r.PlanElement, "PLAN_ELEMENT");
            excel.AddMapping<CommissionRow>(r => r.RecurringPrice, "RECURRING_PRICE");
            excel.AddMapping<CommissionRow>(r => r.SubscriberCount, "SUBSCRIBER_COUNT");
            excel.AddMapping<CommissionRow>(r => r.CommissionAmount, "COMMISSION_AMOUNT");
        }
    }
}
