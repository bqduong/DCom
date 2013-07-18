using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Input;

using DigicomDealerReportGenerator.Commands;
using DigicomDealerReportGenerator.MappingHelper;
using DigicomDealerReportGenerator.Models;

using LinqToExcel;

using OfficeOpenXml;

namespace DigicomDealerReportGenerator.ViewModels
{
    public class CommissionReportGeneratorViewModel : BaseViewModel
    {
        private CommissionReportGeneratorModel commissionReportGeneratorModel;

        private OpenFileDialog openFile;

        private string executionPath;

        private string sourcePath;

        private string destinationPath;

        private string selectedSourceDealerCode;

        private IEnumerable<CommissionRow> masterTransactionList;

        private List<IDealerIdentification> masterDealerIdentificationList;

        public CommissionReportGeneratorViewModel(string executionPath)
            : base(executionPath)
        {
            this.commissionReportGeneratorModel = new CommissionReportGeneratorModel(this);
        }


        public ICommand OpenFileClicked
        {
            get
            {
                return new SimpleCommand(this.Load);
            }
        }

        public ICommand SelectDestinationPathClicked
        {
            get
            {
                return new SimpleCommand(this.SetDestinationPath);
            }
        }

        public ICommand GenerateReportsClicked
        {
            get
            {
                return new SimpleCommand(this.GenerateReports);
            }
        }

        public string SourcePath
        {
            get
            {
                return this.sourcePath;
            }
            set
            {
                if (value != this.sourcePath)
                {
                    this.sourcePath = value;
                    this.NotifyPropertyChanged("SourcePath");
                }
            }
        }

        public string DestinationPath
        {
            get
            {
                return this.destinationPath;
            }
            set
            {
                if (value != this.destinationPath)
                {
                    this.destinationPath = value;
                    this.NotifyPropertyChanged("DestinationPath");
                }
            }
        }

        public string SelectedSourceDealerCode
        {
            get
            {
                return this.selectedSourceDealerCode;
            }
            set
            {
                if (value != this.selectedSourceDealerCode)
                {
                    this.selectedSourceDealerCode = value;
                    this.NotifyPropertyChanged("SelectedSourceDealerCode");
                }
            }
        }

        public List<IDealerIdentification> MasterDealerIdentificationList
        {
            get
            {
                return this.masterDealerIdentificationList;
            }
            set
            {
                if (value != this.masterDealerIdentificationList)
                {
                    this.masterDealerIdentificationList = value;
                    this.NotifyPropertyChanged("MasterDealerIdentificationList");
                }
            }
        }

        public IEnumerable<CommissionRow> MasterTransactionList
        {
            get
            {
                return this.masterTransactionList;
            }
            set
            {
                this.masterTransactionList = value;
            }
        }


        protected void Load(object param = null)
        {
            this.openFile = new OpenFileDialog();
            if (this.openFile.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    this.SourcePath = this.openFile.FileName;
                    this.Excel = new ExcelQueryFactory(this.SourcePath);
                    LinqToExcelMappingHelpers.MapCommissionRowToLinq(ref this.Excel);

                    //populate dropdown list
                    this.MasterTransactionList = DataHelpers.GetMasterListOfCommissionRows(this.Excel);
                    this.MasterDealerIdentificationList = DataHelpers.GenerateAgentListWithDealerCode(this.MasterTransactionList);
                }
                catch (Exception e)
                {
                    MessageBox.Show("Invalid excel file.  Please try again with another file");
                }
            }
        }

        protected void SetDestinationPath(object param = null)
        {
            var dialog = new FolderBrowserDialog();

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                this.DestinationPath = dialog.SelectedPath;
            }
        }

        public void GenerateReports(object param = null)
        {
            var commissionReportGenerator = new CommissionReportGeneratorModel(this);

            if (SelectedSourceDealerCode == "All")
            {
                var dealerCodes =
                    this.MasterDealerIdentificationList.Where(m => m.DoorCode != "All").Select(m => m.DoorCode);

                foreach (var dealerCode in dealerCodes)
                {
                    //using (ExcelPackage package = new ExcelPackage(DataHelpers.GetTemplateFile(this.IsQualified, this.IsSoCalReport, this.executionPath)))
                    //{
                    //    commissionReportGeneratorModel.GenerateSingleReport(dealerCode, package);
                    //}
                }
            }
            else
            {
                //using (ExcelPackage package = new ExcelPackage(DataHelpers.GetTemplateFile(this.IsQualified, this.IsSoCalReport, this.executionPath)))
                //{
                //    commissionReportGeneratorModel.GenerateSingleReport(this.SelectedSourceDealerCode, package);
                //}
            }
            MessageBox.Show("Done processing reports.");
        }
    }
}
