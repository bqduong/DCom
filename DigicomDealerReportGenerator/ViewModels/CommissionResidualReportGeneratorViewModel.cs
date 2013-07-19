using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
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
    public class CommissionResidualReportGeneratorViewModel : BaseViewModel
    {
        private CommissionResidualReportGeneratorModel commissionReportGeneratorModel;

        private OpenFileDialog openFile;

        private string executionPath;

        private string sourcePath;

        private string destinationPath;

        private string selectedCommissionSourceDealerCode;

        private string selectedResidualSourceDealerCode;

        private string weekInput;

        private IEnumerable<CommissionRow> masterCommissionTransactionList;

        private IEnumerable<ResidualRow> masterResidualTransactionList;

        private List<IDealerIdentification> masterCommissionDealerIdentificationList;

        private List<IDealerIdentification> masterResidualDealerIdentificationList;

        public CommissionResidualReportGeneratorViewModel(string executionPath)
            : base(executionPath)
        {
            this.commissionReportGeneratorModel = new CommissionResidualReportGeneratorModel(this);
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

        public ICommand GenerateCommissionReportsClicked
        {
            get
            {
                return new SimpleCommand(this.GenerateCommissionReports);
            }
        }

        public ICommand GenerateResidualReportsClicked
        {
            get
            {
                return new SimpleCommand(this.GenerateResidualReports);
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

        public string SelectedCommissionSourceDealerCode
        {
            get
            {
                return this.selectedCommissionSourceDealerCode;
            }
            set
            {
                if (value != this.selectedCommissionSourceDealerCode)
                {
                    this.selectedCommissionSourceDealerCode = value;
                    this.NotifyPropertyChanged("selectedCommissionSourceDealerCode");
                }
            }
        }

        public string SelectedResidualSourceDealerCode
        {
            get
            {
                return this.selectedResidualSourceDealerCode;
            }
            set
            {
                if (value != this.selectedResidualSourceDealerCode)
                {
                    this.selectedResidualSourceDealerCode = value;
                    this.NotifyPropertyChanged("SelectedResidualSourceDealerCode");
                }
            }
        }

        public List<IDealerIdentification> MasterCommissionDealerIdentificationList
        {
            get
            {
                return this.masterCommissionDealerIdentificationList;
            }
            set
            {
                if (value != this.masterCommissionDealerIdentificationList)
                {
                    this.masterCommissionDealerIdentificationList = value;
                    this.NotifyPropertyChanged("MasterCommissionDealerIdentificationList");
                }
            }
        }

        public List<IDealerIdentification> MasterResidualDealerIdentificationList
        {
            get
            {
                return this.masterResidualDealerIdentificationList;
            }
            set
            {
                if (value != this.masterResidualDealerIdentificationList)
                {
                    this.masterResidualDealerIdentificationList = value;
                    this.NotifyPropertyChanged("MasterResidualDealerIdentificationList");
                }
            }
        }

        public IEnumerable<CommissionRow> MasterCommissionTransactionList
        {
            get
            {
                return this.masterCommissionTransactionList;
            }
            set
            {
                this.masterCommissionTransactionList = value;
            }
        }

        public string WeekInput
        {
            get
            {
                return this.weekInput;
            }
            set
            {
                if (value != this.weekInput)
                {
                    this.weekInput = value;
                    this.NotifyPropertyChanged("WeekInput");
                }
            }
        }

        public IEnumerable<ResidualRow> MasterResidualTransactionList
        {
            get
            {
                return this.masterResidualTransactionList;
            }
            set
            {
                this.masterResidualTransactionList = value;
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
                    LinqToExcelMappingHelpers.MapResidualRowToLinq(ref this.Excel);

                    //populate dropdown list
                    this.MasterCommissionTransactionList = DataHelpers.GetMasterListOfCommissionRows(this.Excel);
                    this.MasterCommissionDealerIdentificationList = DataHelpers.GenerateAgentListWithDealerCode(this.MasterCommissionTransactionList);

                    //populate residual list
                    this.MasterResidualTransactionList = DataHelpers.GetMasterListOfResidualRows(this.Excel);
                    this.MasterResidualDealerIdentificationList = DataHelpers.GenerateAgentListWithDealerCode(this.MasterResidualTransactionList);
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

        public void GenerateCommissionReports(object param = null)
        {
            if (this.SelectedCommissionSourceDealerCode == "[All Dealers]")
            {
                var fullDealerIds =
                    this.MasterCommissionDealerIdentificationList.Where(m => m.DoorCode != "All").Select(m => m.FullDealerIdentification);

                foreach (var fullDealerId in fullDealerIds)
                {
                    using (ExcelPackage package = new ExcelPackage(new FileInfo(this.ExecutionPath + "Digicom Templates\\Commission Report Template.xlsx")))
                    {
                        commissionReportGeneratorModel.GenerateSingleCommissionReport(fullDealerId, package);
                    }
                }
            }
            else
            {
                using (ExcelPackage package = new ExcelPackage(new FileInfo(this.ExecutionPath + "Digicom Templates\\Commission Report Template.xlsx")))
                {
                    commissionReportGeneratorModel.GenerateSingleCommissionReport(this.selectedCommissionSourceDealerCode, package);
                }
            }
            MessageBox.Show("Done processing commission reports.");
        }

        public void GenerateResidualReports(object param = null)
        {
            if (SelectedResidualSourceDealerCode == "[All Dealers]")
            {
                var fullDealerIds =
                    this.MasterCommissionDealerIdentificationList.Where(m => m.DoorCode != "All").Select(m => m.FullDealerIdentification);

                foreach (var fullDealerId in fullDealerIds)
                {
                    using (ExcelPackage package = new ExcelPackage(new FileInfo(this.ExecutionPath + "Digicom Templates\\Residual Report Template.xlsx")))
                    {
                        commissionReportGeneratorModel.GenerateSingleResidualReport(fullDealerId, package);
                    }
                }
            }
            else
            {
                using (ExcelPackage package = new ExcelPackage(new FileInfo(this.ExecutionPath + "Digicom Templates\\Residual Report Template.xlsx")))
                {
                    commissionReportGeneratorModel.GenerateSingleResidualReport(this.SelectedResidualSourceDealerCode, package);
                }
            }
            MessageBox.Show("Done processing residual reports.");
        }
    }
}
