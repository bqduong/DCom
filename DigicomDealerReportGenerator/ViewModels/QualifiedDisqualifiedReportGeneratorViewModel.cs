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
    public class QualifiedDisqualifiedReportGeneratorViewModel : BaseViewModel, INotifyPropertyChanged
    {
        #region Fields

        private QualifiedDisqualifiedReportGeneratorModel qualifiedDisqualifiedReportGeneratorModel;

        private ExcelWorksheet templateWorksheet;

        private OpenFileDialog openFile;

        private string executionPath;

        private string sourcePath;

        private string destinationPath;

        private DateTime startDate;

        private DateTime endDate;

        private bool isQualified;

        private bool isSoCalReport;

        private string selectedSourceDealerDoorCode;

        private IEnumerable<ITransactionRow> masterTransactionList;

        private List<IDealerIdentification> masterDealerIdentificationList;

        public event PropertyChangedEventHandler PropertyChanged;

        #endregion

        public QualifiedDisqualifiedReportGeneratorViewModel(string executionPath) : base(executionPath)
        {
            this.qualifiedDisqualifiedReportGeneratorModel = new QualifiedDisqualifiedReportGeneratorModel(this);
        }

        #region Properties

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

        public DateTime StartDate
        {
            get
            {
                return this.startDate;
            }
            set
            {
                if (value != this.startDate)
                {
                    this.startDate = value;
                    this.NotifyPropertyChanged("StartDate");
                }
            }
        }

        public DateTime EndDate
        {
            get
            {
                return this.endDate;
            }
            set
            {
                if (value != this.endDate)
                {
                    this.endDate = value;
                    this.NotifyPropertyChanged("EndDate");
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

        public string SelectedSourceDealerDoorCode
        {
            get
            {
                return this.selectedSourceDealerDoorCode;
            }
            set
            {
                if (value != this.selectedSourceDealerDoorCode)
                {
                    this.selectedSourceDealerDoorCode = value;
                    this.NotifyPropertyChanged("SelectedSourceDealerDoorCode");
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

        public bool IsQualified
        {
            get
            {
                return this.isQualified;
            }
            set
            {
                this.isQualified = value;
            }
        }

        public IEnumerable<ITransactionRow> MasterTransactionList
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

        public ExcelWorksheet TemplateWorksheet
        {
            get
            {
                return this.templateWorksheet;
            }
            set
            {
                this.templateWorksheet = value;
            }
        }

        public bool IsSoCalReport
        {
            get
            {
                return this.isSoCalReport;
            }
            set
            {
                if (value != this.isSoCalReport)
                {
                    this.isSoCalReport = value;
                    this.NotifyPropertyChanged("IsSocalReport");
                }
            }
        }

        #endregion

        protected void Load(object param = null)
        {
            this.openFile = new OpenFileDialog();
            if (this.openFile.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    this.SourcePath = this.openFile.FileName;
                    this.Excel = new ExcelQueryFactory(this.SourcePath);
                    this.IsQualified = DataHelpers.IsQualified(this.SourcePath);
                    LinqToExcelMappingHelpers.MapToLinq(ref this.Excel, DataHelpers.GetReportType, this.SourcePath);

                    //populate dropdown list
                    this.MasterTransactionList = DataHelpers.GetMasterListOfTransactionRows(this.IsQualified, this.Excel);
                    this.MasterDealerIdentificationList = DataHelpers.GenerateDoorNameListWithDoorCode(this.MasterTransactionList);

                    //populate date range
                    this.StartDate = DataHelpers.GetEarliestDate(this.MasterTransactionList, this.IsQualified);
                    this.EndDate = DataHelpers.GetLatestDate(this.MasterTransactionList, this.IsQualified);
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

        protected void GenerateReports(object param = null)
        {
            var dealerReportGenerator = new QualifiedDisqualifiedReportGeneratorModel(this);

            if (SelectedSourceDealerDoorCode == "All")
            {
                var doorCodes =
                    this.MasterDealerIdentificationList.Where(m => m.DoorCode != "All").Select(m => m.DoorCode);

                foreach (var doorCode in doorCodes)
                {
                    using (ExcelPackage package = new ExcelPackage(DataHelpers.GetTemplateFile(this.IsQualified, this.IsSoCalReport, this.executionPath)))
                    {
                        dealerReportGenerator.GenerateSingleReport(doorCode, package);
                    }
                }
            }
            else
            {
                using (ExcelPackage package = new ExcelPackage(DataHelpers.GetTemplateFile(this.IsQualified, this.IsSoCalReport, this.executionPath)))
                {
                    dealerReportGenerator.GenerateSingleReport(this.SelectedSourceDealerDoorCode, package);
                }
            }
            MessageBox.Show("Done processing reports.");
        }
        
        private void NotifyPropertyChanged(String propertyName = "")
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }
}