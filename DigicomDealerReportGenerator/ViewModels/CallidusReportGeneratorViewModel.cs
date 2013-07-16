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

namespace DigicomDealerReportGenerator.ViewModels
{
    public class CallidusReportGeneratorViewModel : BaseViewModel, INotifyPropertyChanged
    {
        private OpenFileDialog openFile;

        private CallidusReportGeneratorModel callidusReportGeneratorModel;

        private string bayAreaSourcePath;

        private string soCalSourcePath;

        private DateTime dateSelect;

        public event PropertyChangedEventHandler PropertyChanged;

        public CallidusReportGeneratorViewModel(string executionPath) : base(executionPath)
        {
            this.callidusReportGeneratorModel = new CallidusReportGeneratorModel(this);
        }

        public string BayAreaSourcePath
        {
            get
            {
                return this.bayAreaSourcePath;
            }
            set
            {
                if (value != this.bayAreaSourcePath)
                {
                    this.bayAreaSourcePath = value;
                    this.NotifyPropertyChanged("BayAreaSourcePath");
                }
            }
        }

        public string SoCalSourcePath
        {
            get
            {
                return this.soCalSourcePath;
            }
            set
            {
                if (value != this.soCalSourcePath)
                {
                    this.soCalSourcePath = value;
                    this.NotifyPropertyChanged("SoCalSourcePath");
                }
            }
        }
        
        public DateTime DateSelect
        {
            get
            {
                return this.dateSelect;
            }
            set
            {
                if (value != this.dateSelect)
                {
                    this.dateSelect = value;
                    this.callidusReportGeneratorModel.DateSelect = value;
                    this.NotifyPropertyChanged("DateSelect");
                }
            }
        }

        public ICommand LoadBayAreaFileClicked
        {
            get
            {
                return new SimpleCommand(this.LoadBayAreaFile);
            }
        }

        public ICommand LoadSoCalFileClicked
        {
            get
            {
                return new SimpleCommand(this.LoadBayAreaFile);
            }
        }

        public ICommand LoadQPayRetailMasterClicked
        {
            get
            {
                return new SimpleCommand(this.LoadQPayRetailMasterFile);
            }
        }

        public ICommand LoadQPayOnlineMasterClicked
        {
            get
            {
                return new SimpleCommand(this.LoadQPayOnlineMasterFile);
            }
        }

        public ICommand ProcessQPayReportsClicked
        {
            get
            {
                return new SimpleCommand(this.ProcessQPayReports);
            }
        }

        protected void LoadBayAreaFile(object param = null)
        {
            this.openFile = new OpenFileDialog();
            if (this.openFile.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    var sourcePath = this.openFile.FileName;
                    this.Excel = new ExcelQueryFactory(sourcePath);
                    LinqToExcelMappingHelpers.MapToLinq(ref this.Excel, DataHelpers.GetReportType, sourcePath);

                    //populate data list
                    var isQualified = DataHelpers.IsQualified(sourcePath);
                    this.callidusReportGeneratorModel.MasterBayAreaTransactionList = DataHelpers.GetMasterListOfTransactionRows(isQualified, this.Excel);

                    this.BayAreaSourcePath = sourcePath;
                }
                catch (Exception e)
                {
                    this.BayAreaSourcePath = "";
                    MessageBox.Show("Invalid excel file.  Please try again with another file");
                }
            }
        }

        protected void LoadSoCalFile(object param = null)
        {
            this.openFile = new OpenFileDialog();
            if (this.openFile.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    var sourcePath = this.openFile.FileName;
                    this.Excel = new ExcelQueryFactory(sourcePath);
                    LinqToExcelMappingHelpers.MapToLinq(ref this.Excel, DataHelpers.GetReportType, sourcePath);

                    //populate data list
                    var isQualified = DataHelpers.IsQualified(sourcePath);
                    this.callidusReportGeneratorModel.MasterSoCalTransactionList = DataHelpers.GetMasterListOfTransactionRows(isQualified, this.Excel);

                    this.SoCalSourcePath = sourcePath;
                }
                catch (Exception e)
                {
                    this.SoCalSourcePath = "";
                    MessageBox.Show("Invalid excel file.  Please try again with another file");
                }
            }
        }

        protected void LoadQPayRetailMasterFile(object param = null)
        {
            this.openFile = new OpenFileDialog();
            if (this.openFile.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    var sourcePath = this.openFile.FileName;
                    this.Excel = new ExcelQueryFactory(sourcePath);
                    LinqToExcelMappingHelpers.MapToLinq(ref this.Excel, DataHelpers.GetReportType, sourcePath);

                    //populate data list
                    var isQualified = DataHelpers.IsQualified(sourcePath);
                    this.callidusReportGeneratorModel.MasterSoCalTransactionList = DataHelpers.GetMasterListOfTransactionRows(isQualified, this.Excel);

                    this.SoCalSourcePath = sourcePath;
                }
                catch (Exception e)
                {
                    this.SoCalSourcePath = "";
                    MessageBox.Show("Invalid excel file.  Please try again with another file");
                }
            }
        }

        protected void LoadQPayOnlineMasterFile(object param = null)
        {
            this.openFile = new OpenFileDialog();
            if (this.openFile.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    var sourcePath = this.openFile.FileName;
                    this.Excel = new ExcelQueryFactory(sourcePath);
                    LinqToExcelMappingHelpers.MapToLinq(ref this.Excel, DataHelpers.GetReportType, sourcePath);

                    //populate data list
                    var isQualified = DataHelpers.IsQualified(sourcePath);
                    this.callidusReportGeneratorModel.MasterSoCalTransactionList = DataHelpers.GetMasterListOfTransactionRows(isQualified, this.Excel);

                    this.SoCalSourcePath = sourcePath;
                }
                catch (Exception e)
                {
                    this.SoCalSourcePath = "";
                    MessageBox.Show("Invalid excel file.  Please try again with another file");
                }
            }
        }

        protected void ProcessQPayReports(object param = null)
        {
            this.callidusReportGeneratorModel.ProcessQPayReports();   
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