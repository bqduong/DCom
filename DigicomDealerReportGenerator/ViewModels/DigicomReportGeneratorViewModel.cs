using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using System.Windows.Input;

using DigicomDealerReportGenerator.FormattingHelper;
using DigicomDealerReportGenerator.MappingHelper;
using DigicomDealerReportGenerator.Models;

using LinqToExcel;

using DigicomDealerReportGenerator.Annotations;
using DigicomDealerReportGenerator.Commands;

using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace DigicomDealerReportGenerator.ViewModels
{
    public class DigicomReportGeneratorViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        public DigicomReportGeneratorViewModel()
        {
            var executionPath = this.Initialize();

            //dependency injection in the future
            this.QualifiedDisqualifiedReportGeneratorViewModel = new QualifiedDisqualifiedReportGeneratorViewModel(this, executionPath);
            this.CallidusReportGeneratorViewModel = new CallidusReportGeneratorViewModel(this, executionPath);
        }

        #region Properties

        public QualifiedDisqualifiedReportGeneratorViewModel QualifiedDisqualifiedReportGeneratorViewModel { get; set; }

        public CallidusReportGeneratorViewModel CallidusReportGeneratorViewModel { get; set; }

        #endregion Properties

        protected string Initialize()
        {
            return AppDomain.CurrentDomain.BaseDirectory;
        }

        //protected void GenerateCallidusReports(object param = null)
        //{
        //    //change to callidusReportGenerator
        //    var callidusReportGenerator = new CallidusReportGeneratorViewModel(this);

        //    using (ExcelPackage package = new ExcelPackage(DataHelpers.GetTemplateFile(this.IsQualified, this.IsSoCalReport, this.executionPath)))
        //    {
        //        //callidusReportGenerator.GenerateCallidusReport(package);
        //    }
        //}
        
        private void NotifyPropertyChanged(String propertyName = "")
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }
}