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
            this.QualifiedDisqualifiedReportGeneratorViewModel = new QualifiedDisqualifiedReportGeneratorViewModel(executionPath);
            this.CallidusReportGeneratorViewModel = new CallidusReportGeneratorViewModel(executionPath);
            this.CommissionResidualReportGeneratorViewModel = new CommissionResidualReportGeneratorViewModel(executionPath);
        }

        #region Properties

        public QualifiedDisqualifiedReportGeneratorViewModel QualifiedDisqualifiedReportGeneratorViewModel { get; set; }

        public CallidusReportGeneratorViewModel CallidusReportGeneratorViewModel { get; set; }

        public CommissionResidualReportGeneratorViewModel CommissionResidualReportGeneratorViewModel { get; set; }

        #endregion Properties

        protected string Initialize()
        {
            return AppDomain.CurrentDomain.BaseDirectory;
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