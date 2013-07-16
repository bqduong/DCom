using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Input;

using DigicomDealerReportGenerator.Commands;
using DigicomDealerReportGenerator.Models;

namespace DigicomDealerReportGenerator.ViewModels
{
    public class CallidusReportGeneratorViewModel : BaseViewModel, INotifyPropertyChanged
    {
        private OpenFileDialog openFile;

        private DigicomReportGeneratorViewModel digicomReportGeneratorViewModel;

        private CallidusReportGeneratorModel callidusReportGeneratorModel;

        public event PropertyChangedEventHandler PropertyChanged;

        public CallidusReportGeneratorViewModel(DigicomReportGeneratorViewModel digicomReportGeneratorViewModel, string executionPath) : base(executionPath)
        {
            this.digicomReportGeneratorViewModel = digicomReportGeneratorViewModel;
        }

        public ICommand OpenFileClicked
        {
            get
            {
                return new SimpleCommand(this.Load);
            }
        }

        protected void Load(object param = null)
        {
            this.openFile = new OpenFileDialog();
            if (this.openFile.ShowDialog() == DialogResult.OK)
            {
            }
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