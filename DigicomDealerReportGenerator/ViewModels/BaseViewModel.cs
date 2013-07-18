using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using LinqToExcel;

namespace DigicomDealerReportGenerator.ViewModels
{
    public class BaseViewModel : INotifyPropertyChanged
    {
        public BaseViewModel(string executionPath)
        {
            //DI later
            this.ExecutionPath = executionPath;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected ExcelQueryFactory Excel;

        public string ExecutionPath;

        public void NotifyPropertyChanged(String propertyName = "")
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }
}
