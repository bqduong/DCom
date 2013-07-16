using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using LinqToExcel;

namespace DigicomDealerReportGenerator.ViewModels
{
    public class BaseViewModel
    {
        public BaseViewModel(string executionPath)
        {
            //DI later
            this.ExecutionPath = executionPath;
        }

        protected ExcelQueryFactory Excel;

        protected string ExecutionPath;
    }
}
