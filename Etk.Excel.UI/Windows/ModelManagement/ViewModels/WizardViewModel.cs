﻿using Etk.Excel.MvvmBase;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace Etk.Excel.UI.Windows.ModelManagement.ViewModels
{
    public class WizardViewModel : ViewModelBase
    {
        public RequestViewModel Request
        { get; private set; }

        public ViewPropertiesViewModel ViewProperties
        { get; set; }

        public AccessorsParametersViewModel AccessorsParameters
        { get; set; }
        

        #region .ctors and factories
        private WizardViewModel(ExcelInterop.Range caller, ExcelInterop.Range firstOutputRangeAddress)
        {
            Request = new RequestViewModel(this, caller, firstOutputRangeAddress);
        }

        public static WizardViewModel CreateInstance(ExcelInterop.Range caller, ExcelInterop.Range firstOutputRangeAddress)
        {
            return new WizardViewModel(caller, firstOutputRangeAddress);
        }
        #endregion
    }
}
