namespace Etk.Excel.UI.Windows.ModelManagement.ViewModels
{
    using Microsoft.Office.Interop.Excel;
    using MvvmBase;
    using System.Windows.Input;

    public class WizardViewModel : ViewModelBase
    {
        public RequestViewModel Request
        { get; private set; }

        public ViewPropertiesViewModel ViewProperties
        { get; set; }

        public AccessorsParametersViewModel AccessorsParameters
        { get; set; }
        

        #region .ctors and factories
        private WizardViewModel(Range caller, Range firstOutputRangeAddress)
        {
            Request = new RequestViewModel(this, caller, firstOutputRangeAddress);
        }

        static public WizardViewModel CreateInstance(Range caller, Range firstOutputRangeAddress)
        {
            return new WizardViewModel(caller, firstOutputRangeAddress);
        }
        #endregion
    }
}
