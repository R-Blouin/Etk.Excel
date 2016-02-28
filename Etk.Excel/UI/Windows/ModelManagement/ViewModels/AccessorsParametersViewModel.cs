namespace Etk.Excel.UI.Windows.ModelManagement.ViewModels
{
    using Etk.Excel.UI.Windows.Wizard;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    public class AccessorsParametersViewModel : IWizardStep
    {
        private WizardViewModel parent;

        #region .ctors
        public AccessorsParametersViewModel(WizardViewModel parent)
        {
            this.parent = parent;
        }
        #endregion

        #region IWizardStep interface implementation
        public object GetNextStepData()
        {
            return parent.ViewProperties;
        }

        public bool OnNext(object parameters)
        {
            return true;
        }

        public bool OnCancel()
        {
            return true;
        }

        public bool CheckCanNext()
        {
            return false;
        }

        event Action canNext;
        event Action IWizardStep.CanNext
        {
            add { canNext += value; }
            remove { canNext -= value; }
        }
        #endregion
    }
}
