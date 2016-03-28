using System;

namespace Etk.Excel.UI.Windows.Wizard
{
    public interface IWizardStep
    {
        /// <summary>Call before 'OnNext' is called. return a object that will be passed as parameters of the 'OnNext method'</summary>
        /// <returns></returns>
        object GetNextStepData();

        /// <summary>Call jsut before the changing of step</summary>
        bool OnNext(object parameters);

        /// <summary>properties set by the Wizard to its steps. The steps must called it to change the 'Enable' feature if teh button</summary>
        event Action CanNext;

        /// <summary>Return true if the step is ok for the wizard to continue</summary>
        bool CheckCanNext();

        bool OnCancel();
    }
}
