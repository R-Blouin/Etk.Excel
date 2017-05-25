using Etk.Excel.UI.Windows.ModelManagement.Controls;
using Etk.Excel.UI.Windows.ModelManagement.ViewModels;

namespace Etk.Excel.UI.Windows.ModelManagement
{
    public partial class DynamicRequestManagementWindow
    {
        public DynamicRequestManagementWindow(WizardViewModel wizardViewModel)
        {
            InitializeComponent();

            Wizard.AddStep(0, new RequestProperties(wizardViewModel.Request));
            Wizard.AddStep(1, new ViewProperties(wizardViewModel.ViewProperties));
            Wizard.AddStep(2, new Accessorsparameters(wizardViewModel.AccessorsParameters));
        }
    }
}
