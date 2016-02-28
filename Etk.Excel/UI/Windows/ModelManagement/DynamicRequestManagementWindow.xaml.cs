namespace Etk.Excel.UI.Windows.ModelManagement
{
    using Controls;
    using System.Windows.Controls;
    using ViewModels;
    using Wizard;

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
